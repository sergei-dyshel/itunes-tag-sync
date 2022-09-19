import logging
import logging.config
from dataclasses import dataclass
from typing import Any

import click
import colorlog
import eyed3.id3
import tqdm
import os.path
import win32com.client
from tqdm.contrib.logging import logging_redirect_tqdm

WMP = b'Windows Media Player 9 Series'

log = logging.getLogger(__name__)


def rating_as_stars(rating):
    return '★' * rating + '☆' * (5 - rating)


class SetRatingMode:
    TAG = "tag"
    ITUNES = "itunes"
    AUTO = "auto"

    ALL = [TAG, ITUNES, AUTO]


class Song:
    track: Any  # itunes track
    tag: eyed3.id3.Tag
    tag_dirty = False

    def __init__(self, track):
        self.track = track
        audio = eyed3.load(track.location)
        self.tag = audio.tag

    def itunes_rating(self) -> int:
        rating = self.track.Rating
        assert rating % 20 == 0, f'Invalid track rating {rating}'
        return rating // 20

    def tag_rating(self) -> int:
        popm = self.tag.popularities.get(WMP)
        if popm is None:
            return 0
        rating = popm.rating
        if rating <= 31:
            return 1
        if rating <= 95:
            return 2
        if rating <= 159:
            return 3
        if rating <= 221:
            return 4
        return 5

    def set_tag_rating(self, rating: int):
        if self.tag_rating() == rating:
            return
        log.info(f'setting tag rating to {rating}')
        popm = self.tag.popularities
        if rating == 0:
            popm.remove(WMP)
            return
        assert 1 <= rating <= 5
        popm_rating = [None, 1, 64, 128, 196, 255][rating]
        popm.set(WMP, popm_rating, 0)
        self.tag_dirty = True

    def set_itunes_rating(self, rating: int):
        if self.itunes_rating() == rating:
            return
        log.info(f'setting iTunes rating to {rating}')
        self.track.Rating = rating * 20


@dataclass
class Config:
    dry: bool = False
    all: bool = False
    verbose: bool = False
    clean: bool = False


cfg: Config = None

tracks: list[Any] = None


@click.group(chain=True)
@click.option('-a', '--all', is_flag=True, help='Process all tracks, not only selected ones')
@click.option('-n', '--dry', is_flag=True, help='Do not do any actual changes, just print actions')
@click.option('-v', '--verbose', is_flag=True, help='Verbose logging')
@click.option('-c', '--clean', is_flag=True, help='Delete non-existent songs')
def main(**kwargs):
    """
    TODO: add help for
    """
    global cfg
    cfg = Config(**kwargs)

    formatter = logging.Formatter("%(message)s")
    handler = colorlog.StreamHandler()
    handler.setFormatter(formatter)
    log.addHandler(handler)
    log.setLevel(logging.DEBUG if cfg.verbose else logging.INFO)


@main.command()
def update_from_tag():
    """Update iTunes song metadat from MP3 tag"""

    def command(song: Song):
        track = song.track
        tag = song.tag
        if tag.artist == track.Artist and tag.album == track.Album and tag.album_artist == track.AlbumArtist and tag.title == track.Name and tag.genre == track.Genre:
            return
        log.debug('Updating metadata from file tag')
        if cfg.dry:
            return
        song.track.UpdateInfoFromFile()

    return command


@main.command()
@click.argument('mode', type=click.Choice(SetRatingMode.ALL), default=SetRatingMode.AUTO)
def set_rating(mode: SetRatingMode):
    """Set rating based on iTunes metadata or MP3 tag"""

    def command(song: Song):
        log.debug(f'Tag rating {song.tag_rating()}, iTunes rating {song.itunes_rating()}')
        match mode:
            case SetRatingMode.TAG:
                song.set_itunes_rating(song.tag_rating())
            case SetRatingMode.ITUNES:
                song.set_tag_rating(song.itunes_rating())
            case SetRatingMode.AUTO:
                tag_rating = song.tag_rating()
                itunes_rating = song.itunes_rating()
                if tag_rating != 0 and itunes_rating == 0:
                    song.set_itunes_rating(tag_rating)
                if tag_rating == 0 and itunes_rating != 0:
                    song.set_itunes_rating(tag_rating)
                if tag_rating != 0 and itunes_rating != 0 and tag_rating != itunes_rating:
                    log.warn(f'different tag and itunes ratings: {tag_rating} vs {itunes_rating}')

    return command


def fix_eyed3_logs():
    from eyed3.mp3.headers import log
    log.setLevel(logging.ERROR)


@main.result_callback()
def process_commands(commands, **kwargs):
    fix_eyed3_logs()
    itunes = win32com.client.Dispatch("iTunes.Application")
    global tracks
    if cfg.all:
        tracks = itunes.LibraryPlaylist.Tracks
    else:
        tracks = itunes.SelectedTracks
        if tracks is None:
            raise Exception('iTunes not in songs view')

    count = tracks.Count
    print(f"Total {count} tracks selected")
    progress = tqdm.tqdm(iterable=range(1, tracks.Count + 1),
                         unit='tracks',
                         delay=1,
                         bar_format='{bar:20}{r_bar}{l_bar}'
                         )

    with logging_redirect_tqdm(loggers=[log]):
        for i in progress:
            track = tracks.Item(i)
            label = f'{track.Artist} - {track.Name}'
            progress.set_description(label, refresh=False)
            log.handlers[0].setFormatter(colorlog.ColoredFormatter(f'%(log_color)s{label}: %(message)s'))
            if track.Kind == 1:  # FileTrack
                if not os.path.isfile(track.Location):
                    if cfg.clean:
                        log.info('deleting non-existent file')
                        track.Delete()
                    else:
                        log.warn('file does not exists')
                    continue
                try:
                    song = Song(track)
                    for cmd in commands:
                        cmd(song)
                    if song.tag_dirty and not cfg.dry:
                        song.tag.save()
                except Exception as exc:
                    log.exception("", exc_info=exc)


if __name__ == '__main__':
    main()
