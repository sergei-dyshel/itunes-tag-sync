from enum import Enum
from numbers import Integral
from typing import Any
import win32com.client
import logging
import typer
import eyed3
import eyed3.id3
import click
import tqdm
import logging
from tqdm.contrib.logging import logging_redirect_tqdm
from dataclasses import dataclass

WMP = b'Windows Media Player 9 Series'

log = logging.getLogger()

def rating_as_stars(rating):
    return '★' * rating + '☆' * (5 - rating)

@Enum
class SetRatingMode:
    TAG = "tag"
    ITUNES = "itunes"


class Song:
    track: Any # itunes track
    tag: eyed3.id3.Tag

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

    def set_itunes_rating(self, rating: int):
        if self.itunes_rating() == rating:
            return
        log.info(f'setting iTunes rating to {rating}')
        self.track.Rating = rating * 20


@dataclass
class Config:
    dry: bool
    all: bool
    verbose: bool

cfg = Config()

tracks: list[Any] = None


@click.group(chain=True)
@click.option('-a', '--all', is_flag=True, help='Process all tracks, not only selected ones')
@click.option('-n', '--dry', is_flag=True, help='Do not do any actual changes, just print actions')
@click.option('-v', '--verbose', is_flag=True, help='Verbose logging')
def main(**kwargs):
    """
    TODO: add help for
    """
    global cfg
    cfg = Config(**kwargs)

    logging.basicConfig()

@main.command()
def update_from_tag():
    '''Update iTunes song metadat from MP3 tag'''
    def command(song: Song):
        if not cfg.dry:
            log.debug('Updating metadata from file tag')
            song.track.UpdateInfoFromFile()
    return command


@main.command()
@click.argument('mode', type=click.Choice(list(SetRatingMode)), help='TODO')
def set_rating(mode: SetRatingMode):
    '''Set rating based on iTunes metadata or MP3 tag'''
    def command(song: Song):
        match mode:
            case SetRatingMode.TAG:
                song.set_itunes_rating(song.tag_rating())
            case SetRatingMode.ITUNES:
                song.set_tag_rating(song.itunes_rating())
    return command


@main.callback()
def process_commands(commands, **kwargs):
    from eyed3.mp3.headers import log
    log.setLevel(logging.ERROR)

    itunes = win32com.client.Dispatch("iTunes.Application")
    if cfg.all:
        tracks = itunes.LibraryPlaylist.Tracks
    else:
        tracks = itunes.SelectedTracks

    count = tracks.Count
    print(f"Total {count} tracks selected")
    progress = tqdm.tqdm(iterable=range(1, tracks.Count + 1),
        unit='tracks',
        delay=1,
    )

    with logging_redirect_tqdm():
        for i in progress:
            track = tracks.Item(i)
            label = f'{track.Artist} - {track.Name}'
            progress.set_description(label)
            log.handlers[0].setFormatter(logging.Formatter(f'{label}: %(message)s'))
            if track.Kind == 1:  # FileTrack
                try:
                    song = Song(track)
                    for cmd in commands:
                        cmd(song)
                except Exception as exc:
                    log.exception(exc_info=exc)



if __name__ == '__main__':
    main()
