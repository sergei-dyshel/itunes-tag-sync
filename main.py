import datetime
import logging
import logging.config
import os.path
from dataclasses import dataclass
from typing import Any, Iterable, Protocol, Set

import click
import colorlog
import eyed3.id3
from eyed3.mp3.headers import log as eyed3_log

import tqdm
import win32com.client
from tqdm.contrib.logging import logging_redirect_tqdm

dry: bool

WMP = b'Windows Media Player 9 Series'

log = logging.getLogger(__name__)

def rating_as_stars(rating):
    return '★' * rating + '☆' * (5 - rating)


class Track(Protocol):
    Name: str
    Artist: str
    Rating: int
    Location: str
    PlayedDate: datetime.datetime

Tag = eyed3.id3.Tag

def get_tag_rating(tag: Tag) -> int:
    popm = tag.popularities.get(WMP)
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

def set_tag_rating(tag: Tag, rating: int):
    popm = tag.popularities
    if rating == 0:
        popm.remove(WMP)
        return
    assert 1 <= rating <= 5
    popm_rating = [None, 1, 64, 128, 196, 255][rating]
    popm.set(WMP, popm_rating, 0)

def sync_rating(tag: Tag, track: Track) -> bool:
    itunes_rating = track.Rating
    assert itunes_rating % 20 == 0, f'Invalid iTunes rating {itunes_rating}'
    itunes_rating //= 20

    tag_rating = get_tag_rating(tag)

    if tag_rating != 0 and itunes_rating == 0:
        log.info(f'updating iTunes rating to {itunes_rating} => {tag_rating}')
        if not dry:
            track.Rating = tag_rating * 20
        return False
    if tag_rating == 0 and itunes_rating != 0:
        log.info(f'updating tag rating to {tag_rating} => {itunes_rating}')
        if not dry:
            set_tag_rating(tag, itunes_rating)
        return not dry
    if tag_rating != 0 and itunes_rating != 0 and tag_rating != itunes_rating:
        tag_date = datetime.datetime.fromtimestamp(os.path.getmtime(track.Location))
        track_date = track.PlayedDate.replace(tzinfo=None)
        delta = tag_date - track_date
        log.debug(f'file mtime is {tag_date}, track last played date is {track_date}')
        if delta.total_seconds() < 1:
            log.info(f'updating older tag rating {tag_rating} => {itunes_rating}')
            if not dry:
                set_tag_rating(tag, itunes_rating)
            return not dry
        else:
            log.info(f'updating older iTunes rating {itunes_rating} => {tag_rating}')
            if not dry:
                track.Rating = tag_rating * 20
            return False
    return False

def sync_tag(track: Track):
    try:
        audio = eyed3.load(track.Location)
        tag: eyed3.id3.Tag = audio.tag
    except Exception as exc:
        log.error(f'could not load file: {exc}')
        return

    dirty = sync_rating(tag, track)
    if dirty and not dry:
        tag.save()

def tqdm_tracks(tracks: Iterable[Track]) -> Iterable[Track]:
    progress = tqdm.tqdm(iterable=range(1, tracks.Count + 1),
                         unit='tracks',
                         delay=1,
                         bar_format='{bar:20}{r_bar}{l_bar}'
                         )
    for i in progress:
        track: Track = tracks.Item(i)
        if track.Kind != 1:
            continue
        label = f'{track.Artist} - {track.Name}'
        progress.set_description(label, refresh=False)

        yield track

def tqdm_redirect_log():
    return logging_redirect_tqdm(loggers=[log, eyed3_log])

def scan_for_new_files(library):
    log.info('Scanning track locations...')
    locations: Set[str] = set()
    with tqdm_redirect_log():
        for track in tqdm_tracks(library.LibraryPlaylist.Tracks):
            locations.add(track.Location)
    dirs = sorted(os.path.dirname(loc) for loc in locations)
    i = 0
    while i < len(dirs) - 1:
        if dirs[i + 1].startswith(dirs[i]):
            del dirs[i + 1]
        else:
            i += 1

    for directory in dirs:
        log.info(f'Scanning "{directory}"')
        if not dry:
            library.AddFiles(directory)

@click.command()
@click.option('--sel', '--selected', is_flag=True, help='Process only selected tracks')
@click.option('-n', '--dry', '_dry', is_flag=True, help='Do not do any actual changes, just print actions')
@click.option('-v', '--verbose', is_flag=True, help='Verbose logging')
@click.option('--clean', is_flag=True, help='Delete non-existent songs')
@click.option('--update', is_flag=True, help='Force update iTunes metadata from file tag')
@click.option('--sync', is_flag=True, help='Sync file tag with iTunes')
@click.option('--scan', is_flag=True, help='Scan folders for new files')
def main(selected: bool, _dry: bool, verbose: bool, clean: bool, update: bool, sync: bool, scan: bool):
    """
    TODO: add help for
    """

    global dry
    dry = _dry

    logging.basicConfig(format='%(module)s %(message)s')

    formatter = logging.Formatter("%(message)s")
    handler = colorlog.StreamHandler()
    handler.setFormatter(formatter)
    log.addHandler(handler)
    log.setLevel(logging.DEBUG if verbose else logging.INFO)

    eyed3_log.setLevel(logging.ERROR)

    itunes = win32com.client.Dispatch("iTunes.Application")
    library = itunes.LibraryPlaylist
    if selected:
        tracks = library.Tracks
        if tracks is None:
            raise Exception('iTunes not in songs view')
    else:
        tracks = itunes.SelectedTracks
        if tracks is None:
            raise Exception('iTunes not in songs view or no track selected')


    if scan:
        if selected:
            log.error('Scanning for new files is diabled with --selected')
        else:
            scan_for_new_files(library)

    if not update and not sync and not clean:
        return

    with tqdm_redirect_log():
        for track in tqdm_tracks(tracks):
            label = f'{track.Artist} - {track.Name}'
            log.handlers[0].setFormatter(colorlog.ColoredFormatter(f'%(log_color)s{label}: %(message)s'))
            if not os.path.isfile(track.Location):
                if clean:
                    log.info(f'deleting non-existent file {track.Location}')
                    if not dry:
                        track.Delete()
                else:
                    log.warning(f'file does not exists: {track.Location}')
                continue

            try:
                if update and not dry:
                    track.UpdateInfoFromFile()
                if sync:
                    sync_tag(track)
            except Exception as exc:
                log.exception("", exc_info=exc)





if __name__ == '__main__':
    main()
