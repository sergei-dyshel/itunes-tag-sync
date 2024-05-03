from __future__ import annotations

import datetime
import logging
import logging.config
import os.path
from typing import Iterable, Protocol, TypeVar, Literal, Optional

import click
import colorlog
import eyed3.id3
import tqdm
import win32com.client
from eyed3.core import log as core_log
from eyed3.id3.frames import log as frames_log
from eyed3.mp3.headers import log as headers_log
from tqdm.contrib.logging import logging_redirect_tqdm

EYED3_LOGGERS = [core_log, headers_log, frames_log]

dry: bool

WMP = b'Windows Media Player 9 Series'

root_log = logging.getLogger()

log = logging.getLogger(__name__)


# log.propagate = None


def rating_as_stars(rating):
    return '★' * rating + '☆' * (5 - rating)


class Object(Protocol):
    Name: str


class Track(Protocol):
    Name: str
    Artist: str
    Rating: int
    Location: str
    Kind: int
    PlayedDate: datetime.datetime

    def Delete(self): ...

    def UpdateInfoFromFile(self): ...


Item = TypeVar('Item')


class Collection(Protocol[Item]):
    Count: int

    def Item(self, i: int) -> Item: ...


class NamedCollection(Collection[Item]):
    def ItemByName(self, name: str) -> Item: ...


def iterate(coll: Collection[Item]) -> Iterable[Item]:
    for i in range(1, coll.Count + 1):
        yield coll.Item(i)


TrackCollection = Collection[Track]


class Playlist(Object):
    Kind: int
    Tracks: TrackCollection


class UserPlaylist(Playlist):
    Smart: bool
    Parent: Optional[UserPlaylist]


class Library(Playlist):

    def AddFiles(self, paths: list[str]): ...


class Source(Object):
    Kind: int
    Playlists: NamedCollection[Playlist]


class ITunes(Protocol):
    LibraryPlaylist: Library
    SelectedTracks: TrackCollection
    Sources: NamedCollection[Source]


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


ForceRating = Literal['itunes', 'tag']


def sync_rating(tag: Tag, track: Track, force_rating: ForceRating = None) -> bool:
    itunes_rating = track.Rating
    assert itunes_rating % 20 == 0, f'Invalid iTunes rating {itunes_rating}'
    itunes_rating //= 20

    tag_rating = get_tag_rating(tag)

    def update_itunes_rating(msg):
        log.info(f'{msg} iTunes rating {itunes_rating} => {tag_rating}')
        if not dry:
            track.Rating = tag_rating * 20
        return False

    def update_tag_rating(msg):
        log.info(f'{msg} tag rating {tag_rating} => {itunes_rating}')
        if not dry:
            set_tag_rating(tag, itunes_rating)
        return not dry

    if tag_rating != 0 and itunes_rating == 0:
        return update_itunes_rating('setting')
    if tag_rating == 0 and itunes_rating != 0:
        return update_tag_rating('setting')
    if tag_rating != 0 and itunes_rating != 0 and tag_rating != itunes_rating:
        if force_rating == 'itunes':
            return update_tag_rating('FORCE-setting')
        if force_rating == 'tag':
            return update_itunes_rating('FORCE-setting')

        tag_date = datetime.datetime.fromtimestamp(os.path.getmtime(track.Location))
        track_date = track.PlayedDate.replace(tzinfo=None)
        delta = tag_date - track_date
        log.debug(f'file mtime is {tag_date}, track last played date is {track_date}')

        if delta.total_seconds() < 1:
            return update_tag_rating('updating older')
        else:
            return update_itunes_rating('updating older')
    return False


def sync_tag(track: Track, **kwargs):
    try:
        audio = eyed3.load(track.Location)
        tag: eyed3.id3.Tag = audio.tag
    except Exception as exc:
        log.error(f'could not load file: {exc}')
        return

    dirty = sync_rating(tag, track, **kwargs)
    if dirty and not dry:
        tag.save()


def get_label(track: Track) -> str:
    return f'{track.Artist} - {track.Name}'


def tqdm_tracks(tracks: TrackCollection) -> Iterable[Track]:
    progress = tqdm.tqdm(iterable=range(1, tracks.Count + 1),
                         unit='tracks',
                         delay=1,
                         bar_format='{bar:20}{r_bar}{l_bar}',
                         leave=False,
                         colour=''
                         )
    for i in progress:
        track: Track = tracks.Item(i)
        if track.Kind != 1:
            continue
        progress.set_description(get_label(track), refresh=False)

        yield track


def get_all_playlists(itunes: ITunes):
    sources = itunes.Sources
    library_source = sources.ItemByName("Library")
    for pls in iterate(library_source.Playlists):
        log.info(f'{pls.Name} {pls.Kind}')


@click.command()
@click.option('--selected', '--sel', is_flag=True, help='Process only selected tracks')
@click.option('-n', '--dry', '_dry', is_flag=True, help='Do not do any actual changes, just print actions')
@click.option('-v', '--verbose', is_flag=True, help='Verbose logging')
@click.option('--clean', is_flag=True, help='Delete non-existent songs')
@click.option('--update', is_flag=True, help='Force update iTunes metadata from file tag')
@click.option('--sync', is_flag=True, help='Sync file tag with iTunes')
@click.option('--scan', multiple=True, type=click.Path(exists=True), help='Scan given folders for new files')
@click.option('--errors', is_flag=True, help='Log tag parsing errors')
@click.option('--force-rating', type=click.Choice(['itunes', 'tag']),
              help='When ratings differ always take the specified one')
def main(selected: bool, _dry: bool, verbose: bool, clean: bool, update: bool, sync: bool, scan: list[str],
         errors: bool, force_rating: ForceRating):
    """
    TODO: add help for
    """

    global dry
    dry = _dry

    formatter = logging.Formatter('[%(module)s] %(message)s')
    handler = colorlog.StreamHandler()
    handler.setFormatter(formatter)
    root_log.addHandler(handler)
    log.setLevel(logging.DEBUG if verbose else logging.INFO)

    for logger in EYED3_LOGGERS:
        logger.setLevel(logging.WARNING if errors else logging.ERROR)

    itunes: ITunes = win32com.client.Dispatch("iTunes.Application")
    library = itunes.LibraryPlaylist

    sources = itunes.Sources
    library_source = sources.ItemByName("Library")
    for pls in iterate(library_source.Playlists):
        log.info(f'{pls.Name} {pls.Kind}')
    if not selected:
        tracks = library.Tracks
        if tracks is None:
            raise Exception('iTunes not in songs view')
    else:
        tracks = itunes.SelectedTracks
        if tracks is None:
            raise Exception('iTunes not in songs view or no track selected')

    if scan:
        if selected:
            log.error('Scanning for new files is disabled with --selected')
        else:
            for folder in scan:
                log.info(f'Scanning "{folder}" for new files...')
                library.AddFiles([folder])

    if not update and not sync and not clean:
        return

    deleted_tracks: list[Track] = []

    with logging_redirect_tqdm():
        for track in tqdm_tracks(tracks):
            label = f'{track.Artist} - {track.Name}'
            root_log.handlers[0].setFormatter(
                colorlog.ColoredFormatter(f'{label}: %(log_color)s[%(module)s] %(message)s'))
            if track.Location == "":
                log.warning(f'file does not exist')
                deleted_tracks += [track]
                continue
            try:
                if update and not dry:
                    track.UpdateInfoFromFile()
                if sync:
                    sync_tag(track, force_rating=force_rating)
            except Exception as exc:
                log.exception("", exc_info=exc)

    if clean:
        for track in deleted_tracks:
            log.info(f'{get_label(track)}: deleting')
            if not dry:
                track.Delete()


if __name__ == '__main__':
    main()
