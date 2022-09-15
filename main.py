import win32com.client
import logging
import typer
import eyed3

app = typer.Typer()

WMP = b'Windows Media Player 9 Series'


def rating_as_stars(rating):
    return '★' * rating + '☆' * (5 - rating)


def itunes_rating(track):
    rating = track.Rating
    assert rating % 20 == 0, f'Invalid track rating {rating}'
    return rating // 20


def tag_rating(audio):
    tag = audio.tag
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


def set_tag_rating(audio, rating):
    popm = audio.tag.popularities
    if rating == 0:
        popm.remove(WMP)
        return
    assert 1 <= rating <= 5
    popm_rating = [None, 1, 64, 128, 196, 255][rating]
    popm.set(WMP, popm_rating, 0)


def getRating(rate):
    r = int(rate)
    if r == 100:
        return 255
    elif r == 80:
        return 192
    elif r == 60:
        return 128
    elif r == 40:
        return 64
    elif r == 20:
        return 1
    else:
        return 0


@app.command("list")
def list_tracks(ctx: typer.Context):
    tracks = ctx.obj
    count = tracks.Count
    print(f"Total {count} tracks selected")
    for i in range(1, tracks.Count + 1):
        track = tracks.Item(i)
        if track.Kind == 1:  # FileTrack
            # track.UpdateInfoFromFile()
            star_rating = rating_as_stars(track.Rating // 20)
            print(f"{track.Artist} - {track.Name} - {star_rating} - {track.ModificationDate}")
            audio = eyed3.load(track.location)
            rating = itunes_rating(track)
            set_tag_rating(audio, rating)
            # audio.tag.save()


@app.callback()
def main(ctx: typer.Context, all: bool = False):
    """
    Some help
    """

    from eyed3.mp3.headers import log
    log.setLevel(logging.ERROR)

    itunes = win32com.client.Dispatch("iTunes.Application")
    if all:
        ctx.obj = itunes.LibraryPlaylist.Tracks
    else:
        ctx.obj = itunes.SelectedTracks


@app.command()
def refresh():
    print("hello")


if __name__ == '__main__':
    app()
