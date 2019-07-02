# jrkong's Personal Audio Sorter/Backup Script
A very simple audio organizer which organizes and backs up music built using Python 3.

The code will automatically build the directory structure by organizing the music in an `Artist/(album release date) Album Title/audio files` scheme and will use the music file tags to do so. The code assumes the music is properly tagged (if it is not I suggest using [MusicBrainz Picard](https://picard.musicbrainz.org/) to do so).

# Using the script
1. Update the source, target and backup paths in the `findAndMoveAudio()` function
1. If this is the first run on the machine, run `pip install -r requirements.txt` to install all requirements
1. Run the script using `python audioSorter.py`

# TODO List:
- Finalize and implement command line arguments
- Remove assumptions
- Update logic to find source directories/make source crawling logic more robust
- Create tests