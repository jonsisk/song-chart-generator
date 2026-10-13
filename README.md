# Live Charts — local song chart generator

A local web app that turns simple lyric/chord text into a **single portrait PowerPoint deck or PDF**. Each selected song becomes **exactly one slide/page**, in the order shown in the setlist. Designed for the **9th-generation iPad (3:4 portrait)**.

## Launch on Windows

1. Unzip the project into a permanent folder (e.g. `Documents\LiveCharts`).
2. Make sure **Python 3.10 or newer** is installed. During Python installation, select **Add Python to PATH**.
3. Double-click **`start_windows.bat`**. The first launch creates a local virtual environment and installs the three packages from `requirements.txt` (internet access is needed *for that first installation only*).
4. Your browser should open at **http://localhost:8765**. If it doesn't, enter that address yourself.
5. Keep the black command window open while using the app. Close it or press `Ctrl+C` when finished.

For Mac/Linux, run `./start_mac_linux.sh` in a terminal.

## How to use

- Click **+** to add songs. Enter artist, title, and chart text; **changes save automatically** to `data/songs.json`.
- Use **drag and drop** or the **▲/▼** arrows to set slide order.
- Check or uncheck songs to include/exclude them from **the next export**. Unchecked songs **stay in your library**.
- Click **PowerPoint** to export one `.pptx` with each checked song on its own **fully editable slide**, or **PDF** for the matching printable/iPad document.
- Click **Backup library** periodically to save a portable JSON backup. **Import library** restores a backup but replaces the current library (confirmation required).
- The preloaded Blondie chart is based on your example and can be freely edited.

### Input syntax

```text
[Verse 1]
D5 | One way or another...
B5 | One day, maybe next week...

[Chorus]
NOTE: Guitar doubles vocals
I will drive past your house

[Break]
4 BAR DRUM BREAK

[Solo]
NOTE: G# > A on high E string

[Outro]
! Single notes, not chords
Repeat 3 times
```

- Sections in square brackets set the background and label. Available colors: Verse, Chorus, Pre-chorus, Bridge, Intro, Solo, Instrumental, Break, and Outro.
- `D5 | Lyric` puts the chord over that lyric. A standalone chord line followed by a lyric works too.
- `NOTE:`, `CUE:`, or a line beginning with `!` makes a bold reminder.
- Ordinary text creates a lyric line.
- The colored roadmap at the top shows the structure of the whole song.

## Layout behavior

- PowerPoint slide size: **7.5 in wide by 10 in high**, matching the iPad's **3:4** screen ratio in portrait.
- Font sizes adjust slightly to fit, but the generator will **warn about overlong songs**, not silently clip text. Shorten repeated lines, use reminders instead, or simplify dense sections if the warning appears.
- Slide elements in the `.pptx` are normal editable text boxes and rectangles. Editing a generated file **does not update the app's source library**—make permanent changes in the app and export again.
- PDF and PowerPoint use the same measured chart layout, though font rendering can vary slightly between devices.

## Privacy and backup

The app runs on **your computer only** (`127.0.0.1:8765`), has no online service, and stores its song library in `data/songs.json`. It does not send your songs anywhere. Packages are downloaded only during first-time setup. Keep backups of `data/songs.json` / exported JSON files.

Advanced: run `python app.py --port 9999 --no-browser` to choose a different local port. By default the server is accessible only from the computer running it.
