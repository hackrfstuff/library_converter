## A tiny helper to process Roon’s **Skipped files** export and make your library importable

> Always keep a backup. Originals are deleted **only after** the new FLAC verifies OK, but it’s still your library.

---

## Why FLAC?

FLAC is broadly supported (including by Roon), lossless, and less fussy than `.m4a/.mp4`.


## Requirements

- Python **3.9+**
- `ffmpeg` and `ffprobe` available on your system `PATH`
- Python deps:
  ```bash
  pip install pandas openpyxl


## Get the Excel report from Roon

`Settings → Library → (optional: Clean up library) → View skipped files → Export to Excel`

That `.xlsx` is what this tool reads.


## Usage

**Preview (no changes):**

```bash
python fix_skipped.py --xlsx "SkippedFiles.xlsx" --root "C:\path\to\music\folder"
```

**Apply (convert/repair + delete originals on success):**

```bash
python fix_skipped.py --xlsx "SkippedFiles.xlsx" --root "C:\path\to\music\folder" --apply
```

The script will print planned actions, stream **ffmpeg** progress during `--apply`, and write a CSV log next to your files.


## What the script does

* Scans the `.xlsx` for file paths/filenames
* Looks **only** in the `--root` folder (no deep recursive search)
* Plans actions:

  * `.m4a`/`.mp4` → `.flac`
  * broken `.flac` → re-encode `.flac`
* In `--apply`:

  * runs `ffmpeg` with live progress
  * verifies output (has audio stream, non-zero duration)
  * removes the original if verification passes
* Writes a CSV log to `--root`:

  * `fix_log_preview.csv` or `fix_log_apply.csv`


## Notes & Limitations

* Works on Windows/macOS/Linux as long as `ffmpeg`/`ffprobe` are on PATH.
* The script uses paths from the `.xlsx` or matches by **filename** in `--root`.
* If the destination `.flac` exists, a “ (2)”, “ (3)”, … suffix is added.


## Example commands

```bash
# Preview
python fix_skipped.py --xlsx "C:\Users\me\Desktop\SkippedFiles.xlsx" --root "D:\Music\Inbox"

# Apply
python fix_skipped.py --xlsx "C:\Users\me\Desktop\SkippedFiles.xlsx" --root "D:\Music\Inbox" --apply
```


## License

MIT
