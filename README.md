A tiny helper to process Roon’s **Skipped files** export:

- Convert **`.m4a` / `.mp4` → `.flac`**
- “Repair” **corrupt `.flac`** by re-encoding to FLAC
- Preserve tags (`-map_metadata 0`)
- Verify output then **remove the original** on success
- **Preview** mode (no changes) and **Apply** mode
- **Live ffmpeg progress** while converting

> ⚠️ Always keep a backup. The script deletes the original file *after* verifying the new FLAC.

## Why FLAC?
FLAC is broadly supported (including by Roon), lossless, and less fussy about containers. Converting lossy sources (AAC in `.m4a`) to FLAC doesn’t improve quality, but it puts them in a simple, predictable container Roon likes.

## Requirements
- Python 3.9+
- `ffmpeg` and `ffprobe` available on your system `PATH`
- Python deps:
  ```bash
  pip install pandas openpyxl

If you would like more supported formats added, let me know :)