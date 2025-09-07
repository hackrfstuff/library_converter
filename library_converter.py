#!/usr/bin/env python3
"""

I wrote this to clean up Roon’s “Skipped files” export. It does two things:

  1) Convert .m4a/.mp4 to .flac (preserves tags)
  2) “Repair” broken .flac by re-encoding to FLAC

There are only two modes:

  - Preview (default): plan and print what would happen; no changes
  - Apply (--apply): run ffmpeg with live progress, verify output, then delete the original

How to use:

  1) In Roon: Settings → Library → (optional) Clean up library →
               View skipped files → Export to Excel
  2) Run one of:
       python library_converter.py --xlsx "SkippedFiles.xlsx" --root "C:\\path\\to\\folder"
       python library_converter.py --xlsx "SkippedFiles.xlsx" --root "C:\\path\\to\\folder" --apply

Requirements:
  - Python 3.9+
  - pip install pandas openpyxl
  - ffmpeg and ffprobe available on PATH

Notes:
  - The script uses paths from the .xlsx, or matches by filename in --root.
  - If a destination filename already exists, “ (2)”, “ (3)”, etc. is appended.
  - Back up first. Originals are removed only after a verified success,
    but it’s still your library.
"""

import argparse
import csv
import json
import os
import re
import shutil
import subprocess
import sys
from pathlib import Path
from typing import Iterable, List, Optional, Set, Tuple

import pandas as pd

# ------- constants -------
MUSIC_EXTS: Set[str] = {".flac", ".m4a", ".mp4"}
M4A_EXTS: Set[str] = {".m4a", ".mp4"}
FLAC_EXTS: Set[str] = {".flac"}

PATHLIKE_COL_KEYWORDS = [
    "path", "file", "location", "filename", "track file", "file path", "source", "fullpath", "url"
]

TIME_RE = re.compile(r"time=(\d+):(\d+):(\d+(?:\.\d+)?)")
SPEED_RE = re.compile(r"speed=\s*([0-9.]+x)")

# ------- utils -------

def which(cmd: str) -> Optional[str]:
    return shutil.which(cmd)

def require_tools(tools: Iterable[str]) -> None:
    missing = [t for t in tools if which(t) is None]
    if missing:
        print(f"Error: required tool(s) not found on PATH: {', '.join(missing)}", file=sys.stderr)
        sys.exit(2)

def read_all_sheets(xlsx_path: Path) -> List[pd.DataFrame]:
    try:
        xls = pd.ExcelFile(xlsx_path)
        return [xls.parse(sheet_name) for sheet_name in xls.sheet_names]
    except Exception as e:
        print(f"Error reading '{xlsx_path}': {e}", file=sys.stderr)
        sys.exit(1)

def find_path_columns(df: pd.DataFrame) -> List[str]:
    cols = []
    for c in df.columns:
        lc = str(c).strip().lower()
        if any(k in lc for k in PATHLIKE_COL_KEYWORDS):
            cols.append(c)
    return cols

def looks_like_path(s: str) -> bool:
    if len(s) < 5:
        return False
    s2 = s.strip().lower()
    if not (("/" in s2) or ("\\" in s2)):
        # still allow bare filenames with extension
        pass
    return any(s2.endswith(ext) for ext in MUSIC_EXTS)

def extract_candidate_paths(df: pd.DataFrame) -> List[str]:
    paths: List[str] = []
    path_cols = find_path_columns(df)
    for col in path_cols:
        ser = df[col].dropna()
        for v in ser.astype(str).tolist():
            if looks_like_path(v):
                paths.append(v.strip())
    if not paths:
        for v in df.astype(str).fillna("").values.ravel():
            s = v.strip()
            if looks_like_path(s):
                paths.append(s)
    # unique preserve order
    seen = set()
    out = []
    for p in paths:
        if p not in seen:
            out.append(p)
            seen.add(p)
    return out

def normalize_to_root(s: str, root: Path) -> Path:
    p = Path(s)
    if not p.is_absolute():
        p = (root / s)
    return p

def is_under(child: Path, root: Path) -> bool:
    try:
        child.resolve().relative_to(root.resolve())
        return True
    except Exception:
        return False

def unique_target_path(target: Path) -> Path:
    if not target.exists():
        return target
    stem, suf = target.stem, target.suffix
    n = 2
    while True:
        cand = target.with_name(f"{stem} ({n}){suf}")
        if not cand.exists():
            return cand
        n += 1

def ffprobe_json(path: Path) -> Optional[dict]:
    cmd = ["ffprobe", "-v", "error", "-show_streams", "-show_format", "-of", "json", str(path)]
    try:
        out = subprocess.run(cmd, capture_output=True, text=True, check=True)
        return json.loads(out.stdout or "{}")
    except subprocess.CalledProcessError:
        return None

def file_duration_sec(path: Path) -> Optional[float]:
    info = ffprobe_json(path)
    if not info:
        return None
    try:
        return float(info.get("format", {}).get("duration", "0")) or None
    except Exception:
        return None

def ffmpeg_decode_test(path: Path) -> Tuple[bool, str]:
    cmd = ["ffmpeg", "-v", "error", "-xerror", "-nostdin", "-i", str(path), "-f", "null", "-"]
    proc = subprocess.run(cmd, capture_output=True, text=True)
    ok = (proc.returncode == 0)
    stderr = "\n".join(proc.stderr.splitlines()[-6:])
    return ok, stderr

def build_ffmpeg_cmd_convert_to_flac(src: Path, dst: Path) -> List[str]:
    # Lossless re-encode, preserve metadata from source
    return [
        "ffmpeg", "-hide_banner", "-nostdin", "-y",
        "-i", str(src),
        "-map_metadata", "0",
        "-c:a", "flac", "-compression_level", "8",
        "-map", "0:a:0",
        str(dst)
    ]

def verify_audio_ok(path: Path) -> bool:
    info = ffprobe_json(path)
    if not info:
        return False
    fmt = info.get("format", {})
    try:
        dur = float(fmt.get("duration", "0"))
    except Exception:
        dur = 0.0
    has_stream = any(st.get("codec_type") == "audio" for st in info.get("streams", []))
    return has_stream and dur > 0.5

# ------- planning -------

class Action:
    def __init__(self, kind: str, src: Path, dst: Optional[Path], cmd: Optional[List[str]], note: str = ""):
        self.kind = kind              # 'convert_m4a' | 'repair_flac'
        self.src = src
        self.dst = dst
        self.cmd = cmd or []
        self.note = note

def decide_actions_for_path(src: Path) -> List[Action]:
    actions: List[Action] = []
    ext = src.suffix.lower()

    if ext in M4A_EXTS:
        dst = src.with_suffix(".flac")
        if dst.exists():
            dst = unique_target_path(dst)
        cmd = build_ffmpeg_cmd_convert_to_flac(src, dst)
        actions.append(Action("convert_m4a", src, dst, cmd, "m4a→flac"))
        return actions

    if ext in FLAC_EXTS:
        ok, _ = ffmpeg_decode_test(src)
        if ok:
            return []
        dst = src.with_suffix(".flac")
        if dst.exists():
            dst = unique_target_path(dst)
        cmd = build_ffmpeg_cmd_convert_to_flac(src, dst)
        actions.append(Action("repair_flac", src, dst, cmd, "re-encode flac"))
        return actions

    return []

def plan_from_xlsx(xlsx: Path, root: Path):
    rows = []
    actions: List[Action] = []

    sheets = read_all_sheets(xlsx)
    candidates: List[str] = []
    for df in sheets:
        candidates.extend(extract_candidate_paths(df))

    seen = set()
    for raw in candidates:
        if raw in seen:
            continue
        seen.add(raw)

        resolved: Optional[Path] = None

        # 1) absolute/relative join
        p_try = normalize_to_root(raw, root)
        if p_try.exists() and p_try.is_file() and is_under(p_try, root):
            resolved = p_try

        # 2) fallback: just filename inside root (no recursive search)
        if resolved is None:
            base = Path(raw).name
            p2 = root / base
            if p2.exists() and p2.is_file():
                resolved = p2

        if resolved is None:
            rows.append((raw, None, "NOT FOUND"))
            continue

        if resolved.suffix.lower() not in MUSIC_EXTS:
            rows.append((raw, resolved, f"SKIP (unsupported ext: {resolved.suffix})"))
            continue

        acts = decide_actions_for_path(resolved)
        if not acts:
            rows.append((raw, resolved, "OK (no fix required)"))
        else:
            rows.append((raw, resolved, "PLAN: " + ", ".join(a.kind for a in acts)))
            actions.extend(acts)

    return rows, actions

# ------- execution with live progress -------

def parse_time_to_seconds(s: str) -> float:
    h, m, sec = s.split(":")
    return int(h) * 3600 + int(m) * 60 + float(sec)

def run_ffmpeg_with_progress(cmd: List[str], src: Path, idx: int, total: int) -> int:
    duration = file_duration_sec(src) or 0.0
    title = f"[{idx}/{total}] {src.name}"
    print(f"\n▶ Converting {title}")
    # Stream stderr so we can parse progress; also echo ffmpeg lines for transparency
    proc = subprocess.Popen(
        cmd,
        stdout=subprocess.DEVNULL,
        stderr=subprocess.PIPE,
        text=True,
        bufsize=1
    )
    try:
        last_line = ""
        for line in proc.stderr:
            last_line = line.rstrip()
            # try to parse time and speed
            tmatch = TIME_RE.search(last_line)
            smatch = SPEED_RE.search(last_line)
            if tmatch and duration > 0:
                h, m, s = tmatch.groups()
                cur = int(h) * 3600 + int(m) * 60 + float(s)
                pct = max(0.0, min(100.0, (cur / duration) * 100.0))
                spd = smatch.group(1) if smatch else ""
                sys.stdout.write(f"\r   {pct:6.2f}%  time={int(cur//60):02d}:{int(cur%60):02d}  {('speed='+spd) if spd else ''}        ")
                sys.stdout.flush()
            # If not parsable, still show occasional raw lines that contain time/size
            elif "time=" in last_line or "size=" in last_line:
                sys.stdout.write("\r   " + last_line[:100] + " " * 10)
                sys.stdout.flush()
        rc = proc.wait()
        sys.stdout.write("\r")
        sys.stdout.flush()
        return rc
    finally:
        if proc and proc.poll() is None:
            proc.kill()

def safe_remove(path: Path) -> None:
    try:
        os.remove(path)
    except Exception as e:
        print(f"[WARN] Failed to remove '{path}': {e}", file=sys.stderr)

def apply_actions(actions: List[Action], log_csv: Path, preview: bool) -> Tuple[int, int]:
    success = 0
    error = 0
    with log_csv.open("w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow(["kind", "source", "destination", "status", "note"])
        total = sum(1 for a in actions if a.kind in {"convert_m4a", "repair_flac"})
        idx = 0
        for a in actions:
            if a.kind not in {"convert_m4a", "repair_flac"}:
                continue
            idx += 1
            src = a.src
            dst = unique_target_path(a.dst) if a.dst else None
            if preview:
                print(f"[PREVIEW] {a.kind}: '{src.name}' -> '{dst.name if dst else 'N/A'}'")
                w.writerow([a.kind, str(src), str(dst) if dst else "", "PREVIEW", a.note])
                continue

            # Ensure final cmd uses (possibly uniquified) dst
            cmd = list(a.cmd)
            if dst and str(dst) != a.cmd[-1]:
                cmd = a.cmd[:-1] + [str(dst)]

            rc = run_ffmpeg_with_progress(cmd, src, idx, total)
            if rc != 0:
                print(f"✗ Failed: {src.name}")
                w.writerow([a.kind, str(src), str(dst) if dst else "", "ERROR", a.note])
                error += 1
                # remove partial output
                if dst and dst.exists():
                    safe_remove(dst)
                continue

            if not dst or not dst.exists() or not verify_audio_ok(dst):
                print(f"✗ Output verification failed: {src.name}")
                w.writerow([a.kind, str(src), str(dst) if dst else "", "ERROR: verify", a.note])
                error += 1
                if dst and dst.exists():
                    safe_remove(dst)
                continue

            # success → remove original
            safe_remove(src)
            print(f"✓ Done: {src.name}  →  {dst.name}")
            w.writerow([a.kind, str(src), str(dst), "OK", a.note])
            success += 1

    return success, error

# ------- CLI -------

def parse_args() -> argparse.Namespace:
    ap = argparse.ArgumentParser(
        description="Analyze skipped .xlsx and fix listed audio files (preview/apply). "
                    "Only M4A→FLAC conversions and FLAC repairs. No deep search."
    )
    ap.add_argument("--xlsx", required=True, help="Path to the skipped files .xlsx report.")
    ap.add_argument("--root", required=True, help="Root folder where the files should exist.")
    ap.add_argument("--apply", action="store_true", help="Apply fixes (default: preview only).")
    return ap.parse_args()

def main():
    args = parse_args()
    require_tools(["ffmpeg", "ffprobe"])

    xlsx = Path(args.xlsx).expanduser()
    root = Path(args.root).expanduser()
    if not xlsx.exists():
        print(f"Error: xlsx not found: {xlsx}", file=sys.stderr)
        sys.exit(1)
    if not root.exists() or not root.is_dir():
        print(f"Error: root is not a directory: {root}", file=sys.stderr)
        sys.exit(1)

    preview = not args.apply
    log_csv = root / ("fix_log_preview.csv" if preview else "fix_log_apply.csv")

    mode = "PREVIEW" if preview else "APPLY"
    print(f"\n=== {mode} MODE ===")
    print(f"Report: {xlsx}")
    print(f"Root:   {root}\n")

    found_rows, actions = plan_from_xlsx(xlsx, root)

    total_rows = len(found_rows)
    planned = sum(1 for _, _, s in found_rows if str(s).startswith("PLAN:"))
    notfound = sum(1 for _, rp, s in found_rows if rp is None)
    print(f"Rows parsed: {total_rows}")
    print(f"Files planned for fix: {planned}")
    print(f"Not found: {notfound}\n")

    for raw, rp, status in found_rows[:2000]:
        print(f"- {raw}  ->  {str(rp) if rp else '<missing>'}  ::  {status}")

    print("\nPlanned actions:")
    if not actions:
        print("(none)")
    else:
        for a in actions:
            if a.kind in {"convert_m4a", "repair_flac"}:
                print(f"* {a.kind}: '{a.src.name}' -> '{a.dst.name if a.dst else 'N/A'}' [{a.note}]")

    ok, err = apply_actions(actions, log_csv=log_csv, preview=preview)

    print(f"\nDone. Success: {ok}, Errors: {err}")
    print(f"Log written to: {log_csv}\n")

if __name__ == "__main__":
    main()
