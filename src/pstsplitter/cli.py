"""Command-line interface for pstsplitter.

Examples:
  python -m pstsplitter.cli --source path/to/mail.pst --out outdir --mode size --size 5GB \
      --stream --quiet --no-verify --summary summary.csv

Supports filters: --include-folders, --exclude-folders, --sender-domains, --date-range
Date range format: YYYY-MM-DD[:YYYY-MM-DD] (end inclusive). Use single date for >= that date.
"""
from __future__ import annotations

from pathlib import Path
import argparse
import logging
from datetime import datetime
from threading import Event

from .splitter import split_pst
from .util import configure_logging, format_bytes
from .outlook import is_outlook_available


def _parse_size(text: str) -> int:
    units = {"mb": 1024**2, "gb": 1024**3, "tb": 1024**4}
    t = text.strip().lower()
    for u, mul in units.items():
        if t.endswith(u):
            num = float(t[:-len(u)])
            return int(num * mul)
    # default MB if only number
    if t.isdigit():
        return int(t) * units["mb"]
    raise argparse.ArgumentTypeError(f"Cannot parse size '{text}' (expected e.g. 500MB / 5GB / 0.5TB)")


def _parse_date_range(spec: str):
    if not spec:
        return None
    parts = spec.split(":", 1)
    def parse_one(p: str | None):
        if not p:
            return None
        return datetime.strptime(p, "%Y-%m-%d")
    if len(parts) == 1:
        start = parse_one(parts[0])
        return (start, None)
    return (parse_one(parts[0]), parse_one(parts[1]))


def build_arg_parser() -> argparse.ArgumentParser:
    ap = argparse.ArgumentParser(description="Split Outlook PST files into smaller PSTs")
    ap.add_argument('--source', required=True, type=Path, help='Source PST file path')
    ap.add_argument('--out', required=True, type=Path, help='Output directory')
    ap.add_argument('--mode', choices=['size','year','month','folder'], default='size')
    ap.add_argument('--size', type=_parse_size, help='Max size per part (e.g. 2GB, 500MB) for size mode')
    ap.add_argument('--include-folders', help='Comma list of top-level folders to include')
    ap.add_argument('--exclude-folders', help='Comma list of top-level folders to exclude')
    ap.add_argument('--sender-domains', help='Comma list of sender domains to include (others filtered out)')
    ap.add_argument('--date-range', help='Date range filter YYYY-MM-DD or YYYY-MM-DD:YYYY-MM-DD')
    ap.add_argument('--stream', action='store_true', help='Enable streaming size mode')
    ap.add_argument('--fast-enum', action='store_true', help='Fast enumeration (skip pre-pass)')
    ap.add_argument('--include-non-mail', action='store_true', help='Include non-mail item classes')
    ap.add_argument('--move', action='store_true', help='Move items instead of copy')
    ap.add_argument('--quiet', action='store_true', help='Suppress per-item logs')
    ap.add_argument('--no-verify', action='store_true', help='Disable verification')
    ap.add_argument('--throttle-ms', type=int, default=250, help='Progress throttle in ms')
    ap.add_argument('--summary', type=Path, help='Write bucket summary CSV to this path')
    ap.add_argument('--log-level', default='INFO', help='Logging level (INFO/DEBUG/WARN)')
    ap.add_argument('--dry-run', action='store_true', help='Dry run (enumerate and plan only)')
    return ap


def main(argv: list[str] | None = None) -> int:
    ap = build_arg_parser()
    args = ap.parse_args(argv)
    if not is_outlook_available():
        ap.error('Outlook / pywin32 not available on this system')
    if args.mode == 'size' and not args.size:
        ap.error('--size is required for size mode')
    if args.mode != 'size' and args.stream:
        ap.error('--stream is only valid with --mode size')
    args.out.mkdir(parents=True, exist_ok=True)
    stop_event = configure_logging(getattr(logging, args.log_level.upper(), logging.INFO))
    include_set = {s.strip() for s in (args.include_folders or '').split(',') if s.strip()}
    exclude_set = {s.strip() for s in (args.exclude_folders or '').split(',') if s.strip()}
    domain_set = {s.strip().lstrip('@') for s in (args.sender_domains or '').split(',') if s.strip()}
    date_rng = _parse_date_range(args.date_range) if args.date_range else None

    cancel = Event()  # placeholder - could wire signals

    def progress(done_items: int, total_items: int, done_bytes: int, total_bytes: int):
        if total_items:
            pct = done_items / total_items * 100
            logging.info("Progress: %s/%s items (%.1f%%) %s/%s", done_items, total_items, pct, format_bytes(done_bytes), format_bytes(total_bytes))
        else:
            logging.info("Progress: %s items processed (streaming)", done_items)

    result = split_pst(
        args.source,
        args.out,
        args.mode,
        args.size if args.mode == 'size' else None,
        cancel,
        progress_cb=progress,
        dry_run=args.dry_run,
        include_non_mail=args.include_non_mail,
        move_items=args.move,
        verify=not args.no_verify,
        fast_enumeration=args.fast_enum,
        suppress_item_logs=args.quiet,
        stream_size_mode=args.stream,
        throttle_progress_ms=args.throttle_ms,
        include_folders=include_set or None,
        exclude_folders=exclude_set or None,
        sender_domains=domain_set or None,
        date_range=date_rng,
        summary_csv=args.summary,
    )
    logging.info("Completed: %s files, %s items, %s bytes", len(result.created_files), result.total_items, format_bytes(result.total_bytes))
    if result.errors:
        logging.warning("Errors encountered: %s", len(result.errors))
    stop_event.set()
    return 0


if __name__ == '__main__':  # pragma: no cover
    raise SystemExit(main())
