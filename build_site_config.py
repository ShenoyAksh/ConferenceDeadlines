#!/usr/bin/env python3
"""Write runtime configuration for the static conference deadline site."""

from __future__ import annotations

import argparse
import json
from pathlib import Path


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Build site/config.js for the static website.")
    parser.add_argument(
        "-o",
        "--output",
        default="site/config.js",
        help="output JavaScript config file, default: site/config.js",
    )
    parser.add_argument(
        "--urgent-days",
        type=int,
        default=10,
        help="mark deadlines with fewer than this many days as urgent, default: 10",
    )
    parser.add_argument(
        "--soon-days",
        type=int,
        default=30,
        help="mark deadlines within this many days as soon, default: 30",
    )
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    output_path = Path(args.output)
    output_path.parent.mkdir(parents=True, exist_ok=True)

    config = {
        "urgentDays": max(0, args.urgent_days),
        "soonDays": max(0, args.soon_days),
    }
    payload = json.dumps(config, ensure_ascii=False, indent=2)
    output_path.write_text(f"window.CORE_SITE_CONFIG = {payload};\n", encoding="utf-8")

    print(f"Wrote site config to {output_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
