"""Multi-source spreadsheet audit CLI (Python 3.10+)."""

from argparse import ArgumentParser
from pathlib import Path
import sys

from src.audit import run_audit


def main() -> int:
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
        sys.stderr.reconfigure(encoding="utf-8")
    parser = ArgumentParser(description="核对资产、BMC 与采集表，输出差异报告")
    parser.add_argument("--assets", type=Path, required=True)
    parser.add_argument("--bm", type=Path, required=True)
    parser.add_argument("--collected", type=Path, required=True)
    parser.add_argument("--output", type=Path, default=Path("output"))
    args = parser.parse_args()
    try:
        result = run_audit(args.assets, args.bm, args.collected, args.output)
    except (OSError, ValueError) as exc:
        parser.exit(2, f"输入错误: {exc}\n")
    print(f"报告已生成：{args.output}")
    print(f"输入 {result['input_rows']} 行；状态统计 {result['status_counts']}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
