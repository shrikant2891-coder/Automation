#!/usr/bin/env python3
"""Apply a payroll command to payroll.xml (dashboard reads the same file)."""

import sys

from payroll_dashboard.commands import CommandError, apply_command


def main(argv: list[str] | None = None) -> int:
    argv = argv if argv is not None else sys.argv[1:]
    if not argv or argv[0] in ("-h", "--help", "help"):
        from payroll_dashboard.commands import _cmd_help

        print(_cmd_help())
        return 0
    cmd = " ".join(argv)
    try:
        msg = apply_command(cmd)
        print(msg)
        return 0
    except CommandError as e:
        print(f"Error: {e}", file=sys.stderr)
        return 1
    except FileNotFoundError as e:
        print(f"Error: {e}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
