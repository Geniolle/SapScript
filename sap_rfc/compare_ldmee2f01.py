#!/usr/bin/env python3
"""Compare LDMEE2F01 (the include with syntax errors) between DEV and PRD."""

import os
import sys
from pathlib import Path
from dotenv import load_dotenv

project_root = Path(__file__).parent.parent
sys.path.insert(0, str(project_root))

try:
    from pyrfc import Connection
except ImportError:
    print("pyrfc not available")
    sys.exit(1)

import difflib


def connect_env(prefix, label):
    try:
        conn = Connection(
            user=os.getenv(prefix + '_USER'),
            passwd=os.getenv(prefix + '_PASSWD'),
            ashost=os.getenv(prefix + '_ASHOST'),
            sysnr=os.getenv(prefix + '_SYSNR'),
            client=os.getenv(prefix + '_CLIENT'),
            lang=os.getenv(prefix + '_LANG', 'PT'),
        )
        print("[OK] Connected to " + label)
        return conn
    except Exception as e:
        print("[FAIL] " + label + ": " + str(e))
        return None


def read_include(conn, include_name, env_label):
    print("\n[*] Reading " + include_name + " from " + env_label + "...")

    try:
        result = conn.call(
            'RPY_PROGRAM_READ',
            PROGRAM_NAME=include_name,
            LANGUAGE='PT',
        )

        lines = {}
        source_extended = result.get('SOURCE_EXTENDED', [])

        for idx, item in enumerate(source_extended, start=1):
            if isinstance(item, dict):
                line_text = item.get('LINE', '')
            else:
                line_text = str(item)
            lines[idx] = line_text.rstrip()

        print("    [OK] " + str(len(lines)) + " lines read")
        return lines

    except Exception as e:
        print("    [ERROR] " + str(e))
        return None


def main():
    print("=" * 70)
    print("LDMEE2F01 (Include) Syntax Error Analysis: DEV vs PRD")
    print("=" * 70)

    env_file = project_root / ".env"
    load_dotenv(env_file)

    dev_conn = connect_env("SAP_DEV", "DEV")
    prd_conn = connect_env("SAP_PRD", "PRD")

    if not dev_conn or not prd_conn:
        print("\n[ERROR] Cannot connect to both")
        return False

    try:
        # Read LDMEE2F01
        dev_include = read_include(dev_conn, 'LDMEE2F01', 'DEV')
        prd_include = read_include(prd_conn, 'LDMEE2F01', 'PRD')

        if dev_include is None or prd_include is None:
            print("\n[ERROR] Failed to read include")
            return False

        # Compare
        print("\n[*] Comparing...")

        if dev_include == prd_include:
            print("    [✓] IDENTICAL")
        else:
            print("    [!] DIFFERENT")
            print("        DEV: " + str(len(dev_include)) + " lines")
            print("        PRD: " + str(len(prd_include)) + " lines")

        # Analyze error lines
        error_lines = [25, 30, 78, 538, 730, 2465, 2505, 2532, 2617, 2791]
        print("\n[*] Error lines analysis (from PRD report):")
        print("    (Lines are 1-based)")

        has_differences = False
        for err_line in error_lines:
            dev_line = dev_include.get(err_line, "[MISSING]")
            prd_line = prd_include.get(err_line, "[MISSING]")

            if dev_line != prd_line:
                has_differences = True
                print("\n    Line " + str(err_line) + ": DIFFERS")
                print("        DEV: " + dev_line[:70])
                print("        PRD: " + prd_line[:70])

        if not has_differences:
            print("\n    [!] NO DIFFERENCES at reported error lines!")
            print("        Possible causes:")
            print("        - Line numbers in error report might be off by 1 or based on different line counting")
            print("        - Errors might be in dependent objects (not LDMEE2F01 itself)")

        # Full diff
        print("\n[*] Full line-by-line diff...")

        dev_lines = [str(k).rjust(5) + " " + v for k, v in sorted(dev_include.items())]
        prd_lines = [str(k).rjust(5) + " " + v for k, v in sorted(prd_include.items())]

        diff = list(difflib.unified_diff(dev_lines, prd_lines, lineterm='', n=1))

        if diff:
            diff_lines = [l for l in diff if l.startswith('+') or l.startswith('-')]
            print("    Total diff lines: " + str(len(diff_lines)))

            if diff_lines:
                print("\n    First 20 differences:")
                for line in diff_lines[:20]:
                    print("      " + line)
                if len(diff_lines) > 20:
                    print("      ... and " + str(len(diff_lines) - 20) + " more")
        else:
            print("    No differences found")

        print("\n[OK] Analysis complete")
        return True

    finally:
        try:
            dev_conn.close()
        except:
            pass
        try:
            prd_conn.close()
        except:
            pass


if __name__ == "__main__":
    success = main()
    sys.exit(0 if success else 1)
