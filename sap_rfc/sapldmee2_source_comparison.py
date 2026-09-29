#!/usr/bin/env python3
"""
Read and compare SAPLDMEE2 source using RPY_PROGRAM_READ RFC.
"""

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
import json


class SourceComparator:
    def __init__(self):
        pass

    def connect(self, prefix, label):
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

    def read_program_source(self, conn, program_name, env_label):
        print("\n[*] Reading " + program_name + " from " + env_label + "...")

        try:
            result = conn.call(
                'RPY_PROGRAM_READ',
                PROGRAM_NAME=program_name,
                LANGUAGE='PT',
            )

            # Extract source
            source_lines = {}
            source_data = result.get('SOURCE') or result.get('SOURCE_TAB') or []

            if isinstance(source_data, dict):
                source_data = [source_data]

            for idx, item in enumerate(source_data, start=1):
                if isinstance(item, dict):
                    line_text = item.get('LINE') or item.get('TEXT') or ""
                else:
                    line_text = str(item)

                source_lines[idx] = line_text.rstrip()

            print("    [OK] " + str(len(source_lines)) + " lines read")

            # Also get includes
            includes = []
            for key in ('INCLUDE_TAB', 'INCLUDETAB', 'INCLUDES'):
                inc_data = result.get(key)
                if inc_data:
                    if isinstance(inc_data, dict):
                        inc_data = [inc_data]
                    for inc_item in inc_data:
                        if isinstance(inc_item, dict):
                            inc_name = inc_item.get('INCLUDE') or inc_item.get('NAME') or ""
                            if inc_name:
                                includes.append(inc_name.strip().upper())

            if includes:
                print("    Includes: " + ', '.join(includes))

            return source_lines, includes

        except Exception as e:
            print("    [ERROR] " + str(e))
            return None, []

    def analyze_error_lines(self, dev_src, prd_src, error_lines):
        """Show context around error lines"""
        print("\n[!] Checking reported error lines...")

        context_size = 2

        for err_line in error_lines:
            dev_line = dev_src.get(err_line, "[MISSING]")
            prd_line = prd_src.get(err_line, "[MISSING]")

            if dev_line != prd_line:
                print("\n    Line " + str(err_line) + ": DIFFERS")

                # Show context
                start = max(1, err_line - context_size)
                end = min(max(dev_src.keys()), err_line + context_size)

                print("        DEV:")
                for ln in range(start, end + 1):
                    marker = ">>>" if ln == err_line else "   "
                    print(marker + " " + str(ln) + ": " + dev_src.get(ln, "[N/A]")[:60])

                print("        PRD:")
                for ln in range(start, end + 1):
                    marker = ">>>" if ln == err_line else "   "
                    print(marker + " " + str(ln) + ": " + prd_src.get(ln, "[N/A]")[:60])


def main():
    print("=" * 70)
    print("SAPLDMEE2 Source Code Comparison: DEV vs PRD")
    print("=" * 70)

    env_file = project_root / ".env"
    load_dotenv(env_file)

    comp = SourceComparator()

    dev_conn = comp.connect("SAP_DEV", "DEV")
    prd_conn = comp.connect("SAP_PRD", "PRD")

    if not dev_conn or not prd_conn:
        print("\n[ERROR] Cannot connect to both environments")
        return False

    try:
        # Read sources
        dev_source, dev_includes = comp.read_program_source(dev_conn, 'SAPLDMEE2', 'DEV')
        prd_source, prd_includes = comp.read_program_source(prd_conn, 'SAPLDMEE2', 'PRD')

        if dev_source is None or prd_source is None:
            print("\n[ERROR] Failed to read source")
            return False

        # Compare
        print("\n[*] Comparing...")

        if dev_source == prd_source:
            print("    [✓] IDENTICAL code")
            print("\n    Hypothesis: Error is in dependencies, not SAPLDMEE2 itself")
            print("    Check: LDMEE2F01 includes, interface changes, class versioning")
        else:
            print("    [!] DIFFERENT code")
            print("        DEV: " + str(len(dev_source)) + " lines")
            print("        PRD: " + str(len(prd_source)) + " lines")

            # Find first differences
            dev_keys = set(dev_source.keys())
            prd_keys = set(prd_source.keys())
            all_keys = sorted(dev_keys | prd_keys)

            diff_count = 0
            first_diffs = []

            for key in all_keys:
                if dev_source.get(key) != prd_source.get(key):
                    diff_count += 1
                    if len(first_diffs) < 5:
                        first_diffs.append(key)

            print("\n    First differences found at lines: " + ', '.join(str(d) for d in first_diffs))
            print("    Total difference count: " + str(diff_count))

        # Analyze reported error lines
        error_lines = [25, 30, 78, 538, 730, 2465, 2505, 2532, 2617, 2791]
        comp.analyze_error_lines(dev_source, prd_source, error_lines)

        # Unified diff
        print("\n[*] Generating diff...")
        dev_lines = [str(k).rjust(4) + ": " + v for k, v in sorted(dev_source.items())]
        prd_lines = [str(k).rjust(4) + ": " + v for k, v in sorted(prd_source.items())]

        diff = list(difflib.unified_diff(dev_lines, prd_lines, lineterm='', n=1))

        if diff:
            diff_lines = [l for l in diff if l.startswith('+') or l.startswith('-')]
            print("    Total diff lines: " + str(len(diff_lines)))

            if diff_lines:
                print("\n    First 10 diff lines:")
                for line in diff_lines[:10]:
                    print("      " + line)

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
