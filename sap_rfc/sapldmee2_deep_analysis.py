#!/usr/bin/env python3
"""
Advanced SAPLDMEE2 analysis using multiple RFC methods.
Tries different approaches to read source and diagnostics.
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

import json
from datetime import datetime


class SAPLDMEEAnalyzer:
    def __init__(self):
        self.env_prefix = None
        self.conn = None
        self.env_label = None

    def connect(self, prefix, label):
        self.env_prefix = prefix
        self.env_label = label
        try:
            self.conn = Connection(
                user=os.getenv(prefix + '_USER'),
                passwd=os.getenv(prefix + '_PASSWD'),
                ashost=os.getenv(prefix + '_ASHOST'),
                sysnr=os.getenv(prefix + '_SYSNR'),
                client=os.getenv(prefix + '_CLIENT'),
                lang=os.getenv(prefix + '_LANG', 'PT'),
            )
            print("[OK] Connected to " + label)
            return True
        except Exception as e:
            print("[FAIL] " + label + ": " + str(e))
            return False

    def try_read_table(self, table_name, options=None):
        """Try to read a table via RFC_READ_TABLE."""
        try:
            opts = options or []
            opt_list = [{'TEXT': o} for o in opts]

            result = self.conn.call(
                'RFC_READ_TABLE',
                QUERY_TABLE=table_name,
                OPTIONS=opt_list,
                ROWCOUNT=100,
            )

            count = len(result.get('DATA', []))
            print("  " + table_name + ": " + str(count) + " rows")
            return result
        except Exception as e:
            print("  " + table_name + ": " + str(e)[:60])
            return None

    def try_function_module(self, func_name, **params):
        """Try to call a function module."""
        try:
            result = self.conn.call(func_name, **params)
            print("  " + func_name + ": OK")
            return result
        except Exception as e:
            print("  " + func_name + ": " + str(e)[:60])
            return None

    def analyze(self, program_name):
        print("\n=== Analyzing " + program_name + " in " + self.env_label + " ===")

        # Strategy 1: TRDIR + timestamp
        print("\n[1] Program Directory (TRDIR)...")
        self.try_read_table('TRDIR', [program_name])

        # Strategy 2: E071 transport history
        print("\n[2] Transport History (E071)...")
        self.try_read_table('E071', [program_name])

        # Strategy 3: Try alternative source tables
        print("\n[3] Source Tables...")
        self.try_read_table('REPOTD')  # Program text
        self.try_read_table('PROGRAMM')  # Program directory (alternative)

        # Strategy 4: Try function modules for reading source
        print("\n[4] Function Modules...")
        self.try_function_module('RPY_PROGRAM_READ', PROGRAM_NAME=program_name)
        self.try_function_module('ABAP_SOURCE_READ', PROGRAM_NAME=program_name)

        # Strategy 5: Diagnostic/Compilation info
        print("\n[5] Compilation Info...")
        self.try_read_table('REPOSRC_DIAG')
        self.try_read_table('D_OBJ_VERSION')

        print("\n[Done]")


def main():
    print("=" * 60)
    print("SAPLDMEE2 Multi-Method Analysis")
    print("=" * 60)

    env_file = project_root / ".env"
    load_dotenv(env_file)

    # Analyze DEV
    dev_analyzer = SAPLDMEEAnalyzer()
    if dev_analyzer.connect("SAP_DEV", "DEV"):
        dev_analyzer.analyze("SAPLDMEE2")
        try:
            dev_analyzer.conn.close()
        except:
            pass

    # Analyze PRD
    prd_analyzer = SAPLDMEEAnalyzer()
    if prd_analyzer.connect("SAP_PRD", "PRD"):
        prd_analyzer.analyze("SAPLDMEE2")
        try:
            prd_analyzer.conn.close()
        except:
            pass

    print("\n" + "=" * 60)
    print("Analysis complete. Check output above for results.")
    print("If all methods failed: Check RFC authorization and table access.")
    print("=" * 60)


if __name__ == "__main__":
    main()
