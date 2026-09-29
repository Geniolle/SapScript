#!/usr/bin/env python3
"""Debug RPY_PROGRAM_READ response structure."""

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


def connect_env(prefix):
    return Connection(
        user=os.getenv(prefix + '_USER'),
        passwd=os.getenv(prefix + '_PASSWD'),
        ashost=os.getenv(prefix + '_ASHOST'),
        sysnr=os.getenv(prefix + '_SYSNR'),
        client=os.getenv(prefix + '_CLIENT'),
        lang=os.getenv(prefix + '_LANG', 'PT'),
    )


def main():
    env_file = project_root / ".env"
    load_dotenv(env_file)

    print("Connecting to SAP_DEV...")
    conn = connect_env("SAP_DEV")

    print("\nCalling RPY_PROGRAM_READ for SAPLDMEE2...")

    print("\n[Attempt 1: Basic call]")
    result1 = conn.call(
        'RPY_PROGRAM_READ',
        PROGRAM_NAME='SAPLDMEE2',
        LANGUAGE='PT',
    )
    print(f"SOURCE length: {len(result1.get('SOURCE', []))}")
    print(f"SOURCE_EXTENDED: {type(result1.get('SOURCE_EXTENDED'))}")

    print("\n[Attempt 2: With ONLY_SOURCE]")
    result2 = conn.call(
        'RPY_PROGRAM_READ',
        PROGRAM_NAME='SAPLDMEE2',
        LANGUAGE='PT',
        ONLY_SOURCE='X',
    )
    print(f"SOURCE length: {len(result2.get('SOURCE', []))}")

    print("\n[Attempt 3: With READ_LATEST_VERSION]")
    result3 = conn.call(
        'RPY_PROGRAM_READ',
        PROGRAM_NAME='SAPLDMEE2',
        LANGUAGE='PT',
        READ_LATEST_VERSION='X',
    )
    print(f"SOURCE length: {len(result3.get('SOURCE', []))}")

    print("\n[Attempt 4: Trying LDMEE2F01 include]")
    result4 = conn.call(
        'RPY_PROGRAM_READ',
        PROGRAM_NAME='LDMEE2F01',
        LANGUAGE='PT',
        READ_LATEST_VERSION='X',
    )
    print(f"SOURCE length: {len(result4.get('SOURCE', []))}")

    result = result1

    print("\n[RESPONSE KEYS]")
    print("Keys: " + ', '.join(sorted(result.keys())))

    print("\n[CHECKING DATA SOURCES]")

    for key in ['SOURCE', 'SOURCE_TAB', 'SOURCE_TABLE', 'SOURCE_LINES', 'LINES']:
        if key in result:
            data = result[key]
            print(f"\n{key}: {type(data).__name__}")
            if isinstance(data, (list, dict)):
                print(f"  Length: {len(data)}")
                if isinstance(data, list) and len(data) > 0:
                    print(f"  First item type: {type(data[0]).__name__}")
                    if isinstance(data[0], dict):
                        print(f"  First item keys: {list(data[0].keys())[:5]}")
                elif isinstance(data, dict):
                    print(f"  Keys: {list(data.keys())[:5]}")
            if data:
                print(f"  Sample: {str(data)[:100]}")

    print("\n[FULL RESPONSE STRUCTURE]")

    # Pretty print with truncation
    def truncate_deep(obj, max_depth=2, current_depth=0):
        if current_depth >= max_depth:
            return f"<{type(obj).__name__}>"

        if isinstance(obj, dict):
            return {k: truncate_deep(v, max_depth, current_depth+1) for k, v in list(obj.items())[:3]}
        elif isinstance(obj, (list, tuple)):
            return [truncate_deep(v, max_depth, current_depth+1) for v in obj[:2]]
        else:
            s = str(obj)
            return s[:50] if len(s) > 50 else s

    print(json.dumps(truncate_deep(result, max_depth=3), indent=2, default=str))

    conn.close()


if __name__ == "__main__":
    main()
