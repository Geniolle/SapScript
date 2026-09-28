#!/usr/bin/env python3
"""Test reading SAPLDMEE2 source with different RFC calls."""

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


def main():
    env_file = project_root / ".env"
    load_dotenv(env_file)

    print("Connecting to SAP_DEV...")
    conn = Connection(
        user=os.getenv('SAP_DEV_USER'),
        passwd=os.getenv('SAP_DEV_PASSWD'),
        ashost=os.getenv('SAP_DEV_ASHOST'),
        sysnr=os.getenv('SAP_DEV_SYSNR'),
        client=os.getenv('SAP_DEV_CLIENT'),
        lang=os.getenv('SAP_DEV_LANG', 'PT'),
    )

    print("\nTrying RPY_PROGRAM_READ with different parameters...")

    # Try 1
    print("\n[1] Basic call:")
    r = conn.call('RPY_PROGRAM_READ', PROGRAM_NAME='SAPLDMEE2', LANGUAGE='PT')
    print("    Keys: " + ', '.join(r.keys()))
    print("    SOURCE: " + str(len(r.get('SOURCE', []))) + " items")
    print("    SOURCE_EXTENDED: " + str(type(r.get('SOURCE_EXTENDED')).__name__) + ", " + str(len(r.get('SOURCE_EXTENDED', []))) + " items")

    # Try 2
    print("\n[2] With WITH_LOWERCASE:")
    r = conn.call('RPY_PROGRAM_READ', PROGRAM_NAME='SAPLDMEE2', LANGUAGE='PT', WITH_LOWERCASE='X')
    print("    SOURCE: " + str(len(r.get('SOURCE', []))) + " items")
    print("    SOURCE_EXTENDED: " + str(len(r.get('SOURCE_EXTENDED', []))) + " items")

    # Try 3
    print("\n[3] With READ_LATEST_VERSION:")
    r = conn.call('RPY_PROGRAM_READ', PROGRAM_NAME='SAPLDMEE2', LANGUAGE='PT', READ_LATEST_VERSION='X')
    print("    SOURCE: " + str(len(r.get('SOURCE', []))) + " items")
    print("    SOURCE_EXTENDED: " + str(len(r.get('SOURCE_EXTENDED', []))) + " items")

    # Try 4 - Try reading LDMEE2F01 directly
    print("\n[4] Reading LDMEE2F01 (include) directly:")
    r = conn.call('RPY_PROGRAM_READ', PROGRAM_NAME='LDMEE2F01', LANGUAGE='PT')
    print("    SOURCE: " + str(len(r.get('SOURCE', []))) + " items")
    print("    SOURCE_EXTENDED: " + str(len(r.get('SOURCE_EXTENDED', []))) + " items")

    # Now let's inspect SOURCE_EXTENDED structure
    print("\n[5] Inspecting SOURCE_EXTENDED structure:")
    r = conn.call('RPY_PROGRAM_READ', PROGRAM_NAME='SAPLDMEE2', LANGUAGE='PT')
    se = r.get('SOURCE_EXTENDED', [])

    if se:
        if isinstance(se, list) and len(se) > 0:
            print("    List with " + str(len(se)) + " items")
            item = se[0]
            if isinstance(item, dict):
                print("    First item keys: " + ', '.join(item.keys()))
                print("    First item content: " + str(item)[:200])
            else:
                print("    First item type: " + type(item).__name__)
                print("    First item: " + str(item)[:200])
        else:
            print("    Type: " + type(se).__name__)
            print("    Content: " + str(se)[:200])

    # Try SE16 RFC or similar
    print("\n[6] Trying SE16 RFC methods:")
    try:
        r = conn.call('SE16N_GET_DATA', TABLE='SAPLDMEE2')
        print("    SE16N_GET_DATA: Found data")
    except Exception as e:
        print("    SE16N_GET_DATA: " + str(e)[:80])

    conn.close()


if __name__ == "__main__":
    main()
