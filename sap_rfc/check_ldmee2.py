#!/usr/bin/env python3
import os
import sys
from pathlib import Path
from dotenv import load_dotenv

project_root = Path(__file__).parent.parent
sys.path.insert(0, str(project_root))

try:
    from pyrfc import Connection
except ImportError:
    print("ERROR: pyrfc not available")
    sys.exit(1)


def connect_env(prefix):
    conn = Connection(
        user=os.getenv(prefix + '_USER'),
        passwd=os.getenv(prefix + '_PASSWD'),
        ashost=os.getenv(prefix + '_ASHOST'),
        sysnr=os.getenv(prefix + '_SYSNR'),
        client=os.getenv(prefix + '_CLIENT'),
        lang=os.getenv(prefix + '_LANG', 'PT'),
    )
    return conn


def read_trdir_info(conn, program_name, env_label):
    print("\n[*] Reading TRDIR for " + program_name + " in " + env_label + "...")

    try:
        result = conn.call(
            'RFC_READ_TABLE',
            QUERY_TABLE='TRDIR',
            FIELDS=[
                {'FIELDNAME': 'NAME'},
                {'FIELDNAME': 'UNAM'},
                {'FIELDNAME': 'UDAT'},
                {'FIELDNAME': 'UTIME'},
                {'FIELDNAME': 'VARCL'},
                {'FIELDNAME': 'FIXPT'},
            ],
            OPTIONS=[
                {'TEXT': 'NAME = ' + "'" + program_name + "'"}
            ],
            ROWCOUNT=1,
        )

        if result.get('DATA'):
            wa = str(result['DATA'][0].get('WA', ''))
            parts = wa.split('|')
            print("    [OK] Program found")
            if len(parts) > 2:
                print("         Date: " + parts[2])
                print("         Time: " + (parts[3] if len(parts) > 3 else "?"))
            return True
        else:
            print("    [!] Program not found")
            return False

    except Exception as e:
        print("    [ERROR] " + str(e))
        return False


def read_e071_changes(conn, program_name, env_label):
    print("\n[*] Checking E071 (transport history) for " + program_name + " in " + env_label + "...")

    try:
        result = conn.call(
            'RFC_READ_TABLE',
            QUERY_TABLE='E071',
            FIELDS=[
                {'FIELDNAME': 'TRKORR'},
                {'FIELDNAME': 'STRKORR'},
                {'FIELDNAME': 'OBJECT'},
                {'FIELDNAME': 'OBJ_NAME'},
                {'FIELDNAME': 'OPERATION'},
            ],
            OPTIONS=[
                {'TEXT': 'OBJ_NAME = ' + "'" + program_name + "'"},
                {'TEXT': 'OBJECT = ' + "'" + 'PROG' + "'"}
            ],
            ROWCOUNT=50,
        )

        requests = []
        for row in result.get('DATA', []):
            wa = str(row.get('WA', ''))
            parts = wa.split('|')
            if len(parts) >= 2:
                req = parts[0].strip()
                if req:
                    requests.append(req)

        if requests:
            print("    [OK] Found in transports: " + ', '.join(set(requests)))
            return True
        else:
            print("    [!] Not found in recent transports")
            return False

    except Exception as e:
        print("    [ERROR] " + str(e))
        return False


def main():
    print("="*60)
    print("SAPLDMEE2 Analysis: DEV vs PRD")
    print("="*60)

    env_file = project_root / ".env"
    load_dotenv(env_file)

    print("\n[1] Connecting to SAP environments...")

    try:
        dev_conn = connect_env("SAP_DEV")
        print("    [OK] SAP_DEV connected")
    except Exception as e:
        print(f"    [FAIL] SAP_DEV: {e}")
        return False

    try:
        prd_conn = connect_env("SAP_PRD")
        print("    [OK] SAP_PRD connected")
    except Exception as e:
        print(f"    [FAIL] SAP_PRD: {e}")
        return False

    try:
        print("\n[2] Checking program status...")
        dev_found = read_trdir_info(dev_conn, 'SAPLDMEE2', 'DEV')
        prd_found = read_trdir_info(prd_conn, 'SAPLDMEE2', 'PRD')

        if not dev_found or not prd_found:
            print("\n[ERROR] Program not found in one or both environments")
            return False

        print("\n[3] Checking transport history...")
        read_e071_changes(dev_conn, 'SAPLDMEE2', 'DEV')
        read_e071_changes(prd_conn, 'SAPLDMEE2', 'PRD')

        print("\n[4] ANALYSIS RESULTS:")
        print("    NOTE: Full source code access via RFC_READ_TABLE/REPOSRC")
        print("    is not available in this environment.")
        print("    Recommend using transaction SE38 or source control to compare.")
        print("\n[5] Analysis complete")
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
