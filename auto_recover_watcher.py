"""
auto_recover_watcher.py
-----------------------
Run this in a third terminal alongside the collector and uploader.
It runs recover_excel.py every 20 minutes to fix any Excel corruption
before it can crash the other scripts.

Usage:  python auto_recover_watcher.py
Stop:   Ctrl+C
"""

import subprocess
import sys
import time
from datetime import datetime

INTERVAL_SECONDS = 20 * 60  # 20 minutes


def run_recovery():
    print(f"\n[{datetime.now():%H:%M:%S}] Running recover_excel.py …")
    try:
        result = subprocess.run(
            [sys.executable, "recover_excel.py"],
            capture_output=True,
            text=True,
        )
        if result.stdout.strip():
            print(result.stdout.strip())
        if result.returncode != 0:
            print(f"  recover_excel.py exited with code {result.returncode}")
            if result.stderr.strip():
                print(f"  stderr: {result.stderr.strip()}")
        else:
            print(f"  Done — file is healthy.")
    except Exception as e:
        print(f"  Could not run recover_excel.py: {e}")


def main():
    print("Auto-Recover Watcher started.")
    print(f"  Runs recover_excel.py every {INTERVAL_SECONDS // 60} minutes.")
    print("  Press Ctrl+C to stop.\n")

    # Run once immediately so the file is checked right away.
    run_recovery()

    while True:
        time.sleep(INTERVAL_SECONDS)
        run_recovery()


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\nWatcher stopped.")
