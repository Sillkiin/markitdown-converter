#!/usr/bin/env python3
"""Run every bundled test suite and print a combined summary.
Usage:  python tests/run_all.py
Exit code is non-zero if any suite fails."""
import os, sys, subprocess

HERE = os.path.dirname(os.path.abspath(__file__))
SUITES = ["selftest.py", "test_suite.py", "test_adversarial.py", "verify_native.py"]

def main():
    failed = []
    for s in SUITES:
        path = os.path.join(HERE, s)
        print("\n" + "=" * 30 + " " + s + " " + "=" * 30)
        rc = subprocess.run([sys.executable, path]).returncode
        print("--> %s: %s" % (s, "OK" if rc == 0 else "FAILED (rc=%d)" % rc))
        if rc != 0:
            failed.append(s)
    print("\n" + "#" * 70)
    if failed:
        print("RESULT: FAILED suites: " + ", ".join(failed))
        sys.exit(1)
    print("RESULT: ALL SUITES PASSED")

if __name__ == "__main__":
    main()
