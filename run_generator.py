#!/usr/bin/env python3
"""
Runner script - wykonuje generator arkusza Excel
"""
import subprocess
import sys

if __name__ == "__main__":
    print("Generowanie arkusza Excel do oceny kandydatów...")
    result = subprocess.run([sys.executable, "Ocena_Kandydatow.py"], capture_output=True, text=True)
    print(result.stdout)
    if result.stderr:
        print("Błędy:", result.stderr)
    sys.exit(result.returncode)
