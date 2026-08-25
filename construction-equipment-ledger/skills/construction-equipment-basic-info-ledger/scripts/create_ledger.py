#!/usr/bin/env python3
"""Create a new blank 건설기계 기초자료 관리대장 workbook.

Usage: python create_ledger.py <output_path.xlsx>
"""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from ledger_lib import create_ledger

if __name__ == "__main__":
    if len(sys.argv) != 2:
        print("Usage: python create_ledger.py <output_path.xlsx>")
        sys.exit(1)
    path = create_ledger(sys.argv[1])
    print(f"created {path}")
