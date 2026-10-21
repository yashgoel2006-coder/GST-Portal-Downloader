#!/bin/bash
cd "$(dirname "$0")"
command -v python3 >/dev/null || { echo "Install Python 3 from https://www.python.org/downloads/"; read -p "Enter..."; exit 1; }
python3 start.py
