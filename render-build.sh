#!/usr/bin/env bash
# Render.com build script — installs both Python and Node dependencies
set -e

echo "=== Installing Python dependencies ==="
pip install -r requirements.txt

echo "=== Installing Node dependencies (for SheetJS) ==="
npm install

echo "=== Build complete ==="
