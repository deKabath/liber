#!/usr/bin/env python3
"""
Download club logos from Airtable.

Usage:
  1. Install pyairtable: pip install pyairtable
  2. Set your Airtable Personal Access Token: export AIRTABLE_TOKEN="pat..."
  3. Run: python3 scripts/download_logos.py

The script reads logos from the Airtable "Table 3" and saves them to img/logos/.
"""
import os, re, sys, urllib.request

try:
    from pyairtable import Api
except ImportError:
    print("Please install pyairtable: pip install pyairtable")
    sys.exit(1)

TOKEN = os.environ.get("AIRTABLE_TOKEN")
if not TOKEN:
    print("Set AIRTABLE_TOKEN environment variable with your Airtable Personal Access Token")
    sys.exit(1)

BASE_ID = "apppVF94p7J29Hoz3"
TABLE_ID = "tbloTyVMyl3x4ZWEe"

# Ensure we're in the repo root
script_dir = os.path.dirname(os.path.abspath(__file__))
repo_root = os.path.dirname(script_dir)
logos_dir = os.path.join(repo_root, "img", "logos")
os.makedirs(logos_dir, exist_ok=True)

api = Api(TOKEN)
table = api.table(BASE_ID, TABLE_ID)
records = table.all()

downloaded = 0
skipped = 0

for record in records:
    fields = record["fields"]
    name = fields.get("Naam vereniging", "")
    if not name or name == "Gemeenten & Overheden":
        continue

    # Use "Attachments 2" first (logo field), fallback to "Attachments"
    attachments = fields.get("Attachments 2") or fields.get("Attachments") or []
    if not attachments:
        print(f"  SKIP (no logo): {name}")
        skipped += 1
        continue

    logo = attachments[0]
    url = logo["url"]
    filename_orig = logo.get("filename", "logo.png")
    ext = os.path.splitext(filename_orig)[1].lower() or ".png"

    safe_name = re.sub(r"[^a-zA-Z0-9]", "-", name.lower()).strip("-")
    safe_name = re.sub(r"-+", "-", safe_name)
    local_filename = f"{safe_name}{ext}"
    local_path = os.path.join(logos_dir, local_filename)

    try:
        urllib.request.urlretrieve(url, local_path)
        size = os.path.getsize(local_path)
        print(f"  OK: {name} -> {local_filename} ({size} bytes)")
        downloaded += 1
    except Exception as e:
        print(f"  FAIL: {name} - {e}")

print(f"\nDone! Downloaded: {downloaded}, Skipped: {skipped}")
