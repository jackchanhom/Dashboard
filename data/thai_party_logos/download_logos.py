import os
import re
import urllib.parse
import requests
from PIL import Image, ImageDraw
from io import BytesIO
import csv

BASE = "https://www.vote62.com/img/election69/party/"
OUT_DIR = "party_logos_circle"
SIZE = 512
SHEET_ID = "19cLkQfXtcwbnVFR6ilNd7ZeqRTUYU2jrmCCRsyFJ61w"
DISTRICT_GID = "0"
SHEET_URL = f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/export?format=csv&gid={DISTRICT_GID}"

HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)",
    "Referer": "https://www.vote62.com/",
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,*/*;q=0.8",
}

os.makedirs(OUT_DIR, exist_ok=True)

def circle_crop(img: Image.Image, size=512) -> Image.Image:
    img = img.convert("RGBA")
    w, h = img.size
    m = min(w, h)
    img = img.crop(((w-m)//2, (h-m)//2, (w+m)//2, (h+m)//2)).resize((size, size), Image.LANCZOS)

    mask = Image.new("L", (size, size), 0)
    ImageDraw.Draw(mask).ellipse((0, 0, size-1, size-1), fill=255)

    out = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    out.paste(img, (0, 0), mask)
    return out

def fetch_party_names_from_sheet():
    r = requests.get(SHEET_URL, headers=HEADERS, timeout=30)
    r.raise_for_status()
    # Force UTF-8 decoding from raw bytes
    text = r.content.decode('utf-8-sig')
    lines = text.splitlines()
    if not lines:
        return []
    reader = csv.DictReader(lines)
    # Look for the party column (accept common variations)
    party_col = None
    for h in reader.fieldnames or []:
        clean = h.strip().lower()
        # Check for exact matches
        if clean in ("พรรค 2569", "พรรค", "party"):
            party_col = h
            break
        # Check if contains "2569" (likely the party column)
        if "2569" in clean:
            party_col = h
            break
        # Check if contains "พรรค"
        if "พรรค" in clean:
            party_col = h
            break
    if not party_col:
        # Fallback: use 5th column (index 4) which is typically the party column
        if reader.fieldnames and len(reader.fieldnames) >= 5:
            party_col = reader.fieldnames[4]
            print(f"Using fallback column: {party_col}")
        else:
            raise ValueError(f"ไม่พบคอลัมน์ชื่อพรรคใน Google Sheet: {reader.fieldnames}")
    print(f"Using party column: {party_col}")
    parties = []
    for row in reader:
        name = (row.get(party_col) or "").strip()
        if name:
            parties.append(name)
    # unique preserve order
    seen = set()
    unique = []
    for p in parties:
        if p not in seen:
            seen.add(p)
            unique.append(p)
    return unique

def download_one(party_name: str):
    url = BASE + urllib.parse.quote(party_name, safe="") + ".png?t=1"
    r = requests.get(url, headers=HEADERS, timeout=30)
    r.raise_for_status()
    img = Image.open(BytesIO(r.content))
    out = circle_crop(img, SIZE)
    # Save with Thai party name
    filename = f"{party_name}.png"
    out.save(os.path.join(OUT_DIR, filename), "PNG")
    print("OK:", party_name)

def main():
    party_names = fetch_party_names_from_sheet()
    if not party_names:
        print("ไม่พบรายชื่อพรรคจาก Google Sheet")
        return
    print(f"Found {len(party_names)} parties.")
    for name in party_names:
        try:
            download_one(name)
        except Exception as e:
            print("FAIL:", name, e)
    print("Done:", OUT_DIR)

if __name__ == "__main__":
    main()
