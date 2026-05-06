import json
import requests
from pathlib import Path

DATA_DIR = Path(".tmp/data")
THUMB_DIR = Path(".tmp/thumbnails")


def download(video_id, url):
    if not url:
        return None
    try:
        r = requests.get(url, timeout=10)
        r.raise_for_status()
        path = THUMB_DIR / f"{video_id}.jpg"
        path.write_bytes(r.content)
        return str(path)
    except Exception:
        return None


def main():
    THUMB_DIR.mkdir(parents=True, exist_ok=True)

    with open(DATA_DIR / "videos.json", encoding="utf-8") as f:
        videos = json.load(f)["videos"]

    top = videos[:20]
    print(f"[1/1] Downloading thumbnails for top {len(top)} videos...")
    results = []
    for i, v in enumerate(top):
        path = download(v["video_id"], v.get("thumbnail_url"))
        status = "✓" if path else "✗"
        print(f"      [{i+1:02d}/{len(top)}] {status} {v['title'][:55]}")
        if path:
            results.append({"video_id": v["video_id"], "title": v["title"], "path": path})

    out = DATA_DIR / "thumbnails.json"
    with open(out, "w", encoding="utf-8") as f:
        json.dump({"thumbnails": results}, f, indent=2)
    print(f"      {len(results)} thumbnails saved → {THUMB_DIR}")


if __name__ == "__main__":
    main()
