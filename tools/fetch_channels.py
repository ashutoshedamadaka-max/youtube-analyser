import os
import json
from pathlib import Path
from dotenv import load_dotenv
from googleapiclient.discovery import build

load_dotenv()

DATA_DIR = Path(".tmp/data")


def fetch_channel_stats(youtube, channel_ids):
    channels = []
    for i in range(0, len(channel_ids), 50):
        batch = channel_ids[i : i + 50]
        resp = youtube.channels().list(
            id=",".join(batch), part="snippet,statistics"
        ).execute()
        for item in resp.get("items", []):
            s = item.get("snippet", {})
            st = item.get("statistics", {})
            channels.append({
                "channel_id": item["id"],
                "title": s.get("title", ""),
                "description": s.get("description", "")[:300],
                "subscribers": int(st.get("subscriberCount", 0)),
                "total_views": int(st.get("viewCount", 0)),
                "video_count": int(st.get("videoCount", 0)),
            })
    return sorted(channels, key=lambda c: c["subscribers"], reverse=True)


def main():
    DATA_DIR.mkdir(parents=True, exist_ok=True)

    with open(DATA_DIR / "videos.json", encoding="utf-8") as f:
        videos = json.load(f)["videos"]

    channel_ids = list({v["channel_id"] for v in videos[:100]})[:25]
    youtube = build("youtube", "v3", developerKey=os.getenv("YOUTUBE_API_KEY"))

    print(f"[1/1] Fetching {len(channel_ids)} channels...")
    channels = fetch_channel_stats(youtube, channel_ids)
    print(f"      {len(channels)} channels retrieved")

    out = DATA_DIR / "channels.json"
    with open(out, "w", encoding="utf-8") as f:
        json.dump({"channels": channels}, f, indent=2, ensure_ascii=False)
    print(f"      Saved → {out}")


if __name__ == "__main__":
    main()
