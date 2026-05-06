import os
import json
import argparse
from pathlib import Path
from dotenv import load_dotenv
from googleapiclient.discovery import build

load_dotenv()

DATA_DIR = Path(".tmp/data")


def search_video_ids(youtube, topic):
    queries = [
        topic,
        f"{topic} for kids",
        f"best {topic} videos",
        f"{topic} learning",
    ]
    ids = set()
    for q in queries:
        resp = youtube.search().list(
            q=q, part="id", type="video", maxResults=50, order="viewCount"
        ).execute()
        for item in resp.get("items", []):
            ids.add(item["id"]["videoId"])
    return list(ids)


def fetch_video_details(youtube, video_ids):
    videos = []
    for i in range(0, len(video_ids), 50):
        batch = video_ids[i : i + 50]
        resp = youtube.videos().list(
            id=",".join(batch), part="snippet,statistics,contentDetails"
        ).execute()
        for item in resp.get("items", []):
            s = item.get("snippet", {})
            st = item.get("statistics", {})
            cd = item.get("contentDetails", {})
            thumbnails = s.get("thumbnails", {})
            thumb_url = (
                thumbnails.get("maxres", thumbnails.get("high", thumbnails.get("default", {}))).get("url", "")
            )
            videos.append({
                "video_id": item["id"],
                "title": s.get("title", ""),
                "description": s.get("description", "")[:500],
                "channel_id": s.get("channelId", ""),
                "channel_title": s.get("channelTitle", ""),
                "published_at": s.get("publishedAt", ""),
                "tags": s.get("tags", [])[:10],
                "thumbnail_url": thumb_url,
                "views": int(st.get("viewCount", 0)),
                "likes": int(st.get("likeCount", 0)),
                "comments": int(st.get("commentCount", 0)),
                "duration": cd.get("duration", ""),
            })
    return sorted(videos, key=lambda v: v["views"], reverse=True)


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--topic", required=True)
    args = parser.parse_args()

    DATA_DIR.mkdir(parents=True, exist_ok=True)
    youtube = build("youtube", "v3", developerKey=os.getenv("YOUTUBE_API_KEY"))

    print(f"[1/2] Searching: '{args.topic}'")
    ids = search_video_ids(youtube, args.topic)
    print(f"      {len(ids)} unique video IDs found")

    print("[2/2] Fetching video details...")
    videos = fetch_video_details(youtube, ids[:150])
    print(f"      {len(videos)} videos retrieved")

    out = DATA_DIR / "videos.json"
    with open(out, "w", encoding="utf-8") as f:
        json.dump({"topic": args.topic, "videos": videos}, f, indent=2, ensure_ascii=False)
    print(f"      Saved → {out}")


if __name__ == "__main__":
    main()
