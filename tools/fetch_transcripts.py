import json
from pathlib import Path
from youtube_transcript_api import YouTubeTranscriptApi, TranscriptsDisabled, NoTranscriptFound

DATA_DIR = Path(".tmp/data")
MAX_CHARS = 3000


def get_transcript(video_id):
    try:
        entries = YouTubeTranscriptApi.get_transcript(video_id, languages=["en", "en-US"])
        return " ".join(e["text"] for e in entries)[:MAX_CHARS]
    except (TranscriptsDisabled, NoTranscriptFound):
        return None
    except Exception:
        return None


def main():
    DATA_DIR.mkdir(parents=True, exist_ok=True)

    with open(DATA_DIR / "videos.json", encoding="utf-8") as f:
        videos = json.load(f)["videos"]

    top = videos[:15]
    print(f"[1/1] Fetching transcripts for top {len(top)} videos...")
    results = []
    for i, v in enumerate(top):
        text = get_transcript(v["video_id"])
        status = "✓" if text else "✗"
        print(f"      [{i+1:02d}/{len(top)}] {status} {v['title'][:55]}")
        if text:
            results.append({
                "video_id": v["video_id"],
                "title": v["title"],
                "views": v["views"],
                "transcript": text,
            })

    out = DATA_DIR / "transcripts.json"
    with open(out, "w", encoding="utf-8") as f:
        json.dump({"transcripts": results}, f, indent=2, ensure_ascii=False)
    print(f"      {len(results)} transcripts saved → {out}")


if __name__ == "__main__":
    main()
