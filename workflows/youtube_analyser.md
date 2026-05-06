# Workflow: YouTube Space Analyser

## Objective
Scrape YouTube data for a given topic, analyse it with OpenAI, and produce a professional
PowerPoint slide deck delivered via email.

## Required Inputs
- `YOUTUBE_API_KEY` in `.env`
- `OPENAI_API_KEY` in `.env`
- `GMAIL_SENDER_EMAIL` + `GMAIL_APP_PASSWORD` in `.env` (for email delivery)
- A topic string (e.g. `"kids content"`)

## Run Command
```
.venv\Scripts\python tools/run_analysis.py --topic "kids content"
```
The script prompts for a recipient email at the end.

## Tool Sequence

| Step | Script | Output |
|------|--------|--------|
| 1 | `fetch_videos.py --topic "..."` | `.tmp/data/videos.json` |
| 2 | `fetch_channels.py` | `.tmp/data/channels.json` |
| 3 | `fetch_transcripts.py` | `.tmp/data/transcripts.json` |
| 4 | `fetch_thumbnails.py` | `.tmp/thumbnails/*.jpg` |
| 5 | `analyze_content.py` | `.tmp/data/insights.json` |
| 6 | `generate_charts.py` | `.tmp/charts/*.png` |
| 7 | `generate_slides.py` | `output/<topic>_<date>.pptx` |
| 8 | `send_email.py` | Email with PPTX attachment |

## API Quota Budget (per run)
- YouTube Data API: ~570 units (< 6% of 10,000/day free quota)
- Claude Sonnet 4.6: 3 text calls (prompt-cached system prompt)
- Claude Haiku 4.5: 4 vision calls (thumbnail batches of 5)
- youtube-transcript-api: free, no key required

## Slide Deck Contents (12 slides)
1. Title
2. Executive Summary
3. Top 10 Videos
4. Top Channels
5. Engagement Deep Dive
6. Content Themes
7. Title & Hook Patterns
8. Thumbnail Visual Trends
9. Upload Cadence
10. Content Gaps & Opportunities
11. Recommendations
12. Methodology

## Edge Cases & Known Quirks
- Many kids content videos have no English captions (animation/music). Transcripts are skipped
  gracefully; the analysis still runs on titles, descriptions, and thumbnails.
- YouTube API returns `subscriberCount` as hidden for some channels — those show 0 subscribers.
- If Claude returns JSON wrapped in markdown fences, `parse_json()` in `analyze_content.py`
  strips them automatically.
- Thumbnails at `maxres` quality may return 404 for older videos; the script falls back to `high`.
- Gmail App Passwords require 2FA to be enabled on the sender Gmail account.

## Re-running a Topic
All intermediate files in `.tmp/` are overwritten on each run. To keep a previous run's data,
copy `.tmp/data/` and `output/` to a separate folder before re-running.
