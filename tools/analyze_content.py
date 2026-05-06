import os
import json
import base64
import re
from pathlib import Path
from dotenv import load_dotenv
from openai import OpenAI

load_dotenv()

DATA_DIR  = Path(".tmp/data")
THUMB_DIR = Path(".tmp/thumbnails")

client = OpenAI(api_key=os.getenv("OPENAI_API_KEY"))


def parse_json(text):
    text = text.strip()
    text = re.sub(r"^```(?:json)?\s*", "", text)
    text = re.sub(r"\s*```$", "", text)
    return json.loads(text)


# ── Main text analysis (one comprehensive call) ───────────────────────────────

def analyze_all_text(videos, channels, transcripts, topic):
    n_videos   = len(videos)
    n_channels = len(channels)

    # Rich channel context
    total_subs = sum(c["subscribers"] for c in channels)
    channel_rows = "\n".join(
        f"  {i+1}. {c['title']}: {c['subscribers']:,} subs | "
        f"{c['total_views']:,} total views | {c['video_count']} videos"
        for i, c in enumerate(channels)
    )

    # Top 30 videos with full metrics
    top30_rows = "\n".join(
        f"  {i+1}. '{v['title']}' | {v['views']:,} views | "
        f"{v['likes']:,} likes | {v['comments']:,} comments | "
        f"channel: {v['channel_title']} | published: {v['published_at'][:10]}"
        for i, v in enumerate(videos[:30])
    )

    # All titles with view counts for pattern analysis
    all_titles = "\n".join(
        f"  {i+1}. [{v['views']:,}] {v['title']} | tags: {', '.join(v.get('tags',[])[:5])}"
        for i, v in enumerate(videos[:150])
    )

    # Transcript section
    transcript_block = ""
    if transcripts:
        transcript_block = "\nTRANSCRIPT EXCERPTS (top videos):\n" + "\n---\n".join(
            f"  '{t['title']}' ({t['views']:,} views):\n  {t['transcript'][:600]}"
            for t in transcripts[:5]
        )

    prompt = f"""You are a YouTube content strategist analyzing the "{topic}" space.
You have data on {n_channels} channels ({total_subs:,} combined subscribers) and {n_videos} videos.

Analyze this data and return a JSON object. Be SPECIFIC — cite channel names, exact numbers, percentages, and concrete examples. Never be generic.

CHANNEL DATA:
{channel_rows}

TOP 30 VIDEOS (by views):
{top30_rows}

ALL {n_videos} VIDEO TITLES WITH VIEW COUNTS:
{all_titles}
{transcript_block}

Return this exact JSON (no markdown, no fences):
{{
  "executive_summary": {{
    "space_overview": "2-3 sentences. Include: combined subscriber count ({total_subs:,}), the dominant player's market share, and the single most important trend right now. Use specific numbers.",
    "executive_summary_bullets": [
      "Bullet 1: Market share insight with exact %. E.g. 'CoComelon captures X% of total views despite being 1 of {n_channels} channels, creating a winner-take-all dynamic...'",
      "Bullet 2: Engagement or format trend with specific numbers",
      "Bullet 3: Content gap or opportunity with market size",
      "Bullet 4: A contrarian insight that goes against conventional wisdom in this space"
    ],
    "best_posting_insight": "Analyze upload dates of top vs bottom performers. What day patterns appear? How does upload frequency of top 5 channels compare to channels ranked 20-25? Cite specific numbers.",
    "content_gaps": [
      {{
        "gap": "Specific gap citing exact counts. E.g. 'Only X of {n_channels} channels produce shorts-first content despite Shorts driving 3x engagement...'",
        "opportunity": "Concrete action with audience size and reference to a proven format from the data"
      }},
      {{
        "gap": "Second specific gap with exact numbers",
        "opportunity": "Concrete action a solo creator with $500/month could execute"
      }},
      {{
        "gap": "Third specific gap",
        "opportunity": "Concrete action with projected reach based on comparable channels"
      }}
    ],
    "recommendations": [
      {{
        "title": "5 words max",
        "description": "Specific tactic referencing actual channels or video titles from the data that prove this works",
        "rationale": "The exact metric from the data that supports this recommendation"
      }},
      {{
        "title": "5 words max",
        "description": "Specific tactic with implementation steps a solo creator can follow",
        "rationale": "Data-backed reason with specific numbers"
      }},
      {{
        "title": "5 words max",
        "description": "Contrarian or non-obvious tactic backed by the data",
        "rationale": "The data point that makes this counter-intuitive but effective"
      }},
      {{
        "title": "5 words max",
        "description": "Format or platform-specific tactic referencing specific video examples",
        "rationale": "Engagement rate or view multiplier from the data"
      }},
      {{
        "title": "5 words max",
        "description": "Long-term positioning recommendation based on competitive gaps",
        "rationale": "Market concentration data supporting this positioning"
      }}
    ]
  }},
  "title_analysis": {{
    "dominant_content_themes": [
      {{"theme": "Exact theme name", "percentage": 28, "avg_views": "Actual avg like 2.1B", "top_example": "Exact video title from the data"}},
      {{"theme": "Theme 2", "percentage": 22, "avg_views": "...", "top_example": "..."}},
      {{"theme": "Theme 3", "percentage": 18, "avg_views": "...", "top_example": "..."}},
      {{"theme": "Theme 4", "percentage": 15, "avg_views": "...", "top_example": "..."}},
      {{"theme": "Theme 5", "percentage": 17, "avg_views": "...", "top_example": "..."}}
    ],
    "top_hook_patterns": [
      "Pattern with multiplier: e.g. 'Named character titles (Baby Shark, CoComelon) average 4.2x more views than generic nursery rhyme titles'",
      "Pattern 2 with specific comparison",
      "Pattern 3 with view count evidence",
      "Pattern 4 with engagement data",
      "Pattern 5 — include one contrarian finding"
    ],
    "high_performing_keywords": ["keyword1", "keyword2", "keyword3", "keyword4", "keyword5", "keyword6", "keyword7", "keyword8", "keyword9", "keyword10"],
    "view_correlation_insights": "What specific title characteristics appear in videos with 2x+ average views? What patterns consistently underperform? Include actual view count comparisons.",
    "overused_patterns": ["Specific oversaturated pattern 1", "Pattern 2", "Pattern 3"]
  }},
  "competitive_analysis": {{
    "market_concentration": "What % of total views do the top 3 channels capture? Top 5? Is this space concentrated (winner-take-all) or fragmented (room for new entrants)?",
    "channel_archetypes": [
      {{
        "archetype": "The Mega-Brand",
        "channels": ["Actual channel name from data"],
        "strategy": "Their specific content and growth strategy based on the data",
        "weakness": "Their specific vulnerability a new creator could exploit"
      }},
      {{
        "archetype": "The Niche Specialist",
        "channels": ["Channel name"],
        "strategy": "Their focused approach",
        "weakness": "Their limitation"
      }},
      {{
        "archetype": "The Rising Challenger",
        "channels": ["Channel name"],
        "strategy": "What's working for them",
        "weakness": "Where they're vulnerable"
      }}
    ],
    "rising_vs_established": "Based on video count vs views ratios, which channels are punching above their weight? Which established channels show signs of stagnation? Be specific with channel names and metrics.",
    "differentiation_map": "How do the top 5 channels each carve out distinct territory? What's each one's defensible angle? Where is the unclaimed white space?"
  }}
}}

CRITICAL RULES:
- Every insight MUST include a specific number, percentage, or data comparison from the actual data provided
- Never use vague language like 'various', 'diverse', 'multiple' — say the exact count
- Never say 'popular' without saying HOW popular (views, subs, engagement rate)
- Compare channels to each other, not just describe them individually
- Every recommendation must be executable by a solo creator with $500/month budget
- Include at least one contrarian insight that challenges conventional wisdom in this space
- avg_views in dominant_content_themes should be calculated from actual video data"""

    resp = client.chat.completions.create(
        model="gpt-4o",
        max_tokens=4000,
        messages=[
            {"role": "system", "content": "You are a YouTube content strategy analyst. Return only valid JSON with no markdown fences or extra text."},
            {"role": "user", "content": prompt},
        ],
    )
    return parse_json(resp.choices[0].message.content)


# ── Thumbnail vision analysis (kept separate — different model) ───────────────

def analyze_thumbnails(thumbnails):
    if not thumbnails:
        return {"dominant_visual_style": "", "color_palette_trends": [],
                "face_character_usage": "", "text_overlay_patterns": [],
                "emotional_tone": "", "design_recommendations": []}

    all_obs = []
    for i in range(0, min(len(thumbnails), 20), 5):
        batch = thumbnails[i : i + 5]
        content = []
        for t in batch:
            p = Path(t["path"])
            if not p.exists():
                continue
            img_b64 = base64.standard_b64encode(p.read_bytes()).decode()
            content.append({
                "type": "image_url",
                "image_url": {"url": f"data:image/jpeg;base64,{img_b64}", "detail": "low"},
            })
            content.append({"type": "text", "text": f"Video: {t['title'][:60]}"})
        if not content:
            continue
        content.append({
            "type": "text",
            "text": (
                "Analyze each thumbnail. For each note: dominant colors, "
                "faces/characters present, text overlay (yes/no and style), emotional tone. "
                'JSON: {"observations":["obs1","obs2",...]}'
            ),
        })
        resp = client.chat.completions.create(
            model="gpt-4o-mini",
            max_tokens=600,
            messages=[{"role": "user", "content": content}],
        )
        try:
            raw = parse_json(resp.choices[0].message.content).get("observations", [])
            all_obs.extend(o if isinstance(o, str) else json.dumps(o) for o in raw)
        except Exception:
            all_obs.append(resp.choices[0].message.content[:200])

    if not all_obs:
        return {"dominant_visual_style": "No thumbnails analyzed", "color_palette_trends": [],
                "face_character_usage": "", "text_overlay_patterns": [],
                "emotional_tone": "", "design_recommendations": []}

    synth = client.chat.completions.create(
        model="gpt-4o",
        max_tokens=900,
        messages=[
            {"role": "system", "content": "You are a visual design analyst. Return only valid JSON."},
            {"role": "user", "content": (
                "Synthesize these thumbnail observations into key visual trends. "
                "Be specific — cite percentages where patterns are clear.\n\n"
                + "\n".join(all_obs)
                + '\n\nReturn JSON:\n'
                '{"dominant_visual_style":"...",'
                '"color_palette_trends":["Specific: e.g. Yellow/red combo in 7 of top 10 thumbnails"],'
                '"face_character_usage":"What % use faces vs characters vs neither?",'
                '"text_overlay_patterns":["Pattern with frequency"],'
                '"emotional_tone":"...",'
                '"design_recommendations":["Actionable rec citing specific evidence from the thumbnails"]}'
            )},
        ],
    )
    return parse_json(synth.choices[0].message.content)


# ── Main ──────────────────────────────────────────────────────────────────────

def main():
    DATA_DIR.mkdir(parents=True, exist_ok=True)

    with open(DATA_DIR / "videos.json",      encoding="utf-8") as f: vdata       = json.load(f)
    with open(DATA_DIR / "channels.json",    encoding="utf-8") as f: channels    = json.load(f)["channels"]
    with open(DATA_DIR / "transcripts.json", encoding="utf-8") as f: transcripts = json.load(f)["transcripts"]
    with open(DATA_DIR / "thumbnails.json",  encoding="utf-8") as f: thumbnails  = json.load(f)["thumbnails"]

    videos = vdata["videos"]
    topic  = vdata["topic"]

    print("[1/3] Running comprehensive text + competitive analysis...")
    text_results = analyze_all_text(videos, channels, transcripts, topic)

    print("[2/3] Analyzing thumbnails (vision)...")
    thumb_analysis = analyze_thumbnails(thumbnails)

    # Merge: text analysis supplies the placeholder thumbnail section,
    # vision analysis overwrites it with real findings
    text_results["thumbnail_analysis"] = thumb_analysis

    # Also extract transcript_analysis field (fill empty if not present)
    if "transcript_analysis" not in text_results:
        text_results["transcript_analysis"] = {
            "note": "Transcripts unavailable — analysis based on titles and thumbnails",
            "common_themes": [], "hook_structures": [],
            "vocabulary_level": "", "content_pacing": "",
            "engagement_techniques": [],
        }

    print("[3/3] Saving insights...")
    out = DATA_DIR / "insights.json"
    with open(out, "w", encoding="utf-8") as f:
        json.dump(text_results, f, indent=2, ensure_ascii=False)
    print(f"      Saved -> {out}")


if __name__ == "__main__":
    main()
