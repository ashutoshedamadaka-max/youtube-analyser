import json
from pathlib import Path
from datetime import datetime
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import numpy as np

DATA_DIR   = Path(".tmp/data")
CHARTS_DIR = Path(".tmp/charts")

# ── Dark theme palette ────────────────────────────────────────────────────────
BG      = '#0F1117'
CARD    = '#1A1D27'
BORDER  = '#2E3343'
MUTED   = '#5A5F73'
TEXT    = '#E8E9ED'
SUBLABEL = '#8B8FA3'

COLORS = ['#6C5CE7', '#00D2D3', '#FF6B81', '#FECA57',
          '#1DD1A1', '#a29bfe', '#74b9ff', '#fd79a8']

plt.rcParams.update({
    'figure.facecolor':    BG,
    'axes.facecolor':      CARD,
    'axes.edgecolor':      BORDER,
    'axes.labelcolor':     SUBLABEL,
    'text.color':          TEXT,
    'xtick.color':         MUTED,
    'ytick.color':         MUTED,
    'xtick.labelsize':     8.5,
    'ytick.labelsize':     8.5,
    'axes.labelsize':      9,
    'grid.color':          BORDER,
    'grid.alpha':          0.6,
    'grid.linewidth':      0.8,
    'axes.grid':           True,
    'axes.spines.top':     False,
    'axes.spines.right':   False,
    'axes.spines.left':    False,
    'axes.spines.bottom':  False,
    'font.family':         'sans-serif',
})


def fmt(n):
    if n >= 1_000_000_000: return f"{n/1_000_000_000:.1f}B"
    if n >= 1_000_000:     return f"{n/1_000_000:.1f}M"
    if n >= 1_000:         return f"{n/1_000:.0f}K"
    return str(n)


def save(fig, name):
    path = CHARTS_DIR / name
    fig.savefig(path, dpi=150, bbox_inches='tight', facecolor=BG)
    plt.close(fig)
    return str(path)


def chart_top_videos(videos):
    top    = videos[:10]
    labels = [v["title"][:36] + "…" if len(v["title"]) > 36 else v["title"] for v in top]
    vals   = [v["views"] for v in top]
    colors = [COLORS[0]] + [COLORS[1]] * (len(vals) - 1)
    colors[0] = '#FECA57'  # gold for #1

    fig, ax = plt.subplots(figsize=(11, 5.5))
    fig.patch.set_facecolor(BG)
    ax.set_facecolor(CARD)

    bars = ax.barh(range(len(labels)), vals, color=colors, height=0.65, zorder=3)
    ax.set_yticks(range(len(labels)))
    ax.set_yticklabels(labels, fontsize=8.5, color=TEXT)
    ax.invert_yaxis()
    ax.xaxis.set_major_formatter(plt.FuncFormatter(lambda x, _: fmt(x)))
    ax.set_xlabel("Total Views", color=SUBLABEL, fontsize=9, labelpad=8)
    ax.set_title("Top 10 Videos by View Count", fontsize=13, fontweight='bold',
                 color=TEXT, pad=14, loc='left')
    ax.tick_params(axis='x', colors=MUTED)
    ax.tick_params(axis='y', colors=TEXT)
    ax.axvline(0, color=BORDER, linewidth=0.8)

    mx = max(vals)
    for bar, v in zip(bars, vals):
        ax.text(bar.get_width() + mx * 0.01, bar.get_y() + bar.get_height() / 2,
                fmt(v), va='center', fontsize=8, color=SUBLABEL)

    plt.tight_layout(pad=1.5)
    return save(fig, 'top_videos.png')


def chart_engagement(videos):
    top = [v for v in videos[:60] if v["views"] > 0 and v["likes"] > 0]
    if not top:
        return None
    xs  = [v["views"] for v in top]
    ys  = [v["likes"] / v["views"] * 100 for v in top]
    cs  = [v["comments"] for v in top]
    mc  = max(cs) if cs else 1
    sz  = [max(18, c / mc * 280) for c in cs]

    fig, ax = plt.subplots(figsize=(10, 5.2))
    fig.patch.set_facecolor(BG)
    ax.set_facecolor(CARD)

    ax.scatter(xs, ys, s=sz, color=COLORS[2], alpha=0.7, edgecolors=BORDER,
               linewidth=0.5, zorder=3)
    ax.set_xscale('log')
    ax.xaxis.set_major_formatter(plt.FuncFormatter(lambda x, _: fmt(x)))
    ax.set_xlabel("Total Views (log scale)", color=SUBLABEL, fontsize=9, labelpad=8)
    ax.set_ylabel("Engagement Rate  (likes / views %)", color=SUBLABEL, fontsize=9, labelpad=8)
    ax.set_title("Engagement Rate vs. Views\nbubble size = comment count",
                 fontsize=13, fontweight='bold', color=TEXT, pad=14, loc='left')
    ax.tick_params(colors=MUTED)

    plt.tight_layout(pad=1.5)
    return save(fig, 'engagement.png')


def chart_top_channels(channels):
    top    = channels[:10]
    labels = [c["title"][:22] + "…" if len(c["title"]) > 22 else c["title"] for c in top]
    vals   = [c["subscribers"] for c in top]
    bar_colors = [COLORS[i % len(COLORS)] for i in range(len(vals))]

    fig, ax = plt.subplots(figsize=(11, 5.5))
    fig.patch.set_facecolor(BG)
    ax.set_facecolor(CARD)

    bars = ax.barh(range(len(labels)), vals, color=bar_colors, height=0.65, zorder=3)
    ax.set_yticks(range(len(labels)))
    ax.set_yticklabels(labels, fontsize=8.5, color=TEXT)
    ax.invert_yaxis()
    ax.xaxis.set_major_formatter(plt.FuncFormatter(lambda x, _: fmt(x)))
    ax.set_xlabel("Subscribers", color=SUBLABEL, fontsize=9, labelpad=8)
    ax.set_title("Top Channels by Subscriber Count", fontsize=13, fontweight='bold',
                 color=TEXT, pad=14, loc='left')
    ax.tick_params(axis='x', colors=MUTED)
    ax.tick_params(axis='y', colors=TEXT)

    mx = max(vals)
    for bar, v in zip(bars, vals):
        ax.text(bar.get_width() + mx * 0.01, bar.get_y() + bar.get_height() / 2,
                fmt(v), va='center', fontsize=8, color=SUBLABEL)

    plt.tight_layout(pad=1.5)
    return save(fig, 'top_channels.png')


def chart_views_distribution(videos):
    vals = [v["views"] for v in videos if v["views"] > 0]
    if not vals:
        return None

    fig, ax = plt.subplots(figsize=(9, 4.5))
    fig.patch.set_facecolor(BG)
    ax.set_facecolor(CARD)

    ax.hist(vals, bins=28, color=COLORS[3], edgecolor=BG, linewidth=0.5,
            log=True, zorder=3)
    ax.xaxis.set_major_formatter(plt.FuncFormatter(lambda x, _: fmt(x)))
    ax.set_xlabel("Views", color=SUBLABEL, fontsize=9, labelpad=8)
    ax.set_ylabel("Video Count (log)", color=SUBLABEL, fontsize=9, labelpad=8)
    ax.set_title("Distribution of Video Views", fontsize=13, fontweight='bold',
                 color=TEXT, pad=14, loc='left')
    ax.tick_params(colors=MUTED)

    plt.tight_layout(pad=1.5)
    return save(fig, 'views_distribution.png')


def chart_content_themes(insights):
    themes = insights.get("title_analysis", {}).get("dominant_content_themes", [])
    if not themes:
        return None

    labels = [t["theme"] for t in themes]
    sizes  = [t.get("percentage", 20) for t in themes]
    pie_colors = COLORS[:len(labels)]

    fig, ax = plt.subplots(figsize=(7, 6.5))
    fig.patch.set_facecolor(BG)
    ax.set_facecolor(BG)

    wedges, texts, autotexts = ax.pie(
        sizes, labels=labels, autopct="%1.0f%%",
        colors=pie_colors, startangle=90,
        wedgeprops={"edgecolor": BG, "linewidth": 2.5},
        pctdistance=0.76, labeldistance=1.12,
        textprops={"color": TEXT, "fontsize": 9.5},
    )
    for a in autotexts:
        a.set_fontsize(9)
        a.set_color(BG)
        a.set_fontweight('bold')

    # Donut hole
    ax.add_artist(plt.Circle((0, 0), 0.48, fc=BG))
    ax.set_title("Content Theme Distribution", fontsize=13, fontweight='bold',
                 color=TEXT, pad=16)

    plt.tight_layout(pad=1.2)
    return save(fig, 'content_themes.png')


def chart_upload_cadence(videos):
    days   = ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"]
    counts = [0] * 7
    for v in videos:
        try:
            dt = datetime.fromisoformat(v["published_at"].replace("Z", "+00:00"))
            counts[dt.weekday()] += 1
        except Exception:
            pass

    peak = max(counts)
    bar_colors = [COLORS[0] if c == peak else COLORS[1] for c in counts]

    fig, ax = plt.subplots(figsize=(8.5, 4.5))
    fig.patch.set_facecolor(BG)
    ax.set_facecolor(CARD)

    bars = ax.bar(days, counts, color=bar_colors, zorder=3, width=0.6)
    ax.set_ylabel("Videos Published", color=SUBLABEL, fontsize=9, labelpad=8)
    ax.set_title("Upload Day Distribution  (violet = peak)", fontsize=13,
                 fontweight='bold', color=TEXT, pad=14, loc='left')
    ax.tick_params(colors=MUTED)

    for bar, c in zip(bars, counts):
        ax.text(bar.get_x() + bar.get_width() / 2, bar.get_height() + 0.2,
                str(c), ha='center', fontsize=9.5, color=SUBLABEL)

    plt.tight_layout(pad=1.5)
    return save(fig, 'upload_cadence.png')


def main():
    CHARTS_DIR.mkdir(parents=True, exist_ok=True)

    with open(DATA_DIR / "videos.json",   encoding="utf-8") as f: videos   = json.load(f)["videos"]
    with open(DATA_DIR / "channels.json", encoding="utf-8") as f: channels = json.load(f)["channels"]
    with open(DATA_DIR / "insights.json", encoding="utf-8") as f: insights = json.load(f)

    steps = [
        ("Top videos",       lambda: chart_top_videos(videos)),
        ("Engagement",       lambda: chart_engagement(videos)),
        ("Top channels",     lambda: chart_top_channels(channels)),
        ("Views dist",       lambda: chart_views_distribution(videos)),
        ("Content themes",   lambda: chart_content_themes(insights)),
        ("Upload cadence",   lambda: chart_upload_cadence(videos)),
    ]
    for i, (name, fn) in enumerate(steps, 1):
        print(f"[{i}/{len(steps)}] {name}")
        fn()
    print(f"      Charts saved -> {CHARTS_DIR}")


if __name__ == "__main__":
    main()
