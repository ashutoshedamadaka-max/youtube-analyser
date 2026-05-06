"""
YouTube Analyser — full pipeline orchestrator.

Usage:
    python tools/run_analysis.py --topic "kids content"
"""
import argparse
import subprocess
import sys
from pathlib import Path


def run(label, *cmd):
    print(f"\n{'─'*60}")
    print(f"  {label}")
    print(f"{'─'*60}")
    subprocess.run([sys.executable, *cmd], check=True)


def main():
    parser = argparse.ArgumentParser(description="YouTube Analyser — full pipeline")
    parser.add_argument("--topic", required=True, help='Topic to analyse, e.g. "kids content"')
    args = parser.parse_args()

    steps = [
        ("1/7  Fetch Videos",      "tools/fetch_videos.py",     "--topic", args.topic),
        ("2/7  Fetch Channels",     "tools/fetch_channels.py"),
        ("3/7  Fetch Transcripts",  "tools/fetch_transcripts.py"),
        ("4/7  Fetch Thumbnails",   "tools/fetch_thumbnails.py"),
        ("5/7  Analyse Content",    "tools/analyze_content.py"),
        ("6/7  Generate Charts",    "tools/generate_charts.py"),
    ]

    for label, script, *extra in steps:
        run(label, script, *extra)

    # JS slide builder
    print(f"\n{'─'*60}")
    print(f"  7/7  Generate Slides (pptxgenjs)")
    print(f"{'─'*60}")
    subprocess.run(["node", "tools/generate_slides.js"], check=True)

    output_files = sorted(Path("output").glob("*.pptx"),
                          key=lambda p: p.stat().st_mtime, reverse=True)
    if not output_files:
        print("\n⚠  No PPTX found in output/")
        return

    pptx = str(output_files[0])
    print(f"\n{'='*60}")
    print(f"  COMPLETE  ✓")
    print(f"  File: {pptx}")
    print(f"{'='*60}\n")

    recipient = input("Enter recipient email to send the report (or press Enter to skip): ").strip()
    if recipient:
        subprocess.run([sys.executable, "tools/send_email.py", pptx, recipient])


if __name__ == "__main__":
    main()
