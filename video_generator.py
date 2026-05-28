"""
NAKACHI Video Generator
Generates YouTube Shorts (.mp4) from a topic string.

Usage:
    python video_generator.py "Your topic here"
    ANTHROPIC_API_KEY=sk-... python video_generator.py "Nigeria Mini-Grid Regulations"
"""

import os
import json
import subprocess
import tempfile
import shutil
import sys
from pathlib import Path

from PIL import Image, ImageDraw, ImageFont
from moviepy import ImageClip, AudioFileClip, concatenate_videoclips
import anthropic

# ─── Visual constants ────────────────────────────────────────────────────────
WIDTH, HEIGHT = 1080, 1920          # 9:16 portrait (YouTube Shorts)
BG          = (13,  27,  42)        # #0d1b2a  Deep Ink
GOLD        = (197, 160,  89)       # #c5a059  Brushed Gold
WHITE       = (255, 255, 255)
LIGHT       = (200, 215, 230)       # body text
MARGIN      = 90
CONTENT_W   = WIDTH - 2 * MARGIN

FONT_BOLD = "/usr/share/fonts/truetype/liberation/LiberationSans-Bold.ttf"
FONT_REG  = "/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf"


# ─── Helpers ─────────────────────────────────────────────────────────────────

def _load_fonts() -> dict:
    specs = {"xl": (FONT_BOLD, 86), "lg": (FONT_BOLD, 62), "md": (FONT_REG, 46),
             "sm": (FONT_REG,  34), "brand": (FONT_BOLD, 40)}
    out = {}
    for k, (path, size) in specs.items():
        try:
            out[k] = ImageFont.truetype(path, size)
        except OSError:
            out[k] = ImageFont.load_default()
    return out


def _wrap(text: str, font, max_w: int) -> list[str]:
    words, lines, cur = text.split(), [], []
    for w in words:
        test = " ".join(cur + [w])
        if font.getbbox(test)[2] <= max_w:
            cur.append(w)
        else:
            if cur:
                lines.append(" ".join(cur))
            cur = [w]
    if cur:
        lines.append(" ".join(cur))
    return lines or [""]


def _draw_block(draw, text, x, y, font, color, max_w, gap=14) -> int:
    for line in _wrap(text, font, max_w):
        draw.text((x, y), line, fill=color, font=font)
        y += font.getbbox(line)[3] + gap
    return y


def _centered_x(text, font) -> int:
    return (WIDTH - font.getbbox(text)[2]) // 2


# ─── Frame builders ───────────────────────────────────────────────────────────

def _frame_hook(data: dict, fonts: dict) -> Image.Image:
    img = Image.new("RGB", (WIDTH, HEIGHT), BG)
    d   = ImageDraw.Draw(img)

    d.rectangle([0, 0, WIDTH, 8], fill=GOLD)
    d.rectangle([0, HEIGHT - 8, WIDTH, HEIGHT], fill=GOLD)
    d.text((MARGIN, 36), "NAKACHI", fill=GOLD, font=fonts["brand"])

    headline = data.get("headline", "")
    subtext  = data.get("subtext", "")

    h_lines = _wrap(headline, fonts["xl"], CONTENT_W)
    s_lines = _wrap(subtext,  fonts["md"], CONTENT_W) if subtext else []
    total_h = len(h_lines) * 102 + (len(s_lines) * 58 + 48 if s_lines else 0)
    y = max(300, (HEIGHT - total_h) // 2)

    for line in h_lines:
        d.text((_centered_x(line, fonts["xl"]), y), line, fill=GOLD, font=fonts["xl"])
        y += fonts["xl"].getbbox(line)[3] + 14

    if subtext:
        y += 48
        for line in s_lines:
            d.text((_centered_x(line, fonts["md"]), y), line, fill=WHITE, font=fonts["md"])
            y += fonts["md"].getbbox(line)[3] + 12

    return img


def _frame_content(data: dict, fonts: dict) -> Image.Image:
    img = Image.new("RGB", (WIDTH, HEIGHT), BG)
    d   = ImageDraw.Draw(img)

    d.rectangle([0, 0, WIDTH, 8], fill=GOLD)
    d.rectangle([0, HEIGHT - 8, WIDTH, HEIGHT], fill=GOLD)
    d.text((MARGIN, 36), "NAKACHI", fill=GOLD, font=fonts["brand"])

    total = data.get("total", 1)
    idx   = data.get("slide_idx", 1)
    dot_x = MARGIN
    for i in range(total):
        fill = GOLD if i < idx else (40, 65, 90)
        d.ellipse([dot_x, 118, dot_x + 20, 138], fill=fill)
        dot_x += 36

    y = 190
    y = _draw_block(d, data.get("title", ""), MARGIN, y, fonts["lg"], GOLD, CONTENT_W, gap=14)

    y += 22
    d.rectangle([MARGIN, y, MARGIN + 130, y + 5], fill=GOLD)
    y += 46

    body = data.get("body", "")
    if body:
        _draw_block(d, body, MARGIN, y, fonts["md"], LIGHT, CONTENT_W, gap=18)

    return img


def _frame_cta(data: dict, fonts: dict) -> Image.Image:
    img = Image.new("RGB", (WIDTH, HEIGHT), BG)
    d   = ImageDraw.Draw(img)

    d.rectangle([0, 0, WIDTH, 8], fill=GOLD)
    d.rectangle([0, HEIGHT - 8, WIDTH, HEIGHT], fill=GOLD)

    headline = data.get("headline", "Follow for more")

    y = HEIGHT // 2 - 180
    d.rectangle([MARGIN, y, WIDTH - MARGIN, y + 4], fill=GOLD)
    y += 44

    for line in _wrap(headline, fonts["lg"], CONTENT_W):
        d.text((_centered_x(line, fonts["lg"]), y), line, fill=WHITE, font=fonts["lg"])
        y += fonts["lg"].getbbox(line)[3] + 16

    y += 40
    d.rectangle([MARGIN, y, WIDTH - MARGIN, y + 4], fill=GOLD)
    y += 72

    brand = "NAKACHI"
    d.text((_centered_x(brand, fonts["xl"]), y), brand, fill=GOLD, font=fonts["xl"])
    y += fonts["xl"].getbbox(brand)[3] + 16

    sub = "Consulting"
    d.text((_centered_x(sub, fonts["md"]), y), sub, fill=WHITE, font=fonts["md"])

    return img


def make_frame(slide_type: str, data: dict, fonts: dict) -> Image.Image:
    builders = {"hook": _frame_hook, "content": _frame_content, "cta": _frame_cta}
    return builders[slide_type](data, fonts)


# ─── TTS ─────────────────────────────────────────────────────────────────────

def tts(text: str, wav_path: str, voice: str = "en-gb",
        speed: int = 148, pitch: int = 55, amplitude: int = 180) -> None:
    """Generate clean mono WAV using espeak-ng.

    en-gb: authoritative British English, clean 22 kHz mono — no double-track artifact.
    For your own voice, use tts_xtts() once the XTTS v2 model is available.
    """
    subprocess.run(
        ["espeak-ng", "-v", voice, "-s", str(speed),
         "-p", str(pitch), "-a", str(amplitude), "-w", wav_path, text],
        check=True, capture_output=True
    )


# ─── Script generation ────────────────────────────────────────────────────────

SCRIPT_PROMPT = """Create a YouTube Shorts script (45–55 seconds spoken) on: "{topic}"

Brand: NAKACHI Consulting — premium Nigerian business strategy & energy consulting.
Voice: Authoritative, direct, no fluff. Think Bloomberg meets Lagos boardroom.

Return ONLY valid JSON (no markdown fences):
{{
  "hook": {{
    "headline": "Bold statement or question (max 8 words)",
    "subtext": "Supporting line (max 10 words)",
    "voiceover": "1-2 punchy sentences"
  }},
  "slides": [
    {{
      "title": "Key point (max 5 words)",
      "body": "Brief explanation (max 18 words)",
      "voiceover": "2-3 sentences for this point"
    }}
  ],
  "cta": {{
    "headline": "Call to action (max 6 words)",
    "voiceover": "1-2 closing sentences"
  }}
}}

Exactly 3 slides. Content must be specific, expert, and relevant to Nigeria/Africa."""


def generate_script(topic: str) -> dict:
    client = anthropic.Anthropic()
    msg = client.messages.create(
        model="claude-sonnet-4-6",
        max_tokens=1500,
        system="You are a video scriptwriter for NAKACHI Consulting. Return only valid JSON.",
        messages=[{"role": "user", "content": SCRIPT_PROMPT.format(topic=topic)}]
    )
    raw = msg.content[0].text
    return json.loads(raw[raw.find("{") : raw.rfind("}") + 1])


# ─── Main pipeline ────────────────────────────────────────────────────────────

def create_video(topic: str, output_path: str | None = None, output_dir: str = "output") -> str:
    """
    Full pipeline: topic string → .mp4 file.

    Args:
        topic:       Topic for the Short (e.g. "Nigeria NERC Mini-Grid 2026")
        output_path: Filename for the output (auto-generated if None)
        output_dir:  Directory to write the video into

    Returns:
        Absolute path of the generated .mp4 file.
    """
    os.makedirs(output_dir, exist_ok=True)
    tmp = tempfile.mkdtemp(prefix="nakachi_vid_")

    try:
        # 1. Script
        print(f'\n[1/4] Generating script → "{topic}"')
        script = generate_script(topic)

        fonts = _load_fonts()

        # Build ordered section list
        n = len(script["slides"])
        sections = [
            {"type": "hook",    "frame": {"headline": script["hook"]["headline"],
                                          "subtext":  script["hook"].get("subtext", "")},
             "vo": script["hook"]["voiceover"]},
            *[{"type": "content",
               "frame": {"title": s["title"], "body": s.get("body", ""),
                         "slide_idx": i + 1, "total": n},
               "vo": s["voiceover"]}
              for i, s in enumerate(script["slides"])],
            {"type": "cta",    "frame": {"headline": script["cta"]["headline"]},
             "vo": script["cta"]["voiceover"]},
        ]

        # 2. Frames + audio
        print("[2/4] Rendering frames and generating audio...")
        clips = []
        for i, sec in enumerate(sections):
            img_path = os.path.join(tmp, f"frame_{i:02d}.png")
            wav_path = os.path.join(tmp, f"audio_{i:02d}.wav")

            make_frame(sec["type"], sec["frame"], fonts).save(img_path)
            tts(sec["vo"], wav_path)

            audio    = AudioFileClip(wav_path)
            duration = audio.duration + 0.4
            clip     = ImageClip(img_path, duration=duration).with_audio(audio)
            clips.append(clip)

        # 3. Assemble
        print("[3/4] Assembling clips...")
        video = concatenate_videoclips(clips, method="compose")

        # 4. Write
        if output_path is None:
            safe = "".join(c if c.isalnum() or c in " -_" else "_" for c in topic)[:50]
            output_path = safe.strip().replace(" ", "_") + ".mp4"

        dest = os.path.join(output_dir, output_path)
        print(f"[4/4] Writing video → {dest}")
        video.write_videofile(dest, fps=24, codec="libx264", audio_codec="aac", logger=None)

        print(f"\n✓  Video saved: {dest}")
        return os.path.abspath(dest)

    finally:
        shutil.rmtree(tmp, ignore_errors=True)


# ─── CLI entry point ──────────────────────────────────────────────────────────

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python video_generator.py \"Your topic here\"")
        print("Example topics:")
        print("  \"Nigeria NERC Mini-Grid Regulations 2026\"")
        print("  \"Execution Drift: Why African Companies Fail\"")
        print("  \"NCC Telecom Policy and State Revenue Opportunity\"")
        sys.exit(1)

    if not os.environ.get("ANTHROPIC_API_KEY"):
        print("Error: ANTHROPIC_API_KEY environment variable is not set.")
        print("Set it with:  export ANTHROPIC_API_KEY=sk-ant-...")
        sys.exit(1)

    topic = " ".join(sys.argv[1:])
    create_video(topic)
