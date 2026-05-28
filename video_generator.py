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

from PIL import Image, ImageDraw, ImageFont, ImageFilter
from moviepy import ImageClip, AudioFileClip, concatenate_videoclips
import anthropic

# ─── Visual constants ────────────────────────────────────────────────────────
WIDTH, HEIGHT = 1080, 1920          # 9:16 portrait (YouTube Shorts)
BG          = (13,  27,  42)        # #0d1b2a  Deep Ink (fallback only)
GOLD        = (197, 160,  89)       # #c5a059  Brushed Gold
WHITE       = (255, 255, 255)
LIGHT       = (220, 230, 240)       # body text
MARGIN      = 90
CONTENT_W   = WIDTH - 2 * MARGIN

FONT_BOLD = "/usr/share/fonts/truetype/liberation/LiberationSans-Bold.ttf"
FONT_REG  = "/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf"

ASSETS_DIR    = Path(__file__).parent / "assets"
BG_IMAGE      = ASSETS_DIR / "bg_office.png"   # NAKACHI office photo
HEADSHOT_PATH = ASSETS_DIR / "headshot.jpg"     # Dr. Sam Ikoku presenter photo


# ─── Background ──────────────────────────────────────────────────────────────

def _load_bg() -> Image.Image:
    """Scale + centre-crop the office photo to exactly 1080×1920."""
    if not BG_IMAGE.exists():
        return Image.new("RGB", (WIDTH, HEIGHT), (13, 27, 42))
    src = Image.open(BG_IMAGE).convert("RGB")
    sw, sh = src.size
    scale  = max(WIDTH / sw, HEIGHT / sh)
    nw, nh = int(sw * scale), int(sh * scale)
    src    = src.resize((nw, nh), Image.LANCZOS)
    left   = (nw - WIDTH)  // 2
    top    = (nh - HEIGHT) // 2
    return src.crop((left, top, left + WIDTH, top + HEIGHT))


def _overlay(base: Image.Image, alpha: int = 160) -> Image.Image:
    """Lay a Deep Ink tint over the photo (alpha 0–255, higher = darker)."""
    tint = Image.new("RGBA", (WIDTH, HEIGHT), (13, 27, 42, alpha))
    out  = base.convert("RGBA")
    out.alpha_composite(tint)
    return out.convert("RGB")


# ─── Presenter compositing ────────────────────────────────────────────────────

def _load_headshot(target_w: int, target_h: int) -> "Image.Image | None":
    """Load and scale headshot to fit within target_w × target_h."""
    if not HEADSHOT_PATH.exists():
        return None
    src   = Image.open(HEADSHOT_PATH).convert("RGBA")
    sw, sh = src.size
    scale = min(target_w / sw, target_h / sh)
    return src.resize((int(sw * scale), int(sh * scale)), Image.LANCZOS)


def _presenter_hook(img: Image.Image) -> Image.Image:
    """Composite presenter photo into lower 58% of hook frame, fading in at top."""
    shot = _load_headshot(WIDTH, int(HEIGHT * 0.62))
    if shot is None:
        return img
    w, h = shot.size
    mask = Image.new("L", (w, h), 255)
    fade = int(h * 0.38)
    md   = ImageDraw.Draw(mask)
    for row in range(fade):
        md.line([(0, row), (w, row)], fill=int(255 * (row / fade)))
    shot.putalpha(mask)
    base = img.convert("RGBA")
    base.alpha_composite(shot, dest=((WIDTH - w) // 2, HEIGHT - h))
    return base.convert("RGB")


def _presenter_avatar(img: Image.Image, size: int = 270,
                      right_margin: int = 80, bottom_margin: int = 90) -> Image.Image:
    """Paste a circular presenter avatar with gold ring in the bottom-right corner."""
    shot = _load_headshot(size, size)
    if shot is None:
        return img
    sw, sh = shot.size
    sq     = min(sw, sh)
    shot   = shot.crop(((sw - sq) // 2, (sh - sq) // 2,
                         (sw + sq) // 2, (sh + sq) // 2)).resize((size, size), Image.LANCZOS)
    circle = Image.new("L", (size, size), 0)
    ImageDraw.Draw(circle).ellipse([0, 0, size - 1, size - 1], fill=255)
    shot   = shot.convert("RGBA")
    shot.putalpha(circle)

    x, y  = WIDTH - size - right_margin, HEIGHT - size - bottom_margin
    base  = img.convert("RGBA")
    ring  = Image.new("RGBA", (WIDTH, HEIGHT), (0, 0, 0, 0))
    b     = 5
    ImageDraw.Draw(ring).ellipse([x - b, y - b, x + size + b, y + size + b],
                                  outline=(*GOLD, 255), width=b)
    base  = Image.alpha_composite(base, ring)
    base.alpha_composite(shot, dest=(x, y))
    return base.convert("RGB")


def _presenter_cta(img: Image.Image) -> Image.Image:
    """Composite presenter photo in upper portion of CTA frame, fading out at bottom."""
    shot = _load_headshot(int(WIDTH * 0.72), int(HEIGHT * 0.50))
    if shot is None:
        return img
    w, h  = shot.size
    mask  = Image.new("L", (w, h), 255)
    start = int(h * 0.68)
    md    = ImageDraw.Draw(mask)
    for row in range(start, h):
        md.line([(0, row), (w, row)], fill=int(255 * (1 - (row - start) / (h - start))))
    shot.putalpha(mask)
    base = img.convert("RGBA")
    base.alpha_composite(shot, dest=((WIDTH - w) // 2, 75))
    return base.convert("RGB")


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
    # Background + presenter photo (lower portion, gradient fade)
    img = _presenter_hook(_overlay(_load_bg(), alpha=120))
    d   = ImageDraw.Draw(img)

    # Gold bars
    d.rectangle([0, 0, WIDTH, 10], fill=GOLD)
    d.rectangle([0, HEIGHT - 10, WIDTH, HEIGHT], fill=GOLD)

    # Brand mark
    d.text((MARGIN, 38), "NAKACHI", fill=GOLD, font=fonts["brand"])
    d.text((MARGIN + fonts["brand"].getbbox("NAKACHI")[2] + 12, 46),
           "CONSULTING", fill=WHITE, font=fonts["sm"])

    headline = data.get("headline", "")
    subtext  = data.get("subtext", "")

    h_lines = _wrap(headline, fonts["xl"], CONTENT_W)
    s_lines = _wrap(subtext, fonts["md"], CONTENT_W) if subtext else []
    total_h = len(h_lines) * 106 + (len(s_lines) * 60 + 56 if s_lines else 0)
    # Keep text in upper 42% so it clears the presenter photo
    y = max(130, min(700, (int(HEIGHT * 0.42) - total_h) // 2 + 80))

    # Semi-transparent pill behind headline for readability
    pad = 24
    pill_top    = y - pad
    pill_bottom = y + total_h + pad
    overlay = Image.new("RGBA", (WIDTH, HEIGHT), (0, 0, 0, 0))
    od = ImageDraw.Draw(overlay)
    od.rectangle([MARGIN - pad, pill_top, WIDTH - MARGIN + pad, pill_bottom],
                 fill=(13, 27, 42, 195))
    img = Image.alpha_composite(img.convert("RGBA"), overlay).convert("RGB")
    d   = ImageDraw.Draw(img)

    for line in h_lines:
        d.text((_centered_x(line, fonts["xl"]), y), line, fill=GOLD, font=fonts["xl"])
        y += fonts["xl"].getbbox(line)[3] + 16

    if subtext:
        y += 40
        for line in s_lines:
            d.text((_centered_x(line, fonts["md"]), y), line, fill=WHITE, font=fonts["md"])
            y += fonts["md"].getbbox(line)[3] + 14

    return img


def _frame_content(data: dict, fonts: dict) -> Image.Image:
    # Background: heavier tint + circular avatar in bottom-right
    img = _presenter_avatar(_overlay(_load_bg(), alpha=175))
    d   = ImageDraw.Draw(img)

    d.rectangle([0, 0, WIDTH, 10], fill=GOLD)
    d.rectangle([0, HEIGHT - 10, WIDTH, HEIGHT], fill=GOLD)
    d.text((MARGIN, 38), "NAKACHI", fill=GOLD, font=fonts["brand"])

    # Progress dots
    total = data.get("total", 1)
    idx   = data.get("slide_idx", 1)
    dx    = MARGIN
    for i in range(total):
        d.ellipse([dx, 118, dx + 20, 138], fill=GOLD if i < idx else (80, 100, 120))
        dx += 36

    y = 195
    y = _draw_block(d, data.get("title", ""), MARGIN, y, fonts["lg"], GOLD, CONTENT_W, gap=14)

    y += 20
    d.rectangle([MARGIN, y, MARGIN + 140, y + 5], fill=GOLD)
    y += 44

    body = data.get("body", "")
    if body:
        _draw_block(d, body, MARGIN, y, fonts["md"], LIGHT, CONTENT_W, gap=18)

    # Name badge bottom-left, above gold bar
    d.text((MARGIN, HEIGHT - 95), "Dr. Sam Ikoku  ·  NAKACHI Consulting", fill=GOLD,
           font=fonts["sm"])

    return img


def _frame_cta(data: dict, fonts: dict) -> Image.Image:
    # CTA: presenter photo in upper portion, branding panel in lower portion
    img = _presenter_cta(_overlay(_load_bg(), alpha=100))
    d   = ImageDraw.Draw(img)

    d.rectangle([0, 0, WIDTH, 10], fill=GOLD)
    d.rectangle([0, HEIGHT - 10, WIDTH, HEIGHT], fill=GOLD)

    headline = data.get("headline", "Follow for more")

    # Branding panel anchored to bottom half
    panel_top    = int(HEIGHT * 0.52)
    panel_bottom = HEIGHT - 20
    overlay = Image.new("RGBA", (WIDTH, HEIGHT), (0, 0, 0, 0))
    od = ImageDraw.Draw(overlay)
    od.rectangle([0, panel_top, WIDTH, panel_bottom], fill=(13, 27, 42, 218))
    img = Image.alpha_composite(img.convert("RGBA"), overlay).convert("RGB")
    d   = ImageDraw.Draw(img)

    y = panel_top + 40
    d.rectangle([MARGIN, y, WIDTH - MARGIN, y + 4], fill=GOLD)
    y += 36

    for line in _wrap(headline, fonts["lg"], CONTENT_W):
        d.text((_centered_x(line, fonts["lg"]), y), line, fill=WHITE, font=fonts["lg"])
        y += fonts["lg"].getbbox(line)[3] + 16

    y += 28
    d.rectangle([MARGIN, y, WIDTH - MARGIN, y + 4], fill=GOLD)
    y += 52

    brand = "NAKACHI"
    d.text((_centered_x(brand, fonts["xl"]), y), brand, fill=GOLD, font=fonts["xl"])
    y += fonts["xl"].getbbox(brand)[3] + 14

    sub = "CONSULTING"
    d.text((_centered_x(sub, fonts["sm"]), y), sub, fill=WHITE, font=fonts["sm"])
    y += fonts["sm"].getbbox(sub)[3] + 20

    d.text((_centered_x("Dr. Sam Ikoku", fonts["sm"]), y),
           "Dr. Sam Ikoku", fill=GOLD, font=fonts["sm"])

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


def tts_xtts(text: str, wav_path: str, speaker_wav: str,
             model_dir: str = "/root/.local/share/tts/xtts_v2") -> None:
    """Generate audio in the speaker's voice using XTTS v2.

    Requires the 4 XTTS v2 model files in model_dir:
        config.json, model.pth, speakers_xtts.pth, vocab.json
    Upload them from https://huggingface.co/coqui/XTTS-v2 via Google Drive,
    then call: use_voice_clone(speaker_wav_path) to activate.
    """
    import os; os.environ["COQUI_TOS_AGREED"] = "1"
    from TTS.api import TTS as CoquiTTS
    model = CoquiTTS(model_path=model_dir,
                     config_path=f"{model_dir}/config.json",
                     progress_bar=False)
    model.tts_to_file(text=text, speaker_wav=speaker_wav,
                      language="en", file_path=wav_path)


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

def create_video(topic: str, output_path: str | None = None, output_dir: str = "output",
                 speaker_wav: str | None = None) -> str:
    """
    Full pipeline: topic string → .mp4 file.

    Args:
        topic:        Topic for the Short (e.g. "Nigeria NERC Mini-Grid 2026")
        output_path:  Filename for the output (auto-generated if None)
        output_dir:   Directory to write the video into
        speaker_wav:  Path to a .wav voice sample for XTTS v2 cloning.
                      If None, falls back to MBROLA offline TTS.

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
            if speaker_wav:
                tts_xtts(sec["vo"], wav_path, speaker_wav)
            else:
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
