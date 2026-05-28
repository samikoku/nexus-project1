# Skill: NAKACHI Video Generator

## What This Does
Generates branded YouTube Shorts (.mp4) from a single topic string. Each video includes Dr. Sam Ikoku's headshot composited into all frames, the NAKACHI office background, and AI-written script with text-to-speech audio. Output is a fully assembled 45–55 second 1080×1920 portrait video ready for upload.

## How to Run

```bash
ANTHROPIC_API_KEY=sk-ant-... python video_generator.py "Your topic here"
```

Example topics:
- `"Nigeria NERC Mini-Grid Regulations 2026"`
- `"Execution Drift: Why African Companies Fail"`
- `"NCC Telecom Policy and State Revenue Opportunity"`

Output lands in `output/<topic_slug>.mp4`.

## Python API

```python
from video_generator import create_video

path = create_video(
    topic="Nigeria NERC Mini-Grid Regulations 2026",
    output_dir="output",          # optional, default "output"
    speaker_wav="voice_sample.wav" # optional — activates XTTS v2 voice cloning
)
```

## Pipeline (4 steps)

1. **Script** — Claude (`claude-sonnet-4-6`) generates a JSON script: hook + 3 content slides + CTA, calibrated for NAKACHI brand voice ("Bloomberg meets Lagos boardroom").
2. **Frames + Audio** — PIL renders each slide as a 1080×1920 PNG; espeak-ng (`en-gb`, speed=148, pitch=55, amplitude=180) generates clean mono WAV per slide.
3. **Assemble** — moviepy concatenates `ImageClip + AudioFileClip` pairs; each clip duration = audio duration + 0.4 s padding.
4. **Write** — ffmpeg encodes to H.264/AAC .mp4 at 24 fps.

## Frame Layout

### Hook (opening)
- Office background photo + Deep Ink tint (alpha=120)
- Dr. Sam headshot fills lower 58%, gradient fade at top
- Headline text in dark pill, upper 42% of frame
- NAKACHI CONSULTING brand mark top-left

### Content slides (×3)
- Office background + heavier tint (alpha=175)
- Title (gold) + gold rule + body text
- Progress dots below brand mark
- Circular avatar (270px, gold ring) bottom-right
- "Dr. Sam Ikoku · NAKACHI Consulting" name badge bottom-left

### CTA (closing)
- Dr. Sam portrait fills upper 50%, fading into branded panel
- Panel: CTA headline, NAKACHI / CONSULTING, "Dr. Sam Ikoku" credit
- Office background visible at light tint (alpha=100) behind photo

## Assets

| File | Purpose |
|------|---------|
| `assets/bg_office.png` | NAKACHI office photo (1376×768), used as background on all frames |
| `assets/headshot.jpg` | Dr. Sam Ikoku headshot (3240×3828), composited into all frames |
| `assets/dr_sam_ikoku.jpg` | NAKACHI seal/crest logo (not used in video — branding reference) |

## Brand Constants

| Token | Value | Use |
|-------|-------|-----|
| Deep Ink | `#0d1b2a` / `(13, 27, 42)` | Background tint, panel fills |
| Brushed Gold | `#c5a059` / `(197, 160, 89)` | Headlines, brand mark, gold bars, avatar ring |
| White | `(255, 255, 255)` | Sub-headlines, CTA text |
| Light | `(220, 230, 240)` | Body copy |

## Voice Options

| Mode | How | Quality |
|------|-----|---------|
| `espeak-ng en-gb` | Default, offline, no setup | Robotic but intelligible |
| XTTS v2 clone | Pass `speaker_wav=` path | Dr. Sam's actual voice — requires model files |

### Activating Voice Cloning (XTTS v2)
Model files must be placed in `/root/.local/share/tts/xtts_v2/`:
- `config.json`, `model.pth`, `speakers_xtts.pth`, `vocab.json`

Download via Google Colab (HuggingFace is blocked in the cloud container):
```python
# In Colab:
from huggingface_hub import snapshot_download
from google.colab import drive
drive.mount('/content/drive')
snapshot_download("coqui/XTTS-v2", local_dir="/content/drive/MyDrive/xtts_v2_model")
```
Then pull to the container via the Google Drive MCP tool.

## Talking-Head Lip-Sync (Pending)
Current implementation composites a static headshot. True lip-sync (mouth moves with speech) requires SadTalker:
- Download SadTalker checkpoints via Colab → save to Google Drive
- Pull via Drive MCP into container
- Wire `sadtalker_inference()` into `create_video()` after audio generation

## Dependencies

```
anthropic        # Claude API script generation
moviepy          # video assembly (2.x API: with_audio, ImageClip)
Pillow           # frame rendering + alpha compositing
espeak-ng        # offline TTS (apt package)
ffmpeg           # video encoding (apt package)
```

## Repository
Branch: `claude/video-reproduction-capability-HyQkB`  
Repo: `samikoku/nexus-project1`  
Main file: `video_generator.py`
