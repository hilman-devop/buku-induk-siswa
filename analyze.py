"""Ambil video TikTok (atau file lokal) -> frame PNG + transkrip.

Pakai:  .venv/Scripts/python.exe analyze.py <url|file> [--frames 12] [--model small]
Hasil:  out/<slug>/  berisi video, frame_*.png, transcript.txt, meta.json
"""
import argparse
import json
import pathlib
import re
import sys

import cv2


def slug(s: str) -> str:
    return re.sub(r"\W+", "-", s)[:60].strip("-") or "video"


def fetch(src: str, base: pathlib.Path) -> tuple[pathlib.Path, dict, pathlib.Path]:
    """-> (file video, metadata, folder hasil)"""
    if not src.startswith("http"):
        vid = pathlib.Path(src)
        out = base / slug(vid.stem)
        out.mkdir(parents=True, exist_ok=True)
        return vid, {}, out
    import yt_dlp

    out = base / slug(src.split("?")[0])
    out.mkdir(parents=True, exist_ok=True)
    opts = {
        # ponytail: satu file mp4 saja supaya tak butuh ffmpeg untuk merge
        "format": "b[ext=mp4]/b",
        "outtmpl": str(out / "video.%(ext)s"),
        "quiet": True,
        "noprogress": True,
        "noplaylist": True,
    }
    with yt_dlp.YoutubeDL(opts) as ydl:
        info = ydl.extract_info(src, download=True)
    vid = next(p for p in out.iterdir() if p.suffix.lower() in (".mp4", ".webm", ".mkv", ".mov"))
    keep = {k: info.get(k) for k in
            ("id", "title", "description", "uploader", "uploader_id", "webpage_url",
             "duration", "view_count", "like_count", "comment_count", "repost_count",
             "upload_date", "track", "artist", "tags", "timestamp")}
    return vid, keep, out


def grab_frames(vid: pathlib.Path, out: pathlib.Path, n: int) -> list[pathlib.Path]:
    cap = cv2.VideoCapture(str(vid))
    if not cap.isOpened():
        sys.exit(f"gagal buka video: {vid}")
    total = int(cap.get(cv2.CAP_PROP_FRAME_COUNT)) or 0
    fps = cap.get(cv2.CAP_PROP_FPS) or 30.0
    dur = total / fps if total else 0
    marks = [dur * i / max(n - 1, 1) for i in range(n)]
    paths, prev = [], None
    for i, t in enumerate(marks):
        cap.set(cv2.CAP_PROP_POS_MSEC, max(t, 0) * 1000)
        ok, frame = cap.read()
        if not ok or frame is None or frame.size == 0:
            continue  # akhir video / seek gagal
        small = cv2.cvtColor(cv2.resize(frame, (64, 64)), cv2.COLOR_BGR2GRAY).astype("int16")
        if prev is not None and abs(small - prev).mean() < 1.0:
            continue  # ponytail: dedupe kasar, naikkan ke perceptual hash kalau per frame mau presisi
        prev = small
        p = out / f"frame_{i:02d}_{t:05.1f}s.png"
        cv2.imwrite(str(p), frame)
        paths.append(p)
    cap.release()
    return paths


def transcribe(vid: pathlib.Path, model_name: str) -> str:
    from faster_whisper import WhisperModel

    model = WhisperModel(model_name, device="cpu", compute_type="int8")
    segments, info = model.transcribe(str(vid), vad_filter=True, beam_size=5)
    lines = [f"[{s.start:6.1f}-{s.end:6.1f}] {s.text.strip()}" for s in segments]
    head = f"# bahasa={info.language} ({info.language_probability:.2f}) model={model_name}"
    return "\n".join([head, *lines])


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("src")
    ap.add_argument("--frames", type=int, default=12)
    ap.add_argument("--model", default="small")
    ap.add_argument("--no-transcript", action="store_true")
    a = ap.parse_args()

    vid, meta, out = fetch(a.src, pathlib.Path(__file__).parent / "out")
    (out / "meta.json").write_text(json.dumps(meta, ensure_ascii=False, indent=2), "utf-8")
    frames = grab_frames(vid, out, a.frames)
    text = "" if a.no_transcript else transcribe(vid, a.model)
    if text:
        (out / "transcript.txt").write_text(text, "utf-8")
    print(json.dumps({
        "dir": str(out),
        "video": str(vid),
        "meta": meta,
        "frames": [str(p) for p in frames],
        "transcript": text,
    }, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
