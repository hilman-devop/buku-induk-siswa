"""Entrypoint serverless Vercel: semua route di-rewrite ke fungsi ini."""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from app import app  # noqa: E402,F401  (Vercel mencari variabel `app`)
