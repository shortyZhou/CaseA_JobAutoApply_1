"""MVP placeholder: convert/refresh English resume profile.

No-op by default. Future implementation should read resumes/pdf/Resume_EN.pdf
and update profiles/profile_en.json without inventing facts.
"""
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
print({"status": "placeholder", "profile": str(ROOT / "profiles" / "profile_en.json")})
