"""MVP placeholder: check application statuses.

Must not log in automatically or change Excel when no clear status change exists.
"""
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
print({"status": "placeholder", "workbook": str(ROOT / "excel" / "岗位搜索与简历投递.xlsx")})
