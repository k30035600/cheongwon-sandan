import sys
sys.path.insert(0, "tools")
sys.stdout.reconfigure(encoding="utf-8")
from match_g2b_unmatched import load_unmatched

items = [
    i for i in load_unmatched()
    if i["ledger"].startswith("01") and i["unit"] in ("주", "본", "EA", "ea")
]
print(len(items), "tree-like units")
for i in items:
    print(f"  {i['name'][:45]} | {i['spec'][:35]} | {i['unit']} qty={i['qty']}")
