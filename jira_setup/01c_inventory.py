"""Phase 1c · Inventory existing fields/types to identify reuse vs create-new."""
import json
from client import get, checkpoint_save

# What we need vs what might already exist (by name keyword)
NEEDS = {
    "value stream":        ["value_stream", "valuestream"],
    "opco":                ["opco"],
    "sponsor":             ["sponsor"],
    "business value":      ["business value", "bv"],
    "acceptance criteria": ["acceptance"],
    "customer impact":     ["customer impact", "impact tier"],
    "type / category":     ["demand type", "type"],
    "size / t-shirt":      ["size", "t-shirt", "tshirt"],
    "class of service":    ["class of service", "cos"],
    "run grow transform":  ["run grow", "rgt", "run/grow"],
    "pain score":          ["pain score", "painscore"],
    "ux score":            ["ux score", "ux"],
    "safety score":        ["safety score", "safety"],
    "effort":              ["effort", "complexity", "estimate"],
    "opco multiplier":     ["multiplier"],
    "parent initiative":   ["parent initiative", "epic link"],
    "discovery report":    ["discovery"],
    "concept prototype":   ["prototype", "concept"],
}

ISSUE_TYPE_NEEDS = ["Demand", "Story", "Epic", "Bug", "Sub-task", "Initiative"]


def main():
    print("=" * 70)
    print("PHASE 1c · INVENTORY · existing fields/types we can REUSE")
    print("=" * 70)

    # === Custom fields ===
    print("\n[1/2] Inventorying custom fields ...")
    code, body = get("/rest/api/3/field")
    if code != 200:
        print(f"  ✗ FAIL · {code}"); return
    custom = [f for f in body if f.get("custom")]
    print(f"  ✓ {len(custom)} custom fields globally\n")
    print(f"  {'NEED':25s} → {'CANDIDATES (reuse?)'}")
    print(f"  {'-'*25}   {'-'*60}")
    field_matches = {}
    for need_label, keywords in NEEDS.items():
        matches = []
        for f in custom:
            n = f.get("name", "").lower()
            if any(kw in n for kw in keywords):
                matches.append(f)
        field_matches[need_label] = matches
        if matches:
            for m in matches[:3]:
                print(f"  {need_label:25s} → {m['id']:25s} {m['name'][:55]}")
            if len(matches) > 3:
                print(f"  {'':25s}   ... ({len(matches)-3} more)")
        else:
            print(f"  {need_label:25s} → ✗ NO MATCH (need to create)")

    # === Issue types ===
    print("\n[2/2] Inventorying issue types ...")
    code, body = get("/rest/api/3/issuetype")
    if code != 200:
        print(f"  ✗ FAIL · {code}"); return
    types = body
    type_matches = {}
    print(f"\n  {'NEED':20s} → {'CANDIDATE'}")
    print(f"  {'-'*20}   {'-'*40}")
    for need in ISSUE_TYPE_NEEDS:
        matches = [t for t in types if t.get("name", "").lower() == need.lower()
                   or need.lower() in t.get("name", "").lower()]
        type_matches[need] = matches
        if matches:
            for m in matches[:2]:
                print(f"  {need:20s} → id={m['id']}  name={m['name']}  subtask={m.get('subtask', False)}")
        else:
            print(f"  {need:20s} → ✗ NO MATCH (need to create)")

    # Show ALL existing issue types for reference
    print(f"\n  All {len(types)} unique issue types:")
    for t in sorted(types, key=lambda x: x.get("name", "")):
        print(f"    · id={t['id']:6s}  {t['name']:35s}  hierarchy: {t.get('hierarchyLevel', '?')}")

    # Save checkpoint
    checkpoint_save("01c_inventory", {
        "fields_count": len(custom),
        "fields_to_reuse": {k: [{"id": f["id"], "name": f["name"]} for f in v[:2]]
                            for k, v in field_matches.items() if v},
        "fields_to_create": [k for k, v in field_matches.items() if not v],
        "issue_types_to_reuse": {k: [{"id": t["id"], "name": t["name"]} for t in v[:2]]
                                  for k, v in type_matches.items() if v},
        "issue_types_to_create": [k for k, v in type_matches.items() if not v],
    })
    print("\n  ✓ checkpoint saved · .checkpoints/01c_inventory.json")
    print("\nPHASE 1c COMPLETE.")


if __name__ == "__main__":
    main()
