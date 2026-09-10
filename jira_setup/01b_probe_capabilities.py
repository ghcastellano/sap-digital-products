"""Probe what we can/can't do · list projects · test project creation · check project types."""
import sys
from client import get, post, checkpoint_save


def main():
    print("=" * 70)
    print("PHASE 1b · PROBE CAPABILITIES")
    print("=" * 70)

    # 1. List existing projects (paginated · just first page is enough to confirm access)
    print("\n[1/4] GET /rest/api/3/project/search ...")
    code, body = get("/rest/api/3/project/search", query={"maxResults": 50})
    if code != 200:
        print(f"  ✗ FAIL · {code} · {body}")
        sys.exit(1)
    projects = body.get("values", [])
    total = body.get("total", len(projects))
    print(f"  ✓ {len(projects)} projects accessible (total {total})")
    for p in projects[:10]:
        print(f"    · {p['key']:12s}  {p['name'][:50]:50s}  type: {p.get('projectTypeKey')}  style: {p.get('style', '?')}")

    # Check if our target key DGTLPROD already exists
    target = next((p for p in projects if p["key"] == "DGTLPROD"), None)
    if target:
        print(f"\n  ! Project DGTLPROD ALREADY EXISTS · id={target['id']}")
    else:
        print(f"\n  · Project DGTLPROD does NOT exist yet · safe to create")

    # 2. Get project types available (Cloud has scrum / kanban / business / service-desk · team-managed vs company-managed)
    print("\n[2/4] GET /rest/api/3/project/type ...")
    code, body = get("/rest/api/3/project/type")
    if code == 200 and isinstance(body, list):
        for pt in body:
            print(f"    · {pt.get('key'):20s}  {pt.get('formattedKey', pt.get('description', ''))[:60]}")
    else:
        print(f"  ! ({code}) {body}")

    # 3. List existing custom fields (so we know what's there + don't recreate)
    print("\n[3/4] GET /rest/api/3/field  (filter to CUSTOM only) ...")
    code, body = get("/rest/api/3/field")
    if code == 200 and isinstance(body, list):
        custom = [f for f in body if f.get("custom")]
        print(f"  ✓ {len(custom)} custom fields exist globally (showing 10)")
        for f in custom[:10]:
            print(f"    · {f['id']:35s}  {f['name']}")
        # Check if any of our planned fields already exist
        planned_names = ["Value Stream Destination", "OPCO", "Customer Impact Tier", "Pain Score",
                         "BV Score", "UX Score", "Safety Score", "Effort Score"]
        existing_match = []
        for f in custom:
            for n in planned_names:
                if n.lower() in f.get("name", "").lower():
                    existing_match.append(f)
                    break
        if existing_match:
            print(f"\n  ! WARNING · {len(existing_match)} fields with similar names already exist:")
            for f in existing_match:
                print(f"    · {f['id']}  {f['name']}")
    else:
        print(f"  ! ({code}) {body}")

    # 4. Probe issue types
    print("\n[4/4] GET /rest/api/3/issuetype ...")
    code, body = get("/rest/api/3/issuetype")
    if code == 200 and isinstance(body, list):
        unique = {}
        for it in body:
            unique[it["name"]] = it
        print(f"  ✓ {len(unique)} unique issue types globally:")
        for name, it in sorted(unique.items()):
            print(f"    · {it['id']:6s}  {name}  (subtask={it.get('subtask', False)})")
    else:
        print(f"  ! ({code}) {body}")

    # Save checkpoint
    checkpoint_save("01b_probe_capabilities", {
        "existing_projects_count": total,
        "target_dgtlprod_exists": target is not None,
        "target_project_id": target.get("id") if target else None,
        "custom_fields_count": len(custom) if isinstance(body, list) else 0,
    })
    print("\n  ✓ checkpoint saved")
    print("\nPHASE 1b COMPLETE.")


if __name__ == "__main__":
    main()
