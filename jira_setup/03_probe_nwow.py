"""Phase 3 · Probe NWOW sandbox

Goal: map exactly what NWOW exposes so we can decide what part of the TOM
is feasible to prototype here (no admin = no new fields/types/workflow).

Reports:
- Project meta (lead, type, style)
- Issue types available in NWOW
- Components in NWOW
- Versions in NWOW
- Statuses (workflow) per issue type
- Which custom fields from our REUSE list are actually visible on NWOW
  create screen (sampled by trying createmeta)
"""
from client import get, checkpoint_save, safe_print

NWOW = "NWOW"

REUSE_FIELDS = {
    "OpCo": "customfield_10800",
    "Business Value": "customfield_11440",
    "Acceptance Criteria": "customfield_10795",
    "Customer Impact": "customfield_11061",
    "Demand Type": "customfield_11682",
    "T-Shirt-Sizing": "customfield_10850",
    "Safety": "customfield_11438",
    "Effort": "customfield_11282",
    "Discovery": "customfield_11646",
    "Epic Link": "customfield_10014",
}


def main():
    report = {"project": NWOW, "issue_types": [], "components": [],
              "versions": [], "statuses": {}, "createmeta_fields": {},
              "reuse_field_visibility": {}}

    print(f"\n=== Probing {NWOW} ===\n")

    # 1. project meta
    code, proj = get(f"/rest/api/3/project/{NWOW}")
    safe_print(f"GET project {NWOW}", code, proj)
    if code >= 300:
        return
    report["lead"] = proj.get("lead", {}).get("displayName")
    report["projectTypeKey"] = proj.get("projectTypeKey")
    report["style"] = proj.get("style")
    report["name"] = proj.get("name")
    report["id"] = proj.get("id")
    print(f"    lead     : {report['lead']}")
    print(f"    type     : {report['projectTypeKey']}")
    print(f"    style    : {report['style']}")

    # 2. issue types in NWOW
    its = proj.get("issueTypes", [])
    for it in its:
        report["issue_types"].append({
            "id": it.get("id"),
            "name": it.get("name"),
            "subtask": it.get("subtask", False),
        })
    print(f"\n  issue types ({len(its)}):")
    for it in report["issue_types"]:
        print(f"    · {it['name']} (id={it['id']}, subtask={it['subtask']})")

    # 3. components
    code, comps = get(f"/rest/api/3/project/{NWOW}/components")
    safe_print(f"GET components", code, comps)
    if code < 300 and isinstance(comps, list):
        report["components"] = [{"id": c.get("id"), "name": c.get("name")} for c in comps]
        print(f"\n  components ({len(comps)}):")
        for c in report["components"][:25]:
            print(f"    · {c['name']}")
        if len(comps) > 25:
            print(f"    … {len(comps) - 25} more")

    # 4. versions
    code, vers = get(f"/rest/api/3/project/{NWOW}/versions")
    if code < 300 and isinstance(vers, list):
        report["versions"] = [{"id": v.get("id"), "name": v.get("name")} for v in vers]
        print(f"\n  versions: {len(report['versions'])}")

    # 5. statuses per issue type
    code, statuses = get(f"/rest/api/3/project/{NWOW}/statuses")
    if code < 300 and isinstance(statuses, list):
        for entry in statuses:
            it_name = entry.get("name")
            st_list = [s.get("name") for s in entry.get("statuses", [])]
            report["statuses"][it_name] = st_list
        print(f"\n  workflow statuses per issue type:")
        for it_name, st_list in report["statuses"].items():
            print(f"    {it_name}:")
            for s in st_list:
                print(f"        → {s}")

    # 6. create-meta to see which fields are exposed on create screen
    print(f"\n  create-meta fields per issue type:")
    for it in report["issue_types"]:
        if it.get("subtask"):
            continue
        code, meta = get(
            f"/rest/api/3/issue/createmeta/{NWOW}/issuetypes/{it['id']}",
            query={"maxResults": 200},
        )
        if code >= 300:
            print(f"    ✗ {it['name']}: HTTP {code}")
            continue
        fields = meta.get("fields", []) if isinstance(meta, dict) else []
        names = []
        for f in fields:
            fid = f.get("fieldId") or f.get("key")
            fname = f.get("name")
            req = f.get("required", False)
            names.append({"id": fid, "name": fname, "required": req})
        report["createmeta_fields"][it["name"]] = names
        print(f"    {it['name']}: {len(names)} fields")
        for n in names:
            mark = "★" if n["required"] else " "
            print(f"      {mark} {n['name']}  ({n['id']})")

    # 7. cross-check our REUSE list against NWOW createmeta
    print(f"\n  REUSE field visibility in NWOW:")
    flat = set()
    for it_name, lst in report["createmeta_fields"].items():
        for f in lst:
            flat.add(f["id"])
    for label, fid in REUSE_FIELDS.items():
        seen = fid in flat
        report["reuse_field_visibility"][label] = seen
        print(f"    {'✓' if seen else '✗'}  {label}  ({fid})")

    checkpoint_save("03_probe_nwow", report)
    print(f"\n✅ Saved to .checkpoints/03_probe_nwow.json")


if __name__ == "__main__":
    main()
