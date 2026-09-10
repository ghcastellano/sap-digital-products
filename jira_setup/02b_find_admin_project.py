"""Phase 2b · Find existing projects where the user is admin OR find a sandbox to use.

Strategy: list all 136 projects, check which ones we can administer,
or look for "demand", "digital", "sandbox" candidates we can reuse.
"""
from client import get, checkpoint_save


def main():
    print("=" * 70)
    print("PHASE 2b · FIND EXISTING PROJECT · admin OR demand-related")
    print("=" * 70)

    # Get all projects (paginated)
    all_projects = []
    start_at = 0
    while True:
        code, body = get("/rest/api/3/project/search",
                         query={"maxResults": 100, "startAt": start_at,
                                "expand": "lead,description,projectKeys,permissions"})
        if code != 200: break
        all_projects.extend(body.get("values", []))
        if body.get("isLast", True): break
        start_at += 100
    print(f"\n  Total projects: {len(all_projects)}")

    # Filter: where current user is project lead
    user_chk = __import__("client").checkpoint_load("01_validate_access")
    my_id = user_chk["user"]["accountId"]
    led_by_me = [p for p in all_projects
                 if (p.get("lead") or {}).get("accountId") == my_id]

    print(f"\n[1/3] Projects where YOU are lead ({len(led_by_me)}):")
    for p in led_by_me[:10]:
        print(f"    · {p['key']:12s}  {p['name'][:45]}  type:{p.get('projectTypeKey')} style:{p.get('style')}")

    # Search for demand/digital-related projects (potential reuse)
    print(f"\n[2/3] Demand/Digital/Products-related projects:")
    keywords = ["demand", "digital", "products", "sandbox", "test", "tom",
                "intake", "portfolio", "operating model"]
    matches = []
    for p in all_projects:
        n = (p.get("name", "") + " " + p.get("key", "")).lower()
        for kw in keywords:
            if kw in n:
                matches.append((p, kw))
                break
    for p, kw in matches[:20]:
        print(f"    · [{kw:10s}] {p['key']:14s}  {p['name'][:55]}  type:{p.get('projectTypeKey')}")

    # Test which projects user can edit (via mypermissions per project)
    print(f"\n[3/3] Probe edit/admin permission on candidates ...")
    test_keys = [p["key"] for p in led_by_me[:5]] + [p[0]["key"] for p in matches[:8]]
    test_keys = list(dict.fromkeys(test_keys))[:10]  # dedupe + cap
    admin_projects = []
    for k in test_keys:
        code, body = get("/rest/api/3/mypermissions",
                         query={"projectKey": k,
                                "permissions": "ADMINISTER_PROJECTS,CREATE_ISSUES,EDIT_ISSUES"})
        if code != 200: continue
        perms = body.get("permissions", {})
        admin = perms.get("ADMINISTER_PROJECTS", {}).get("havePermission", False)
        create = perms.get("CREATE_ISSUES", {}).get("havePermission", False)
        edit = perms.get("EDIT_ISSUES", {}).get("havePermission", False)
        marker = "ADMIN" if admin else ("CREATE+EDIT" if create and edit else "BROWSE")
        print(f"    · {k:14s}  → {marker}")
        if admin:
            admin_projects.append(k)

    checkpoint_save("02b_find_admin_project", {
        "led_by_me": [p["key"] for p in led_by_me],
        "demand_related": [p["key"] for p, _ in matches[:20]],
        "admin_projects": admin_projects,
    })
    print(f"\n  ✓ checkpoint saved")
    print(f"\n  Recommendation:")
    if admin_projects:
        print(f"  · Use one of these projects (you are admin): {admin_projects[:3]}")
    elif led_by_me:
        print(f"  · Lead of: {[p['key'] for p in led_by_me[:3]]}  (may not equal admin)")
    else:
        print(f"  · Request 'create next-gen projects' permission from Jira admin")


if __name__ == "__main__":
    main()
