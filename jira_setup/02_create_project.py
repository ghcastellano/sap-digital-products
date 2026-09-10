"""Phase 2 · Create Team-managed Software/Kanban project DGTLPROD.

Team-managed projects have local fields/workflows/screens (project-scoped),
which avoids needing Site Admin to add global custom fields.
"""
import json
import sys
from client import get, post, find_project_by_key, checkpoint_save, checkpoint_load
from config import JIRA_PROJECT_KEY, JIRA_PROJECT_NAME, VALUE_STREAMS


# Available templates in Avangrid Cloud (probed via /rest/api/3/project/type)
# Software templates (team-managed kanban):
KANBAN_TEMPLATE = "com.pyxis.greenhopper.jira:gh-simplified-agility-kanban"
SCRUM_TEMPLATE  = "com.pyxis.greenhopper.jira:gh-simplified-agility-scrum"


def main():
    print("=" * 70)
    print(f"PHASE 2 · CREATE PROJECT · {JIRA_PROJECT_KEY} ({JIRA_PROJECT_NAME})")
    print("=" * 70)

    # Resume guard · skip if already exists
    existing = find_project_by_key(JIRA_PROJECT_KEY)
    if existing:
        print(f"\n  ✓ Project {JIRA_PROJECT_KEY} already exists · id={existing['id']}")
        checkpoint_save("02_create_project", {
            "id": existing["id"],
            "key": existing["key"],
            "name": existing["name"],
            "projectTypeKey": existing.get("projectTypeKey"),
            "style": existing.get("style"),
        })
        return

    # Get my accountId for lead
    user_chk = checkpoint_load("01_validate_access")
    if not user_chk:
        print("  ✗ Run 01_validate_access.py first")
        sys.exit(1)
    lead_id = user_chk["user"]["accountId"]

    # === Create Team-managed Kanban Software Project ===
    print(f"\n[1/2] POST /rest/api/3/project · template=kanban (team-managed) ...")
    body = {
        "key": JIRA_PROJECT_KEY,
        "name": JIRA_PROJECT_NAME,
        "description": (
            "Avangrid Digital Products · TOM (Target Operating Model) · "
            "demand intake · triage · portfolio governance · squad delivery. "
            "Covers CS (Customer Services) + ITBA (IT Business Applications)."
        ),
        "projectTypeKey": "software",
        "projectTemplateKey": KANBAN_TEMPLATE,
        "leadAccountId": lead_id,
        "assigneeType": "PROJECT_LEAD",
    }
    code, resp = post("/rest/api/3/project", body)
    if code >= 300:
        print(f"  ✗ FAIL · {code} · {json.dumps(resp, indent=2, ensure_ascii=False)[:600]}")
        sys.exit(1)
    print(f"  ✓ Project created · id={resp['id']}  key={resp['key']}  self={resp.get('self', '')}")
    project_id = resp["id"]
    project_key = resp["key"]

    # === Create Components per Value Stream ===
    print(f"\n[2/2] Creating {len(VALUE_STREAMS)} components (one per VS) ...")
    components = []
    for vs_name, meta in VALUE_STREAMS.items():
        cbody = {
            "name": vs_name,
            "description": f"{meta['area']} · Director: {meta['director']}",
            "project": project_key,
            "assigneeType": "PROJECT_DEFAULT",
        }
        ccode, cresp = post("/rest/api/3/component", cbody)
        if ccode >= 300:
            print(f"    ✗ {vs_name:35s}  · {ccode} · {cresp}")
        else:
            print(f"    ✓ {vs_name:35s}  · id={cresp['id']}  area={meta['area']}")
            components.append({
                "id": cresp["id"], "name": vs_name, "area": meta["area"],
                "director": meta["director"]
            })

    # Save checkpoint
    checkpoint_save("02_create_project", {
        "id": project_id,
        "key": project_key,
        "name": JIRA_PROJECT_NAME,
        "lead_accountId": lead_id,
        "components": components,
    })
    print(f"\n  ✓ checkpoint saved · {len(components)}/{len(VALUE_STREAMS)} components created")
    print(f"\nPHASE 2 PASSED.")


if __name__ == "__main__":
    main()
