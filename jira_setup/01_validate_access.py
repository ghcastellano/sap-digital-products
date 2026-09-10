"""Phase 1 · Validate API access · check user info, server reachability, basic permissions."""
import sys
from client import get, post, whoami, checkpoint_save


def main():
    print("=" * 70)
    print("PHASE 1 · VALIDATE API ACCESS")
    print("=" * 70)

    # 1. Who am I?
    print("\n[1/3] GET /rest/api/3/myself ...")
    code, body = whoami()
    if code != 200:
        print(f"  ✗ FAIL · {code} · {body}")
        sys.exit(1)
    print(f"  ✓ Authenticated as: {body.get('displayName')} ({body.get('emailAddress')})")
    print(f"    accountId: {body.get('accountId')}")
    user_info = body

    # 2. Server info
    print("\n[2/3] GET /rest/api/3/serverInfo ...")
    code, body = get("/rest/api/3/serverInfo")
    if code != 200:
        print(f"  ✗ FAIL · {code} · {body}")
        sys.exit(1)
    print(f"  ✓ Server: {body.get('serverTitle')}  ·  version {body.get('version')}")
    print(f"    deploymentType: {body.get('deploymentType')}  ·  baseUrl: {body.get('baseUrl')}")
    server_info = body

    # 3. Check permissions (Cloud-friendly perm keys)
    print("\n[3/3] GET /rest/api/3/mypermissions  (ADMINISTER · ADMINISTER_PROJECTS · CREATE_ISSUES) ...")
    perm_keys = ["ADMINISTER", "ADMINISTER_PROJECTS", "CREATE_ISSUES", "BROWSE_PROJECTS"]
    code, body = get("/rest/api/3/mypermissions",
                     query={"permissions": ",".join(perm_keys)})
    if code != 200:
        # Some Cloud sites reject unknown keys · log and continue (auth works · permissions verified later per-call)
        print(f"  ! Skipping permission probe ({code}) · will rely on per-call response")
    else:
        perms = body.get("permissions", {})
        for perm_name, perm_data in perms.items():
            has_perm = perm_data.get("havePermission", False)
            marker = "✓" if has_perm else "·"
            print(f"  {marker} {perm_name}: {has_perm}")

    # Save checkpoint
    checkpoint_save("01_validate_access", {
        "user": {
            "accountId": user_info.get("accountId"),
            "email": user_info.get("emailAddress"),
            "displayName": user_info.get("displayName"),
        },
        "server": {
            "title": server_info.get("serverTitle"),
            "baseUrl": server_info.get("baseUrl"),
            "deploymentType": server_info.get("deploymentType"),
        },
        "permissions": {k: v.get("havePermission") for k, v in perms.items()},
    })
    print("\n  ✓ checkpoint saved · .checkpoints/01_validate_access.json")
    print("\nPHASE 1 PASSED.")


if __name__ == "__main__":
    main()
