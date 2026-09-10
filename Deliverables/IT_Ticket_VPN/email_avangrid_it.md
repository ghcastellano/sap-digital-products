# Email to Avangrid IT — Citrix Gateway access blocked

**To:** [Avangrid IT Service Desk]
**Cc:** Carlos Garbin (IT Business Applications Director); Vitor Cruz (VP IT Business Solutions Delivery)
**Subject:** Citrix Gateway access blocked — user AMER\E738953 — NSGAuthStatus 11 / possible password expiry?

---

Hi IT team,

I am **Gustavo Castellano**, NTT DATA consultant engaged on the **SAP Digital Products transformation project** sponsored by Carlos Garbin and Vitor Cruz. I need remote access to the Avangrid Citrix environment to continue the Wave 1 discovery work, and my login has been failing since this morning (2026-04-20).

## What is happening

When I try to authenticate at `https://www5avgec.avangrid.com` (or `www5avgecfs.avangrid.com`), the gateway displays the generic message **"Tente novamente mais tarde ou entre em contato com o suporte técnico"** immediately after I submit username/password — **before any MFA prompt appears**.

The NetScaler Gateway client logs show:

```
NSGAuthStatus(rawValue: 11)
NSGAuthService - Gateway discovery error - 13
Discover call returned NOT OK : 302   (repeating)
Gateway unreachable  (intermittent)
URL: https://www5avgec.avangrid.com/cgi/login
Timestamps: 2026-04-20 09:54 – 10:14 GMT-3
```

## Diagnostics already performed on my side

To save you time, I already validated the following from my machine:

| Check | Result |
|---|---|
| DNS resolution of `www5avgec.avangrid.com` | OK — resolves to `157.225.204.50` |
| TCP/TLS handshake on :443 | OK — valid cert (DigiCert, Avangrid Service Company, expiring Nov 2026) |
| Login page (`/logon/LogonPoint/tmindex.html`) | Loads (HTTP 200) |
| System clock drift | +0.04s (irrelevant) |
| My MFA token | Already enrolled |

So the gateway is reachable and my credentials are reaching the server — the rejection is happening at the authentication layer, not the network layer.

## My hypothesis — could this be an expired AD password?

Based on the symptom (error **before** MFA is requested), my first suspicion is that my **AD password has expired**. Avangrid rotates passwords periodically, and because I mostly work with NTT-managed systems, I may have missed the rotation window.

Other possibilities I would like you to confirm or rule out:

1. **AD password expired** — can you check the password age for `AMER\E738953` and, if needed, trigger a reset flow?
2. **Account lockout from failed attempts** — I did retry a few times this morning; has the account been temporarily locked?
3. **Gateway policy / device compliance (EPA)** — my Mac is NTT-issued (not Avangrid-managed, no MDM enrollment). Is EPA blocking non-corporate devices at the authentication stage, or is BYOD access allowed with an exception for consultants?
4. **Group membership for the external Citrix Gateway** — could the project team access policy have been rolled back or not yet applied to my user?

## What I need

- Confirmation of which of the above (1–4) is causing the block.
- If it is a password expiry: the reset procedure (self-service URL or ticket path).
- If it is EPA / device compliance: the correct path for a consulting partner (is there a BYOD enrollment, a laptop provisioning workflow, or a web-only portal that bypasses EPA?).

I am blocked from the project workspace until this is resolved. Any prioritization you can give this would be appreciated — Carlos and Vitor are copied for visibility.

## Reference data for ticket

- **User ID:** `AMER\E738953`
- **Email:** ghcastellano@[…]
- **Gateway URL:** `https://www5avgec.avangrid.com`
- **Error code:** `NSGAuthStatus 11 / Gateway discovery error 13`
- **My public IP at the time of failure:** `201.43.163.121` (Brazil)
- **Failure timestamps:** `2026-04-20 09:54:48 – 10:14:43 GMT-3`
- **Client:** Citrix Workspace / NetScaler Gateway on macOS

Thanks for the help,

**Gustavo Castellano**
NTT DATA Consulting — SAP Digital Products (Avangrid)
