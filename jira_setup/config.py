"""Centralized config loaded from .env."""
import os
from pathlib import Path

ENV_PATH = Path(__file__).parent / ".env"

def _load_env():
    if not ENV_PATH.exists():
        raise FileNotFoundError(f".env not found at {ENV_PATH}")
    for line in ENV_PATH.read_text().splitlines():
        line = line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        k, v = line.split("=", 1)
        os.environ.setdefault(k.strip(), v.strip())

_load_env()

JIRA_EMAIL       = os.environ["JIRA_EMAIL"]
JIRA_API_TOKEN   = os.environ["JIRA_API_TOKEN"]
JIRA_BASE_URL    = os.environ["JIRA_BASE_URL"].rstrip("/")
JIRA_PROJECT_KEY = os.environ["JIRA_PROJECT_KEY"]
JIRA_PROJECT_NAME = os.environ["JIRA_PROJECT_NAME"]

# Checkpoint dir for idempotency
CHECKPOINT_DIR = Path(__file__).parent / ".checkpoints"
CHECKPOINT_DIR.mkdir(exist_ok=True)

# Value Streams (10) for components + boards
VALUE_STREAMS = {
    # ITBA (5)
    "Field Operations":              {"area": "ITBA", "director": "Carlos Garbin"},
    "Engineering & Asset Mgmt":      {"area": "ITBA", "director": "Carlos Garbin"},
    "Network Operations & Control":  {"area": "ITBA", "director": "Carlos Garbin"},
    "Corporate & Regulatory":        {"area": "ITBA", "director": "Carlos Garbin"},
    "Iberdrola Corp Services":       {"area": "ITBA", "director": "Carlos Garbin"},
    # CS (4 DVS + 1 Enabling)
    "Meter & Device Operations":     {"area": "CS",   "director": "Kumar Jaladanki"},
    "Billing Operations":            {"area": "CS",   "director": "Kumar Jaladanki"},
    "Revenue & Collections":         {"area": "CS",   "director": "Kumar Jaladanki"},
    "Customer Care & Engagement":    {"area": "CS",   "director": "Heather LaJoie"},
    "Enabling Team — Foundation":    {"area": "CS",   "director": "Kumar Jaladanki"},
}

OPCO_OPTIONS = ["UIL", "CMP", "NY", "Iberdrola", "Cross-OPCO"]
TYPE_OPTIONS = ["Production Support", "Regulatory", "Non-Regulatory"]
SIZE_OPTIONS = ["S", "M", "L"]  # XL removed
COS_OPTIONS  = ["Standard", "Fixed-Date", "Expedite", "Intangible"]
RGT_OPTIONS  = ["Run", "Grow", "Transform"]
CIT_OPTIONS  = ["P1 Critical", "P2 High", "P3 Medium", "P4 Low"]
