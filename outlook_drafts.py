#!/usr/bin/env python3
"""
Outlook Draft Generator — brouillons-outlook-yann
Cron: 0 9,18 * * *  (9h et 18h tous les jours)

Usage:
    python outlook_drafts.py           # crée les brouillons
    python outlook_drafts.py --dry-run # affiche sans créer
"""

import os
import sys
from dotenv import load_dotenv

load_dotenv()

from src.outlook_drafter import run_drafts


def main():
    dry_run = "--dry-run" in sys.argv
    user_email = os.environ.get("SENDER_EMAIL")
    user_name = os.environ.get("USER_DISPLAY_NAME", "Yann")

    if not user_email:
        print("❌ SENDER_EMAIL non défini dans .env")
        sys.exit(1)

    print(f"{'='*60}")
    print(f"Brouillons Outlook — {user_name} <{user_email}>")
    if dry_run:
        print("  [MODE DRY-RUN — aucun brouillon ne sera créé]")

    created = run_drafts(user_email=user_email, user_name=user_name, dry_run=dry_run)

    print(f"\n{'='*60}")
    print(f"✓ {len(created)} brouillon(s) créé(s) dans Outlook")
    for d in created:
        print(f"  • {d['subject']} ← {d['from_name'] or d['from']}")
        print(f"    Raison : {d['reason']}")


if __name__ == "__main__":
    main()
