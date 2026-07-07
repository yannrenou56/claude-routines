#!/bin/bash
# Configure les tâches cron pour claude-routines
# Usage: bash setup_cron.sh

set -e

REPO_DIR="$(cd "$(dirname "$0")" && pwd)"
LOG_DIR="$REPO_DIR/logs"
PYTHON=$(which python3)

mkdir -p "$LOG_DIR"

echo "=== Configuration des tâches cron ==="
echo "Répertoire : $REPO_DIR"
echo "Python     : $PYTHON"

# Remove existing claude-routines cron entries
TMP=$(mktemp)
crontab -l 2>/dev/null | grep -v "claude-routines" > "$TMP" || true

# Add new entries
cat >> "$TMP" << CRON
# claude-routines — brouillons-outlook-yann
0 9,18 * * * cd "$REPO_DIR" && $PYTHON outlook_drafts.py >> "$LOG_DIR/drafts.log" 2>&1
CRON

crontab "$TMP"
rm "$TMP"

echo ""
echo "✓ Tâche cron installée :"
crontab -l | grep "claude-routines"
echo ""
echo "Logs dans : $LOG_DIR/drafts.log"
echo ""
echo "Pour tester maintenant (sans attendre 9h/18h) :"
echo "  cd $REPO_DIR && python3 outlook_drafts.py --dry-run"
