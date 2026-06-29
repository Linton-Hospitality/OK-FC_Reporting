#!/bin/bash
# Syncs the latest Promised Land Cloudbeds exports from Google Drive into this
# repo and pushes, so the Monday GitHub Actions run has fresh data instead of
# a stale committed copy. Run this anytime after a new export lands in Drive —
# ideally before Monday 10am Pacific.
set -euo pipefail

GDRIVE="/Users/stephcleung/Library/CloudStorage/GoogleDrive-stephanie@lintonhospitality.com/Shared drives/Ferncrest/1_Locations/01_PA-PL/6_Reports_PA_PL/AssetManagement"
REPO_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"

FILES=(
  "PA-PL_occupancy_2026ytd.xlsx"
  "PA-PL_occupancy_pace_150day.xlsx"
  "PA-PL_occupancy_2025.xlsx"
)

cd "$REPO_DIR"

CHANGED=0
for f in "${FILES[@]}"; do
  if [ ! -f "$GDRIVE/$f" ]; then
    echo "⚠️  Not found in Drive: $f — skipping"
    continue
  fi
  if ! cmp -s "$GDRIVE/$f" "$REPO_DIR/$f"; then
    cp "$GDRIVE/$f" "$REPO_DIR/$f"
    echo "✅ Synced: $f"
    CHANGED=1
  else
    echo "·  Unchanged: $f"
  fi
done

if [ "$CHANGED" -eq 0 ]; then
  echo "Nothing to sync — repo already matches Drive."
  exit 0
fi

git add "${FILES[@]}"
git commit -m "Sync Promised Land Cloudbeds export from Drive — $(date +%Y-%m-%d)"
git pull --rebase origin main
git push

echo "✅ Pushed. Next scheduled/manual pipeline run will use this data."
