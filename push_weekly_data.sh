#!/bin/bash

CD_PATH="/Users/stephcleung/Library/Mobile Documents/com~apple~CloudDocs/Desktop/Ferncrest/_Franchise/Franchisees/2_OK-FC/AssetManagement"

cd "$CD_PATH"

git add ok-fc_occupancy_2026ytd.xlsx ok-fc_occupancy_pace_150day.xlsx ferncrest_ok-fc_v2.xlsx

# Check if there's anything to commit
if git diff --cached --quiet; then
    osascript -e 'display notification "No changes to push — files already up to date." with title "Ferncrest Weekly Push" subtitle "OK-FC"'
    exit 0
fi

git commit -m "Weekly data update $(date '+%Y-%m-%d')"

if git push; then
    osascript -e 'display notification "Excel files pushed to GitHub successfully." with title "Ferncrest Weekly Push ✅" subtitle "OK-FC"'
else
    osascript -e 'display notification "Push failed — check your internet connection and push manually." with title "Ferncrest Weekly Push ❌" subtitle "OK-FC"'
fi
