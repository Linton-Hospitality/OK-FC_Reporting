#!/bin/bash

CD_PATH="/Users/stephcleung/Library/Mobile Documents/com~apple~CloudDocs/Desktop/Ferncrest/_Franchise/Franchisees/2_OK-FC/AssetManagement"

cd "$CD_PATH"

# On the first Monday of the month, push both actuals + pace
# All other Mondays, push pace only
DAY_OF_MONTH=$(date '+%d')

if [ "$DAY_OF_MONTH" -le 7 ]; then
    # First Monday of the month — push both files
    git add ok-fc_occupancy_2026ytd.xlsx ok-fc_occupancy_pace_150day.xlsx ferncrest_ok-fc_v2.xlsx
    PUSH_LABEL="actuals + pace"
else
    # Regular Monday — pace only
    git add ok-fc_occupancy_pace_150day.xlsx ferncrest_ok-fc_v2.xlsx
    PUSH_LABEL="pace only"
fi

# Check if there's anything to commit
if git diff --cached --quiet; then
    osascript -e 'display notification "No changes to push — files already up to date." with title "Ferncrest Weekly Push" subtitle "OK-FC"'
    exit 0
fi

git commit -m "Weekly data update $(date '+%Y-%m-%d') ($PUSH_LABEL)"

if git push; then
    osascript -e "display notification \"Pushed $PUSH_LABEL to GitHub successfully.\" with title \"Ferncrest Weekly Push ✅\" subtitle \"OK-FC\""
else
    osascript -e "display notification \"Push failed ($PUSH_LABEL) — check your internet and push manually before 10am.\" with title \"Ferncrest Weekly Push ❌\" subtitle \"OK-FC\""
fi
