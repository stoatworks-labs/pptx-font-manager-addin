#!/bin/bash
# Sideload into Mac PowerPoint: drop the manifest in the wef folder, then quit
# and reopen PowerPoint. The dev server (npm run dev) must be running first, and
# the dev cert trusted (npm run certs).
#
# NOTE: the custom Home-tab button may only appear after you launch the add-in
# once from  Home ▸ Add-ins ▸ Developer Add-ins ▸ Font Manager  — a known
# wef-sideload quirk. After that first launch the ribbon button sticks.
set -euo pipefail
HERE="$(cd "$(dirname "$0")/.." && pwd)"
WEF="$HOME/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef"
mkdir -p "$WEF"
cp "$HERE/manifest.xml" "$WEF/pptx-font-manager.manifest.xml"
echo "Sideloaded → $WEF/pptx-font-manager.manifest.xml"
echo "Quit PowerPoint (Cmd-Q) and reopen it. Then: Home ▸ Add-ins ▸ Developer Add-ins ▸ Font Manager."
