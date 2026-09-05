#!/bin/sh
# Capture one store screenshot from the live screen: the 1280×800 logical-pixel
# region whose top-left is the Sheets viewport's top-left, saved as
# docs/store/assets/screenshot-N.png. On a Retina display the raw capture is
# 2560×1600 and is resampled to exactly 1280×800, which the store requires.
#
#   sh docs/store/capture.sh N X Y
#
# X, Y = the viewport's screen position in logical pixels. Get them from the
# spreadsheet tab's console: X = window.screenX, Y = window.screenY +
# (window.outerHeight - window.innerHeight); make the viewport 1280×800 first
# (window.innerWidth/innerHeight). Open the sheet with ?hl=en for the English
# listing's shots and ?hl=ja for the Japanese ones; move the mouse off the
# viewport before capturing. Verified 2026-09-05 with X=0 Y=216.
set -e
n=${1:?screenshot number}; x=${2:?viewport screen X}; y=${3:?viewport screen Y}
out="$(dirname "$0")/assets/screenshot-$n.png"
tmp=$(mktemp -t scopion-shot).png
screencapture -x -R"$x,$y,1280,800" "$tmp"
sips -z 800 1280 "$tmp" --out "$out" >/dev/null
rm -f "$tmp"
sips -g pixelWidth -g pixelHeight "$out" | tail -2
