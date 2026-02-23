#!/bin/bash
# Run Brandenburg patch: auto-record to ~/Desktop/MyBrandenburg.wav (250 sec), then open it.
cd "$(dirname "$0")"
SCLANG="/Applications/SuperCollider.app/Contents/MacOS/sclang"
if [ ! -x "$SCLANG" ]; then
  echo "SuperCollider not found at $SCLANG"
  exit 1
fi
echo "Starting Brandenburg — recording 250s to Desktop/MyBrandenburg.wav, then opening."
exec "$SCLANG" run_brandenburg_midi_liquid_dnb_v3.scd
