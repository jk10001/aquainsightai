#!/usr/bin/env bash
set -euo pipefail

# Clean temporary user profile
PROFILE_DIR=/tmp/lo_profile
mkdir -p "$PROFILE_DIR"

LOG=/var/log/soffice.log
: > "$LOG"

echo "[start-lo] Starting LibreOffice headless listener..." | tee -a "$LOG"

/usr/bin/soffice \
  --headless --norestore --nolockcheck --nodefault --nofirststartwizard \
  -env:UserInstallation=file://$PROFILE_DIR \
  --accept="socket,host=0.0.0.0,port=2002;urp;StarOffice.ComponentContext" \
  >> "$LOG" 2>&1 &

SOFFICE_PID=$!
echo "[start-lo] soffice PID: $SOFFICE_PID" | tee -a "$LOG"

# If soffice exits, the container will exit with the same code
wait "$SOFFICE_PID"
EXIT_CODE=$?
echo "[start-lo] soffice exited with code $EXIT_CODE" | tee -a "$LOG"
exit "$EXIT_CODE"
