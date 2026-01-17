#!/usr/bin/env bash

set -euo pipefail

run_calendar_tests() {
  if skip "calendar"; then
    echo "==> calendar (skipped)"
    return 0
  fi

  read -r START END DAY1 DAY2 <<<"$($PY - <<'PY'
import datetime
now=datetime.datetime.now(datetime.timezone.utc).replace(minute=0, second=0, microsecond=0)
start=now + datetime.timedelta(hours=1)
end=start + datetime.timedelta(hours=1)
print(start.strftime('%Y-%m-%dT%H:%M:%SZ'), end.strftime('%Y-%m-%dT%H:%M:%SZ'), start.strftime('%Y-%m-%d'), (start+datetime.timedelta(days=1)).strftime('%Y-%m-%d'))
PY
)"

  run_required "calendar" "calendar list" gog calendar calendars --json --max 1 >/dev/null
  run_required "calendar" "calendar acl" gog calendar acl primary --json --max 1 >/dev/null
  run_required "calendar" "calendar colors" gog calendar colors --json >/dev/null
  run_required "calendar" "calendar time" gog calendar time --json >/dev/null

  local ev_json ev_id
  ev_json=$(gog calendar create primary --summary "gogcli-smoke-$TS" --from "$START" --to "$END" --location "Test" --send-updates none --json)
  ev_id=$(extract_id "$ev_json")
  [ -n "$ev_id" ] || { echo "Failed to parse calendar event id" >&2; exit 1; }

  run_required "calendar" "calendar event get" gog calendar event primary "$ev_id" --json >/dev/null
  run_required "calendar" "calendar propose-time" gog calendar propose-time primary "$ev_id" --json >/dev/null
  run_required "calendar" "calendar update" gog calendar update primary "$ev_id" --summary "gogcli-smoke-updated-$TS" --json >/dev/null
  run_required "calendar" "calendar events list" gog calendar events primary --from "$START" --to "$END" --json --max 5 >/dev/null
  run_required "calendar" "calendar search" gog calendar search "gogcli-smoke" --from "$START" --to "$END" --json --max 5 >/dev/null
  run_required "calendar" "calendar freebusy" gog calendar freebusy primary --from "$START" --to "$END" --json >/dev/null
  run_required "calendar" "calendar conflicts" gog calendar conflicts --from "$START" --to "$END" --json >/dev/null

  if [ -n "${GOG_LIVE_CALENDAR_RESPOND:-}" ]; then
    run_optional "calendar-respond" "calendar respond" gog calendar respond primary "$ev_id" --status accepted --json >/dev/null
  else
    echo "==> calendar respond (skipped; needs invite from another account)"
  fi

  run_required "calendar" "calendar delete event" gog calendar delete primary "$ev_id" --force >/dev/null

  if ! skip "calendar-enterprise"; then
    run_optional "calendar-enterprise" "calendar focus-time" gog calendar create primary --event-type focus-time --from "$START" --to "$END" --json >/dev/null 2>&1 || true
    run_optional "calendar-enterprise" "calendar out-of-office" gog calendar create primary --event-type out-of-office --from "$DAY1" --to "$DAY2" --all-day --json >/dev/null 2>&1 || true
    run_optional "calendar-enterprise" "calendar working-location" gog calendar create primary --event-type working-location --working-location-type office --working-office-label "HQ" --from "$DAY1" --to "$DAY2" --json >/dev/null 2>&1 || true
  fi

  if [ -n "${GOG_LIVE_CALENDAR_RECURRENCE:-}" ]; then
    local rec_json rec_id
    rec_json=$(gog calendar create primary --summary "gogcli-recurring-$TS" --from "$START" --to "$END" --rrule "RRULE:FREQ=DAILY;COUNT=2" --reminder "popup:30m" --json)
    rec_id=$(extract_id "$rec_json")
    if [ -n "$rec_id" ]; then
      run_required "calendar" "calendar delete recurring" gog calendar delete primary "$rec_id" --force >/dev/null
    fi
  else
    echo "==> calendar recurrence/reminders (skipped; set GOG_LIVE_CALENDAR_RECURRENCE=1)"
  fi

  if [ -n "${GOG_LIVE_GROUP_EMAIL:-}" ] && ! is_consumer_account "$ACCOUNT"; then
    run_optional "calendar-team" "calendar team" gog calendar team "$GOG_LIVE_GROUP_EMAIL" --json --max 5 >/dev/null
  fi

  if is_consumer_account "$ACCOUNT"; then
    echo "==> calendar users (skipped; Workspace only)"
  else
    run_optional "calendar-users" "calendar users list" gog calendar users --json --max 1 >/dev/null
  fi
}
