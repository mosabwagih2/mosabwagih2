#!/usr/bin/env bash
# Guard for the Pump & Punish campaign: Meta Ads write tools are pre-approved in
# .claude/settings.json so the scheduled checks can run unattended. This hook
# narrows that approval to the one campaign and to budget/status changes only.
set -euo pipefail

CAMPAIGN_ID="120249758271140590"
# Campaign, its ad set and its four ads: the only objects that may be activated.
ACTIVATABLE_IDS="120249758271140590 120249758274010590 120249758275290590 120249758275550590 120249758275710590 120249758276000590"
MIN_BUDGET=100000     # 1,000 EGP in piasters
MAX_BUDGET=10000000   # 100,000 EGP in piasters

input="$(cat)"
tool="$(jq -r '.tool_name // ""' <<<"$input")"

deny() {
  jq -n --arg reason "$1" '{hookSpecificOutput: {hookEventName: "PreToolUse", permissionDecision: "deny", permissionDecisionReason: $reason}}'
  exit 0
}

case "$tool" in
  mcp__Meta_Ads__ads_update_entity)
    entity_id="$(jq -r '.tool_input.entity_id // ""' <<<"$input")"
    entity_type="$(jq -r '.tool_input.entity_type // ""' <<<"$input")"
    [[ "$entity_id" == "$CAMPAIGN_ID" && "$entity_type" == "campaign" ]] \
      || deny "Meta guard: updates are allowed only on campaign $CAMPAIGN_ID."

    fields="$(jq -c '.tool_input.fields | if type == "string" then fromjson else . end' <<<"$input" 2>/dev/null)" \
      || deny "Meta guard: fields is not valid JSON."
    extra="$(jq -r 'keys - ["daily_budget", "status"] | join(",")' <<<"$fields")"
    [[ -z "$extra" ]] || deny "Meta guard: only daily_budget and status may change (got: $extra)."

    status="$(jq -r '.status // ""' <<<"$fields")"
    [[ -z "$status" || "$status" == "PAUSED" ]] \
      || deny "Meta guard: status may only be set to PAUSED (got: $status)."

    if jq -e 'has("daily_budget")' <<<"$fields" >/dev/null; then
      budget="$(jq -r '.daily_budget' <<<"$fields")"
      [[ "$budget" =~ ^[0-9]+$ ]] || deny "Meta guard: daily_budget must be an integer in piasters."
      (( budget >= MIN_BUDGET && budget <= MAX_BUDGET )) \
        || deny "Meta guard: daily_budget $budget is outside $MIN_BUDGET..$MAX_BUDGET."
    fi
    ;;
  mcp__Meta_Ads__ads_activate_entity)
    ids="$(jq -r '[.tool_input.entity_id] + (.tool_input.object_ids // []) | map(select(. != null)) | .[]' <<<"$input")"
    for id in $ids; do
      [[ " $ACTIVATABLE_IDS " == *" $id "* ]] \
        || deny "Meta guard: activation is allowed only for the Pump & Punish campaign objects (got: $id)."
    done
    ;;
esac
exit 0
