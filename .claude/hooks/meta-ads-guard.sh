#!/usr/bin/env bash
# Guard for the managed Meta Ads campaigns: Meta Ads write tools are pre-approved
# in .claude/settings.json so the scheduled checks can run unattended. This hook
# narrows that approval to the managed objects and to budget/status changes only.
set -euo pipefail

# Pump & Punish: CBO, so only the campaign budget moves.
CAMPAIGN_ID="120249758271140590"
MIN_BUDGET=100000     # 1,000 EGP in piasters
MAX_BUDGET=10000000   # 100,000 EGP in piasters

# Creative tests (Bundle 30, Best-selling bundle): ABO, so each ad set budget
# moves, within a tighter cap.
TEST_CAMPAIGN_IDS="120249762316840590 120249762646600590"
TEST_ADSET_IDS="120249762318740590 120249762319330590 120249762320090590 120249762649160590 120249762650720590 120249762651310590 120249762652320590 120249762652900590"
TEST_MIN_BUDGET=40000    # 400 EGP
TEST_MAX_BUDGET=200000   # 2,000 EGP
# Test campaign objects that may only be paused (the ads).
TEST_PAUSABLE_IDS="120249762322260590 120249762322490590 120249762322660590 120249762653560590 120249762654030590 120249762654330590 120249762655430590 120249762657330590"

# Objects of all managed campaigns: the only ones that may be activated.
ACTIVATABLE_IDS="120249758271140590 120249758274010590 120249758275290590 120249758275550590 120249758275710590 120249758276000590 $TEST_CAMPAIGN_IDS $TEST_ADSET_IDS $TEST_PAUSABLE_IDS"

input="$(cat)"
tool="$(jq -r '.tool_name // ""' <<<"$input")"

deny() {
  jq -n --arg reason "$1" '{hookSpecificOutput: {hookEventName: "PreToolUse", permissionDecision: "deny", permissionDecisionReason: $reason}}'
  exit 0
}

in_list() { [[ " $2 " == *" $1 "* ]]; }

case "$tool" in
  mcp__Meta_Ads__ads_update_entity)
    entity_id="$(jq -r '.tool_input.entity_id // ""' <<<"$input")"
    entity_type="$(jq -r '.tool_input.entity_type // ""' <<<"$input")"
    if [[ "$entity_id" == "$CAMPAIGN_ID" && "$entity_type" == "campaign" ]]; then
      min=$MIN_BUDGET; max=$MAX_BUDGET; budget_allowed=1
    elif [[ "$entity_type" == "ad_set" ]] && in_list "$entity_id" "$TEST_ADSET_IDS"; then
      min=$TEST_MIN_BUDGET; max=$TEST_MAX_BUDGET; budget_allowed=1
    elif [[ "$entity_type" == "ad" ]] && in_list "$entity_id" "$TEST_PAUSABLE_IDS"; then
      budget_allowed=0
    else
      deny "Meta guard: updates are allowed only on the managed Pump & Punish and creative-test objects."
    fi

    fields="$(jq -c '.tool_input.fields | if type == "string" then fromjson else . end' <<<"$input" 2>/dev/null)" \
      || deny "Meta guard: fields is not valid JSON."
    allowed='["daily_budget", "status"]'
    # Test ad sets may also have their start time moved.
    [[ "$entity_type" == "ad_set" ]] && in_list "$entity_id" "$TEST_ADSET_IDS" && allowed='["daily_budget", "status", "start_time"]'
    extra="$(jq -r --argjson allowed "$allowed" 'keys - $allowed | join(",")' <<<"$fields")"
    [[ -z "$extra" ]] || deny "Meta guard: only $allowed may change (got: $extra)."

    status="$(jq -r '.status // ""' <<<"$fields")"
    [[ -z "$status" || "$status" == "PAUSED" ]] \
      || deny "Meta guard: status may only be set to PAUSED (got: $status)."

    if jq -e 'has("daily_budget")' <<<"$fields" >/dev/null; then
      (( budget_allowed )) || deny "Meta guard: ads have no budget to change."
      budget="$(jq -r '.daily_budget' <<<"$fields")"
      [[ "$budget" =~ ^[0-9]+$ ]] || deny "Meta guard: daily_budget must be an integer in piasters."
      (( budget >= min && budget <= max )) \
        || deny "Meta guard: daily_budget $budget is outside $min..$max for $entity_id."
    fi
    ;;
  mcp__Meta_Ads__ads_activate_entity)
    ids="$(jq -r '[.tool_input.entity_id] + (.tool_input.object_ids // []) | map(select(. != null)) | .[]' <<<"$input")"
    for id in $ids; do
      in_list "$id" "$ACTIVATABLE_IDS" \
        || deny "Meta guard: activation is allowed only for the managed campaign objects (got: $id)."
    done
    ;;
esac
exit 0
