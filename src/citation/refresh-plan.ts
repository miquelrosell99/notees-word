import type { CitationEmbed } from "./embed";
import type { ScannedField } from "./ooxml";

export interface RefreshPlanEntry {
  /** Position of the field in the scan order — the index WordFieldAccess updates by. */
  fieldIndex: number;
  /** The parsed embed: carries the styleId + locator the field was inserted with. */
  embed: CitationEmbed;
}

/**
 * Pure refresh plan: keep only citation fields whose ADDIN code parses to a
 * well-formed embed. Foreign fields (Zotero, …) and hand-corrupted codes
 * (prefix present, JSON garbage) classify as foreign upstream and are skipped
 * here — never deleted, never re-rendered (skip-not-corrupt contract).
 *
 * v1 renders once per planned field; batching renders by distinct styleId is
 * a recorded optimization, not this task.
 */
export function planRefresh(fields: ScannedField[]): RefreshPlanEntry[] {
  const plan: RefreshPlanEntry[] = [];
  fields.forEach((field, fieldIndex) => {
    if (field.kind !== "citation" || field.embed === null) return;
    if (!("nodeIds" in field.embed)) return;
    plan.push({ fieldIndex, embed: field.embed });
  });
  return plan;
}
