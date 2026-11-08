import { BridgeOfflineError } from "../bridge/client";
import type { Bridge, BridgeResponse } from "../bridge/types";
import { planRefresh } from "../citation/refresh-plan";
import { classify, scanDocumentXml, type ScannedField } from "../citation/ooxml";
import { CapabilityError } from "./document";

export interface RefreshOutcome {
  /** Fields whose display text was rewritten from a fresh render. */
  refreshed: number;
  /** Foreign/corrupted fields left untouched. */
  skipped: number;
  /** Fields whose render or update failed; their existing text is preserved. */
  failed: number;
}

/**
 * The ONLY Word-runtime seam refresh goes through. `listFields` returns every
 * field's instruction code + current display text in stable document order;
 * `updateResultText` rewrites the display text of the field at that index.
 * Implemented over `body.fields` (Word desktop) and over a read-only OOXML
 * scan (other hosts) — and faked in tests, which is why this is an interface.
 */
export interface WordFieldAccess {
  listFields(): Promise<{ code: string; resultText: string }[]>;
  updateResultText(index: number, text: string): Promise<void>;
}

/**
 * Rescan every field and re-render each notees citation in place, using the
 * styleId + locator stored in the field's own embed (a document-level style
 * switch never rewrites fields implicitly).
 *
 * Never throws and never corrupts: a field whose render fails keeps its text;
 * a hand-corrupted embed is skipped, not deleted; a thrown BridgeOfflineError
 * aborts the remaining fields and the counts report completed work only
 * (unattempted fields appear in none of the three counters).
 */
export async function refreshCitations(
  bridge: Bridge,
  fieldsApi: WordFieldAccess,
): Promise<RefreshOutcome> {
  let raw: { code: string; resultText: string }[];
  try {
    raw = await fieldsApi.listFields();
  } catch {
    // The document could not be scanned at all — report zero progress rather
    // than throwing; the pane shows the scan failure from the zero outcome.
    return { refreshed: 0, skipped: 0, failed: 0 };
  }
  // Same classification the OOXML scanner applies, so a code that scan would
  // call foreign is skipped here too. Codes are trimmed as the scanner trims.
  const fields: ScannedField[] = raw.map((f) => {
    const code = f.code.trim();
    return { code, displayText: f.resultText, ...classify(code) };
  });
  const plan = planRefresh(fields);
  const outcome: RefreshOutcome = {
    refreshed: 0,
    skipped: raw.length - plan.length,
    failed: 0,
  };
  for (const entry of plan) {
    let res: BridgeResponse<{ citation: string }>;
    try {
      res = await bridge.renderCitation(entry.embed.nodeIds, entry.embed.styleId, entry.embed.locator);
    } catch (e) {
      if (e instanceof BridgeOfflineError) break;
      // Shape/host render error: per-field failure, never aborts the rest.
      outcome.failed += 1;
      continue;
    }
    if (!res.ok) {
      outcome.failed += 1;
      continue;
    }
    try {
      await fieldsApi.updateResultText(entry.fieldIndex, res.citation);
      outcome.refreshed += 1;
    } catch {
      outcome.failed += 1;
    }
  }
  return outcome;
}

/** Pane copy for a refresh run: "2 refreshed · 1 skipped (foreign/corrupted) · 0 failed". */
export function describeRefreshOutcome(outcome: RefreshOutcome): string {
  return `${String(outcome.refreshed)} refreshed · ${String(outcome.skipped)} skipped (foreign/corrupted) · ${String(outcome.failed)} failed`;
}

/**
 * Desktop field access over `body.fields` (WordApi 1.5): codes and result
 * text load in one sync; an update replaces the field's result-range text,
 * which keeps the field itself intact.
 */
export function bodyFieldAccess(): WordFieldAccess {
  return {
    listFields: () =>
      Word.run(async (context) => {
        const fields = context.document.body.fields;
        fields.load("code, result");
        await context.sync();
        return fields.items.map((f) => ({ code: f.code, resultText: f.result.text }));
      }),
    updateResultText: (index, text) =>
      Word.run(async (context) => {
        const fields = context.document.body.fields;
        fields.load("items");
        await context.sync();
        const field = fields.items[index];
        if (field === undefined) {
          throw new CapabilityError("field-update", "Field disappeared before its text could be updated.");
        }
        field.result.insertText(text, Word.InsertLocation.replace);
        await context.sync();
      }),
  };
}

/**
 * Fallback for hosts without WordApi 1.5 field access (Word for the web):
 * fields are listed via an OOXML body dump + the pure scanner, but their
 * display text cannot be rewritten — `updateResultText` rejects with a
 * CapabilityError. Refresh on such a host reports the degradation up front
 * and never touches the document.
 */
export function readOnlyFieldAccess(): WordFieldAccess {
  return {
    listFields: async () => {
      const xml = await Word.run(async (context) => {
        const ooxml = context.document.body.getOoxml();
        await context.sync();
        return ooxml.value;
      });
      return scanDocumentXml(xml).map((f) => ({ code: f.code, resultText: f.displayText }));
    },
    updateResultText: () =>
      Promise.reject(
        new CapabilityError(
          "field-update",
          "Citation fields are read-only on this Word host — open the document in the Word desktop app to refresh.",
        ),
      ),
  };
}

/**
 * Pick the field access for the current host: writable `body.fields` when the
 * host supports WordApi 1.5, else the read-only OOXML scan. Any capability
 * probe failure degrades to read-only — never to a broken document.
 */
export function wordFieldAccess(): { access: WordFieldAccess; writable: boolean } {
  try {
    if (Office.context.requirements.isSetSupported("WordApi", "1.5")) {
      return { access: bodyFieldAccess(), writable: true };
    }
  } catch {
    // Fall through to the read-only access.
  }
  return { access: readOnlyFieldAccess(), writable: false };
}
