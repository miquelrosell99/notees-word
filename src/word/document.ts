import type { CitationEmbed } from "../citation/embed";
import { citationFieldXml } from "../citation/ooxml";

/** A Word-host capability is missing or the host rejected the call. */
export class CapabilityError extends Error {
  readonly capability: string;

  constructor(capability: string, message?: string, options?: { cause?: unknown }) {
    super(message ?? `Word capability unavailable: ${capability}`, options);
    this.name = "CapabilityError";
    this.capability = capability;
  }
}

/**
 * The thinnest possible seam over the Word runtime. Everything that touches
 * `Word`/`Office` globals lives behind this interface so the composition
 * logic is unit-testable in node with a fake runtime.
 */
export interface WordRuntime {
  isSetSupported(set: string, version: string): boolean;
  insertOoxmlAtSelection(xml: string): Promise<void>;
}

/** The real adapter, bound to the Office.js globals. Only used at the edge. */
export function wordRuntime(): WordRuntime {
  return {
    isSetSupported: (set, version) => {
      try {
        return Office.context.requirements.isSetSupported(set, version);
      } catch {
        return false;
      }
    },
    insertOoxmlAtSelection: (xml) =>
      Word.run(async (context) => {
        const range = context.document.getSelection().getRange();
        range.insertOoxml(xml, Word.InsertLocation.replace);
        await context.sync();
      }),
  };
}

/**
 * Full paragraph snippet for `insertOoxml`: the fldSimple field run must be a
 * direct child of `w:p` (fldSimple inside `w:r` is not valid WordprocessingML),
 * and the `w` namespace must be declared for a bare paragraph fragment.
 */
export function citationParagraphXml(embed: CitationEmbed, displayText: string): string {
  return `<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">${citationFieldXml(embed, displayText)}</w:p>`;
}

/**
 * Insert the citation field at the current selection, replacing it.
 *
 * `embed.citationId` must be a fresh uuid per call — the caller owns id
 * minting so undo/reinsert semantics stay outside this adapter. Throws
 * `CapabilityError("insert-ooxml")` when the host lacks OOXML insertion
 * (Word for the web) or rejects the insert — the pane surfaces that as a
 * user-visible message.
 */
export async function insertCitationFieldWith(
  runtime: WordRuntime,
  embed: CitationEmbed,
  displayText: string,
): Promise<void> {
  if (embed.nodeIds.length === 0) {
    throw new RangeError("insertCitationField requires at least one node id");
  }
  if (!runtime.isSetSupported("WordApi", "1.5")) {
    throw new CapabilityError(
      "insert-ooxml",
      "This Word host cannot insert citations (OOXML insertion requires WordApi 1.5 — use the Word desktop app).",
    );
  }
  try {
    await runtime.insertOoxmlAtSelection(citationParagraphXml(embed, displayText));
  } catch (cause) {
    throw new CapabilityError("insert-ooxml", "The Word host rejected the citation insert.", {
      cause,
    });
  }
}

/** Selection-bound entry point used by the pane; delegates to the injectable core. */
export async function insertCitationField(embed: CitationEmbed, displayText: string): Promise<void> {
  await insertCitationFieldWith(wordRuntime(), embed, displayText);
}
