/* global document, Office, Word */

type StatusKind = "info" | "success" | "error";

Office.onReady((info) => {
  if (info.host !== Office.HostType.Word) return;

  // Hide sideload message, show app UI
  const sideload = document.getElementById("sideload-msg");
  const appBody = document.getElementById("app-body");
  if (sideload) sideload.style.display = "none";
  if (appBody) appBody.style.display = "block";

  document.getElementById("insertTemplate")?.addEventListener("click", insertTemplate);
  document.getElementById("bulletsFromSelection")?.addEventListener("click", bulletsFromSelection);

  setStatus("Ready.", "success");
});

function getTitle(): string {
  const el = document.getElementById("titleInput") as HTMLInputElement | null;
  const title = (el?.value ?? "").trim();
  return title || "McCarren Office Challenge";
}

function setBusy(isBusy: boolean) {
  const insertBtn = document.getElementById("insertTemplate") as HTMLButtonElement | null;
  const bulletsBtn = document.getElementById("bulletsFromSelection") as HTMLButtonElement | null;
  const titleInput = document.getElementById("titleInput") as HTMLInputElement | null;

  if (insertBtn) insertBtn.disabled = isBusy;
  if (bulletsBtn) bulletsBtn.disabled = isBusy;
  if (titleInput) titleInput.disabled = isBusy;
}

function setStatus(msg: string, kind: StatusKind = "info") {
  const el = document.getElementById("status");
  if (!el) return;

  el.textContent = msg;
  el.classList.remove("status--info", "status--success", "status--error");
  el.classList.add(`status--${kind}`);
}

/**
 * Convert Word selection text into a clean array of bullet items.
 * Handles:
 * - real newlines (\n), carriage returns (\r)
 * - literal "\r" and "\n" sequences that sometimes appear
 * - \u2028 / \u2029 separators
 * - backslash-separated lists: apple\banana\cherry
 * - comma-separated lists: apple, banana, cherry
 * - strips leading bullets/numbering like: •, -, *, 1), 1.
 */
function parseItems(selectionText: string): string[] {
  if (!selectionText) return [];

  let t = selectionText.trim();
  if (!t) return [];

  // If Word returns literal "\r" (two chars), normalize it
  t = t.replace(/\\r/g, "\n").replace(/\\n/g, "\n");

  // Normalize real line breaks and separators into \n
  t = t
    .replace(/\r\n/g, "\n")
    .replace(/\r/g, "\n")
    .replace(/\u2028/g, "\n")
    .replace(/\u2029/g, "\n");

  // If no newlines, allow "\" or "," as separators
  if (!t.includes("\n") && (t.includes("\\") || t.includes(","))) {
    return t
      .split(/[\\,]/g)
      .map((s) => s.trim())
      .filter(Boolean)
      .map(stripPrefix);
  }

  // Standard: split by lines
  return t
    .split("\n")
    .map((s) => s.trim())
    .filter(Boolean)
    .map(stripPrefix);
}

function stripPrefix(s: string): string {
  // Remove common bullet/number prefixes: •, -, *, 1), 1., (1)
  return s.replace(/^(\u2022|\-|\*|\(\d+\)|\d+[.)])\s*/g, "").trim();
}

async function insertTemplate() {
  const title = getTitle();
  setBusy(true);
  setStatus("Inserting template...", "info");

  try {
    await Word.run(async (context) => {
      const body = context.document.body;

      // Optional spacer if user already has content
      // (harmless even if empty)
      body.insertParagraph("", Word.InsertLocation.end);

      const h1 = body.insertParagraph(title, Word.InsertLocation.end);
      h1.styleBuiltIn = Word.BuiltInStyleName.heading1;

      const meta = body.insertParagraph(
        `Generated: ${new Date().toLocaleString()}`,
        Word.InsertLocation.end
      );
      meta.font.italic = true;

      const h2Summary = body.insertParagraph("Summary", Word.InsertLocation.end);
      h2Summary.styleBuiltIn = Word.BuiltInStyleName.heading2;

      // Real paragraphs (no weird embedded breaks)
      body.insertParagraph("• Purpose: (write 1 sentence)", Word.InsertLocation.end);
      body.insertParagraph("• Outcome: (write 1 sentence)", Word.InsertLocation.end);
      body.insertParagraph("• Next Step: (write 1 sentence)", Word.InsertLocation.end);

      const h2Key = body.insertParagraph("Key Points", Word.InsertLocation.end);
      h2Key.styleBuiltIn = Word.BuiltInStyleName.heading2;

      // Reliable bulleted list
      const first = body.insertParagraph("What was done", Word.InsertLocation.end);
      const list = first.startNewList();
      list.insertParagraph("Why it matters", Word.InsertLocation.end);
      list.insertParagraph("What happens next", Word.InsertLocation.end);

      const h2Notes = body.insertParagraph("Notes", Word.InsertLocation.end);
      h2Notes.styleBuiltIn = Word.BuiltInStyleName.heading2;

      const notes = body.insertParagraph("(Add any details here.)", Word.InsertLocation.end);
      notes.font.color = "#2b579a"; // subtle link-ish blue, easy to spot

      await context.sync();
    });

    setStatus("Template inserted.", "success");
  } catch (err: any) {
    console.error(err);
    setStatus(`Error inserting template: ${err?.message || "Check console."}`, "error");
  } finally {
    setBusy(false);
  }
}

async function bulletsFromSelection() {
  setBusy(true);
  setStatus("Creating bullets...", "info");

  try {
    await Word.run(async (context) => {
      const range = context.document.getSelection();
      range.load("text");
      await context.sync();

      const items = parseItems(range.text);

      if (items.length === 0) {
        setStatus("Select text first (one item per line, or apple\\banana\\cherry).", "error");
        return;
      }

      // Replace the selection with a bulleted list
      range.insertText("", Word.InsertLocation.replace);

      const first = range.insertParagraph(items[0], Word.InsertLocation.after);
      const list = first.startNewList();

      for (let i = 1; i < items.length; i++) {
        list.insertParagraph(items[i], Word.InsertLocation.end);
      }

      await context.sync();
    });

    setStatus("Bullets created.", "success");
  } catch (err: any) {
    console.error(err);
    setStatus(`Error creating bullets: ${err?.message || "Check console."}`, "error");
  } finally {
    setBusy(false);
  }
}
