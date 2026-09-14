/* Client-side extraction core: PDF annotations + DOCX tracked changes,
   ported 1:1 from the Python extractor/formatter. No server, ever. */

/* ------------------------------------------------------------------ utils */
const W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const W14_NS = "http://schemas.microsoft.com/office/word/2010/wordml";
const W15_NS = "http://schemas.microsoft.com/office/word/2012/wordml";

function xmlAttr(el, name) {
  return el.getAttribute(name) ?? el.getAttributeNS(W_NS, name) ?? el.getAttributeNS(W14_NS, name);
}

function parsePdfDate(raw) {
  if (!raw) return "";
  let s = raw.startsWith("D:") ? raw.slice(2) : raw;
  try {
    const clean = s.replace(/'/g, "");
    const dt = new Date(
      Date.UTC(+clean.slice(0, 4), +clean.slice(4, 6) - 1, +clean.slice(6, 8),
               +clean.slice(8, 10), +clean.slice(10, 12), +clean.slice(12, 14))
    );
    if (isNaN(dt.getTime())) return raw;
    const p = (n) => String(n).padStart(2, "0");
    let out = `${dt.getUTCFullYear()}-${p(dt.getUTCMonth() + 1)}-${p(dt.getUTCDate())} ` +
              `${p(dt.getUTCHours())}:${p(dt.getUTCMinutes())}:${p(dt.getUTCSeconds())}`;
    const tz = clean.slice(14);
    if (tz && (tz.startsWith("+") || tz.startsWith("-"))) out += ` (UTC${tz})`;
    return out;
  } catch { return raw; }
}

function parseDocxDate(raw) {
  if (!raw) return "";
  const t = Date.parse(raw);
  if (isNaN(t)) return raw;
  const d = new Date(t), p = (n) => String(n).padStart(2, "0");
  return `${d.getFullYear()}-${p(d.getMonth() + 1)}-${p(d.getDate())} ` +
         `${p(d.getHours())}:${p(d.getMinutes())}:${p(d.getSeconds())}`;
}

const DEFAULT_INSTRUCTION =
  "> For each item, locate the comment in the source document " +
  "using the page (when given) and the quoted text.";

/* ------------------------------------------------------------- formatter */
function toMarkdown(items, filename, includeInstructions) {
  const md = [`# Review Comments: \`${filename}\``, `**Total Items:** ${items.length}  `, ""];
  if (includeInstructions) { md.push(DEFAULT_INSTRUCTION, ""); }
  if (!items.length) { md.push("*No comments or annotations found in this document.*"); return md.join("\n"); }

  items.forEach((c, i) => {
    const pageLabel = c.page ? `Page ${c.page} — ` : "";
    md.push(`## [${i + 1}] ${pageLabel}${c.type}`);
    const meta = [];
    if (c.author && c.author !== "Unknown Reviewer") meta.push(`**Author:** ${c.author}`);
    if (c.date) meta.push(`**Date:** ${c.date}`);
    if (meta.length) md.push(meta.join(" | ") + "  ");
    if (c.quoted_text) {
      md.push(`> **Quoted Text:** *"${String(c.quoted_text).replace(/\n/g, " ")}"*`, "");
    }
    if (c.comment) {
      md.push("**Comment:**");
      for (const line of c.comment.split("\n")) md.push(`> ${line}`);
    } else {
      md.push("**Comment:** *(Annotation mark without text)*");
    }
    md.push("", "---", "");
  });
  return md.join("\n");
}

/* ------------------------------------------------- DOCX (OpenXML) walker */
function _paragraphText(p) {
  let out = "";
  p.querySelectorAll("*").forEach((n) => {
    const ln = n.localName;
    if ((ln === "t" || ln === "delText") && n.textContent) out += n.textContent;
    else if (ln === "br") out += "\n";
    else if (ln === "tab") out += " ";
  });
  return out;
}

function _parseThreads(doc, comments, paraToCid) {
  if (!doc) return {};
  // doc: Document for commentsExtended.xml; returns {root_cid: [reply_cids]}
  const replyParent = {};
  doc.querySelectorAll("*").forEach((ex) => {
    if (ex.localName !== "commentEx") return;
    const pid = ex.getAttributeNS(W15_NS, "paraId") || ex.getAttribute("w15:paraId");
    const ppid = ex.getAttributeNS(W15_NS, "paraIdParent") || ex.getAttribute("w15:paraIdParent");
    if (!ppid) return;
    const cid = paraToCid[pid], pcid = paraToCid[ppid];
    if (cid && pcid && cid !== pcid) replyParent[cid] = pcid;
  });
  void comments;
  const threads = {};
  for (const cid of Object.keys(comments)) {
    let root = cid; const seen = new Set();
    while (replyParent[root] && !seen.has(root)) { seen.add(root); root = replyParent[root]; }
    if (root !== cid) (threads[root] ??= []).push(cid);
  }
  return threads;
}

function extractDocx(bytes, name) {
  const files = fflate.unzipSync(new Uint8Array(bytes));
  const parser = new DOMParser();
  const getDoc = (part) => {
    const data = files[part];
    return data ? parser.parseFromString(new TextDecoder().decode(data), "text/xml") : null;
  };

  const comments = {};      // cid -> record
  const paraToCid = {};
  const cDoc = getDoc("word/comments.xml");
  if (cDoc) {
    cDoc.querySelectorAll("*").forEach((c) => {
      if (c.localName !== "comment") return;
      const cid = c.getAttributeNS(W_NS, "id") ?? c.getAttribute("w:id");
      if (cid == null) return;
      const paras = [];
      c.querySelectorAll("*").forEach((n) => {
        if (n.localName === "p") paras.push(_paragraphText(n));
      });
      const text = paras.filter(Boolean).join("\n").trim();
      comments[cid] = {
        id: cid,
        author: (c.getAttributeNS(W_NS, "author") || c.getAttribute("w:author") || "").trim() || "Unknown Reviewer",
        date: parseDocxDate(c.getAttributeNS(W_NS, "date") || c.getAttribute("w:date") || ""),
        comment: text,
        type: "Comment",
        quoted_text: "",
        page: 1,
      };
      c.querySelectorAll("*").forEach((p) => {
        if (p.localName !== "p") return;
        const pid = p.getAttributeNS(W14_NS, "paraId") || p.getAttribute("w14:paraId");
        if (pid) paraToCid[pid] = cid;
      });
    });
  }

  const docEl = getDoc("word/document.xml");
  if (!docEl) return [];
  const body = [...docEl.querySelectorAll("*")].find((n) => n.localName === "body");
  if (!body) return [];

  const threadMembers = {};
  {
    // resolve reply chains to roots
    const rootOf = (cid) => {
      const parentOf = {};
      // replyParent map is rebuilt inside _parseThreads; replicate via call:
      return cid;
    };
    void rootOf;
  }
  const threads = _parseThreads(getDoc("word/commentsExtended.xml"), comments, paraToCid);
  const allReplies = new Set(Object.values(threads).flat());

  // ---- event walk ---------------------------------------------------------
  const CHANGE_TAGS = { ins: "ins_text", del: "del_text", moveTo: "move_to", moveFrom: "move_from" };
  const CHANGE_KINDS = new Set(["del_text", "ins_text", "move_from", "move_to"]);

  const events = [];
  (function walk(node, changeKind, changeAuthor, changeDate) {
    const tag = node.localName;
    if (tag === "lastRenderedPageBreak" || (tag === "br" && xmlAttr(node, "type") === "page")) {
      events.push(["page_break"]);
    } else if (tag === "commentRangeStart") {
      events.push(["comment_start", xmlAttr(node, "id")]);
    } else if (tag === "commentRangeEnd") {
      events.push(["comment_end", xmlAttr(node, "id")]);
    } else if (CHANGE_TAGS[tag]) {
      const kind = CHANGE_TAGS[tag];
      const author = xmlAttr(node, "author") || "Unknown Reviewer";
      const date = parseDocxDate(xmlAttr(node, "date") || "");
      for (const child of node.children) walk(child, kind, author, date);
      return;
    } else if (tag === "t" || tag === "delText") {
      if (node.textContent) {
        if (changeKind) events.push([changeKind, changeAuthor, changeDate, node.textContent]);
        else if (tag === "t") events.push(["text", node.textContent]);
        else events.push(["del_text", "Unknown Reviewer", changeDate || "", node.textContent]);
      }
      return;
    } else if (tag === "p") {
      for (const child of node.children) walk(child, changeKind, changeAuthor, changeDate);
      events.push(["para_break"]);
      return;
    }
    for (const child of node.children) walk(child, changeKind, changeAuthor, changeDate);
  })(body);

  // ---- event consumption (mirrors the Python extractor) -------------------
  let current_page = 1, sawPageBreak = false;
  const active = new Set();
  const quotes = {}, pages = {}, emitted = new Set();
  const items = [];
  let recentPlain = [];

  const rootText = (cid) => {
    const members = [cid, ...(threads[cid] || [])];
    const quoted = quotes[cid]?.join("").trim() ||
      (threads[cid] || []).map((m) => (quotes[m] || []).join("").trim()).find(Boolean) || "";
    let text = comments[cid].comment;
    for (const m of (threads[cid] || [])) {
      const rd = comments[m];
      if (!rd.comment) continue;
      const stamp = rd.date ? ` (${rd.date})` : "";
      text += `\n↳ ${rd.author}${stamp}: ${rd.comment.replace(/\n/g, " ")}`;
    }
    let page = pages[cid];
    if (page == null) page = (threads[cid] || []).map((m) => pages[m]).find((p) => p != null) ?? 1;
    return { page, quoted, text, author: comments[cid].author, date: comments[cid].date };
  };

  const closeComment = (cid) => {
    if (!(cid in comments) || emitted.has(cid)) return;
    if (threads[cid] || !allReplies.has(cid)) {
      const members = [cid, ...(threads[cid] || [])];
      const root = rootText(cid);
      items.push({
        page: root.page, type: "Comment", comment: root.text,
        quoted_text: root.quoted, author: root.author, date: root.date, id: cid,
      });
      for (const m of members) emitted.add(m);
    }
  };

  const contextSnippet = (pre, post) => {
    const parts = [];
    if (pre) parts.push(`...${pre}`);
    parts.push("[...]");
    if (post) parts.push(`${post}...`);
    return parts.join(" ");
  };
  const emitDeletion = (block, page) => {
    const del = block[3].trim();
    if (!del) return;
    items.push({ page, type: "Deletion", comment: `Delete: "${del}"`,
      quoted_text: del, del_text: del, author: block[1], date: block[2] });
  };
  const emitInsertion = (block, page, snippet) => {
    const ins = block[3].trim();
    if (!ins) return;
    items.push({ page, type: "Insertion", comment: `Insert: "${ins}"`,
      quoted_text: snippet, ins_text: ins, author: block[1], date: block[2] });
  };
  const emitMove = (block, page, snippet, isSource) => {
    const t = block[3].trim();
    if (!t) return;
    const label = isSource ? "Move (source)" : "Move (target)";
    const action = isSource ? "moved away from this position" : "moved into this position";
    items.push({ page, type: label, comment: `Move: "${t}" ${action}`,
      quoted_text: snippet, del_text: isSource ? t : "", ins_text: isSource ? "" : t,
      author: block[1], date: block[2] });
  };

  let i = 0;
  while (i < events.length) {
    const ev = events[i], kind = ev[0];
    if (kind === "page_break") { sawPageBreak = true; current_page += 1; i++; continue; }
    if (kind === "para_break") { recentPlain = []; i++; continue; }
    if (kind === "comment_start") {
      const cid = ev[1];
      if (cid in comments) { active.add(cid); if (!(cid in pages)) pages[cid] = current_page; }
      i++; continue;
    }
    if (kind === "comment_end") {
      active.delete(ev[1]);
      closeComment(ev[1]);
      i++; continue;
    }
    if (kind === "text") {
      for (const cid of active) (quotes[cid] ??= []).push(ev[1]);
      recentPlain.push(ev[1]);
      i++; continue;
    }
    if (CHANGE_KINDS.has(kind)) {
      const cluster = [];
      let j = i;
      while (j < events.length && (CHANGE_KINDS.has(events[j][0]) || events[j][0] === "comment_start" || events[j][0] === "comment_end")) {
        if (events[j][0] === "comment_start") {
          const cid = events[j][1];
          if (cid in comments) { active.add(cid); if (!(cid in pages)) pages[cid] = current_page; }
        } else if (events[j][0] === "comment_end") {
          active.delete(events[j][1]);
          closeComment(events[j][1]);
        } else {
          cluster.push(events[j]);
          if (active.size) for (const cid of active) (quotes[cid] ??= []).push(events[j][3]);
        }
        j++;
      }
      let future = [];
      for (let fj = j; fj < events.length && events[fj][0] !== "para_break" && future.join("").length < 50; fj++) {
        if (events[fj][0] === "text") future.push(events[fj][1]);
      }
      const pre = recentPlain.join("").slice(-35).trim();
      const post = future.join("").slice(0, 35).trim();
      const snippet = ((pre ? `...${pre} ` : "") + "[...] " + (post ? `${post}...` : "")).trim();

      const blocks = [];
      for (const b of cluster) {
        const [k, author, date, txt] = b;
        if (blocks.length && blocks[blocks.length - 1].k === k && blocks[blocks.length - 1].author === author) {
          blocks[blocks.length - 1].text += txt;
        } else blocks.push({ k, author, date, text: txt });
      }

      let k = 0;
      while (k < blocks.length) {
        const curr = blocks[k], nxt = blocks[k + 1] ?? null;
        if (nxt) {
          const pairKey = new Set([curr.k, nxt.k]);
          if (pairKey.has("del_text") && pairKey.has("ins_text")) {
            const delB = curr.k === "del_text" ? curr : nxt;
            const insB = curr.k === "ins_text" ? curr : nxt;
            const del = delB.text.trim(), ins = insB.text.trim();
            if (!ins) emitDeletion(delB, current_page);
            else if (!del) emitInsertion(insB, current_page, snippet);
            else items.push({
              page: current_page, type: "Replacement",
              comment: `Replace "${del}" with: "${ins}"`, quoted_text: del,
              del_text: del, ins_text: ins,
              author: insB.author || delB.author, date: insB.date || delB.date,
            });
            k += 2; continue;
          }
          if (pairKey.has("move_from") && pairKey.has("move_to")) {
            const src = curr.k === "move_from" ? curr : nxt;
            const tgt = curr.k === "move_to" ? curr : nxt;
            const t = src.text.trim() || tgt.text.trim();
            if (t) items.push({
              page: current_page, type: "Moved",
              comment: `Move: "${t}" was moved (removed here, re-inserted nearby)`,
              quoted_text: snippet, del_text: src.text.trim(), ins_text: tgt.text.trim(),
              author: tgt.author || src.author, date: tgt.date || src.date,
            });
            k += 2; continue;
          }
        }
        if (curr.k === "del_text") emitDeletion(curr, current_page);
        else if (curr.k === "ins_text") emitInsertion(curr, current_page, snippet);
        else if (curr.k === "move_from") emitMove(curr, current_page, snippet, true);
        else if (curr.k === "move_to") emitMove(curr, current_page, snippet, false);
        k += 1;
      }
      i = j;
      continue;
    }
    i++;
  }

  for (const cid of Object.keys(comments)) closeComment(cid);

  if (!sawPageBreak) for (const it of items) it.page = null;
  return items;
}

function extractDocxParagraphs(bytes) {
  const files = fflate.unzipSync(new Uint8Array(bytes));
  const data = files["word/document.xml"];
  if (!data) return [];
  const doc = new DOMParser().parseFromString(new TextDecoder().decode(data), "text/xml");
  const paras = [];
  doc.querySelectorAll("*").forEach((n) => {
    if (n.localName === "p") {
      const t = _paragraphText(n).trim();
      if (t) paras.push(t);
    }
  });
  return paras;
}

/* ------------------------------------------------- PDF (pdf.js) extraction */
async function extractPdf(bytes, name) {
  void name;
  const pdfjsLib = window.pdfjsLib;
  pdfjsLib.GlobalWorkerOptions.workerSrc = "./vendor/pdf.worker.min.js";
  const pdf = await pdfjsLib.getDocument({ data: new Uint8Array(bytes).slice(), standardFontDataUrl: "./vendor/standard_fonts/", useSystemFonts: true }).promise;
  const items = [];
  for (let pno = 1; pno <= pdf.numPages; pno++) {
    const page = await pdf.getPage(pno);
    const annotations = await page.getAnnotations();
    const textContent = await page.getTextContent();
    const pageHeight = page.getViewport({ scale: 1 }).height;

    const textUnder = (r) => {
      // r and text transforms are both in PDF user space (origin bottom-left)
      const parts = [];
      for (const item of textContent.items) {
        if (!item.str) continue;
        const t = item.transform;
        const x = t[4], baseline = t[5], w = item.width, h = item.height || 10;
        const itemBottom = baseline - 0.25 * h;
        const itemTop = baseline + 0.85 * h;
        const overlap = Math.min(r[2], x + w) - Math.max(r[0], x);
        const overY = Math.min(r[3], itemTop) - Math.max(r[1], itemBottom);
        if (overlap > 0.5 && overY > 0) parts.push([x, item.str]);
      }
      parts.sort((a, b) => a[0] - b[0]);
      return parts.map((p) => p[1]).join(" ").replace(/\s+/g, " ").trim();
    };

    for (const a of annotations) {
      const s = a.subtype;
      let type = null;
      if (s === "Highlight") type = "Highlight";
      else if (s === "Underline") type = "Underline";
      else if (s === "Squiggly") type = "Squiggly";
      else if (s === "StrikeOut") type = "StrikeOut";
      else if (s === "Text") type = "Sticky Note";
      else if (s === "FreeText") type = "FreeText";
      else if (s === "Caret") type = "Caret";
      else if (s === "Stamp") type = "Stamp";
      else if (s === "Ink") type = "Ink";
      else if (s === "Line" || s === "Square" || s === "Circle" || s === "Polygon" || s === "PolyLine" || s === "Redaction") type = s;
      if (!type) continue;

      const comment = (a.contentsObj?.str ?? a.contents ?? "").trim();
      const quoted = ["Highlight", "Underline", "Squiggly", "StrikeOut"].includes(s) ? textUnder(a.rect) : "";
      const subject = (a.subject || "").trim();
      if (!comment && !quoted) {
        if (subject && (s === "FreeText")) items.push({
          page: pno, type, comment: subject, quoted_text: "",
          author: (a.titleObj?.str ?? a.title ?? "") || "Unknown Reviewer",
          date: parsePdfDate(a.creationDate || ""),
        });
        continue;
      }
      items.push({
        page: pno, type,
        comment, quoted_text: quoted,
        author: (a.titleObj?.str ?? a.title ?? "") || "Unknown Reviewer",
        date: parsePdfDate(a.creationDate || ""),
        rect: a.rect,
      });
    }
  }
  return items.sort((a, b) => (a.page - b.page) || (a.rect ? a.rect[1] : 0) - (b.rect ? b.rect[1] : 0));
}

/* --------------------------------------------- PDF page rendering (pdf.js) */
async function renderPdfPage(bytes, pageNo, scale) {
  const pdfjsLib = window.pdfjsLib;
  pdfjsLib.GlobalWorkerOptions.workerSrc = "./vendor/pdf.worker.min.js";
  const pdf = await pdfjsLib.getDocument({ data: new Uint8Array(bytes).slice(), standardFontDataUrl: "./vendor/standard_fonts/", useSystemFonts: true }).promise;
  const page = await pdf.getPage(pageNo);
  const viewport = page.getViewport({ scale });
  const canvas = document.createElement("canvas");
  canvas.width = viewport.width; canvas.height = viewport.height;
  await page.render({ canvasContext: canvas.getContext("2d"), viewport }).promise;
  const png = canvas.toDataURL("image/png");
  // annotations for this page (viewport-space boxes)
  const boxes = (await page.getAnnotations())
    .filter((a) => !["Link", "Popup"].includes(a.subtype))
    .map((a) => ({ rect: a.rect, subtype: a.subtype }));
  const textHits = await page.getTextContent();
  return {
    page: pageNo, total_pages: pdf.numPages, scale,
    width: viewport.width / scale, height: viewport.height / scale,
    image: png, boxes, textItems: textHits.items.length,
  };
}

async function searchPdfPage(bytes, pageNo, needle) {
  const pdfjsLib = window.pdfjsLib;
  pdfjsLib.GlobalWorkerOptions.workerSrc = "./vendor/pdf.worker.min.js";
  const pdf = await pdfjsLib.getDocument({ data: new Uint8Array(bytes).slice(), standardFontDataUrl: "./vendor/standard_fonts/", useSystemFonts: true }).promise;
  const page = await pdf.getPage(pageNo);
  return page.getAnnotations(); // caller picks rects
}

async function findTextRects(pdfBytes, needle) {
  const pdfjsLib = window.pdfjsLib;
  pdfjsLib.GlobalWorkerOptions.workerSrc = "./vendor/pdf.worker.min.js";
  const pdf = await pdfjsLib.getDocument({ data: new Uint8Array(pdfBytes).slice(), standardFontDataUrl: "./vendor/standard_fonts/", useSystemFonts: true }).promise;
  const needleL = needle.toLowerCase();
  const hits = [];
  for (let pno = 1; pno <= pdf.numPages && hits.length < 50; pno++) {
    const page = await pdf.getPage(pno);
    const pageHeight = page.getViewport({ scale: 1 }).height;
    const tc = await page.getTextContent();
    for (const item of tc.items) {
      if (!item.str) continue;
      const idx = item.str.toLowerCase().indexOf(needleL);
      if (idx === -1) continue;
      const t = item.transform;
      const w = (needleL.length / Math.max(item.str.length, 1)) * item.width;
      const h = item.height || 10;
      const x = t[4] + (idx / Math.max(item.str.length, 1)) * item.width;
      const yTop = pageHeight - (t[5] + h);
      hits.push({ page: pno, rect: [x, yTop, x + w, yTop + h] });
      break; // first hit per page
    }
  }
  return hits;
}
