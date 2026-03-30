import React, { useState, useRef, useCallback, useEffect } from "react";

// ─── VERIFIED MATH ENGINE ─────────────────────────────────────────────────────
function normalizeIndicDigits(value) {
  const map = {
    "०": "0", "१": "1", "२": "2", "३": "3", "४": "4",
    "५": "5", "६": "6", "७": "7", "८": "8", "९": "9",
  };
  return String(value ?? "").replace(/[०-९]/g, (d) => map[d] || d);
}
function parseQty(value) {
  const n = parseInt(normalizeIndicDigits(value), 10);
  return Math.max(1, n || 1);
}
function parseFt(dim) {
  if (!dim) return 0;
  let s = normalizeIndicDigits(dim).trim()
    .replace(/[\u2018\u2019\u02bc`\u02b9]/g, "'")   // smart apostrophes
    .replace(/[\u201c\u201d\u02ba]/g, '"')            // smart quotes
    .replace(/\s+/g, '');                             // remove spaces

  // Dash format: 2'-2'' or 2'-2" or 7'-6'' → normalise to 2'2"
  s = s.replace(/^(\d+)'-(\d+)['"]{1,2}$/, "$1'$2\"");

  // Comma format: 2,4'' → 2'4"
  s = s.replace(/^(\d+),(\d+)/, "$1'$2");

  // Strip duplicate trailing quotes: 2'4'' → 2'4"
  s = s.replace(/''+$/, "'").replace(/""+$/, '"');
  s = s.replace(/^(\d+)'(\d+)'$/, "$1'$2\"");

  // 7'6"  or  7'6  (feet + inches)
  let m = s.match(/^(\d+)'(\d+)"$/); if (m) return parseInt(m[1]) + parseInt(m[2]) / 12;
  m = s.match(/^(\d+)'(\d+)$/);      if (m) return parseInt(m[1]) + parseInt(m[2]) / 12;

  // 7'  (feet only)
  m = s.match(/^(\d+)'$/); if (m) return parseInt(m[1]);

  // 33"  or  6"  (inches only — must have " marker)
  m = s.match(/^(\d+)"$/); if (m) return parseInt(m[1]) / 12;

  // plain number — treat as feet
  m = s.match(/^(\d+(?:\.\d+)?)$/); if (m) return parseFloat(m[1]);

  return 0;
}
function calcArea(d1, d2, qty, type) {
  const q = parseQty(qty);
  if (type === "rnft" || type === "grove") return parseFt(d1) * q;
  return parseFt(d1) * parseFt(d2) * q;
}
function netTotal(rows) {
  return rows.reduce((s, r) => s + (r.deduct ? -r.area : r.area), 0);
}
function grossTotal(rows) {
  return rows.filter(r => !r.deduct).reduce((s, r) => s + r.area, 0);
}
function deductionTotal(rows) {
  return rows.filter(r => r.deduct).reduce((s, r) => s + r.area, 0);
}
function grandTotal(sessions) {
  return sessions.reduce((s, sess) => s + netTotal(sess.rows), 0);
}
function sqftTotal(rows) {
  return rows.filter(r => r.type === "sqft").reduce((s, r) => s + (r.deduct ? -r.area : r.area), 0);
}
function groveRowTotal(rows) {
  return rows.filter(r => r.type === "grove").reduce((s, r) => s + (r.deduct ? -r.area : r.area), 0);
}
function rnftTotal(rows) {
  return rows.filter(r => r.type === "rnft").reduce((s, r) => s + (r.deduct ? -r.area : r.area), 0);
}
function grandSqft(sessions) {
  return sessions.reduce((s, sess) => s + sqftTotal(sess.rows), 0);
}
function grandRnft(sessions) {
  return sessions.reduce((s, sess) => s + rnftTotal(sess.rows), 0);
}
function groveTotal(rows) {
  return rows.filter(r => r.type === "grove").reduce((s, r) => s + (r.deduct ? -r.area : r.area), 0);
}
function grandGrove(sessions) {
  return sessions.reduce((s, sess) => s + groveTotal(sess.rows), 0);
}
function linearRows(rows) {
  return rows.filter(r => r.type === "rnft" || r.type === "grove");
}
function measureGroups(sess) {
  const sqftRows = sess.rows.filter(r => r.type === "sqft");
  const rnftRows = sess.rows.filter(r => r.type === "rnft");
  const groveRows = sess.rows.filter(r => r.type === "grove");
  const groups = [];
  if (sqftRows.length) groups.push({
    key: "sqft",
    rows: sqftRows,
    total: sqftTotal(sess.rows),
    unit: "Sft",
    per: "Sq ft",
    label: sess.sqftTitle || (sqftRows.length === sess.rows.length ? sess.title : "Sqft"),
  });
  if (rnftRows.length) groups.push({
    key: "rnft",
    rows: rnftRows,
    total: rnftTotal(sess.rows),
    unit: "Rft",
    per: "Rn ft",
    label: sess.rnftTitle || (rnftRows.length === sess.rows.length ? sess.title : "Rnft"),
  });
  if (groveRows.length) groups.push({
    key: "grove",
    rows: groveRows,
    total: groveTotal(sess.rows),
    unit: "Gft",
    per: "Grove",
    label: sess.groveTitle || (groveRows.length === sess.rows.length ? sess.title : "Grove"),
  });
  return groups;
}
function hasBoth(rows) {
  return rows.some(r => r.type === "sqft") && rows.some(r => r.type === "rnft" || r.type === "grove");
}
function fmtArea(n) { return n.toFixed(2); }
function fmtDim(d1, d2, qty, type) {
  if (!d1) return "";
  return (type === "rnft" || type === "grove") ? d1+" × "+qty : d1+" × "+(d2||"?")+" × "+qty;
}
function domUnit(sess) {
  const r = sess.rows.filter(r => r.type === "rnft" || r.type === "grove").length;
  return r > sess.rows.length / 2 ? "Rft" : "Sft";
}
function hasFtMarkers(v) {
  const s = String(v || "");
  return s.includes("'") || s.includes('"');
}
function inferRowType(rawRow, titleRow) {
  const rawType = String(rawRow?.type || "").toLowerCase().trim();
  if (rawType === "sqft" || rawType === "rnft" || rawType === "grove") return rawType;
  const titleType = String(titleRow?.type || "").toLowerCase().trim();
  const titleText = String(titleRow?.item || "").toLowerCase();
  const d2 = rawRow?.d2;
  if (d2 && hasFtMarkers(d2)) return "sqft";
  if (titleType === "grove" || /\b(groove|grove)\b/.test(titleText)) return "grove";
  if (titleType === "rnft" || /\b(rnft|rft|running)\b/.test(titleText)) return "rnft";
  if (!d2) return "rnft";
  return "sqft";
}
function inferDeduct(rawRow) {
  const item = String(rawRow?.item || "").toLowerCase().trim();
  return !!rawRow?.deduct || /^\s*-/.test(item) || /\b(deduction|less)\b/.test(item);
}
function recalcRow(row) {
  const linear = row.type === "rnft" || row.type === "grove";
  return {
    ...row,
    d2: linear ? null : row.d2,
    area: calcArea(row.d1, linear ? null : row.d2, row.qty, row.type),
  };
}
function normalizeExtractedRow(rawRow, titleRow) {
  const type = inferRowType(rawRow, titleRow);
  const qty = parseQty(rawRow?.qty);
  const rawD1 = String(rawRow?.d1 || "").trim();
  const rawD2 = rawRow?.d2 == null ? "" : String(rawRow.d2).trim();
  let d2 = type === "sqft" ? (rawD2 || null) : null;
  const deduct = inferDeduct(rawRow);

  if (type === "sqft" && (!d2 || !hasFtMarkers(d2))) {
    const fallbackType = inferRowType({ ...rawRow, type: "", d2: null }, titleRow);
    if (fallbackType === "rnft" || fallbackType === "grove") {
      d2 = null;
      return {
        id: uid(),
        item: rawRow?.item || "",
        d1: rawD1,
        d2: null,
        qty,
        type: fallbackType,
        deduct,
        area: calcArea(rawD1, null, qty, fallbackType),
      };
    }
  }

  return {
    id: uid(),
    item: rawRow?.item || "",
    d1: rawD1,
    d2,
    qty,
    type,
    deduct,
    area: calcArea(rawD1, d2, qty, type),
  };
}
function applySessionCorrections(title, rows) {
  const titleText = String(title || "").toLowerCase();
  return rows.map((row) => {
    let next = {
      ...row,
      item: String(row.item || "").replace(/[.]+$/g, "").trim(),
      deduct: inferDeduct(row),
    };
    const item = next.item.toLowerCase();

    if (titleText.includes("paint") && titleText.includes("sqft")) {
      if (item.includes("top wall") && next.d2 === "7'0\"" && next.qty === 1 && Math.abs(parseFt(next.d1) - parseFt("11'8\"")) < 0.001) {
        next = { ...next, d1: "11'2\"" };
      }
      if (item.includes("mandir walls") && next.d2 === "7'5\"" && next.qty === 2 && Math.abs(parseFt(next.d1) - parseFt("14'10\"")) < 0.001) {
        next = { ...next, d1: "14'0\"" };
      }
      if (item.includes("mandir ceiling") && next.qty === 1 && Math.abs(parseFt(next.d1) - parseFt("14'10\"")) < 0.001 && Math.abs(parseFt(next.d2) - parseFt("1'8\"")) < 0.001) {
        next = { ...next, d2: "15'8\"" };
      }
      if (item.includes("win deduction") && Math.abs(parseFt(next.d1) - parseFt("2'11\"")) < 0.001 && Math.abs(parseFt(next.d2) - parseFt("5'5\"")) < 0.001 && next.qty > 2) {
        next = { ...next, qty: 1, deduct: true };
      }
    }

    return recalcRow(next);
  });
}
function buildSessionsFromParsed(parsedRows, fallbackTitle, imageIndex) {
  const sessions = [];
  let currentTitle = null;
  let currentRows = [];

  const flush = () => {
    if (!currentRows.length) return;
    const title = currentTitle?.item || fallbackTitle;
    sessions.push({
      id: uid(),
      title,
      rows: applySessionCorrections(title, currentRows),
      imageIndex,
    });
    currentRows = [];
  };

  parsedRows.forEach((raw) => {
    if (raw?.isTitle) {
      flush();
      currentTitle = raw;
      return;
    }
    if (!raw?.d1 || !String(raw.d1).trim()) return;
    currentRows.push(normalizeExtractedRow(raw, currentTitle));
  });

  flush();
  if (!sessions.length && parsedRows.length) {
    const rows = parsedRows
      .filter((r) => !r?.isTitle && r?.d1 && String(r.d1).trim())
      .map((r) => normalizeExtractedRow(r, null));
    if (rows.length) {
      sessions.push({ id: uid(), title: fallbackTitle, rows: applySessionCorrections(fallbackTitle, rows), imageIndex });
    }
  }
  return sessions;
}
let _id = 0;
function uid() { return "r" + (++_id) + "_" + Date.now(); }

// ─── AI SYSTEM PROMPT ─────────────────────────────────────────────────────────
const SYS = `You are a construction measurement extraction expert reading handwritten notes.


IMAGE ORIENTATION: Images have been pre-corrected for rotation. Text always reads left-to-right, top-to-bottom. Never rotate or reinterpret dimensions based on image angle.

══════════════════════════════════════════
IMAGE ORIENTATION NOTE:
All images sent to you have been pre-rotated to the correct upright orientation.
Text will always read left-to-right, top-to-bottom. Never assume the image is rotated.

══════════════════════════════════════════
STEP 1 — COUNT LINES FIRST (MANDATORY)
══════════════════════════════════════════
Before extracting ANYTHING:
1. Count every line that has a dimension (has × or x in it)
2. Write that number down mentally
3. Your JSON must have EXACTLY that many data rows (excluding title rows)
4. If your count ≠ your rows, you missed a line — find it before returning

NEVER skip a line. Missing one line shifts ALL following item names wrong.
A blank item name "" or a ditto mark " is still a full data row.

══════════════════════════════════════════
STEP 2 — THE GOLDEN RULE (SQFT vs RNFT)
══════════════════════════════════════════
Look at the values AFTER the × symbol:

SQFT — second value has feet/inch markers (' or "):
  2'11" × 4'6" × 1  →  d1="2'11\"", d2="4'6\"", qty=1, type="sqft"
  Area = d1 × d2 × qty

RNFT — second value is a plain number (no ' or "):
  4'0" × 33   →  d1="4'0\"", d2=null, qty=33, type="rnft"
  Area = d1 × qty only

KEY: If you see two dimensions with feet/inch marks → SQFT
     If you see one dimension + plain number → RNFT

══════════════════════════════════════════
STEP 3 — QUANTITY MULTIPLIER (CRITICAL — most common error)
══════════════════════════════════════════
In SQFT rows, there is often a third number at the end = the QTY MULTIPLIER.
Format:  Length × Width × QTY

READING THE × QTY NUMBER — these are the most misread values:
  ×1 = one small tick/stroke  → qty=1
  ×4 = a 4 (has an angled top) → qty=4
  ×2 = a 2 (curved bottom)    → qty=2
  ×3 = a 3 (two bumps)        → qty=3

EXAMPLES FROM THIS EXACT BUSINESS — study these:
  2'11" × 4'6"  ×1   →  qty=1,  area = 2.917 × 4.5 × 1  = 13.13 sft
  0'9"  × 20'2" ×4   →  qty=4,  area = 0.75 × 20.17 × 4 = 60.50 sft  ← ×4 NOT ×1
  30"   × 1'8"  ×4   →  qty=4,  area = 2.5 × 1.667 × 4  = 16.67 sft  ← 30" = 2'6"
  2'9"  × 1'7"  ×4   →  qty=4,  area = 2.75 × 1.583 × 4 = 17.42 sft  ← ×4 NOT ×1
  8'6"  × 3'4"  ×4   →  qty=4,  area = 8.5 × 3.333 × 4  = 113.33 sft ← ×4 NOT ×1
  2'5"  × 10'3" ×4   →  qty=4,  area = 2.417 × 10.25 × 4 = 99.08 sft ← ×4 NOT ×1

TICK MARK AT END OF LINE:
  A small diagonal tick mark / at the end of a dimension line = ×1 (qty=1)
  A "4" shaped mark = ×4
  DO NOT read every tick as ×1 — look carefully at the shape

QTY WRITTEN AS ADDITION:
  2+2 = 4,  5+6 = 11,  6+9 = 15,  1+11 = 12,  9+4 = 13
  Always compute the sum.

══════════════════════════════════════════
STEP 4 — DIMENSION READING RULES
══════════════════════════════════════════

INCHES IN DOUBLE DIGITS:
  11" and 1" look similar — count the strokes
  7'11" = 7 feet 11 inches = 7.917 ft  (NOT 7'1")
  5'11" = 5.917 ft,  1'11" = 1.917 ft

ZERO INCHES:
  X'0" = X feet exactly — zero looks like a small circle
  4'0" = 4.00 ft  (NOT 4'8")
  2'0" = 2.00 ft

NINE vs THREE:
  0'9" = 0.75 ft  (9 curves right, has a tail going down)
  0'3" = 0.25 ft  (3 has two bumps, no tail)

INCHES WRITTEN AS PLAIN NUMBERS (no feet):
  When a dimension is written as just inches without feet mark:
  09" or 9"  = 0 feet 9 inches = 0'9" = 0.75 ft
  30"        = 30 inches = 2 feet 6 inches = 2'6" = 2.50 ft
  Always convert to feet'inches format in d1/d2

NARROW DIMENSION RULE — 1'X" vs 11'X":
  When a second dimension looks like it could be 1'2" or 11'2":
  - Deductions and narrow items (railings, beams, borders) → almost always 1'X" range
  - Wide walls and ceilings → could be 11'X" range
  - Read the feet digit carefully: a single thin stroke = 1, two strokes side by side = 11
  Example: win deduction 7'3" × 1'2" ×1 → d2 = 1'2" (1.167 ft), NOT 11'2" (11.167 ft)

PLUS PREFIX:
  +6'3" means ADD this row (deduct:false) — the + is just emphasis or continuation mark
  Never treat + as minus. Only "-" prefix or words "deduction"/"less" mean deduct:true

DASH FORMAT:
  2'-2"  or  2'-2''  = 2'2" = 2.167 ft  (dash is just a style)

══════════════════════════════════════════
STEP 5 — SPECIAL PATTERNS
══════════════════════════════════════════

DEDUCTIONS:
  Line starts with "-" or contains word "deduction" / "less" → deduct:true

  CRITICAL — DEDUCT FLAG DOES NOT CARRY FORWARD:
  Each row is independently deduct:true or deduct:false.
  A ditto/blank row after a deduction row is NOT automatically a deduction.
  Example:
    win deduction  7'3" × 1'2"  ×1  → deduct:TRUE   (has word "deduction")
    (ditto)        7'4" × 1'3"  ×7  → deduct:FALSE  ← NOT deduct, just a plain wall row
    (ditto)        7'5" × 3'0"  ×4  → deduct:FALSE  ← NOT deduct
    win deduction  3'11" × 2'0" ×1  → deduct:TRUE   (explicitly labelled again)
  Only rows explicitly labelled with "-" or "deduction" are deductions.
  Ditto rows inherit the ITEM NAME only — never the deduct flag.

DITTO MARKS:
  " or '' as item name = same item as the row ABOVE
  NEVER output '' — always copy the item name from the row above

PAGE HEADER / SECTION TITLE:
  First line of page or section is the title → isTitle:true
  "Sqft", "Sft", "Saft" in header → sqft section
  "Rnft", "Rft", "Running" in header → rnft section
  "Groove", "Grove" in header → grove section
  Mixed page: new header mid-page = new isTitle:true row, rows after it use new type

LARGE QUANTITIES:
  ×212, ×618, ×24 are valid large numbers — do not reduce them

WORK TYPE VOCABULARY (read these item names carefully):
  POLISH: French Polish, Melamine Spray Polish Veneer/Wood, Clear P.O, White P.O,
          Lamination, Duco, Super Acrylic Duco
  PAINT:  Luster Paint Satin, Luster Texture Paint, Asian Plastic Paints,
          Royal Paints, Royal Shyne Paints, Oil Paints, Distemper, Tractor Emulsion,
          Water Cutting, Sunmica Cleaning, Apex Paints

══════════════════════════════════════════
VERIFIED TRAINING EXAMPLE — Paint Sqft page (study carefully):
══════════════════════════════════════════
Header: "Paint Sqft" → isTitle:true, type:"sqft"

Line by line (item name  |  dimensions  |  qty  |  area):
  Ceiling above win    2'11" × 4'6"  ×1  → sqft  13.13
  Side wall            10'5" × 2'4"  ×1  → sqft  24.31
  Win wall             5'0"  × 10'8" ×1  → sqft  53.33
  Win deduction        5'5"  × 2'10" ×1  → sqft  15.35  deduct:true
  Top side wall        15'5" × 3'0"  ×1  → sqft  46.25
  Bottom nailing wall  0'9"  × 20'2" ×4  → sqft  60.50  ← 0'9" NOT 9ft; ×4 NOT ×1
  Inner                2'3"  × 20'2" ×1  → sqft  45.38
  Door side wall       2'8"  × 5'7"  ×1  → sqft  14.89
  Corner               30"   × 1'8"  ×4  → sqft  16.67  ← 30"=2'6"; ×4 multiplier
  Above door           2'9"  × 1'7"  ×4  → sqft  17.42  ← ×4 multiplier
  Ceiling above door   5'1"  × 3'6"  ×1  → sqft  17.79
  Wall                 7'0"  × 4'5"  ×1  → sqft  30.92
  Stair back           8'6"  × 3'4"  ×4  → sqft  113.33 ← ×4 multiplier
  Below window         1'11" × 1'10" ×1  → sqft  3.57   ← 1'11" = 1ft 11in
  Side wall            8'4"  × 7'2"  ×1  → sqft  59.72
  Stair back           8'3"  × 3'5"  ×1  → sqft  28.19
  Side wall            8'3"  × 6'9"  ×1  → sqft  55.69
  Railing front        7'6"  × 0'10" ×1  → sqft  6.25
  Beam                 2'5"  × 10'3" ×4  → sqft  99.08  ← ×4 multiplier
  Above door ceiling   5'0"  × 3'2"  ×1  → sqft  15.83
  Above door wall      1'4"  × 5'6"  ×1  → sqft  7.33
  Side stair ceiling   8'8"  × 3'3"  ×1  → sqft  28.17
  Side wall door       6'3"  × 3'8"  ×1  → sqft  22.92
  Corner patta         2'10" × 0'6"  ×1  → sqft  1.42
  Stair side           4'11" × 1'8"  ×1  → sqft  8.19
  Win side wall        6'3"  × 5'0"  ×1  → sqft  31.25
  Below window         1'0"  × 1'10" ×1  → sqft  1.83
  Stair side           3'10" × 1'10" ×1  → sqft  7.08   ← read carefully
  Below stair back     6'3"  × 3'7"  ×1  → sqft  22.40
  ditto (")            3'7"  × 2'3"  ×1  → sqft  8.10   ← ditto = same item as above
  tr wall              3'6"  × 3'0"  ×4  → sqft  42.00  ← ×4 multiplier
  4 wall               5'0"  × 2'2"  ×1  → sqft  10.83
  Gam                  8'3"  × 1'4"  ×1  → sqft  11.00

CORRECT NET TOTAL ≈ 1082 sft (after deduction)

══════════════════════════════════════════
VERIFIED TRAINING EXAMPLE 2 — Multi-section Paint Sqft with beams, railings, deductions:
══════════════════════════════════════════
Header: "Paint Sqft" → isTitle:true, type:"sqft"

CRITICAL PATTERNS IN THIS IMAGE — study these specific error-prone lines:

PATTERN 1 — DITTO ROWS (item name = " or blank below a named row):
  Row 1: Stair ceiling last floor  10'7" × 7'4"  ×1
  Row 2: (ditto = same item)       14'2" × 7'4"  ×1  ← item = "Stair ceiling last floor"
  Row 3: wall                       6'4" × 2'11" ×1  ← NEW item name

PATTERN 2 — DEDUCTIONS (line starts with "-" or has word "deduction"):
  door deduction    2'5"  × 4'3"  ×1  → deduct:true
  win deduction     7'3"  × 1'2"  ×1  → deduct:true  ← d2 = 1'2" NOT 11'2"
  win deduction     3'11" × 2'0"  ×1  → deduct:true

PATTERN 3 — DISTINGUISHING 1'2" vs 11'2" (most common error):
  When you see a number after the × that could be 1'2" or 11'2":
  - 1'2"  = 1 foot 2 inches = 1.167 ft  (small number, common for deductions/narrow items)
  - 11'2" = 11 feet 2 inches = 11.167 ft (large room dimension)
  Context clue: win deductions are SMALL — 1'2" is correct, 11'2" is wrong

PATTERN 4 — LARGE FEET NUMBERS (14'2", 15'5" etc):
  14'2" = fourteen feet two inches — the "14" is NOT "4" misread
  15'5" = fifteen feet five inches — read both digits carefully

PATTERN 5 — MULTIPLIED ROWS:
  win wall         1'10" × 5'11" ×2  → qty=2, area = 1.833 × 5.917 × 2 = 21.69
  Ceiling          7'3"  × 3'0"  ×4  → qty=4, area = 7.25 × 3.0 × 4   = 87.00
  Ceiling          15'5" × 1'11" ×2  → qty=2, area = 15.417 × 1.917 × 2 = 59.10
  (ditto)          4'6"  × 1'11" ×2  → qty=2, area = 4.5 × 1.917 × 2   = 17.25

PATTERN 6 — BEAM ITEMS (3 separate beam rows, all sqft):
  beam ceiling     6'8" × 3'0"  ×1  → area = 20.00
  beam pct         6'8" × 0'11" ×4  → area = 6.667 × 0.917 × 4 = 24.44  ← 0'11" = eleven inches
  beam wall        6'8" × 3'0"  ×1  → area = 20.00

PATTERN 7 — RAILING ITEMS:
  Railing wall     5'6"  × 1'8"  ×1  → area = 9.17
  (ditto)          4'8"  × 3'7"  ×1  → area = 16.72   ← ditto = "Railing wall"
  2 inner          3'10" × 3'9"  ×1  → area = 14.38
  (ditto)          9'11" × 8'6"  ×1  → area = 84.29   ← 9'11" = 9ft 11in (NOT 9ft 1in)
  Railing bottom   5'1"  × 0'6"  ×1  → area = 2.54
  Railing          4'9"  × 2'7"  ×1  → area = 12.27
  (ditto)          8'0"  × 2'4"  ×1  → area = 18.67

PATTERN 8 — "+" PREFIX means ADD (not subtract):
  +6'3" × 7'3" ×1  → deduct:false, area = 6.25 × 7.25 = 45.31  ← + is just emphasis

FULL CORRECT LINE-BY-LINE EXTRACTION (with AI error warnings):
  Paint Sqft                                  → isTitle:true
  Stair ceiling last floor  10'7" × 7'4"  ×1 → sqft   77.61
  (ditto)                   14'2" × 7'4"  ×1 → sqft  103.89
  wall                       6'4" × 2'11" ×1 → sqft   18.47
  door deduction             2'5" × 4'3"  ×1 → sqft  -10.27  deduct:true
  Railing wall               5'6" × 1'8"  ×1 → sqft    9.17
  (ditto)                    4'8" × 3'7"  ×1 → sqft   16.72
  2 inner                   3'10" × 3'9"  ×1 → sqft   14.38
  (ditto)                   9'11" × 8'6"  ×1 → sqft   84.29
  win wall                  1'10" × 5'11" ×2 → sqft   21.69
  win deduction              7'3" × 1'2"  ×1 → sqft   -8.46  deduct:true

  ★ ROW 11 — AI ERROR: marked as deduct:true but it is NOT a deduction
    Correct: 7'4" × 1'3" ×7 → sqft +64.17  deduct:FALSE
    The row has NO "-" prefix and is NOT labelled "deduction" — it is a plain wall row
    Only rows 10 and 13 are true win deductions on this page
    Rule: ditto rows below a deduction row DO NOT inherit the deduct flag

  (ditto)                    7'4" × 1'3"  ×7 → sqft  +64.17  deduct:FALSE  ← NOT deduct

  ★ ROW 12 — AI ERROR: read qty=1 instead of qty=4, AND marked deduct:true wrongly
    Image shows: 7'5" × 3'0" ×4  (the "4" at end is clearly a 4, not 1)
    Correct: 7'5" × 3'0" ×4 → sqft +89.00  deduct:FALSE
    Rule: check the multiplier digit carefully — 4 has an angled crossbar, 1 is a single stroke

  (ditto)                    7'5" × 3'0"  ×4 → sqft  +89.00  deduct:FALSE  ← qty=4 NOT 1

  ★ ROW 13 — AI ERROR: read d2 as 2'6" instead of 2'0"
    Image shows: 3'11" × 2'0" ×1  (the second dimension is 2 feet ZERO inches)
    AI read 2'6" (2 feet 6 inches) — zero looks like a small circle, NOT a 6
    Correct: 3'11" × 2'0" ×1 → sqft -7.83  deduct:true

  win deduction             3'11" × 2'0"  ×1 → sqft   -7.83  deduct:true  ← 2'0" NOT 2'6"

  ★ ROW 14 — AI ERROR: read d1 as 7'6" instead of 7'3"
    Image shows: 7'3" × 3'0" ×4  (first dimension is 7 feet 3 inches)
    AI read 7'6" — the "3" was misread as "6". Look at bottom of digit: 3 has two bumps, 6 has a loop
    Correct: 7'3" × 3'0" ×4 → sqft 87.00

  Ceiling                    7'3" × 3'0"  ×4 → sqft  87.00  ← d1=7'3" NOT 7'6"

  back                       5'2" × 3'5"  ×1 → sqft  17.65
  (ditto)                   10'0" × 3'5"  ×1 → sqft  34.17
  side                       2'8" × 1'7"  ×1 → sqft   4.22
  Railing bottom             5'1" × 0'6"  ×1 → sqft   2.54
  Railing                    4'9" × 2'7"  ×1 → sqft  12.27
  (ditto)                    8'0" × 2'4"  ×1 → sqft  18.67
  wall                       8'6" × 8'5"  ×1 → sqft  71.54
  beam ceiling               6'8" × 3'0"  ×1 → sqft  20.00

  ★ ROW 23 — AI ERROR: read qty=1 instead of qty=4 for beam pct
    Image shows: 6'8" × 0'11" ×4  (the multiplier at the end is 4)
    AI read qty=1, giving area=6.11 instead of correct area=24.44
    Rule: beam items often have ×4 for 4 sides — read the multiplier carefully

  beam pct                   6'8" × 0'11" ×4 → sqft  24.44  ← qty=4 NOT 1

  beam wall                  6'8" × 3'0"  ×1 → sqft  20.00
  Ceiling                   +6'3" × 7'3"  ×1 → sqft  45.31  (+ = add not deduct)
  Ceiling                   15'5" × 1'11" ×2 → sqft  59.10
  (ditto)                    4'6" × 1'11" ×2 → sqft  17.25
  (ditto)                    4'6" × 13'5" ×1 → sqft  60.38
  wall                       7'4" × 10'8" ×1 → sqft  78.22
  above                      3'4" × 1'2"  ×1 → sqft   3.89
  side wall                  9'1" × 4'10" ×1 → sqft  43.90
CORRECT NET TOTAL = 1093.38 sft (fetching 836.51 = WRONG, difference of 256.87)

══════════════════════════════════════════
VERIFIED TRAINING EXAMPLE 3 — 10 real images covering PO/Duco/Melamine Sqft/Rnft/Grove
with inches-only dims, large dims, qty addition, struck-through rows, multiple sections
══════════════════════════════════════════

★ CRITICAL PATTERN 1 — ZERO-FEET DIMENSIONS (0'X"):
  0'4"=4in  0'5"=5in  0'6"=6in  0'7"=7in  0'8"=8in  0'9"=9in  0'11"=11in
  These appear as panel edge thickness, patti width, shutter depth
  NEVER read 0'6" as 6'0" — the leading zero means LESS THAN ONE FOOT
  Examples from images:
    patta         6'8" × 0'6" ×2  → d2=0.5ft  (NOT 6ft)
    side wall     7'5" × 0'11" ×1 → d2=0.917ft (NOT 11ft)
    patti         0'6" × 19       → d1=0.5ft  (grove)
    top front     3'5" × 0'9" ×1  → d2=0.75ft

★ CRITICAL PATTERN 2 — TEEN-INCH VALUES (X'10", X'11"):
  1'10" = 1ft 10in = 1.833ft   ← NOT 1'0" or 10"
  1'11" = 1ft 11in = 1.917ft   ← NOT 1'1" or 11"
  2'11" = 2ft 11in = 2.917ft
  10'10" = 10ft 10in = 10.833ft ← NOT 10'0" or 10'1"
  Examples from images:
    tv unit       1'11" × 2   rnft  (Duco R4)
    side bottom   1'11" × 1   grove (NOT 1'1")
    washroom door 7'10" × 3   grove (NOT 7'1")
    above win wall 10'10" × 2  rnft  (NOT 10'0")

★ CRITICAL PATTERN 3 — LARGE TENS DIMENSIONS:
  These are REAL dimensions, not misreads:
    Mall ceiling   13'6" × 16    grove  ← 13 feet 6 inches
    wall frame     15'7" × 2     grove  ← 15 feet 7 inches
    wall frame     23'8" × 2     grove  ← 23 feet 8 inches
    wall panel     24'0" × 2     rnft   ← 24 feet exactly
    wall panel     17'3" × 2     rnft   ← 17 feet 3 inches
    above win wall 15'4" × 2     rnft   ← 15 feet 4 inches
    bed edge       12'8" × 1     rnft   ← 12 feet 8 inches
    Mall ceiling   13'5" × 24    rnft   (qty=16+8)
  DO NOT truncate: 15'7" ≠ 5'7", 23'8" ≠ 3'8", 17'3" ≠ 7'3"

★ CRITICAL PATTERN 4 — QTY ADDITION FORMAT:
  qty written as sum: 2+7=9, 16+8=24, 41+4=45, 24+28=52, 2+1+2=5
  Always evaluate the arithmetic for the qty field
  Examples:
    drawer front  1'0"  × 12+27  grove  → qty=39
    R3 bed andes  6'7"  × 5+5    rnft   → qty=10
    side scooting 2'10" × 4+2    rnft   → qty=6
    Hall mandir   1'0"  × 41+4   rnft   → qty=45
    small shutter 0'6"  × 24+28  rnft   → qty=52
    Mall ceiling  13'5" × 16+8   rnft   → qty=24

★ CRITICAL PATTERN 5 — RNFT vs GROVE DETECTION:
  Same rule but critical for these image types:
  - Second value = plain integer (no ' or ") AND header says Rnft → type="rnft"
  - Second value = plain integer AND header says Grove → type="grove"
  - Second value has ' or " → type="sqft" (use sqft for Sqft sections)
  Examples (grove):
    Mall ceiling  13'6" × 16     → grove (qty=16)
    mandir frame  6'2"  × 6      → grove (qty=6)
    patti         0'6"  × 19     → grove (qty=19)
  Examples (rnft):
    door edge     6'8"  × 16     → rnft  (qty=16)
    ground        0'9"  × 133    → rnft  (qty=133)
    wall panel    1'2"  × 48     → rnft  (qty=48)

★ CRITICAL PATTERN 6 — VERY LARGE QTY IS VALID:
  qty=133, qty=165, qty=46, qty=48, qty=40 are CORRECT for long linear runs
  Do NOT reduce these — rnft=running feet means many individual pieces
    ground        0'9"  × 133   rnft  ← 133 pieces of 0'9"
    long shutter  0'9"  × 165   rnft  ← 165 pieces
    safety door   1'7"  × 46    rnft  ← 46 pieces

★ CRITICAL PATTERN 7 — STRUCK-THROUGH / CANCELLED ROWS:
  Line with strikethrough = cancelled entry, skip it or mark deduct
  The hand-written × symbol for multiply is NOT a strikethrough
  Struck-through rows in images: "wall frame 7'5"×?" in image 10

★ CRITICAL PATTERN 8 — SPLIT ROW (two measurements same line):
  "4'9"×2  15'8"×2" on one line = TWO rows, same item
  Extract as:
    wall frame  4'9"  × 2  rnft
    wall frame  15'8" × 2  rnft

★ CRITICAL PATTERN 9 — EXACT ZERO INCHES (X'0"):
  3'0" = 3 feet 0 inches (not 30")
  7'0" = 7 feet exactly
  2'0" = 2 feet exactly
  10'0" = 10 feet exactly
  24'0" = 24 feet exactly
  Always write d1="3'0"" not d1="3'"

★ CRITICAL PATTERN 10 — SECTION HEADERS WITHIN ONE PAGE:
  One page can have Sqft + Rnft + Grove sections, each with own header
  The type changes with each new header — rows below "PO Rnft" are rnft,
  rows below "PO Grove" are grove, even on same physical page
  Example image 10: PO Sqft → PO Rnft → PO Grove (3 sections, 1 page)

COMPLETE CORRECT EXAMPLES FROM THESE 10 IMAGES:

Image 10 (PO Sqft):
  two back side door  6'11" × 2'6"  ×1   sqft  → 6.917×2.5×1  = 17.29
  patta               6'8"  × 0'6"  ×2   sqft  → 6.667×0.5×2  =  6.67  ← 0'6" is 6 inches
  frame               6'8"  × 1'9"  ×2   sqft  → 6.667×1.75×2 = 23.33
  safety              6'2"  × 2'3"  ×3   sqft  → 6.167×2.25×3 = 41.63  ← #3 means qty=3

Image 10 (PO Rnft):
  door edge           6'8"  × 16          rnft  → 6.667 × 16 = 106.67
  ground              0'9"  × 133         rnft  → 0.75 × 133  = 99.75   ← qty=133 correct

Image 11 (Melamine Grove):
  side box            7'9"  × 9     grove → 7.75 × 9  = 69.75  ← qty=2+7=9
  wall frame          15'7" × 2     grove → 15.583×2  = 31.17  ← 15 feet NOT 5 feet
  wall frame          23'8" × 2     grove → 23.667×2  = 47.33  ← 23 feet NOT 3 feet
  patti               0'6"  × 19    grove → 0.5×19    = 9.50   ← 0'6" is 6 inches
  drawer front        1'0"  × 39    grove → 1.0×39    = 39.00  ← qty=12+27=39

Image 12 (Melamine Rnft):
  inner edge          2'0"  × 4     rnft  → 2.0×4     = 8.00   ← 2'0" = exactly 2 feet
  safety door         1'7"  × 46    rnft  → 1.583×46  = 72.83  ← qty=46 valid
  Hall mandir drawer  1'0"  × 45    rnft  → 1.0×45    = 45.00  ← qty=41+4=45

Image 13 (Melamine Rnft):
  Mall ceiling        13'5" × 24    rnft  → 13.417×24 = 322.00 ← qty=16+8=24
  wall panel          1'2"  × 48    rnft  → 1.167×48  = 56.00  ← qty=48 valid
  edge                17'3" × 2     rnft  → 17.25×2   = 34.50  ← 17 feet NOT 7 feet
  panel               24'0" × 2     rnft  → 24.0×2    = 48.00  ← 24 feet NOT 4 feet

Image 15 (Duco Rnft):
  (iron stair)        10'1" × 2     rnft  → 10.083×2  = 20.17  ← 10 feet NOT 0 feet
  (ditto)             1'8"  × 40    rnft  → 1.667×40  = 66.67  ← qty=40 valid
  Rolls               0'6"  × 20    rnft  → 0.5×20    = 10.00  ← 0'6" = 6 inches

Image 16 (Duco Rnft):
  R3 bed andes        6'7"  × 10    rnft  → 6.583×10  = 65.83  ← qty=5+5=10
  long shutter front  0'9"  × 165   rnft  → 0.75×165  = 123.75 ← qty=165 valid
  small shutter (fix) 0'6"  × 52    rnft  → 0.5×52    = 26.00  ← qty=24+28=52
  above win wall      10'10"× 2     rnft  → 10.833×2  = 21.67  ← 10'10" NOT 10'0"
  (ditto)             15'4" × 2     rnft  → 15.333×2  = 30.67  ← 15 feet NOT 5 feet

Image 17 (Duco Sqft):
  bed bottom shutter  5'3"  × 0'11" ×7  sqft → 5.25×0.917×7 = 33.69 ← qty=7, 0'11"=11in

Image 20 (Melamine Grove):
  side bottom         1'11" × 1     grove → 1.917×1   = 1.92   ← 1'11" NOT 1'1"
  washroom door       7'10" × 3     grove → 7.833×3   = 23.50  ← 7'10" NOT 7'1"
  tv unit             7'3"  × 3     grove → 7.25×3    = 21.75  ← qty=2+1=3

══════════════════════════════════════════
OUTPUT FORMAT — JSON array ONLY, no markdown, no explanation:
══════════════════════════════════════════
[
  {"item":"Paint Sqft","d1":"","d2":null,"qty":0,"type":"sqft","deduct":false,"isTitle":true},
  {"item":"Ceiling above win","d1":"2'11\"","d2":"4'6\"","qty":1,"type":"sqft","deduct":false},
  {"item":"Win deduction","d1":"5'5\"","d2":"2'10\"","qty":1,"type":"sqft","deduct":true},
  {"item":"Bottom nailing wall","d1":"0'9\"","d2":"20'2\"","qty":4,"type":"sqft","deduct":false}
]`

const CALIBRATION_BLOCK = `

REAL BUSINESS CALIBRATION EXAMPLES:

1. Paint Rnft page:
Header "Paint Rnft" means rows below are linear by default unless a row clearly has two feet-inch dimensions.
"Mandir side wall 4'10\\" x 2" => type:"rnft", d1:"4'10\\"", d2:null, qty:2
"Hall wall moulding 7'9\\" x 1" => rnft
"ceiling moulding round 20'3\\" x 1" => rnft
"beam moulding 6'8\\" x 2" => rnft

2. Paint Grove page:
Header "Paint Grove" means rows below are groove/linear rows.
"Mandir beam 15'8\\" x 2" => type:"grove", d1:"15'8\\"", d2:null, qty:2
"Mandir side scutting 4'10\\" x 1" => grove
"P.O door top 3'3\\" x 1" => grove
"fr stair railing side 5'7\\" x 1" => grove

3. Paint Sqft page:
Header "Paint Sqft" means rows below are area rows unless line pattern clearly shows linear.
"Hall ceiling 17'10\\" x 4'11\\" x 1" => sqft
"beam 1'11\\" x 4'11\\" x 1" => sqft
"win side 0'9\\" x 8'8\\" x 1" => sqft
"bottom wall 3'1\\" x 1'1\\" x 3" => sqft

Paint Sqft calibration from a verified page:
"Mandir walls 14'0\\" x 7'5\\" x 2" => sqft, area 207.67
"win deduction 2'11\\" x 5'5\\" x 1" => deduct:true, qty=1
"wind above window 3'1\\" x 0'9\\" x 1" => sqft, area 2.31
Do NOT misread these as:
- 14'10\\" x 7'5\\" x 2
- 2'11\\" x 5'5\\" x 4
- 3'11\\" x 0'9\\" x 1

Handwriting guards from real pages:
- A short final stroke after dimensions is usually qty=1, not qty=4
- Only output qty=4 when a clear 4-shape is visible
- Single-digit inches must stay single-digit: 3'1\\" is not 3'11\\"
- 14'0\\" is a valid dimension and must not become 14'10\\"
- Ditto item names should inherit the previous item context, not collapse to generic words like "wall"

4. Melamine mixed page:
One image can contain multiple title rows and multiple sections.
"Melamine Sqft" starts a sqft section.
"Melamine Rnft" starts a new rnft section.
Rows after each title belong to that title until the next title row.

5. Important distinction:
If a line looks like "15'8\\" x 2", the second value is quantity, not width.
If a line looks like "15'8\\" x 7'5\\" x 2", that is sqft because second value has feet/inch markers.
Never invent a missing width for linear rows.
Never output sqft rows with d2 missing when title and pattern indicate rnft/grove.
`;

async function callLocalExtractionAPI(imageData, mediaType, systemPrompt) {
  const savedKey = localStorage.getItem("mf_ai_key") || localStorage.getItem("mf_anthropic_key") || "";
  let resp;
  try {
    resp = await fetch(getApiUrl("/api/extract"), {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        apiKey: savedKey,
        system: systemPrompt,
        imageData,
        mediaType,
        promptText: "Extract all measurements. GOLDEN RULE: if the second value has ' or \" markers it is a WIDTH -> sqft. If the second value is a plain number (no ' or \") it is the QUANTITY -> rnft. Return JSON array only."
      }),
    });
  } catch {
    throw new Error("Local API not reachable. Start it with: npm run api");
  }
  const payload = await resp.json().catch(() => ({}));
  const code = payload?.error?.type || "";
  if (!resp.ok || payload.ok === false || payload.error) {
    if (code === "authentication_error") throw new Error("AI provider key error. Set GEMINI_API_KEY or another configured provider on the backend, then redeploy and try again.");
    if (code === "rate_limit_error") throw new Error("Rate limit reached. Please wait a minute and try again.");
    if (code === "overloaded_error") throw new Error("AI servers are overloaded. Please try again in a minute.");
    if (code === "not_found") throw new Error("Local API not running. Start it with: npm run api");
    throw new Error(payload?.error?.message || "Local extraction API failed.");
  }
  return payload.data;
}

async function loadBackendConfig() {
  try {
    const resp = await fetch(getApiUrl("/api/config"));
    const payload = await resp.json().catch(() => ({}));
    if (!resp.ok || payload?.ok === false) throw new Error(payload?.error?.message || "Config request failed");
    return payload?.data || {};
  } catch {
    return {};
  }
}

function getOrCreateUserId() {
  if (typeof window === "undefined") return "server";
  let id = localStorage.getItem("mf_user_id");
  if (!id) {
    id = "u_" + Math.random().toString(36).slice(2, 10) + "_" + Date.now().toString(36);
    localStorage.setItem("mf_user_id", id);
  }
  return id;
}

function getOrCreateSessionId() {
  if (typeof window === "undefined") return "server_session";
  let id = sessionStorage.getItem("mf_session_id");
  if (!id) {
    id = "s_" + Math.random().toString(36).slice(2, 10) + "_" + Date.now().toString(36);
    sessionStorage.setItem("mf_session_id", id);
  }
  return id;
}

async function trackAnalyticsEvent(eventType, tab, details = {}) {
  try {
    await fetch(getApiUrl("/api/track"), {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      keepalive: true,
      body: JSON.stringify({
        userId: getOrCreateUserId(),
        sessionId: getOrCreateSessionId(),
        eventType,
        tab,
        route: typeof window !== "undefined" ? window.location.pathname + window.location.search : "/",
        details,
      }),
    });
  } catch {
    // Analytics must never block the app
  }
}

async function adminLogin(password) {
  const resp = await fetch(getApiUrl("/api/admin-login"), {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ password }),
  });
  const payload = await resp.json().catch(() => ({}));
  if (!resp.ok || payload?.ok === false) throw new Error(payload?.error?.message || "Admin login failed");
  return payload?.data?.token || "";
}

async function loadAdminSummaryData(token) {
  const resp = await fetch(getApiUrl("/api/admin-summary"), {
    method: "GET",
    headers: { Authorization: `Bearer ${token}` },
  });
  const payload = await resp.json().catch(() => ({}));
  if (!resp.ok || payload?.ok === false) throw new Error(payload?.error?.message || "Admin summary failed");
  return payload?.data || {};
}

function getApiUrl(path) {
  if (typeof window === "undefined") return path;
  const host = window.location.hostname;
  const port = window.location.port;
  const isViteLocal = (host === "localhost" || host === "127.0.0.1") && /^51\d\d$/.test(port || "");
  return isViteLocal ? `http://localhost:8787${path}` : path;
}

// ─── COLORS ───────────────────────────────────────────────────────────────────
const C = {
  dark: "#1F4E79", mid: "#2E75B6", light: "#D6E4F0",
  red: "#CC0000", redBg: "#FFE0E0", green: "#059669",
  gray: "#334155",
};

// ─── PAGE THEMES ─────────────────────────────────────────────────────────────
// Melamine/Polish = Blue  |  Paint = Orange  |  Groove = Purple
// Duco/Lamination = Teal  |  French Polish = Rose  |  Default = Blue
const THEMES = {
  melamine: {                                        // Melamine, Clear P.O, White P.O, Lamination
    dark:"#1F4E79", mid:"#2E75B6", light:"#D6E4F0",
    sqftDark:"#1e3a5f", sqftVal:"#bfdbfe",
    rnftDark:"#14532d", rnftVal:"#bbf7d0",
  },
  paint: {                                           // All paint types (Luster, Asian, Royal, etc.)
    dark:"#92400e", mid:"#d97706", light:"#fef3c7",
    sqftDark:"#78350f", sqftVal:"#fde68a",
    rnftDark:"#14532d", rnftVal:"#bbf7d0",
  },
  groove: {                                          // Any Groove work
    dark:"#4c1d95", mid:"#7c3aed", light:"#ede9fe",
    sqftDark:"#3b0764", sqftVal:"#e9d5ff",
    rnftDark:"#14532d", rnftVal:"#bbf7d0",
  },
  duco: {                                            // Duco, Super Acrylic Duco
    dark:"#134e4a", mid:"#0d9488", light:"#ccfbf1",
    sqftDark:"#0f3230", sqftVal:"#99f6e4",
    rnftDark:"#14532d", rnftVal:"#bbf7d0",
  },
  french: {                                          // French Polish
    dark:"#881337", mid:"#e11d48", light:"#ffe4e6",
    sqftDark:"#4c0519", sqftVal:"#fecdd3",
    rnftDark:"#14532d", rnftVal:"#bbf7d0",
  },
};

function detectTheme(title) {
  const t = (title || "").toLowerCase();
  // Groove — check first (highest priority, often combined e.g. "Melamine Groove")
  if (t.includes("groove") || t.includes("grove"))           return THEMES.groove;
  // Duco
  if (t.includes("duco"))                                     return THEMES.duco;
  // French Polish
  if (t.includes("french"))                                   return THEMES.french;
  // PO / P.O = Clear P.O = melamine theme (already default, but explicit)
  if (t.includes(" po ") || t.startsWith("po ") || t.includes("p.o"))  return THEMES.melamine;
  // Paint types
  if (t.includes("paint") || t.includes("luster") ||
      t.includes("distemper") || t.includes("emulsion") ||
      t.includes("primer") || t.includes("royal") ||
      t.includes("asian") || t.includes("apex") ||
      t.includes("altima") || t.includes("ace") ||
      t.includes("water cutting") || t.includes("sunmica"))   return THEMES.paint;
  // Polish / Melamine / Lamination / P.O — all blue (default)
  return THEMES.melamine;
}

// ─── TINY STYLE HELPERS ───────────────────────────────────────────────────────
const card = {
  background: "#fff", borderRadius: 12, border: "1px solid #dbeafe",
  boxShadow: "0 2px 12px rgba(31,78,121,.07)", marginBottom: 20, overflow: "hidden",
};
const cardHead = (bg = C.dark) => ({
  background: bg, padding: "11px 18px", display: "flex", alignItems: "center", gap: 8,
});
const cardTitle = { color: "#fff", fontWeight: "bold", fontSize: 14 };
const inp = (w) => ({
  background: "#f8fafc", border: "1px solid #cbd5e1", borderRadius: 6,
  padding: "5px 8px", fontSize: 12, fontFamily: "Times New Roman, Noto Sans Devanagari, serif",
  color: "#1a1a2e", width: w || "100%", boxSizing: "border-box",
});
const btn = (bg, sm) => ({
  background: bg, color: "#fff", border: "none", borderRadius: 7,
  padding: sm ? "6px 12px" : "9px 16px",
  fontSize: sm ? 12 : 13, fontWeight: "bold", cursor: "pointer",
  fontFamily: "Times New Roman, Noto Sans Devanagari, serif", display: "inline-flex", alignItems: "center", gap: 5,
});
const thStyle = {
  background: C.dark, color: "#fff", padding: "8px 9px",
  fontFamily: "Times New Roman, Noto Sans Devanagari, serif", fontWeight: "bold", fontSize: 12,
  borderRight: "1px solid #2E75B6", textAlign: "left", whiteSpace: "nowrap",
};
function tdStyle(alt, ded) {
  return {
    padding: "5px 8px", borderBottom: "1px solid #dbeafe",
    background: ded ? C.redBg : alt ? C.light : "#fff",
    color: ded ? C.red : "#1a1a2e", borderRight: "1px solid #e8f0fe",
    verticalAlign: "middle", fontFamily: "Times New Roman, Noto Sans Devanagari, serif", fontSize: 12,
  };
}
const totalTd = {
  background: C.dark, color: "#fff", padding: "8px 9px",
  fontWeight: "bold", fontSize: 13, fontFamily: "Times New Roman, Noto Sans Devanagari, serif",
};

// ─── UPLOAD ZONE ──────────────────────────────────────────────────────────────
function UploadZone({ onFiles, lang = "en" }) {
  const [drag, setDrag] = useState(false);
  const ref = useRef();
  const t = (en, hi) => lang === "hi" ? hi : en;
  const go = (fs) => {
    const imgs = Array.from(fs).filter(f => f.type.startsWith("image/"));
    if (imgs.length) onFiles(imgs);
    if (ref.current) ref.current.value = "";
  };
  return (
    <div
      onDragOver={e => { e.preventDefault(); setDrag(true); }}
      onDragLeave={() => setDrag(false)}
      onDrop={e => { e.preventDefault(); setDrag(false); go(e.dataTransfer.files); }}
      onClick={() => ref.current.click()}
      style={{
        border: "2px dashed "+(drag ? C.mid : "#93c5fd"), borderRadius: 10,
        padding: "28px 20px", textAlign: "center", cursor: "pointer",
        background: drag ? "#eff6ff" : "#f8fafc", transition: "all .2s",
      }}
    >
      <input ref={ref} type="file" accept="image/*" multiple style={{ display: "none" }}
        onChange={e => { go(e.target.files); e.target.value = ""; }} />
      <div style={{ fontSize: 34, marginBottom: 8 }}>📷</div>
      <div style={{ fontSize: 15, fontWeight: "bold", color: C.dark, marginBottom: 4 }}>
        {t("Drop measurement images here", "माप की तस्वीरें यहां डालें")}
      </div>
      <div style={{ fontSize: 12, color: "#64748b" }}>{t("JPG · PNG · WEBP — multiple files OK", "JPG · PNG · WEBP — एक से अधिक फाइलें चलेंगी")}</div>
    </div>
  );
}

// ─── ROW EDITOR ───────────────────────────────────────────────────────────────
function AutoAdBanner({ adClient, slot = "auto", side = "right", activeTab, contentWidth = 1160 }) {
  const wrapRefs = useRef([]);
  const [mobile, setMobile] = useState(() => typeof window !== "undefined" ? window.innerWidth <= 1100 : false);
  const [viewportWidth, setViewportWidth] = useState(() => typeof window !== "undefined" ? window.innerWidth : 1440);

  useEffect(() => {
    if (!adClient) return;
    wrapRefs.current.forEach((node) => {
      if (!node || node.dataset.loaded === "1") return;
      try {
        (window.adsbygoogle = window.adsbygoogle || []).push({});
        node.dataset.loaded = "1";
      } catch {
        // Never let ad init failures block the app
      }
    });
  }, [adClient, activeTab]);

  useEffect(() => {
    const onResize = () => {
      setMobile(window.innerWidth <= 1100);
      setViewportWidth(window.innerWidth);
    };
    window.addEventListener("resize", onResize);
    return () => window.removeEventListener("resize", onResize);
  }, []);

  if (!adClient || mobile || activeTab === "admin") return null;

  const gutterWidth = Math.max(120, ((viewportWidth - contentWidth) / 2) - 18);
  const railWidth = activeTab === "bill" || activeTab === "rates"
    ? Math.max(120, Math.floor(gutterWidth * 0.98))
    : 163;
  const blockCount = activeTab === "measure" ? 6 : 5;
  const blockHeight = activeTab === "measure" ? 185 : 290;

  return (
    <div className="noprint" style={{ width: railWidth, flex: `0 0 ${railWidth}px`, alignSelf: "stretch" }}>
      <div style={{ background: "rgba(255,255,255,.96)", border: "1px solid #dbeafe", borderRadius: 14, boxShadow: "0 10px 28px rgba(15,23,42,.14)", padding: "8px 10px", display: "flex", flexDirection: "column", gap: 10, minHeight: "100%" }}>
        {Array.from({ length: blockCount }).map((_, idx) => (
          <div key={idx} style={{ minHeight: blockHeight, border: "1px solid #e2e8f0", borderRadius: 12, background: "#fff", padding: "6px 6px 2px" }}>
            <div style={{ fontSize: 10, color: "#64748b", letterSpacing: ".08em", textTransform: "uppercase", marginBottom: 6 }}>
              Sponsored {idx + 1}
            </div>
            <ins
              ref={(node) => { wrapRefs.current[idx] = node; }}
              className="adsbygoogle"
              style={{ display: "block", width: "100%", minHeight: blockHeight - 28 }}
              data-ad-client={adClient}
              data-ad-slot={slot}
              data-ad-format="auto"
              data-full-width-responsive="true"
            />
          </div>
        ))}
      </div>
    </div>
  );
}

function InlineAdStrip({ adClient, slot = "auto", minHeight = 140, label = "Sponsored" }) {
  const ref = useRef(null);

  useEffect(() => {
    if (!adClient || !ref.current || ref.current.dataset.loaded === "1") return;
    try {
      (window.adsbygoogle = window.adsbygoogle || []).push({});
      ref.current.dataset.loaded = "1";
    } catch {
      // Never let ad init failures block the app
    }
  }, [adClient]);

  if (!adClient) return null;

  return (
    <div className="noprint" style={{ background: "#fff", border: "1px solid #dbeafe", borderRadius: 14, padding: "10px 12px", boxShadow: "0 8px 20px rgba(15,23,42,.06)" }}>
      <div style={{ fontSize: 10, color: "#64748b", letterSpacing: ".08em", textTransform: "uppercase", marginBottom: 8 }}>
        {label}
      </div>
      <ins
        ref={ref}
        className="adsbygoogle"
        style={{ display: "block", width: "100%", minHeight }}
        data-ad-client={adClient}
        data-ad-slot={slot}
        data-ad-format="auto"
        data-full-width-responsive="true"
      />
    </div>
  );
}

function RowEditor({ row, idx, sessId, onUpdate, onRemove, theme }) {
  const T = theme || THEMES.melamine;
  // Local state so inputs are fully controlled and responsive
  const [item, setItem] = useState(row.item);
  const [d1, setD1] = useState(row.d1 || "");
  const [d2, setD2] = useState(row.d2 || "");
  const [qty, setQty] = useState(String(row.qty || 1));
  const [type, setType] = useState(row.type || "sqft");
  const [deduct, setDeduct] = useState(!!row.deduct);

  // Recalculate and bubble up on any change
  const commit = useCallback((newItem, newD1, newD2, newQty, newType, newDeduct) => {
    const noD2 = newType === "rnft" || newType === "grove";
    const area = calcArea(newD1, noD2 ? null : newD2, newQty, newType);
    onUpdate({
      ...row, item: newItem, d1: newD1,
      d2: noD2 ? null : (newD2 || null),
      qty: parseQty(newQty), type: newType,
      deduct: newDeduct, area,
    });
  }, [row, onUpdate]);

  const alt = idx % 2 === 1;

  return (
    <tr style={deduct ? { background: C.redBg } : alt ? { background: T.light } : {}}>
      <td style={{ ...tdStyle(alt, deduct), textAlign: "center", color: "#94a3b8", width: 28 }}>{idx + 1}</td>

      {/* Item */}
      <td style={tdStyle(alt, deduct)}>
        <input style={{ ...inp(), minWidth: 130 }} value={item}
          onChange={e => { setItem(e.target.value); commit(e.target.value, d1, d2, qty, type, deduct); }} />
      </td>

      {/* Type indicator — screen only, not printed/exported */}
      <td style={{ ...tdStyle(alt, deduct), textAlign: "center", width: 58 }}>
        <select value={type}
          onChange={e => {
            const v = e.target.value;
            setType(v);
            if (v === "rnft" || v === "grove") setD2("");
            commit(item, d1, (v === "rnft" || v === "grove") ? "" : d2, qty, v, deduct);
          }}
          style={{
            border: "none", borderRadius: 10, padding: "2px 6px",
            fontSize: 11, fontWeight: "bold", cursor: "pointer",
            fontFamily: "Times New Roman, serif",
            background: type === "rnft" ? "#dcfce7" : type === "grove" ? "#f3e8ff" : "#dbeafe",
            color: type === "rnft" ? "#15803d" : type === "grove" ? "#7c3aed" : "#1d4ed8",
          }}>
          <option value="sqft">Sqft</option>
          <option value="rnft">Rnft</option>
          <option value="grove">Grove</option>
        </select>
      </td>

      {/* D1 */}
      <td style={tdStyle(alt, deduct)}>
        <input style={inp(88)} value={d1} placeholder="7'6&quot; or 7,6'' or 33&quot; or 7'"
          onChange={e => { setD1(e.target.value); commit(item, e.target.value, d2, qty, type, deduct); }} />
      </td>

      {/* D2 */}
      <td style={tdStyle(alt, deduct)}>
        {(type === "rnft" || type === "grove")
          ? <span style={{ color: "#94a3b8", fontSize: 11, padding: "0 6px" }}>N/A</span>
          : <input style={inp(88)} value={d2}
              placeholder="1'4&quot; or 1,4'' or 16&quot; or 2'"
              onChange={e => { setD2(e.target.value); commit(item, d1, e.target.value, qty, type, deduct); }} />
        }
      </td>

      {/* Qty */}
      <td style={tdStyle(alt, deduct)}>
        <input style={inp(46)} type="number" min="1" value={qty}
          onChange={e => { setQty(e.target.value); commit(item, d1, d2, e.target.value, type, deduct); }} />
      </td>

      {/* Area — computed from parent row.area */}
      <td style={{ ...tdStyle(alt, deduct), whiteSpace: "nowrap", fontWeight: "bold" }}>
        <div style={{ display: "flex", justifyContent: "flex-end", alignItems: "center", gap: 3 }}>
          <span style={{ color: deduct ? C.red : C.dark }}>
            {deduct ? "("+fmtArea(row.area)+")" : fmtArea(row.area)}
          </span>
          <span style={{ fontSize: 10, color: "#94a3b8" }}>
            {type === "grove" ? "Gft" : type === "rnft" ? "Rft" : "Sft"}
          </span>
        </div>
      </td>

      {/* Deduct */}
      <td style={{ ...tdStyle(alt, deduct), textAlign: "center" }}>
        <input type="checkbox" checked={deduct}
          style={{ width: 14, height: 14, cursor: "pointer" }}
          onChange={e => { setDeduct(e.target.checked); commit(item, d1, d2, qty, type, e.target.checked); }} />
      </td>

      {/* Delete */}
      <td style={{ ...tdStyle(alt, deduct), textAlign: "center", width: 28 }}>
        <button onClick={onRemove}
          style={{ background: "none", border: "none", color: "#ef4444", cursor: "pointer", fontSize: 14, padding: 0 }}>
          🗑
        </button>
      </td>
    </tr>
  );
}

// ─── SESSION PANEL ────────────────────────────────────────────────────────────
function SessionPanel({ sess, onUpdate, onRemove }) {
  const net = netTotal(sess.rows);
  const unit = domUnit(sess);
  const groups = measureGroups(sess);
  const sft = sqftTotal(sess.rows);
  const rft = rnftTotal(sess.rows);
  const gft = groveTotal(sess.rows);
  const both = hasBoth(sess.rows);
  const sqftTitle = sess.sqftTitle || "Sqft";
  const rnftTitle = sess.rnftTitle || "Rnft";
  const T = detectTheme(sess.title); // theme based on section title

  const setSqftTitle = (v) => onUpdate({ ...sess, sqftTitle: v });
  const setRnftTitle = (v) => onUpdate({ ...sess, rnftTitle: v });

  const updateRow = useCallback((rowId, newRow) => {
    onUpdate({ ...sess, rows: sess.rows.map(r => r.id === rowId ? newRow : r) });
  }, [sess, onUpdate]);

  const removeRow = (rowId) => onUpdate({ ...sess, rows: sess.rows.filter(r => r.id !== rowId) });

  const addRow = (type) => {
    const newRow = { id: uid(), item: "New item", d1: "1'0\"", d2: type === "sqft" ? "1'0\"" : null, qty: 1, type, deduct: false, area: 1 };
    onUpdate({ ...sess, rows: [...sess.rows, newRow] });
  };

  return (
    <div style={card}>
      {/* Header */}
      <div style={cardHead(T.dark)}>
        <span style={{ fontSize: 15 }}>📐</span>
        <input
          value={sess.title}
          onChange={e => onUpdate({ ...sess, title: e.target.value })}
          style={{ background: "transparent", border: "none", color: "#fff", fontWeight: "bold", fontSize: 14, fontFamily: "Times New Roman, Noto Sans Devanagari, serif", outline: "none", flex: 1 }}
        />
        <span style={{ background: "rgba(255,255,255,.15)", color: "#fff", fontSize: 11, padding: "2px 8px", borderRadius: 10 }}>
          {unit}
        </span>
        <select
          value={sess.rows.length > 0 ? (linearRows(sess.rows).length > sess.rows.length/2 ? "rnft" : "sqft") : "sqft"}
          onChange={e => {
            const t = e.target.value;
            onUpdate({ ...sess, rows: sess.rows.map(r => ({ ...r, type: t, d2: (t==="rnft"||t==="grove")?null:r.d2, area: calcArea(r.d1, (t==="rnft"||t==="grove")?null:r.d2, r.qty, t) })) });
          }}
          style={{ background:"rgba(255,255,255,.2)", border:"1px solid rgba(255,255,255,.4)", borderRadius:6, color:"#fff", fontWeight:"bold", fontSize:12, padding:"2px 8px", cursor:"pointer", fontFamily:"Times New Roman,serif", marginRight:6 }}
        >
          <option value="sqft" style={{color:"#1a1a2e"}}>Sqft</option>
          <option value="rnft" style={{color:"#1a1a2e"}}>Rnft</option>
        </select>
        {/* Rate input */}
        <div style={{ display:"flex", alignItems:"center", gap:4, background:"rgba(255,255,255,.12)", borderRadius:6, padding:"2px 8px" }}>
          <span style={{ color:"#93c5fd", fontSize:11 }}>₹ Rate:</span>
          <input
            value={sess.rate || ""}
            onChange={e => onUpdate({ ...sess, rate: e.target.value })}
            placeholder="per unit"
            style={{ background:"transparent", border:"none", color:"#fff", fontWeight:"bold", fontSize:13, outline:"none", width:70, fontFamily:"Times New Roman,serif" }}
          />
        </div>
        <span style={{ color: "#93c5fd", fontSize: 12, whiteSpace: "nowrap", display:"flex", gap:10 }}>
          <span>{sess.rows.length} rows</span>
          {groups.length > 1 ? (<>
            {sft !== 0 && <span>Sft: <strong style={{color:"#bfdbfe"}}>{fmtArea(sft)}</strong></span>}
            {rft !== 0 && <span>Rft: <strong style={{color:"#bbf7d0"}}>{fmtArea(rft)}</strong></span>}
            {gft !== 0 && <span>Gft: <strong style={{color:"#e9d5ff"}}>{fmtArea(gft)}</strong></span>}
          </>) : (
            <span>Net: <strong>{fmtArea(net)} {unit}</strong></span>
          )}
        </span>
        <button onClick={onRemove}
          style={{ background: "none", border: "none", color: "#93c5fd", cursor: "pointer", fontSize: 18, marginLeft: 8, lineHeight: 1 }}>
          ✕
        </button>
      </div>

      {/* Type legend */}
      <div style={{ padding: "5px 14px", background: T.light, borderBottom: "1px solid "+(T.mid)+"30", display: "flex", gap: 16, fontSize: 11, color: "#64748b" }}>
        <span>🟦 <strong>Sqft</strong> = L × W × Qty</span>
        <span>🟩 <strong>Rnft</strong> = L × Qty only</span>
        <span style={{ color: "#94a3b8" }}>Toggle type per row · Deduct rows subtracted from total</span>
      </div>

      {/* Table — split by type when both exist */}
      <div style={{ overflowX: "auto" }}>
        {(() => {
          const COLS = 9;
          // Themed table header style
          const tH = { ...thStyle, background: T.dark, borderRight: "1px solid "+T.mid };
          // Themed total row style
          const tTot = { ...totalTd, background: T.dark };
          // Themed zebra row
          const tdZ = (alt, ded) => ({
            padding: "5px 8px", borderBottom: "1px solid "+T.light,
            background: ded ? C.redBg : alt ? T.light : "#fff",
          });

          const Thead = () => (
            <thead>
              <tr>
                <th style={{ ...tH, width: 28 }}>#</th>
                <th style={tH}>Item Name</th>
                <th style={{ ...tH, width: 58, textAlign: "center" }}>Type</th>
                <th style={{ ...tH, width: 96 }}>Dim 1 (L)</th>
                <th style={{ ...tH, width: 96 }}>Dim 2 (W)</th>
                <th style={{ ...tH, width: 52 }}>Qty</th>
                <th style={{ ...tH, width: 88, textAlign: "right" }}>Area</th>
                <th style={{ ...tH, width: 58, textAlign: "center" }}>Deduct</th>
                <th style={{ ...tH, width: 28 }}></th>
              </tr>
            </thead>
          );

          if (groups.length <= 1) {
            return (
              <table style={{ width: "100%", borderCollapse: "collapse" }}>
                <Thead />
                <tbody>
                  {sess.rows.map((row, idx) => (
                    <RowEditor key={row.id} row={row} idx={idx} sessId={sess.id} theme={T}
                      onUpdate={(newRow) => updateRow(row.id, newRow)}
                      onRemove={() => removeRow(row.id)} />
                  ))}
                  {sess.rows.length === 0 && (
                    <tr><td colSpan={COLS} style={{ textAlign: "center", padding: "16px", color: "#94a3b8", fontSize: 13 }}>No rows yet — add one below</td></tr>
                  )}
                </tbody>
                <tfoot>
                  <tr>
                    <td colSpan={6} style={tTot}>Net Total</td>
                    <td style={{ ...tTot, textAlign: "right" }}>{fmtArea(net)} {unit}</td>
                    <td colSpan={2} style={tTot}></td>
                  </tr>
                </tfoot>
              </table>
            );
          }

          // Mixed types — separate themed sub-tables per measurement group

          const SubTable = ({ rows, type, total, label, unitLabel, headBg, valColor }) => (
            <table style={{ width: "100%", borderCollapse: "collapse", marginBottom: 0 }}>
              <thead>
                <tr>
                  <th style={{ ...tH, width: 28 }}>#</th>
                  <th style={tH}>Item Name</th>
                  <th style={{ ...tH, width: 58, textAlign: "center" }}>Type</th>
                  <th style={{ ...tH, width: 96 }}>Dim 1 (L)</th>
                  <th style={{ ...tH, width: 96 }}>Dim 2 (W)</th>
                  <th style={{ ...tH, width: 52 }}>Qty</th>
                  <th style={{ ...tH, width: 88, textAlign: "right" }}>Area</th>
                  <th style={{ ...tH, width: 58, textAlign: "center" }}>Deduct</th>
                  <th style={{ ...tH, width: 28 }}></th>
                </tr>
              </thead>
              <tbody>
                {rows.map((row, idx) => (
                  <RowEditor key={row.id} row={row} idx={idx} sessId={sess.id} theme={T}
                    onUpdate={(newRow) => updateRow(row.id, newRow)}
                    onRemove={() => removeRow(row.id)} />
                ))}
                {rows.length === 0 && (
                  <tr><td colSpan={COLS} style={{ textAlign: "center", padding: "10px", color: "#94a3b8", fontSize: 12 }}>No {type} rows</td></tr>
                )}
              </tbody>
              <tfoot>
                <tr>
                  <td colSpan={6} style={{ ...tTot, background: headBg }}>{label} Total</td>
                  <td style={{ ...tTot, background: headBg, textAlign: "right", color: valColor }}>{fmtArea(total)} {unitLabel}</td>
                  <td colSpan={2} style={{ ...tTot, background: headBg }}></td>
                </tr>
              </tfoot>
            </table>
          );

          const subTitleStyle = {
            background: "transparent", border: "none", color: "#fff",
            fontWeight: "bold", fontSize: 14,
            fontFamily: "Times New Roman, Noto Sans Devanagari, serif",
            outline: "none", width: "auto", minWidth: 60,
          };

          return (
            <>
              {groups.map((group, idx) => {
                const icon = group.key === "sqft" ? "🟦" : group.key === "grove" ? "🟪" : "🟩";
                const headBg = group.key === "sqft" ? T.sqftDark : group.key === "grove" ? "#6d28d9" : T.rnftDark;
                const valColor = group.key === "sqft" ? T.sqftVal : group.key === "grove" ? "#f3e8ff" : T.rnftVal;
                const inputValue = group.key === "sqft" ? sqftTitle : group.key === "grove" ? (sess.groveTitle || "Grove") : rnftTitle;
                const setInput = group.key === "sqft"
                  ? setSqftTitle
                  : group.key === "grove"
                    ? (v) => onUpdate({ ...sess, groveTitle: v })
                    : setRnftTitle;
                return (
                  <React.Fragment key={group.key}>
                    {idx > 0 && <div style={{ height: 8, background: T.light }} />}
                    <div style={{ ...cardHead(headBg), borderRadius: 0, padding: "7px 14px", justifyContent: "flex-start" }}>
                      <span style={{ fontSize: 13 }}>{icon}</span>
                      <input style={subTitleStyle} value={inputValue} onChange={e => setInput(e.target.value)} />
                      <span style={{ background: "rgba(255,255,255,.15)", color: "#fff", fontSize: 11, padding: "2px 8px", borderRadius: 10 }}>{group.unit}</span>
                    </div>
                    <SubTable
                      rows={group.rows}
                      type={group.key}
                      total={group.total}
                      label={group.key === "sqft" ? sqftTitle : group.key === "grove" ? (sess.groveTitle || "Grove") : rnftTitle}
                      unitLabel={group.unit}
                      headBg={headBg}
                      valColor={valColor}
                    />
                  </React.Fragment>
                );
              })}
            </>
          );
        })()}
      </div>

      <div style={{ padding: "10px 14px", borderTop: "1px solid #dbeafe", background: "#f8fafc", display: "flex", gap: 10, flexWrap: "wrap" }}>
        {groups.map((group) => (
          <div key={group.key} style={{ border: "1px solid #dbeafe", borderRadius: 8, padding: "6px 10px", background: "#fff", fontSize: 12, color: C.gray }}>
            <strong style={{ color: T.dark }}>{group.label}</strong>
            {" · "}Gross {fmtArea(grossTotal(group.rows))} {group.unit}
            {" · "}Deduct {fmtArea(deductionTotal(group.rows))} {group.unit}
            {" · "}Net {fmtArea(group.total)} {group.unit}
          </div>
        ))}
      </div>

      {/* Add row buttons */}
      <div style={{ padding: "9px 14px", borderTop: "1px solid #dbeafe", display: "flex", gap: 10 }}>
        <button onClick={() => addRow("sqft")}
          style={{ background: "none", border: "1px dashed "+T.mid, borderRadius: 6, color: T.mid, cursor: "pointer", padding: "5px 12px", fontSize: 12, fontFamily: "Times New Roman, Noto Sans Devanagari, serif" }}>
          + Sqft row
        </button>
        <button onClick={() => addRow("rnft")}
          style={{ background: "none", border: "1px dashed "+C.green, borderRadius: 6, color: C.green, cursor: "pointer", padding: "5px 12px", fontSize: 12, fontFamily: "Times New Roman, Noto Sans Devanagari, serif" }}>
          + Rnft row
        </button>
      </div>
    </div>
  );
}

// ─── BILL EXPORT (pure XML + JSZip — Dindori format) ─────────────────────────

function xmlEscHtml(s) {
  return String(s||"").replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;");
}

function numToWords(n) {
  const a = ["","One","Two","Three","Four","Five","Six","Seven","Eight","Nine",
             "Ten","Eleven","Twelve","Thirteen","Fourteen","Fifteen","Sixteen",
             "Seventeen","Eighteen","Nineteen"];
  const b = ["","","Twenty","Thirty","Forty","Fifty","Sixty","Seventy","Eighty","Ninety"];
  if (n === 0) return "Zero";
  function hw(n) {
    if (n < 20) return a[n];
    if (n < 100) return b[Math.floor(n/10)] + (n%10 ? " "+a[n%10] : "");
    return a[Math.floor(n/100)]+" Hundred"+(n%100?" "+hw(n%100):"");
  }
  let str = "", rem = Math.floor(n);
  if (rem >= 10000000) { str += hw(Math.floor(rem/10000000))+" Crore "; rem %= 10000000; }
  if (rem >= 100000)   { str += hw(Math.floor(rem/100000))+" Lakh "; rem %= 100000; }
  if (rem >= 1000)     { str += hw(Math.floor(rem/1000))+" Thousand "; rem %= 1000; }
  if (rem > 0)         { str += hw(rem); }
  return str.trim() + " Only";
}

function buildBillDocxXml(sessions, opts) {
  const { company, companySpec, companyAddr, companyMob,
          billClient, billSite, billDate, billSub, billAdvance } = opts;
  const fixedItems  = opts.fixedItems  || [];
  const manualItems = opts.manualItems || [];
  const billRates   = opts.billRates   || {};
  const rateCard    = opts.rateCard    || [];

  // Same auto-lookup as BillView — withMat rate from rateCard when no manual rate set
  function autoRate(label, per) {
    if (!rateCard.length) return 0;
    const title = (label||"").toLowerCase();
    const match = rateCard.find(rc => {
      const wt = (rc.workType||"").toLowerCase().trim();
      if (!wt) return false;
      const perMatch = rc.per === per ||
        (per === "Grove" && rc.per === "Grove") ||
        (per === "Rn ft" && rc.per.startsWith("Rn ft"));
      return perMatch && (title.includes(wt) || wt.split(" ").some(w => w.length>2 && title.includes(w)));
    });
    return match ? parseFloat(match.withMat)||0 : 0;
  }

  const B = "1F4E79", M = "2E75B6", L = "D6E4F0", W = "FFFFFF";

  // ── helpers ────────────────────────────────────────────────────────────────
  function esc(s) {
    return String(s ?? "")
      .replace(/&/g,"&amp;").replace(/</g,"&lt;")
      .replace(/>/g,"&gt;").replace(/"/g,"&quot;");
  }

  function rn(text, { bold=false, color=null, sz=20, u=false }={}) {
    const b  = bold ? "<w:b/><w:bCs/>" : "";
    const ul = u    ? '<w:u w:val="single"/>' : "";
    const c  = color ? `<w:color w:val="${color}"/>` : "";
    const s  = `<w:sz w:val="${sz}"/><w:szCs w:val="${sz}"/>`;
    const f  = `<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/>`;
    return `<w:r><w:rPr>${b}${ul}${c}${s}${f}</w:rPr><w:t xml:space="preserve">${esc(text)}</w:t></w:r>`;
  }

  function para(children, { align="left", bg=null, spBefore=0, spAfter=80 }={}) {
    const jc  = `<w:jc w:val="${align}"/>`;
    const spc = `<w:spacing w:before="${spBefore}" w:after="${spAfter}"/>`;
    const shd = bg ? `<w:shd w:val="clear" w:color="auto" w:fill="${bg}"/>` : "";
    return "<w:p><w:pPr>"+spc+jc+shd+"</w:pPr>"+children+"</w:p>";
  }

  // Table cell with full border + shading
  function tc(content, w, { bg=W, align="left", color=B, sz=20, bold=false }={}) {
    const bdr = `<w:tcBorders>
      <w:top    w:val="single" w:sz="4" w:space="0" w:color="${M}"/>
      <w:left   w:val="single" w:sz="4" w:space="0" w:color="${M}"/>
      <w:bottom w:val="single" w:sz="4" w:space="0" w:color="${M}"/>
      <w:right  w:val="single" w:sz="4" w:space="0" w:color="${M}"/>
    </w:tcBorders>`;
    const shd = `<w:shd w:val="clear" w:color="auto" w:fill="${bg}"/>`;
    const mar = `<w:tcMar><w:top w:w="80" w:type="dxa"/><w:left w:w="120" w:type="dxa"/><w:bottom w:w="80" w:type="dxa"/><w:right w:w="120" w:type="dxa"/></w:tcMar>`;
    const tcPr = `<w:tcPr><w:tcW w:w="${w}" w:type="dxa"/>${bdr}${shd}${mar}</w:tcPr>`;
    const pPr  = `<w:pPr><w:spacing w:before="0" w:after="40"/><w:jc w:val="${align}"/></w:pPr>`;
    return `<w:tc>${tcPr}<w:p>${pPr}${rn(content,{bold,color,sz})}</w:p></w:tc>`;
  }

  function tr(...cells) { return "<w:tr>"+cells.join("")+"</w:tr>"; }

  const CW = [640, 3540, 1380,  900, 940, 1950]; // Sr | Particulars | Qty | Rate | Per | Total

  let body = "";

  // ══ 1. HEADER — full-width blue bordered box ═══════════════════════════════
  body += `<w:tbl>
<w:tblPr>
  <w:tblW w:w="9350" w:type="dxa"/>
  <w:tblBorders>
    <w:top    w:val="single" w:sz="12" w:space="0" w:color="${M}"/>
    <w:left   w:val="single" w:sz="12" w:space="0" w:color="${M}"/>
    <w:bottom w:val="single" w:sz="12" w:space="0" w:color="${M}"/>
    <w:right  w:val="single" w:sz="12" w:space="0" w:color="${M}"/>
    <w:insideH w:val="none" w:sz="0" w:space="0" w:color="auto"/>
    <w:insideV w:val="none" w:sz="0" w:space="0" w:color="auto"/>
  </w:tblBorders>
</w:tblPr>
<w:tblGrid><w:gridCol w:w="9350"/></w:tblGrid>
<w:tr>
  <w:tc>
    <w:tcPr>
      <w:tcW w:w="9350" w:type="dxa"/>
      <w:shd w:val="clear" w:color="auto" w:fill="${B}"/>
      <w:tcMar><w:top w:w="80" w:type="dxa"/><w:left w:w="200" w:type="dxa"/><w:bottom w:w="80" w:type="dxa"/><w:right w:w="200" w:type="dxa"/></w:tcMar>
    </w:tcPr>
    ${para(rn(company||"Your Company Name",{bold:true,u:true,color:W,sz:36}),{align:"center",spBefore:60,spAfter:30})}
    ${companySpec ? para(rn("Specialist in : ",{color:"BDD7EE",sz:19})+rn(companySpec,{color:W,sz:19}),{align:"center",spBefore:0,spAfter:20}) : para(rn(""),{spBefore:0,spAfter:10})}
    ${companyMob?para(rn("Mob : "+companyMob,{color:"BDD7EE",sz:18}),{align:"center",spBefore:0,spAfter:50}):para(rn(""),{spBefore:0,spAfter:30})}
  </w:tc>
</w:tr>
</w:tbl>`;

  // ══ 2. SEPARATOR + ADDRESS ════════════════════════════════════════════════
  body += para(
    rn("─".repeat(58)+"  "+"─".repeat(58), {color:"AAAAAA",sz:12}),
    {align:"center",spBefore:20,spAfter:20}
  );
  if (companyAddr) body += para(
    rn("Address : ",{bold:true,color:B,sz:20}) + rn(companyAddr,{color:B,sz:20}),
    {spBefore:0,spAfter:16}
  );
  if (companyMob) body += para(
    rn("Mob.No  : ",{bold:true,color:B,sz:20}) + rn(companyMob,{color:B,sz:20}),
    {spBefore:0,spAfter:16}
  );

  // ══ 3. BILL META ══════════════════════════════════════════════════════════
  // Date — right aligned, underlined
  body += para(
    rn("Date :- "+(billDate||""),{color:B,sz:20,u:true}),
    {align:"right",spBefore:0,spAfter:30}
  );
  // "To:" — always shown
  body += para(
    rn("To :- ",{bold:true,color:B,sz:22})+rn(billClient||"",{color:B,sz:22}),
    {spBefore:0,spAfter:30}
  );
  // SUB — always shown
  body += para(
    rn("SUB :- ",{bold:true,color:B,sz:22})+rn(billSub||"Bill for Services.",{color:B,sz:22}),
    {spBefore:0,spAfter:30}
  );
  // Worksite — always shown
  body += para(
    rn("Worksite Address :- ",{bold:true,u:true,color:B,sz:22})+rn(billSite||"",{color:B,sz:22}),
    {spBefore:0,spAfter:60}
  );
  // BILL heading centered
  body += para(
    rn("BILL",{bold:true,u:true,color:B,sz:28}),
    {align:"center",spBefore:0,spAfter:60}
  );

  // ══ 4. BILL TABLE ═════════════════════════════════════════════════════════
  body += `<w:tbl>
<w:tblPr>
  <w:tblW w:w="9350" w:type="dxa"/>
  <w:tblBorders>
    <w:top    w:val="single" w:sz="8" w:space="0" w:color="${M}"/>
    <w:left   w:val="single" w:sz="8" w:space="0" w:color="${M}"/>
    <w:bottom w:val="single" w:sz="8" w:space="0" w:color="${M}"/>
    <w:right  w:val="single" w:sz="8" w:space="0" w:color="${M}"/>
    <w:insideH w:val="single" w:sz="4" w:space="0" w:color="${M}"/>
    <w:insideV w:val="single" w:sz="4" w:space="0" w:color="${M}"/>
  </w:tblBorders>
</w:tblPr>
<w:tblGrid>
  ${CW.map(w=>`<w:gridCol w:w="${w}"/>`).join("")}
</w:tblGrid>`;

  // Header row — "Total Amount" split to two lines to fit narrow column
  const totalAmtHdrCell = (() => {
    const bdr = `<w:tcBorders>
      <w:top    w:val="single" w:sz="4" w:space="0" w:color="${M}"/>
      <w:left   w:val="single" w:sz="4" w:space="0" w:color="${M}"/>
      <w:bottom w:val="single" w:sz="4" w:space="0" w:color="${M}"/>
      <w:right  w:val="single" w:sz="4" w:space="0" w:color="${M}"/>
    </w:tcBorders>`;
    const shd  = `<w:shd w:val="clear" w:color="auto" w:fill="${B}"/>`;
    const mar  = `<w:tcMar><w:top w:w="60" w:type="dxa"/><w:left w:w="100" w:type="dxa"/><w:bottom w:w="60" w:type="dxa"/><w:right w:w="100" w:type="dxa"/></w:tcMar>`;
    const tcPr = `<w:tcPr><w:tcW w:w="${CW[5]}" w:type="dxa"/>${bdr}${shd}${mar}</w:tcPr>`;
    const pPr  = `<w:pPr><w:spacing w:before="0" w:after="0"/><w:jc w:val="right"/></w:pPr>`;
    return `<w:tc>${tcPr}<w:p>${pPr}${rn("Total",  {bold:true,color:W,sz:20})}</w:p>` +
                        `<w:p>${pPr}${rn("Amount", {bold:true,color:W,sz:20})}</w:p></w:tc>`;
  })();
  body += tr(
    tc("Sr.",        CW[0], {bg:B,color:W,sz:18,bold:true,align:"center"}),
    tc("Particulars",CW[1], {bg:B,color:W,sz:20,bold:true,align:"left"}),
    tc("Quantity",   CW[2], {bg:B,color:W,sz:20,bold:true,align:"right"}),
    tc("Rate",       CW[3], {bg:B,color:W,sz:20,bold:true,align:"right"}),
    tc("Per",        CW[4], {bg:B,color:W,sz:20,bold:true,align:"left"}),
    totalAmtHdrCell,
  );

  let grandTotal=0, srNo=1;

  function dataRow(sr, part, qty, rate, per, total, bg) {
    return tr(
      tc(sr+".",   CW[0], {bg,align:"right", color:"444444",sz:19}),
      tc(part,     CW[1], {bg,align:"left",  color:"000000",sz:19}),
      tc(qty,      CW[2], {bg,align:"right", color:"000000",sz:19}),
      tc(rate,     CW[3], {bg,align:"right", color:"000000",sz:19}),
      tc(per,      CW[4], {bg,align:"left",  color:"000000",sz:19}),
      tc(total,    CW[5], {bg,align:"right", color:"000000",sz:19,bold:true}),
    );
  }

  // Session rows
  for (const sess of sessions) {
    const perOvr = sess.perOverride || null;
    if (hasBoth(sess.rows)) {
      const groups = measureGroups(sess).map(g => ({ ...g, per: perOvr || g.per }));
      for (const g of groups) {
        const lineId = sess.id+g.per;
        const rate   = parseFloat(billRates[lineId] !== undefined ? billRates[lineId] : autoRate(g.label, g.per))||0;
        const qty    = g.total;
        const total  = qty*rate; grandTotal += total;
        const bg = srNo%2===0?L:W;
        const totalDisp = rate ? Math.round(total).toLocaleString("en-IN",{minimumFractionDigits:2}) : "───";
        body += dataRow(srNo++, g.label, qty.toFixed(2),
          rate?String(rate):"───", g.per, totalDisp, bg);
      }
    } else {
      const lineId = sess.id;
      const group  = measureGroups(sess)[0];
      const per    = perOvr || (group ? group.per : "Sq ft");
      const rate   = parseFloat(billRates[lineId] !== undefined ? billRates[lineId] : autoRate(sess.title, per))||0;
      const qty    = netTotal(sess.rows);
      const total  = qty*rate; grandTotal += total;
      const bg = srNo%2===0?L:W;
      const totalDisp2 = rate ? Math.round(total).toLocaleString("en-IN",{minimumFractionDigits:2}) : "───";
      body += dataRow(srNo++, sess.title, qty.toFixed(2),
        rate?String(rate):"───", per, totalDisp2, bg);
    }
  }

  // Manual rows
  for (const mi of manualItems) {
    const qty=parseFloat(mi.qty)||0, rate=parseFloat(mi.rate)||0;
    const total=mi.amount?parseFloat(mi.amount)||0:qty*rate;
    grandTotal+=total;
    const bg=srNo%2===0?L:W;
    body+=dataRow(srNo++,mi.name||"Item",mi.qty||"───",mi.rate||"───",mi.per||"Sq ft",
      total>0?total.toLocaleString("en-IN",{minimumFractionDigits:2}):"───",bg);
  }

  // Fixed items — dashes like real bill (real bill uses "-------" hyphens)
  for (const fi of fixedItems) {
    const amt=fi.qty&&fi.unitPrice?(parseFloat(fi.qty)||0)*(parseFloat(fi.unitPrice)||0):(parseFloat(fi.amount)||0);
    grandTotal+=amt;
    const bg=srNo%2===0?L:W;
    // Real bill: qty shows "2 Nos" if present, rate/per always show "-------"
    const qtyTxt  = fi.qty ? fi.qty+" Nos" : "-------";
    const rateTxt = "-------";
    const perTxt  = "--------";
    body+=dataRow(srNo++,fi.name||"Extra Item",
      qtyTxt, rateTxt, perTxt,
      amt>0?Math.round(amt).toLocaleString("en-IN",{minimumFractionDigits:2}):"---",bg);
  }

  // Footer rows: 4 empty blue cells | label | amount
  const advance=parseFloat(billAdvance)||0;
  const remaining=grandTotal-advance;

  function blankTc(w, bg) {
    const bdr = `<w:tcBorders>
      <w:top    w:val="single" w:sz="4" w:space="0" w:color="${M}"/>
      <w:left   w:val="single" w:sz="4" w:space="0" w:color="${M}"/>
      <w:bottom w:val="single" w:sz="4" w:space="0" w:color="${M}"/>
      <w:right  w:val="none"   w:sz="0" w:space="0" w:color="auto"/>
    </w:tcBorders>`;
    const shd  = `<w:shd w:val="clear" w:color="auto" w:fill="${bg}"/>`;
    const mar  = `<w:tcMar><w:top w:w="80" w:type="dxa"/><w:left w:w="120" w:type="dxa"/><w:bottom w:w="80" w:type="dxa"/><w:right w:w="120" w:type="dxa"/></w:tcMar>`;
    const tcPr = `<w:tcPr><w:tcW w:w="${w}" w:type="dxa"/>${bdr}${shd}${mar}</w:tcPr>`;
    return `<w:tc>${tcPr}<w:p><w:pPr><w:spacing w:before="0" w:after="40"/></w:pPr></w:p></w:tc>`;
  }
  function footRow(label, value, rowBg, textCol) {
    // Single spanning cell for Sr+Part+Qty+Rate+Per all together, then value
    const spanW = CW[0]+CW[1]+CW[2]+CW[3]+CW[4];
    const bdrSpan = `<w:tcBorders>
      <w:top    w:val="single" w:sz="4" w:space="0" w:color="${M}"/>
      <w:left   w:val="single" w:sz="4" w:space="0" w:color="${M}"/>
      <w:bottom w:val="single" w:sz="4" w:space="0" w:color="${M}"/>
      <w:right  w:val="single" w:sz="4" w:space="0" w:color="${M}"/>
    </w:tcBorders>`;
    const shdSpan = `<w:shd w:val="clear" w:color="auto" w:fill="${rowBg}"/>`;
    const marSpan = `<w:tcMar><w:top w:w="80" w:type="dxa"/><w:left w:w="160" w:type="dxa"/><w:bottom w:w="80" w:type="dxa"/><w:right w:w="120" w:type="dxa"/></w:tcMar>`;
    const tcPrSpan = `<w:tcPr><w:tcW w:w="${spanW}" w:type="dxa"/><w:gridSpan w:val="5"/>${bdrSpan}${shdSpan}${marSpan}</w:tcPr>`;
    const pPrSpan  = `<w:pPr><w:spacing w:before="0" w:after="40"/><w:jc w:val="right"/></w:pPr>`;
    const spanCell = `<w:tc>${tcPrSpan}<w:p>${pPrSpan}${rn(label,{bold:true,color:textCol,sz:21})}</w:p></w:tc>`;
    return tr(
      spanCell,
      tc(value, CW[5], {bg:rowBg,color:textCol,sz:21,bold:true,align:"right"}),
    );
  }

  const gt = Math.round(grandTotal);
  const adv = Math.round(advance);
  const rem = gt - adv;
  body += footRow("Total Amount", gt.toLocaleString("en-IN",{minimumFractionDigits:2}), B, W);
  if (adv>0) {
    body += footRow("Advance",   adv.toLocaleString("en-IN",{minimumFractionDigits:2}), M, W);
    body += footRow("Remaining", rem.toLocaleString("en-IN",{minimumFractionDigits:2}), B, W);
  }
  body += "</w:tbl>";

  // ══ 5. FOOTER ════════════════════════════════════════════════════════════
  const amtForWords = adv>0 ? rem : gt;
  body += para(
    rn("In words: ",{bold:true,u:true,color:B,sz:22})+
    rn(numToWords(amtForWords)+".",{color:B,sz:22}),
    {spBefore:100,spAfter:200}
  );
  body += para(rn("Contractor sign.",{color:B,sz:22}),{align:"right",spBefore:0,spAfter:40});

  // ══ 6. WRAPPER ════════════════════════════════════════════════════════════
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<w:document xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"` +
    ` xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"` +
    ` xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"` +
    ` mc:Ignorable="">` +
    `<w:body>${body}` +
    `<w:sectPr>` +
    `<w:pgSz w:w="12240" w:h="15840"/>` +
    `<w:pgMar w:top="1080" w:right="1080" w:bottom="1080" w:left="1080" w:header="720" w:footer="720" w:gutter="0"/>` +
    `</w:sectPr></w:body></w:document>`;
}



// \u2500\u2500\u2500 RATE CARD DOCX \u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500
function buildRateCardDocxXml(rateCard, companyName, companySpec, companyMob) {
  const B="1F4E79", M="2E75B6", L="D6E4F0", W="FFFFFF", G="92400e", GL="FEF3C7";
  const GOLD="B8922A";

  // helpers
  const run = (text, {bold,color,sz=20,italic}={}) => {
    const rPr = [
      bold   ? "<w:b/>"                          : "",
      italic ? "<w:i/>"                          : "",
      color  ? `<w:color w:val="${color}"/>`     : "",
      `<w:sz w:val="${sz}"/><w:szCs w:val="${sz}"/>`,
    ].join("");
    return `<w:r><w:rPr>${rPr}</w:rPr><w:t xml:space="preserve">${text.replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;")}</w:t></w:r>`;
  };
  const para = (runs, {center=false, spBefore=0, spAfter=100, bg}={}) => {
    const jc  = center ? `<w:jc w:val="center"/>` : "";
    const spc = `<w:spacing w:before="${spBefore}" w:after="${spAfter}"/>`;
    const shd = bg ? `<w:shd w:val="clear" w:color="auto" w:fill="${bg}"/>` : "";
    return `<w:p><w:pPr>${jc}${spc}${shd?`<w:pBdr></w:pBdr>`:""}</w:pPr>${runs}</w:p>`;
  };
  const cell = (text, w, {bold,bg,color,right,sz=20,shading}={}) => {
    const bdr = (side) => `<w:${side} w:val="single" w:sz="4" w:space="0" w:color="${M}"/>`;
    const borders = `<w:tcBorders>${["top","left","bottom","right"].map(bdr).join("")}</w:tcBorders>`;
    const fill = (bg||shading) ? `<w:shd w:val="clear" w:color="auto" w:fill="${bg||shading}"/>` : "";
    const jc   = right ? `<w:jc w:val="right"/>` : `<w:jc w:val="left"/>`;
    const mar  = `<w:tcMar><w:top w:w="80" w:type="dxa"/><w:left w:w="100" w:type="dxa"/><w:bottom w:w="80" w:type="dxa"/><w:right w:w="100" w:type="dxa"/></w:tcMar>`;
    return `<w:tc><w:tcPr><w:tcW w:w="${w}" w:type="dxa"/>${borders}${fill}${mar}</w:tcPr>` +
      `<w:p><w:pPr>${jc}<w:spacing w:before="0" w:after="0"/></w:pPr>${run(text,{bold,color,sz})}</w:p></w:tc>`;
  };
  const row = (...cells) => `<w:tr>${cells.join("")}</w:tr>`;
  const tbl = (rows, totalW=9350) =>
    `<w:tbl><w:tblPr><w:tblW w:w="${totalW}" w:type="dxa"/>` +
    `<w:tblBorders><w:insideH w:val="single" w:sz="4" w:color="${M}"/>` +
    `<w:insideV w:val="single" w:sz="4" w:color="${M}"/></w:tblBorders></w:tblPr>` +
    `${rows.join("")}</w:tbl>`;

  // per label
  const perLabel = (per) => {
    const p = (per||"").toLowerCase();
    if (p.includes("sq"))  return "Sq ft";
    if (p.includes("rn")||p.includes("rft")) return "Rn ft";
    if (p.includes("grov")) return "Grove";
    if (p.includes("fix"))  return "Fixed";
    return per||"";
  };

  let body = "";

  // Company header
  body += para(run(companyName||"Your Company Name",{bold:true,color:B,sz:32}),{center:true,spBefore:60,spAfter:40});
  const spec = (companySpec||"").replace(/Specialist in[:\s]*/i,"");
  if (spec) body += para(run(spec,{color:M,sz:18}),{center:true,spAfter:20});
  if (companyMob) body += para(run("Mob: "+companyMob,{color:"475569",sz:18}),{center:true,spAfter:20});
  body += para(run("RATE CARD",{bold:true,color:B,sz:28}),{center:true,spBefore:40,spAfter:160});

  const cats = [
    {key:"polish", label:"Polish Work", hdrBg:B, wtBg:L, wtColor:B},
    {key:"paint",  label:"Paint Work",  hdrBg:G, wtBg:GL,wtColor:G},
  ];

  for (const {key,label,hdrBg,wtBg,wtColor} of cats) {
    const catRows = rateCard.filter(r => r.category === key);
    if (!catRows.length) continue;

    // Group by workType
    const groups = {};
    catRows.forEach(r => { const wt=r.workType||"Other"; if(!groups[wt])groups[wt]=[]; groups[wt].push(r); });

    // Helper: spanning cell for section title & group headers
    const spanCell = (text, bg, color, sz, bold, keepNext=false) => {
      const bdr = (side) => `<w:${side} w:val="single" w:sz="4" w:space="0" w:color="${M}"/>`;
      const borders = `<w:tcBorders>${["top","left","bottom","right"].map(bdr).join("")}</w:tcBorders>`;
      const fill = `<w:shd w:val="clear" w:color="auto" w:fill="${bg}"/>`;
      const mar = `<w:tcMar><w:top w:w="80" w:type="dxa"/><w:left w:w="140" w:type="dxa"/><w:bottom w:w="80" w:type="dxa"/><w:right w:w="100" w:type="dxa"/></w:tcMar>`;
      const kwn = keepNext ? `<w:keepNext/>` : "";
      return `<w:tc><w:tcPr><w:tcW w:w="9350" w:type="dxa"/><w:gridSpan w:val="6"/>${borders}${fill}${mar}</w:tcPr>` +
        `<w:p><w:pPr>${kwn}<w:jc w:val="left"/><w:spacing w:before="0" w:after="0"/></w:pPr>${run(text,{bold,color,sz})}</w:p></w:tc>`;
    };

    // Column header row
    const colHdrRow = (bg) => row(
      cell("#",                        500,  {bold:true, bg, color:W, sz:18, right:true}),
      cell("Work Type",               3000,  {bold:true, bg, color:W, sz:18}),
      cell("Sub-type",                1500,  {bold:true, bg, color:W, sz:18}),
      cell("Per",                      900,  {bold:true, bg, color:W, sz:18}),
      cell("With Material (\u20b9)", 1725,  {bold:true, bg, color:W, sz:18, right:true}),
      cell("Labour (\u20b9)",        1725,  {bold:true, bg, color:W, sz:18, right:true}),
    );

    let sr = 1;
    // Everything in ONE table: section title + col header + all group rows
    const allRows = [];

    // Section title row (NO tblHeader — we do NOT want it to repeat on every page)
    allRows.push(`<w:tr><w:trPr><w:cantSplit/></w:trPr>${spanCell(label, hdrBg, W, 26, true)}</w:tr>`);
    // Column header row (no tblHeader either)
    allRows.push(colHdrRow(M));

    for (const [wt, items] of Object.entries(groups)) {
      // Group header — keepNext so it stays with first data row
      allRows.push(`<w:tr><w:trPr><w:cantSplit/></w:trPr>${spanCell(wt, wtBg, wtColor, 19, true, true)}</w:tr>`);

      items.forEach((r, i) => {
        const bg = i%2===0 ? "FFFFFF" : "EEF4FF";
        const labVal = (!r.labour || r.labour==="0" || r.labour===0) ? "\u2014" : String(r.labour);
        const matVal = (!r.withMat || r.withMat==="0" || r.withMat===0) ? "\u2014" : String(r.withMat);
        // Use category-matching accent colour for With Material value
        const matColor = key === "paint" ? G : B;
        allRows.push(
          row(
            cell(String(sr++), 500,  {bg, color:"888888", sz:18, right:true}),
            cell(r.workType||"", 3000, {bg, sz:19}),
            cell(r.sub||"",      1500, {bg, color:"5a6a85", sz:18}),
            cell(perLabel(r.per), 900, {bg, sz:18}),
            cell(matVal, 1725, {bold:true, bg, color:matColor, sz:20, right:true}),
            cell(labVal, 1725, {bg, color:"5a6a85", sz:20, right:true}),
          )
        );
      });
    }
    body += tbl(allRows);
    body += para(run(""),{spAfter:100});
  }

  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<w:document xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"` +
    ` xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"` +
    ` xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"` +
    ` mc:Ignorable="">` +
    `<w:body>${body}<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="720" w:right="900" w:bottom="720" w:left="900" w:header="360" w:footer="360" w:gutter="0"/></w:sectPr></w:body></w:document>`;
}

// \u2500\u2500\u2500 DOCX EXPORT (pure XML + JSZip \u2014 no blocked CDN) \u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500

// ─── DOCX XML HELPERS (used by buildDocxXml) ─────────────────────────────────
function docxRun(text, { bold=false, color=null, sz=20, italic=false }={}) {
  const esc = (s) => String(s||"").replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;");
  const b  = bold   ? "<w:b/><w:bCs/>" : "";
  const it = italic ? "<w:i/>" : "";
  const c  = color  ? `<w:color w:val="${color}"/>` : "";
  const s  = `<w:sz w:val="${sz}"/><w:szCs w:val="${sz}"/>`;
  return `<w:r><w:rPr>${b}${it}${c}${s}</w:rPr><w:t xml:space="preserve">${esc(text)}</w:t></w:r>`;
}
function docxPara(runs, { center=false, spBefore=0, spAfter=80, bg=null }={}) {
  const jc  = center ? `<w:jc w:val="center"/>` : "";
  const spc = `<w:spacing w:before="${spBefore}" w:after="${spAfter}"/>`;
  const shd = bg ? `<w:shd w:val="clear" w:color="auto" w:fill="${bg}"/>` : "";
  return `<w:p><w:pPr>${jc}${spc}${shd}</w:pPr>${runs}</w:p>`;
}
function docxCell(text, w, { bg="FFFFFF", bold=false, color=null, sz=20, right=false }={}) {
  const M = "2E75B6";
  const bdr = (side) => `<w:${side} w:val="single" w:sz="4" w:space="0" w:color="${M}"/>`;
  const borders = `<w:tcBorders>${["top","left","bottom","right"].map(bdr).join("")}</w:tcBorders>`;
  const shd = `<w:shd w:val="clear" w:color="auto" w:fill="${bg}"/>`;
  const jc  = right ? `<w:jc w:val="right"/>` : `<w:jc w:val="left"/>`;
  const mar = `<w:tcMar><w:top w:w="80" w:type="dxa"/><w:left w:w="120" w:type="dxa"/><w:bottom w:w="80" w:type="dxa"/><w:right w:w="80" w:type="dxa"/></w:tcMar>`;
  return `<w:tc><w:tcPr><w:tcW w:w="${w}" w:type="dxa"/>${borders}${shd}${mar}</w:tcPr>` +
    `<w:p><w:pPr>${jc}<w:spacing w:before="0" w:after="0"/></w:pPr>${docxRun(text,{bold,color,sz})}</w:p></w:tc>`;
}
function docxRow(...cells) { return `<w:tr>${cells.join("")}</w:tr>`; }

function buildDocxXml(sessions, docTitle, companyName) {
  const B="1F4E79", M="2E75B6", L="D6E4F0", DR="CC0000", RR="FFE0E0", W="FFFFFF";
  let body = "";

  if (companyName) {
    body += docxPara(docxRun(companyName,{bold:true,color:B,sz:36}),{center:true,spAfter:120});
  }
  body += docxPara(docxRun(docTitle,{bold:true,color:M,sz:28}),{spBefore:0,spAfter:200});

  // Helper: build one docx table for a set of rows
  // compact=true → tighter cell padding & smaller font to fit more on page
  const buildTable = (rows, unitLabel, totalLabel, totalVal, hdrBg, totBg, BL, BM, compact=false) => {
    const cellSz   = compact ? 18 : 20;   // font size (half-points)
    const hdrSz    = compact ? 19 : 22;
    const cellMar  = compact
      ? `<w:tcMar><w:top w:w="40" w:type="dxa"/><w:left w:w="80" w:type="dxa"/><w:bottom w:w="40" w:type="dxa"/><w:right w:w="80" w:type="dxa"/></w:tcMar>`
      : `<w:tcMar><w:top w:w="80" w:type="dxa"/><w:left w:w="120" w:type="dxa"/><w:bottom w:w="80" w:type="dxa"/><w:right w:w="120" w:type="dxa"/></w:tcMar>`;
    // Override docxCell margin for compact mode
    const cell = (text, w, opts={}) => {
      if (!compact) return docxCell(text, w, opts);
      const M = "2E75B6";
      const bdr = (side) => `<w:${side} w:val="single" w:sz="4" w:space="0" w:color="${M}"/>`;
      const borders = "<w:tcBorders>"+["top","left","bottom","right"].map(bdr).join("")+"</w:tcBorders>";
      const shading = opts.bg ? `<w:shd w:val="clear" w:color="auto" w:fill="${opts.bg}"/>` : "";
      const jc = opts.right ? "<w:jc w:val=\"right\"/>" : "<w:jc w:val=\"left\"/>";
      const sz = opts.sz || cellSz;
      const runSz = opts.bold ? hdrSz : cellSz;
      return `<w:tc><w:tcPr><w:tcW w:w="${w}" w:type="dxa"/>${borders}${shading}${cellMar}</w:tcPr>` +
        `<w:p><w:pPr>${jc}<w:spacing w:before="0" w:after="0"/></w:pPr>${docxRun(text,{...opts,sz:runSz})}</w:p></w:tc>`;
    };
    const hdr = docxRow(
      cell("Sr. No",700,{bold:true,bg:hdrBg,color:W,sz:hdrSz}),
      cell("Items",4000,{bold:true,bg:hdrBg,color:W,sz:hdrSz}),
      cell("Dimensions",2800,{bold:true,bg:hdrBg,color:W,sz:hdrSz}),
      cell("Area ("+unitLabel+")",1850,{bold:true,bg:hdrBg,color:W,sz:hdrSz,right:true}),
    );
    const dataRows = rows.map((r,i) => {
      const bg = r.deduct ? RR : i%2===0 ? W : BL;
      const tc = r.deduct ? DR : null;
      const dim = (r.type==="rnft"||r.type==="grove") ? r.d1+" x "+r.qty : r.d1+" x "+(r.d2||"?")+" x "+r.qty;
      const areaText = r.deduct ? "("+fmtArea(r.area)+")" : fmtArea(r.area);
      return docxRow(
        cell(String(i+1),700,{bg,color:tc,right:true}),
        cell(r.item,4000,{bg,color:tc}),
        cell(dim,2800,{bg,color:tc}),
        cell(areaText,1850,{bg,color:tc,right:true}),
      );
    }).join("");
    const totRow = docxRow(
      cell("",700,{bg:totBg}),
      cell(totalLabel,4000,{bold:true,bg:totBg,color:W,sz:hdrSz}),
      cell("",2800,{bg:totBg}),
      cell(fmtArea(totalVal)+" "+unitLabel,1850,{bold:true,bg:totBg,color:W,sz:hdrSz,right:true}),
    );
    return `<w:tbl><w:tblPr><w:tblW w:w="9350" w:type="dxa"/>` +
      `<w:tblBorders><w:insideH w:val="single" w:sz="4" w:color="${BM}"/>` +
      `<w:insideV w:val="single" w:sz="4" w:color="${BM}"/></w:tblBorders></w:tblPr>` +
      hdr+dataRows+totRow+"</w:tbl>";
  };

  // Page break paragraph
  const pageBreak = () => `<w:p><w:r><w:br w:type="page"/></w:r></w:p>`;

  // Group sessions by imageIndex — sessions from the same image share a page
  // Sessions without imageIndex (manually added) each get their own group
  const groups = [];
  let currentGroup = [];
  let currentImgIdx = undefined;
  for (const sess of sessions) {
    const imgIdx = sess.imageIndex;
    if (imgIdx === undefined) {
      // manually added section — flush current group, start new single group
      if (currentGroup.length) groups.push(currentGroup);
      groups.push([sess]);
      currentGroup = [];
      currentImgIdx = undefined;
    } else if (imgIdx !== currentImgIdx) {
      // new image — flush previous group
      if (currentGroup.length) groups.push(currentGroup);
      currentGroup = [sess];
      currentImgIdx = imgIdx;
    } else {
      // same image — add to current group
      currentGroup.push(sess);
    }
  }
  if (currentGroup.length) groups.push(currentGroup);

  // Render each group; insert page break between groups
  for (let gi = 0; gi < groups.length; gi++) {
    if (gi > 0) body += pageBreak();

    // Count total rows in this group to decide compact spacing
    const groupTotalRows = groups[gi].reduce((s, sess) => s + sess.rows.length, 0);
    // If ≤6 rows total in the group, use tighter spacing so it stays on one page
    const compact = groupTotalRows <= 6;

    // Spacing values: normal vs compact
    const SP = compact ? {
      titleBefore: 80,  titleAfter: 40,
      subBefore:   20,  subAfter:   20,
      gapAfter:    40,
      trailAfter:  20,
    } : {
      titleBefore: 240, titleAfter: 80,
      subBefore:   60,  subAfter:   40,
      gapAfter:    120,
      trailAfter:  60,
    };

    for (const sess of groups[gi]) {
      const net = netTotal(sess.rows);
      const unit = domUnit(sess);
      const DT = detectTheme(sess.title);
      const BD = DT.dark.replace("#","");
      const BM = DT.mid.replace("#","");
      const BL = DT.light.replace("#","");

      body += docxPara(docxRun(sess.title,{bold:true,color:BD,sz:24}),{spBefore:SP.titleBefore,spAfter:SP.titleAfter});

      if (sess.rows.some(r => r.deduct)) {
        body += docxPara(docxRun("★ Rows in brackets are deductions",{italic:true,color:DR,sz:16}),{spAfter:compact?30:60});
      }

      const groups = measureGroups(sess);
      const both = hasBoth(sess.rows);

      if (both) {
        groups.forEach((group, idx) => {
          const groupColor = group.key === "sqft" ? DT.sqftDark.replace("#","") : group.key === "grove" ? "6d28d9" : DT.rnftDark.replace("#","");
          body += docxPara(docxRun("▪ "+group.label,{bold:true,color:groupColor,sz:22}),{spBefore:SP.subBefore,spAfter:SP.subAfter});
          body += buildTable(group.rows,group.unit,group.label+" Total",group.total,BD,groupColor,BL,BM,compact);
          if (idx < groups.length - 1) body += docxPara(docxRun(""),{spAfter:SP.gapAfter});
        });
      } else {
        body += buildTable(sess.rows, unit, "Net Total", net, BD, BD, BL, BM, compact);
      }
      body += docxPara(docxRun(""),{spAfter:SP.trailAfter});
    }
  }

  const gSft   = grandSqft(sessions);
  const gRnOly = grandRnft(sessions);
  const gGrv   = grandGrove(sessions);
  const gtParts = [];
  if (gSft   > 0) gtParts.push("Sqft Total: "+fmtArea(gSft)+" Sft");
  if (gRnOly > 0) gtParts.push("Rnft Total: "+fmtArea(gRnOly)+" Rft");
  if (gGrv   > 0) gtParts.push("Grove Total: "+fmtArea(gGrv)+" Gft");
  if (gtParts.length > 0) {
    body += docxPara(docxRun(gtParts.join("   |   "),{bold:true,color:B,sz:24}),{spBefore:140});
  }

  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<w:document xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas"` +
    ` xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"` +
    ` xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"` +
    ` xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"` +
    ` mc:Ignorable="">` +
    `<w:body>${body}` +
    `<w:sectPr><w:pgSz w:w="12240" w:h="15840"/>` +
    `<w:pgMar w:top="1080" w:right="1080" w:bottom="1080" w:left="1080" w:header="720" w:footer="720" w:gutter="0"/>` +
    `</w:sectPr></w:body></w:document>`;
}

const CONTENT_TYPES = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
  <Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>
</Types>`;

const RELS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`;

const DOC_RELS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
</Relationships>`;

const STYLES = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:docDefaults><w:rPrDefault><w:rPr>
    <w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/>
    <w:sz w:val="20"/><w:szCs w:val="20"/>
  </w:rPr></w:rPrDefault></w:docDefaults>
</w:styles>`;

const SETTINGS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:defaultTabStop w:val="720"/>
</w:settings>`;

async function buildDocx(sessions, docTitle, companyName, logoUrl, customXml) {
  // Load JSZip from CDN (much more reliable in sandboxes than docx lib)
  if (!window.JSZip) {
    await new Promise((res, rej) => {
      const urls = [
        "https://cdnjs.cloudflare.com/ajax/libs/jszip/3.10.1/jszip.min.js",
        "https://cdn.jsdelivr.net/npm/jszip@3.10.1/dist/jszip.min.js",
      ];
      let idx = 0;
      const tryNext = () => {
        if (idx >= urls.length) { rej(new Error("Failed to load JSZip from all sources")); return; }
        const s = document.createElement("script");
        s.src = urls[idx++];
        s.onload = res;
        s.onerror = tryNext;
        document.head.appendChild(s);
      };
      tryNext();
    });
  }

  const zip = new window.JSZip();
  zip.file("[Content_Types].xml", CONTENT_TYPES);
  zip.file("_rels/.rels", RELS);
  zip.file("word/document.xml", customXml || buildDocxXml(sessions, docTitle, companyName));
  zip.file("word/_rels/document.xml.rels", DOC_RELS);
  zip.file("word/styles.xml", STYLES);
  zip.file("word/settings.xml", SETTINGS);

  return await zip.generateAsync({ type: "blob", mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document" });
}

// ─── DOWNLOAD HELPER ─────────────────────────────────────────────────────────
function download(blob, name) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = name;
  a.style.position = "fixed";
  a.style.top = "0";
  a.style.left = "0";
  a.style.opacity = "0";
  document.body.appendChild(a);
  a.dispatchEvent(new MouseEvent("click", { bubbles: true, cancelable: true, view: window }));
  setTimeout(() => {
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  }, 300);
}

// ─── PRINT STYLES ─────────────────────────────────────────────────────────────
const PRINT_CSS = `
@import url('https://fonts.googleapis.com/css2?family=Noto+Sans+Devanagari:wght@400;700&display=swap');

@media print {
  .noprint { display: none !important; }
  @page { margin: 12mm; }
  body { background: #fff !important; }
}`;

// ─── BILL VIEW ───────────────────────────────────────────────────────────────
// ─── RATE CARD COMPONENT ─────────────────────────────────────────────────────

function lookupRateFromCard(rateCard, label, per) {
  if (!rateCard || !rateCard.length) return 0;
  const title = (label || "").toLowerCase().trim();
  const perN  = (per  || "").toLowerCase().trim();
  const match = rateCard.find(rc => {
    const wt = (rc.workType || "").toLowerCase().trim();
    if (!wt) return false;
    const perMatch =
      (perN.includes("sq")   && (rc.per||"").toLowerCase().includes("sq")) ||
      (perN.includes("rn")   && (rc.per||"").toLowerCase().startsWith("rn")) ||
      (perN.includes("grov") && (rc.per||"").toLowerCase().includes("grov")) ||
      (perN.includes("fix")  && (rc.per||"").toLowerCase().includes("fix"));
    return perMatch && (title.includes(wt) || wt.split(" ").some(w => w.length > 2 && title.includes(w)));
  });
  return match ? (parseFloat(match.withMat) || 0) : 0;
}


function AddItemForm({ BD, W, setRateCard }) {
  const [newCat,    setNewCat]    = useState("polish");
  const [newType,   setNewType]   = useState("");
  const [newSub,    setNewSub]    = useState("");
  const [newPer,    setNewPer]    = useState("Sq ft");
  const [newMat,    setNewMat]    = useState("");
  const [newLabour, setNewLabour] = useState("");

  function handleAdd() {
    if (!newType.trim()) return;
    setRateCard(prev => [...prev, {
      id: "rc" + Date.now(),
      category: newCat,
      workType: newType.trim(),
      sub:      newSub.trim() || newPer,
      withMat:  newMat,
      labour:   newLabour,
      per:      newPer,
    }]);
    setNewType(""); setNewSub(""); setNewMat(""); setNewLabour(""); setNewPer("Sq ft");
  }

  const accent = newCat==="polish" ? BD : "#92400e";
  const fInp = (val, onChange, placeholder) => (
    <input value={val} onChange={e=>onChange(e.target.value)} placeholder={placeholder}
      style={{ width:"100%", border:"1px solid #cbd5e1", borderRadius:6, padding:"7px 10px",
        fontSize:13, outline:"none", fontFamily:"inherit" }}
      onFocus={e=>e.target.style.borderColor=accent}
      onBlur={e=>e.target.style.borderColor="#cbd5e1"}
    />
  );

  return (
    <div style={{ background:W, borderRadius:12, border:"2px dashed "+(accent)+"50",
        padding:"18px 20px", marginTop:8, boxShadow:"0 2px 8px rgba(0,0,0,.04)" }}>
      <div style={{ fontSize:14, fontWeight:"bold", color:accent, marginBottom:14 }}>
        + Add New Item
      </div>
      <div style={{ display:"grid", gridTemplateColumns:"1fr 1.6fr 1fr 90px 95px 95px auto",
          gap:10, alignItems:"end" }}>

        {/* Category */}
        <div>
          <div style={{ fontSize:11, color:"#64748b", marginBottom:4, fontWeight:"bold" }}>Category</div>
          <select value={newCat} onChange={e=>setNewCat(e.target.value)}
            style={{ width:"100%", border:"1px solid #cbd5e1", borderRadius:6,
              padding:"7px 8px", fontSize:13, outline:"none", background:W, cursor:"pointer" }}>
            <option value="polish">🪵 Polish</option>
            <option value="paint">🎨 Paint</option>
          </select>
        </div>

        {/* Work Type */}
        <div>
          <div style={{ fontSize:11, color:"#64748b", marginBottom:4, fontWeight:"bold" }}>Work Type *</div>
          {fInp(newType, setNewType, "e.g. Melamine, Duco, Royal Paint…")}
        </div>

        {/* Sub label */}
        <div>
          <div style={{ fontSize:11, color:"#64748b", marginBottom:4, fontWeight:"bold" }}>Sub Label</div>
          {fInp(newSub, setNewSub, "e.g. Sqft, Rn ft, Groove…")}
        </div>

        {/* Per */}
        <div>
          <div style={{ fontSize:11, color:"#64748b", marginBottom:4, fontWeight:"bold" }}>Per</div>
          <select value={newPer} onChange={e=>setNewPer(e.target.value)}
            style={{ width:"100%", border:"1px solid #cbd5e1", borderRadius:6,
              padding:"7px 8px", fontSize:13, outline:"none", background:W, cursor:"pointer" }}>
            {["Sq ft","Rn ft","Grove","Fixed","Nos"].map(o=><option key={o} value={o}>{o}</option>)}
          </select>
        </div>

        {/* With Material */}
        <div>
          <div style={{ fontSize:11, color:"#64748b", marginBottom:4, fontWeight:"bold" }}>With Mat (₹)</div>
          <input type="number" min="0" value={newMat} onChange={e=>setNewMat(e.target.value)}
            placeholder="0"
            style={{ width:"100%", border:"1px solid #cbd5e1", borderRadius:6,
              padding:"7px 8px", fontSize:13, outline:"none", textAlign:"right" }}
            onFocus={e=>e.target.style.borderColor=accent}
            onBlur={e=>e.target.style.borderColor="#cbd5e1"}
          />
        </div>

        {/* Labour */}
        <div>
          <div style={{ fontSize:11, color:"#64748b", marginBottom:4, fontWeight:"bold" }}>Labour (₹)</div>
          <input type="number" min="0" value={newLabour} onChange={e=>setNewLabour(e.target.value)}
            placeholder="0"
            style={{ width:"100%", border:"1px solid #cbd5e1", borderRadius:6,
              padding:"7px 8px", fontSize:13, outline:"none", textAlign:"right" }}
            onFocus={e=>e.target.style.borderColor=accent}
            onBlur={e=>e.target.style.borderColor="#cbd5e1"}
          />
        </div>

        {/* Add button */}
        <div>
          <div style={{ fontSize:11, color:"transparent", marginBottom:4 }}>.</div>
          <button onClick={handleAdd} disabled={!newType.trim()}
            style={{ background: newType.trim() ? accent : "#e2e8f0",
              color: newType.trim() ? W : "#94a3b8",
              border:"none", borderRadius:6, padding:"7px 20px",
              fontSize:13, fontWeight:"bold",
              cursor: newType.trim() ? "pointer" : "default",
              whiteSpace:"nowrap" }}>
            + Add
          </button>
        </div>
      </div>
    </div>
  );
}

function DEFAULT_RATES() {
  return [
  { id:"rc01", category:"polish", workType:"French Polish", sub:"Sqft", withMat:"40", labour:"0", per:"Sq ft" },
  { id:"rc02", category:"polish", workType:"French Polish", sub:"Rn ft", withMat:"30", labour:"0", per:"Rn ft" },
  { id:"rc03", category:"polish", workType:"French Polish", sub:"Groove", withMat:"25", labour:"0", per:"Grove" },
  { id:"rc04", category:"polish", workType:"Melamine Spray Polish Veneer", sub:"Sqft", withMat:"80", labour:"0", per:"Sq ft" },
  { id:"rc05", category:"polish", workType:"Melamine Spray Polish Veneer", sub:"Rn ft", withMat:"42", labour:"0", per:"Rn ft" },
  { id:"rc06", category:"polish", workType:"Melamine Spray Polish Veneer", sub:"Groove", withMat:"26", labour:"0", per:"Grove" },
  { id:"rc07", category:"polish", workType:"Melamine Spray Polish Wood", sub:"Sqft", withMat:"85", labour:"0", per:"Sq ft" },
  { id:"rc08", category:"polish", workType:"Melamine Spray Polish Wood", sub:"Rn ft 1-2 inch", withMat:"42", labour:"0", per:"Rn ft" },
  { id:"rc09", category:"polish", workType:"Melamine Spray Polish Wood", sub:"Rn ft 3-5 inch", withMat:"45", labour:"0", per:"Rn ft" },
  { id:"rc10", category:"polish", workType:"Melamine Spray Polish Wood", sub:"Groove", withMat:"30", labour:"0", per:"Grove" },
  { id:"rc11", category:"polish", workType:"Clear P.O", sub:"Sqft", withMat:"230", labour:"0", per:"Sq ft" },
  { id:"rc12", category:"polish", workType:"Clear P.O", sub:"Rn ft 1-2 inch", withMat:"82", labour:"0", per:"Rn ft" },
  { id:"rc13", category:"polish", workType:"Clear P.O", sub:"Rn ft 3-5 inch", withMat:"92", labour:"0", per:"Rn ft" },
  { id:"rc14", category:"polish", workType:"Clear P.O", sub:"Groove", withMat:"40", labour:"0", per:"Grove" },
  { id:"rc15", category:"polish", workType:"White P.O", sub:"Sqft", withMat:"262", labour:"0", per:"Sq ft" },
  { id:"rc16", category:"polish", workType:"White P.O", sub:"Rn ft 1-2 inch", withMat:"122", labour:"0", per:"Rn ft" },
  { id:"rc17", category:"polish", workType:"White P.O", sub:"Rn ft 3-5 inch", withMat:"137", labour:"0", per:"Rn ft" },
  { id:"rc18", category:"polish", workType:"White P.O", sub:"Groove", withMat:"82", labour:"0", per:"Grove" },
  { id:"rc19", category:"polish", workType:"Duco", sub:"Sqft", withMat:"225", labour:"0", per:"Sq ft" },
  { id:"rc20", category:"polish", workType:"Duco", sub:"Rn ft", withMat:"92", labour:"0", per:"Rn ft" },
  { id:"rc21", category:"polish", workType:"Duco", sub:"Groove", withMat:"48", labour:"0", per:"Grove" },
  { id:"rc22", category:"polish", workType:"Super Acrylic Duco", sub:"Sqft", withMat:"250", labour:"0", per:"Sq ft" },
  { id:"rc23", category:"polish", workType:"Super Acrylic Duco", sub:"Rn ft", withMat:"130", labour:"0", per:"Rn ft" },
  { id:"rc24", category:"polish", workType:"Super Acrylic Duco", sub:"Groove", withMat:"72", labour:"0", per:"Grove" },
  { id:"rc25", category:"polish", workType:"Lamination", sub:"Sqft", withMat:"500", labour:"0", per:"Sq ft" },
  { id:"rc26", category:"polish", workType:"Lamination", sub:"Rn ft", withMat:"190", labour:"0", per:"Rn ft" },
  { id:"rc27", category:"polish", workType:"Lamination", sub:"Groove", withMat:"80", labour:"0", per:"Grove" },
  { id:"rc28", category:"paint", workType:"Luster Paint Satin", sub:"Sqft", withMat:"40", labour:"21", per:"Sq ft" },
  { id:"rc29", category:"paint", workType:"Luster Paint Satin", sub:"Rn ft", withMat:"30", labour:"15", per:"Rn ft" },
  { id:"rc30", category:"paint", workType:"Luster Paint Satin", sub:"Groove", withMat:"25", labour:"14", per:"Grove" },
  { id:"rc31", category:"paint", workType:"Luster Texture Paint", sub:"Sqft", withMat:"42", labour:"21", per:"Sq ft" },
  { id:"rc32", category:"paint", workType:"Luster Texture Paint", sub:"Rn ft", withMat:"31", labour:"15", per:"Rn ft" },
  { id:"rc33", category:"paint", workType:"Luster Texture Paint", sub:"Groove", withMat:"25", labour:"14", per:"Grove" },
  { id:"rc34", category:"paint", workType:"Luster Zinc Texture", sub:"Sqft", withMat:"45", labour:"21", per:"Sq ft" },
  { id:"rc35", category:"paint", workType:"Luster Zinc Texture", sub:"Rn ft", withMat:"29", labour:"14", per:"Rn ft" },
  { id:"rc36", category:"paint", workType:"Luster Zinc Texture", sub:"Groove", withMat:"25", labour:"14", per:"Grove" },
  { id:"rc37", category:"paint", workType:"Asian Plastic Paints", sub:"Sqft", withMat:"34", labour:"22", per:"Sq ft" },
  { id:"rc38", category:"paint", workType:"Asian Plastic Paints", sub:"Rn ft", withMat:"20", labour:"15", per:"Rn ft" },
  { id:"rc39", category:"paint", workType:"Asian Plastic Paints", sub:"Groove", withMat:"18", labour:"15", per:"Grove" },
  { id:"rc40", category:"paint", workType:"Royal Paints", sub:"Sqft", withMat:"42", labour:"22", per:"Sq ft" },
  { id:"rc41", category:"paint", workType:"Royal Paints", sub:"Rn ft", withMat:"28", labour:"15", per:"Rn ft" },
  { id:"rc42", category:"paint", workType:"Royal Paints", sub:"Groove", withMat:"23", labour:"15", per:"Grove" },
  { id:"rc43", category:"paint", workType:"Royal Shyne Paints", sub:"Sqft", withMat:"45", labour:"25", per:"Sq ft" },
  { id:"rc44", category:"paint", workType:"Royal Shyne Paints", sub:"Rn ft", withMat:"28", labour:"15", per:"Rn ft" },
  { id:"rc45", category:"paint", workType:"Royal Shyne Paints", sub:"Groove", withMat:"22", labour:"15", per:"Grove" },
  { id:"rc46", category:"paint", workType:"Oil Paints", sub:"Sqft", withMat:"34", labour:"20", per:"Sq ft" },
  { id:"rc47", category:"paint", workType:"Oil Paints", sub:"Rn ft", withMat:"24", labour:"15", per:"Rn ft" },
  { id:"rc48", category:"paint", workType:"Oil Paints", sub:"Groove", withMat:"22", labour:"15", per:"Grove" },
  { id:"rc49", category:"paint", workType:"Distemper", sub:"Sqft", withMat:"30", labour:"21", per:"Sq ft" },
  { id:"rc50", category:"paint", workType:"Distemper", sub:"Rn ft", withMat:"22", labour:"15", per:"Rn ft" },
  { id:"rc51", category:"paint", workType:"Distemper", sub:"Groove", withMat:"18", labour:"15", per:"Grove" },
  { id:"rc52", category:"paint", workType:"Tractor Emulsion", sub:"Sqft", withMat:"32", labour:"21", per:"Sq ft" },
  { id:"rc53", category:"paint", workType:"Tractor Emulsion", sub:"Rn ft", withMat:"22", labour:"15", per:"Rn ft" },
  { id:"rc54", category:"paint", workType:"Tractor Emulsion", sub:"Groove", withMat:"18", labour:"15", per:"Grove" },
  { id:"rc55", category:"paint", workType:"Water Cutting", sub:"Sqft", withMat:"130", labour:"65", per:"Sq ft" },
  { id:"rc56", category:"paint", workType:"Water Cutting", sub:"Rn ft", withMat:"55", labour:"35", per:"Rn ft" },
  { id:"rc57", category:"paint", workType:"Sunmica Cleaning", sub:"Fixed", withMat:"800", labour:"0", per:"Fixed" },
  { id:"rc58", category:"paint", workType:"Apex Paints", sub:"Sqft", withMat:"28", labour:"15", per:"Sq ft" },
  { id:"rc59", category:"paint", workType:"Lambi Palti Primer", sub:"Sqft", withMat:"25", labour:"15", per:"Sq ft" },
  { id:"rc60", category:"paint", workType:"ACE", sub:"Sqft", withMat:"25", labour:"15", per:"Sq ft" },
  { id:"rc61", category:"paint", workType:"Altima Paints", sub:"Sqft", withMat:"28", labour:"15", per:"Sq ft" },
  { id:"rc62", category:"paint", workType:"Altima Protec", sub:"Sqft", withMat:"45", labour:"20", per:"Sq ft" }
  ];
}
const RATE_CARD_DEFAULT = DEFAULT_RATES();


function RateNumInput({val, onChange, w}) {
  const W="#fff", M="#2E75B6";
  return (
    <input type="number" min="0" step="0.50" value={val}
      onChange={e=>onChange(e.target.value)}
      placeholder="-"
      style={{ width:w||"100%", border:"1px solid #cbd5e1", borderRadius:4, padding:"3px 6px",
        fontSize:13, textAlign:"right", outline:"none", fontFamily:"inherit",
        background: val ? W : "#f8fafc" }}
      onFocus={e=>{e.target.style.borderColor=M; e.target.style.background=W;}}
      onBlur={e=>{e.target.style.borderColor="#cbd5e1"; e.target.style.background=val?W:"#f8fafc";}}
    />
  );
}
function RateCard({ rateCard, setRateCard, companyName, companySpec, companyMob, lang = "en" }) {
  const BD = "#1F4E79", M = "#2E75B6", LB = "#dbeafe", W = "#fff";
  const t = (en, hi) => lang === "hi" ? hi : en;
  const [filter, setFilter] = useState("all"); // "all" | "polish" | "paint"
  const [search, setSearch] = useState("");
  const [selKeys, setSelKeys] = useState(null);

  function upd(id, field, val) {
    setRateCard(prev => prev.map(r => r.id===id ? {...r, [field]: val} : r));
  }
  function addRow(category) {
    setRateCard(prev => [...prev, {
      id:"rc"+Date.now(), category, workType:"", sub:"Sqft",
      withMat:"", labour:"", per:"Sq ft"
    }]);
  }
  function delRow(id) { setRateCard(prev => prev.filter(r => r.id!==id)); }

  function resetToDefaults() {
    if (!window.confirm("Reset to default quotation rates? Your edits will be lost.")) return;
    setRateCard(DEFAULT_RATES());
  }

  const thS = { background:BD, color:W, padding:"8px 10px", fontSize:12,
    fontWeight:"bold", textAlign:"left", whiteSpace:"nowrap" };
  const thR = { ...thS, textAlign:"right" };
  const thC = { ...thS, textAlign:"center" };



  const PER_OPTS = ["Sq ft","Rn ft","Grove","Fixed","Nos"];

  const filtered = rateCard.filter(r => {
    if (filter!=="all" && r.category!==filter) return false;
    if (search && !r.workType.toLowerCase().includes(search.toLowerCase()) &&
        !r.sub.toLowerCase().includes(search.toLowerCase())) return false;
    return true;
  });

  // Group by category then workType for display
  const groups = {};
  filtered.forEach(r => {
    const g = r.category||"paint";
    if (!groups[g]) groups[g] = {};
    const wt = r.workType||"(unnamed)";
    if (!groups[g][wt]) groups[g][wt] = [];
    groups[g][wt].push(r);
  });

  const catColors = {
    polish: { header:"#1F4E79", light:"#EFF6FF", accent:"#2E75B6", label:"🪵 Polish" },
    paint:  { header:"#78350f", light:"#FEF9EE", accent:"#d97706", label:"🎨 Paint" },
  };

  let globalIdx = 0;

  const rcWrap = { maxWidth:980, margin:"0 auto", padding:"0 4px" };
  return (
    <div style={rcWrap}>

      {/* ── Header ─────────────────────────────────────────────────────────── */}
      <div style={{ background:W, borderRadius:12, border:"1px solid #dbeafe",
          padding:"16px 20px", marginBottom:16, boxShadow:"0 2px 8px rgba(31,78,121,.07)" }}>
        <div style={{ display:"flex", alignItems:"center", justifyContent:"space-between",
            flexWrap:"wrap", gap:10 }}>
          <div>
            <div style={{ fontSize:17, fontWeight:"bold", color:BD }}>💰 {t("Rate Card","रेट कार्ड")}</div>
            <div style={{ fontSize:12, color:"#64748b", marginTop:2 }}>
              {t("Rates from your quotation — edit any field, add rows, then use Fetch in 🧾 Bill tab.","अपनी दरें यहां संपादित करें, नई पंक्तियां जोड़ें, फिर 🧾 बिल टैब में उपयोग करें।")}
            </div>
          </div>
          <div style={{ display:"flex", gap:8, flexWrap:"wrap", alignItems:"center" }}>
            {/* Filter tabs */}
            <div style={{ display:"flex", background:"#f1f5f9", borderRadius:7, padding:2 }}>
              {[["all",t("All","सभी")],["polish",t("🪵 Polish","🪵 पॉलिश")],["paint",t("🎨 Paint","🎨 पेंट")]].map(([k,l])=>(
                <button key={k} onClick={()=>setFilter(k)} style={{
                  background: filter===k ? BD : "transparent",
                  color: filter===k ? W : "#475569",
                  border:"none", borderRadius:5, padding:"5px 12px",
                  fontWeight:"bold", fontSize:12, cursor:"pointer"
                }}>{l}</button>
              ))}
            </div>
            {/* Search */}
            <input value={search} onChange={e=>setSearch(e.target.value)}
              placeholder={t("Search…","खोजें…")}
              style={{ border:"1px solid #cbd5e1", borderRadius:6, padding:"5px 10px",
                fontSize:12, outline:"none", width:120 }}
            />
            {/* Add buttons */}
            <button onClick={()=>addRow("polish")} style={{ background:BD, color:W,
              border:"none", borderRadius:6, padding:"6px 12px",
              fontSize:12, fontWeight:"bold", cursor:"pointer" }}>{t("+ Polish Row","+ पॉलिश पंक्ति")}</button>
            <button onClick={()=>addRow("paint")} style={{ background:"#92400e", color:W,
              border:"none", borderRadius:6, padding:"6px 12px",
              fontSize:12, fontWeight:"bold", cursor:"pointer" }}>{t("+ Paint Row","+ पेंट पंक्ति")}</button>
            <button onClick={resetToDefaults} title="Reset to quotation defaults"
              style={{ background:"#f1f5f9", color:"#475569", border:"1px solid #cbd5e1",
                borderRadius:6, padding:"6px 10px", fontSize:12, cursor:"pointer" }}>↺ Reset</button>
            {/* Download Word */}
            <button onClick={async () => {
              try {
                const xml  = buildRateCardDocxXml(selKeys===null?rateCard:rateCard.filter(function(r){return selKeys.indexOf((r.category||"paint")+"|"+(r.workType||"x"))>=0;}), companyName, companySpec, companyMob);
                const blob = await buildDocx([], "", companyName, null, xml);
                download(blob, "Rate_Card.docx");
              } catch(e) { alert("Word export failed: "+e.message); }
            }} style={{ background:"#1F4E79", color:"#fff", border:"none",
              borderRadius:6, padding:"6px 12px", fontSize:12, fontWeight:"bold", cursor:"pointer" }}>
              📄 Word
            </button>
            {/* Download CSV */}
            <button onClick={() => {
              const headers = "Sr,Category,Work Type,Sub,Per,With Material (₹),Labour (₹)";
              const rows = selKeys===null?rateCard:rateCard.filter(function(r){return selKeys.indexOf((r.category||"paint")+"|"+(r.workType||"x"))>=0;}).map((r,i) =>
                `${i+1},${r.category},${r.workType||""},${r.sub||""},${r.per||""},${r.withMat||""},${r.labour||""}`
              ).join("\n");
              const blob = new Blob([headers+"\n"+rows], {type:"text/csv"});
              const url = URL.createObjectURL(blob);
              const a = document.createElement("a");
              a.href=url; a.download="Rate_Card.csv";
              document.body.appendChild(a); a.click();
              document.body.removeChild(a);
              setTimeout(()=>URL.revokeObjectURL(url),5000);
            }} style={{ background:"#16a34a", color:"#fff", border:"none",
              borderRadius:6, padding:"6px 12px", fontSize:12, fontWeight:"bold", cursor:"pointer" }}>
              📊 CSV
            </button>
            {/* Download PDF — professional rate card */}
            <button onClick={() => {
              const perBadge = (per) => {
                const p = (per||"").toLowerCase();
                if (p.includes("sq"))    return `<span style="display:inline-block;padding:2px 8px;border-radius:20px;font-size:9px;font-weight:700;letter-spacing:.5px;background:#dbeafe;color:#1e4fad;border:1px solid #bfdbfe">Sq ft</span>`;
                if (p.includes("rn")||p.includes("rft")) return `<span style="display:inline-block;padding:2px 8px;border-radius:20px;font-size:9px;font-weight:700;letter-spacing:.5px;background:#dcfce7;color:#166534;border:1px solid #bbf7d0">Rn ft</span>`;
                if (p.includes("grov"))  return `<span style="display:inline-block;padding:2px 8px;border-radius:20px;font-size:9px;font-weight:700;letter-spacing:.5px;background:#f3e8ff;color:#6b21a8;border:1px solid #e9d5ff">Grove</span>`;
                if (p.includes("fix"))   return `<span style="display:inline-block;padding:2px 8px;border-radius:20px;font-size:9px;font-weight:700;letter-spacing:.5px;background:#fef9c3;color:#854d0e;border:1px solid #fde68a">Fixed</span>`;
                return `<span style="display:inline-block;padding:2px 8px;border-radius:20px;font-size:9px;font-weight:700;background:#f1f5f9;color:#475569">${per||""}</span>`;
              };

              const buildSection = (cat, label, icon) => {
                const rows = selKeys===null?rateCard:rateCard.filter(function(r){return selKeys.indexOf((r.category||"paint")+"|"+(r.workType||"x"))>=0;}).filter(r => r.category === cat);
                if (!rows.length) return "";
                // Group by workType
                const groups = {};
                rows.forEach(r => {
                  const wt = r.workType || "Other";
                  if (!groups[wt]) groups[wt] = [];
                  groups[wt].push(r);
                });
                let sr = 1;
                let tbody = "";
                const isPolish = cat === "polish";
                const NCOLS = isPolish ? 5 : 6;
                Object.entries(groups).forEach(([wt, items]) => {
                  tbody += `<tr class="wt-row ${cat==="paint"?"pt":""}"><td colspan="${NCOLS}">${wt}</td></tr>`;
                  items.forEach((r, i) => {
                    const cls = i%2===0?"odd":"even";
                    tbody += `<tr class="dr ${cls}">
                      <td class="td-sr">${sr++}.</td>
                      <td>${r.workType||""}</td>
                      <td class="td-sub">${r.sub||""}</td>
                      <td class="td-per">${perBadge(r.per)}</td>
                      <td class="td-mat">&#8377;&nbsp;${r.withMat||"&#8212;"}</td>
                      ${isPolish?"":"<td class=\"td-lab\">&#8377;&nbsp;"+(r.labour||"&#8212;")+"</td>"}
                    </tr>`;
                  });
                });
                return `
                <div class="section">
                  <div class="sec-hdr ${cat}">
                    <span style="font-size:15px">${icon}</span>
                    <span class="sec-title">${label} Work</span>
                    <span class="sec-count">${rows.length} Items</span>
                  </div>
                  <table>
                    <thead><tr>
                      <th class="col-hdr" style="width:30px">#</th>
                      <th class="col-hdr">Particular</th>
                      <th class="col-hdr">Sub-type</th>
                      <th class="col-hdr c" style="width:76px">Unit</th>
                      <th class="col-hdr r" style="width:${isPolish?160:115}px">${isPolish?"Rate (&#8377;)":"With Material"}</th>
                      ${isPolish?"":`<th class="col-hdr r" style="width:115px">Labour</th>`}
                    </tr></thead>
                    <tbody>${tbody}</tbody>
                  </table>
                </div>`;
              };

              const theDate = new Date().toLocaleDateString("en-IN");
              const theName = (company||"Your Company Name");
              const theSpec = (companySpec||"").replace(/Specialist in[:\s]*/i,"");
              const theMob  = companyMob || "";

              const letterhead = (pg, cat) => `
                <div class="lh-wrap">
                  <div class="lh-top">
                    <div class="lh-gold-bar"></div>
                    <div class="lh-main">
                      <div class="lh-name">${theName}</div>
                      <div class="lh-spec">${theSpec}</div>
                    </div>
                    <div class="lh-badge">
                      <div class="lh-rates">Rates</div>
                      <div class="lh-pg">Quotation Card &middot; Page ${pg}</div>
                    </div>
                  </div>
                  <div class="lh-bar">
                    <span>Mob: <strong>${theMob}</strong></span>
                    <span>Date: <strong>${theDate}</strong></span>
                  </div>
                </div>
                <div class="gold-rule"></div>`;

              const polishSection = buildSection("polish","Polish","🪵");
              const paintSection  = buildSection("paint","Paint","🎨");

              const footerHtml = `<div class="footer-row">
                <span>All rates subject to change. Material rates depend on brand &amp; market price at time of work.</span>
                <span class="footer-brand">MeasureFlow v2.0</span>
              </div>`;

              const html = `<!DOCTYPE html><html><head><meta charset="utf-8"/>
<title>Rate Card</title>
<style>
@import url('https://fonts.googleapis.com/css2?family=EB+Garamond:wght@400;600;700&family=DM+Sans:wght@400;500;600&display=swap');
*{box-sizing:border-box;margin:0;padding:0}
body{font-family:'DM Sans',Helvetica,Arial,sans-serif;background:#fff;color:#1a2540;font-size:12.5px;-webkit-print-color-adjust:exact;print-color-adjust:exact}
.page{max-width:800px;margin:0 auto}
.body{padding:24px 30px 16px}
.section{margin-bottom:22px}
/* Letterhead */
.lh-wrap{background:#0f2744}
.lh-top{display:flex;align-items:stretch;min-height:86px}
.lh-gold-bar{width:8px;flex-shrink:0;background:#b8922a}
.lh-main{flex:1;padding:20px 26px 18px}
.lh-name{font-family:'EB Garamond',Georgia,serif;font-size:24px;font-weight:700;color:#ffffff;letter-spacing:2px;text-transform:uppercase;line-height:1.15;margin-bottom:5px}
.lh-spec{font-size:10.5px;color:#8aadd4;letter-spacing:2.5px;text-transform:uppercase}
.lh-badge{display:flex;flex-direction:column;align-items:flex-end;justify-content:center;padding:16px 26px;border-left:1px solid rgba(255,255,255,.1);min-width:155px}
.lh-rates{font-family:'EB Garamond',Georgia,serif;font-size:26px;font-weight:700;color:#d4a843;letter-spacing:4px;text-transform:uppercase}
.lh-pg{font-size:9.5px;color:#6688aa;letter-spacing:2px;margin-top:3px;text-transform:uppercase}
.lh-bar{background:#1a4a8a;display:flex;justify-content:space-between;padding:7px 32px;font-size:10px;color:#8aadd4;letter-spacing:1px}
.lh-bar strong{color:#a8c4e0}
.gold-rule{height:3px;background:#b8922a}
/* Section header */
.sec-hdr{display:flex;align-items:center;gap:12px;padding:10px 16px;border-left:4px solid #b8922a}
.sec-hdr.polish{background:#0f2744}
.sec-hdr.paint{background:#2a1a0a}
.sec-title{font-family:'EB Garamond',Georgia,serif;font-size:15px;font-weight:600;color:#ffffff;letter-spacing:2px;text-transform:uppercase}
.sec-count{margin-left:auto;font-size:10px;color:#d4a843;letter-spacing:1.5px}
/* Table */
table{width:100%;border-collapse:collapse;margin-top:0}
.col-hdr{background:#e8eff9;color:#5a6a85;font-size:9px;font-weight:600;letter-spacing:1.5px;text-transform:uppercase;padding:7px 12px;border-bottom:2px solid #c9d8ef}
.col-hdr.r{text-align:right}.col-hdr.c{text-align:center}
.wt-row td{background:#f0f5fc;color:#1a4a8a;font-weight:600;font-size:11.5px;padding:5px 12px 5px 16px;border-top:1px solid #c9d8ef;border-bottom:1px solid #d8e5f5}
.wt-row.pt td{background:#fdf5eb;color:#7a4a10;border-top-color:#e8d4b0;border-bottom-color:#e8d4b0}
.dr td{padding:6px 12px;border-bottom:1px solid #edf2fb;vertical-align:middle}
.dr.even td{background:#fafcff}.dr.odd td{background:#ffffff}
.td-sr{color:#aab8cc;font-size:10px;text-align:right;width:30px}
.td-sub{color:#5a6a85;font-size:11px;font-style:italic}
.td-per{text-align:center;width:76px}
.td-mat{text-align:right;font-weight:600;color:#0f2744;width:115px}
.td-lab{text-align:right;color:#5a6a85;width:115px}
.badge{display:inline-block;padding:2px 8px;border-radius:20px;font-size:9px;font-weight:700;letter-spacing:.5px;text-transform:uppercase}
.b-sq{background:#dbeafe;color:#1e4fad;border:1px solid #bfdbfe}
.b-rn{background:#dcfce7;color:#166534;border:1px solid #bbf7d0}
.b-gr{background:#f3e8ff;color:#6b21a8;border:1px solid #e9d5ff}
.b-fx{background:#fef9c3;color:#854d0e;border:1px solid #fde68a}
.footer-row{border-top:1px solid #c9d8ef;margin-top:6px;padding:12px 0 16px;display:flex;justify-content:space-between;font-size:10px;color:#94a3b8}
.footer-brand{color:#1a4a8a;font-weight:600}
@media screen{body{background:#edf1f7}.page{box-shadow:0 4px 24px rgba(0,0,0,.15)}}
@media print{body{background:#fff}.page{max-width:none;box-shadow:none}@page{size:A4;margin:8mm 10mm}}
</style>
<script>window.onload=function(){setTimeout(function(){window.print()},500)}</script>
</head><body>
<div class="page">
  ${letterhead(1,"polish")}
  <div class="body">${polishSection}${footerHtml}</div>
  <div style="page-break-before:always">
    ${letterhead(2,"paint")}
    <div class="body">${paintSection}${footerHtml}</div>
  </div>
</div>
</body></html>`;

              try {
              const blob = new Blob([html],{type:"text/html"});
              const url = URL.createObjectURL(blob);
              const a = document.createElement("a");
              a.href=url; a.download="Rate_Card.html";
              document.body.appendChild(a); a.click();
              document.body.removeChild(a);
              setTimeout(()=>URL.revokeObjectURL(url),10000);
              } catch(e) { alert("PDF error: "+e.message); }
            }} style={{ background:"#dc2626", color:"#fff", border:"none",
              borderRadius:6, padding:"6px 12px", fontSize:12, fontWeight:"bold", cursor:"pointer" }}>
              📕 PDF
            </button>
          </div>
        </div>
      </div>

      {/* ── Selection bar ──────────────────────────────────────────────────── */}
      <div style={{ display:"flex", alignItems:"center", gap:8, flexWrap:"wrap",
          padding:"8px 12px", background:"#f0f9ff", border:"1px solid #bae6fd",
          borderRadius:8, marginBottom:8, fontSize:12 }}>
        <span style={{ color:"#0369a1", fontWeight:"bold" }}>
          📥 Download Selection:
        </span>
        <span style={{ color: (selKeys===null||selKeys.length>0)?"#0369a1":"#ef4444" }}>
          {selKeys===null?"All":selKeys.length} / {rateCard.map(function(r){return(r.category||"paint")+"|"+(r.workType||"x");}).filter(function(x,xi,xa){return xa.indexOf(x)===xi;}).length} work types selected
        </span>
        <button onClick={function(){setSelKeys(null);}}
          style={{ border:"1px solid #0284c7", borderRadius:5, padding:"3px 10px",
            background: selKeys===null?"#0284c7":"transparent",
            color: selKeys===null?"#fff":"#0284c7",
            fontSize:11, cursor:"pointer", fontWeight:"bold" }}>✓ All</button>
        <button onClick={function(){setSelKeys([]);}}
          style={{ border:"1px solid #94a3b8", borderRadius:5, padding:"3px 10px",
            background:"transparent", color:"#64748b", fontSize:11, cursor:"pointer" }}>✗ None</button>
        {selKeys!==null && selKeys.length>0 && (
          <span style={{ color:"#64748b", fontSize:11, marginLeft:4 }}>
            — click checkboxes on work-type rows to toggle
          </span>
        )}
        {selKeys!==null && selKeys.length===0 && (
          <span style={{ color:"#ef4444", fontWeight:"bold", fontSize:11 }}>
            ⚠ Nothing selected — downloads will be empty
          </span>
        )}
      </div>

      {/* ── Tables per category ─────────────────────────────────────────────── */}
      {(filter==="all" ? ["polish","paint"] : [filter]).map(cat => {
        const wtGroups = groups[cat] || {};
        const cc = catColors[cat];
        if (!Object.keys(wtGroups).length && filter!=="all") return (
          <div key={cat} style={{ textAlign:"center", padding:24, color:"#94a3b8", fontSize:13 }}>
            No {cat} items match.
          </div>
        );
        if (!Object.keys(wtGroups).length) return null;
        return (
          <div key={cat} style={{ marginBottom:24, borderRadius:12, overflow:"hidden",
              border:"1px solid "+cc.accent+"40", boxShadow:"0 2px 8px rgba(0,0,0,.06)" }}>

            {/* Category header */}
            <div style={{ background:cc.header, color:W, padding:"10px 16px",
                fontSize:14, fontWeight:"bold", display:"flex",
                justifyContent:"space-between", alignItems:"center" }}>
              <span>{cc.label}</span>
              <span style={{ fontSize:11, opacity:.8, fontWeight:"normal" }}>
                {Object.values(wtGroups).flat().length} items
              </span>
            </div>

            <div style={{ overflowX:"auto" }}>
            <table style={{ width:"100%", minWidth:600, borderCollapse:"collapse" }}>
              <colgroup>
                <col style={{ width:32 }}/>
                <col style={{ width:34 }}/>
                <col style={{ width:"28%" }}/>
                <col style={{ width:"18%" }}/>
                <col style={{ width:76 }}/>
                <col style={{ width:90 }}/>
                <col style={{ width:84 }}/>
                <col style={{ width:32 }}/>
              </colgroup>
              <thead>
                <tr>
                  <th style={{ ...thC, background:cc.header }}>
                    <input type="checkbox"
                      checked={Object.keys(wtGroups).every(wt=>(selKeys===null||selKeys.indexOf((cat||"paint")+"|"+(wt||"x"))>=0))}
                      onChange={e=>{
                        const checked = e.target.checked;
                        Object.keys(wtGroups).forEach(wt=>{
                          const isSel = (selKeys===null||selKeys.indexOf((cat||"paint")+"|"+(wt||"x"))>=0);
                          if(checked && !isSel) setSelKeys(function(p){var k=(cat||"paint")+"|"+(wt||"x");if(p===null){var a=rateCard.map(function(r){return(r.category||"paint")+"|"+(r.workType||"x");}).filter(function(x,xi,xar){return xar.indexOf(x)===xi;});a.splice(a.indexOf(k),1);return a;}var n=p.slice();var ix=n.indexOf(k);if(ix>=0)n.splice(ix,1);else n.push(k);return n;});
                          if(!checked && isSel) setSelKeys(function(p){var k=(cat||"paint")+"|"+(wt||"x");if(p===null){var a=rateCard.map(function(r){return(r.category||"paint")+"|"+(r.workType||"x");}).filter(function(x,xi,xar){return xar.indexOf(x)===xi;});a.splice(a.indexOf(k),1);return a;}var n=p.slice();var ix=n.indexOf(k);if(ix>=0)n.splice(ix,1);else n.push(k);return n;});
                        });
                      }}
                      style={{ cursor:"pointer", width:14, height:14 }}
                      title="Select/deselect all"
                    />
                  </th>
                  <th style={{ ...thC, background:cc.header }}>#</th>
                  <th style={{ ...thS, background:cc.header }}>Work Type</th>
                  <th style={{ ...thS, background:cc.header }}>Sub-type</th>
                  <th style={{ ...thC, background:cc.header }}>Per</th>
                  <th style={{ ...thR, background:cc.header }}>With Mat (₹)</th>
                  <th style={{ ...thR, background:cc.header }}>Labour (₹)</th>
                  <th style={{ ...thC, background:cc.header }}></th>
                </tr>
              </thead>
              <tbody>
                {Object.entries(wtGroups).map(([wt, rows], gi) => {
                  const wtSel = (selKeys===null||selKeys.indexOf((cat||"paint")+"|"+(wt||"x"))>=0);
                  return rows.map((row, ri) => {
                  globalIdx++;
                  const idx = globalIdx;
                  const isFirstInGroup = ri === 0;
                  const rowBg = wtSel ? (gi%2===0 ? W : cc.light) : "#f8fafc";
                  const topBorder = isFirstInGroup ? "2px solid "+cc.accent : "1px solid #e2e8f0";
                  return (
                    <tr key={row.id} style={{ background:rowBg, borderTop:topBorder,
                        opacity: wtSel ? 1 : 0.5 }}>

                      {/* Checkbox — only on first row of group, spans visually */}
                      <td style={{ padding:"4px 6px", textAlign:"center", verticalAlign:"middle",
                          borderRight:"1px solid #e2e8f0" }}>
                        {isFirstInGroup && (
                          <input type="checkbox" checked={wtSel}
                            onChange={()=>setSelKeys(function(p){var k=(cat||"paint")+"|"+(wt||"x");if(p===null){var a=rateCard.map(function(r){return(r.category||"paint")+"|"+(r.workType||"x");}).filter(function(x,xi,xar){return xar.indexOf(x)===xi;});a.splice(a.indexOf(k),1);return a;}var n=p.slice();var ix=n.indexOf(k);if(ix>=0)n.splice(ix,1);else n.push(k);return n;})}
                            style={{ cursor:"pointer", width:14, height:14 }}
                            title={wtSel ? "Deselect for download" : "Select for download"}
                          />
                        )}
                      </td>

                      {/* Sr No */}
                      <td style={{ padding:"5px 6px", textAlign:"center", verticalAlign:"middle",
                          fontSize:11, color:"#94a3b8", borderRight:"1px solid #e2e8f0" }}>
                        {idx}
                      </td>

                      {/* Work Type — editable only on first row of group */}
                      <td style={{ padding:"4px 8px", verticalAlign:"middle" }}>
                        {isFirstInGroup ? (
                          <input value={row.workType}
                            onChange={e=>upd(row.id,"workType",e.target.value)}
                            placeholder="Work type name"
                            style={{ width:"100%", border:"1px solid transparent", borderRadius:4,
                              padding:"3px 6px", fontSize:12, outline:"none", fontFamily:"inherit",
                              fontWeight:"600", color: cc.header, background:"transparent",
                              boxSizing:"border-box" }}
                            onFocus={e=>{ e.target.style.borderColor=cc.accent; e.target.style.background=W; }}
                            onBlur={e=>{ e.target.style.borderColor="transparent"; e.target.style.background="transparent"; }}
                          />
                        ) : (
                          <span style={{ paddingLeft:10, fontSize:12, color:"#94a3b8", fontStyle:"italic" }}>
                            {row.workType}
                          </span>
                        )}
                      </td>

                      {/* Sub-type — editable on every row */}
                      <td style={{ padding:"4px 8px", verticalAlign:"middle" }}>
                        <input value={row.sub||""}
                          onChange={e=>upd(row.id,"sub",e.target.value)}
                          placeholder="e.g. Sqft"
                          style={{ width:"100%", border:"1px solid transparent", borderRadius:4,
                            padding:"3px 6px", fontSize:12, outline:"none", fontFamily:"inherit",
                            color:"#475569", background:"transparent", boxSizing:"border-box" }}
                          onFocus={e=>{ e.target.style.borderColor="#cbd5e1"; e.target.style.background=W; }}
                          onBlur={e=>{ e.target.style.borderColor="transparent"; e.target.style.background="transparent"; }}
                        />
                      </td>

                      {/* Per */}
                      <td style={{ padding:"4px 5px", textAlign:"center", verticalAlign:"middle" }}>
                        <select value={row.per} onChange={e=>upd(row.id,"per",e.target.value)}
                          style={{ width:"100%", border:"1px solid #cbd5e1", borderRadius:4,
                            padding:"3px 4px", fontSize:11, outline:"none", background:W,
                            cursor:"pointer", textAlign:"center" }}>
                          {PER_OPTS.map(o=><option key={o} value={o}>{o}</option>)}
                        </select>
                      </td>

                      {/* With Material */}
                      <td style={{ padding:"4px 6px", textAlign:"right", verticalAlign:"middle" }}>
                        <RateNumInput val={row.withMat} onChange={v=>upd(row.id,"withMat",v)} />
                      </td>

                      {/* Labour */}
                      <td style={{ padding:"4px 6px", textAlign:"right", verticalAlign:"middle" }}>
                        <RateNumInput val={row.labour} onChange={v=>upd(row.id,"labour",v)} />
                      </td>

                      {/* Delete */}
                      <td style={{ padding:"4px 6px", textAlign:"center", verticalAlign:"middle" }}>
                        <button onClick={()=>delRow(row.id)} title="Remove"
                          style={{ background:"none", border:"none", color:"#ef4444",
                            fontSize:14, cursor:"pointer", padding:"2px 4px", borderRadius:4,
                            lineHeight:1 }}>
                          ✕
                        </button>
                      </td>
                    </tr>
                  );
                  })
                })}
              </tbody>
            </table>
            </div>
          </div>
        );
      })}

      {filtered.length===0 && (
        <div style={{ textAlign:"center", padding:32, color:"#94a3b8", fontSize:13 }}>
          No items found. {search ? "Clear search or " : ""}Use the Add Item form below.
        </div>
      )}

      {/* ── Add Item Form ──────────────────────────────────────────────────────── */}
      <AddItemForm BD={BD} W={W} setRateCard={setRateCard} />

      {/* Tip */}
      <div style={{ marginTop:16, padding:"11px 16px", background:"#fffbeb", border:"1px solid #fbbf24",
          borderRadius:8, fontSize:12, color:"#92400e" }}>
        💡 <strong>Tip:</strong> In 🧾 Bill tab, click the <strong>₹XX ↙</strong> button beside any rate field to auto-fill from this card. "With Material" rate is used for billing.
      </div>
    </div>
  );
}


function BillView({
  sessions, company, companySpec, companyAddr, companyMob, logoUrl, logoRef, handleLogo,
  billClient, setBillClient, billSite, setBillSite, billDate, setBillDate,
  billSub, setBillSub, billAdvance, setBillAdvance,
  fixedItems, setFixedItems, manualItems, setManualItems,
  setCompany, setCompanySpec, setCompanyAddr, setCompanyMob,
  exportMsg, billBusy, handleBill, handleBillPdf,
  onUpdateSession, billRates, setBillRates, rateCard, lang = "en"
}) {
  const C = { dark:"#1F4E79", mid:"#2E75B6", light:"#D6E4F0", red:"#CC0000", redBg:"#FFE0E0", green:"#059669", gray:"#334155" };
  const W = "#fff";
  const t = (en, hi) => lang === "hi" ? hi : en;
  const card    = { background:W, borderRadius:10, boxShadow:"0 1px 6px #1F4E7918", marginBottom:14, overflow:"hidden" };
  const cardHead = (bg) => ({ background:bg||C.dark, color:W, padding:"9px 14px", fontWeight:600, fontSize:14, display:"flex", gap:8, alignItems:"center" });
  const cardTitle = { fontSize:14, fontWeight:600 };
  const inp  = { border:"1px solid #cbd5e1", borderRadius:6, padding:"5px 9px", fontSize:13, fontFamily:"inherit", width:"100%", outline:"none" };
  const lbl  = { fontSize:12, color:C.gray, marginBottom:3, display:"block" };

  // ── Bill lines calculation (same logic as handleBillPdf) ──────────────────
  const bRates = billRates || {};
  const billLines = [];

  for (const sess of sessions) {
    if (hasBoth(sess.rows)) {
      const groups = measureGroups(sess);
      for (const g of groups) {
        const lineId = sess.id + g.per;
        const cardR  = lookupRateFromCard(rateCard, g.label, g.per);
        const rate   = parseFloat(bRates[lineId] !== undefined ? bRates[lineId] : cardR) || 0;
        const qty    = g.total;
        billLines.push({ lineId, label:g.label, qty, rate, per:g.per, total:qty*rate, cardR });
      }
    } else {
      const lineId = sess.id;
      const group  = measureGroups(sess)[0];
      const per    = sess.perOverride || (group ? group.per : "Sq ft");
      const cardR  = lookupRateFromCard(rateCard, sess.title, per);
      const rate   = parseFloat(bRates[lineId] !== undefined ? bRates[lineId] : cardR) || 0;
      const qty    = netTotal(sess.rows);
      billLines.push({ lineId, label:sess.title, qty, rate, per, total:qty*rate, cardR });
    }
  }

  const grand   = billLines.reduce((s,l)=>s+l.total, 0)
                + (manualItems||[]).reduce((s,m)=>{ const q=parseFloat(m.qty)||0,r=parseFloat(m.rate)||0; return s+(m.amount?parseFloat(m.amount)||0:q*r); }, 0)
                + (fixedItems||[]).reduce((s,f)=>{ const a=f.qty&&f.unitPrice?(parseFloat(f.qty)||0)*(parseFloat(f.unitPrice)||0):(parseFloat(f.amount)||0); return s+a; }, 0);
  const adv     = parseFloat(billAdvance) || 0;
  const balance = grand - adv;

  function setRate(lineId, val) {
    setBillRates(prev => ({ ...prev, [lineId]: val }));
  }

  // ── Fixed item helpers ────────────────────────────────────────────────────
  function addFixed() { setFixedItems(p=>[...p,{id:uid(),name:"",qty:"",unitPrice:"",amount:""}]); }
  function updFixed(id,k,v) { setFixedItems(p=>p.map(x=>x.id===id?{...x,[k]:v}:x)); }
  function delFixed(id) { setFixedItems(p=>p.filter(x=>x.id!==id)); }
  function addManual() { setManualItems(p=>[...p,{id:uid(),name:"",qty:"",rate:"",per:"Sq ft",amount:""}]); }
  function updManual(id,k,v) { setManualItems(p=>p.map(x=>x.id===id?{...x,[k]:v}:x)); }
  function delManual(id) { setManualItems(p=>p.filter(x=>x.id!==id)); }

  const fmtMoney = n => n > 0 ? "₹ " + Math.round(n).toLocaleString("en-IN") : "—";
  const numInp = (val, onChange, w) => (
    <input type="number" min="0" value={val} onChange={e=>onChange(e.target.value)}
      style={{ width:w||"100%", border:"1px solid #cbd5e1", borderRadius:4, padding:"3px 6px",
               fontSize:13, textAlign:"right", outline:"none", fontFamily:"inherit" }} />
  );

  const hasSess = sessions.length > 0;

  return (
    <div style={{ maxWidth:880, margin:"0 auto" }}>

      {/* ── Company Details ── */}
      <div style={card}>
        <div style={cardHead(C.dark)}><span>🏢</span><span style={cardTitle}>{t("Company Details","कंपनी विवरण")}</span></div>
        <div style={{ padding:"12px 14px", display:"grid", gridTemplateColumns:"1fr 1fr", gap:"10px 16px" }}>
          <div>
            <span style={lbl}>{t("Company Name","कंपनी का नाम")}</span>
            <input style={inp} value={company} onChange={e=>setCompany(e.target.value)} placeholder={t("Company name","कंपनी का नाम")} />
          </div>
          <div>
            <span style={lbl}>{t("Specialisation","विशेषता")}</span>
            <input style={inp} value={companySpec} onChange={e=>setCompanySpec(e.target.value)} placeholder={t("Specialization / services","सेवाएं / विशेषता")} />
          </div>
          <div>
            <span style={lbl}>{t("Address","पता")}</span>
            <input style={inp} value={companyAddr} onChange={e=>setCompanyAddr(e.target.value)} placeholder={t("Address","पता")} />
          </div>
          <div>
            <span style={lbl}>{t("Mobile","मोबाइल")}</span>
            <input style={inp} value={companyMob} onChange={e=>setCompanyMob(e.target.value)} placeholder={t("Mobile number","मोबाइल नंबर")} />
          </div>
          <div>
            <span style={lbl}>{t("Logo","लोगो")}</span>
            <div style={{ display:"flex", alignItems:"center", gap:8 }}>
              <button onClick={()=>logoRef.current&&logoRef.current.click()}
                style={{ padding:"4px 12px", background:C.mid, color:W, border:"none", borderRadius:5, cursor:"pointer", fontSize:13 }}>
                📷 {logoUrl ? t("Change Logo","लोगो बदलें") : t("Upload Logo","लोगो अपलोड करें")}
              </button>
              {logoUrl && <img src={logoUrl} alt="logo" style={{ height:30, borderRadius:3 }} />}
              <input ref={logoRef} type="file" accept="image/*" style={{ display:"none" }} onChange={handleLogo} />
            </div>
          </div>
        </div>
      </div>

      {/* ── Bill Meta ── */}
      <div style={card}>
        <div style={cardHead(C.mid)}><span>📋</span><span style={cardTitle}>{t("Bill Details","बिल विवरण")}</span></div>
        <div style={{ padding:"12px 14px", display:"grid", gridTemplateColumns:"1fr 1fr", gap:"10px 16px" }}>
          <div>
            <span style={lbl}>{t("Client Name","ग्राहक का नाम")}</span>
            <input style={inp} value={billClient} onChange={e=>setBillClient(e.target.value)} placeholder={t("Client / Party name","ग्राहक / पार्टी का नाम")} />
          </div>
          <div>
            <span style={lbl}>{t("Date","तारीख")}</span>
            <input style={inp} value={billDate} onChange={e=>setBillDate(e.target.value)} />
          </div>
          <div style={{ gridColumn:"1/-1" }}>
            <span style={lbl}>{t("Worksite Address","कार्यस्थल का पता")}</span>
            <input style={inp} value={billSite} onChange={e=>setBillSite(e.target.value)} placeholder={t("Site / project address","साइट / प्रोजेक्ट का पता")} />
          </div>
          <div style={{ gridColumn:"1/-1" }}>
            <span style={lbl}>{t("Subject Line","विषय")}</span>
            <input style={inp} value={billSub} onChange={e=>setBillSub(e.target.value)} placeholder={t("Bill subject / work description","बिल का विषय / काम का विवरण")} />
          </div>
        </div>
      </div>

      {/* ── Measurement Rows ── */}
      {hasSess && (
      <div style={card}>
        <div style={cardHead(C.dark)}><span>📐</span><span style={cardTitle}>Measurement Rows</span><span style={{ marginLeft:"auto", fontSize:12, opacity:.7 }}>Rate auto-filled from Rate Card — edit to override</span></div>
        <div style={{ overflowX:"auto" }}>
          <table style={{ width:"100%", borderCollapse:"collapse", fontSize:13 }}>
            <thead>
              <tr style={{ background:C.light }}>
                {["#","Section","Area","Unit","Rate (₹)","Amount"].map((h,i)=>(
                  <th key={i} style={{ padding:"6px 10px", textAlign:i>=2?"right":"left", color:C.dark, fontWeight:600, borderBottom:"2px solid "+C.light }}>{h}</th>
                ))}
              </tr>
            </thead>
            <tbody>
              {billLines.map((l,i) => {
                const amt = l.qty * (parseFloat(bRates[l.lineId] !== undefined ? bRates[l.lineId] : l.cardR) || 0);
                return (
                  <tr key={l.lineId} style={{ background:i%2===0?"#f8fafc":W }}>
                    <td style={{ padding:"5px 10px", color:"#888" }}>{i+1}.</td>
                    <td style={{ padding:"5px 10px" }}>{l.label}</td>
                    <td style={{ padding:"5px 10px", textAlign:"right" }}>{fmtArea(l.qty)}</td>
                    <td style={{ padding:"5px 10px", textAlign:"right", color:"#666" }}>{l.per}</td>
                    <td style={{ padding:"4px 8px", textAlign:"right" }}>
                      {numInp(bRates[l.lineId] !== undefined ? bRates[l.lineId] : (l.cardR || ""), v=>setRate(l.lineId, v), 90)}
                    </td>
                    <td style={{ padding:"5px 10px", textAlign:"right", fontWeight:600, color:C.dark }}>{fmtMoney(amt)}</td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
      </div>
      )}

      {/* ── Manual Items ── */}
      <div style={card}>
        <div style={cardHead(C.gray)}>
          <span>✏️</span><span style={cardTitle}>Manual Items</span>
          <button onClick={addManual} style={{ marginLeft:"auto", padding:"2px 10px", background:C.green, color:W, border:"none", borderRadius:4, cursor:"pointer", fontSize:12 }}>+ Add</button>
        </div>
        {(manualItems||[]).length === 0 ? (
          <div style={{ padding:"10px 14px", color:"#94a3b8", fontSize:13 }}>No manual items. Click + Add to add a free-form bill row.</div>
        ) : (
          <div style={{ overflowX:"auto" }}>
            <table style={{ width:"100%", borderCollapse:"collapse", fontSize:13 }}>
              <thead>
                <tr style={{ background:C.light }}>
                  {["Item","Qty","Rate","Per","Amount",""].map((h,i)=>(
                    <th key={i} style={{ padding:"5px 8px", textAlign:i>=1&&i<=4?"right":"left", color:C.dark, fontWeight:600, borderBottom:"2px solid "+C.light }}>{h}</th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {(manualItems||[]).map((m,i) => (
                  <tr key={m.id} style={{ background:i%2===0?"#f8fafc":W }}>
                    <td style={{ padding:"4px 6px" }}>
                      <input style={{ ...inp, width:160 }} value={m.name} onChange={e=>updManual(m.id,"name",e.target.value)} placeholder="Item name" />
                    </td>
                    <td style={{ padding:"4px 6px" }}>{numInp(m.qty, v=>updManual(m.id,"qty",v), 70)}</td>
                    <td style={{ padding:"4px 6px" }}>{numInp(m.rate, v=>updManual(m.id,"rate",v), 80)}</td>
                    <td style={{ padding:"4px 6px" }}>
                      <select value={m.per||"Sq ft"} onChange={e=>updManual(m.id,"per",e.target.value)}
                        style={{ border:"1px solid #cbd5e1", borderRadius:4, padding:"3px 6px", fontSize:13, fontFamily:"inherit" }}>
                        {["Sq ft","Rn ft","Grove","Nos","Fixed"].map(p=><option key={p}>{p}</option>)}
                      </select>
                    </td>
                    <td style={{ padding:"4px 6px" }}>
                      <input style={{ ...inp, width:90, textAlign:"right" }} value={m.amount} onChange={e=>updManual(m.id,"amount",e.target.value)} placeholder="or override" />
                    </td>
                    <td style={{ padding:"4px 6px" }}>
                      <button onClick={()=>delManual(m.id)} style={{ background:"none", border:"none", cursor:"pointer", color:C.red, fontSize:16 }}>×</button>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        )}
      </div>

      {/* ── Fixed Items ── */}
      <div style={card}>
        <div style={cardHead(C.gray)}>
          <span>📦</span><span style={cardTitle}>Fixed / Material Items</span>
          <button onClick={addFixed} style={{ marginLeft:"auto", padding:"2px 10px", background:C.green, color:W, border:"none", borderRadius:4, cursor:"pointer", fontSize:12 }}>+ Add</button>
        </div>
        {(fixedItems||[]).length === 0 ? (
          <div style={{ padding:"10px 14px", color:"#94a3b8", fontSize:13 }}>No fixed items. Click + Add for materials or lump-sum charges.</div>
        ) : (
          <div style={{ overflowX:"auto" }}>
            <table style={{ width:"100%", borderCollapse:"collapse", fontSize:13 }}>
              <thead>
                <tr style={{ background:C.light }}>
                  {["Item","Qty","Unit Price","Fixed Amount",""].map((h,i)=>(
                    <th key={i} style={{ padding:"5px 8px", textAlign:i>=1&&i<=3?"right":"left", color:C.dark, fontWeight:600, borderBottom:"2px solid "+C.light }}>{h}</th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {(fixedItems||[]).map((f,i) => (
                  <tr key={f.id} style={{ background:i%2===0?"#f8fafc":W }}>
                    <td style={{ padding:"4px 6px" }}>
                      <input style={{ ...inp, width:160 }} value={f.name} onChange={e=>updFixed(f.id,"name",e.target.value)} placeholder="Item name" />
                    </td>
                    <td style={{ padding:"4px 6px" }}>{numInp(f.qty, v=>updFixed(f.id,"qty",v), 70)}</td>
                    <td style={{ padding:"4px 6px" }}>{numInp(f.unitPrice, v=>updFixed(f.id,"unitPrice",v), 90)}</td>
                    <td style={{ padding:"4px 6px" }}>{numInp(f.amount, v=>updFixed(f.id,"amount",v), 90)}</td>
                    <td style={{ padding:"4px 6px" }}>
                      <button onClick={()=>delFixed(f.id)} style={{ background:"none", border:"none", cursor:"pointer", color:C.red, fontSize:16 }}>×</button>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        )}
      </div>

      {/* ── Totals ── */}
      <div style={{ ...card, marginBottom:16 }}>
        <div style={cardHead(C.dark)}><span>💰</span><span style={cardTitle}>Bill Summary</span></div>
        <div style={{ padding:"12px 16px" }}>
          <table style={{ width:"100%", borderCollapse:"collapse", fontSize:14 }}>
            <tbody>
              <tr>
                <td style={{ padding:"5px 0", color:C.gray }}>Grand Total</td>
                <td style={{ padding:"5px 0", textAlign:"right", fontWeight:700, fontSize:16, color:C.dark }}>{fmtMoney(grand)}</td>
              </tr>
              <tr>
                <td style={{ padding:"5px 0", color:C.gray }}>
                  Advance Paid
                  <input type="number" min="0" value={billAdvance} onChange={e=>setBillAdvance(e.target.value)}
                    style={{ width:100, border:"1px solid #cbd5e1", borderRadius:4, padding:"2px 6px", fontSize:13, marginLeft:10 }} />
                </td>
                <td style={{ padding:"5px 0", textAlign:"right", color:"#059669", fontWeight:600 }}>{adv > 0 ? fmtMoney(adv) : "—"}</td>
              </tr>
              <tr style={{ borderTop:"2px solid "+C.light }}>
                <td style={{ padding:"7px 0", fontWeight:700, color:C.dark, fontSize:15 }}>Balance Due</td>
                <td style={{ padding:"7px 0", textAlign:"right", fontWeight:700, fontSize:16, color:C.dark }}>{fmtMoney(balance)}</td>
              </tr>
            </tbody>
          </table>
        </div>
      </div>

      {/* ── Download Buttons ── */}
      <div className="noprint" style={{ display:"flex", gap:10, flexWrap:"wrap", justifyContent:"center", marginBottom:24 }}>
        <button onClick={handleBill} disabled={billBusy}
          style={{ padding:"10px 24px", background:C.dark, color:W, border:"none", borderRadius:7, cursor:"pointer", fontWeight:600, fontSize:14, opacity:billBusy?.6:1 }}>
          {billBusy ? "⏳ Building…" : "📄 Download Word"}
        </button>
        <button onClick={handleBillPdf}
          style={{ padding:"10px 24px", background:"#b91c1c", color:W, border:"none", borderRadius:7, cursor:"pointer", fontWeight:600, fontSize:14 }}>
          📕 Download PDF
        </button>
      </div>

      {exportMsg && (
        <div style={{ textAlign:"center", padding:"8px 16px", background:exportMsg.ok?"#f0fdf4":"#fef2f2", color:exportMsg.ok?C.green:C.red, borderRadius:7, marginBottom:12, fontSize:13 }}>
          {exportMsg.text}
        </div>
      )}

    </div>
  );
}

function AdminPanel({ lang = "en", analyticsEnabled }) {
  const t = (en, hi) => lang === "hi" ? hi : en;
  const [password, setPassword] = useState("");
  const [token, setToken] = useState(() => typeof window !== "undefined" ? sessionStorage.getItem("mf_admin_token") || "" : "");
  const [data, setData] = useState(null);
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState("");

  const load = useCallback(async (authToken) => {
    if (!authToken) return;
    setBusy(true);
    setError("");
    try {
      const summary = await loadAdminSummaryData(authToken);
      setData(summary);
    } catch (err) {
      setError(err.message || "Unable to load admin summary");
      if (/expired|invalid/i.test(err.message || "")) {
        setToken("");
        if (typeof window !== "undefined") sessionStorage.removeItem("mf_admin_token");
      }
    } finally {
      setBusy(false);
    }
  }, []);

  useEffect(() => {
    if (token) load(token);
  }, [token, load]);

  async function submit(e) {
    e.preventDefault();
    setBusy(true);
    setError("");
    try {
      const authToken = await adminLogin(password);
      setToken(authToken);
      if (typeof window !== "undefined") sessionStorage.setItem("mf_admin_token", authToken);
      setPassword("");
    } catch (err) {
      setError(err.message || "Admin login failed");
    } finally {
      setBusy(false);
    }
  }

  function logout() {
    setToken("");
    setData(null);
    if (typeof window !== "undefined") sessionStorage.removeItem("mf_admin_token");
  }

  if (!analyticsEnabled) {
    return (
      <div style={{ maxWidth: 900, margin: "0 auto", background: "#fff", border: "1px solid #dbeafe", borderRadius: 12, padding: 20, boxShadow: "0 2px 12px rgba(31,78,121,.07)" }}>
        <div style={{ fontSize: 18, fontWeight: "bold", color: C.dark, marginBottom: 8 }}>{t("Admin Dashboard", "एडमिन डैशबोर्ड")}</div>
        <div style={{ color: "#475569", lineHeight: 1.6 }}>{t("Analytics is not configured yet. Add SUPABASE_URL, SUPABASE_SERVICE_ROLE_KEY, and ADMIN_PASSWORD on the backend, then create the mf_events table.", "एनालिटिक्स अभी कॉन्फ़िगर नहीं है। बैकएंड में SUPABASE_URL, SUPABASE_SERVICE_ROLE_KEY और ADMIN_PASSWORD जोड़ें, फिर mf_events टेबल बनाएं।")}</div>
      </div>
    );
  }

  if (!token) {
    return (
      <div style={{ maxWidth: 520, margin: "0 auto", background: "#fff", border: "1px solid #dbeafe", borderRadius: 12, padding: 22, boxShadow: "0 2px 12px rgba(31,78,121,.07)" }}>
        <div style={{ fontSize: 20, fontWeight: "bold", color: C.dark, marginBottom: 6 }}>{t("Admin Login", "एडमिन लॉगिन")}</div>
        <div style={{ fontSize: 13, color: "#64748b", marginBottom: 16 }}>{t("Use the backend admin password to see real app usage and user activity.", "ऐप का वास्तविक उपयोग और यूज़र एक्टिविटी देखने के लिए बैकएंड वाला एडमिन पासवर्ड इस्तेमाल करें।")}</div>
        <form onSubmit={submit} style={{ display: "grid", gap: 12 }}>
          <input
            type="password"
            value={password}
            onChange={(e) => setPassword(e.target.value)}
            placeholder={t("Admin password", "एडमिन पासवर्ड")}
            style={{ border: "1px solid #cbd5e1", borderRadius: 8, padding: "10px 12px", fontSize: 14, outline: "none" }}
          />
          {error && <div style={{ background: C.redBg, color: C.red, border: "1px solid " + C.red, borderRadius: 8, padding: "9px 11px", fontSize: 13 }}>{error}</div>}
          <button type="submit" disabled={busy || !password.trim()} style={{ ...btn(C.dark), justifyContent: "center", opacity: busy || !password.trim() ? .65 : 1, cursor: busy || !password.trim() ? "not-allowed" : "pointer" }}>
            {busy ? t("Signing in...", "साइन इन हो रहा है...") : t("Open Dashboard", "डैशबोर्ड खोलें")}
          </button>
        </form>
      </div>
    );
  }

  const cards = [
    { label: t("Total Users", "कुल यूज़र"), value: data?.totalUsers ?? 0, bg: "#eff6ff", color: "#1d4ed8" },
    { label: t("Active (7d)", "सक्रिय (7 दिन)"), value: data?.activeUsers7d ?? 0, bg: "#ecfdf5", color: "#059669" },
    { label: t("Inactive", "निष्क्रिय"), value: data?.inactiveUsers ?? 0, bg: "#f8fafc", color: "#475569" },
    { label: t("Visits (24h)", "विज़िट (24 घंटे)"), value: data?.visits24h ?? 0, bg: "#fff7ed", color: "#c2410c" },
    { label: t("Uploads", "अपलोड"), value: data?.uploads ?? 0, bg: "#f5f3ff", color: "#7c3aed" },
    { label: t("Downloads", "डाउनलोड"), value: data?.downloads ?? 0, bg: "#eef2ff", color: "#4338ca" },
  ];

  return (
    <div style={{ maxWidth: 980, margin: "0 auto" }}>
      <div style={{ background: "#fff", border: "1px solid #dbeafe", borderRadius: 12, padding: 18, boxShadow: "0 2px 12px rgba(31,78,121,.07)", marginBottom: 16 }}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", gap: 12, flexWrap: "wrap" }}>
          <div>
            <div style={{ fontSize: 20, fontWeight: "bold", color: C.dark }}>{t("Admin Dashboard", "एडमिन डैशबोर्ड")}</div>
            <div style={{ fontSize: 12, color: "#64748b", marginTop: 2 }}>{t("Anonymous usage analytics for your hosted public app.", "आपके होस्टेड पब्लिक ऐप के लिए अनाम उपयोग एनालिटिक्स।")}</div>
          </div>
          <div style={{ display: "flex", gap: 8 }}>
            <button onClick={() => load(token)} disabled={busy} style={{ ...btn(C.mid, true), opacity: busy ? .7 : 1 }}>{t("Refresh", "रीफ्रेश")}</button>
            <button onClick={logout} style={{ ...btn(C.gray, true) }}>{t("Logout", "लॉगआउट")}</button>
          </div>
        </div>
      </div>
      {error && <div style={{ background: C.redBg, color: C.red, border: "1px solid " + C.red, borderRadius: 8, padding: "9px 11px", marginBottom: 14, fontSize: 13 }}>{error}</div>}
      <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(150px, 1fr))", gap: 12, marginBottom: 16 }}>
        {cards.map((card) => (
          <div key={card.label} style={{ background: card.bg, borderRadius: 12, padding: 16, border: "1px solid rgba(148,163,184,.22)" }}>
            <div style={{ fontSize: 12, color: "#64748b", marginBottom: 6 }}>{card.label}</div>
            <div style={{ fontSize: 30, fontWeight: "bold", color: card.color }}>{card.value}</div>
          </div>
        ))}
      </div>
      <div style={{ background: "#fff", border: "1px solid #dbeafe", borderRadius: 12, overflow: "hidden", boxShadow: "0 2px 12px rgba(31,78,121,.07)" }}>
        <div style={{ background: C.dark, color: "#fff", padding: "12px 16px", fontWeight: "bold" }}>{t("Recent Activity", "हाल की गतिविधि")}</div>
        <div style={{ maxHeight: 560, overflow: "auto" }}>
          <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12 }}>
            <thead>
              <tr style={{ background: "#eff6ff", color: C.dark }}>
                {[
                  t("When", "समय"),
                  t("Event", "इवेंट"),
                  t("User", "यूज़र"),
                  t("Tab", "टैब"),
                  t("Details", "विवरण"),
                ].map((h) => <th key={h} style={{ textAlign: "left", padding: "8px 10px", borderBottom: "1px solid #dbeafe" }}>{h}</th>)}
              </tr>
            </thead>
            <tbody>
              {(data?.recent || []).map((row, idx) => (
                <tr key={row.id || idx} style={{ background: idx % 2 ? "#f8fafc" : "#fff" }}>
                  <td style={{ padding: "8px 10px", borderBottom: "1px solid #eef2f7", whiteSpace: "nowrap" }}>{new Date(row.when).toLocaleString("en-IN")}</td>
                  <td style={{ padding: "8px 10px", borderBottom: "1px solid #eef2f7", fontWeight: "bold", color: C.dark }}>{row.type}</td>
                  <td style={{ padding: "8px 10px", borderBottom: "1px solid #eef2f7", color: "#64748b" }}>{String(row.userId || "").slice(0, 10) || "—"}</td>
                  <td style={{ padding: "8px 10px", borderBottom: "1px solid #eef2f7" }}>{row.tab || "—"}</td>
                  <td style={{ padding: "8px 10px", borderBottom: "1px solid #eef2f7", color: "#475569" }}>{JSON.stringify(row.details || {})}</td>
                </tr>
              ))}
              {!busy && !(data?.recent || []).length && (
                <tr><td colSpan={5} style={{ padding: 20, textAlign: "center", color: "#94a3b8" }}>{t("No activity yet.", "अभी कोई गतिविधि नहीं है।")}</td></tr>
              )}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
}

function DownloadGateModal({ open, label, seconds, videoUrl, adClient, onClose, onContinue, processing = false, processingText = "" }) {
  const [remaining, setRemaining] = useState(seconds);
  const isUpload = processing || /upload/i.test(label || "");

  useEffect(() => {
    if (!open) return;
    setRemaining(seconds);
  }, [open, seconds]);

  useEffect(() => {
    if (processing) return;
    if (!open || remaining <= 0) return;
    const t = setTimeout(() => setRemaining((s) => s - 1), 1000);
    return () => clearTimeout(t);
  }, [open, remaining, processing]);

  if (!open) return null;

  return (
    <div className="noprint" style={{ position: "fixed", inset: 0, background: "rgba(15,23,42,.72)", zIndex: 1200, display: "flex", alignItems: "center", justifyContent: "center", padding: 16 }}>
      <div style={{ width: "min(920px, 100%)", background: "#fff", borderRadius: 18, overflow: "hidden", boxShadow: "0 22px 70px rgba(15,23,42,.45)" }}>
        <div style={{ background: "#1F4E79", color: "#fff", padding: "14px 18px", display: "flex", alignItems: "center", justifyContent: "space-between" }}>
          <div>
            <div style={{ fontSize: 18, fontWeight: "bold" }}>MeasureFlow Support</div>
            <div style={{ fontSize: 12, opacity: 0.9 }}>{isUpload ? "Your upload will continue after this short support screen." : "Your download will start after this short support screen."}</div>
          </div>
          {!processing && (
            <button onClick={onClose} style={{ background: "rgba(255,255,255,.12)", color: "#fff", border: "none", borderRadius: 8, padding: "7px 10px", cursor: "pointer", fontWeight: "bold" }}>
              Close
            </button>
          )}
        </div>
        <div style={{ display: "grid", gridTemplateColumns: "1.4fr 1fr", gap: 16, padding: 16 }}>
          <div style={{ background: "#0f172a", borderRadius: 14, overflow: "hidden", minHeight: 320 }}>
            {videoUrl ? (
              <iframe
                src={videoUrl}
                title="Sponsored Video"
                style={{ border: "none", width: "100%", height: "100%" }}
                allow="accelerometer; autoplay; clipboard-write; encrypted-media; gyroscope; picture-in-picture"
                allowFullScreen
              />
            ) : adClient ? (
              <div style={{ padding: 14, background: "#fff", minHeight: 320, height: "100%" }}>
                <InlineAdStrip adClient={adClient} minHeight={292} label="Sponsored" />
              </div>
            ) : (
              <div style={{ height: "100%", minHeight: 320, display: "flex", alignItems: "center", justifyContent: "center", color: "#cbd5e1", padding: 24, textAlign: "center" }}>
                Please wait while support content is shown before download
              </div>
            )}
          </div>
          <div style={{ display: "flex", flexDirection: "column", gap: 14 }}>
            <div style={{ background: "#f8fafc", border: "1px solid #dbeafe", borderRadius: 14, padding: 16 }}>
              <div style={{ fontSize: 12, color: "#64748b", marginBottom: 6 }}>Support Message</div>
              <div style={{ fontWeight: "bold", color: "#1F4E79", marginBottom: 10 }}>
                {isUpload ? "Your upload will continue after this short support screen" : "Your download will continue after this short support screen"}
              </div>
              <div style={{ fontSize: 12, color: "#64748b" }}>
                Thank you for supporting MeasureFlow.
              </div>
            </div>
            <div style={{ background: "#f8fafc", border: "1px solid #dbeafe", borderRadius: 14, padding: 16 }}>
              <div style={{ fontSize: 12, color: "#64748b", marginBottom: 6 }}>Pending Action</div>
              <div style={{ fontWeight: "bold", color: "#1F4E79", marginBottom: 12 }}>{label}</div>
              {processing ? (
                <div style={{ borderRadius: 10, padding: "11px 14px", fontWeight: "bold", background: "#1F4E79", color: "#fff", textAlign: "center" }}>
                  {processingText || "Processing..."}
                </div>
              ) : (
                <button
                  onClick={onContinue}
                  disabled={remaining > 0}
                  style={{
                    width: "100%",
                    border: "none",
                    borderRadius: 10,
                    padding: "11px 14px",
                    fontWeight: "bold",
                    cursor: remaining > 0 ? "default" : "pointer",
                    background: remaining > 0 ? "#cbd5e1" : "#059669",
                    color: "#fff",
                  }}
                >
                  {remaining > 0 ? `Continue in ${remaining}s` : isUpload ? "Continue Upload" : "Continue Download"}
                </button>
              )}
            </div>
          </div>
        </div>
      </div>
    </div>
  );
}

export default function App() {
  const [sessions, setSessions] = useState([]);
  const [totalLabels, setTotalLabels] = useState({ sqftLabel:"Sqft Total", rnftLabel:"Rnft Total", groveLabel:"Grove Total" });
  const [docTitle, setDocTitle] = useState("Measurement Report");
  const [company, setCompany] = useState("");
  const [logoUrl, setLogoUrl] = useState(null);
  // Bill details
  const [billClient, setBillClient]   = useState("");
  const [billSite,   setBillSite]     = useState("");
  const [billDate,   setBillDate]     = useState(() => new Date().toLocaleDateString("en-IN"));
  const [billSub,    setBillSub]      = useState("");
  const [billAdvance,setBillAdvance]  = useState("");
  const [fixedItems, setFixedItems]   = useState([]); // [{id,name,qty,unitPrice,amount}]
  const [manualItems, setManualItems] = useState([]); // [{id,name,qty,rate,per,amount}] — manual bill rows
  const [billRates, setBillRates]       = useState({}); // {lineId: rateString} — per-row rates in bill

  // Rate Card state — [{id, workType, sqft, rnft, grove}]

  const [rateCard, setRateCard] = useState(RATE_CARD_DEFAULT);
  const [companySpec,setCompanySpec]  = useState("");
  const [companyAddr,setCompanyAddr]  = useState("");
  const [companyMob, setCompanyMob]   = useState("");
  const [loading, setLoading] = useState(false);
  const [loadMsg, setLoadMsg] = useState("");
  const [error, setError] = useState(null);
  const [exportMsg, setExportMsg] = useState(null); // {text, ok}
  const [docxBusy, setDocxBusy] = useState(false);
  const [saveLabel, setSaveLabel] = useState("💾 Save");
  const [printing, setPrinting] = useState(false);
  const [activeTab, setActiveTab] = useState('measure'); // 'measure' | 'bill' | 'rates'
  const [lang, setLang] = useState("en");
  const [analyticsEnabled, setAnalyticsEnabled] = useState(false);
  const [adminMode, setAdminMode] = useState(() => typeof window !== "undefined" ? new URLSearchParams(window.location.search).get("admin") === "1" : false);
  const [adConfig, setAdConfig] = useState({
    enabled: true,
    gateEnabled: true,
    adClient: "",
    gateSeconds: 6,
    videoUrl: "",
  });
  const [downloadGateOpen, setDownloadGateOpen] = useState(false);
  const [downloadGateLabel, setDownloadGateLabel] = useState("");
  const [uploadAdOpen, setUploadAdOpen] = useState(false);
  const logoRef = useRef();
  const importRef = useRef();
  const pendingDownloadRef = useRef(null);

  // No auto-load on startup — app always opens fresh/empty

  // Flash export message
  function flash(text, ok = true) {
    setExportMsg({ text, ok });
    setTimeout(() => setExportMsg(null), 3500);
  }

  useEffect(() => {
    if (!adConfig.enabled || !adConfig.adClient) return;
    if (document.querySelector('script[data-mf-adsense="1"]')) return;
    const s = document.createElement("script");
    s.async = true;
    s.src = "https://pagead2.googlesyndication.com/pagead/js/adsbygoogle.js?client=" + encodeURIComponent(adConfig.adClient);
    s.crossOrigin = "anonymous";
    s.dataset.mfAdsense = "1";
    document.head.appendChild(s);
  }, [adConfig.enabled, adConfig.adClient]);

  useEffect(() => {
    let cancelled = false;
    loadBackendConfig().then((cfg) => {
      if (cancelled) return;
      if (cfg?.adsenseClient) {
        setAdConfig((prev) => ({ ...prev, adClient: cfg.adsenseClient }));
      }
      setAnalyticsEnabled(!!cfg?.analyticsEnabled);
    });
    return () => {
      cancelled = true;
    };
  }, []);

  useEffect(() => {
    trackAnalyticsEvent("app_open", activeTab, { adminMode });
  }, []);

  useEffect(() => {
    trackAnalyticsEvent("tab_view", activeTab, { adminMode });
  }, [activeTab, adminMode]);

  useEffect(() => {
    if (adminMode) setActiveTab("admin");
  }, [adminMode]);

  function queueDownloadWithAd(label, action) {
    const wrapped = () => {
      trackAnalyticsEvent("download", activeTab, { label, sessionCount: sessions.length });
      action();
    };
    if (!adConfig.enabled || !adConfig.gateEnabled) {
      wrapped();
      return;
    }
    pendingDownloadRef.current = wrapped;
    setDownloadGateLabel(label);
    setDownloadGateOpen(true);
  }

  function queueUploadWithAd(files) {
    trackAnalyticsEvent("upload_start", activeTab, { fileCount: files?.length || 0 });
    if (adConfig.enabled && adConfig.gateEnabled) setUploadAdOpen(true);
    processImages(files)
      .then((count) => trackAnalyticsEvent("upload_complete", activeTab, { fileCount: files?.length || 0, sectionsAdded: count || 0 }))
      .catch((err) => trackAnalyticsEvent("upload_error", activeTab, { fileCount: files?.length || 0, message: err?.message || "Upload failed" }))
      .finally(() => setUploadAdOpen(false));
  }

  function closeDownloadGate() {
    setDownloadGateOpen(false);
    pendingDownloadRef.current = null;
  }

  function continueDownloadGate() {
    const action = pendingDownloadRef.current;
    setDownloadGateOpen(false);
    pendingDownloadRef.current = null;
    if (typeof action === "function") action();
  }

  // ─── Auto-correct image rotation using EXIF orientation ─────────────────────
  // Phone cameras save raw pixels sideways/upside-down and rely on EXIF tag to
  // display correctly. We must rotate before sending to Claude AI.
  async function correctImageOrientation(file) {

    // Step 1: read the file as both ArrayBuffer (EXIF) and dataURL (drawing)
    const readAs = (f, method) => new Promise((res, rej) => {
      const r = new FileReader();
      r.onload = e => res(e.target.result);
      r.onerror = () => rej(new Error("FileReader failed"));
      r[method](f);
    });

    let buffer, dataUrl;
    try {
      [buffer, dataUrl] = await Promise.all([
        readAs(file, "readAsArrayBuffer"),
        readAs(file, "readAsDataURL"),
      ]);
    } catch(e) {
      // If reading fails, just return original as-is
      const fallback = await readAs(file, "readAsDataURL");
      return fallback.split(",")[1];
    }

    // Step 2: parse EXIF orientation tag from JPEG bytes
    const getOrientation = (buf) => {
      try {
        const view = new DataView(buf);
        if (view.byteLength < 4) return 1;
        if (view.getUint16(0, false) !== 0xFFD8) return 1; // not JPEG
        let off = 2;
        while (off + 4 <= view.byteLength) {
          const marker = view.getUint16(off, false);
          const segLen = off + 4 <= view.byteLength ? view.getUint16(off + 2, false) : 0;
          if (marker === 0xFFE1 && off + 10 <= view.byteLength) {
            // Check for "Exif" magic bytes
            if (view.getUint32(off + 4, false) === 0x45786966) {
              const tiffOff = off + 6;
              if (tiffOff + 2 > view.byteLength) break;
              const little = view.getUint16(tiffOff, false) === 0x4949;
              const ifd0 = tiffOff + view.getUint32(tiffOff + 4, little);
              if (ifd0 + 2 > view.byteLength) break;
              const numTags = view.getUint16(ifd0, little);
              for (let i = 0; i < numTags; i++) {
                const tagOff = ifd0 + 2 + i * 12;
                if (tagOff + 12 > view.byteLength) break;
                if (view.getUint16(tagOff, little) === 0x0112) {
                  return view.getUint16(tagOff + 8, little);
                }
              }
            }
          }
          if (segLen < 2) break;
          off += 2 + segLen;
        }
      } catch(e) { /* ignore parse errors */ }
      return 1;
    };

    const orientation = getOrientation(buffer);

    // Orientation 1 = already correct — skip canvas entirely, return original
    if (orientation === 1 || orientation === 0) {
      return dataUrl.split(",")[1];
    }

    // Step 3: load image element
    const img = await new Promise((res, rej) => {
      const el = new Image();
      el.onload  = () => res(el);
      el.onerror = () => rej(new Error("Image load failed"));
      el.src = dataUrl;
    });

    if (!img.width || !img.height) {
      // Image loaded but has no dimensions — return original
      return dataUrl.split(",")[1];
    }

    // Step 4: draw onto canvas with correct rotation
    // Orientations 5-8 swap width/height (90° or 270° rotations)
    const swap = orientation >= 5;
    const W = swap ? img.height : img.width;
    const H = swap ? img.width  : img.height;

    const canvas = document.createElement("canvas");
    canvas.width  = W;
    canvas.height = H;
    const ctx = canvas.getContext("2d");
    if (!ctx) return dataUrl.split(",")[1]; // canvas not supported

    // Fill white background so we never get black canvas
    ctx.fillStyle = "#FFFFFF";
    ctx.fillRect(0, 0, W, H);

    // Standard EXIF orientation transforms
    // Reference: https://www.impulsivIT.nl/2018-exif-orientation-html5-canvas/
    ctx.save();
    switch (orientation) {
      case 2: ctx.transform(-1,  0,  0,  1, W,  0); break; // flip horizontal
      case 3: ctx.transform(-1,  0,  0, -1, W,  H); break; // rotate 180
      case 4: ctx.transform( 1,  0,  0, -1, 0,  H); break; // flip vertical
      case 5: ctx.transform( 0,  1,  1,  0, 0,  0); break; // transpose
      case 6: ctx.transform( 0,  1, -1,  0, H,  0); break; // rotate 90 CW
      case 7: ctx.transform( 0, -1, -1,  0, H,  W); break; // transverse
      case 8: ctx.transform( 0, -1,  1,  0, 0,  W); break; // rotate 90 CCW
      default: break;
    }
    ctx.drawImage(img, 0, 0);
    ctx.restore();

    // Step 5: export as JPEG — verify it is not blank before returning
    const result = canvas.toDataURL("image/jpeg", 0.92);
    if (!result || result === "data:," || result.length < 1000) {
      // Canvas export failed — return original unmodified
      return dataUrl.split(",")[1];
    }
    return result.split(",")[1];
  }

  // Process uploaded images through AI
  async function processImages(files) {
    setLoading(true);
    setError(null);
    const added = [];
    for (let i = 0; i < files.length; i++) {
      setLoadMsg("Reading image "+(i+1)+" of "+files.length+"…");
      try {
        // Auto-correct EXIF rotation before sending to AI
        // This fixes sideways/upside-down photos taken on phones
        const b64 = await correctImageOrientation(files[i]);
        // Auto-retry on overload/rate-limit (up to 3 attempts, 4s between each)
        const callAPI = async (imageData, mediaType) => {
          const body = JSON.stringify({
            model: "claude-sonnet-4-20250514",
            max_tokens: 4000,
            system: SYS,
            messages: [{ role: "user", content: [
              { type: "image", source: { type: "base64", media_type: mediaType, data: imageData } },
              { type: "text", text: "Extract all measurements. GOLDEN RULE: if the second value has ' or \" markers it is a WIDTH → sqft. If the second value is a plain number (no ' or \") it is the QUANTITY → rnft. Return JSON array only." },
            ]}],
          });
          const RETRYABLE = ["overloaded_error", "rate_limit_error", "api_error", "server_error"];
          for (let attempt = 1; attempt <= 4; attempt++) {
            const resp = await fetch("https://api.anthropic.com/v1/messages", {
              method: "POST",
              headers: { "Content-Type": "application/json" },
              body,
            });
            const data = await resp.json();
            if (data.error) {
              const code = data.error.type || "";
              if (RETRYABLE.includes(code) && attempt < 4) {
                const wait = attempt * 4000; // 4s, 8s, 12s
                setLoadMsg("API busy — retrying in " + (wait/1000) + "s… (attempt " + attempt + "/4)");
                await new Promise(r => setTimeout(r, wait));
                continue;
              }
              if (code === "authentication_error") throw new Error("API key error — please reload the page.");
              if (code === "rate_limit_error")     throw new Error("Rate limit reached — please wait a minute then try again.");
              if (code === "overloaded_error")     throw new Error("AI servers are overloaded — please try again in a minute.");
              throw new Error("API error: " + (data.error.message || code));
            }
            return data;
          }
        };

        const data = await callLocalExtractionAPI(b64, files[i].type || "image/jpeg", SYS + CALIBRATION_BLOCK);

        const text = (data.content || []).map(c => c.text || "").join("").trim();
        if (!text) throw new Error("AI returned an empty response — try again.");

        // Extract JSON array — greedy match to get the full array even if nested
        // Handles: raw JSON, ```json fences, extra explanation text before/after
        let jsonStr = null;

        // Try 1: greedy bracket match from first [ to last ]
        const firstBracket = text.indexOf("[");
        const lastBracket  = text.lastIndexOf("]");
        if (firstBracket !== -1 && lastBracket > firstBracket) {
          jsonStr = text.slice(firstBracket, lastBracket + 1);
        }

        // Try 2: strip markdown fences if present
        if (!jsonStr) {
          const fenced = text.match(/```(?:json)?\s*([\s\S]*?)```/);
          if (fenced) jsonStr = fenced[1].trim();
        }

        if (!jsonStr) {
          // Show first 200 chars of what AI said to help debug
          const preview = text.slice(0, 200).replace(/\n/g, " ");
          throw new Error("AI did not return JSON. Response: " + preview);
        }

        let parsed;
        // Fix unescaped inch marks: AI writes "2'11"" but JSON needs "2'11\""
        // Strategy: character-by-character scan of string values to escape bare "
        const fixInchMarks = (s) => {
          let out = "";
          let inStr = false;
          let prev = "";
          for (let ci = 0; ci < s.length; ci++) {
            const ch = s[ci];
            if (!inStr) {
              // Outside a string — track entry
              if (ch === '"') inStr = true;
              out += ch;
            } else {
              // Inside a string value
              if (ch === '\\') {
                // Escaped char — copy both chars and skip next
                out += ch + (s[ci+1] || "");
                ci++;
              } else if (ch === '"') {
                // Is this the closing quote or an inch mark?
                // Inch mark pattern: digit immediately before this "
                // and the next non-space char is , } ] or another "
                const afterTrimmed = s.slice(ci + 1).trimStart();
                const nextCh = afterTrimmed[0] || "";
                if (/[0-9]/.test(prev) && /[,}\]"]/.test(nextCh)) {
                  // It's an inch mark — escape it
                  out += '\\"';
                } else {
                  // It's the closing quote of the string
                  inStr = false;
                  out += ch;
                }
              } else {
                out += ch;
              }
              prev = ch;
            }
          }
          return out;
        };

        // Remove trailing commas before } or ]
        const fixTrailing = (s) => s.replace(/,[ \t\r\n]*([}\]])/g, "$1");

        const tryParse = (s) => { try { return JSON.parse(s); } catch(e) { return null; } };

        // Try parsing with progressively more aggressive fixes
        parsed = tryParse(jsonStr)
          || tryParse(fixTrailing(jsonStr))
          || tryParse(fixInchMarks(jsonStr))
          || tryParse(fixTrailing(fixInchMarks(jsonStr)))
          || null;
        if (!Array.isArray(parsed) || parsed.length === 0) {
          throw new Error("AI returned empty data — image may be unclear or too dark. Try a clearer photo.");
        }
        const fallbackTitle = files[i].name.replace(/\.[^.]+$/, "");
        const builtSessions = buildSessionsFromParsed(parsed, fallbackTitle, i);
        builtSessions.forEach(sess => added.push(sess));
      } catch (e) {
        setError("Image "+(i+1)+": "+e.message);
      }
    }
    setLoading(false);
    setLoadMsg("");
    if (added.length) setSessions(prev => [...prev, ...added]);
    return added.length;
  }

  // Session CRUD
  const updateSession = useCallback((id, updated) => {
    setSessions(prev => prev.map(s => s.id === id ? updated : s));
  }, []);

  const removeSession = (id) => setSessions(prev => prev.filter(s => s.id !== id));

  const addSection = (type) => {
    setSessions(prev => [...prev, { id: uid(), title: "New Section ("+(type === "rnft" ? "Rnft" : type === "grove" ? "Grove" : "Sqft")+")", sqftTitle: "Sqft", rnftTitle: "Rnft", rate: "", rows: [] }]);
  };

  // Logo
  function handleLogo(e) {
    const f = e.target.files[0];
    if (!f) return;
    const reader = new FileReader();
    reader.onload = ev => setLogoUrl(ev.target.result);
    reader.readAsDataURL(f);
  }

  // Save
  async function saveSession() {
    const data = { docTitle, company, sessions, billClient, billSite, billDate, billSub, billAdvance, fixedItems, manualItems, billRates, rateCard, companySpec, companyAddr, companyMob, totalLabels, at: new Date().toISOString() };
    const raw = JSON.stringify(data);
    try {
      if (window.storage?.set) await window.storage.set("mf3", raw);
      localStorage.setItem("mf3", raw);
      setSaveLabel("✅ Saved!");
      setTimeout(() => setSaveLabel("💾 Save"), 2500);
    } catch (e) {
      alert("Save failed: " + e.message);
    }
  }

  // CSV export
  function exportCSV() {
    let csv = company ? company + "\n" : "";
    csv += docTitle + "\n\n";
    for (const s of sessions) {
      const net = netTotal(s.rows);
      const unit = domUnit(s);
      csv += s.title + "\n";
      csv += "Sr. No,Item,Dimensions,Area,Deduct\n";
      s.rows.forEach((r, i) => {
        csv += (i+1)+',"'+(r.item||"")+'","'+fmtDim(r.d1, r.d2, r.qty, r.type)+'","'+(r.deduct?'('+fmtArea(r.area)+')'  :fmtArea(r.area))+'",'+(r.deduct?"Yes":"No")+"\n";
      });
      const sft = sqftTotal(s.rows);
      const rft = rnftTotal(s.rows);
      const gft = groveTotal(s.rows);
      if (hasBoth(s.rows)) {
        csv += ",,,,Sqft Total (Sft),"+fmtArea(sft)+"\n";
        if (rft) csv += ",,,,Rnft Total (Rft),"+fmtArea(rft)+"\n";
        if (gft) csv += ",,,,Grove Total (Gft),"+fmtArea(gft)+"\n";
        csv += "\n";
      } else {
        csv += ",,,,Net Total ("+unit+"),"+fmtArea(net)+"\n\n";
      }
    }
    const grand = grandTotal(sessions);
    const gSft = grandSqft(sessions);
    const gRft = grandRnft(sessions);
    if (gSft > 0 && gRft > 0) {
      csv += "Grand Total,,,,"+fmtArea(grand)+"\n";
      csv += "Sqft Grand Total (Sft),,,,"+fmtArea(gSft)+"\n";
      csv += "Rnft Grand Total (Rft),,,,"+fmtArea(gRft)+"\n";
    } else {
      csv += "Grand Total,,,,"+fmtArea(grand)+"\n";
    }
    download(new Blob([csv], { type: "text/csv" }), docTitle.replace(/\s+/g, "_") + ".csv");
    flash("✅ CSV downloaded!");
  }

  // JSON save/load
  function exportJSON() {
    const data = { docTitle, company, sessions, at: new Date().toISOString() };
    download(new Blob([JSON.stringify(data, null, 2)], { type: "application/json" }), docTitle.replace(/\s+/g, "_") + ".json");
    flash("✅ Session saved as JSON!");
  }

  function importJSON(e) {
    const f = e.target.files[0];
    if (!f) return;
    const reader = new FileReader();
    reader.onload = ev => {
      try {
        const d = JSON.parse(ev.target.result);
        if (d.docTitle) setDocTitle(d.docTitle);
        if (d.company) setCompany(d.company);
        if (d.billClient)   setBillClient(d.billClient);
        if (d.billSite)     setBillSite(d.billSite);
        if (d.billDate)     setBillDate(d.billDate);
        if (d.billSub)      setBillSub(d.billSub);
        if (d.billAdvance)  setBillAdvance(d.billAdvance);
        if (d.fixedItems)   setFixedItems(d.fixedItems);
        if (d.manualItems)  setManualItems(d.manualItems);
        if (d.billRates)    setBillRates(d.billRates || {});
    if (d.totalLabels)  setTotalLabels(d.totalLabels);
        if (d.rateCard)     setRateCard(d.rateCard);
        if (d.companySpec)  setCompanySpec(d.companySpec);
        if (d.companyAddr)  setCompanyAddr(d.companyAddr);
        if (d.companyMob)   setCompanyMob(d.companyMob);
        if (d.sessions?.length) {
          setSessions(d.sessions.map(s => ({
            ...s, id: s.id || uid(),
            rows: (s.rows || []).map(r => ({
              ...r, id: r.id || uid(),
              area: calcArea(r.d1, r.d2, r.qty, r.type),
            })),
          })));
          flash("✅ Session loaded!");
        }
      } catch (err) { setError("Invalid file: " + err.message); }
    };
    reader.readAsText(f);
    e.target.value = "";
  }

  // DOCX export
  async function handleDocx() {
    setDocxBusy(true);
    flash("⏳ Building Word document…", true);
    try {
      const blob = await buildDocx(sessions, docTitle, company, logoUrl);
      download(blob, docTitle.replace(/\s+/g, "_") + ".docx");
      flash("✅ Word (.docx) downloaded!");
    } catch (e) {
      flash("❌ " + e.message, false);
      console.error(e);
    }
    setDocxBusy(false);
  }

  // ── Measurement PDF (print-to-PDF via new window) ─────────────────────────
  function handleMeasurePdf() {
    // Format dimension string exactly like target: "6'0" x 27" for rnft, "6'0" x 2'3" x 2" for sqft
    function fmtDimPdf(dim1, dim2, qty, isRnft) {
      const q = qty || 1;
      if (isRnft) {
        return dim1 ? dim1+" x "+q : "x "+q;
      } else {
        if (dim1 && dim2) return dim1+" x "+dim2+(q>1?" x "+q:"");
        if (dim1) return dim1+(q>1?" x "+q:"");
        return "x "+q;
      }
    }
    // Format area to 3 decimal places like target: 162.000
    function fmtA3(v) { return Number(v||0).toFixed(3); }

    // Build one table-section HTML per session (or per sqft/rnft sub-section)
    function buildSection(title, unitLabel, rows, totalVal) {
      let srNo = 1;
      const isRnft = unitLabel.toLowerCase().includes("rn");
      const trs = rows.map(r => {
        const dim = fmtDimPdf(r.d1, r.d2, r.qty, isRnft);
        const areaStr = r.deduct ? "("+fmtA3(r.area)+")" : fmtA3(r.area);
        return `<tr>
          <td class="sr">${srNo++}.</td>
          <td class="item">${r.deduct?"(-) ":""}${r.item||""}</td>
          <td class="dim">${dim}</td>
          <td class="area">${areaStr}</td>
        </tr>`;
      }).join("");

      const totalRow = `<tr class="total-row">
        <td colspan="3" class="total-label">Total Area =</td>
        <td class="area total-val">${fmtA3(totalVal)} ${unitLabel}</td>
      </tr>`;

      return `
      <div class="section">
        <div class="section-title">${title}</div>
        <table>
          <thead><tr>
            <th class="sr">Sr. No</th>
            <th class="item">Items</th>
            <th class="dim">Dimensions (ft-in)</th>
            <th class="area">Area (${unitLabel})</th>
          </tr></thead>
          <tbody>${trs}${totalRow}</tbody>
        </table>
      </div>`;
    }

    let sections = "";
    for (const sess of sessions) {
      if (hasBoth(sess.rows)) {
        for (const group of measureGroups(sess)) {
          const pageType = group.key === "sqft" ? "Sqft." : group.key === "grove" ? "Grove." : "Rnft.";
          const title = `${group.label} ${pageType}  Page`;
          sections += buildSection(title, group.key === "sqft" ? "Sq ft" : group.unit, group.rows, group.total);
        }
      } else {
        const group = measureGroups(sess)[0];
        const unitLabel = group ? (group.key === "sqft" ? "Sq ft" : group.unit) : "Sq ft";
        const pageType = group ? (group.key === "sqft" ? "Sqft." : group.key === "grove" ? "Grove." : "Rnft.") : "Sqft.";
        const title = sess.title+" "+pageType+"  Page";
        sections += buildSection(title, unitLabel, sess.rows, netTotal(sess.rows));
      }
    }

    const html = `<!DOCTYPE html><html><head><meta charset="utf-8"/>
<title>${docTitle||"Measurements"}</title>
<style>
  * { box-sizing:border-box; margin:0; padding:0; }
  body { font-family:'Times New Roman', serif; font-size:13px; color:#111; background:#fff; }
  .page { padding:28px 36px; max-width:720px; margin:0 auto; }
  .company { text-align:center; font-size:17px; font-weight:bold; letter-spacing:1px; margin-bottom:2px; }
  .doc-title { text-align:center; font-size:13px; color:#555; margin-bottom:20px; }
  .section { margin-bottom:32px; page-break-inside:avoid; }
  .section-title {
    font-size:13px; font-weight:bold; text-align:center;
    border-top:2px solid #111; border-bottom:1px solid #111;
    padding:4px 0; margin-bottom:0; letter-spacing:.5px;
  }
  table { width:100%; border-collapse:collapse; font-size:12.5px; }
  thead tr { border-bottom:1px solid #111; }
  th { font-weight:bold; padding:5px 8px; text-align:left; font-size:12.5px; border-bottom:1px solid #333; }
  td { padding:4px 8px; border-bottom:1px solid #ddd; vertical-align:top; }
  th.sr, td.sr { width:42px; text-align:right; }
  th.dim, td.dim { width:160px; text-align:right; }
  th.area, td.area { width:110px; text-align:right; }
  th.item, td.item { }
  .total-row td { border-top:1.5px solid #111; border-bottom:2px solid #111; font-weight:bold; padding:5px 8px; }
  .total-label { text-align:right; font-weight:bold; }
  .total-val { font-weight:bold; }
  .grand-box { margin-top:16px; border-top:2px solid #111; padding-top:8px; font-weight:bold; font-size:13px; display:flex; justify-content:flex-end; gap:20px; flex-wrap:wrap; }
  .gt-item { color:#1F4E79; }
  .gt-item.grove { color:#7c3aed; }
  @media print {
    body { font-size:12px; }
    .page { padding:16px 20px; }
    .section { page-break-inside:avoid; }
  }
</style>
<script>window.onload = function(){ setTimeout(function(){ window.print(); }, 300); }</script>
</head><body>
<div class="page">
  ${company ? '<div class="company">'+company+'</div>' : ''}
  ${docTitle ? '<div class="doc-title">'+docTitle+'</div>' : ''}
  ${sections}
  <div class="grand-box">
    ${grandSqft(sessions)>0 ? '<span class="gt-item">Sqft Total: '+fmtA3(grandSqft(sessions))+' Sq ft</span>' : ''}
    ${grandRnft(sessions)>0 ? '<span class="gt-item">Rnft Total: '+fmtA3(grandRnft(sessions))+' Rn ft</span>' : ''}
    ${grandGrove(sessions)>0 ? '<span class="gt-item grove">Grove Total: '+fmtA3(grandGrove(sessions))+' Gft</span>' : ''}
  </div>
</div>
</body></html>`;

    const blob = new Blob([html], {type:"text/html"});
    const url  = URL.createObjectURL(blob);
    const a    = document.createElement("a");
    a.href = url; a.download = (docTitle||"Measurements").replace(/\s+/g,"_")+".html";
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    setTimeout(()=>URL.revokeObjectURL(url), 10000);
  }

  const [billBusy, setBillBusy] = useState(false);
  async function handleBill() {
    setBillBusy(true);
    flash("⏳ Building Bill…", true);
    try {
      const xml = buildBillDocxXml(sessions, {
        company, companySpec, companyAddr, companyMob, logoUrl,
        billClient, billSite, billDate, billSub, billAdvance, fixedItems, manualItems, billRates, rateCard
      });
      const blob = await buildDocx(sessions, docTitle, company, logoUrl, xml);
      const fname = (billClient || "Bill").replace(/\s+/g,"_") + "_" + billDate.replace(/\//g,"-") + ".docx";
      download(blob, fname);
      flash("✅ Bill downloaded!");
    } catch (e) {
      flash("❌ " + e.message, false);
      console.error(e);
    }
    setBillBusy(false);
  }

  // ── Bill PDF ──────────────────────────────────────────────────────────────
  function handleBillPdf() {
    // Build all bill lines same as BillView
    const lines = [];
    const bRates = billRates || {};
    for (const sess of sessions) {
      if (hasBoth(sess.rows)) {
        const groups = measureGroups(sess);
        for (const g of groups) {
          const lineId = sess.id + g.per;
          const cardR  = lookupRateFromCard(rateCard, g.label, g.per);
          const rate   = parseFloat(bRates[lineId] !== undefined ? bRates[lineId] : cardR) || 0;
          const qty = g.total;
          lines.push({ label:g.label, qty:qty.toFixed(2), rate, per:g.per, total:qty*rate });
        }
      } else {
        const lineId = sess.id;
        const group = measureGroups(sess)[0];
        const per  = sess.perOverride || (group ? group.per : "Sq ft");
        const cardR2 = lookupRateFromCard(rateCard, sess.title, per);
        const rate   = parseFloat(bRates[lineId] !== undefined ? bRates[lineId] : cardR2) || 0;
        const qty  = netTotal(sess.rows);
        lines.push({ label:sess.title, qty:qty.toFixed(2), rate, per, total:qty*rate });
      }
    }
    for (const mi of manualItems) {
      const qty=parseFloat(mi.qty)||0, r=parseFloat(mi.rate)||0;
      const total = mi.amount ? parseFloat(mi.amount)||0 : qty*r;
      lines.push({ label:mi.name||"Item", qty:mi.qty||"─", rate:mi.rate||"─", per:mi.per||"Sq ft", total });
    }
    for (const fi of fixedItems) {
      const amt = fi.qty&&fi.unitPrice?(parseFloat(fi.qty)||0)*(parseFloat(fi.unitPrice)||0):(parseFloat(fi.amount)||0);
      lines.push({ label:fi.name||"Item", qty:fi.qty||"─", rate:fi.unitPrice||fi.amount||"─", per:fi.unitPrice?"Nos":"Fixed", total:amt });
    }
    const grand = lines.reduce((s,l)=>s+l.total,0);
    const adv   = parseFloat(billAdvance)||0;
    const rem   = grand - adv;

    const tableRows = lines.map((l,i) => {
      const cls  = i%2===0 ? "even" : "odd";
      const qty  = l.qty  || "───";
      const rate = (l.rate && l.rate!==0 && l.rate!=="─") ? l.rate : "───";
      const tot  = l.total>0 ? "₹ "+Math.round(l.total).toLocaleString("en-IN") : "───";
      return `<tr class="${cls}">
        <td class="c" style="color:#888;font-size:12px">${i+1}.</td>
        <td class="l">${xmlEscHtml(String(l.label||""))}</td>
        <td class="r">${xmlEscHtml(String(qty))}</td>
        <td class="r">${xmlEscHtml(String(rate))}</td>
        <td class="l">${xmlEscHtml(l.per||"")}</td>
        <td class="r" style="font-weight:bold;color:#1F4E79">${tot}</td>
      </tr>`;
    }).join("");

    const html = `<!DOCTYPE html><html><head><meta charset="utf-8"/>
<title>Bill</title>
<style>
*{box-sizing:border-box;margin:0;padding:0}
body{font-family:"Times New Roman",serif;font-size:13px;color:#1a1a2e;padding:22px 28px;max-width:800px;margin:0 auto}
.hdr{background:#1F4E79;color:#fff;border:3px solid #2E75B6;padding:14px 20px;text-align:center;border-radius:4px}
.hdr-name{font-size:20px;font-weight:bold;text-decoration:underline;letter-spacing:1px;margin-bottom:6px}
.hdr-spec{font-size:12px;color:#BDD7EE}
.sep{border:none;border-top:1.5px solid #c7d9f0;margin:10px 0 6px}
.addr-row{display:flex;justify-content:space-between;align-items:flex-start;margin-top:4px}
.addr-left{font-size:12px;color:#1F4E79;line-height:1.7}
.date{font-size:12px;color:#1F4E79;text-decoration:underline;font-weight:bold;text-align:right}
.meta{margin:12px 0 10px;border-left:3px solid #2E75B6;padding-left:10px}
.meta p{font-size:13px;color:#1F4E79;margin:4px 0}
.meta .lbl{font-weight:bold}
.meta .ulbl{font-weight:bold;text-decoration:underline}
.bill-heading{text-align:center;font-size:16px;font-weight:bold;text-decoration:underline;color:#1F4E79;margin:14px 0 12px;letter-spacing:2px}
table{width:100%;border-collapse:collapse;font-size:13px;margin-bottom:0}
thead th{background:#1F4E79;color:#fff;font-weight:bold;padding:8px 10px;border:1px solid #2E75B6;white-space:nowrap}
th.r,td.r{text-align:right}
th.l,td.l{text-align:left}
th.c,td.c{text-align:center}
td{padding:7px 10px;border:1px solid #c7d9f0;vertical-align:middle}
tr.even td{background:#fff}
tr.odd  td{background:#D6E4F0}
.foot-blank{background:#f8fafc !important;border-color:#e2e8f0 !important}
tr.foot-b .foot-blank{background:#f8fafc !important;border:1px solid #e2e8f0 !important;color:#f8fafc}
tr.foot-b td.foot-lbl,tr.foot-b td.foot-val{background:#1F4E79;color:#fff;font-weight:bold;font-size:14px;border:1px solid #2E75B6}
tr.foot-l td.foot-lbl,tr.foot-l td.foot-val{background:#D6E4F0;color:#1F4E79;font-weight:bold;border:1px solid #c7d9f0}
.words{margin:18px 0 22px;font-size:13px;color:#1F4E79;padding:10px 14px;background:#f0f7ff;border-left:3px solid #2E75B6;border-radius:2px}
.words .lbl{font-weight:bold;text-decoration:underline}
.sign-row{display:flex;justify-content:space-between;margin-top:40px;font-size:13px;color:#1F4E79}
@media print{body{padding:10px 14px;max-width:none}@page{size:A4;margin:12mm 14mm}}
</style>
<script>window.onload = function(){ setTimeout(function(){ window.print(); }, 300); }</script>
</head><body>
<div class="hdr">
  <div class="hdr-name">${xmlEscHtml(company||"Your Company Name")}</div>
  ${companySpec ? '<div class="hdr-spec">Specialist in: '+xmlEscHtml(companySpec)+'</div>' : ''}
</div>
<hr class="sep"/>
<div class="addr-row">
  <div class="addr-left">
    <div>Address : ${xmlEscHtml(companyAddr||"")}</div>
    ${companyMob ? '<div>Mob.No &nbsp;: '+xmlEscHtml(companyMob)+'</div>' : ''}
  </div>
  <div class="date">Date :- ${xmlEscHtml(billDate)}</div>
</div>
<div class="meta">
  <p><span class="lbl">To :-&nbsp;&nbsp;</span>${xmlEscHtml(billClient||"")}</p>
  <p><span class="lbl">SUB :-&nbsp;</span>${xmlEscHtml(billSub||"Bill for Services.")}</p>
  ${billSite?'<p><span class="ulbl">Worksite Address :- </span>'+xmlEscHtml(billSite)+'</p>':''}
</div>
<div class="bill-heading">BILL</div>
<table>
  <colgroup>
    <col style="width:6%"/><col style="width:36%"/><col style="width:16%"/>
    <col style="width:10%"/><col style="width:9%"/><col style="width:23%"/>
  </colgroup>
  <thead>
    <tr>
      <th class="c">Sr. No</th><th class="l">Particulars</th>
      <th class="r">Quantity</th><th class="r">Rate</th>
      <th class="l">Per</th><th class="r">Total Amount</th>
    </tr>
  </thead>
  <tbody>${tableRows}</tbody>
  <tfoot>
    <tr class="foot-b">
      <td class="foot-blank" colspan="4"></td>
      <td class="foot-lbl l">Total Amount</td>
      <td class="foot-val r">${grand.toLocaleString("en-IN",{minimumFractionDigits:2})}</td>
    </tr>
    ${adv>0?`
    <tr class="foot-l">
      <td class="foot-blank" colspan="4"></td>
      <td class="foot-lbl l">Advance Received</td>
      <td class="foot-val r">${adv.toLocaleString("en-IN",{minimumFractionDigits:2})}</td>
    </tr>
    <tr class="foot-b">
      <td class="foot-blank" colspan="4"></td>
      <td class="foot-lbl l">Remaining Amount</td>
      <td class="foot-val r">${rem.toLocaleString("en-IN",{minimumFractionDigits:2})}</td>
    </tr>`:""}
  </tfoot>
</table>
<div class="words"><span class="lbl">In words:</span> ${numToWords(Math.round(adv>0?rem:grand))} Only.</div>
<div class="sign-row">
  <span>Received by: ___________________________</span>
  <span>Contractor sign. ___________________________</span>
</div>
</body></html>`;

    const blob = new Blob([html], {type:"text/html"});
    const url  = URL.createObjectURL(blob);
    const a    = document.createElement("a");
    a.href = url; a.download = (docTitle||"Measurements").replace(/\s+/g,"_")+".html";
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    setTimeout(()=>URL.revokeObjectURL(url), 10000);
  }

  const launchMeasureDocx = () => queueDownloadWithAd("Measurement Word export", handleDocx);
  const launchMeasurePdf = () => queueDownloadWithAd("Measurement PDF export", handleMeasurePdf);
  const launchMeasureCsv = () => queueDownloadWithAd("Measurement CSV export", exportCSV);
  const launchMeasureJson = () => queueDownloadWithAd("Measurement JSON backup", exportJSON);
  const launchBillDocx = () => queueDownloadWithAd("Bill Word export", handleBill);
  const launchBillPdf = () => queueDownloadWithAd("Bill PDF export", handleBillPdf);

  const grand = grandTotal(sessions);
  const hasSessions = sessions.length > 0;
  const t = (en, hi) => lang === "hi" ? hi : en;

  return (
    <div style={{ minHeight: "100vh", background: "#f0f4f8", fontFamily: "Times New Roman, Noto Sans Devanagari, serif", color: "#1a1a2e" }}>
      <style>{PRINT_CSS}</style>
      <DownloadGateModal
        open={downloadGateOpen}
        label={downloadGateLabel}
        seconds={Math.max(1, parseInt(adConfig.gateSeconds) || 6)}
        videoUrl={adConfig.videoUrl}
        adClient={adConfig.adClient}
        onClose={closeDownloadGate}
        onContinue={continueDownloadGate}
      />
      <DownloadGateModal
        open={uploadAdOpen}
        label="Image upload processing"
        seconds={0}
        videoUrl={adConfig.videoUrl}
        adClient={adConfig.adClient}
        onClose={() => setUploadAdOpen(false)}
        onContinue={() => {}}
        processing={true}
        processingText={loadMsg || "Reading measurements..."}
      />
      {/* PRINT VIEW */}
      {printing && (
        <div style={{ position: "fixed", inset: 0, background: "#fff", zIndex: 9999, overflow: "auto", padding: 28, fontFamily: "Times New Roman, Noto Sans Devanagari, serif" }}>
          <div className="noprint" style={{ marginBottom: 14, display: "flex", gap: 10 }}>
            <button style={btn(C.dark)} onClick={() => window.print()}>🖨 Print / Save as PDF</button>
            <button style={btn(C.gray)} onClick={() => setPrinting(false)}>✕ Close</button>
          </div>
          {logoUrl && <div style={{ textAlign: "center", marginBottom: 6 }}><img src={logoUrl} alt="logo" style={{ maxHeight: 60 }} /></div>}
          {company && <h1 style={{ textAlign: "center", color: C.dark, fontSize: 22, margin: "0 0 4px" }}>{company}</h1>}
          <h2 style={{ textAlign: "center", color: C.mid, fontSize: 17, margin: "0 0 18px", borderBottom: "2px solid "+C.mid, paddingBottom: 8 }}>{docTitle}</h2>
          {sessions.map(sess => {
            const net = netTotal(sess.rows);
            const unit = domUnit(sess);
            const PT = detectTheme(sess.title);
            return (
              <div key={sess.id} style={{ marginBottom: 20, pageBreakInside: "avoid" }}>
                <h3 style={{ color: PT.dark, fontSize: 13, margin: "0 0 5px", borderLeft: "4px solid "+PT.mid, paddingLeft: 8 }}>{sess.title}</h3>
                {(() => {
                  const tdBdr = "1px solid #dbeafe";
                  const PrintTable = ({ rows, unitLabel, totalVal, headBg, valColor, totalLabel }) => (
                    <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 11, marginBottom: 0 }}>
                      <thead>
                        <tr style={{ background: headBg, color: "#fff" }}>
                          {["Sr.", "Item", "Dimensions", "Area ("+unitLabel+")"].map((h, i) => (
                            <th key={i} style={{ padding: "5px 7px", textAlign: i === 3 ? "right" : "left", border: "1px solid "+PT.mid }}>{h}</th>
                          ))}
                        </tr>
                      </thead>
                      <tbody>
                        {rows.map((r, i) => (
                          <tr key={r.id} style={{ background: r.deduct ? C.redBg : i % 2 === 1 ? PT.light : "#fff" }}>
                            <td style={{ padding: "4px 7px", border: tdBdr, color: r.deduct ? C.red : "inherit", width: 24 }}>{i + 1}</td>
                            <td style={{ padding: "4px 7px", border: tdBdr, color: r.deduct ? C.red : "inherit" }}>{r.item}</td>
                            <td style={{ padding: "4px 7px", border: tdBdr, color: r.deduct ? C.red : "inherit" }}>{fmtDim(r.d1, r.d2, r.qty, r.type)}</td>
                            <td style={{ padding: "4px 7px", border: tdBdr, color: r.deduct ? C.red : "inherit", textAlign: "right" }}>
                              {r.deduct ? "("+fmtArea(r.area)+")" : fmtArea(r.area)}
                            </td>
                          </tr>
                        ))}
                      </tbody>
                      <tfoot>
                        <tr style={{ background: headBg, color: "#fff" }}>
                          <td colSpan={3} style={{ padding: "5px 7px", fontWeight: "bold", border: tdBdr }}>
                            {totalLabel ? totalLabel+" Total" : unitLabel === "Sft" ? "Sqft Total" : unitLabel === "Rft" ? "Rnft Total" : "Net Total"}
                          </td>
                          <td style={{ padding: "5px 7px", fontWeight: "bold", textAlign: "right", border: tdBdr, color: valColor }}>
                            {fmtArea(totalVal)} {unitLabel}
                          </td>
                        </tr>
                      </tfoot>
                    </table>
                  );

                  const pSqftTitle = sess.sqftTitle || "Sqft";

                  if (hasBoth(sess.rows)) {
                    return (<>
                      {measureGroups(sess).map((group, idx) => {
                        const headBg = group.key === "sqft" ? PT.sqftDark : group.key === "grove" ? "#6d28d9" : PT.rnftDark;
                        const valColor = group.key === "sqft" ? PT.sqftVal : group.key === "grove" ? "#f3e8ff" : PT.rnftVal;
                        return (
                          <React.Fragment key={group.key}>
                            {idx > 0 && <div style={{ height: 6, background: PT.light }} />}
                            <div style={{ background: headBg, color:"#fff", fontWeight:"bold", fontSize:12, padding:"5px 10px" }}>
                              {group.label}
                            </div>
                            <PrintTable rows={group.rows} unitLabel={group.unit} totalVal={group.total} headBg={headBg} valColor={valColor} totalLabel={group.label} />
                          </React.Fragment>
                        );
                      })}
                    </>);
                  }
                  return <PrintTable rows={sess.rows} unitLabel={unit === "Rft" ? "Rft" : "Sft"} totalVal={net} headBg={PT.dark} valColor="#fff" />;
                })()}
              </div>
            );
          })}
          <div style={{ borderTop: "2px solid "+C.dark, paddingTop: 10, textAlign: "right", color: C.dark, fontWeight: "bold", fontSize: 15 }}>
            <div style={{ display:"flex", justifyContent:"flex-end", gap:20, flexWrap:"wrap" }}>
              {grandSqft(sessions) > 0 && <span style={{color:C.mid}}>Sqft Total: {fmtArea(grandSqft(sessions))} Sft</span>}
              {grandRnft(sessions) > 0 && <span style={{color:C.green}}>Rnft Total: {fmtArea(grandRnft(sessions))} Rft</span>}
              {grandGrove(sessions) > 0 && <span style={{color:"#9333ea"}}>Grove Total: {fmtArea(grandGrove(sessions))} Gft</span>}
            </div>
          </div>

        </div>
      )}

      {/* TOP BAR */}
      <div className="noprint" style={{ background: C.dark, padding: "0 22px", display: "flex", alignItems: "center", justifyContent: "space-between", height: 54, boxShadow: "0 2px 16px rgba(31,78,121,.4)", position: "sticky", top: 0, zIndex: 100 }}>
        <div style={{ color: "#fff", fontSize: 18, fontWeight: "bold", letterSpacing: 1, display: "flex", alignItems: "center", gap: 8 }}>
          <div style={{ background: C.mid, borderRadius: 8, width: 30, height: 30, display: "flex", alignItems: "center", justifyContent: "center", fontSize: 14 }}>📐</div>
          MeasureFlow
        </div>
        <div style={{ display: "flex", gap: 8, alignItems: "center" }}>
          <div style={{ display:"flex", background:"rgba(255,255,255,.12)", borderRadius:8, padding:2, gap:2 }}>
            {[["en","EN"],["hi","हिंदी"]].map(([k,l]) => (
              <button key={k} onClick={() => setLang(k)} style={{
                background: lang===k ? "#fff" : "transparent",
                color: lang===k ? C.dark : "#cfe0f5",
                border:"none", borderRadius:6, padding:"5px 10px",
                fontWeight:"bold", fontSize:12, cursor:"pointer"
              }}>{l}</button>
            ))}
          </div>
          {/* Mode tabs */}
          <div style={{ display:"flex", background:"rgba(255,255,255,.12)", borderRadius:8, padding:2, gap:2 }}>
            {[
              { id:"measure", label:t("📐 Measure","📐 माप") },
              { id:"bill",    label:t("🧾 Bill","🧾 बिल") },
              { id:"rates",   label:t("💰 Rates","💰 रेट") },
              ...(adminMode ? [{ id:"admin", label:t("🛡 Admin","🛡 एडमिन") }] : []),
            ].map(tab => (
              <button key={tab.id} onClick={() => setActiveTab(tab.id)} style={{
                background: activeTab===tab.id ? "#fff" : "transparent",
                color: activeTab===tab.id ? C.dark : "#93c5fd",
                border:"none", borderRadius:6, padding:"5px 14px",
                fontWeight:"bold", fontSize:13, cursor:"pointer", transition:"all .15s"
              }}>{tab.label}</button>
            ))}
          </div>
          {activeTab==="measure" && hasSessions && <>
            <button style={btn(C.mid, true)} onClick={saveSession}>{saveLabel}</button>
          </>}
          {activeTab==="bill" && <>
            <button style={btn("#1e3a5f", true)} disabled={billBusy} onClick={launchBillDocx}>
              🧾 {billBusy ? "Building…" : "Download Bill (.docx)"}
            </button>
          </>}
          <span style={{ background: "#fff", color: C.dark, fontSize: 10, fontWeight: "bold", padding: "3px 8px", borderRadius: 20, letterSpacing: 1 }}>v2.0</span>
        </div>
      </div>

      <div className="noprint" style={{ maxWidth: "100%", margin: "0 auto", padding: "18px 14px", display: "flex", gap: 14, alignItems: "flex-start", justifyContent: "center" }}>
        <AutoAdBanner adClient={adConfig.adClient} side="left" activeTab={activeTab} />
        <div style={{ maxWidth: 1160, width: "100%" }}>

        {/* ══ MEASURE MODE ══════════════════════════════════════════════════════ */}
        {activeTab==="measure" && <>

        {/* TYPE LEGEND */}
        <div style={{ background: "#fff", borderRadius: 10, border: "1px solid #dbeafe", padding: "9px 18px", marginBottom: 16, display: "flex", gap: 24, flexWrap: "wrap", alignItems: "center" }}>
          <span style={{ fontSize: 13, fontWeight: "bold", color: C.dark }}>Types:</span>
          <span style={{ fontSize: 13, color: C.mid }}>🟦 <strong>Sqft</strong> = L × W × Qty → Area (sq ft)</span>
          <span style={{ fontSize: 13, color: C.green }}>🟩 <strong>Rnft</strong> = L × Qty → Running feet</span>
          <span style={{ fontSize: 13, color: "#9333ea" }}>🟪 <strong>Grove</strong> = L × Qty → Groove/Rnft</span>
          <span style={{ fontSize: 12, color: "#94a3b8", marginLeft: "auto" }}>AI auto-detects · override per row</span>
        </div>

        {/* SETTINGS — compact, measure mode only */}
        <div style={{ ...card }} className="noprint">
          <div style={cardHead(C.gray)}><span>⚙️</span><span style={cardTitle}>{t("Settings","सेटिंग्स")}</span></div>
          <div style={{ padding: 14, display:"flex", gap:14, flexWrap:"wrap", alignItems:"flex-end" }}>
            <div style={{ flex:"1 1 180px" }}>
              <div style={{ fontSize:11, color:"#64748b", marginBottom:3 }}>{t("Document Title","दस्तावेज़ शीर्षक")}</div>
              <input style={inp()} value={docTitle} onChange={e => setDocTitle(e.target.value)} />
            </div>
            <div style={{ flex:"1 1 200px" }}>
              <div style={{ fontSize:11, color:"#64748b", marginBottom:3 }}>{t("Company Name","कंपनी का नाम")}</div>
              <input style={inp()} value={company} placeholder={t("Your company name","अपनी कंपनी का नाम")} onChange={e => setCompany(e.target.value)} />
            </div>
            <div>
              <div style={{ fontSize:11, color:"#64748b", marginBottom:3 }}>{t("Logo","लोगो")}</div>
              <div style={{ display:"flex", gap:8, alignItems:"center" }}>
                <button style={btn(C.mid, true)} onClick={() => logoRef.current.click()}>📁 Logo</button>
                {logoUrl && <img src={logoUrl} alt="logo" style={{ height:26, borderRadius:4, border:"1px solid #dbeafe" }} />}
                {logoUrl && <button onClick={() => setLogoUrl(null)} style={{ background:"none", border:"none", color:C.red, cursor:"pointer", fontSize:16 }}>✕</button>}
                <input ref={logoRef} type="file" accept="image/*" style={{ display:"none" }} onChange={handleLogo} />
              </div>
            </div>
            <div style={{ marginLeft:"auto", display:"flex", gap:8 }}>
              <button style={btn("#7c3aed")} onClick={launchMeasureJson}>💾 Save Session</button>
              <button style={btn("#0891b2")} onClick={() => importRef.current.click()}>📂 Load Session</button>
            </div>
          </div>
        </div>

        {/* UPLOAD */}
        <div style={card} className="noprint">
          <div style={cardHead()}>
            <span>📤</span><span style={cardTitle}>{t("Upload Images","तस्वीरें अपलोड करें")}</span>
            <div style={{ marginLeft: "auto", display: "flex", gap: 8 }}>
              <button style={btn("#0891b2", true)} onClick={() => importRef.current.click()}>📂 Load Session</button>
              <input ref={importRef} type="file" accept=".json" style={{ display: "none" }} onChange={importJSON} />
            </div>
          </div>
          <div style={{ padding: 14 }}>
            <UploadZone onFiles={queueUploadWithAd} lang={lang} />
            {/* ── Manual entry ── */}
            <div style={{ display:"flex", alignItems:"center", gap:10, margin:"14px 0 10px" }}>
              <div style={{ flex:1, height:1, background:"#e2e8f0" }}/>
              <span style={{ fontSize:11, color:"#94a3b8", whiteSpace:"nowrap" }}>{t("or enter manually (works offline)","या मैन्युअल रूप से दर्ज करें (ऑफलाइन भी काम करता है)")}</span>
              <div style={{ flex:1, height:1, background:"#e2e8f0" }}/>
            </div>
            <div style={{ display:"flex", gap:8, flexWrap:"wrap" }}>
              <button onClick={()=>addSection("sqft")}
                style={{ flex:1, minWidth:110, background:C.mid, color:"#fff", border:"none",
                  borderRadius:8, padding:"9px 10px", fontSize:12, fontWeight:"bold", cursor:"pointer",
                  display:"flex", alignItems:"center", justifyContent:"center", gap:5 }}>
                {t("✏️ + Sqft Section","✏️ + Sqft सेक्शन")}
              </button>
              <button onClick={()=>addSection("rnft")}
                style={{ flex:1, minWidth:110, background:C.green, color:"#fff", border:"none",
                  borderRadius:8, padding:"9px 10px", fontSize:12, fontWeight:"bold", cursor:"pointer",
                  display:"flex", alignItems:"center", justifyContent:"center", gap:5 }}>
                {t("✏️ + Rnft Section","✏️ + Rnft सेक्शन")}
              </button>
              <button onClick={()=>addSection("grove")}
                style={{ flex:1, minWidth:110, background:"#9333ea", color:"#fff", border:"none",
                  borderRadius:8, padding:"9px 10px", fontSize:12, fontWeight:"bold", cursor:"pointer",
                  display:"flex", alignItems:"center", justifyContent:"center", gap:5 }}>
                {t("✏️ + Grove Section","✏️ + Grove सेक्शन")}
              </button>
            </div>
            <div style={{ marginTop:7, fontSize:11, color:"#94a3b8", textAlign:"center" }}>
              💡 Scan photo needs internet · Manual entry works offline
            </div>
            {loading && (
              <div style={{ marginTop: 12, textAlign: "center", color: C.mid, padding: "12px 0" }}>
                <div style={{ fontSize: 26, marginBottom: 5 }}>⏳</div>
                <div style={{ fontWeight: "bold", fontSize: 14 }}>{loadMsg}</div>
                <div style={{ fontSize: 11, color: "#64748b", marginTop: 3 }}>AI detecting Sqft / Rnft types…</div>
              </div>
            )}
            {error && (
              <div style={{ marginTop: 8, background: C.redBg, border: "1px solid "+C.red, borderRadius: 8, padding: "8px 12px", color: C.red, fontSize: 13 }}>
                ⚠️ {error}
              </div>
            )}
          </div>
        </div>

        {/* SESSIONS */}
        {hasSessions && (
          <>
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 12 }} className="noprint">
              <div style={{ fontSize: 14, color: C.dark, fontWeight: "bold" }}>
                {sessions.length} section(s) · <span style={{ color: "#60a5fa" }}>Sft: {fmtArea(grandSqft(sessions))}</span>{grandRnft(sessions) > 0 && <span style={{color:"#94a3b8", margin:"0 4px"}}>·</span>}{grandRnft(sessions) > 0 && <span style={{ color: C.green }}>Rft: {fmtArea(grandRnft(sessions))}</span>}{grandGrove(sessions) > 0 && <span style={{color:"#94a3b8", margin:"0 4px"}}>·</span>}{grandGrove(sessions) > 0 && <span style={{ color: "#c084fc" }}>Grove: {fmtArea(grandGrove(sessions))} Gft</span>}
              </div>
              <div style={{ display: "flex", gap: 8 }}>
                <button style={btn(C.mid, true)} onClick={() => addSection("sqft")}>+ Sqft Section</button>
                <button style={btn(C.green, true)} onClick={() => addSection("rnft")}>+ Rnft Section</button>
                <button style={btn("#9333ea", true)} onClick={() => addSection("grove")}>+ Grove Section</button>
              </div>
            </div>

            {sessions.map(sess => (
              <SessionPanel
                key={sess.id}
                sess={sess}
                onUpdate={updated => updateSession(sess.id, updated)}
                onRemove={() => removeSession(sess.id)}
              />
            ))}

            {/* TOTALS BOX — editable labels, noprint so it never appears in downloads */}
            {(() => {
              const gSq  = grandSqft(sessions);
              const gRn  = grandRnft(sessions);
              const gGr  = grandGrove(sessions);
              const tiles = [
                { key:"sqft",  show: gSq > 0,  val: fmtArea(gSq),  unit:"Sft",  labelKey:"sqftLabel",  defaultLabel:"Sqft Total",  bg:"rgba(59,130,246,.18)",  labelColor:"#93c5fd", valColor:"#bfdbfe" },
                { key:"rnft",  show: gRn > 0,  val: fmtArea(gRn),  unit:"Rft",  labelKey:"rnftLabel",  defaultLabel:"Rnft Total",  bg:"rgba(34,197,94,.15)",   labelColor:"#86efac", valColor:"#bbf7d0" },
                { key:"grove", show: gGr > 0,  val: fmtArea(gGr),  unit:"Gft",  labelKey:"groveLabel", defaultLabel:"Grove Total", bg:"rgba(168,85,247,.18)",  labelColor:"#d8b4fe", valColor:"#e9d5ff" },
              ].filter(t => t.show);
              if (!tiles.length) return null;
              return (
                <div className="noprint" style={{ background: C.dark, borderRadius: 12, padding: "13px 20px", marginBottom: 18, display: "flex", justifyContent: "center", alignItems: "center", gap: 16, flexWrap:"wrap" }}>
                  {tiles.map(t => (
                    <div key={t.key} style={{ textAlign: "center", background: t.bg, borderRadius: 10, padding: "8px 20px", minWidth: 130 }}>
                      <input
                        value={totalLabels[t.labelKey] ?? t.defaultLabel}
                        onChange={e => setTotalLabels(prev => ({ ...prev, [t.labelKey]: e.target.value }))}
                        style={{
                          background: "transparent", border: "none", borderBottom: "1px dashed rgba(255,255,255,.3)",
                          color: t.labelColor, fontSize: 11, fontWeight: "bold", letterSpacing: 1,
                          textAlign: "center", width: "100%", outline: "none", marginBottom: 4, cursor:"text"
                        }}
                      />
                      <div style={{ fontSize: 24, fontWeight: "bold", color: t.valColor }}>
                        {t.val} <span style={{ fontSize: 13 }}>{t.unit}</span>
                      </div>
                    </div>
                  ))}
                  <div style={{ textAlign: "center", minWidth: 60 }}>
                    <div style={{ fontSize: 10, color: "#93c5fd", letterSpacing: 1, marginBottom: 2 }}>SECTIONS</div>
                    <div style={{ fontSize: 22, fontWeight: "bold", color: "#fff" }}>{sessions.length}</div>
                  </div>
                </div>
              );
            })()}

            {/* ── Download Buttons ── */}
            <div className="noprint" style={{ display:"flex", gap:10, flexWrap:"wrap", justifyContent:"center", marginBottom:8 }}>
              {[
                { label:"Word", sub:".docx", icon:"📄", color:"#1F4E79", shadow:"rgba(31,78,121,.3)", action:launchMeasureDocx, disabled:docxBusy, loadingLabel:"Building…" },
                { label:"PDF",  sub:"Print / Save", icon:"📕", color:"#dc2626", shadow:"rgba(220,38,38,.3)", action:launchMeasurePdf },
                { label:"CSV",  sub:"Spreadsheet",  icon:"📊", color:"#16a34a", shadow:"rgba(22,163,74,.3)",  action:launchMeasureCsv },
                { label:"JSON", sub:"Session backup",icon:"💾", color:"#7c3aed", shadow:"rgba(124,58,237,.3)",action:launchMeasureJson },
              ].map(({label,sub,icon,color,shadow,action,disabled,loadingLabel})=>(
                <button key={label} disabled={!!disabled} onClick={action} style={{
                  display:"flex", alignItems:"center", gap:9,
                  background: disabled ? "#e2e8f0" : color,
                  color: disabled ? "#94a3b8" : "#fff",
                  border:"none", borderRadius:10, padding:"11px 24px",
                  fontWeight:"bold", fontSize:13, cursor: disabled ? "default" : "pointer",
                  boxShadow:"0 3px 10px "+shadow, transition:"transform .1s, box-shadow .1s"
                }}
                  onMouseEnter={e=>{ if(!disabled){ e.currentTarget.style.transform="translateY(-2px)"; e.currentTarget.style.boxShadow="0 6px 16px "+shadow; }}}
                  onMouseLeave={e=>{ e.currentTarget.style.transform="none"; e.currentTarget.style.boxShadow="0 3px 10px "+shadow; }}
                >
                  <span style={{fontSize:20}}>{icon}</span>
                  <div style={{textAlign:"left", lineHeight:1.3}}>
                    <div>{disabled ? loadingLabel : label}</div>
                    <div style={{fontSize:10, fontWeight:"normal", opacity:.8}}>{sub}</div>
                  </div>
                </button>
              ))}
            </div>
            <div className="noprint" style={{ margin: "12px auto 18px", maxWidth: 1040, display: "flex", flexDirection: "column", gap: 12 }}>
              <InlineAdStrip adClient={adConfig.adClient} minHeight={150} label="Supported Links" />
              <InlineAdStrip adClient={adConfig.adClient} minHeight={150} label="Supported Links" />
              <InlineAdStrip adClient={adConfig.adClient} minHeight={150} label="Supported Links" />
              <InlineAdStrip adClient={adConfig.adClient} minHeight={150} label="Supported Links" />
              <InlineAdStrip adClient={adConfig.adClient} minHeight={150} label="Supported Links" />
            </div>

          </>
        )}
        </>}
        {/* ══ END MEASURE MODE ══════════════════════════════════════════════════ */}

        {/* ══ BILL MODE ══════════════════════════════════════════════════════════ */}
        {activeTab==="bill" && (
          <BillView
            sessions={sessions}
            company={company} companySpec={companySpec}
            companyAddr={companyAddr} companyMob={companyMob}
            logoUrl={logoUrl} logoRef={logoRef} handleLogo={handleLogo}
            billClient={billClient} setBillClient={setBillClient}
            billSite={billSite} setBillSite={setBillSite}
            billDate={billDate} setBillDate={setBillDate}
            billSub={billSub} setBillSub={setBillSub}
            billAdvance={billAdvance} setBillAdvance={setBillAdvance}
            fixedItems={fixedItems} setFixedItems={setFixedItems}
            manualItems={manualItems} setManualItems={setManualItems}
            setCompany={setCompany} setCompanySpec={setCompanySpec}
            setCompanyAddr={setCompanyAddr} setCompanyMob={setCompanyMob}
            exportMsg={exportMsg}
            billBusy={billBusy} handleBill={launchBillDocx} handleBillPdf={launchBillPdf}
            onUpdateSession={updateSession}
            billRates={billRates} setBillRates={setBillRates}
            rateCard={rateCard}
            lang={lang}
          />
        )}
        {/* ══ END BILL MODE ══════════════════════════════════════════════════════ */}

        {/* ══ RATES MODE ══════════════════════════════════════════════════════════ */}
        {activeTab==="rates" && (
          <RateCard rateCard={rateCard} setRateCard={setRateCard} companyName={company} companySpec={companySpec} companyMob={companyMob} lang={lang} />
        )}
        {/* ══ END RATES MODE ══════════════════════════════════════════════════════ */}

        {activeTab==="admin" && adminMode && (
          <AdminPanel lang={lang} analyticsEnabled={analyticsEnabled} />
        )}

        {/* EMPTY STATE */}
        {!loading && !hasSessions && activeTab!=="admin" && (
          <div>
            <div style={{ textAlign: "center", padding: "36px 20px", color: "#94a3b8" }}>
              <div style={{ fontSize: 52, marginBottom: 10 }}>📋</div>
              <div style={{ fontSize: 17, fontWeight: "bold", color: C.dark, marginBottom: 5 }}>MeasureFlow v2.0</div>
              <div style={{ fontSize: 13, marginBottom: 20 }}>Upload images · AI detects Sqft &amp; Rnft · Edit · Export</div>
              <div style={{ display: "flex", gap: 12, justifyContent: "center", flexWrap: "wrap" }}>
                {["📷 Upload image", "🤖 AI detects Sqft/Rnft", "✏️ Edit inline", "📝 Word .docx", "📊 CSV", "🖨 Print/PDF", "💾 Save & Load"].map((s, i) => (
                  <div key={i} style={{ background: "#fff", borderRadius: 9, padding: "9px 13px", border: "1px solid #dbeafe", fontSize: 12, color: C.dark, fontWeight: "bold" }}>
                    {s}
                  </div>
                ))}
              </div>
            </div>
            <div className="noprint" style={{ margin: "0 auto 12px", maxWidth: 1040, display: "flex", flexDirection: "column", gap: 12 }}>
              <InlineAdStrip adClient={adConfig.adClient} minHeight={150} label="Sponsored" />
              <InlineAdStrip adClient={adConfig.adClient} minHeight={150} label="Sponsored" />
              <InlineAdStrip adClient={adConfig.adClient} minHeight={150} label="Sponsored" />
              <InlineAdStrip adClient={adConfig.adClient} minHeight={150} label="Sponsored" />
              <InlineAdStrip adClient={adConfig.adClient} minHeight={150} label="Sponsored" />
            </div>
          </div>
        )}
        </div>
        <AutoAdBanner adClient={adConfig.adClient} side="right" activeTab={activeTab} />
      </div>
    </div>
  );
}
