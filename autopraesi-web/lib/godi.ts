// API-Client für den GoDi-Plan-Reiter: liest und schreibt die GoDi-Plan-Excel
// (komplettes Tabellenblatt inkl. Ansicht) über das Cloud-Backend (/api/godi/*).

export interface GodiFile {
  name: string;
  excel_path: string;
  is_current_quarter: boolean;
}

export interface GodiSheet {
  name: string;
  is_helper: boolean;
}

export interface UpcomingSunday {
  excel_path: string | null;
  sheet: string | null;
  date: string;
}

export interface CellStyle {
  b?: boolean; // bold
  i?: boolean; // italic
  sz?: number; // font size (pt)
  fg?: string; // font color #RRGGBB
  bg?: string; // fill color #RRGGBB
  a?: string; // horizontal align: left|center|right
  va?: string; // vertical align
  wrap?: boolean; // wrap text
  bt?: boolean; // border top
  br?: boolean; // border right
  bb?: boolean; // border bottom
  bl?: boolean; // border left
}

export interface Cell {
  v?: string; // display value
  s?: CellStyle;
}

export interface Grid {
  sheet: string;
  max_row: number;
  max_col: number;
  col_widths: Record<string, number>; // colIndex (1-based) -> Excel char width
  row_heights: Record<string, number>; // rowIndex -> points
  default_col_width: number;
  default_row_height: number;
  merged: { r1: number; c1: number; r2: number; c2: number }[];
  freeze: string | null; // e.g. "A4"
  cells: Record<string, Cell>; // "row:col" -> Cell
  rev: string | null;
}

export type GridOp =
  | { op: "set"; row: number; col: number; value: string }
  | { op: "format"; row: number; col: number; bold?: boolean; italic?: boolean; color?: string; fill?: string | null; align?: string }
  | { op: "insertRow"; index: number; count?: number }
  | { op: "deleteRow"; index: number; count?: number }
  | { op: "insertCol"; index: number; count?: number }
  | { op: "deleteCol"; index: number; count?: number }
  | { op: "merge"; r1: number; c1: number; r2: number; c2: number }
  | { op: "unmerge"; r1: number; c1: number; r2: number; c2: number };

export interface SaveResult {
  success: boolean;
  changed: number;
  backup?: string | null;
  rev: string | null;
}

export async function getGodiFiles(): Promise<GodiFile[]> {
  const res = await fetch(`/api/godi/files`);
  if (!res.ok) throw new Error("GoDi-Plan-Dateien konnten nicht geladen werden");
  return res.json();
}

export async function getGodiSheets(excelPath: string): Promise<GodiSheet[]> {
  const params = new URLSearchParams({ excel_path: excelPath });
  const res = await fetch(`/api/godi/sheets?${params}`);
  if (!res.ok) throw new Error("Tabellenblätter konnten nicht geladen werden");
  return res.json();
}

export async function getUpcomingSunday(): Promise<UpcomingSunday> {
  const res = await fetch(`/api/godi/upcoming-sunday`);
  if (!res.ok) throw new Error("Kommender Sonntag konnte nicht ermittelt werden");
  return res.json();
}

export async function getGrid(excelPath: string, sheet: string): Promise<Grid> {
  const params = new URLSearchParams({ excel_path: excelPath, sheet });
  const res = await fetch(`/api/godi/grid?${params}`);
  if (!res.ok) throw new Error(`Blatt „${sheet}" konnte nicht geladen werden`);
  return res.json();
}

export async function saveGodi(req: {
  excel_path: string;
  sheet: string;
  operations: GridOp[];
  base_rev: string | null;
}): Promise<SaveResult> {
  const res = await fetch(`/api/godi/save`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(req),
  });
  if (!res.ok) {
    const err = await res.json().catch(() => ({ detail: "Unbekannter Fehler" }));
    throw new Error(err.detail || "Speichern fehlgeschlagen");
  }
  return res.json();
}

// --- Hilfen für Spaltenbezeichnungen (A, B, …, Z, AA, AB, …) ---

export function colLetter(col: number): string {
  let s = "";
  while (col > 0) {
    const m = (col - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    col = Math.floor((col - 1) / 26);
  }
  return s;
}

// Excel-Zeichenbreite → CSS-Pixel (Näherung wie in Excel: 7px/Zeichen + 5px Rand)
export function colWidthPx(charWidth: number | undefined, fallback: number): number {
  const w = charWidth ?? fallback;
  return Math.round(w * 7 + 5);
}

// Excel-Punkthöhe → CSS-Pixel
export function rowHeightPx(points: number | undefined, fallback: number): number {
  const h = points ?? fallback;
  return Math.max(22, Math.round(h * 1.333));
}
