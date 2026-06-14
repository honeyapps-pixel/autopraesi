"use client";

import { useState, useEffect, useCallback, useRef } from "react";
import {
  GodiFile,
  GodiSheet,
  Grid,
  Cell,
  CellStyle,
  GridOp,
  getGodiFiles,
  getGodiSheets,
  getUpcomingSunday,
  getGrid,
  saveGodi,
  colLetter,
  colWidthPx,
  rowHeightPx,
} from "@/lib/godi";

// "row:col"
const keyOf = (r: number, c: number) => `${r}:${c}`;

interface Range {
  r1: number;
  c1: number;
  r2: number;
  c2: number;
}

const norm = (a: { r: number; c: number }, b: { r: number; c: number }): Range => ({
  r1: Math.min(a.r, b.r),
  c1: Math.min(a.c, b.c),
  r2: Math.max(a.r, b.r),
  c2: Math.max(a.c, b.c),
});

const inRange = (r: number, c: number, sel: Range | null) =>
  !!sel && r >= sel.r1 && r <= sel.r2 && c >= sel.c1 && c <= sel.c2;

// --- Lokale, optimistische Anwendung einer Operation auf das Raster ---
// Spiegelt die Semantik des Backends (openpyxl), damit die Tabelle sofort
// reagiert und beim Sync exakt dasselbe Ergebnis entsteht.
function applyLocal(grid: Grid, op: GridOp): Grid {
  const cells = { ...grid.cells };
  const g: Grid = { ...grid, cells };

  const shiftRowKeys = (from: number, delta: number) => {
    const next: Record<string, Cell> = {};
    for (const [k, v] of Object.entries(grid.cells)) {
      const [r, c] = k.split(":").map(Number);
      const nr = r >= from ? r + delta : r;
      if (nr >= 1) next[keyOf(nr, c)] = v;
    }
    g.cells = next;
  };
  const shiftColKeys = (from: number, delta: number) => {
    const next: Record<string, Cell> = {};
    for (const [k, v] of Object.entries(grid.cells)) {
      const [r, c] = k.split(":").map(Number);
      const nc = c >= from ? c + delta : c;
      if (nc >= 1) next[keyOf(r, nc)] = v;
    }
    g.cells = next;
  };

  switch (op.op) {
    case "set": {
      const k = keyOf(op.row, op.col);
      cells[k] = { ...(cells[k] || {}), v: op.value };
      if (op.value === "") delete cells[k].v;
      break;
    }
    case "format": {
      const k = keyOf(op.row, op.col);
      const prev = cells[k] || {};
      const s: CellStyle = { ...(prev.s || {}) };
      if (op.bold !== undefined) s.b = op.bold || undefined;
      if (op.italic !== undefined) s.i = op.italic || undefined;
      if (op.color !== undefined) s.fg = op.color || undefined;
      if (op.fill !== undefined) s.bg = op.fill || undefined;
      if (op.align !== undefined) s.a = op.align || undefined;
      cells[k] = { ...prev, s };
      break;
    }
    case "insertRow":
      shiftRowKeys(op.index, op.count || 1);
      g.max_row = grid.max_row + (op.count || 1);
      g.merged = grid.merged.map((m) => ({
        ...m,
        r1: m.r1 >= op.index ? m.r1 + (op.count || 1) : m.r1,
        r2: m.r2 >= op.index ? m.r2 + (op.count || 1) : m.r2,
      }));
      break;
    case "deleteRow": {
      const cnt = op.count || 1;
      const next: Record<string, Cell> = {};
      for (const [k, v] of Object.entries(grid.cells)) {
        const [r, c] = k.split(":").map(Number);
        if (r >= op.index && r < op.index + cnt) continue;
        const nr = r >= op.index + cnt ? r - cnt : r;
        next[keyOf(nr, c)] = v;
      }
      g.cells = next;
      g.max_row = Math.max(1, grid.max_row - cnt);
      g.merged = grid.merged
        .filter((m) => !(m.r1 >= op.index && m.r2 < op.index + cnt))
        .map((m) => ({
          ...m,
          r1: m.r1 >= op.index + cnt ? m.r1 - cnt : m.r1,
          r2: m.r2 >= op.index + cnt ? m.r2 - cnt : m.r2,
        }));
      break;
    }
    case "insertCol":
      shiftColKeys(op.index, op.count || 1);
      g.max_col = grid.max_col + (op.count || 1);
      g.merged = grid.merged.map((m) => ({
        ...m,
        c1: m.c1 >= op.index ? m.c1 + (op.count || 1) : m.c1,
        c2: m.c2 >= op.index ? m.c2 + (op.count || 1) : m.c2,
      }));
      break;
    case "deleteCol": {
      const cnt = op.count || 1;
      const next: Record<string, Cell> = {};
      for (const [k, v] of Object.entries(grid.cells)) {
        const [r, c] = k.split(":").map(Number);
        if (c >= op.index && c < op.index + cnt) continue;
        const nc = c >= op.index + cnt ? c - cnt : c;
        next[keyOf(r, nc)] = v;
      }
      g.cells = next;
      g.max_col = Math.max(1, grid.max_col - cnt);
      g.merged = grid.merged
        .filter((m) => !(m.c1 >= op.index && m.c2 < op.index + cnt))
        .map((m) => ({
          ...m,
          c1: m.c1 >= op.index + cnt ? m.c1 - cnt : m.c1,
          c2: m.c2 >= op.index + cnt ? m.c2 - cnt : m.c2,
        }));
      break;
    }
    case "merge":
      g.merged = [...grid.merged, { r1: op.r1, c1: op.c1, r2: op.r2, c2: op.c2 }];
      break;
    case "unmerge":
      g.merged = grid.merged.filter(
        (m) => !(m.r1 === op.r1 && m.c1 === op.c1 && m.r2 === op.r2 && m.c2 === op.c2)
      );
      break;
  }
  return g;
}

export default function GodiPlanTab() {
  const [files, setFiles] = useState<GodiFile[]>([]);
  const [excelPath, setExcelPath] = useState<string | null>(null);
  const [sheets, setSheets] = useState<GodiSheet[]>([]);
  const [sheet, setSheet] = useState<string | null>(null);
  const [upcomingSheet, setUpcomingSheet] = useState<string | null>(null);
  const [grid, setGrid] = useState<Grid | null>(null);
  const [baseRev, setBaseRev] = useState<string | null>(null);

  const [loading, setLoading] = useState(true);
  const [error, setError] = useState<string | null>(null);

  const [ops, setOps] = useState<GridOp[]>([]);
  const [saving, setSaving] = useState(false);
  const [status, setStatus] = useState<string | null>(null);

  const [anchor, setAnchor] = useState<{ r: number; c: number } | null>(null);
  const [focusCell, setFocusCell] = useState<{ r: number; c: number } | null>(null);
  const [editing, setEditing] = useState<{ r: number; c: number; value: string } | null>(null);

  const sel = anchor && focusCell ? norm(anchor, focusCell) : null;
  const dirty = ops.length > 0;
  const gridScrollRef = useRef<HTMLDivElement>(null);

  // Kompakte Darstellung auf Mobil (kleinere Zellen/Spalten)
  const [compact, setCompact] = useState(false);
  useEffect(() => {
    const mq = window.matchMedia("(max-width: 768px)");
    const update = () => setCompact(mq.matches);
    update();
    mq.addEventListener("change", update);
    return () => mq.removeEventListener("change", update);
  }, []);

  // Verfügbare Rasterbreite messen (für „ganze Tabelle auf einen Blick" am Desktop)
  const [containerW, setContainerW] = useState(0);
  useEffect(() => {
    const el = gridScrollRef.current;
    if (!el) return;
    const update = () => setContainerW(el.clientWidth);
    const ro = new ResizeObserver(update);
    ro.observe(el);
    update();
    return () => ro.disconnect();
  }, [grid]);

  // --- Initial: kommender Sonntag + Datei-Liste ---
  useEffect(() => {
    let cancelled = false;
    (async () => {
      setLoading(true);
      try {
        const [up, fileList] = await Promise.all([getUpcomingSunday(), getGodiFiles()]);
        if (cancelled) return;
        setFiles(fileList);
        setUpcomingSheet(up.sheet);
        const path =
          up.excel_path ||
          fileList.find((f) => f.is_current_quarter)?.excel_path ||
          fileList[0]?.excel_path ||
          null;
        setExcelPath(path);
        if (path) {
          const sh = await getGodiSheets(path);
          if (cancelled) return;
          setSheets(sh);
          const initial =
            (up.excel_path === path && up.sheet && sh.find((s) => s.name === up.sheet)?.name) ||
            sh.find((s) => !s.is_helper)?.name ||
            sh[0]?.name ||
            null;
          setSheet(initial);
        } else {
          setLoading(false);
        }
      } catch (e) {
        if (!cancelled) {
          setError((e as Error).message);
          setLoading(false);
        }
      }
    })();
    return () => {
      cancelled = true;
    };
  }, []);

  // --- Datei gewechselt → Blätter neu laden ---
  const switchFile = async (path: string) => {
    if (path === excelPath) return;
    if (dirty && !confirm("Ungespeicherte Änderungen verwerfen?")) return;
    setExcelPath(path);
    setSheet(null); // verhindert ein kurzes Laden der alten Mappe in der neuen Datei
    setOps([]);
    setGrid(null);
    setLoading(true);
    try {
      const sh = await getGodiSheets(path);
      setSheets(sh);
      setSheet(sh.find((s) => !s.is_helper)?.name || sh[0]?.name || null);
    } catch (e) {
      setError((e as Error).message);
      setLoading(false);
    }
  };

  // --- Blatt laden ---
  const loadGrid = useCallback(async (path: string, sh: string) => {
    setLoading(true);
    setError(null);
    setOps([]);
    setAnchor(null);
    setFocusCell(null);
    setEditing(null);
    try {
      const g = await getGrid(path, sh);
      setGrid(g);
      setBaseRev(g.rev);
    } catch (e) {
      setError((e as Error).message);
      setGrid(null);
    } finally {
      setLoading(false);
    }
  }, []);

  useEffect(() => {
    if (excelPath && sheet) loadGrid(excelPath, sheet);
  }, [excelPath, sheet, loadGrid]);

  // DOM-Fokus der Auswahl nachführen, damit Pfeiltasten-Navigation funktioniert
  useEffect(() => {
    if (editing || !anchor) return;
    const el = gridScrollRef.current?.querySelector<HTMLElement>(
      `[data-cell="${anchor.r}:${anchor.c}"]`
    );
    el?.focus({ preventScroll: false });
  }, [anchor, editing]);

  const selectSheet = (name: string) => {
    if (name === sheet) return;
    if (dirty && !confirm("Ungespeicherte Änderungen verwerfen?")) return;
    setSheet(name);
  };

  // --- Operationen anwenden (lokal + vormerken) ---
  const pushOps = useCallback((newOps: GridOp[]) => {
    if (newOps.length === 0) return;
    setGrid((g) => (g ? newOps.reduce(applyLocal, g) : g));
    setOps((prev) => [...prev, ...newOps]);
    setStatus(null);
  }, []);

  // --- Zelle bearbeiten ---
  const displayValue = (r: number, c: number) => grid?.cells[keyOf(r, c)]?.v ?? "";

  const startEdit = (r: number, c: number) => {
    setEditing({ r, c, value: displayValue(r, c) });
  };

  const commitEdit = (move?: "down" | "right") => {
    if (!editing || !grid) return;
    const original = displayValue(editing.r, editing.c);
    if (editing.value !== original) {
      pushOps([{ op: "set", row: editing.r, col: editing.c, value: editing.value }]);
    }
    const { r, c } = editing;
    setEditing(null);
    if (move === "down" && r < grid.max_row) {
      setAnchor({ r: r + 1, c });
      setFocusCell({ r: r + 1, c });
    } else if (move === "right" && c < grid.max_col) {
      setAnchor({ r, c: c + 1 });
      setFocusCell({ r, c: c + 1 });
    }
  };

  // --- Formatierung auf Auswahl anwenden ---
  const formatSelection = (patch: Partial<Extract<GridOp, { op: "format" }>>) => {
    if (!sel) return;
    const list: GridOp[] = [];
    for (let r = sel.r1; r <= sel.r2; r++)
      for (let c = sel.c1; c <= sel.c2; c++) list.push({ op: "format", row: r, col: c, ...patch });
    pushOps(list);
  };

  const activeStyle: CellStyle | undefined = anchor ? grid?.cells[keyOf(anchor.r, anchor.c)]?.s : undefined;

  const toggleBold = () => formatSelection({ bold: !activeStyle?.b });
  const toggleItalic = () => formatSelection({ italic: !activeStyle?.i });

  const mergeSelection = () => {
    if (!sel || (sel.r1 === sel.r2 && sel.c1 === sel.c2)) return;
    pushOps([{ op: "merge", ...sel }]);
  };
  const unmergeSelection = () => {
    if (!sel || !grid) return;
    const m = grid.merged.find((m) => m.r1 === sel.r1 && m.c1 === sel.c1);
    if (m) pushOps([{ op: "unmerge", ...m }]);
  };

  const insertRow = () => anchor && pushOps([{ op: "insertRow", index: anchor.r }]);
  const deleteRow = () => anchor && pushOps([{ op: "deleteRow", index: anchor.r }]);
  const insertCol = () => anchor && pushOps([{ op: "insertCol", index: anchor.c }]);
  const deleteCol = () => anchor && pushOps([{ op: "deleteCol", index: anchor.c }]);

  // --- Sync ---
  const sync = async () => {
    if (!excelPath || !sheet || ops.length === 0) return;
    setSaving(true);
    setError(null);
    setStatus(null);
    try {
      const res = await saveGodi({ excel_path: excelPath, sheet, operations: ops, base_rev: baseRev });
      setOps([]);
      setStatus(`Gespeichert und mit Dropbox synchronisiert${res.backup ? " · Sicherung angelegt" : ""}.`);
      // Frisch laden, damit serverseitige Normalisierungen sichtbar sind
      await loadGrid(excelPath, sheet);
    } catch (e) {
      setError((e as Error).message);
    } finally {
      setSaving(false);
    }
  };

  // --- Tastatur im Raster ---
  const onCellKeyDown = (e: React.KeyboardEvent, r: number, c: number) => {
    if (editing) return;
    if (e.key === "Enter" || e.key === "F2") {
      e.preventDefault();
      startEdit(r, c);
    } else if (e.key.length === 1 && !e.metaKey && !e.ctrlKey) {
      setEditing({ r, c, value: "" });
    } else if (["ArrowUp", "ArrowDown", "ArrowLeft", "ArrowRight"].includes(e.key) && grid) {
      e.preventDefault();
      const nr = Math.min(grid.max_row, Math.max(1, r + (e.key === "ArrowDown" ? 1 : e.key === "ArrowUp" ? -1 : 0)));
      const nc = Math.min(grid.max_col, Math.max(1, c + (e.key === "ArrowRight" ? 1 : e.key === "ArrowLeft" ? -1 : 0)));
      setAnchor({ r: nr, c: nc });
      setFocusCell({ r: nr, c: nc });
    }
  };

  // ---------- Render ----------
  if (loading && !grid) {
    return (
      <div className="max-w-[120rem] mx-auto px-4 py-10">
        <div className="flex flex-col items-center justify-center py-24 gap-3">
          <span className="inline-block w-6 h-6 border-2 border-[var(--accent)] border-t-transparent rounded-full animate-spin" />
          <p className="text-sm text-[var(--text-secondary)]">Lade Gottesdienstplan…</p>
        </div>
      </div>
    );
  }

  return (
    <div className="max-w-[120rem] mx-auto px-2 sm:px-4 py-3 sm:py-5">
      <Toolbar
        files={files}
        excelPath={excelPath}
        onSwitchFile={switchFile}
        activeStyle={activeStyle}
        hasSelection={!!sel}
        onBold={toggleBold}
        onItalic={toggleItalic}
        onColor={(color) => formatSelection({ color })}
        onFill={(fill) => formatSelection({ fill })}
        onAlign={(align) => formatSelection({ align })}
        onMerge={mergeSelection}
        onUnmerge={unmergeSelection}
        onInsertRow={insertRow}
        onDeleteRow={deleteRow}
        onInsertCol={insertCol}
        onDeleteCol={deleteCol}
        dirty={dirty}
        saving={saving}
        onSync={sync}
        pendingCount={ops.length}
      />

      {error && (
        <div className="mt-3 rounded-xl border border-[var(--danger)]/30 bg-[var(--danger)]/5 px-4 py-2.5 text-sm text-[var(--danger)]">
          {error}
        </div>
      )}
      {status && !error && (
        <div className="mt-3 rounded-xl border border-[var(--success)]/30 bg-[var(--success)]/5 px-4 py-2.5 text-sm text-[var(--success)] flex items-center gap-2" role="status" aria-live="polite">
          <span className="status-dot found" />
          {status}
        </div>
      )}

      {/* Raster */}
      <div className="relative mt-3">
      <div
        ref={gridScrollRef}
        className={`godi-grid-scroll overflow-auto rounded-xl border border-[var(--card-border)] bg-white shadow-sm transition-opacity duration-150 ${
          compact ? "godi-compact" : ""
        } ${loading ? "opacity-40" : "opacity-100"}`}
      >
        {grid && (
          <GridTable
            grid={grid}
            compact={compact}
            containerW={containerW}
            sel={sel}
            anchor={anchor}
            editing={editing}
            onSelect={(r, c, shift) => {
              if (shift && anchor) setFocusCell({ r, c });
              else {
                setAnchor({ r, c });
                setFocusCell({ r, c });
              }
            }}
            onStartEdit={startEdit}
            onEditChange={(value) => setEditing((e) => (e ? { ...e, value } : e))}
            onCommit={commitEdit}
            onCancel={() => setEditing(null)}
            onCellKeyDown={onCellKeyDown}
          />
        )}
      </div>
      {loading && grid && (
        <div className="absolute inset-0 z-40 flex items-center justify-center rounded-xl">
          <div className="flex items-center gap-2.5 rounded-full bg-white/90 px-4 py-2 shadow-md border border-[var(--card-border)]">
            <span className="inline-block w-4 h-4 border-2 border-[var(--accent)] border-t-transparent rounded-full animate-spin" />
            <span className="text-sm text-[var(--text-secondary)]">Lade Mappe…</span>
          </div>
        </div>
      )}
      </div>

      {/* Mappen-Tabs (Sonntage) – wie in Excel unten */}
      <SheetTabs sheets={sheets} active={sheet} upcoming={upcomingSheet} onSelect={selectSheet} />
    </div>
  );
}

// ============================ Toolbar ============================

const FILL_SWATCHES = [
  { c: null, label: "Keine Füllung" },
  { c: "#C5E0B4", label: "Grün (Lied)" },
  { c: "#FFA7A7", label: "Rot (Predigt)" },
  { c: "#F8CBAD", label: "Orange (Lesung)" },
  { c: "#FFE699", label: "Gelb" },
  { c: "#BDD7EE", label: "Blau" },
];

function Toolbar(props: {
  files: GodiFile[];
  excelPath: string | null;
  onSwitchFile: (p: string) => void;
  activeStyle?: CellStyle;
  hasSelection: boolean;
  onBold: () => void;
  onItalic: () => void;
  onColor: (c: string) => void;
  onFill: (c: string | null) => void;
  onAlign: (a: string) => void;
  onMerge: () => void;
  onUnmerge: () => void;
  onInsertRow: () => void;
  onDeleteRow: () => void;
  onInsertCol: () => void;
  onDeleteCol: () => void;
  dirty: boolean;
  saving: boolean;
  onSync: () => void;
  pendingCount: number;
}) {
  const s = props.activeStyle;
  const dis = !props.hasSelection;
  return (
    <div className="rounded-xl border border-[var(--card-border)] bg-[var(--card)] backdrop-blur px-2 py-2 sm:px-3 space-y-2">
      {/* Reihe 1: Datei-Auswahl + Sync */}
      <div className="flex items-center gap-2">
        <select
          className="text-sm bg-white border border-[var(--card-border)] rounded-lg px-2.5 py-1.5 outline-none cursor-pointer focus:border-[var(--accent)] min-w-0 flex-1 sm:flex-none sm:max-w-[15rem]"
          value={props.excelPath || ""}
          onChange={(e) => props.onSwitchFile(e.target.value)}
          title="Excel-Datei wählen"
        >
          {props.files.map((f) => (
            <option key={f.excel_path} value={f.excel_path}>
              {f.name.replace(/\.xlsx$/, "")}
              {f.is_current_quarter ? "  · aktuell" : ""}
            </option>
          ))}
        </select>

        <div className="ml-auto flex items-center gap-2 shrink-0">
          {props.dirty && (
            <span className="text-xs text-[var(--warning)] font-medium tabular-nums whitespace-nowrap">
              {props.pendingCount} {props.pendingCount === 1 ? "Änderung" : "Änderungen"}
            </span>
          )}
          <button
            onClick={props.onSync}
            disabled={!props.dirty || props.saving}
            className="flex items-center gap-2 rounded-lg bg-[var(--accent)] text-white text-sm font-medium px-3 sm:px-4 py-1.5 transition-all duration-150 enabled:hover:bg-[var(--accent-hover)] disabled:opacity-40 disabled:cursor-not-allowed cursor-pointer whitespace-nowrap"
          >
            {props.saving ? (
              <>
                <span className="inline-block w-3.5 h-3.5 border-2 border-white border-t-transparent rounded-full animate-spin" />
                <span className="hidden sm:inline">Synchronisiere…</span>
                <span className="sm:hidden">Sync…</span>
              </>
            ) : (
              <>
                <SyncIcon />
                <span className="hidden sm:inline">Mit Dropbox synchronisieren</span>
                <span className="sm:hidden">Sync</span>
              </>
            )}
          </button>
        </div>
      </div>

      {/* Reihe 2: Werkzeuge – auf Mobil horizontal scrollbar, auf Desktop umbrechend */}
      <div className="flex items-center gap-1.5 sm:gap-2 overflow-x-auto sm:flex-wrap sm:overflow-visible -mx-1 px-1 [&>*]:shrink-0">
        <TbBtn onClick={props.onBold} active={!!s?.b} disabled={dis} label="Fett" className="font-bold">B</TbBtn>
        <TbBtn onClick={props.onItalic} active={!!s?.i} disabled={dis} label="Kursiv" className="italic font-serif">I</TbBtn>

        <Divider />

        {/* Ausrichtung */}
        {([["left", "Links"], ["center", "Mitte"], ["right", "Rechts"]] as const).map(([a, lbl]) => (
          <TbBtn key={a} onClick={() => props.onAlign(a)} active={s?.a === a} disabled={dis} label={`Ausrichtung ${lbl}`}>
            <AlignIcon a={a} />
          </TbBtn>
        ))}

        <Divider />

        {/* Textfarbe */}
        <div className="flex items-center gap-1" title="Textfarbe">
          <button disabled={dis} onClick={() => props.onColor("#000000")} className="tb-swatch" style={{ background: "#1d1d1f" }} aria-label="Schrift schwarz" />
          <button disabled={dis} onClick={() => props.onColor("#FFFFFF")} className="tb-swatch border border-[var(--card-border)]" style={{ background: "#fff" }} aria-label="Schrift weiß" />
        </div>

        <Divider />

        {/* Füllung */}
        <div className="flex items-center gap-1" title="Zellfarbe">
          {FILL_SWATCHES.map((sw) => (
            <button
              key={sw.label}
              disabled={dis}
              onClick={() => props.onFill(sw.c)}
              className="tb-swatch border border-[var(--card-border)] relative"
              style={{ background: sw.c || "#fff" }}
              aria-label={sw.label}
              title={sw.label}
            >
              {sw.c === null && <span className="absolute inset-0 flex items-center justify-center text-[var(--danger)] text-xs leading-none">⁄</span>}
            </button>
          ))}
        </div>

        <Divider />

        <TbBtn onClick={props.onMerge} disabled={dis} label="Zellen verbinden"><MergeIcon /></TbBtn>
        <TbBtn onClick={props.onUnmerge} disabled={dis} label="Verbindung lösen"><UnmergeIcon /></TbBtn>

        <Divider />

        {/* Zeilen / Spalten */}
        <TbBtn onClick={props.onInsertRow} disabled={dis} label="Zeile einfügen">+Zeile</TbBtn>
        <TbBtn onClick={props.onDeleteRow} disabled={dis} label="Zeile löschen" danger>−Zeile</TbBtn>
        <TbBtn onClick={props.onInsertCol} disabled={dis} label="Spalte einfügen">+Spalte</TbBtn>
        <TbBtn onClick={props.onDeleteCol} disabled={dis} label="Spalte löschen" danger>−Spalte</TbBtn>
      </div>
    </div>
  );
}

function TbBtn({
  children,
  onClick,
  active,
  disabled,
  label,
  className = "",
  danger,
}: {
  children: React.ReactNode;
  onClick: () => void;
  active?: boolean;
  disabled?: boolean;
  label: string;
  className?: string;
  danger?: boolean;
}) {
  return (
    <button
      type="button"
      onClick={onClick}
      disabled={disabled}
      title={label}
      aria-label={label}
      aria-pressed={active}
      className={`h-8 min-w-8 px-2 inline-flex items-center justify-center rounded-lg text-sm transition-colors duration-150 cursor-pointer disabled:opacity-30 disabled:cursor-not-allowed ${
        active
          ? "bg-[var(--accent)] text-white"
          : danger
          ? "text-[var(--text-secondary)] enabled:hover:bg-[var(--danger)]/10 enabled:hover:text-[var(--danger)]"
          : "text-[var(--text-primary)] enabled:hover:bg-black/5"
      } ${className}`}
    >
      {children}
    </button>
  );
}

const Divider = () => <span className="w-px h-5 bg-[var(--card-border)]" aria-hidden />;

// ============================ Grid ============================

function GridTable({
  grid,
  compact,
  containerW,
  sel,
  anchor,
  editing,
  onSelect,
  onStartEdit,
  onEditChange,
  onCommit,
  onCancel,
  onCellKeyDown,
}: {
  grid: Grid;
  compact: boolean;
  containerW: number;
  sel: Range | null;
  anchor: { r: number; c: number } | null;
  editing: { r: number; c: number; value: string } | null;
  onSelect: (r: number, c: number, shift: boolean) => void;
  onStartEdit: (r: number, c: number) => void;
  onEditChange: (v: string) => void;
  onCommit: (move?: "down" | "right") => void;
  onCancel: () => void;
  onCellKeyDown: (e: React.KeyboardEvent, r: number, c: number) => void;
}) {
  // Verdeckte Zellen (Teil eines Merges, nicht oben-links) ermitteln
  const covered = new Set<string>();
  const spanOf: Record<string, { rs: number; cs: number }> = {};
  for (const m of grid.merged) {
    spanOf[keyOf(m.r1, m.c1)] = { rs: m.r2 - m.r1 + 1, cs: m.c2 - m.c1 + 1 };
    for (let r = m.r1; r <= m.r2; r++)
      for (let c = m.c1; c <= m.c2; c++) if (!(r === m.r1 && c === m.c1)) covered.add(keyOf(r, c));
  }

  const rows = Array.from({ length: grid.max_row }, (_, i) => i + 1);
  const cols = Array.from({ length: grid.max_col }, (_, i) => i + 1);

  // Auf Mobil Spalten schmaler skalieren und deckeln, damit das Raster aufs Display passt
  const colW = (c: number) => {
    const w = colWidthPx(grid.col_widths[c], grid.default_col_width);
    return compact ? Math.max(30, Math.min(150, Math.round(w * 0.58))) : w;
  };
  const rowH = (r: number) => {
    const h = rowHeightPx(grid.row_heights[r], grid.default_row_height);
    return compact ? Math.max(20, Math.round(h * 0.78)) : h;
  };
  const gutter = compact ? 28 : 44;

  // Fit-to-Width (Desktop): das Raster so weit herunterzoomen, dass alle Spalten
  // ohne horizontales Scrollen sichtbar sind. `zoom` skaliert alles proportional
  // (Breiten, Schrift, Höhen) und erhält dabei Layout + Sticky-Header.
  let naturalW = gutter;
  for (const c of cols) naturalW += colW(c);
  const fitZoom =
    !compact && containerW > 0 && naturalW > containerW
      ? Math.max(0.4, (containerW - 2) / naturalW)
      : 1;

  return (
    <table
      className="border-separate border-spacing-0 select-none"
      style={{ tableLayout: "fixed", zoom: fitZoom }}
    >
      <colgroup>
        <col style={{ width: gutter }} />
        {cols.map((c) => (
          <col key={c} style={{ width: colW(c) }} />
        ))}
      </colgroup>
      <thead>
        <tr>
          <th className="godi-corner sticky top-0 left-0 z-30" />
          {cols.map((c) => (
            <th
              key={c}
              className={`godi-colhead sticky top-0 z-20 ${sel && c >= sel.c1 && c <= sel.c2 ? "godi-head-active" : ""}`}
            >
              {colLetter(c)}
            </th>
          ))}
        </tr>
      </thead>
      <tbody>
        {rows.map((r) => (
          <tr key={r} style={{ height: rowH(r) }}>
            <th
              className={`godi-rowhead sticky left-0 z-10 ${sel && r >= sel.r1 && r <= sel.r2 ? "godi-head-active" : ""}`}
            >
              {r}
            </th>
            {cols.map((c) => {
              const k = keyOf(r, c);
              if (covered.has(k)) return null;
              const cell = grid.cells[k];
              const span = spanOf[k];
              const isAnchor = anchor?.r === r && anchor?.c === c;
              const isEditing = editing?.r === r && editing?.c === c;
              const style = cell?.s;
              return (
                <td
                  key={c}
                  data-cell={k}
                  rowSpan={span?.rs}
                  colSpan={span?.cs}
                  tabIndex={0}
                  onMouseDown={(e) => onSelect(r, c, e.shiftKey)}
                  onDoubleClick={() => onStartEdit(r, c)}
                  onKeyDown={(e) => onCellKeyDown(e, r, c)}
                  className={`godi-cell ${inRange(r, c, sel) ? "godi-sel" : ""} ${isAnchor ? "godi-anchor" : ""}`}
                  style={{
                    background: style?.bg,
                    color: style?.fg,
                    fontWeight: style?.b ? 700 : undefined,
                    fontStyle: style?.i ? "italic" : undefined,
                    fontSize: style?.sz ? `${style.sz}px` : undefined,
                    textAlign: (style?.a as React.CSSProperties["textAlign"]) || undefined,
                    verticalAlign: style?.va === "center" ? "middle" : style?.va,
                    whiteSpace: style?.wrap ? "normal" : "nowrap",
                    boxShadow: borderShadow(style),
                  }}
                >
                  {isEditing ? (
                    <input
                      autoFocus
                      value={editing!.value}
                      onChange={(e) => onEditChange(e.target.value)}
                      onBlur={() => onCommit()}
                      onKeyDown={(e) => {
                        if (e.key === "Enter") {
                          e.preventDefault();
                          onCommit("down");
                        } else if (e.key === "Tab") {
                          e.preventDefault();
                          onCommit("right");
                        } else if (e.key === "Escape") {
                          e.preventDefault();
                          onCancel();
                        }
                      }}
                      className="godi-input"
                    />
                  ) : (
                    cell?.v
                  )}
                </td>
              );
            })}
          </tr>
        ))}
      </tbody>
    </table>
  );
}

// Rahmen einer Zelle als inset box-shadow (überlagert die Standard-Gitterlinie)
function borderShadow(s?: CellStyle): string | undefined {
  if (!s) return undefined;
  const col = "#9aa0a6";
  const parts: string[] = [];
  if (s.bt) parts.push(`inset 0 1px 0 0 ${col}`);
  if (s.bb) parts.push(`inset 0 -1px 0 0 ${col}`);
  if (s.bl) parts.push(`inset 1px 0 0 0 ${col}`);
  if (s.br) parts.push(`inset -1px 0 0 0 ${col}`);
  return parts.length ? parts.join(", ") : undefined;
}

// ============================ Sheet-Tabs ============================

function SheetTabs({
  sheets,
  active,
  upcoming,
  onSelect,
}: {
  sheets: GodiSheet[];
  active: string | null;
  upcoming: string | null;
  onSelect: (n: string) => void;
}) {
  return (
    <div className="mt-2 flex items-stretch gap-0.5 overflow-x-auto border-t border-[var(--card-border)] pt-2">
      {sheets.map((s) => {
        const isActive = s.name === active;
        const isUpcoming = s.name === upcoming;
        return (
          <button
            key={s.name}
            onClick={() => onSelect(s.name)}
            title={isUpcoming ? "Kommender Sonntag" : undefined}
            className={`group relative whitespace-nowrap rounded-t-lg px-3.5 py-1.5 text-[13px] transition-colors duration-150 cursor-pointer border border-b-0 ${
              isActive
                ? "bg-white border-[var(--card-border)] text-[var(--text-primary)] font-semibold -mb-px"
                : "bg-transparent border-transparent text-[var(--text-secondary)] hover:bg-black/5"
            } ${s.is_helper ? "opacity-60 italic" : ""}`}
          >
            {isUpcoming && (
              <span className={`inline-block w-1.5 h-1.5 rounded-full mr-1.5 align-middle ${isActive ? "bg-[var(--accent)]" : "bg-[var(--accent)]/60"}`} />
            )}
            {s.name}
          </button>
        );
      })}
    </div>
  );
}

// ============================ Icons ============================

function AlignIcon({ a }: { a: "left" | "center" | "right" }) {
  const lines =
    a === "left"
      ? ["M3 5h14", "M3 9h9", "M3 13h12", "M3 17h7"]
      : a === "center"
      ? ["M3 5h14", "M5 9h10", "M4 13h12", "M6 17h8"]
      : ["M3 5h14", "M8 9h9", "M5 13h12", "M10 17h7"];
  return (
    <svg width="16" height="16" viewBox="0 0 20 22" fill="none" stroke="currentColor" strokeWidth="1.6" strokeLinecap="round">
      {lines.map((d, i) => (
        <path key={i} d={d} />
      ))}
    </svg>
  );
}

function MergeIcon() {
  return (
    <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.7" strokeLinecap="round" strokeLinejoin="round">
      <rect x="3" y="5" width="18" height="14" rx="1.5" />
      <path d="M9 5v14M15 5v14" strokeDasharray="2 2" opacity="0.5" />
      <path d="M8 12h8M13 9l3 3-3 3" />
    </svg>
  );
}

function UnmergeIcon() {
  return (
    <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.7" strokeLinecap="round" strokeLinejoin="round">
      <rect x="3" y="5" width="18" height="14" rx="1.5" />
      <path d="M12 5v14" />
    </svg>
  );
}

function SyncIcon() {
  return (
    <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
      <path d="M21 12a9 9 0 0 1-9 9c-2.5 0-4.8-1-6.4-2.6L3 16" />
      <path d="M3 12a9 9 0 0 1 9-9c2.5 0 4.8 1 6.4 2.6L21 8" />
      <path d="M21 3v5h-5M3 21v-5h5" />
    </svg>
  );
}
