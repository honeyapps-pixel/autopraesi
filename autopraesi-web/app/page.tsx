"use client";

import { useState, useEffect, useCallback, useRef } from "react";
import { motion, AnimatePresence } from "framer-motion";
import {
  DndContext,
  closestCenter,
  KeyboardSensor,
  PointerSensor,
  useSensor,
  useSensors,
  DragEndEvent,
} from "@dnd-kit/core";
import {
  arrayMove,
  SortableContext,
  sortableKeyboardCoordinates,
  useSortable,
  verticalListSortingStrategy,
} from "@dnd-kit/sortable";
import { CSS } from "@dnd-kit/utilities";
import {
  Sheet,
  SheetData,
  SectionInfo,
  InvitationEvent,
  SongSearchResult,
  getSheets,
  getSections,
  getSheetData,
  getSheetRows,
  generate,
  uploadImage,
  searchSong,
  downloadUrl,
  imagePreviewUrl,
  ExcelRow,
  GenerateResult,
} from "@/lib/api";

// --- Slide Text Layout ---
interface TextLayout {
  x: number; y: number; w: number; h: number; fontSize: number;
}

// --- Draggable Slide Text Element ---
function DraggableText({
  text, layout, color, italic, onLayoutChange, containerRef,
}: {
  text: string;
  layout: TextLayout;
  color: string;
  italic?: boolean;
  onLayoutChange: (l: TextLayout) => void;
  containerRef: React.RefObject<HTMLDivElement | null>;
}) {
  const dragStart = useRef<{ startX: number; startY: number; origX: number; origY: number } | null>(null);

  const handlePointerDown = (e: React.PointerEvent) => {
    e.preventDefault();
    e.stopPropagation();
    const el = e.currentTarget as HTMLElement;
    el.setPointerCapture(e.pointerId);
    dragStart.current = { startX: e.clientX, startY: e.clientY, origX: layout.x, origY: layout.y };
  };

  const handlePointerMove = (e: React.PointerEvent) => {
    if (!dragStart.current || !containerRef.current) return;
    const rect = containerRef.current.getBoundingClientRect();
    const dx = ((e.clientX - dragStart.current.startX) / rect.width) * 100;
    const dy = ((e.clientY - dragStart.current.startY) / rect.height) * 100;
    onLayoutChange({
      ...layout,
      x: Math.max(0, Math.min(100 - layout.w, dragStart.current.origX + dx)),
      y: Math.max(0, Math.min(100 - layout.h, dragStart.current.origY + dy)),
    });
  };

  const handlePointerUp = () => { dragStart.current = null; };

  const shadow = color === "white"
    ? "0 2px 6px rgba(0,0,0,0.8), 0 0 20px rgba(0,0,0,0.3)"
    : "0 2px 6px rgba(255,255,255,0.6), 0 0 20px rgba(255,255,255,0.2)";

  return (
    <div
      className="absolute cursor-move select-none hover:outline hover:outline-2 hover:outline-dashed hover:outline-white/50 rounded"
      style={{
        left: `${layout.x}%`, top: `${layout.y}%`,
        width: `${layout.w}%`, height: `${layout.h}%`,
        color, textShadow: shadow,
        fontSize: `${layout.fontSize}cqi`,
        fontStyle: italic ? "italic" : "normal",
        fontWeight: "bold",
        lineHeight: 1.1,
        display: "flex", alignItems: "center", justifyContent: "center",
        textAlign: "center",
        touchAction: "none",
      }}
      onPointerDown={handlePointerDown}
      onPointerMove={handlePointerMove}
      onPointerUp={handlePointerUp}
    >
      <span style={{ wordBreak: "break-word" }}>{text}</span>
    </div>
  );
}

// --- Slide Preview ---
function SlidePreview({
  imageUrl, theme, dateStr, kirchenkalender, textColor,
  titleLayout, subtitleLayout, onTitleLayoutChange, onSubtitleLayoutChange,
}: {
  imageUrl: string | null;
  theme: string;
  dateStr: string;
  kirchenkalender: string;
  textColor: string;
  titleLayout: TextLayout;
  subtitleLayout: TextLayout;
  onTitleLayoutChange: (l: TextLayout) => void;
  onSubtitleLayoutChange: (l: TextLayout) => void;
}) {
  const containerRef = useRef<HTMLDivElement>(null);

  if (!imageUrl) {
    return (
      <div className="flex items-center gap-3 mb-3">
        <span className="status-dot missing" />
        <span className="text-sm">Kein Hintergrundbild gefunden</span>
      </div>
    );
  }

  return (
    <div
      ref={containerRef}
      className="relative rounded-xl overflow-hidden border border-[var(--card-border)] select-none"
      style={{ aspectRatio: "4 / 3", containerType: "inline-size" }}
    >
      <img src={imageUrl} alt="" className="absolute inset-0 w-full h-full object-cover" draggable={false} />
      <DraggableText
        text={theme || "Thema"}
        layout={titleLayout}
        color={textColor}
        italic
        onLayoutChange={onTitleLayoutChange}
        containerRef={containerRef}
      />
      <DraggableText
        text={`Gottesdienst am ${dateStr}, ${kirchenkalender}`}
        layout={subtitleLayout}
        color={textColor}
        onLayoutChange={onSubtitleLayoutChange}
        containerRef={containerRef}
      />
    </div>
  );
}

// --- Excel Table (shared) ---
function ExcelTable({ rows }: { rows: ExcelRow[] }) {
  return (
    <table className="w-full text-sm">
      <thead>
        <tr className="text-[10px] uppercase tracking-wider text-[var(--text-secondary)] border-b border-[var(--card-border)]">
          <th className="text-left py-1.5 px-2 w-14">Zeit</th>
          <th className="text-left py-1.5 px-2">Programm</th>
          <th className="text-left py-1.5 px-2">Details</th>
        </tr>
      </thead>
      <tbody>
        {rows.map((r) => (
          <tr key={r.row} className="border-b border-[var(--card-border)]/30 even:bg-white/30">
            <td className="py-1 px-2 text-[11px] font-mono text-[var(--text-secondary)] whitespace-nowrap">{r.uhrzeit}</td>
            <td className="py-1 px-2 text-xs font-medium">{r.programmpunkt}</td>
            <td className="py-1 px-2 text-[11px] text-[var(--text-secondary)] max-w-[180px] truncate">{r.details}</td>
          </tr>
        ))}
      </tbody>
    </table>
  );
}

// --- Mobile Excel Card ---
function MobileExcelCard({ rows, loading, open, onToggle }: {
  rows: ExcelRow[]; loading: boolean; open: boolean; onToggle: () => void;
}) {
  return (
    <div className="glass-card mb-4 md:hidden">
      <button type="button" onClick={onToggle} className="w-full flex items-center justify-between">
        <h3 className="text-xs font-semibold uppercase tracking-wider text-[var(--text-secondary)]">
          GoDi-Plan (Excel)
        </h3>
        <svg
          width="16" height="16" viewBox="0 0 24 24" fill="none"
          stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"
          className={`text-[var(--text-secondary)] transition-transform duration-200 ${open ? "rotate-180" : ""}`}
        >
          <polyline points="6 9 12 15 18 9" />
        </svg>
      </button>
      <AnimatePresence>
        {open && (
          <motion.div
            initial={{ height: 0, opacity: 0 }}
            animate={{ height: "auto", opacity: 1 }}
            exit={{ height: 0, opacity: 0 }}
            transition={{ duration: 0.2 }}
            className="overflow-hidden"
          >
            <div className="mt-3 max-h-72 overflow-y-auto">
              {loading ? (
                <p className="text-xs text-[var(--text-secondary)] text-center py-4">Lade...</p>
              ) : rows.length > 0 ? (
                <ExcelTable rows={rows} />
              ) : (
                <p className="text-xs text-[var(--text-secondary)] text-center py-4">Keine Daten</p>
              )}
            </div>
          </motion.div>
        )}
      </AnimatePresence>
    </div>
  );
}

// --- Desktop Excel Sidebar ---
function ExcelSidebar({ rows, loading }: { rows: ExcelRow[]; loading: boolean }) {
  return (
    <div className="glass-card">
      <h3 className="text-xs font-semibold uppercase tracking-wider text-[var(--text-secondary)] mb-3">
        GoDi-Plan (Excel)
      </h3>
      {loading ? (
        <div className="text-center py-8">
          <div className="inline-block w-5 h-5 border-2 border-[var(--accent)] border-t-transparent rounded-full animate-spin" />
          <p className="text-xs text-[var(--text-secondary)] mt-2">Lade Excel-Daten...</p>
        </div>
      ) : rows.length > 0 ? (
        <ExcelTable rows={rows} />
      ) : (
        <p className="text-xs text-[var(--text-secondary)] text-center py-8">Keine Daten</p>
      )}
    </div>
  );
}

function StatusDot({ status }: { status: "found" | "missing" | "empty" }) {
  return <span className={`status-dot ${status}`} />;
}

// --- Sortable Section Item ---
function SortableSection({
  section,
  detail,
  checked,
  onChange,
}: {
  section: SectionInfo;
  detail?: string;
  checked: boolean;
  onChange: (v: boolean) => void;
}) {
  const { attributes, listeners, setNodeRef, transform, transition, isDragging } =
    useSortable({ id: section.key });

  const style = {
    transform: CSS.Transform.toString(transform),
    transition,
    opacity: isDragging ? 0.5 : 1,
  };

  return (
    <div
      ref={setNodeRef}
      style={style}
      className="flex items-center justify-between py-2.5 px-1 group"
    >
      <div className="flex items-center gap-3 flex-1 min-w-0">
        {/* Drag Handle */}
        <button
          type="button"
          className="touch-none cursor-grab active:cursor-grabbing text-[var(--text-secondary)] opacity-40 group-hover:opacity-100 transition-opacity shrink-0"
          {...attributes}
          {...listeners}
        >
          <svg width="16" height="16" viewBox="0 0 24 24" fill="currentColor">
            <circle cx="9" cy="6" r="1.5" />
            <circle cx="15" cy="6" r="1.5" />
            <circle cx="9" cy="12" r="1.5" />
            <circle cx="15" cy="12" r="1.5" />
            <circle cx="9" cy="18" r="1.5" />
            <circle cx="15" cy="18" r="1.5" />
          </svg>
        </button>
        <div className="flex-1 min-w-0">
          <span className={`text-sm ${!checked ? "text-[var(--text-secondary)] line-through" : ""}`}>
            {section.label}
          </span>
          {detail && (
            <span className={`block text-xs truncate ${checked ? "text-[var(--text-secondary)]" : "text-[var(--text-secondary)]/50 line-through"}`}>
              {detail}
            </span>
          )}
        </div>
      </div>
      <button
        type="button"
        role="switch"
        aria-checked={checked}
        onClick={() => onChange(!checked)}
        className={`relative inline-flex h-6 w-11 shrink-0 rounded-full transition-colors duration-200 ml-3 ${
          checked ? "bg-[var(--success)]" : "bg-gray-300"
        }`}
      >
        <span
          className={`inline-block h-5 w-5 transform rounded-full bg-white shadow transition-transform duration-200 mt-0.5 ${
            checked ? "translate-x-5.5 ml-0.5" : "translate-x-0.5"
          }`}
        />
      </button>
    </div>
  );
}

// --- Extra Song State ---
interface ExtraSong {
  key: string;
  raw: string;
  searching: boolean;
  result: SongSearchResult | null;
}

// --- Song Slot Label ---
function songLabel(slot: string): string {
  if (slot.startsWith("song_extra")) return `+${slot.replace("song_extra", "")}`;
  return slot.replace("song", "");
}

// --- Extra Song Row with Live Search ---
function ExtraSongRow({
  extra,
  onSearch,
  onRemove,
}: {
  extra: ExtraSong;
  onSearch: (key: string, raw: string) => void;
  onRemove: (key: string) => void;
}) {
  const timerRef = useRef<ReturnType<typeof setTimeout> | null>(null);

  const handleChange = (value: string) => {
    if (timerRef.current) clearTimeout(timerRef.current);
    timerRef.current = setTimeout(() => {
      onSearch(extra.key, value);
    }, 500);
  };

  return (
    <div className="py-2 px-3 rounded-xl bg-white/50 space-y-1">
      <div className="flex items-center gap-3">
        <span className="text-xs font-medium text-[var(--accent)] w-6">
          {songLabel(extra.key)}
        </span>
        <StatusDot
          status={extra.searching ? "empty" : extra.result?.found ? "found" : extra.raw ? "missing" : "empty"}
        />
        <input
          className="input-field flex-1"
          defaultValue={extra.raw}
          placeholder="z.B. GLS 428 - Näher noch näher"
          onChange={(e) => handleChange(e.target.value)}
        />
        <button
          type="button"
          onClick={() => onRemove(extra.key)}
          className="w-7 h-7 flex items-center justify-center rounded-lg text-[var(--text-secondary)] hover:text-[var(--danger)] hover:bg-[var(--danger)]/10 transition-colors shrink-0"
          title="Lied entfernen"
        >
          <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
            <line x1="18" y1="6" x2="6" y2="18" /><line x1="6" y1="6" x2="18" y2="18" />
          </svg>
        </button>
      </div>
      {extra.searching && (
        <div className="flex items-center gap-2 ml-9">
          <span className="inline-block w-3 h-3 border border-[var(--accent)] border-t-transparent rounded-full animate-spin" />
          <span className="text-xs text-[var(--text-secondary)]">Suche...</span>
        </div>
      )}
      {!extra.searching && extra.result?.found && (
        <div className="flex items-center gap-2 ml-9">
          <span className="text-xs text-[var(--success)] font-medium">Gefunden:</span>
          <span className="text-xs text-[var(--text-secondary)]">{extra.result.file_name}</span>
        </div>
      )}
      {!extra.searching && extra.raw && extra.result && !extra.result.found && (
        <div className="ml-9">
          <span className="text-xs text-[var(--danger)]">Nicht in der Bibliothek gefunden</span>
        </div>
      )}
    </div>
  );
}

// --- Song Table ---
function SongTable({
  songs,
  extraSongs,
  onSongChange,
  onAddSong,
  onSearchExtra,
  onRemoveExtra,
}: {
  songs: SheetData["songs"];
  extraSongs: ExtraSong[];
  onSongChange?: (slot: string, value: string) => void;
  onAddSong?: () => void;
  onSearchExtra?: (key: string, raw: string) => void;
  onRemoveExtra?: (key: string) => void;
}) {
  return (
    <div className="space-y-2">
      {songs.map((s) => (
        <div
          key={s.slot_key}
          className="flex items-center gap-3 py-2 px-3 rounded-xl bg-white/50"
        >
          <span className="text-xs font-medium text-[var(--text-secondary)] w-6">
            {songLabel(s.slot_key)}
          </span>
          <StatusDot
            status={!s.raw ? "empty" : s.found ? "found" : "missing"}
          />
          <input
            className="input-field flex-1"
            defaultValue={s.raw}
            placeholder="z.B. FJ1 77 - Komm in unser dürres Leben"
            onChange={(e) => onSongChange?.(s.slot_key, e.target.value)}
          />
          {s.found && (
            <span className="text-xs text-[var(--text-secondary)] hidden sm:block truncate max-w-[180px]">
              {s.file_name}
            </span>
          )}
        </div>
      ))}
      {extraSongs.map((ex) => (
        <ExtraSongRow
          key={ex.key}
          extra={ex}
          onSearch={onSearchExtra!}
          onRemove={onRemoveExtra!}
        />
      ))}
      <button
        type="button"
        onClick={onAddSong}
        className="w-full py-2 text-xs font-medium text-[var(--accent)] hover:underline"
      >
        + Lied hinzufügen
      </button>
    </div>
  );
}

// --- Info Row ---
function InfoRow({
  label,
  value,
  onChange,
  multiline,
}: {
  label: string;
  value: string;
  onChange?: (v: string) => void;
  multiline?: boolean;
}) {
  return (
    <div className="flex items-start gap-4 py-2">
      <span className="text-sm font-medium text-[var(--text-secondary)] w-28 shrink-0">
        {label}
      </span>
      {multiline ? (
        <textarea
          className="input-field flex-1 min-h-[60px] resize-y"
          defaultValue={value}
          onChange={(e) => onChange?.(e.target.value)}
        />
      ) : (
        <input
          className="input-field flex-1"
          defaultValue={value}
          onChange={(e) => onChange?.(e.target.value)}
        />
      )}
    </div>
  );
}

// --- Main Page ---
export default function Home() {
  const [sheets, setSheets] = useState<Sheet[]>([]);
  const [sections, setSections] = useState<SectionInfo[]>([]);
  const [sectionOrder, setSectionOrder] = useState<string[]>([]);
  const [selected, setSelected] = useState<Sheet | null>(null);
  const [data, setData] = useState<SheetData | null>(null);
  const [loading, setLoading] = useState(false);
  const [generating, setGenerating] = useState(false);
  const [result, setResult] = useState<GenerateResult | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [overrides, setOverrides] = useState<Record<string, unknown>>({});
  const [songOverrides, setSongOverrides] = useState<Record<string, string>>({});
  const [uploadedImagePath, setUploadedImagePath] = useState<string | null>(null);
  const [sectionToggles, setSectionToggles] = useState<Record<string, boolean>>({});
  const [eventOverrides, setEventOverrides] = useState<InvitationEvent[] | null>(null);
  const [manualExtraSongs, setManualExtraSongs] = useState<ExtraSong[]>([]);
  const [textColor, setTextColor] = useState<"white" | "black">("white");
  const [excelRows, setExcelRows] = useState<ExcelRow[]>([]);
  const [excelLoading, setExcelLoading] = useState(false);
  const [excelOpen, setExcelOpen] = useState(false);
  // Layout-Positionen in % der Folie (aus PowerPoint: 9144000 x 6858000 EMU)
  // Titel: left=6.9% top=2.6% w=86.2% h=31.5% fontSize=66pt
  // Subtitle: left=10.6% top=83.7% w=78.9% h=10.1% fontSize=28pt
  const [titleLayout, setTitleLayout] = useState({ x: 6.9, y: 2.6, w: 86.2, h: 31.5, fontSize: 6.6 });
  const [subtitleLayout, setSubtitleLayout] = useState({ x: 10.6, y: 83.7, w: 78.9, h: 10.1, fontSize: 2.8 });

  const sensors = useSensors(
    useSensor(PointerSensor, { activationConstraint: { distance: 5 } }),
    useSensor(KeyboardSensor, { coordinateGetter: sortableKeyboardCoordinates })
  );

  // Load sheets + sections on mount
  useEffect(() => {
    getSheets()
      .then(setSheets)
      .catch((e) => setError(e.message));
    getSections()
      .then((secs) => {
        setSections(secs);
        setSectionOrder(secs.map((s) => s.key));
        const defaults: Record<string, boolean> = {};
        secs.forEach((s) => (defaults[s.key] = s.default_enabled));
        setSectionToggles(defaults);
      })
      .catch(() => {});
  }, []);

  // Load sheet data on selection
  const loadSheet = useCallback(
    async (sheet: Sheet) => {
      setSelected(sheet);
      setData(null);
      setResult(null);
      setError(null);
      setOverrides({});
      setSongOverrides({});
      setUploadedImagePath(null);
      setEventOverrides(null);
      setManualExtraSongs([]);
      setTextColor("white");
      setTitleLayout({ x: 6.9, y: 2.6, w: 86.2, h: 31.5, fontSize: 6.6 });
      setSubtitleLayout({ x: 10.6, y: 83.7, w: 78.9, h: 10.1, fontSize: 2.8 });
      setExcelRows([]);
      setExcelOpen(false);
      const defaults: Record<string, boolean> = {};
      sections.forEach((s) => (defaults[s.key] = s.default_enabled));
      setSectionToggles(defaults);
      setSectionOrder(sections.map((s) => s.key));
      setLoading(true);
      setExcelLoading(true);
      // Parallel laden
      getSheetRows(sheet.name, sheet.excel_path)
        .then(setExcelRows)
        .catch(() => {})
        .finally(() => setExcelLoading(false));
      try {
        const d = await getSheetData(sheet.name, sheet.excel_path);
        setData(d);

        // Extra-Songs aus der Excel in Sections einfügen
        const apiExtras = d.songs.filter((s) => s.slot_key.startsWith("song_extra"));
        if (apiExtras.length > 0) {
          const baseOrder = sections.map((s) => s.key);
          const einladungIdx = baseOrder.indexOf("einladung");
          const extraKeys = apiExtras.map((s) => s.slot_key);
          if (einladungIdx >= 0) {
            baseOrder.splice(einladungIdx, 0, ...extraKeys);
          } else {
            baseOrder.push(...extraKeys);
          }
          setSectionOrder(baseOrder);
          setSectionToggles((prev) => {
            const next = { ...prev };
            extraKeys.forEach((k) => (next[k] = true));
            return next;
          });
        }
      } catch (e) {
        setError((e as Error).message);
      } finally {
        setLoading(false);
      }
    },
    [sections]
  );

  const handleOverride = (key: string, value: string) => {
    setOverrides((prev) => ({ ...prev, [key]: value }));
  };

  const handleSongOverride = (slot: string, value: string) => {
    setSongOverrides((prev) => ({ ...prev, [slot]: value }));
  };

  const handleImageUpload = async (file: File) => {
    try {
      const path = await uploadImage(file);
      setUploadedImagePath(path);
    } catch (e) {
      setError((e as Error).message);
    }
  };

  const getEvents = (): InvitationEvent[] =>
    eventOverrides || data?.invitation_events || [];

  const updateEvent = (index: number, field: keyof InvitationEvent, value: string) => {
    const current = [...getEvents()];
    current[index] = { ...current[index], [field]: value };
    setEventOverrides(current);
  };

  const addEvent = () => {
    setEventOverrides([...getEvents(), { date_str: "", time_str: "", event_name: "", note: "" }]);
  };

  const removeEvent = (index: number) => {
    const current = [...getEvents()];
    current.splice(index, 1);
    setEventOverrides(current);
  };

  const addManualSong = () => {
    const apiExtras = data?.songs.filter((s) => s.slot_key.startsWith("song_extra")).length || 0;
    const nextNum = apiExtras + manualExtraSongs.length + 1;
    const key = `song_extra${nextNum}`;
    setManualExtraSongs((prev) => [...prev, { key, raw: "", searching: false, result: null }]);
    // Automatisch in Folien-Abschnitte einfügen (vor einladung)
    setSectionOrder((prev) => {
      const idx = prev.indexOf("einladung");
      if (idx >= 0) {
        const next = [...prev];
        next.splice(idx, 0, key);
        return next;
      }
      return [...prev, key];
    });
    setSectionToggles((prev) => ({ ...prev, [key]: true }));
  };

  const removeManualSong = (key: string) => {
    setManualExtraSongs((prev) => prev.filter((s) => s.key !== key));
    setSongOverrides((prev) => {
      const next = { ...prev };
      delete next[key];
      return next;
    });
    setSectionOrder((prev) => prev.filter((k) => k !== key));
    setSectionToggles((prev) => {
      const next = { ...prev };
      delete next[key];
      return next;
    });
  };

  const handleExtraSongSearch = async (key: string, raw: string) => {
    // Update raw + set searching
    setManualExtraSongs((prev) =>
      prev.map((s) => (s.key === key ? { ...s, raw, searching: true, result: null } : s))
    );
    setSongOverrides((prev) => ({ ...prev, [key]: raw }));

    if (!raw.trim()) {
      setManualExtraSongs((prev) =>
        prev.map((s) => (s.key === key ? { ...s, raw, searching: false, result: null } : s))
      );
      return;
    }

    try {
      const result = await searchSong(raw);
      setManualExtraSongs((prev) =>
        prev.map((s) => (s.key === key ? { ...s, searching: false, result } : s))
      );
    } catch {
      setManualExtraSongs((prev) =>
        prev.map((s) => (s.key === key ? { ...s, searching: false, result: null } : s))
      );
    }
  };

  const handleDragEnd = (event: DragEndEvent) => {
    const { active, over } = event;
    if (over && active.id !== over.id) {
      setSectionOrder((prev) => {
        const oldIndex = prev.indexOf(active.id as string);
        const newIndex = prev.indexOf(over.id as string);
        return arrayMove(prev, oldIndex, newIndex);
      });
    }
  };

  const handleGenerate = async () => {
    if (!selected || !data) return;
    setGenerating(true);
    setResult(null);
    setError(null);

    const finalOverrides: Record<string, unknown> = { ...overrides };

    if (Object.keys(songOverrides).length > 0) {
      finalOverrides.songs = songOverrides;
    }
    if (uploadedImagePath) {
      finalOverrides.image_path = uploadedImagePath;
    }
    if (eventOverrides) {
      finalOverrides.invitation_events = eventOverrides;
    }

    const disabledSections = Object.entries(sectionToggles)
      .filter(([, enabled]) => !enabled)
      .map(([key]) => key);

    try {
      const res = await generate({
        sheet_name: selected.name,
        excel_path: selected.excel_path,
        overrides: Object.keys(finalOverrides).length > 0 ? finalOverrides : undefined,
        fetch_bible: true,
        disabled_sections: disabledSections.length > 0 ? disabledSections : undefined,
        section_order: sectionOrder,
        text_color: textColor,
        title_layout: titleLayout,
        subtitle_layout: subtitleLayout,
      });
      setResult(res);
    } catch (e) {
      setError((e as Error).message);
    } finally {
      setGenerating(false);
    }
  };

  // Detail-Info für jeden Abschnitt aus den geladenen Daten
  const getSectionDetail = (key: string): string | undefined => {
    if (!data) return undefined;
    const songMap: Record<string, string> = {};
    data.songs.forEach((s) => {
      songMap[s.slot_key] = s.raw || "";
    });
    switch (key) {
      case "begruessung":
        return data.greeting_verse || undefined;
      case "song1":
        return songMap.song1 || undefined;
      case "song2":
        return songMap.song2 || undefined;
      case "song3":
        return songMap.song3 || undefined;
      case "kinderstunde": {
        const s3 = songMap.song3;
        return s3 ? `Kinderlied: ${s3}` : undefined;
      }
      case "song4":
        return songMap.song4 || undefined;
      case "song5":
        return songMap.song5 || undefined;
      case "song6":
        return songMap.song6 || undefined;
      case "song7":
        return songMap.song7 || undefined;
      case "lesung":
        return data.lesung_reference || undefined;
      case "predigt1":
        return [data.predigt1_title, data.predigt1_reference].filter(Boolean).join(" — ") || undefined;
      case "predigt2":
        return [data.predigt2_title, data.predigt2_reference].filter(Boolean).join(" — ") || undefined;
      case "einladung":
        return data.invitation_events.length > 0
          ? `${data.invitation_events.length} Termine`
          : undefined;
      default: {
        // Extra-Song Details (aus Excel oder manuell)
        if (key.startsWith("song_extra")) {
          // Erst in API-Songs schauen
          const apiSong = data.songs.find((s) => s.slot_key === key);
          if (apiSong) return apiSong.found ? apiSong.file_name : apiSong.raw;
          // Dann in manuellen Extras
          const extra = manualExtraSongs.find((s) => s.key === key);
          if (extra?.result?.found) return extra.result.file_name;
          if (extra?.raw) return extra.raw;
        }
        return undefined;
      }
    }
  };

  // Alle Sections: API + Extra-Songs aus Excel + manuell hinzugefügte
  const apiExtraSongs = data?.songs.filter((s) => s.slot_key.startsWith("song_extra")) || [];
  const allSections: SectionInfo[] = [
    ...sections,
    ...apiExtraSongs.map((s) => ({
      key: s.slot_key,
      label: s.found
        ? `Extra: ${s.file_name.replace(/\.[^/.]+$/, "")}`
        : `Extra: ${s.raw}`,
      default_enabled: true,
    })),
    ...manualExtraSongs.map((ex) => ({
      key: ex.key,
      label: ex.result?.found
        ? `Extra: ${ex.result.file_name.replace(/\.[^/.]+$/, "")}`
        : ex.raw
        ? `Extra: ${ex.raw}`
        : "Extra-Lied (leer)",
      default_enabled: true,
    })),
  ];

  const orderedSections = sectionOrder
    .map((key) => allSections.find((s) => s.key === key))
    .filter((s): s is SectionInfo => !!s);

  return (
    <div className="max-w-6xl mx-auto px-4 py-6 sm:py-10 md:flex md:gap-6">
    <div className="flex-1 max-w-2xl min-w-0">
      {/* Header */}
      <header className="mb-8">
        <div className="flex items-center gap-3 mb-5">
          <img src="/logo.jpg" alt="Gemeindelogo" className="h-10 w-10 rounded-full object-cover shadow-sm" />
          <div>
            <h1 className="text-2xl font-bold tracking-tight leading-none">AutoPräsi</h1>
            <p className="text-xs text-[var(--text-secondary)] mt-0.5 tracking-wide uppercase">
              Ev.-luth. Christus-Brüdergemeinde
            </p>
          </div>
        </div>
        <select
          className="input-field w-full text-base"
          value={selected?.name || ""}
          onChange={(e) => {
            const s = sheets.find((sh) => sh.name === e.target.value);
            if (s) loadSheet(s);
          }}
        >
          <option value="" disabled>
            Gottesdienst wählen...
          </option>
          {sheets.map((s) => (
            <option key={`${s.name}-${s.excel_path}`} value={s.name}>
              {s.name}
            </option>
          ))}
        </select>
      </header>

      {/* Mobile Excel Preview */}
      {selected && data && !loading && (
        <MobileExcelCard rows={excelRows} loading={excelLoading} open={excelOpen} onToggle={() => setExcelOpen((o) => !o)} />
      )}

      {/* Loading */}
      {loading && (
        <div className="glass-card text-center py-12">
          <div className="inline-block w-6 h-6 border-2 border-[var(--accent)] border-t-transparent rounded-full animate-spin" />
          <p className="text-sm text-[var(--text-secondary)] mt-3">Lade Daten...</p>
        </div>
      )}

      {/* Content */}
      <AnimatePresence mode="wait">
        {data && selected && !loading && (
          <motion.div
            key={selected.name}
            initial={{ opacity: 0, y: 12 }}
            animate={{ opacity: 1, y: 0 }}
            exit={{ opacity: 0, y: -12 }}
            transition={{ duration: 0.25, ease: "easeOut" }}
            className="space-y-4"
          >
            {/* Header Card */}
            <div className="glass-card">
              <div className="flex items-center justify-between mb-1">
                <h2 className="text-xl font-semibold">{selected.name}</h2>
              </div>
              <input
                className="input-field mt-1 text-sm"
                defaultValue={data.service_header}
                placeholder="z.B. Passionsandacht am 18.03.2026"
                onChange={(e) => handleOverride("service_header", e.target.value)}
              />
            </div>

            {/* Sections Toggle Card with Drag & Drop */}
            <div className="glass-card">
              <h3 className="text-xs font-semibold uppercase tracking-wider text-[var(--text-secondary)] mb-3">
                Folien-Abschnitte
              </h3>
              <DndContext
                sensors={sensors}
                collisionDetection={closestCenter}
                onDragEnd={handleDragEnd}
              >
                <SortableContext
                  items={sectionOrder}
                  strategy={verticalListSortingStrategy}
                >
                  <div className="divide-y divide-[var(--card-border)]">
                    {orderedSections.map((sec) => (
                      <SortableSection
                        key={sec.key}
                        section={sec}
                        detail={getSectionDetail(sec.key)}
                        checked={sectionToggles[sec.key] ?? true}
                        onChange={(v) =>
                          setSectionToggles((prev) => ({ ...prev, [sec.key]: v }))
                        }
                      />
                    ))}
                  </div>
                </SortableContext>
              </DndContext>
            </div>

            {/* Overview Card */}
            <div className="glass-card">
              <h3 className="text-xs font-semibold uppercase tracking-wider text-[var(--text-secondary)] mb-3">
                Übersicht
              </h3>
              <InfoRow label="Thema" value={data.theme} onChange={(v) => handleOverride("theme", v)} />
              <InfoRow label="Begrüßung" value={data.greeting_verse} multiline onChange={(v) => handleOverride("greeting_verse", v)} />
              <InfoRow label="Lesung" value={data.lesung_reference} onChange={(v) => handleOverride("lesung_reference", v)} />

              {/* Predigt 1 — Thema + Bibelstelle */}
              <div className="py-2 space-y-1">
                <span className="text-xs font-semibold uppercase tracking-wider text-[var(--text-secondary)]">
                  Predigt 1
                </span>
                <div className="flex items-start gap-4">
                  <span className="text-sm font-medium text-[var(--text-secondary)] w-28 shrink-0">Thema</span>
                  <input
                    className="input-field flex-1"
                    defaultValue={data.predigt1_title}
                    placeholder="Predigtthema"
                    onChange={(e) => handleOverride("predigt1_title", e.target.value)}
                  />
                </div>
                <div className="flex items-start gap-4">
                  <span className="text-sm font-medium text-[var(--text-secondary)] w-28 shrink-0">Predigt zu</span>
                  <input
                    className="input-field flex-1"
                    defaultValue={data.predigt1_reference}
                    placeholder="Bibelstelle"
                    onChange={(e) => handleOverride("predigt1_reference", e.target.value)}
                  />
                </div>
              </div>

              {/* Predigt 2 — Thema + Bibelstelle */}
              <div className="py-2 space-y-1">
                <span className="text-xs font-semibold uppercase tracking-wider text-[var(--text-secondary)]">
                  Predigt 2
                </span>
                <div className="flex items-start gap-4">
                  <span className="text-sm font-medium text-[var(--text-secondary)] w-28 shrink-0">Thema</span>
                  <input
                    className="input-field flex-1"
                    defaultValue={data.predigt2_title}
                    placeholder="Predigtthema"
                    onChange={(e) => handleOverride("predigt2_title", e.target.value)}
                  />
                </div>
                <div className="flex items-start gap-4">
                  <span className="text-sm font-medium text-[var(--text-secondary)] w-28 shrink-0">Predigt zu</span>
                  <input
                    className="input-field flex-1"
                    defaultValue={data.predigt2_reference}
                    placeholder="Bibelstelle"
                    onChange={(e) => handleOverride("predigt2_reference", e.target.value)}
                  />
                </div>
              </div>

              {data.is_abendmahl && (
                <div className="mt-2 text-xs font-medium text-[var(--accent)]">
                  Abendmahl
                </div>
              )}
            </div>

            {/* Songs Card */}
            <div className="glass-card">
              <h3 className="text-xs font-semibold uppercase tracking-wider text-[var(--text-secondary)] mb-3">
                Lieder
              </h3>
              <SongTable
                songs={data.songs}
                extraSongs={manualExtraSongs}
                onSongChange={handleSongOverride}
                onAddSong={addManualSong}
                onSearchExtra={handleExtraSongSearch}
                onRemoveExtra={removeManualSong}
              />
            </div>

            {/* Herzliche Einladung Tabelle */}
            <div className="glass-card">
              <div className="flex items-center justify-between mb-3">
                <h3 className="text-xs font-semibold uppercase tracking-wider text-[var(--text-secondary)]">
                  Herzliche Einladung
                </h3>
                <button
                  type="button"
                  onClick={addEvent}
                  className="text-xs font-medium text-[var(--accent)] hover:underline"
                >
                  + Zeile hinzufügen
                </button>
              </div>
              <div className="grid grid-cols-[0.9fr_0.5fr_1.5fr_1fr_auto] gap-2 mb-2 px-1">
                <span className="text-xs font-semibold text-[var(--text-secondary)]">Datum</span>
                <span className="text-xs font-semibold text-[var(--text-secondary)]">Uhrzeit</span>
                <span className="text-xs font-semibold text-[var(--text-secondary)]">Veranstaltung</span>
                <span className="text-xs font-semibold text-[var(--text-secondary)]">Hinweis</span>
                <span className="w-7" />
              </div>
              <div className="space-y-1">
                {getEvents().map((evt, i) => (
                  <div key={i} className="grid grid-cols-[0.9fr_0.5fr_1.5fr_1fr_auto] gap-2 items-center">
                    <input
                      className="input-field text-sm"
                      value={evt.date_str}
                      placeholder="Di 17.03.26"
                      onChange={(e) => updateEvent(i, "date_str", e.target.value)}
                    />
                    <input
                      className="input-field text-sm"
                      value={evt.time_str}
                      placeholder="19:00"
                      onChange={(e) => updateEvent(i, "time_str", e.target.value)}
                    />
                    <input
                      className="input-field text-sm"
                      value={evt.event_name}
                      placeholder="Gebetsstunde"
                      onChange={(e) => updateEvent(i, "event_name", e.target.value)}
                    />
                    <input
                      className="input-field text-sm"
                      value={evt.note}
                      placeholder="z.B. fällt aus"
                      onChange={(e) => updateEvent(i, "note", e.target.value)}
                    />
                    <button
                      type="button"
                      onClick={() => removeEvent(i)}
                      className="w-7 h-7 flex items-center justify-center rounded-lg text-[var(--text-secondary)] hover:text-[var(--danger)] hover:bg-[var(--danger)]/10 transition-colors"
                      title="Zeile entfernen"
                    >
                      <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
                        <line x1="18" y1="6" x2="6" y2="18" /><line x1="6" y1="6" x2="18" y2="18" />
                      </svg>
                    </button>
                  </div>
                ))}
              </div>
              {getEvents().length === 0 && (
                <p className="text-sm text-[var(--text-secondary)] text-center py-4">
                  Keine Einträge. Klicke &quot;+ Zeile hinzufügen&quot; um einen Termin einzutragen.
                </p>
              )}
            </div>

            {/* Folien-Vorschau — exakte PowerPoint-Positionen */}
            <div className="glass-card">
              <h3 className="text-xs font-semibold uppercase tracking-wider text-[var(--text-secondary)] mb-3">
                Folien-Vorschau
              </h3>

              <SlidePreview
                imageUrl={data.image_found && data.image_path ? imagePreviewUrl(data.image_path) : null}
                theme={(overrides.theme as string) || data.theme || ""}
                dateStr={data.date_str}
                kirchenkalender={data.kirchenkalender}
                textColor={textColor}
                titleLayout={titleLayout}
                subtitleLayout={subtitleLayout}
                onTitleLayoutChange={setTitleLayout}
                onSubtitleLayoutChange={setSubtitleLayout}
              />

              {/* Controls */}
              <div className="flex items-center gap-4 mt-3">
                <div className="flex items-center gap-2">
                  <span className="text-xs text-[var(--text-secondary)]">Textfarbe:</span>
                  <button
                    type="button"
                    onClick={() => setTextColor("white")}
                    className={`w-7 h-7 rounded-full border-2 bg-white transition-all ${
                      textColor === "white" ? "border-[var(--accent)] scale-110 shadow-md" : "border-gray-300"
                    }`}
                    title="Weiß"
                  />
                  <button
                    type="button"
                    onClick={() => setTextColor("black")}
                    className={`w-7 h-7 rounded-full border-2 bg-black transition-all ${
                      textColor === "black" ? "border-[var(--accent)] scale-110 shadow-md" : "border-gray-300"
                    }`}
                    title="Schwarz"
                  />
                </div>
                <div className="flex items-center gap-2">
                  <span className="text-xs text-[var(--text-secondary)]">Titel:</span>
                  <button type="button" onClick={() => setTitleLayout(l => ({...l, fontSize: Math.max(3, l.fontSize - 0.5)}))} className="text-xs px-1.5 py-0.5 rounded bg-gray-100 hover:bg-gray-200">A-</button>
                  <span className="text-xs text-[var(--text-secondary)] w-6 text-center">{titleLayout.fontSize.toFixed(1)}</span>
                  <button type="button" onClick={() => setTitleLayout(l => ({...l, fontSize: Math.min(12, l.fontSize + 0.5)}))} className="text-xs px-1.5 py-0.5 rounded bg-gray-100 hover:bg-gray-200">A+</button>
                </div>
                <div className="flex items-center gap-2">
                  <span className="text-xs text-[var(--text-secondary)]">Datum:</span>
                  <button type="button" onClick={() => setSubtitleLayout(l => ({...l, fontSize: Math.max(1, l.fontSize - 0.25)}))} className="text-xs px-1.5 py-0.5 rounded bg-gray-100 hover:bg-gray-200">A-</button>
                  <span className="text-xs text-[var(--text-secondary)] w-6 text-center">{subtitleLayout.fontSize.toFixed(1)}</span>
                  <button type="button" onClick={() => setSubtitleLayout(l => ({...l, fontSize: Math.min(6, l.fontSize + 0.25)}))} className="text-xs px-1.5 py-0.5 rounded bg-gray-100 hover:bg-gray-200">A+</button>
                </div>
              </div>

              {/* Bild Upload */}
              <label className="mt-3 block">
                <input
                  type="file"
                  accept="image/*"
                  className="hidden"
                  onChange={(e) => {
                    const f = e.target.files?.[0];
                    if (f) handleImageUpload(f);
                  }}
                />
                <div className="mt-1 border-2 border-dashed border-[var(--card-border)] rounded-xl p-3 text-center text-sm text-[var(--text-secondary)] cursor-pointer hover:border-[var(--accent)] transition-colors">
                  {data.image_found || uploadedImagePath ? "Anderes Bild wählen..." : "Bild auswählen..."}
                </div>
              </label>
            </div>

            {/* Generate Button */}
            <button
              className="btn-primary"
              disabled={generating}
              onClick={handleGenerate}
            >
              {generating ? (
                <span className="flex items-center justify-center gap-2">
                  <span className="inline-block w-4 h-4 border-2 border-white border-t-transparent rounded-full animate-spin" />
                  Generiere...
                </span>
              ) : (
                "Präsentation generieren"
              )}
            </button>

            {/* Error */}
            {error && (
              <motion.div
                initial={{ opacity: 0 }}
                animate={{ opacity: 1 }}
                className="glass-card border-[var(--danger)]/30 bg-[var(--danger)]/5"
              >
                <p className="text-sm text-[var(--danger)]">{error}</p>
              </motion.div>
            )}

            {/* Result */}
            {result && (
              <motion.div
                initial={{ opacity: 0, y: 8 }}
                animate={{ opacity: 1, y: 0 }}
                className="glass-card border-[var(--success)]/30 bg-[var(--success)]/5"
              >
                <div className="flex items-center gap-2 mb-2">
                  <StatusDot status="found" />
                  <span className="text-sm font-medium">Präsentation erstellt</span>
                </div>
                <a
                  href={downloadUrl(result.output_name)}
                  className="inline-block mt-1 text-sm text-[var(--accent)] font-medium hover:underline"
                  download
                >
                  {result.output_name} herunterladen
                </a>
                {result.missing_songs.length > 0 && (
                  <p className="text-xs text-[var(--warning)] mt-2">
                    Fehlende Lieder: {result.missing_songs.join(", ")}
                  </p>
                )}
              </motion.div>
            )}
          </motion.div>
        )}
      </AnimatePresence>

      {/* Empty State */}
      {!selected && !loading && (
        <div className="glass-card text-center py-16">
          <p className="text-lg text-[var(--text-secondary)]">
            Wähle einen Gottesdienst aus um zu beginnen
          </p>
        </div>
      )}
    </div>

    {/* Desktop Excel Sidebar */}
    {selected && data && !loading && (
      <aside className="hidden md:block w-80 lg:w-96 shrink-0">
        <div className="sticky top-6 max-h-[calc(100vh-3rem)] overflow-y-auto rounded-2xl">
          <ExcelSidebar rows={excelRows} loading={excelLoading} />
        </div>
      </aside>
    )}
    </div>
  );
}
