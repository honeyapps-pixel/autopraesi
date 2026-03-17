"use client";

import { useState, useEffect, useCallback } from "react";
import { motion, AnimatePresence } from "framer-motion";
import {
  Sheet,
  SheetData,
  SectionInfo,
  getSheets,
  getSections,
  getSheetData,
  generate,
  uploadImage,
  downloadUrl,
  GenerateResult,
} from "@/lib/api";

function StatusDot({ status }: { status: "found" | "missing" | "empty" }) {
  return <span className={`status-dot ${status}`} />;
}

// --- Toggle Switch ---
function Toggle({
  label,
  checked,
  onChange,
}: {
  label: string;
  checked: boolean;
  onChange: (v: boolean) => void;
}) {
  return (
    <label className="flex items-center justify-between py-2 cursor-pointer">
      <span className="text-sm">{label}</span>
      <button
        type="button"
        role="switch"
        aria-checked={checked}
        onClick={() => onChange(!checked)}
        className={`relative inline-flex h-6 w-11 shrink-0 rounded-full transition-colors duration-200 ${
          checked ? "bg-[var(--success)]" : "bg-gray-300"
        }`}
      >
        <span
          className={`inline-block h-5 w-5 transform rounded-full bg-white shadow transition-transform duration-200 mt-0.5 ${
            checked ? "translate-x-5.5 ml-0.5" : "translate-x-0.5"
          }`}
        />
      </button>
    </label>
  );
}

// --- Song Table ---
function SongTable({
  songs,
  editable,
  onSongChange,
}: {
  songs: SheetData["songs"];
  editable: boolean;
  onSongChange?: (slot: string, value: string) => void;
}) {
  return (
    <div className="space-y-2">
      {songs.map((s) => (
        <div
          key={s.slot_key}
          className="flex items-center gap-3 py-2 px-3 rounded-xl bg-white/50"
        >
          <span className="text-xs font-medium text-[var(--text-secondary)] w-6">
            {s.slot_key.replace("song", "")}
          </span>
          <StatusDot
            status={!s.raw ? "empty" : s.found ? "found" : "missing"}
          />
          {editable ? (
            <input
              className="input-field flex-1"
              defaultValue={s.raw}
              placeholder="z.B. FJ1 77 - Komm in unser dürres Leben"
              onChange={(e) => onSongChange?.(s.slot_key, e.target.value)}
            />
          ) : (
            <span className="flex-1 text-sm truncate">
              {s.raw || <span className="text-[var(--text-secondary)]">—</span>}
            </span>
          )}
          {s.found && (
            <span className="text-xs text-[var(--text-secondary)] hidden sm:block truncate max-w-[180px]">
              {s.file_name}
            </span>
          )}
        </div>
      ))}
    </div>
  );
}

// --- Info Row ---
function InfoRow({
  label,
  value,
  editable,
  onChange,
  multiline,
}: {
  label: string;
  value: string;
  editable: boolean;
  onChange?: (v: string) => void;
  multiline?: boolean;
}) {
  return (
    <div className="flex items-start gap-4 py-2">
      <span className="text-sm font-medium text-[var(--text-secondary)] w-28 shrink-0">
        {label}
      </span>
      {editable ? (
        multiline ? (
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
        )
      ) : (
        <span className="text-sm flex-1 whitespace-pre-line">{value || "—"}</span>
      )}
    </div>
  );
}

// --- Main Page ---
export default function Home() {
  const [sheets, setSheets] = useState<Sheet[]>([]);
  const [sections, setSections] = useState<SectionInfo[]>([]);
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
  const [announcementOverrides, setAnnouncementOverrides] = useState<string[] | null>(null);

  // Load sheets + sections on mount
  useEffect(() => {
    getSheets()
      .then(setSheets)
      .catch((e) => setError(e.message));
    getSections()
      .then((secs) => {
        setSections(secs);
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
      setAnnouncementOverrides(null);
      // Reset section toggles to defaults
      const defaults: Record<string, boolean> = {};
      sections.forEach((s) => (defaults[s.key] = s.default_enabled));
      setSectionToggles(defaults);
      setLoading(true);
      try {
        const d = await getSheetData(sheet.name, sheet.excel_path);
        setData(d);
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

  const handleAnnouncementChange = (index: number, value: string) => {
    const current = announcementOverrides || [...(data?.announcements || [])];
    const updated = [...current];
    updated[index] = value;
    setAnnouncementOverrides(updated);
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
    if (announcementOverrides) {
      finalOverrides.announcements = announcementOverrides;
    }

    // Disabled sections
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
      });
      setResult(res);
    } catch (e) {
      setError((e as Error).message);
    } finally {
      setGenerating(false);
    }
  };

  const announcements = announcementOverrides || data?.announcements || [];

  return (
    <div className="max-w-2xl mx-auto px-4 py-8 sm:py-12">
      {/* Header */}
      <div className="flex items-center justify-between mb-8">
        <div>
          <h1 className="text-3xl font-semibold tracking-tight">AutoPräsi</h1>
          <p className="text-sm text-[var(--text-secondary)] mt-1">
            Gottesdienst-Präsentation
          </p>
        </div>

        {/* Sheet Selector */}
        <select
          className="input-field w-auto min-w-[160px]"
          value={selected?.name || ""}
          onChange={(e) => {
            const s = sheets.find((sh) => sh.name === e.target.value);
            if (s) loadSheet(s);
          }}
        >
          <option value="" disabled>
            Sheet wählen...
          </option>
          {sheets.map((s) => (
            <option key={`${s.name}-${s.excel_path}`} value={s.name}>
              {s.name}
            </option>
          ))}
        </select>
      </div>

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
            {/* Header Card — editierbarer Service Header */}
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

            {/* Sections Toggle Card (immer sichtbar) */}
            <div className="glass-card">
              <h3 className="text-xs font-semibold uppercase tracking-wider text-[var(--text-secondary)] mb-3">
                Folien-Abschnitte
              </h3>
              <div className="divide-y divide-[var(--card-border)]">
                {sections.map((sec) => (
                  <Toggle
                    key={sec.key}
                    label={sec.label}
                    checked={sectionToggles[sec.key] ?? true}
                    onChange={(v) =>
                      setSectionToggles((prev) => ({ ...prev, [sec.key]: v }))
                    }
                  />
                ))}
              </div>
            </div>

            {/* Overview Card */}
            <div className="glass-card">
              <h3 className="text-xs font-semibold uppercase tracking-wider text-[var(--text-secondary)] mb-3">
                Übersicht
              </h3>
              <InfoRow label="Thema" value={data.theme} editable onChange={(v) => handleOverride("theme", v)} />
              <InfoRow label="Begrüßung" value={data.greeting_verse} editable multiline onChange={(v) => handleOverride("greeting_verse", v)} />
              <InfoRow label="Lesung" value={data.lesung_reference} editable onChange={(v) => handleOverride("lesung_reference", v)} />
              <InfoRow
                label="Predigt 1"
                value={`${data.predigt1_reference}${data.predigt1_title ? " – " + data.predigt1_title : ""}`}
                editable
                onChange={(v) => handleOverride("predigt1_reference", v)}
              />
              <InfoRow
                label="Predigt 2"
                value={`${data.predigt2_reference}${data.predigt2_title ? " – " + data.predigt2_title : ""}`}
                editable
                onChange={(v) => handleOverride("predigt2_reference", v)}
              />
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
                editable
                onSongChange={handleSongOverride}
              />
            </div>

            {/* Announcements Card */}
            {announcements.length > 0 && (
              <div className="glass-card">
                <h3 className="text-xs font-semibold uppercase tracking-wider text-[var(--text-secondary)] mb-3">
                  Abkündigungen
                </h3>
                <div className="space-y-2">
                  {announcements.map((a, i) => (
                    <div key={i} className="py-1">
                        <input
                          className="input-field"
                          defaultValue={a}
                          onChange={(e) => handleAnnouncementChange(i, e.target.value)}
                        />
                    </div>
                  ))}
                </div>
              </div>
            )}

            {/* Image Card */}
            <div className="glass-card">
              <h3 className="text-xs font-semibold uppercase tracking-wider text-[var(--text-secondary)] mb-3">
                Hintergrundbild
              </h3>
              <div className="flex items-center gap-3">
                <StatusDot status={data.image_found || uploadedImagePath ? "found" : "missing"} />
                <span className="text-sm">
                  {uploadedImagePath
                    ? "Bild hochgeladen"
                    : data.image_found
                    ? "Automatisch gefunden"
                    : "Nicht gefunden"}
                </span>
              </div>
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
                  <div className="mt-2 border-2 border-dashed border-[var(--card-border)] rounded-xl p-4 text-center text-sm text-[var(--text-secondary)] cursor-pointer hover:border-[var(--accent)] transition-colors">
                    {uploadedImagePath ? "Anderes Bild wählen..." : "Bild auswählen..."}
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
            Wähle ein Sheet aus um zu beginnen
          </p>
        </div>
      )}
    </div>
  );
}
