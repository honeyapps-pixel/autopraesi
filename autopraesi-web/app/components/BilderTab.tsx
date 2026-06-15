"use client";

import { useState, useEffect, useCallback, useRef } from "react";
import { motion, AnimatePresence } from "framer-motion";
import {
  Sheet,
  getSheets,
  getCurrentQuarter,
  getSheetData,
  saveSundayImage,
} from "@/lib/api";
import {
  GenImage,
  PromptInput,
  generate as genImages,
  regenerate as regenImage,
  deleteImage as delImage,
  pollStatus,
  fetchImageObjectUrl,
  health as genHealth,
} from "@/lib/imagegen";

type Candidate = GenImage & { objectUrl?: string };

/** Konvertiert eine (lokale blob:) Bild-URL im Browser nach JPEG – für den Upload. */
async function objectUrlToJpegBlob(url: string): Promise<Blob> {
  const res = await fetch(url);
  if (!res.ok) throw new Error("Bild konnte nicht geladen werden");
  const bitmap = await createImageBitmap(await res.blob());
  const canvas = document.createElement("canvas");
  canvas.width = bitmap.width;
  canvas.height = bitmap.height;
  const ctx = canvas.getContext("2d");
  if (!ctx) throw new Error("Canvas nicht verfügbar");
  ctx.drawImage(bitmap, 0, 0);
  bitmap.close();
  return await new Promise<Blob>((resolve, reject) =>
    canvas.toBlob(
      (b) => (b ? resolve(b) : reject(new Error("JPEG-Konvertierung fehlgeschlagen"))),
      "image/jpeg",
      0.9
    )
  );
}

export default function BilderTab() {
  const [sheets, setSheets] = useState<Sheet[]>([]);
  const [quarterPattern, setQuarterPattern] = useState("");
  const [showAllQuarters, setShowAllQuarters] = useState(false);
  const [selected, setSelected] = useState<Sheet | null>(null);

  // Aus der Excel vorbefüllt
  const [theme, setTheme] = useState("");
  const [wochenspruch, setWochenspruch] = useState("");
  const [freitext, setFreitext] = useState("");
  const [dateStr, setDateStr] = useState("");

  const [count, setCount] = useState(2);
  const [images, setImages] = useState<Candidate[]>([]);
  const [generating, setGenerating] = useState(false);
  const [confirmingId, setConfirmingId] = useState<string | null>(null);
  // „Bearbeiten": welcher Kandidat gerade ein Eingabefeld zeigt + dessen Text
  const [editingId, setEditingId] = useState<string | null>(null);
  const [editText, setEditText] = useState("");
  const [savedName, setSavedName] = useState<string | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [loadingSheet, setLoadingSheet] = useState(false);

  // Generator-Erreichbarkeit
  const [online, setOnline] = useState<boolean | null>(null);
  const [checking, setChecking] = useState(false);

  useEffect(() => {
    getCurrentQuarter().then(setQuarterPattern).catch(() => {});
    getSheets().then(setSheets).catch((e) => setError(e.message));
  }, []);

  const checkHealth = useCallback(async () => {
    setChecking(true);
    try {
      setOnline(await genHealth());
    } finally {
      setChecking(false);
    }
  }, []);

  useEffect(() => {
    checkHealth();
  }, [checkHealth]);

  const loadSheet = useCallback(async (sheet: Sheet) => {
    setSelected(sheet);
    setLoadingSheet(true);
    setImages([]);
    setSavedName(null);
    setError(null);
    try {
      const d = await getSheetData(sheet.name, sheet.excel_path);
      setTheme(d.theme || "");
      setWochenspruch(d.greeting_verse || "");
      setDateStr(d.date_str || "");
      setFreitext("");
    } catch (e) {
      setError((e as Error).message);
    } finally {
      setLoadingSheet(false);
    }
  }, []);

  const promptInput = (): PromptInput => ({ theme, wochenspruch, freitext });

  // Solange Bilder "pending" sind, regelmäßig den Status pollen (~1 Min/Bild).
  const pollKey = images.map((i) => `${i.id}:${i.status}`).join(",");
  useEffect(() => {
    const pendingIds = images.filter((i) => i.status === "pending").map((i) => i.id);
    if (pendingIds.length === 0) return;
    let active = true;
    const tick = async () => {
      try {
        const states = await pollStatus(pendingIds);
        if (!active) return;
        setImages((prev) =>
          prev.map((img) => {
            const s = states.find((x) => x.id === img.id);
            return s && s.status !== img.status ? { ...img, status: s.status, error: s.error } : img;
          })
        );
      } catch {
        /* nächster Tick versucht es erneut */
      }
    };
    const iv = setInterval(tick, 2500);
    tick();
    return () => {
      active = false;
      clearInterval(iv);
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [pollKey]);

  // Fertige Bilder als Blob (mit ngrok-Header) laden und als Object-URL anzeigen.
  const loadingBlobs = useRef<Set<string>>(new Set());
  useEffect(() => {
    images.forEach((img) => {
      if (img.status === "done" && !img.objectUrl && !loadingBlobs.current.has(img.id)) {
        loadingBlobs.current.add(img.id);
        fetchImageObjectUrl(img.id)
          .then((url) => setImages((prev) => prev.map((x) => (x.id === img.id ? { ...x, objectUrl: url } : x))))
          .catch(() => {})
          .finally(() => loadingBlobs.current.delete(img.id));
      }
    });
  }, [images]);

  const handleGenerate = async () => {
    setError(null);
    setSavedName(null);
    setGenerating(true);
    try {
      const imgs = await genImages(promptInput(), count);
      setImages((prev) => [...imgs, ...prev]);
    } catch (e) {
      setError((e as Error).message);
    } finally {
      setGenerating(false);
    }
  };

  // „Bearbeiten": Kandidat anhand der eingegebenen Bildbeschreibung neu erzeugen.
  const handleRegenerate = async (img: Candidate, editFreitext: string) => {
    setError(null);
    setEditingId(null);
    try {
      const fresh = await regenImage({ theme, wochenspruch, freitext: editFreitext });
      if (img.objectUrl) URL.revokeObjectURL(img.objectUrl);
      delImage(img.id).catch(() => {});
      setImages((prev) => prev.map((x) => (x.id === img.id ? (fresh as Candidate) : x)));
    } catch (e) {
      setError((e as Error).message);
    }
  };

  const startEditing = (img: Candidate) => {
    setEditingId(img.id);
    setEditText(freitext);
  };

  const handleDelete = (img: Candidate) => {
    if (img.objectUrl) URL.revokeObjectURL(img.objectUrl);
    setImages((prev) => prev.filter((x) => x.id !== img.id));
    delImage(img.id).catch(() => {});
  };

  const handleConfirm = async (img: Candidate) => {
    if (img.status !== "done" || !img.objectUrl) return;
    if (!dateStr) {
      setError("Kein Datum – bitte zuerst einen Gottesdienst wählen.");
      return;
    }
    setError(null);
    setConfirmingId(img.id);
    try {
      const jpeg = await objectUrlToJpegBlob(img.objectUrl);
      const res = await saveSundayImage(jpeg, dateStr);
      setSavedName(res.name);
    } catch (e) {
      setError((e as Error).message);
    } finally {
      setConfirmingId(null);
    }
  };

  const filteredSheets = showAllQuarters
    ? sheets
    : quarterPattern
    ? sheets.filter((s) => s.excel_path.includes(quarterPattern))
    : sheets;

  return (
    <div className="max-w-2xl mx-auto px-4 py-6 sm:py-10 min-w-0">
      {/* Generator-Erreichbarkeit: nur Prüfen-Button + Meldung */}
      <div className="flex items-center gap-3 mb-4">
        <button
          type="button"
          onClick={checkHealth}
          disabled={checking}
          className="text-sm font-medium px-4 py-2 rounded-lg bg-gray-100 hover:bg-gray-200 disabled:opacity-50"
        >
          {checking ? "Prüfe…" : "Generator prüfen"}
        </button>
        {online !== null && !checking && (
          <span className="flex items-center gap-1.5 text-sm">
            <span className={`status-dot ${online ? "found" : "missing"}`} />
            {online ? "Generator erreichbar" : "Generator nicht erreichbar (läuft er auf dem Mac Mini?)"}
          </span>
        )}
      </div>

      {/* Gottesdienst-Auswahl */}
      <header className="mb-4">
        <select
          className="input-field w-full text-base"
          value={selected ? `${selected.name}|${selected.excel_path}` : ""}
          onChange={(e) => {
            const [name, ...rest] = e.target.value.split("|");
            const path = rest.join("|");
            const s = sheets.find((sh) => sh.name === name && sh.excel_path === path);
            if (s) loadSheet(s);
          }}
        >
          <option value="" disabled>
            Gottesdienst wählen…
          </option>
          {Object.entries(
            filteredSheets.reduce<Record<string, Sheet[]>>((groups, s) => {
              const file = s.excel_path.split("/").pop()?.replace(".xlsx", "") || s.excel_path;
              (groups[file] ??= []).push(s);
              return groups;
            }, {})
          ).map(([file, group]) => (
            <optgroup key={file} label={file}>
              {group.map((s) => (
                <option key={`${s.name}|${s.excel_path}`} value={`${s.name}|${s.excel_path}`}>
                  {s.name}
                </option>
              ))}
            </optgroup>
          ))}
        </select>
        <div className="flex items-center justify-between mt-1.5">
          <button
            type="button"
            onClick={() => setShowAllQuarters((v) => !v)}
            className="text-xs text-[var(--text-secondary)] hover:text-[var(--accent)] transition-colors"
          >
            {showAllQuarters ? "Nur aktuelles Quartal" : "Alle Quartale anzeigen"}
          </button>
        </div>
      </header>

      {loadingSheet && (
        <div className="glass-card text-center py-8">
          <div className="inline-block w-5 h-5 border-2 border-[var(--accent)] border-t-transparent rounded-full animate-spin" />
        </div>
      )}

      {selected && !loadingSheet && (
        <motion.div initial={{ opacity: 0, y: 8 }} animate={{ opacity: 1, y: 0 }} className="space-y-4">
          {/* Prompt-Formular */}
          <div className="glass-card space-y-3">
            <h3 className="text-xs font-semibold uppercase tracking-wider text-[var(--text-secondary)]">
              Bildbeschreibung
            </h3>
            <div>
              <label className="text-sm font-medium text-[var(--text-secondary)]">Thema</label>
              <input className="input-field w-full mt-1" value={theme} onChange={(e) => setTheme(e.target.value)} placeholder="Thema des Gottesdienstes" />
            </div>
            <div>
              <label className="text-sm font-medium text-[var(--text-secondary)]">Wochenspruch</label>
              <textarea className="input-field w-full mt-1 min-h-[56px] resize-y" value={wochenspruch} onChange={(e) => setWochenspruch(e.target.value)} placeholder="Falls in der Excel vorhanden, automatisch übernommen" />
            </div>
            <div>
              <label className="text-sm font-medium text-[var(--text-secondary)]">Freitext (Bildidee)</label>
              <textarea className="input-field w-full mt-1 min-h-[56px] resize-y" value={freitext} onChange={(e) => setFreitext(e.target.value)} placeholder="z.B. ruhiger See bei Sonnenaufgang, sanftes Licht" />
            </div>

            <div className="flex items-center justify-between pt-1">
              <div className="flex items-center gap-2">
                <span className="text-sm text-[var(--text-secondary)]">Anzahl:</span>
                {[1, 2, 3].map((n) => (
                  <button
                    key={n}
                    type="button"
                    onClick={() => setCount(n)}
                    className={`w-8 h-8 rounded-lg text-sm font-medium transition-all ${
                      count === n ? "bg-[var(--accent)] text-white" : "bg-gray-100 hover:bg-gray-200"
                    }`}
                  >
                    {n}
                  </button>
                ))}
              </div>
            </div>

            <button className="btn-primary" disabled={generating} onClick={handleGenerate}>
              {generating ? (
                <span className="flex items-center justify-center gap-2">
                  <span className="inline-block w-4 h-4 border-2 border-white border-t-transparent rounded-full animate-spin" />
                  Generiere {count} Bild{count > 1 ? "er" : ""}…
                </span>
              ) : (
                `${count} Bild${count > 1 ? "er" : ""} generieren`
              )}
            </button>
            <p className="text-xs text-[var(--text-secondary)]">
              Generierung läuft lokal auf dem Mac Mini – ca. 1 Minute pro Bild.
            </p>
          </div>

          {/* Kandidaten */}
          {images.length > 0 && (
            <div className="glass-card">
              <h3 className="text-xs font-semibold uppercase tracking-wider text-[var(--text-secondary)] mb-3">
                Kandidaten — eines bestätigen
              </h3>
              <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
                {images.map((img) => {
                  const isConfirming = confirmingId === img.id;
                  const pending = img.status === "pending" || (img.status === "done" && !img.objectUrl);
                  const failed = img.status === "error";
                  const ready = img.status === "done" && !!img.objectUrl;
                  return (
                    <div key={img.id} className="space-y-2">
                      <div className="relative rounded-xl overflow-hidden border border-[var(--card-border)] bg-gray-50" style={{ aspectRatio: "4 / 3" }}>
                        {ready && (
                          // eslint-disable-next-line @next/next/no-img-element
                          <img src={img.objectUrl} alt="" className="absolute inset-0 w-full h-full object-cover" />
                        )}
                        {pending && !failed && (
                          <div className="absolute inset-0 flex flex-col items-center justify-center gap-2 text-[var(--text-secondary)]">
                            <span className="inline-block w-6 h-6 border-2 border-[var(--accent)] border-t-transparent rounded-full animate-spin" />
                            <span className="text-xs">wird erzeugt… (~1 Min)</span>
                          </div>
                        )}
                        {failed && (
                          <div className="absolute inset-0 flex items-center justify-center p-3 text-center">
                            <span className="text-xs text-[var(--danger)]">{img.error || "Fehler bei der Generierung"}</span>
                          </div>
                        )}
                        {isConfirming && (
                          <div className="absolute inset-0 flex items-center justify-center bg-black/30">
                            <span className="inline-block w-6 h-6 border-2 border-white border-t-transparent rounded-full animate-spin" />
                          </div>
                        )}
                      </div>
                      <div className="flex items-center gap-1.5">
                        <button
                          type="button"
                          disabled={!ready || isConfirming}
                          onClick={() => handleConfirm(img)}
                          className="flex-1 text-xs font-medium py-1.5 rounded-lg bg-[var(--success)] text-white hover:opacity-90 disabled:opacity-40"
                        >
                          Bestätigen
                        </button>
                        <button
                          type="button"
                          disabled={img.status === "pending" || isConfirming}
                          onClick={() => (editingId === img.id ? setEditingId(null) : startEditing(img))}
                          className={`text-xs font-medium px-2.5 py-1.5 rounded-lg disabled:opacity-40 ${
                            editingId === img.id ? "bg-[var(--accent)] text-white" : "bg-gray-100 hover:bg-gray-200"
                          }`}
                          title="Bildbeschreibung anpassen und neu erzeugen"
                        >
                          Bearbeiten
                        </button>
                        <button
                          type="button"
                          disabled={isConfirming}
                          onClick={() => handleDelete(img)}
                          className="text-xs font-medium px-2.5 py-1.5 rounded-lg text-[var(--danger)] hover:bg-[var(--danger)]/10 disabled:opacity-40"
                        >
                          Löschen
                        </button>
                      </div>

                      {/* Bearbeiten-Feld: Bildbeschreibung anpassen → neu erzeugen */}
                      {editingId === img.id && (
                        <div className="space-y-2 rounded-lg bg-[var(--accent)]/5 border border-[var(--accent)]/20 p-2">
                          <textarea
                            className="input-field w-full text-sm min-h-[56px] resize-y"
                            value={editText}
                            onChange={(e) => setEditText(e.target.value)}
                            placeholder="Beschreibe, wie das neue Bild aussehen soll, z.B. „ruhiger See bei Sonnenuntergang, warmes Licht“"
                            autoFocus
                          />
                          <div className="flex items-center gap-1.5">
                            <button
                              type="button"
                              onClick={() => handleRegenerate(img, editText)}
                              className="flex-1 text-xs font-medium py-1.5 rounded-lg bg-[var(--accent)] text-white hover:opacity-90"
                            >
                              Neu erzeugen
                            </button>
                            <button
                              type="button"
                              onClick={() => setEditingId(null)}
                              className="text-xs font-medium px-2.5 py-1.5 rounded-lg bg-gray-100 hover:bg-gray-200"
                            >
                              Abbrechen
                            </button>
                          </div>
                        </div>
                      )}
                    </div>
                  );
                })}
              </div>
            </div>
          )}

          {/* Erfolg */}
          <AnimatePresence>
            {savedName && (
              <motion.div
                initial={{ opacity: 0, y: 8 }}
                animate={{ opacity: 1, y: 0 }}
                exit={{ opacity: 0 }}
                className="glass-card border-[var(--success)]/30 bg-[var(--success)]/5"
              >
                <div className="flex items-center gap-2">
                  <span className="status-dot found" />
                  <span className="text-sm font-medium">
                    Gespeichert als „{savedName}". Die Folien-Vorschau im Reiter „Präsentation" erkennt es jetzt.
                  </span>
                </div>
              </motion.div>
            )}
          </AnimatePresence>

          {error && (
            <div className="glass-card border-[var(--danger)]/30 bg-[var(--danger)]/5">
              <p className="text-sm text-[var(--danger)]">{error}</p>
            </div>
          )}
        </motion.div>
      )}

      {!selected && !loadingSheet && (
        <div className="glass-card text-center py-16">
          <p className="text-lg text-[var(--text-secondary)]">Wähle einen Gottesdienst, um Bilder zu erzeugen</p>
        </div>
      )}
    </div>
  );
}
