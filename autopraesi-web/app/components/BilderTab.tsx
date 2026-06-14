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
  imageUrl,
  health as genHealth,
  getImagegenBase,
  setImagegenBase,
} from "@/lib/imagegen";

/** Lädt ein Kandidaten-PNG und konvertiert es im Browser nach JPEG (für den Upload). */
async function pngUrlToJpegBlob(url: string): Promise<Blob> {
  const res = await fetch(url, { cache: "no-store" });
  if (!res.ok) throw new Error("Bild konnte nicht geladen werden");
  const pngBlob = await res.blob();
  const bitmap = await createImageBitmap(pngBlob);
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
  const [images, setImages] = useState<GenImage[]>([]);
  const [generating, setGenerating] = useState(false);
  const [confirmingId, setConfirmingId] = useState<string | null>(null);
  const [savedName, setSavedName] = useState<string | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [loadingSheet, setLoadingSheet] = useState(false);

  // Generator-URL (Tunnel) + Erreichbarkeit
  const [genUrl, setGenUrl] = useState("");
  const [online, setOnline] = useState<boolean | null>(null);

  useEffect(() => {
    setGenUrl(getImagegenBase());
    getCurrentQuarter().then(setQuarterPattern).catch(() => {});
    getSheets().then(setSheets).catch((e) => setError(e.message));
  }, []);

  const checkHealth = useCallback(async () => {
    setOnline(await genHealth());
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

  // Solange Bilder "pending" sind, regelmäßig den Status pollen (Generierung ~1 Min/Bild).
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

  const saveGenUrl = () => {
    setImagegenBase(genUrl);
    setOnline(null);
    checkHealth();
  };

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

  // „Bearbeiten": Kandidat mit aktuellem (ggf. geändertem) Prompt + neuem Seed neu erzeugen.
  const handleRegenerate = async (img: GenImage) => {
    setError(null);
    try {
      const fresh = await regenImage(promptInput());
      delImage(img.id).catch(() => {});
      setImages((prev) => prev.map((x) => (x.id === img.id ? fresh : x)));
    } catch (e) {
      setError((e as Error).message);
    }
  };

  const handleDelete = (img: GenImage) => {
    setImages((prev) => prev.filter((x) => x.id !== img.id));
    delImage(img.id).catch(() => {});
  };

  const handleConfirm = async (img: GenImage) => {
    if (img.status !== "done") return;
    if (!dateStr) {
      setError("Kein Datum – bitte zuerst einen Gottesdienst wählen.");
      return;
    }
    setError(null);
    setConfirmingId(img.id);
    try {
      const jpeg = await pngUrlToJpegBlob(imageUrl(img));
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
      {/* Generator-Status / Tunnel-URL */}
      <div className="glass-card mb-4">
        <div className="flex items-center justify-between mb-2">
          <h3 className="text-xs font-semibold uppercase tracking-wider text-[var(--text-secondary)]">
            Bild-Generator (Mac Mini)
          </h3>
          <span className="flex items-center gap-1.5 text-xs">
            <span className={`status-dot ${online ? "found" : online === false ? "missing" : "empty"}`} />
            {online === null ? "prüfe…" : online ? "erreichbar" : "nicht erreichbar"}
          </span>
        </div>
        <div className="flex items-center gap-2">
          <input
            className="input-field flex-1 text-sm"
            value={genUrl}
            placeholder="https://…trycloudflare.com  (Tunnel-URL)"
            onChange={(e) => setGenUrl(e.target.value)}
          />
          <button type="button" onClick={saveGenUrl} className="text-xs font-medium px-3 py-2 rounded-lg bg-gray-100 hover:bg-gray-200">
            Speichern
          </button>
          <button type="button" onClick={checkHealth} className="text-xs font-medium px-3 py-2 rounded-lg bg-gray-100 hover:bg-gray-200">
            Prüfen
          </button>
        </div>
        {online === false && (
          <p className="text-xs text-[var(--text-secondary)] mt-2">
            Läuft <code>start-imagegen.sh</code> auf dem Mac Mini und ist die Tunnel-URL korrekt?
          </p>
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

            <button className="btn-primary" disabled={generating || online === false} onClick={handleGenerate}>
              {generating ? (
                <span className="flex items-center justify-center gap-2">
                  <span className="inline-block w-4 h-4 border-2 border-white border-t-transparent rounded-full animate-spin" />
                  Generiere {count} Bild{count > 1 ? "er" : ""}…
                </span>
              ) : (
                `${count} Bild${count > 1 ? "er" : ""} generieren`
              )}
            </button>
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
                  const pending = img.status === "pending";
                  const failed = img.status === "error";
                  const done = img.status === "done";
                  return (
                    <div key={img.id} className="space-y-2">
                      <div className="relative rounded-xl overflow-hidden border border-[var(--card-border)] bg-gray-50" style={{ aspectRatio: "4 / 3" }}>
                        {done && (
                          // eslint-disable-next-line @next/next/no-img-element
                          <img src={imageUrl(img)} alt="" className="absolute inset-0 w-full h-full object-cover" />
                        )}
                        {pending && (
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
                          disabled={!done || isConfirming}
                          onClick={() => handleConfirm(img)}
                          className="flex-1 text-xs font-medium py-1.5 rounded-lg bg-[var(--success)] text-white hover:opacity-90 disabled:opacity-40"
                        >
                          Bestätigen
                        </button>
                        <button
                          type="button"
                          disabled={pending || isConfirming}
                          onClick={() => handleRegenerate(img)}
                          className="text-xs font-medium px-2.5 py-1.5 rounded-lg bg-gray-100 hover:bg-gray-200 disabled:opacity-40"
                          title="Mit aktuellem Prompt neu erzeugen"
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
