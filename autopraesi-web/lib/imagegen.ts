// Client für den LOKALEN Bild-Generator (MFLUX auf dem Mac Mini, via Cloudflared-Tunnel).
// Anders als lib/api.ts (Cloud-Backend über den Vercel-Proxy) gehen diese Aufrufe DIREKT
// an den Tunnel – der Generator läuft nicht in der Cloud.
//
// Basis-URL: NEXT_PUBLIC_IMAGEGEN_URL (z.B. stabile Tunnel-URL), zur Laufzeit per
// localStorage überschreibbar (für wechselnde Quick-Tunnel-URLs ohne Rebuild).

const LS_KEY = "autopraesi_imagegen_url";

export function getImagegenBase(): string {
  if (typeof window !== "undefined") {
    const override = window.localStorage.getItem(LS_KEY);
    if (override) return override.replace(/\/+$/, "");
  }
  return (process.env.NEXT_PUBLIC_IMAGEGEN_URL || "").replace(/\/+$/, "");
}

export function setImagegenBase(url: string): void {
  if (typeof window === "undefined") return;
  const clean = url.trim().replace(/\/+$/, "");
  if (clean) window.localStorage.setItem(LS_KEY, clean);
  else window.localStorage.removeItem(LS_KEY);
}

export type GenStatus = "pending" | "done" | "error" | "deleted" | "unknown";

export interface GenImage {
  id: string;
  seed: number | null;
  url: string; // relativer Pfad am Generator (/image/<id>)
  status: GenStatus;
  error?: string | null;
}

export interface JobStatus {
  id: string;
  status: GenStatus;
  error: string | null;
}

export interface PromptInput {
  theme: string;
  wochenspruch: string;
  freitext: string;
}

function base(): string {
  const b = getImagegenBase();
  if (!b) throw new Error("Keine Generator-URL gesetzt. Bitte oben die Tunnel-URL eintragen.");
  return b;
}

/** Vollständige URL zum Anzeigen eines Kandidatenbildes. */
export function imageUrl(img: GenImage): string {
  return `${getImagegenBase()}${img.url}`;
}

export async function health(): Promise<boolean> {
  const b = getImagegenBase();
  if (!b) return false;
  try {
    const res = await fetch(`${b}/health`, { cache: "no-store" });
    if (!res.ok) return false;
    const data = await res.json();
    return !!data.ok;
  } catch {
    return false;
  }
}

// Legt Jobs an; die Bilder kommen mit status "pending" zurück und werden per
// pollStatus() aktualisiert, bis sie "done" sind.
export async function generate(input: PromptInput, count: number, seed?: number): Promise<GenImage[]> {
  const res = await fetch(`${base()}/generate`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ ...input, count, seed }),
  });
  if (!res.ok) {
    const err = await res.json().catch(() => ({ detail: "Generierung fehlgeschlagen" }));
    throw new Error(err.detail || "Generierung fehlgeschlagen");
  }
  const data = await res.json();
  return data.images as GenImage[];
}

export async function regenerate(input: PromptInput, seed?: number): Promise<GenImage> {
  const res = await fetch(`${base()}/regenerate`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ ...input, seed }),
  });
  if (!res.ok) {
    const err = await res.json().catch(() => ({ detail: "Bearbeiten fehlgeschlagen" }));
    throw new Error(err.detail || "Bearbeiten fehlgeschlagen");
  }
  return (await res.json()) as GenImage;
}

export async function pollStatus(ids: string[]): Promise<JobStatus[]> {
  if (ids.length === 0) return [];
  const res = await fetch(`${base()}/status?ids=${encodeURIComponent(ids.join(","))}`, { cache: "no-store" });
  if (!res.ok) throw new Error("Status-Abfrage fehlgeschlagen");
  return (await res.json()) as JobStatus[];
}

export async function deleteImage(id: string): Promise<void> {
  await fetch(`${base()}/image/${encodeURIComponent(id)}`, { method: "DELETE" });
}
