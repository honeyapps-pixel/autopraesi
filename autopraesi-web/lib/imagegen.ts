// Client für den LOKALEN Bild-Generator (MFLUX auf dem Mac Mini, via ngrok-Tunnel).
// Anders als lib/api.ts (Cloud-Backend über den Vercel-Proxy) gehen diese Aufrufe DIREKT
// an den Tunnel – der Generator läuft nicht in der Cloud.
//
// Basis-URL kommt aus NEXT_PUBLIC_IMAGEGEN_URL (feste ngrok-Domain, in Vercel gesetzt).
// ngrok-Free zeigt sonst eine HTML-Warnseite statt der API-Antwort – deshalb senden wir
// bei JEDEM Aufruf den Header "ngrok-skip-browser-warning".

const NGROK_HEADER = { "ngrok-skip-browser-warning": "true" } as const;

export function getImagegenBase(): string {
  return (process.env.NEXT_PUBLIC_IMAGEGEN_URL || "").replace(/\/+$/, "");
}

export type GenStatus = "pending" | "done" | "error" | "deleted" | "unknown";

export interface GenImage {
  id: string;
  seed: number | null;
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
  if (!b) throw new Error("Bild-Generator ist nicht konfiguriert.");
  return b;
}

export async function health(): Promise<boolean> {
  const b = getImagegenBase();
  if (!b) return false;
  try {
    const res = await fetch(`${b}/health`, { cache: "no-store", headers: { ...NGROK_HEADER } });
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
    headers: { "Content-Type": "application/json", ...NGROK_HEADER },
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
    headers: { "Content-Type": "application/json", ...NGROK_HEADER },
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
  const res = await fetch(`${base()}/status?ids=${encodeURIComponent(ids.join(","))}`, {
    cache: "no-store",
    headers: { ...NGROK_HEADER },
  });
  if (!res.ok) throw new Error("Status-Abfrage fehlgeschlagen");
  return (await res.json()) as JobStatus[];
}

/** Lädt ein fertiges Bild als Blob (mit ngrok-Header) und gibt eine Object-URL zurück. */
export async function fetchImageObjectUrl(id: string): Promise<string> {
  const res = await fetch(`${base()}/image/${encodeURIComponent(id)}`, {
    cache: "no-store",
    headers: { ...NGROK_HEADER },
  });
  if (!res.ok) throw new Error("Bild konnte nicht geladen werden");
  const blob = await res.blob();
  return URL.createObjectURL(blob);
}

export async function deleteImage(id: string): Promise<void> {
  await fetch(`${base()}/image/${encodeURIComponent(id)}`, {
    method: "DELETE",
    headers: { ...NGROK_HEADER },
  });
}
