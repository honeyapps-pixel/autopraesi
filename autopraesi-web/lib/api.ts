const API = process.env.NEXT_PUBLIC_API_URL || "http://localhost:8000";

export interface Sheet {
  name: string;
  excel_path: string;
}

export interface SongStatus {
  slot_key: string;
  raw: string;
  category: string;
  book: string;
  number: string;
  title: string;
  found: boolean;
  file_name: string;
}

export interface InvitationEvent {
  date_str: string;
  time_str: string;
  event_name: string;
  note: string;
}

export interface SectionInfo {
  key: string;
  label: string;
  default_enabled: boolean;
}

export interface SheetData {
  service_header: string;
  theme: string;
  date_str: string;
  kirchenkalender: string;
  greeting_verse: string;
  lesung_reference: string;
  predigt1_reference: string;
  predigt1_title: string;
  predigt2_reference: string;
  predigt2_title: string;
  is_abendmahl: boolean;
  songs: SongStatus[];
  announcements: string[];
  invitation_events: InvitationEvent[];
  image_found: boolean;
  image_path: string | null;
}

export interface GenerateResult {
  success: boolean;
  output_path: string;
  output_name: string;
  missing_songs: string[];
}

export async function getSheets(): Promise<Sheet[]> {
  const res = await fetch(`${API}/api/sheets`);
  if (!res.ok) throw new Error("Sheets konnten nicht geladen werden");
  return res.json();
}

export async function getSections(): Promise<SectionInfo[]> {
  const res = await fetch(`${API}/api/sections`);
  if (!res.ok) throw new Error("Sections konnten nicht geladen werden");
  return res.json();
}

export async function getSheetData(name: string, excelPath: string): Promise<SheetData> {
  const params = new URLSearchParams({ excel_path: excelPath });
  const res = await fetch(`${API}/api/sheet/${encodeURIComponent(name)}?${params}`);
  if (!res.ok) throw new Error(`Sheet '${name}' konnte nicht geladen werden`);
  return res.json();
}

export async function uploadImage(file: File): Promise<string> {
  const form = new FormData();
  form.append("file", file);
  const res = await fetch(`${API}/api/upload-image`, { method: "POST", body: form });
  if (!res.ok) throw new Error("Bild-Upload fehlgeschlagen");
  const data = await res.json();
  return data.path;
}

export async function generate(req: {
  sheet_name: string;
  excel_path: string;
  overrides?: Record<string, unknown>;
  fetch_bible?: boolean;
  disabled_sections?: string[];
}): Promise<GenerateResult> {
  const res = await fetch(`${API}/api/generate`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(req),
  });
  if (!res.ok) {
    const err = await res.json().catch(() => ({ detail: "Unbekannter Fehler" }));
    throw new Error(err.detail || "Generierung fehlgeschlagen");
  }
  return res.json();
}

export function downloadUrl(filename: string): string {
  return `${API}/api/download/${encodeURIComponent(filename)}`;
}
