import { NextRequest, NextResponse } from "next/server";

// Backend-URL des Cloud-Dienstes (Render). Per Vercel-Env NEXT_PUBLIC_API_URL gesetzt.
const BACKEND =
  process.env.NEXT_PUBLIC_API_URL ||
  "https://autopraesi-api.onrender.com";

// Erstellen (PPTX bauen, Bibeltexte laden) kann länger dauern; Render-Free
// braucht zudem bis zu ~60 s zum Aufwachen.
export const maxDuration = 300;

function backendError(e: unknown): NextResponse {
  const msg = e instanceof Error ? e.message : String(e);
  return NextResponse.json(
    { detail: `Backend nicht erreichbar (${msg}). Ggf. schläft der Server – bitte in 1 Minute erneut versuchen.` },
    { status: 502 }
  );
}

export async function GET(
  req: NextRequest,
  { params }: { params: Promise<{ path: string[] }> }
) {
  const { path } = await params;
  const url = new URL(req.url);
  const target = `${BACKEND}/api/${path.join("/")}${url.search}`;
  try {
    return await proxy(await fetch(target));
  } catch (e) {
    return backendError(e);
  }
}

export async function POST(
  req: NextRequest,
  { params }: { params: Promise<{ path: string[] }> }
) {
  const { path } = await params;
  const url = new URL(req.url);
  const target = `${BACKEND}/api/${path.join("/")}${url.search}`;
  const contentType = req.headers.get("content-type") || "";

  try {
    const res = await fetch(target, {
      method: "POST",
      headers: {
        "Content-Type": contentType,
      },
      body: contentType.includes("json") ? await req.text() : await req.arrayBuffer(),
    });
    return await proxy(res);
  } catch (e) {
    return backendError(e);
  }
}

async function proxy(res: Response): Promise<NextResponse> {
  const contentType = res.headers.get("content-type") || "";

  if (contentType.includes("application/json") || contentType.includes("text/")) {
    const body = await res.text();
    return new NextResponse(body, {
      status: res.status,
      headers: { "Content-Type": contentType },
    });
  }

  // Binary (e.g. file downloads, images)
  const body = await res.arrayBuffer();
  return new NextResponse(body, {
    status: res.status,
    headers: {
      "Content-Type": contentType,
      ...(res.headers.get("content-disposition")
        ? { "Content-Disposition": res.headers.get("content-disposition")! }
        : {}),
    },
  });
}
