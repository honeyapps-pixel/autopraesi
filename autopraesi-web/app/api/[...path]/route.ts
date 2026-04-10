import { NextRequest, NextResponse } from "next/server";

const BACKEND =
  process.env.NEXT_PUBLIC_API_URL ||
  "https://tawniest-uxorially-boyce.ngrok-free.dev";

export async function GET(
  req: NextRequest,
  { params }: { params: Promise<{ path: string[] }> }
) {
  const { path } = await params;
  const url = new URL(req.url);
  const target = `${BACKEND}/api/${path.join("/")}${url.search}`;
  const res = await fetch(target, {
    headers: { "ngrok-skip-browser-warning": "1" },
  });
  return proxy(res);
}

export async function POST(
  req: NextRequest,
  { params }: { params: Promise<{ path: string[] }> }
) {
  const { path } = await params;
  const url = new URL(req.url);
  const target = `${BACKEND}/api/${path.join("/")}${url.search}`;
  const contentType = req.headers.get("content-type") || "";

  const res = await fetch(target, {
    method: "POST",
    headers: {
      "ngrok-skip-browser-warning": "1",
      "Content-Type": contentType,
    },
    body: contentType.includes("json") ? await req.text() : await req.arrayBuffer(),
  });
  return proxy(res);
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
