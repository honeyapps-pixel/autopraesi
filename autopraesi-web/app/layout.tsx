import type { Metadata } from "next";
import "./globals.css";

export const metadata: Metadata = {
  title: "AutoPräsi",
  description: "Gottesdienst-Präsentation Generator",
};

export default function RootLayout({
  children,
}: {
  children: React.ReactNode;
}) {
  return (
    <html lang="de">
      <body className="min-h-screen">{children}</body>
    </html>
  );
}
