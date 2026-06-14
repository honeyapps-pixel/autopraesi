"use client";

import { useState } from "react";
import { motion } from "framer-motion";
import PresentationTab from "./components/PresentationTab";
import GodiPlanTab from "./components/GodiPlanTab";
import BilderTab from "./components/BilderTab";

type Tab = "praesi" | "godi" | "bilder";

const TABS: { id: Tab; label: string }[] = [
  { id: "praesi", label: "Präsentation" },
  { id: "godi", label: "GoDi-Plan" },
  { id: "bilder", label: "Bilder generieren" },
];

export default function Home() {
  const [tab, setTab] = useState<Tab>("praesi");
  // GoDi- und Bilder-Reiter erst beim ersten Öffnen laden, danach montiert lassen.
  const [godiSeen, setGodiSeen] = useState(false);
  const [bilderSeen, setBilderSeen] = useState(false);

  const openTab = (id: Tab) => {
    if (id === "godi") setGodiSeen(true);
    if (id === "bilder") setBilderSeen(true);
    setTab(id);
  };

  return (
    <div className="min-h-dvh">
      {/* Globale Kopfzeile + Reiter */}
      <header className="sticky top-0 z-40 border-b border-[var(--card-border)] bg-[var(--bg)]/80 backdrop-blur-xl">
        <div className="max-w-6xl mx-auto px-4 flex items-center gap-5 h-14">
          <div className="flex items-center gap-2.5 shrink-0">
            <img src="/logo.jpg" alt="Gemeindelogo" className="h-8 w-8 rounded-full object-cover shadow-sm" />
            <div className="leading-none">
              <span className="text-[15px] font-bold tracking-tight">AutoPräsi</span>
              <span className="hidden sm:block text-[10px] text-[var(--text-secondary)] mt-0.5 tracking-wide uppercase">
                Ev.-luth. Christus-Brüdergemeinde
              </span>
            </div>
          </div>

          <nav className="flex items-end gap-1 h-full ml-1" role="tablist" aria-label="Bereiche">
            {TABS.map((t) => {
              const active = tab === t.id;
              return (
                <button
                  key={t.id}
                  role="tab"
                  aria-selected={active}
                  onClick={() => openTab(t.id)}
                  className={`relative px-3.5 h-full text-sm font-medium transition-colors duration-150 cursor-pointer ${
                    active ? "text-[var(--text-primary)]" : "text-[var(--text-secondary)] hover:text-[var(--text-primary)]"
                  }`}
                >
                  {t.label}
                  {active && (
                    <motion.span
                      layoutId="tab-underline"
                      className="absolute left-2 right-2 -bottom-px h-0.5 rounded-full bg-[var(--accent)]"
                      transition={{ type: "spring", stiffness: 500, damping: 38 }}
                    />
                  )}
                </button>
              );
            })}
          </nav>
        </div>
      </header>

      {/* Präsentations-Reiter (immer montiert, Zustand bleibt) */}
      <div className={tab === "praesi" ? "" : "hidden"}>
        <PresentationTab />
      </div>

      {/* GoDi-Plan-Reiter (lazy beim ersten Öffnen) */}
      {godiSeen && (
        <div className={tab === "godi" ? "" : "hidden"}>
          <GodiPlanTab />
        </div>
      )}

      {/* Bilder-Reiter (lazy beim ersten Öffnen) */}
      {bilderSeen && (
        <div className={tab === "bilder" ? "" : "hidden"}>
          <BilderTab />
        </div>
      )}
    </div>
  );
}
