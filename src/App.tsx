import React, { useState } from "react";
import { ArrowRight, BarChart3, BookOpen, LayoutDashboard } from "lucide-react";
import SupplierRiskOpsDashboard from "./SupplierRiskOpsDashboard";

export default function App() {
  const [view, setView] = useState<"landing" | "dashboard">("landing");

  if (view === "dashboard") {
    return <SupplierRiskOpsDashboard />;
  }

  return (
    <div className="min-h-screen bg-[#f4f6f8] text-slate-900">
      <div className="relative mx-auto flex min-h-screen max-w-6xl flex-col px-6 py-10">
        <header className="flex items-center gap-3 text-xs font-semibold uppercase tracking-[0.2em] text-slate-400">
          <span className="h-2 w-2 rounded-full bg-[#14b8a6]" />
          AirLiquide Platform
        </header>

        <main className="flex flex-1 flex-col justify-center">
          <div className="grid gap-10 lg:grid-cols-[1.2fr_1fr] lg:items-center">
            <div>
              <h1 className="text-4xl font-extrabold leading-tight text-slate-900 md:text-6xl">
                Less Manual Work.
                <span className="block text-[#0f766e]">More Profit.</span>
              </h1>
              <p className="mt-4 max-w-xl text-lg text-slate-500">
                Centralized supplier risk operations with clean UX, fast workflows, and full auditability.
              </p>
              <button
                className="mt-8 inline-flex items-center gap-2 rounded-full bg-[#0f766e] px-6 py-3 text-sm font-semibold text-white shadow hover:bg-[#0d6e66]"
                onClick={() => setView("dashboard")}
              >
                Open Dashboard <ArrowRight className="h-4 w-4" />
              </button>
            </div>
            <div className="hidden justify-center lg:flex">
              <div className="h-48 w-48 rounded-full bg-[#0f766e]/15" />
            </div>
          </div>

          <div className="mt-12 grid gap-6 md:grid-cols-3">
            <button
              type="button"
              onClick={() => setView("dashboard")}
              className="group rounded-3xl border border-slate-100 bg-white p-6 text-left shadow-sm transition hover:-translate-y-1 hover:shadow-lg"
            >
              <div className="mb-4 flex h-12 w-12 items-center justify-center rounded-2xl bg-[#0f766e] text-white">
                <LayoutDashboard className="h-6 w-6" />
              </div>
              <div className="text-lg font-semibold">Ops Dashboard</div>
              <p className="mt-2 text-sm text-slate-500">Central command center for supplier risk and compliance.</p>
              <div className="mt-6 text-xs font-semibold uppercase tracking-wide text-slate-400 group-hover:text-slate-700">
                Open Module →
              </div>
            </button>

            <div className="rounded-3xl border border-slate-100 bg-white p-6 text-left opacity-70">
              <div className="mb-4 flex h-12 w-12 items-center justify-center rounded-2xl bg-amber-500 text-white">
                <BarChart3 className="h-6 w-6" />
              </div>
              <div className="text-lg font-semibold">Analytics</div>
              <p className="mt-2 text-sm text-slate-500">Deep dives across categories and countries.</p>
              <div className="mt-6 text-xs font-semibold uppercase tracking-wide text-slate-400">Coming soon</div>
            </div>

            <div className="rounded-3xl border border-slate-100 bg-white p-6 text-left opacity-70">
              <div className="mb-4 flex h-12 w-12 items-center justify-center rounded-2xl bg-slate-900 text-white">
                <BookOpen className="h-6 w-6" />
              </div>
              <div className="text-lg font-semibold">How to Use</div>
              <p className="mt-2 text-sm text-slate-500">Playbooks and onboarding material.</p>
              <div className="mt-6 text-xs font-semibold uppercase tracking-wide text-slate-400">Coming soon</div>
            </div>
          </div>
        </main>

        <footer className="mt-12 text-center text-xs text-slate-400">
          Created by Kristina Kalde. Building better automation paths.
        </footer>
      </div>
    </div>
  );
}
