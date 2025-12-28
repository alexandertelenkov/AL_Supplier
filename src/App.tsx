import React, { useState } from "react";
import { motion } from "framer-motion";
import { BarChart3, BookOpen, ChevronRight, LayoutDashboard } from "lucide-react";
import SupplierRiskOpsDashboard from "./SupplierRiskOpsDashboard";

const LandingCard = ({ title, subtitle, icon: Icon, colorClass, onClick, index }: any) => {
  return (
    <motion.div
      initial={{ opacity: 0, y: 20 }}
      animate={{ opacity: 1, y: 0 }}
      transition={{ delay: index * 0.1 }}
      onClick={onClick}
      className="group relative cursor-pointer overflow-hidden rounded-3xl border border-slate-100 bg-white p-8 shadow-xl transition-all hover:-translate-y-1 hover:shadow-2xl"
    >
      <div className={`absolute -right-6 -top-6 h-32 w-32 rounded-full opacity-10 transition-transform group-hover:scale-150 ${colorClass}`} />

      <div className="relative z-10">
        <div className={`mb-6 flex h-14 w-14 items-center justify-center rounded-2xl ${colorClass} text-white shadow-lg`}>
          <Icon className="h-7 w-7" />
        </div>

        <h3 className="mb-3 text-2xl font-bold text-slate-900">{title}</h3>
        <p className="leading-relaxed text-slate-500">{subtitle}</p>

        <div className="mt-8 flex items-center text-sm font-bold uppercase tracking-wider text-slate-400 transition-colors group-hover:text-slate-900">
          <span>Open Module</span>
          <ChevronRight className="ml-2 h-4 w-4" />
        </div>
      </div>
    </motion.div>
  );
};

function LandingPage({ onNavigate }: { onNavigate: (view: "dashboard" | "analytics" | "docs") => void }) {
  return (
    <div className="min-h-screen bg-[#f4f6f8] font-sans text-slate-900 selection:bg-[#0f766e] selection:text-white">
      <div className="pointer-events-none fixed inset-0 z-0 overflow-hidden">
        <div className="absolute left-[-10%] top-[-20%] h-[50%] w-[50%] rounded-full bg-[#0f766e]/5 blur-[120px]" />
        <div className="absolute bottom-[-10%] right-[-5%] h-[40%] w-[40%] rounded-full bg-amber-500/5 blur-[100px]" />
      </div>

      <div className="relative z-10 flex min-h-screen flex-col">
        <header className="px-6 py-8">
          <div className="mx-auto flex max-w-7xl items-center gap-3">
            <div className="h-3 w-3 rounded-full bg-[#14b8a6] shadow-[0_0_0_4px_rgba(20,184,166,0.1)]" />
            <span className="text-xs font-bold uppercase tracking-[0.2em] text-slate-400">AirLiquide Platform</span>
          </div>
        </header>

        <main className="flex flex-1 flex-col justify-center px-6 pb-20">
          <div className="mx-auto w-full max-w-7xl">
            <div className="mb-24 grid items-center gap-16 lg:grid-cols-2">
              <div>
                <h1 className="mb-8 text-5xl font-extrabold leading-[1.1] tracking-tight text-slate-900 lg:text-7xl">
                  Less Manual Work. <br />
                  <span className="bg-gradient-to-br from-[#0f766e] to-[#14b8a6] bg-clip-text text-transparent">
                    More Profit.
                  </span>
                </h1>
                <p className="max-w-lg text-xl leading-relaxed text-slate-500">
                  We orchestrate your data intelligence and strategic operations. Automate the path from chaos to clarity.
                </p>
              </div>

              <div className="relative hidden lg:block">
                <div className="relative mx-auto aspect-square max-w-md">
                  <div className="absolute inset-0 animate-pulse rounded-full bg-gradient-to-tr from-slate-200 to-slate-50 opacity-20" />
                  <svg viewBox="0 0 200 200" className="h-full w-full text-[#0f766e] opacity-20">
                    <path
                      fill="currentColor"
                      d="M45.7,118c-13.9-28.9,3.1-66.8,36.4-74.9c29.3-7.1,58.7,11.5,70.9,40c11.6,27.1,0.5,60.9-25.8,75.9C95.5,177.3,57.1,141.7,45.7,118z"
                    />
                  </svg>
                </div>
              </div>
            </div>

            <div className="grid gap-8 md:grid-cols-3">
              <LandingCard
                index={0}
                title="Ops Dashboard"
                subtitle="Centralized command center for supplier risk, contracts, and compliance status."
                icon={LayoutDashboard}
                colorClass="bg-[#0f766e]"
                onClick={() => onNavigate("dashboard")}
              />
              <LandingCard
                index={1}
                title="Analytics"
                subtitle="Deep dive into category spend, country distribution, and risk exposure."
                icon={BarChart3}
                colorClass="bg-[#f59e0b]"
                onClick={() => onNavigate("analytics")}
              />
              <LandingCard
                index={2}
                title="How to Use"
                subtitle="Documentation, standard operating procedures, and system guides."
                icon={BookOpen}
                colorClass="bg-slate-800"
                onClick={() => onNavigate("docs")}
              />
            </div>
          </div>
        </main>

        <footer className="py-8 text-center text-sm text-slate-400">
          Created by Kristina Kalde. Building better automation paths.
        </footer>
      </div>
    </div>
  );
}

export default function App() {
  const [view, setView] = useState<"landing" | "dashboard">("landing");

  if (view === "dashboard") {
    return <SupplierRiskOpsDashboard />;
  }

  return <LandingPage onNavigate={(next) => (next === "dashboard" ? setView("dashboard") : setView("landing"))} />;
}
