import React from 'react';
import { ArrowRight, Building2, GraduationCap } from 'lucide-react';

interface LandingScreenProps {
  onEnter: () => void;
}

export const LandingScreen: React.FC<LandingScreenProps> = ({ onEnter }) => {
  return (
    <div className="min-h-screen bg-slate-50 text-slate-900 flex flex-col justify-between">
      {/* Top Institutional Bar */}
      <header className="border-b border-slate-200 bg-white">
        <div className="max-w-6xl mx-auto px-6 py-4 flex items-center justify-between gap-8">
          <div className="flex items-center gap-3 whitespace-nowrap shrink-0">
            <img
              src="/assets/sib_logo.jpg"
              alt="Sammilito Islami Bank PLC"
              referrerPolicy="no-referrer"
              className="h-9 w-auto object-contain"
              onError={(e) => {
                (e.currentTarget as HTMLImageElement).style.display = 'none';
              }}
            />
            <span className="font-display text-lg font-bold tracking-tight text-slate-900">
              Memo Engine by Tuhin
            </span>
          </div>

          <div className="hidden md:flex items-center gap-4 text-sm text-slate-600">
            <span>Special Asset Management Division</span>
            <span aria-hidden="true">·</span>
            <span>Sammilito Islami Bank PLC</span>
          </div>

          <button
            type="button"
            onClick={onEnter}
            className="inline-flex items-center gap-2 px-5 py-2.5 text-sm font-semibold text-white bg-emerald-800 hover:bg-emerald-900 rounded-lg transition-colors whitespace-nowrap shrink-0 cursor-pointer focus-visible:outline-2 focus-visible:outline-offset-2 focus-visible:outline-emerald-800"
          >
            <span>Enter</span>
            <ArrowRight className="w-4 h-4" />
          </button>
        </div>
      </header>

      {/* Main Hero Section */}
      <main className="flex-1 flex items-center py-12 px-6">
        <div className="max-w-6xl mx-auto w-full grid grid-cols-1 lg:grid-cols-12 gap-10 items-stretch">
          {/* Left Column: Application Brief & Primary Action */}
          <div className="lg:col-span-7 flex flex-col justify-center space-y-8">
            <div className="space-y-4">
              <p className="text-sm font-medium text-emerald-800 tracking-wide">
                Settlement &amp; Waiver Memorandum Automation
              </p>
              <h1 className="font-display text-4xl sm:text-5xl font-bold tracking-tight text-slate-900 text-balance leading-[1.12]">
                Memo Engine by Tuhin
              </h1>
              <p className="text-base sm:text-lg text-slate-600 max-w-2xl leading-relaxed">
                Browser-native banking document processor for Exit with Waiver proposals. Extracts
                figures from seven core banking statements, applies unit and source-precedence rules,
                validates every placeholder, and builds ready-to-submit Word (.docx) memoranda.
              </p>
            </div>

            <div className="grid grid-cols-1 sm:grid-cols-3 gap-4 border-t border-b border-slate-200 py-6">
              <div>
                <p className="text-xs font-medium text-slate-500">Input Documents</p>
                <p className="mt-1 font-mono-tabular text-xl font-semibold text-slate-900">7 Sources</p>
                <p className="mt-0.5 text-xs text-slate-500">PDF, Scanned Proposal &amp; Excel</p>
              </div>
              <div>
                <p className="text-xs font-medium text-slate-500">Target Outputs</p>
                <p className="mt-1 font-mono-tabular text-xl font-semibold text-slate-900">4 Memos</p>
                <p className="mt-0.5 text-xs text-slate-500">IC, MEC, Board &amp; Sanction</p>
              </div>
              <div>
                <p className="text-xs font-medium text-slate-500">Document Export</p>
                <p className="mt-1 font-mono-tabular text-xl font-semibold text-slate-900">Client .docx</p>
                <p className="mt-0.5 text-xs text-slate-500">1:1 Reference Layout</p>
              </div>
            </div>

            <div className="flex flex-wrap items-center gap-4">
              <button
                type="button"
                onClick={onEnter}
                className="inline-flex items-center gap-2.5 px-7 py-3.5 text-base font-semibold text-white bg-emerald-800 hover:bg-emerald-900 rounded-lg transition-colors cursor-pointer focus-visible:outline-2 focus-visible:outline-offset-2 focus-visible:outline-emerald-800"
              >
                <span>Enter Workspace</span>
                <ArrowRight className="w-5 h-5" />
              </button>
              <span className="text-sm text-slate-500">
                All document parsing and .docx generation run directly in your browser.
              </span>
            </div>
          </div>

          {/* Right Column: Developer Profile Card */}
          <div className="lg:col-span-5 flex flex-col">
            <div className="bg-white border border-slate-200 rounded-xl p-8 flex-1 flex flex-col justify-between">
              <div className="space-y-6">
                <div className="flex items-start gap-4">
                  <div className="relative w-32 h-36 rounded-xl overflow-hidden border border-slate-200 bg-white shrink-0">
                    <img
                      src="/assets/aistudio/Tuhin_Pic_6c5894de.jpeg"
                      alt="Mohammad Towhidul Akbar Tuhin"
                      referrerPolicy="no-referrer"
                      className="w-full h-full object-cover object-top"
                    />
                  </div>
                </div>

                <div className="space-y-2">
                  <p className="text-xs font-medium text-slate-500">Designed &amp; Built By</p>
                  <h2 className="font-display text-2xl font-bold text-slate-900 leading-snug">
                    Mohammad Towhidul Akbar Tuhin
                  </h2>
                </div>

                <div className="space-y-3 pt-2 border-t border-slate-200 text-sm text-slate-700">
                  <div className="flex items-start gap-3">
                    <Building2 className="w-4 h-4 text-emerald-800 mt-0.5 shrink-0" />
                    <p className="leading-relaxed">
                      Officer, <span className="font-semibold text-slate-900">Sammilito Islami Bank PLC</span>{' '}
                      (Transferee of Assets &amp; Liabilities of Union Bank PLC)
                    </p>
                  </div>
                  <div className="flex items-start gap-3">
                    <GraduationCap className="w-4 h-4 text-emerald-800 mt-0.5 shrink-0" />
                    <div className="space-y-1">
                      <p>BBA, Heriot Watt University, UK</p>
                      <p>MBA, University of Chittagong</p>
                    </div>
                  </div>
                </div>
              </div>

              <div className="mt-8 pt-4 border-t border-slate-200 flex items-center justify-between text-xs text-slate-500">
                <span>Union Bank PLC Unit</span>
                <span>Head Office, Dhaka</span>
              </div>
            </div>
          </div>
        </div>
      </main>

      <footer className="border-t border-slate-200 bg-white py-4 px-6">
        <div className="max-w-6xl mx-auto flex flex-wrap items-center justify-between gap-4 text-xs text-slate-500">
          <span>Memo Engine by Tuhin · Sammilito Islami Bank PLC</span>
          <span>IC Memo Operational · MEC, Board &amp; Sanction Ready for Rules</span>
        </div>
      </footer>
    </div>
  );
};
