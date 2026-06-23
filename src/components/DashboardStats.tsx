/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React from 'react';
import { ProjectAnalysis } from '../types';
import { FileText, CheckCircle, AlertTriangle, Layers, Type, Percent, Sliders } from 'lucide-react';

interface DashboardStatsProps {
  analysis: ProjectAnalysis;
}

export default function DashboardStats({ analysis }: DashboardStatsProps) {
  const { totalFiles, totalParagraphs, passedParagraphs, failedParagraphs, totalIssues } = analysis.summary;
  
  // Calculate aggregate pass percentage
  const passRate = totalParagraphs > 0 ? Math.round((passedParagraphs / totalParagraphs) * 100) : 100;

  // Compute issue category logs
  let fontNameIssues = 0;
  let fontSizeIssues = 0;
  let boldIssues = 0;
  let alignIssues = 0;
  let indentIssues = 0;

  analysis.files.forEach(file => {
    file.paragraphs.forEach(para => {
      para.issues.forEach(issue => {
        if (issue.type === 'fontName') fontNameIssues++;
        else if (issue.type === 'fontSize') fontSizeIssues++;
        else if (issue.type === 'bold') boldIssues++;
        else if (issue.type === 'alignment') alignIssues++;
        else if (issue.type === 'indent') indentIssues++;
      });
    });
  });

  return (
    <div id="dashboard_stats_panel" class="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-4 gap-4 mb-6 animate-fade-in font-sans">
      
      {/* KPI Card 1: Documents */}
      <div class="bg-white p-4 rounded-xl border border-slate-200 shadow-sm">
        <p class="text-xs font-semibold text-slate-500 uppercase tracking-wider mb-1">Files Scanned</p>
        <div class="flex items-end justify-between">
          <div class="flex items-end space-x-2">
            <span class="text-2xl font-bold text-slate-800 font-display">{totalFiles}</span>
            <span class="text-xs text-slate-400 mb-1">.docx files</span>
          </div>
          <div class="p-1 px-1.5 bg-indigo-50 text-indigo-600 rounded-md">
            <FileText class="w-4 h-4" />
          </div>
        </div>
        <p class="text-[10px] text-slate-400 font-mono mt-1">Size: {(analysis.files.reduce((sum, f) => sum + f.fileSize, 0) / 1024).toFixed(1)} KB aggregate</p>
      </div>

      {/* KPI Card 2: Total Paragraphs Checked */}
      <div class="bg-white p-4 rounded-xl border border-slate-200 shadow-sm">
        <p class="text-xs font-semibold text-slate-500 uppercase tracking-wider mb-1">Total Paragraphs</p>
        <div class="flex items-end justify-between">
          <div class="flex items-end space-x-2">
            <span class="text-2xl font-bold text-slate-800 font-display">{totalParagraphs}</span>
            <span class="text-xs text-indigo-500 mb-1 font-medium">+{passedParagraphs} unique</span>
          </div>
          <div class="p-1 px-1.5 bg-sky-50 text-sky-600 rounded-md">
            <Layers class="w-4 h-4" />
          </div>
        </div>
        <div class="flex items-center gap-1.5 mt-1 text-[10px] font-medium text-slate-400">
          <span class="text-emerald-600 font-bold">{passedParagraphs} passed</span>
          <span>/</span>
          <span class="text-rose-500 font-bold">{failedParagraphs} failed</span>
        </div>
      </div>

      {/* KPI Card 3: Violations */}
      <div class="bg-white p-4 rounded-xl border border-slate-200 shadow-sm">
        <p class="text-xs font-semibold text-slate-500 uppercase tracking-wider mb-1">Violations</p>
        <div class="flex items-end justify-between">
          <div class="flex items-end space-x-2">
            <span class={`text-2xl font-bold font-display ${totalIssues > 0 ? 'text-rose-600' : 'text-emerald-600'}`}>
              {totalIssues}
            </span>
            <span class={`text-xs mb-1 font-medium ${totalIssues > 0 ? 'text-rose-400' : 'text-emerald-400'}`}>
              {totalIssues > 0 ? 'Requires Fix' : 'Verified'}
            </span>
          </div>
          <div class={`p-1 px-1.5 rounded-md ${totalIssues > 0 ? 'bg-rose-50 text-rose-500' : 'bg-emerald-50 text-emerald-500'}`}>
            <AlertTriangle class="w-4 h-4" />
          </div>
        </div>
        <p class="text-[10px] text-slate-400 font-medium mt-1">
          {totalIssues > 0 ? 'Formatting compliance gaps' : 'Standards-compliant content'}
        </p>
      </div>

      {/* KPI Card 4: Compliance Rate */}
      <div class="bg-white p-4 rounded-xl border border-slate-200 shadow-sm">
        <p class="text-xs font-semibold text-slate-500 uppercase tracking-wider mb-1">Compliance Score</p>
        <div class="flex items-end justify-between">
          <div class="flex items-end space-x-2 w-full justify-between">
            <span class={`text-2xl font-bold font-display ${passRate >= 90 ? 'text-emerald-600' : passRate >= 70 ? 'text-amber-600' : 'text-rose-600'}`}>{passRate}%</span>
            <div class="w-12 h-1.5 bg-slate-100 rounded-full mb-2 overflow-hidden">
              <div 
                class={`h-full transition-all duration-500 ${passRate >= 90 ? 'bg-emerald-500' : passRate >= 70 ? 'bg-amber-400' : 'bg-rose-500'}`}
                style={{ width: `${passRate}%` }}
              ></div>
            </div>
          </div>
          <div class={`p-1 px-1.5 rounded-md hidden sm:block ${passRate >= 90 ? 'bg-emerald-50 text-emerald-600' : 'bg-amber-50 text-amber-600'}`}>
            <Percent class="w-4 h-4" />
          </div>
        </div>
        <p class="text-[10px] text-slate-400 font-medium mt-1">Weighted validation threshold</p>
      </div>

      {totalIssues > 0 && (
        <div class="col-span-1 md:col-span-2 lg:col-span-4 bg-slate-50 border border-slate-200 rounded-xl p-4 flex flex-wrap items-center gap-6 justify-between animate-fade-in text-xs text-slate-600 font-medium">
          <div class="flex items-center gap-2">
            <Sliders class="w-4 h-4 text-indigo-500" />
            <span class="font-bold text-slate-700">Detailed Deviation Categories:</span>
          </div>
          <div class="flex flex-wrap items-center gap-4 text-[10px] font-mono">
            <span class="bg-white border border-slate-200 rounded-md px-2.5 py-1 flex items-center gap-1.5 shadow-sm text-slate-600">
              <Type class="w-3.5 h-3.5 text-indigo-500" /> Font: <strong class="text-indigo-600 font-semibold">{fontNameIssues}</strong>
            </span>
            <span class="bg-white border border-slate-200 rounded-md px-2.5 py-1 shadow-sm text-slate-600">
              📏 Size: <strong class="text-indigo-600 font-semibold">{fontSizeIssues}</strong>
            </span>
            <span class="bg-white border border-slate-200 rounded-md px-2.5 py-1 shadow-sm text-slate-600">
              <b>B</b> Bold: <strong class="text-indigo-600 font-semibold">{boldIssues}</strong>
            </span>
            <span class="bg-white border border-slate-200 rounded-md px-2.5 py-1 shadow-sm text-slate-600">
              ⇔ Align: <strong class="text-indigo-600 font-semibold">{alignIssues}</strong>
            </span>
            <span class="bg-white border border-slate-200 rounded-md px-2.5 py-1 shadow-sm text-slate-600">
              ⇥ Indent: <strong class="text-indigo-600 font-semibold">{indentIssues}</strong>
            </span>
          </div>
        </div>
      )}
    </div>
  );
}
