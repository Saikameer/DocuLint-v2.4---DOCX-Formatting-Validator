/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React, { useState } from 'react';
import { FileAnalysis, ParagraphAnalysis } from '../types';
import { Search, Filter, AlertCircle, CheckCircle, ChevronDown, ChevronUp, Eye, FileSpreadsheet, EyeOff } from 'lucide-react';

interface FileReportProps {
  fileAnalysis: FileAnalysis;
  onDownloadCsv: () => void;
}

export default function FileReport({ fileAnalysis, onDownloadCsv }: FileReportProps) {
  const [searchTerm, setSearchTerm] = useState('');
  const [showFailedOnly, setShowFailedOnly] = useState(false);
  const [selectedLevelFilter, setSelectedLevelFilter] = useState('ALL');
  const [expandedParas, setExpandedParas] = useState<Record<number, boolean>>({});

  const toggleExpand = (pNum: number) => {
    setExpandedParas(prev => ({ ...prev, [pNum]: !prev[pNum] }));
  };

  // Filter logic
  const filteredParagraphs = fileAnalysis.paragraphs.filter(p => {
    const matchesSearch = p.text.toLowerCase().includes(searchTerm.toLowerCase());
    const matchesFailed = !showFailedOnly || p.validationStatus === 'Fail';
    const matchesLevel = selectedLevelFilter === 'ALL' || p.matchedLevelId === selectedLevelFilter;
    return matchesSearch && matchesFailed && matchesLevel;
  });

  // Unique standard levels found in this file to populate standard filters
  const levelOptionIds = Array.from(new Set(fileAnalysis.paragraphs.map(p => p.matchedLevelId)));

  return (
    <div id="file_report_panel" class="bg-white rounded-xl border border-slate-200 shadow-sm overflow-hidden animate-fade-in font-sans flex flex-col">
      
      {/* File Header Details Area */}
      <div class="p-4 border-b border-slate-200 flex flex-col sm:flex-row sm:items-center justify-between gap-4 bg-slate-50/50">
        <div>
          <h2 class="text-sm font-bold text-slate-700 font-display flex items-center gap-2">
            Analysis Results &mdash; <span class="text-indigo-600">{fileAnalysis.fileName}</span>
          </h2>
          <p class="text-[11px] text-slate-400 mt-0.5">
            Size: {(fileAnalysis.fileSize / 1024).toFixed(1)} KB | processed: {fileAnalysis.summary.totalParagraphs} lines
          </p>
        </div>
        <div class="flex items-center gap-2">
          <span class="text-[10px] font-bold bg-slate-100 text-slate-600 px-2 py-0.5 rounded font-mono">
            {fileAnalysis.summary.passed} Passed / {fileAnalysis.summary.failed} Failed
          </span>
          <button
            onClick={onDownloadCsv}
            class="px-3 py-1.5 bg-white text-slate-700 border border-slate-200 rounded-lg text-xs font-semibold hover:bg-slate-50 active:scale-95 transition flex items-center gap-1.5 shadow-sm"
          >
            <FileSpreadsheet class="w-3.5 h-3.5 text-slate-400" />
            Export CSV
          </button>
        </div>
      </div>

      {/* Advanced Filter and Control Bar */}
      <div class="p-4 border-b border-slate-100 flex flex-col md:flex-row gap-4 items-center justify-between bg-white text-xs">
        
        {/* Search input */}
        <div class="relative w-full md:w-80">
          <span class="absolute inset-y-0 left-0 flex items-center pl-3">
            <Search class="h-4 w-4 text-slate-400" />
          </span>
          <input
            type="text"
            placeholder="Search paragraph text..."
            value={searchTerm}
            onChange={(e) => setSearchTerm(e.target.value)}
            class="w-full pl-9 pr-4 py-1.5 bg-slate-50 hover:bg-slate-100 border-slate-200 focus:bg-white border rounded-lg text-xs outline-none focus:ring-2 focus:ring-indigo-500 transition-all font-medium"
          />
        </div>

        {/* Filters Group */}
        <div class="flex flex-wrap items-center gap-4 w-full md:w-auto justify-end">
          
          {/* Level selector */}
          <div class="flex items-center gap-1.5 text-slate-500">
            <Filter class="w-3.5 h-3.5 text-slate-400" />
            <span class="font-semibold">Target Level:</span>
            <select
              value={selectedLevelFilter}
              onChange={(e) => setSelectedLevelFilter(e.target.value)}
              class="text-xs border-slate-200 rounded bg-white py-1 pl-2 pr-6 text-slate-600 outline-none focus:border-indigo-500 transition"
            >
              <option value="ALL">All Categories</option>
              {levelOptionIds.map(levelId => {
                const label = fileAnalysis.paragraphs.find(p => p.matchedLevelId === levelId)?.styleName || levelId;
                return (
                  <option key={levelId} value={levelId}>
                    {label}
                  </option>
                );
              })}
            </select>
          </div>

          {/* Toggle for only failed items */}
          <label class="flex items-center text-xs text-slate-600 select-none cursor-pointer">
            <input 
              type="checkbox" 
              checked={showFailedOnly}
              onChange={() => setShowFailedOnly(!showFailedOnly)}
              class="mr-2 rounded border-slate-300 text-indigo-600 focus:ring-indigo-500" 
            />
            Show Errors Only
          </label>
        </div>
      </div>

      {/* Results details table view */}
      <div class="overflow-x-auto flex-1">
        <table class="w-full text-left border-collapse">
          <thead>
            <tr class="sticky top-0 bg-white z-10 text-[11px] uppercase tracking-wider text-slate-400 border-b border-slate-100 font-sans">
              <th class="px-4 py-3 font-semibold text-center w-12">P#</th>
              <th class="px-4 py-3 font-semibold">Content Snippet</th>
              <th class="px-4 py-3 font-semibold w-56">Detected Formatting</th>
              <th class="px-4 py-3 font-semibold w-40">Expected Standard</th>
              <th class="px-4 py-3 font-semibold text-center w-24">Status</th>
              <th class="px-4 py-3 font-semibold w-72">Issue Description</th>
            </tr>
          </thead>
          <tbody class="text-xs divide-y divide-slate-100">
            {filteredParagraphs.length === 0 ? (
              <tr>
                <td colSpan={6} class="text-center py-12 px-6 text-slate-400 text-xs font-medium">
                  No paragraphs matched the active filter conditions.
                </td>
              </tr>
            ) : (
              filteredParagraphs.map((para) => {
                const isFailed = para.validationStatus === 'Fail';
                const rowStyle = isFailed 
                  ? 'bg-rose-50/20 hover:bg-rose-50/35 transition-colors' 
                  : 'hover:bg-slate-50 transition-colors';

                const isExpanded = !!expandedParas[para.paragraphNumber];

                return (
                  <React.Fragment key={para.paragraphNumber}>
                    <tr class={`${rowStyle} group transition duration-150`}>
                      
                      {/* Num column */}
                      <td class="px-4 py-3 text-xs font-mono text-slate-400 text-center select-none font-semibold">
                        {para.paragraphNumber}
                      </td>

                      {/* Paragraph body texts */}
                      <td class="px-4 py-3 max-w-lg">
                        <p class="italic text-slate-600 font-normal pr-4 select-all selection:bg-indigo-100 break-words leading-relaxed">
                          "{para.text}"
                        </p>
                        
                        {/* Submetadata tags */}
                        <div class="mt-2 flex flex-wrap items-center gap-x-2 gap-y-1 text-[10px] font-medium text-slate-400">
                          <span class="bg-slate-100 text-slate-500 px-1.5 py-0.5 rounded font-mono">
                            Style: {para.styleName || 'Normal'}
                          </span>
                          <span>•</span>
                          <span>Alignment: <strong class="text-slate-500 font-bold">{para.alignment || 'LEFT'}</strong></span>
                          <span>•</span>
                          <span>Indentation: <strong class="text-slate-500 font-mono">{para.indent || 0}pt</strong></span>
                          
                          {para.runs.length > 1 && (
                            <>
                              <span>•</span>
                              <button 
                                onClick={() => toggleExpand(para.paragraphNumber)}
                                class="text-indigo-600 hover:text-indigo-800 flex items-center gap-0.5 transition font-semibold"
                              >
                                {isExpanded ? 'Hide inline runs' : `Expanded runs view (${para.runs.length})`}
                                {isExpanded ? <ChevronUp class="w-3 h-3" /> : <ChevronDown class="w-3 h-3" />}
                              </button>
                            </>
                          )}
                        </div>
                      </td>

                      {/* Formats properties */}
                      <td class="px-4 py-3 text-xs space-y-1 font-sans selection:bg-indigo-100">
                        <div class="flex items-center gap-1.5">
                          <span class="text-[10px] bg-slate-100 px-1.5 py-0.5 rounded text-slate-600 font-mono">
                            {para.dominantFont} / {para.dominantSize}pt / {para.dominantBold ? 'Bold' : 'Regular'}
                          </span>
                        </div>
                      </td>

                      {/* Matched expected Level Name */}
                      <td class="px-4 py-3 text-xs font-semibold text-slate-600">
                        {para.matchedLevelId !== 'unmatched' ? (
                          <span class="text-indigo-600 font-mono text-[11px]">
                            {fileAnalysis.paragraphs.find(p => p.matchedLevelId === para.matchedLevelId)?.styleName || para.matchedLevelId}
                          </span>
                        ) : (
                          <span class="text-rose-500 font-semibold font-mono text-[10px]">Unmatched</span>
                        )}
                      </td>

                      {/* Status Badging */}
                      <td class="px-4 py-3 text-center select-none font-bold text-[10px]">
                        {para.validationStatus === 'Pass' ? (
                          <span class="text-emerald-500">PASSED</span>
                        ) : (
                          <span class="text-rose-500">FAILED</span>
                        )}
                      </td>

                      {/* Warnings and errors Descriptions */}
                      <td class="px-4 py-3 text-xs text-slate-600">
                        {para.issues.length === 0 ? (
                          <span class="text-slate-400 font-normal">
                            Matches {fileAnalysis.paragraphs.find(p => p.matchedLevelId === para.matchedLevelId)?.styleName || 'Normal'} standard
                          </span>
                        ) : (
                          <div class="space-y-1">
                            {para.issues.map((issue, idx) => (
                              <div key={idx} class="font-sans leading-relaxed text-slate-600">
                                <span class="font-medium text-rose-600">{issue.message}</span>
                                {issue.expected && (
                                  <span class="text-[10px] text-slate-400 block font-mono">
                                    Expected: {issue.expected} &bull; Got: {issue.actual}
                                  </span>
                                )}
                              </div>
                            ))}
                          </div>
                        )}
                      </td>
                    </tr>

                    {/* Expanding nested run analysis section */}
                    {isExpanded && para.runs.length > 1 && (
                      <tr class="bg-indigo-50/10">
                        <td colSpan={6} class="px-4 py-3 border-b border-indigo-100/40 select-text">
                          <div class="p-3 bg-white border border-slate-100 rounded-lg shadow-inner">
                            <h4 class="text-[10px] font-bold text-indigo-700 mb-2 font-display uppercase tracking-wide">
                              Nested XML Text Runs Breakdown
                            </h4>
                            <div class="space-y-1.5">
                              {para.runs.map((run, rIdx) => (
                                <div 
                                  key={rIdx} 
                                  class="flex items-center justify-between gap-4 p-2 bg-slate-50 border border-slate-100 rounded text-xs hover:border-slate-200 transition"
                                >
                                  <div class="font-mono text-slate-400 select-none text-[10px]">Run #{rIdx + 1}</div>
                                  <div class="flex-1 bg-white border border-slate-100 px-2 py-0.5 text-slate-700 font-sans italic max-w-lg truncate">
                                    "{run.text}"
                                  </div>
                                  <div class="flex items-center gap-x-2 text-[10px] text-slate-500 font-mono">
                                    <span>Font: <strong class="text-slate-700 font-semibold">{run.fontName || 'Not Set'}</strong></span>
                                    <span>Size: <strong class="text-slate-700 font-semibold">{run.fontSize ? `${run.fontSize}pt` : 'Not Set'}</strong></span>
                                    <span>Bold: <strong class="text-slate-700 font-semibold">{run.bold ? 'True' : 'False'}</strong></span>
                                  </div>
                                </div>
                              ))}
                            </div>
                          </div>
                        </td>
                      </tr>
                    )}
                  </React.Fragment>
                );
              })
            )}
          </tbody>
        </table>
      </div>
    </div>
  );
}
