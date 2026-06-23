/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React, { useState, useRef } from 'react';
import { StandardRule, ProjectAnalysis, FileAnalysis } from './types';
import { parseDocxFile, DEFAULT_STANDARDS, runValidation } from './utils/docxParser';
import { PYTHON_APP, PYTHON_REQUIREMENTS, PYTHON_README } from './utils/pythonTemplates';
import JSZip from 'jszip';

// Components
import StandardsEditor from './components/StandardsEditor';
import DashboardStats from './components/DashboardStats';
import FileReport from './components/FileReport';
import ExportHub from './components/ExportHub';

// Icons
import {
  FolderOpen,
  FileText,
  AlertCircle,
  Play,
  ArrowRight,
  BookOpen,
  FileCode,
  Download,
  AlertTriangle,
  UploadCloud,
  CheckCircle2,
  Trash2,
  Terminal,
  Clock
} from 'lucide-react';

export default function App() {
  const [standards, setStandards] = useState<StandardRule[]>(DEFAULT_STANDARDS);
  const [projectAnalysis, setProjectAnalysis] = useState<ProjectAnalysis | null>(null);
  const [isAnalyzing, setIsAnalyzing] = useState(false);
  const [rawFiles, setRawFiles] = useState<File[]>([]);
  const [folderPath, setFolderPath] = useState('');
  const [activeFileIndex, setActiveFileIndex] = useState(0);
  const [viewTab, setViewTab] = useState<'applet' | 'python'>('applet');

  const fileInputRef = useRef<HTMLInputElement>(null);
  const folderInputRef = useRef<HTMLInputElement>(null);

  // Trigger re-validation when standards are edited
  const handleUpdateStandards = (newStandards: StandardRule[]) => {
    setStandards(newStandards);
    if (projectAnalysis) {
      // Re-validate existing files with the overridden standards instantly!
      const updatedFiles = projectAnalysis.files.map(file => {
        const updatedParagraphs = file.paragraphs.map(p => {
          return runValidation(
            p.paragraphNumber,
            p.text,
            p.alignment,
            p.indent,
            p.styleName,
            p.runs,
            p.dominantFont,
            p.dominantSize,
            p.dominantBold,
            newStandards
          );
        });

        const passed = updatedParagraphs.filter(p => p.validationStatus === 'Pass').length;
        const failed = updatedParagraphs.length - passed;
        const issuesCount = updatedParagraphs.reduce((sum, p) => sum + p.issues.length, 0);

        return {
          ...file,
          paragraphs: updatedParagraphs,
          summary: {
            totalParagraphs: updatedParagraphs.length,
            passed,
            failed,
            issuesCount
          }
        };
      });

      const totalParagraphs = updatedFiles.reduce((sum, f) => sum + f.summary.totalParagraphs, 0);
      const passedParagraphs = updatedFiles.reduce((sum, f) => sum + f.summary.passed, 0);
      const failedParagraphs = totalParagraphs - passedParagraphs;
      const totalIssues = updatedFiles.reduce((sum, f) => sum + f.summary.issuesCount, 0);

      setProjectAnalysis({
        files: updatedFiles,
        summary: {
          totalFiles: updatedFiles.length,
          totalParagraphs,
          passedParagraphs,
          failedParagraphs,
          totalIssues
        }
      });
    }
  };

  const handleResetStandards = () => {
    handleUpdateStandards(DEFAULT_STANDARDS);
  };

  // Process selected file lists
  const processFiles = async (files: File[]) => {
    const docxFiles = files.filter(f => f.name.endsWith('.docx'));
    if (docxFiles.length === 0) {
      alert('Formatting Analyzer scans only .docx files. Please select a valid folder or group of Word files.');
      return;
    }

    setIsAnalyzing(true);
    setRawFiles(docxFiles);

    try {
      const reports: FileAnalysis[] = [];
      
      for (const file of docxFiles) {
        try {
          const paragraphs = await parseDocxFile(file, standards);
          
          const total = paragraphs.length;
          const passed = paragraphs.filter(p => p.validationStatus === 'Pass').length;
          const failed = total - passed;
          const issuesCount = paragraphs.reduce((sum, p) => sum + p.issues.length, 0);

          reports.push({
            fileName: file.name,
            fileSize: file.size,
            paragraphs,
            summary: {
              totalParagraphs: total,
              passed,
              failed,
              issuesCount
            }
          });
        } catch (err) {
          console.error(`Failed parsing file ${file.name}:`, err);
        }
      }

      const totalParagraphs = reports.reduce((sum, f) => sum + f.summary.totalParagraphs, 0);
      const passedParagraphs = reports.reduce((sum, f) => sum + f.summary.passed, 0);
      const failedParagraphs = totalParagraphs - passedParagraphs;
      const totalIssues = reports.reduce((sum, f) => sum + f.summary.issuesCount, 0);

      setProjectAnalysis({
        files: reports,
        summary: {
          totalFiles: reports.length,
          totalParagraphs,
          passedParagraphs,
          failedParagraphs,
          totalIssues
        }
      });
      setActiveFileIndex(0);
    } catch (error) {
      console.error('Core scan error:', error);
      alert('An error occurred during file scanning.');
    } finally {
      setIsAnalyzing(false);
    }
  };

  const handleFileChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    if (e.target.files && e.target.files.length > 0) {
      processFiles(Array.from(e.target.files));
    }
  };

  const handleFolderChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    if (e.target.files && e.target.files.length > 0) {
      const fileList = Array.from(e.target.files) as any[];
      // Grab folder name from first item path
      const firstPath = fileList[0]?.webkitRelativePath;
      if (firstPath) {
        const pathParts = firstPath.split('/');
        if (pathParts.length > 1) {
          setFolderPath(`./${pathParts[0]}`);
        }
      }
      processFiles(fileList as File[]);
    }
  };

  const handleDragOver = (e: React.DragEvent) => {
    e.preventDefault();
  };

  const handleDrop = (e: React.DragEvent) => {
    e.preventDefault();
    if (e.dataTransfer.files && e.dataTransfer.files.length > 0) {
      processFiles(Array.from(e.dataTransfer.files));
    }
  };

  // Export dynamically-generated combined JSON report
  const downloadJsonReport = () => {
    if (!projectAnalysis) return;
    const blob = new Blob([JSON.stringify(projectAnalysis, null, 2)], { type: 'application/json' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `docx_formatting_validation_report_${new Date().toISOString().slice(0, 10)}.json`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  };

  // Export tabular CSV of current validation report
  const downloadCsvReportForActiveFile = () => {
    if (!projectAnalysis) return;
    const fileReport = projectAnalysis.files[activeFileIndex];
    if (!fileReport) return;

    let csvContent = "data:text/csv;charset=utf-8,";
    csvContent += "Paragraph Number,Text,Main Font,Main Font Size,Is Bold,Alignment,Indent (pt),Expected Level Matches,Status,Identified Issues\n";

    fileReport.paragraphs.forEach(p => {
      const textEscaped = p.text.replace(/"/g, '""').replace(/\n/g, ' ');
      const issuesEscaped = p.issues.map(i => i.message).join('; ').replace(/"/g, '""');
      
      const row = [
        p.paragraphNumber,
        `"${textEscaped}"`,
        `"${p.dominantFont}"`,
        `"${p.dominantSize}"`,
        `"${p.dominantBold}"`,
        `"${p.alignment || 'LEFT'}"`,
        `"${p.indent || 0}"`,
        `"${fileReport.paragraphs.find(x => x.matchedLevelId === p.matchedLevelId)?.styleName || p.matchedLevelId}"`,
        `"${p.validationStatus}"`,
        `"${issuesEscaped || 'None'}"`
      ].join(",");
      csvContent += row + "\n";
    });

    const encodedUri = encodeURI(csvContent);
    const a = document.createElement('a');
    a.href = encodedUri;
    a.download = `${fileReport.fileName.replace('.docx', '')}_formatting_validation_report.csv`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
  };

  const downloadAllFilesCsvReport = () => {
    if (!projectAnalysis) return;
    
    let csvContent = "\ufeff"; // BOM for excel
    csvContent += "File Name,Paragraph Number,Text,Main Font,Size (pt),Bold,Alignment,Indent (pt),Expected Standard,Status,Issues Highlighted\n";

    projectAnalysis.files.forEach(file => {
      file.paragraphs.forEach(p => {
        const textEscaped = p.text.replace(/"/g, '""').replace(/\n/g, ' ');
        const issuesEscaped = p.issues.map(i => i.message).join('; ').replace(/"/g, '""');
        const matchedLabel = file.paragraphs.find(x => x.matchedLevelId === p.matchedLevelId)?.styleName || p.matchedLevelId;

        const row = [
          `"${file.fileName}"`,
          p.paragraphNumber,
          `"${textEscaped}"`,
          `"${p.dominantFont}"`,
          p.dominantSize,
          p.dominantBold ? "True" : "False",
          `"${p.alignment || 'LEFT'}"`,
          p.indent || 0,
          `"${matchedLabel}"`,
          `"${p.validationStatus}"`,
          `"${issuesEscaped || 'None'}"`
        ].join(",");
        csvContent += row + "\n";
      });
    });

    const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `docx_scanned_folder_formatting_validation_report.csv`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  };

  // Export full python local starter tool as ready ZIP
  const downloadPythonBundleZip = async () => {
    try {
      const zip = new JSZip();
      zip.file('app.py', PYTHON_APP);
      zip.file('requirements.txt', PYTHON_REQUIREMENTS);
      zip.file('README.md', PYTHON_README);
      
      const blob = await zip.generateAsync({ type: 'blob' });
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = 'docx_standards_analyzer_python_app.zip';
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      URL.revokeObjectURL(url);
    } catch (err) {
      alert('Zip packing failed: ' + err);
    }
  };

  const clearScans = () => {
    setProjectAnalysis(null);
    setRawFiles([]);
    setFolderPath('');
  };

  const activeFileAnalysis = projectAnalysis?.files[activeFileIndex];

  return (
    <div class="min-h-screen bg-slate-50 text-slate-800 flex flex-col p-6 font-sans selection:bg-indigo-100 max-w-[1280px] mx-auto">
      
      {/* Primary Bento Top App Bar */}
      <header class="flex flex-col md:flex-row md:items-center justify-between mb-6 bg-white p-4 rounded-xl border border-slate-200 shadow-sm gap-4 select-none">
        <div class="flex items-center space-x-3">
          <div class="bg-indigo-600 p-2 rounded-lg text-white">
            <FileText class="w-6 h-6" />
          </div>
          <div>
            <h1 class="text-xl font-bold tracking-tight text-slate-800 font-display flex items-center gap-1.5 leading-none">
              DocuLint <span class="text-indigo-600 font-semibold">v2.4</span>
            </h1>
            <p class="text-[10px] text-slate-400 mt-1 font-mono truncate max-w-xs md:max-w-none">
              DOCX Structural Formatting Standards Validator
            </p>
          </div>
        </div>

        {/* Dynamic Folder Path Mockup */}
        <div class="flex-1 max-w-xl md:mx-6">
          <div class="relative w-full">
            <span class="absolute inset-y-0 left-0 flex items-center pl-3">
              <FolderOpen class="h-4 w-4 text-slate-400" />
            </span>
            <input 
              type="text" 
              readOnly 
              class="w-full pl-9 pr-4 py-1.5 bg-slate-100/70 border-transparent rounded-lg text-xs font-mono text-slate-600 border focus:bg-white transition-all outline-none" 
              value={folderPath || "./Workspace/WordDocuments/Standards"} 
            />
          </div>
        </div>
        
        {/* Main Tabs and Actions */}
        <div class="flex items-center gap-3">
          {projectAnalysis && viewTab === 'applet' && (
            <ExportHub analysis={projectAnalysis} standards={standards} />
          )}
          <div class="flex items-center gap-1 bg-slate-100 p-1 rounded-lg shrink-0">
            <button
              onClick={() => setViewTab('applet')}
              class={`px-3 py-1 text-xs font-semibold rounded-md flex items-center gap-1.5 transition ${
                viewTab === 'applet' 
                  ? 'bg-white text-slate-900 shadow-sm' 
                  : 'text-slate-500 hover:text-slate-900'
              }`}
            >
              Analyze Live
            </button>
            <button
              onClick={() => setViewTab('python')}
              class={`px-3 py-1 text-xs font-semibold rounded-md flex items-center gap-1.5 transition ${
                viewTab === 'python' 
                  ? 'bg-white text-slate-900 shadow-sm' 
                  : 'text-slate-500 hover:text-slate-900'
              }`}
            >
              Python Package
            </button>
          </div>
        </div>
      </header>

      {/* Main Content Pane */}
      <main class="flex-1 flex flex-col gap-6">
        
        {viewTab === 'applet' ? (
          <div class="space-y-6 flex flex-col">
            
            {/* Standards definition manager bento box */}
            <StandardsEditor 
              standards={standards} 
              onUpdateStandards={handleUpdateStandards} 
              onReset={handleResetStandards} 
            />

            {/* Verification Results Panel or Initial File Selector Box */}
            {!projectAnalysis ? (
              <div 
                onDragOver={handleDragOver}
                onDrop={handleDrop}
                class="bg-white border border-slate-200 rounded-xl p-10 text-center shadow-sm max-w-2xl mx-auto flex flex-col items-center gap-5 w-full mt-4"
              >
                <div class="p-4 bg-indigo-50 text-indigo-600 rounded-full">
                  <UploadCloud class="w-10 h-10 animate-bounce" />
                </div>
                <div>
                  <h3 class="text-md font-bold text-slate-800 font-display">Scan folders or files for validation</h3>
                  <p class="text-xs text-slate-500 mt-1 max-w-sm mx-auto leading-relaxed">
                    Instantly parse paragraph-level XML styles directly. All calculations are executed securely and entirely in your local browser sandbox.
                  </p>
                </div>

                <div class="flex flex-col sm:flex-row items-center gap-3 w-full max-w-sm justify-center">
                  {/* Select Folder Button */}
                  <button
                    onClick={() => folderInputRef.current?.click()}
                    disabled={isAnalyzing}
                    class="w-full sm:w-auto px-4 py-2 bg-indigo-600 text-white rounded-lg text-xs font-semibold hover:bg-indigo-700 active:scale-95 transition-all shadow-sm flex items-center justify-center gap-1.5"
                  >
                    <FolderOpen class="w-4 h-4" />
                    Select Folder
                  </button>

                  {/* Select Files */}
                  <button
                    onClick={() => fileInputRef.current?.click()}
                    disabled={isAnalyzing}
                    class="w-full sm:w-auto px-4 py-2 bg-white text-slate-700 border border-slate-200 rounded-lg text-xs font-semibold hover:bg-slate-50 active:scale-95 transition-all flex items-center justify-center gap-1.5"
                  >
                    <FileText class="w-4 h-4" />
                    Select Individual Files
                  </button>
                </div>

                {/* Subtext info */}
                <div class="text-[10px] text-slate-400 flex items-center gap-2 font-medium bg-slate-50 px-3 py-1.5 rounded-full select-none">
                  <Clock class="w-3.5 h-3.5 text-indigo-500" />
                  <span>Supports WebKit directory stream parsing for quick folder imports</span>
                </div>

                {/* Hidden HTML upload actions */}
                <input 
                  type="file" 
                  ref={folderInputRef}
                  onChange={handleFolderChange} 
                  webkitdirectory="" 
                  directory="" 
                  multiple 
                  class="hidden" 
                />
                <input 
                  type="file" 
                  ref={fileInputRef}
                  onChange={handleFileChange} 
                  multiple 
                  accept=".docx"
                  class="hidden" 
                />
              </div>
            ) : (
              <div class="space-y-6 flex flex-col">
                
                {/* Dashboard summary KPIs */}
                <DashboardStats analysis={projectAnalysis} />

                {/* Scanned Outputs dashboard split */}
                <div class="grid grid-cols-1 lg:grid-cols-12 gap-6 items-start">
                  
                  {/* Left panel: List of files inside analyzed folder */}
                  <div class="lg:col-span-4 bg-white rounded-xl border border-slate-200 shadow-sm p-4 h-[580px] flex flex-col">
                    <div class="flex items-center justify-between mb-4 pb-2 border-b border-slate-100">
                      <div class="flex items-center gap-2">
                        <FolderOpen class="w-4 h-4 text-indigo-600" />
                        <h4 class="text-xs font-bold uppercase tracking-wider text-slate-400">
                          Scanned Files ({projectAnalysis.files.length})
                        </h4>
                      </div>
                      <button
                        onClick={clearScans}
                        class="p-1 px-2 text-[10px] bg-rose-50 text-rose-600 rounded flex items-center gap-1 hover:bg-rose-100 active:scale-95 transition font-semibold"
                      >
                        <Trash2 class="w-3.5 h-3.5" />
                        Clear Folder
                      </button>
                    </div>

                    {/* Scroll list of doc outputs */}
                    <div class="flex-1 overflow-y-auto space-y-2 pr-1">
                      {projectAnalysis.files.map((file, idx) => {
                        const isActive = idx === activeFileIndex;
                        const issuesCount = file.summary.issuesCount;

                        return (
                          <button
                            key={idx}
                            onClick={() => setActiveFileIndex(idx)}
                            class={`w-full text-left p-3.5 rounded-xl border flex items-center justify-between gap-3 transition-all ${
                              isActive
                                ? 'bg-indigo-600 text-white border-indigo-600 shadow-sm'
                                : 'bg-slate-50/70 border-slate-100 text-slate-700 hover:bg-slate-50'
                            }`}
                          >
                            <div class="truncate space-y-0.5 flex-1">
                              <p class="text-xs font-bold truncate font-display tracking-tight pr-1">
                                {file.fileName}
                              </p>
                              <p class={`text-[10px] font-mono ${isActive ? 'text-indigo-200' : 'text-slate-400'}`}>
                                {(file.fileSize / 1024).toFixed(1)} KB &bull; {file.summary.totalParagraphs} lines
                              </p>
                            </div>
                            <div class="shrink-0 text-[10px] font-bold">
                              {issuesCount > 0 ? (
                                <span class={`px-2 py-0.5 rounded-full ${
                                  isActive ? 'bg-indigo-700 text-white shadow-inner' : 'bg-rose-50 text-rose-500'
                                }`}>
                                  {issuesCount} FAILED
                                </span>
                              ) : (
                                <span class={`px-2 py-0.5 rounded-full ${
                                  isActive ? 'bg-indigo-700 text-indigo-100' : 'bg-emerald-50 text-emerald-600'
                                }`}>
                                  ✓ PASSED
                                </span>
                              )}
                            </div>
                          </button>
                        );
                      })}
                    </div>

                    {/* Batch download full sheet button */}
                    <div class="mt-4 pt-3 border-t border-slate-150 flex gap-2">
                      <button
                        onClick={downloadAllFilesCsvReport}
                        class="flex-1 px-4 py-2 bg-indigo-600 text-white rounded-lg text-xs font-semibold hover:bg-indigo-700 active:scale-95 transition-all shadow-sm flex items-center justify-center gap-1.5"
                      >
                        <Download class="w-3.5 h-3.5" />
                        Export Folder CSV
                      </button>
                      <button
                        onClick={downloadJsonReport}
                        class="p-2 bg-white border border-slate-200 hover:border-indigo-500 hover:bg-slate-50 text-slate-500 hover:text-indigo-600 rounded-lg shadow-sm transition"
                        title="Download aggregate JSON report"
                      >
                        <FileCode class="w-4 h-4" />
                      </button>
                    </div>
                  </div>

                  {/* Right panel: Active file results tables */}
                  <div class="lg:col-span-8 space-y-6">
                    {activeFileAnalysis && (
                      <FileReport 
                        fileAnalysis={activeFileAnalysis} 
                        onDownloadCsv={downloadCsvReportForActiveFile} 
                      />
                    )}
                  </div>

                </div>

              </div>
            )}

            {/* Spinner loader state */}
            {isAnalyzing && (
              <div class="fixed inset-0 bg-slate-900/40 backdrop-blur-sm flex items-center justify-center z-50 animate-fade-in">
                <div class="bg-white p-6 rounded-2xl border border-slate-200 shadow-2xl flex flex-col items-center gap-4 text-center max-w-sm">
                  <div class="relative w-12 h-12">
                    <div class="absolute inset-0 rounded-full border-4 border-slate-100"></div>
                    <div class="absolute inset-0 rounded-full border-4 border-indigo-600 border-t-transparent animate-spin"></div>
                  </div>
                  <div>
                    <h4 class="text-sm font-bold text-slate-800 font-display">Scanning Word Documents</h4>
                    <p class="text-xs text-slate-500 mt-1 max-w-[240px]">
                      Parsing style relationships, processing XML elements chains and formatting standards validations...
                    </p>
                  </div>
                </div>
              </div>
            )}

          </div>
        ) : (
          /* View Python bundle configuration panel */
          <div class="bg-white rounded-xl border border-slate-200 shadow-sm overflow-hidden p-6 animate-fade-in font-sans">
            <div class="max-w-3xl mx-auto space-y-8">
              
              {/* Feature intro */}
              <div class="flex items-start gap-4">
                <div class="p-3 bg-indigo-50 text-indigo-600 rounded-xl">
                  <Terminal class="w-8 h-8" />
                </div>
                <div>
                  <h3 class="text-lg font-bold text-slate-900 font-display">Stand-alone Local Python Tool Package</h3>
                  <p class="text-xs text-slate-500 mt-1 leading-relaxed">
                    If you want to validate folders of formatting standards directly on physical devices or terminal windows, we have compiled a full, working Python + Flask server script. Drag, drop or select directories of docx blocks and get automatic CSV sheets saved directly to disks.
                  </p>
                </div>
              </div>

              {/* Local File Structure Visualization */}
              <div class="p-4 bg-slate-50 border border-slate-100 rounded-xl">
                <h4 class="text-xs font-bold text-slate-400 uppercase tracking-wider mb-2">Package Contents Structure</h4>
                <div class="font-mono text-xs text-slate-600 space-y-1 select-none">
                  <div>📁 docx_validator/</div>
                  <div>├── 📄 app.py <span class="text-indigo-600 font-sans">&mdash; Full paragraph-by-paragraph XML, word parsing & standards analyzer engine web UI</span></div>
                  <div>├── 📄 requirements.txt <span class="text-indigo-600 font-sans">&mdash; Package definitions (flask, python-docx, lxml)</span></div>
                  <div>└── 📄 README.md <span class="text-indigo-600 font-sans">&mdash; Comprehensive steps & guidelines</span></div>
                </div>
              </div>

              {/* Package download CTA */}
              <div class="flex flex-col sm:flex-row items-center justify-between gap-4 p-5 bg-indigo-50/50 border border-indigo-100 rounded-xl">
                <div>
                  <h4 class="text-xs font-bold text-indigo-900">Download Local Python Code Package</h4>
                  <p class="text-[11px] text-indigo-700/80 mt-0.5 font-medium">Includes full standards, alignment & indentation checks, CSV reports engine, and local Flask server.</p>
                </div>
                <button
                  onClick={downloadPythonBundleZip}
                  class="bg-indigo-600 hover:bg-indigo-700 text-white font-bold text-xs px-5 py-3 rounded-lg flex items-center gap-1.5 shadow-sm active:scale-95 transition whitespace-nowrap shrink-0"
                >
                  <Download class="w-4 h-4" />
                  Download Local Python Package (.zip)
                </button>
              </div>

              {/* Quick instructions steps */}
              <div class="space-y-4">
                <h4 class="text-xs font-bold uppercase tracking-wider text-slate-400">Local Setup Steps in 3 Minutes</h4>
                
                <div class="grid grid-cols-1 md:grid-cols-3 gap-6 font-sans">
                  
                  <div class="space-y-1.5">
                    <div class="flex items-center gap-2">
                      <span class="w-5 h-5 rounded-full bg-slate-100 text-slate-600 flex items-center justify-center font-bold text-xs select-none">1</span>
                      <h4 class="text-xs font-bold text-slate-800">Install Packages</h4>
                    </div>
                    <p class="text-[11px] text-slate-500 leading-normal">
                      Open your local command line inside folder and run pip:
                    </p>
                    <pre class="bg-slate-900 text-slate-300 font-mono text-[10px] p-2 rounded leading-tight">pip install -r requirements.txt</pre>
                  </div>

                  <div class="space-y-1.5">
                    <div class="flex items-center gap-2">
                      <span class="w-5 h-5 rounded-full bg-slate-100 text-slate-600 flex items-center justify-center font-bold text-xs select-none">2</span>
                      <h4 class="text-xs font-bold text-slate-800">Launch Backend</h4>
                    </div>
                    <p class="text-[11px] text-slate-500 leading-normal">
                      Start the Flask application runner:
                    </p>
                    <pre class="bg-slate-900 text-slate-300 font-mono text-[10px] p-2 rounded leading-tight">python app.py</pre>
                  </div>

                  <div class="space-y-1.5">
                    <div class="flex items-center gap-2">
                      <span class="w-5 h-5 rounded-full bg-slate-100 text-slate-600 flex items-center justify-center font-bold text-xs select-none">3</span>
                      <h4 class="text-xs font-bold text-slate-800">Open Dashboard</h4>
                    </div>
                    <p class="text-[11px] text-slate-500 leading-normal">
                      Go to local dashboard link in your browser:
                    </p>
                    <pre class="bg-slate-900 text-slate-300 font-mono text-[10px] p-2 rounded leading-tight">http://127.0.0.1:5000</pre>
                  </div>

                </div>

              </div>

            </div>
          </div>
        )}

      </main>
      
      {/* Footer bar */}
      <footer class="mt-8 py-6 text-center text-xs text-slate-400 font-sans font-medium border-t border-slate-200">
        📄 Word DOCX Standards Formatting Validator &mdash; Standard compliant React UI
      </footer>

    </div>
  );
}
