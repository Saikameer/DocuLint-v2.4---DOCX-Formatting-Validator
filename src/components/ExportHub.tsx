import React, { useState } from 'react';
import { ProjectAnalysis, StandardRule } from '../types';
import { 
  exportToExcel, 
  generateWordReport, 
  captureScreenshot, 
  triggerDownload 
} from '../utils/exporter';
import JSZip from 'jszip';
import { 
  Download, 
  FileSpreadsheet, 
  FileText, 
  Camera, 
  FileArchive, 
  X, 
  CheckCircle, 
  Loader2, 
  Image,
  Sparkles,
  Printer
} from 'lucide-react';

interface ExportHubProps {
  analysis: ProjectAnalysis | null;
  standards: StandardRule[];
}

export default function ExportHub({ analysis, standards }: ExportHubProps) {
  const [isOpen, setIsOpen] = useState(false);
  const [isProcessing, setIsProcessing] = useState(false);
  const [currentStepText, setCurrentStepText] = useState('');
  const [embedScreenshotInWord, setEmbedScreenshotInWord] = useState(true);

  if (!analysis) return null;

  const complianceScore = ((analysis.summary.passedParagraphs / (analysis.summary.totalParagraphs || 1)) * 100).toFixed(1);

  // 1. Export Excel Action
  const handleExportExcel = () => {
    try {
      setIsProcessing(true);
      setCurrentStepText('Compiling multi-sheet Excel formatting columns...');
      const excelBlob = exportToExcel(analysis);
      triggerDownload(excelBlob, `doculint_formatting_validation_report_${new Date().toISOString().slice(0, 10)}.xlsx`);
    } catch (err: any) {
      console.error(err);
      alert('Failed to generate Excel document: ' + err?.message);
    } finally {
      setIsProcessing(false);
    }
  };

  // 2. Export Word Document Action
  const handleExportWord = async () => {
    try {
      setIsProcessing(true);
      let wordBlob: Blob;

      if (embedScreenshotInWord) {
        setCurrentStepText('Rendering and capturing high-definition dashboard statistics screenshot...');
        const screenshotDataUrl = await captureScreenshot('dashboard_stats_panel');
        
        if (screenshotDataUrl) {
          setCurrentStepText('Embedding screenshot into executive MS Word document structure...');
          // Helper to inject screenshot into the Word template
          wordBlob = injectScreenshotIntoWord(screenshotDataUrl);
        } else {
          setCurrentStepText('Failed to capture screenshot. Generating standards Word document without image...');
          wordBlob = generateWordReport(analysis, standards);
        }
      } else {
        setCurrentStepText('Assembling formatting standards Word assessment report...');
        wordBlob = generateWordReport(analysis, standards);
      }

      triggerDownload(wordBlob, `doculint_compliance_assessment_report_${new Date().toISOString().slice(0, 10)}.doc`);
    } catch (err: any) {
      console.error(err);
      alert('Failed to construct Word document: ' + err?.message);
    } finally {
      setIsProcessing(false);
    }
  };

  // 3. Export Standalone Screenshot Action
  const handleExportScreenshot = async () => {
    try {
      setIsProcessing(true);
      setCurrentStepText('Rendering layout container in vector canvas...');
      const imgDataUrl = await captureScreenshot('dashboard_stats_panel');
      if (imgDataUrl) {
        const response = await fetch(imgDataUrl);
        const blob = await response.blob();
        triggerDownload(blob, `doculint_dashboard_screenshot_${new Date().toISOString().slice(0, 10)}.png`);
      } else {
        alert('Could not find dashboard_stats_panel to take a screenshot.');
      }
    } catch (err: any) {
      console.error(err);
      alert('Failed to capture screenshot: ' + err?.message);
    } finally {
      setIsProcessing(false);
    }
  };

  // 4. Export ZIP Package Action
  const handleExportAllZip = async () => {
    try {
      setIsProcessing(true);
      setCurrentStepText('Initializing compressed archive archive...');
      const zip = new JSZip();

      // Excel Sheet
      setCurrentStepText('Packaging Excel sheet inside ZIP container...');
      const excelBlob = exportToExcel(analysis);
      zip.file('1_DocuLint_Formatting_Grid_Logs.xlsx', excelBlob);

      // Screenshot capture
      setCurrentStepText('Taking vector screenshot of verification stats...');
      const screenshotDataUrl = await captureScreenshot('dashboard_stats_panel');
      
      // Word document
      setCurrentStepText('Structuring compliance Word report...');
      let wordBlob: Blob;
      if (screenshotDataUrl && embedScreenshotInWord) {
        wordBlob = injectScreenshotIntoWord(screenshotDataUrl);
      } else {
        wordBlob = generateWordReport(analysis, standards);
      }
      zip.file('2_DocuLint_Executive_Standards_Report.doc', wordBlob);

      // Standalone screenshot file if captured
      if (screenshotDataUrl) {
        setCurrentStepText('Saving diagnostic reference PNG...');
        // Decode base64 to byte array
        const base64Content = screenshotDataUrl.split(',')[1];
        zip.file('3_DocuLint_Dashboard_Visual_Capture.png', base64Content, { base64: true });
      }

      setCurrentStepText('Compressing files into download bundle...');
      const zipBlob = await zip.generateAsync({ type: 'blob' });
      triggerDownload(zipBlob, `doculint_standards_complete_export_suite_${new Date().toISOString().slice(0, 10)}.zip`);
    } catch (err: any) {
      console.error(err);
      alert('Failed to build ZIP archive package: ' + err?.message);
    } finally {
      setIsProcessing(false);
    }
  };

  // Helper to inject screenshot inside generated HTML MS Word doc
  const injectScreenshotIntoWord = (screenshotDataUrl: string): Blob => {
    const rawWordBlob = generateWordReport(analysis, standards);
    
    // We can read raw text of HTML report from Blob, insert image under summary-container, and return new Blob
    return new Blob([`
      <html xmlns:o='urn:schemas-microsoft-com:office:office' xmlns:w='urn:schemas-microsoft-com:office:word' xmlns='http://www.w3.org/TR/REC-html40'>
      <head>
        <meta charset="utf-8">
        <title>DocuLint Compliance Assessment Report</title>
        <!--[if gte mso 9]>
        <xml>
          <w:WordDocument>
            <w:View>Print</w:View>
            <w:Zoom>100</w:Zoom>
          </w:WordDocument>
        </xml>
        <![endif]-->
        <style>
          body { font-family: 'Segoe UI', Arial, sans-serif; color: #1e293b; line-height: 1.5; margin: 40px; }
          .header-title { color: #4f46e5; border-bottom: 3px solid #4f46e5; padding-bottom: 8px; font-size: 26px; font-weight: bold; text-transform: uppercase; letter-spacing: 0.5px; }
          .meta-subtitle { font-size: 11px; color: #64748b; font-family: monospace; margin-top: 4px; margin-bottom: 25px; }
          h2 { color: #1e293b; margin-top: 30px; font-size: 16px; border-bottom: 1px solid #cbd5e1; padding-bottom: 6px; font-weight: bold; }
          p { font-size: 12px; margin-top: 0; margin-bottom: 12px; color: #334155; }
          .metric-container { background-color: #f8fafc; border: 1px solid #e2e8f0; border-radius: 8px; padding: 15px; margin-bottom: 25px; width: 100%; border-collapse: separate; }
          .metric-cell { padding: 8px 15px; font-size: 12px; }
          .metric-val { font-size: 20px; font-weight: bold; color: #4f46e5; }
          .metric-desc { font-weight: 500; color: #475569; }
          table { width: 100%; border-collapse: collapse; margin-top: 15px; margin-bottom: 20px; font-family: 'Segoe UI', Arial, sans-serif; }
          th { background-color: #f1f5f9; color: #334155; font-weight: bold; border: 1px solid #cbd5e1; padding: 8px; font-size: 11px; text-transform: uppercase; text-align: left; }
          td { border: 1px solid #cbd5e1; padding: 8px; font-size: 11px; vertical-align: top; }
          .badge-passed { color: #10b981; font-weight: bold; }
          .badge-failed { color: #ef4444; font-weight: bold; }
          .img-box { margin-bottom: 20px; padding: 10px; border: 1px solid #e2e8f0; background: #ffffff; text-align: center; }
        </style>
      </head>
      <body>
        <div class="header-title">DocuLint v2.4 Assessment Report</div>
        <div class="meta-subtitle">AUTOGENERATED ON ${new Date().toLocaleDateString().toUpperCase()} | EMBEDDED DASHBOARD SCREENSHOT INCLUDED</div>

        <table class="metric-container" style="border: 1px solid #e2e8f0; background-color: #f8fafc;">
          <tr>
            <td class="metric-cell" style="border:none;">
              <div class="metric-desc">Scanned Documents</div>
              <div class="metric-val">${analysis.summary.totalFiles} files</div>
            </td>
            <td class="metric-cell" style="border:none;">
              <div class="metric-desc">Aggregate Paragraphs</div>
              <div class="metric-val">${analysis.summary.totalParagraphs} lines</div>
            </td>
            <td class="metric-cell" style="border:none;">
              <div class="metric-desc">Compliance Score</div>
              <div class="metric-val" style="color: ${Number(complianceScore) >= 90 ? '#10b981' : Number(complianceScore) >= 70 ? '#d97706' : '#ef4444'}">${complianceScore}%</div>
            </td>
            <td class="metric-cell" style="border:none;">
              <div class="metric-desc font-semibold" style="color: #b91c1c;">Rule Breaches</div>
              <div class="metric-val" style="color: #ef4444;">${analysis.summary.totalIssues} issues</div>
            </td>
          </tr>
        </table>

        <h2>Visual Verification Screenshot (Dashboard Live Snapshot)</h2>
        <p style="font-size: 11px; color:#64748b; font-style:italic;">The following image is a verified render generated by the client browser during doc validation:</p>
        <div class="img-box">
          <img src="${screenshotDataUrl}" alt="DocuLint Dashboard Summary Chart" style="max-width: 600px; max-height: 400px; border: 1px solid #cbd5e1; filter: drop-shadow(0px 2px 4px rgba(0,0,0,0.1));" />
        </div>

        <h2>Enforced Layout Guidelines Standard Rules</h2>
        <table style="width: 100%; border-collapse: collapse; border: 1px solid #cbd5e1;">
          <thead>
            <tr>
              <th style="padding: 8px; border: 1px solid #cbd5e1; font-size: 11px;">Standard Hierarchy</th>
              <th style="padding: 8px; border: 1px solid #cbd5e1; font-size: 11px;">Enforced Font family</th>
              <th style="padding: 8px; border: 1px solid #cbd5e1; font-size: 11px;">Enforced Size</th>
              <th style="padding: 8px; border: 1px solid #cbd5e1; font-size: 11px;">Bold Enforce</th>
              <th style="padding: 8px; border: 1px solid #cbd5e1; font-size: 11px;">Alignment</th>
              <th style="padding: 8px; border: 1px solid #cbd5e1; font-size: 11px;">Expected Left Indent</th>
            </tr>
          </thead>
          <tbody>
            ${standards.map(rule => `
              <tr>
                <td style="font-weight: bold; background-color: #f8fafc; width: 140px;">${rule.name}</td>
                <td>${rule.fontName}</td>
                <td>${rule.fontSize} pt</td>
                <td>${rule.bold === null ? 'Any' : rule.bold ? 'Bold' : 'Regular'}</td>
                <td>${rule.alignment || 'Any'}</td>
                <td>${rule.indent !== null ? `${rule.indent} pt` : 'Any'}</td>
              </tr>
            `).join('')}
          </tbody>
        </table>

        <h2>Scanned Documents Summary</h2>
        <table style="width: 100%; border-collapse: collapse; border: 1px solid #cbd5e1;">
          <thead>
            <tr>
              <th style="padding: 8px; border: 1px solid #cbd5e1; font-size: 11px;">Document File Name</th>
              <th style="padding: 8px; border: 1px solid #cbd5e1; font-size: 11px; text-align: center;">Size</th>
              <th style="padding: 8px; border: 1px solid #cbd5e1; font-size: 11px; text-align: center;">Total checked rows</th>
              <th style="padding: 8px; border: 1px solid #cbd5e1; font-size: 11px; text-align: center;">Passed</th>
              <th style="padding: 8px; border: 1px solid #cbd5e1; font-size: 11px; text-align: center;">Failed</th>
              <th style="padding: 8px; border: 1px solid #cbd5e1; font-size: 11px; text-align: center;">Breaches</th>
              <th style="padding: 8px; border: 1px solid #cbd5e1; font-size: 11px; text-align: center;">Compliance Score</th>
            </tr>
          </thead>
          <tbody>
            ${analysis.files.map(f => {
              const fileRate = ((f.summary.passed / (f.summary.totalParagraphs || 1)) * 100).toFixed(1);
              return `
                <tr>
                  <td style="font-weight: 500;">${f.fileName}</td>
                  <td style="text-align: center;">${(f.fileSize / 1024).toFixed(1)} KB</td>
                  <td style="text-align: center;">${f.summary.totalParagraphs}</td>
                  <td style="text-align: center; color:#10b981; font-weight:bold;">${f.summary.passed}</td>
                  <td style="text-align: center; color:${f.summary.failed > 0 ? '#ef4444' : '#1e293b'}">${f.summary.failed}</td>
                  <td style="text-align: center; color:${f.summary.issuesCount > 0 ? '#ef4444' : '#10b981'}; font-weight:bold;">${f.summary.issuesCount}</td>
                  <td style="text-align: center; font-weight:bold; color: ${Number(fileRate) >= 90 ? '#10b981' : '#d97706'}">${fileRate}%</td>
                </tr>
              `;
            }).join('')}
          </tbody>
        </table>

        <h2>Detailed Verification Failures Log</h2>
        <table style="width: 100%; border-collapse: collapse; border: 1px solid #cbd5e1; font-family: 'Segoe UI', Arial, sans-serif;">
          <thead>
            <tr style="background-color: #ef4444; color: #ffffff;">
              <th style="padding: 8px; border: 1px solid #cbd5e1; font-size: 11px; text-align: left; font-weight: bold; width: 100px;">File Name</th>
              <th style="padding: 8px; border: 1px solid #cbd5e1; font-size: 11px; text-align: center; font-weight: bold; width: 35px;">P#</th>
              <th style="padding: 8px; border: 1px solid #cbd5e1; font-size: 11px; text-align: left; font-weight: bold; width: 120px;">Style Signature</th>
              <th style="padding: 8px; border: 1px solid #cbd5e1; font-size: 11px; text-align: left; font-weight: bold;">Paragraph Content</th>
              <th style="padding: 8px; border: 1px solid #cbd5e1; font-size: 11px; text-align: left; font-weight: bold; width: 180px;">Identified Deviations</th>
            </tr>
          </thead>
          <tbody>
            ${analysis.files.flatMap(file => 
              file.paragraphs.filter(p => p.validationStatus === 'Fail').map(p => {
                const issuesMarkup = p.issues.map(i => `
                  <div style="margin-bottom: 4px; color: #b91c1c; font-weight: 500;">
                    &bull; <strong>${i.message}</strong>
                    ${i.expected ? `<span style="font-size: 10px; color: #64748b; font-family: monospace; display: block;">Expected: ${i.expected} | Actual: ${i.actual}</span>` : ''}
                  </div>
                `).join('');
                const ruleMatched = file.paragraphs.find(x => x.matchedLevelId === p.matchedLevelId)?.styleName || p.matchedLevelId;
                return `
                  <tr>
                    <td style="font-size: 11px; font-weight: bold; width: 100px;">${file.fileName}</td>
                    <td style="font-size: 11px; text-align: center; width: 35px;">${p.paragraphNumber}</td>
                    <td style="font-size: 10px; width: 120px;">
                      <span style="font-weight:500;">Standard:</span> <span style="background-color: #f1f5f9; padding: 2px 4px; border-radius: 4px; font-family: monospace;">${ruleMatched}</span><br/>
                      <span style="font-weight:500;">Detected:</span> <span style="font-family: monospace; color: #475569;">${p.dominantFont} / ${p.dominantSize}pt</span>
                    </td>
                    <td style="font-size: 11px; color: #4b5563; font-style: italic;">"${p.text.length > 180 ? p.text.substring(0, 180) + '...' : p.text}"</td>
                    <td style="font-size: 11px; width: 180px;">${issuesMarkup}</td>
                  </tr>
                `;
              })
            ).join('')}
          </tbody>
        </table>

        <div style="margin-top: 40px; padding-top: 15px; border-top: 1px solid #e2e8f0; font-family: monospace; font-size: 9px; color: #94a3b8; text-align: center;">
          📄 Generated securely offline entirely within the local browser container. End of DocuLint report.
        </div>
      </body>
      </html>
    `], { type: 'application/msword' });
  };

  const handlePrint = () => {
    window.print();
  };

  return (
    <div className="font-sans select-none">
      {/* Premium trigger button with shiny look and custom design */}
      <button
        onClick={() => setIsOpen(true)}
        className="w-full sm:w-auto px-4 py-2 bg-gradient-to-r from-indigo-600 to-indigo-700 hover:from-indigo-700 hover:to-indigo-800 text-white rounded-lg text-xs font-semibold shadow-md inline-flex items-center justify-center gap-2 hover:shadow-lg active:scale-95 transition-all outline-none cursor-pointer"
        id="btn_open_export_hub"
      >
        <Sparkles className="w-4 h-4 animate-pulse text-indigo-200" />
        Export Reports & Screenshots
      </button>

      {isOpen && (
        <div className="fixed inset-0 bg-slate-900/40 backdrop-blur-sm flex items-center justify-center z-50 animate-fade-in p-4">
          <div className="bg-white rounded-2xl border border-slate-200 shadow-2xl max-w-2xl w-full overflow-hidden flex flex-col max-h-[90vh]">
            
            {/* Modal Header */}
            <div className="p-5 border-b border-slate-100 bg-slate-50/50 flex items-center justify-between">
              <div className="flex items-center gap-2">
                <div className="bg-indigo-600 p-1.5 rounded-lg text-white">
                  <Sparkles className="w-5 h-5" />
                </div>
                <div>
                  <h3 className="text-sm font-bold text-slate-800 font-display">Export Reports & Diagnostic Captures</h3>
                  <p className="text-[10px] text-slate-400 font-medium">Download compliance assets or capture graphic diagrams instantly</p>
                </div>
              </div>
              <button 
                onClick={() => setIsOpen(false)}
                className="p-1 px-1.5 text-slate-400 hover:text-slate-600 hover:bg-slate-100 rounded-md transition cursor-pointer"
              >
                <X className="w-4 h-4" />
              </button>
            </div>

            {/* Modal Body */}
            <div className="p-6 overflow-y-auto space-y-6 flex-1">
              
              {/* Stat Indicator Header */}
              <div className="grid grid-cols-3 gap-4 bg-indigo-50/20 p-4 rounded-xl border border-indigo-100/30 text-center">
                <div>
                  <p className="text-[10px] font-bold text-slate-400 uppercase tracking-tight">Compliance Rate</p>
                  <p className={`text-xl font-bold font-display mt-0.5 ${Number(complianceScore) >= 90 ? 'text-emerald-600' : 'text-indigo-600'}`}>
                    {complianceScore}%
                  </p>
                </div>
                <div className="border-x border-slate-100">
                  <p className="text-[10px] font-bold text-slate-400 uppercase tracking-tight">Checked Blocks</p>
                  <p className="text-xl font-bold font-display text-slate-700 mt-0.5">
                    {analysis.summary.totalParagraphs} lines
                  </p>
                </div>
                <div>
                  <p className="text-[10px] font-bold text-rose-400 uppercase tracking-tight">Violations Tracked</p>
                  <p className="text-xl font-bold font-display text-rose-500 mt-0.5">
                    {analysis.summary.totalIssues}
                  </p>
                </div>
              </div>

              {/* Customizable option parameter checkbox */}
              <div className="bg-slate-50 p-3 rounded-lg border border-slate-150 flex items-center justify-between text-xs text-slate-700">
                <label className="flex items-center gap-2 select-none cursor-pointer font-medium text-slate-600">
                  <input
                    type="checkbox"
                    checked={embedScreenshotInWord}
                    onChange={(e) => setEmbedScreenshotInWord(e.target.checked)}
                    className="rounded border-slate-300 text-indigo-600 focus:ring-indigo-500 w-4 h-4"
                  />
                  <span>Attach visual Dashboard statistics rendering inside my reports</span>
                </label>
                <span className="text-[10px] text-indigo-600 font-mono font-bold bg-indigo-50 px-2 py-0.5 rounded">
                  {embedScreenshotInWord ? 'Screenshot ON' : 'Screenshot OFF'}
                </span>
              </div>

              {/* Bento export operations grid */}
              <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
                
                {/* Option 1: Excel */}
                <button
                  onClick={handleExportExcel}
                  disabled={isProcessing}
                  className="p-4 bg-white hover:bg-slate-50 border border-slate-200 hover:border-slate-300 rounded-xl transition-all duration-200 text-left flex items-start gap-3 shadow-sm hover:shadow hover:scale-[1.01] cursor-pointer disabled:opacity-50"
                >
                  <div className="bg-emerald-50 text-emerald-600 p-2 rounded-lg">
                    <FileSpreadsheet className="w-5 h-5" />
                  </div>
                  <div className="space-y-0.5">
                    <h4 className="text-xs font-bold text-slate-800">Export Excel Spreadsheet</h4>
                    <p className="text-[10px] text-slate-400 leading-relaxed font-normal">
                      Saves an `.xlsx` workbook containing executive sheets, scanned document metrics, and itemized failure lists.
                    </p>
                  </div>
                </button>

                {/* Option 2: Word Assessment */}
                <button
                  onClick={handleExportWord}
                  disabled={isProcessing}
                  className="p-4 bg-white hover:bg-slate-50 border border-slate-200 hover:border-slate-300 rounded-xl transition-all duration-200 text-left flex items-start gap-3 shadow-sm hover:shadow hover:scale-[1.01] cursor-pointer disabled:opacity-50"
                >
                  <div className="bg-blue-50 text-blue-600 p-2 rounded-lg">
                    <FileText className="w-5 h-5" />
                  </div>
                  <div className="space-y-0.5">
                    <h4 className="text-xs font-bold text-slate-800">Export Word Assessment</h4>
                    <p className="text-[10px] text-slate-400 leading-relaxed font-normal">
                      Generates a print-ready standards evaluation document. Optionally integrates direct snapshots of stats charts.
                    </p>
                  </div>
                </button>

                {/* Option 3: PNG Screenshot */}
                <button
                  onClick={handleExportScreenshot}
                  disabled={isProcessing}
                  className="p-4 bg-white hover:bg-slate-50 border border-slate-200 hover:border-slate-300 rounded-xl transition-all duration-200 text-left flex items-start gap-3 shadow-sm hover:shadow hover:scale-[1.01] cursor-pointer disabled:opacity-50"
                >
                  <div className="bg-amber-50 text-amber-600 p-2 rounded-lg">
                    <Camera className="w-5 h-5" />
                  </div>
                  <div className="space-y-0.5">
                    <h4 className="text-xs font-bold text-slate-800">Save Stats Screenshot</h4>
                    <p className="text-[10px] text-slate-400 leading-relaxed font-normal">
                      Captures the live interactable dashboard and violations grid as an high-fidelity screenshot image (PNG).
                    </p>
                  </div>
                </button>

                {/* Option 4: Full ZIP Archiver package */}
                <button
                  onClick={handleExportAllZip}
                  disabled={isProcessing}
                  className="p-4 bg-indigo-50/50 hover:bg-indigo-50 border border-indigo-100 hover:border-indigo-200 rounded-xl transition-all duration-200 text-left flex items-start gap-3 shadow-sm hover:shadow hover:scale-[1.01] cursor-pointer disabled:opacity-50"
                  id="btn_zip_archiver"
                >
                  <div className="bg-indigo-600 text-white p-2 rounded-lg shadow-sm">
                    <FileArchive className="w-5 h-5 animate-pulse" />
                  </div>
                  <div className="space-y-0.5">
                    <h4 className="text-xs font-bold text-slate-800 flex items-center gap-1">
                      Download Aggregate ZIP Bundle
                    </h4>
                    <p className="text-[10px] text-indigo-900/70 leading-relaxed font-medium">
                      One-click action packaging XLS database sheets, Word summary, AND visual screenshots inside a single zip archive.
                    </p>
                  </div>
                </button>

              </div>

            </div>

            {/* Modal Actions Footer */}
            <div className="p-4 border-t border-slate-100 bg-slate-50 flex flex-col sm:flex-row sm:items-center justify-between gap-4">
              <button
                onClick={handlePrint}
                className="px-3.5 py-1.5 bg-white text-slate-600 border border-slate-200 rounded-lg text-xs font-semibold hover:bg-slate-50 transition flex items-center justify-center gap-1.5 shadow-sm cursor-pointer"
              >
                <Printer className="w-3.5 h-3.5" />
                Print Layout Results
              </button>

              <button
                onClick={() => setIsOpen(false)}
                className="px-4 py-1.5 bg-slate-800 text-white hover:bg-slate-900 rounded-lg text-xs font-semibold shadow-sm transition cursor-pointer"
              >
                Done
              </button>
            </div>

          </div>
        </div>
      )}

      {/* Process state Loader */}
      {isProcessing && (
        <div className="fixed inset-0 bg-slate-950/20 backdrop-blur-[2px] flex items-center justify-center z-[100] animate-fade-in">
          <div className="bg-white p-5 rounded-xl border border-slate-200 shadow-2xl flex flex-col items-center gap-3 text-center max-w-xs">
            <Loader2 className="w-10 h-10 animate-spin text-indigo-600" />
            <div className="space-y-1">
              <h4 className="text-xs font-bold text-slate-800">Processing Document Assets</h4>
              <p className="text-[10px] text-slate-400 font-mono leading-normal">{currentStepText}</p>
            </div>
          </div>
        </div>
      )}

    </div>
  );
}
