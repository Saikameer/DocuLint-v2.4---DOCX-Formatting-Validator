import * as XLSX from 'xlsx';
import html2canvas from 'html2canvas';
import { ProjectAnalysis, StandardRule } from '../types';

/**
 * Trigger immediate client-side download of a file from a Blob
 */
export function triggerDownload(blob: Blob, filename: string) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}

/**
 * 1. Export analytical dataset as a multi-sheet Excel (.xlsx) file
 */
export function exportToExcel(analysis: ProjectAnalysis) {
  const wb = XLSX.utils.book_new();

  // ----- Sheet 1: Executive Dashboard Summary -----
  const overviewRows = [
    ["DOCULINT FORMATTING STANDARDS COMPLIANCE REPORT"],
    ["Generated On", new Date().toLocaleString()],
    [],
    ["Metric Description", "Value"],
    ["Total Documents Checked", analysis.summary.totalFiles],
    ["Total Paragraphs Validated", analysis.summary.totalParagraphs],
    ["Passed Paragraphs", analysis.summary.passedParagraphs],
    ["Failed Paragraphs", analysis.summary.failedParagraphs],
    ["Identified Rule Deviations", analysis.summary.totalIssues],
    [
      "Formatting Compliance Score", 
      `${((analysis.summary.passedParagraphs / (analysis.summary.totalParagraphs || 1)) * 100).toFixed(1)}%`
    ]
  ];
  const wsOverview = XLSX.utils.aoa_to_sheet(overviewRows);
  
  // Set basic column widths for overview
  wsOverview['!cols'] = [{ wch: 32 }, { wch: 18 }];
  XLSX.utils.book_append_sheet(wb, wsOverview, "Executive Summary");

  // ----- Sheet 2: File-by-File Statistics -----
  const fileHeaders = ["Document File Name", "File Size (KB)", "Total Paragraphs", "Passed Lines", "Failed Lines", "Rule Deviations Count", "Compliance Score %"];
  const fileRows = analysis.files.map(f => [
    f.fileName,
    (f.fileSize / 1024).toFixed(1),
    f.summary.totalParagraphs,
    f.summary.passed,
    f.summary.failed,
    f.summary.issuesCount,
    `${((f.summary.passed / (f.summary.totalParagraphs || 1)) * 100).toFixed(1)}%`
  ]);
  const wsFiles = XLSX.utils.aoa_to_sheet([fileHeaders, ...fileRows]);
  wsFiles['!cols'] = [{ wch: 25 }, { wch: 15 }, { wch: 15 }, { wch: 15 }, { wch: 15 }, { wch: 20 }, { wch: 20 }];
  XLSX.utils.book_append_sheet(wb, wsFiles, "Document Compliance Overview");

  // ----- Sheet 3: Full Block Line Diagnostics -----
  const diagnosticHeaders = [
    "File Name", 
    "Paragraph #", 
    "Expected Standard Match", 
    "Validation Status", 
    "Extracted Text Snippet", 
    "Dominant Font Style", 
    "Size (pt)", 
    "Bold Enforced", 
    "Alignment Mode", 
    "Indentation (pt)", 
    "Deviations Highlighted"
  ];
  
  const diagnosticRows: any[] = [];
  analysis.files.forEach(file => {
    file.paragraphs.forEach(p => {
      const issuesCombined = p.issues.map(i => `${i.message} (Expected: ${i.expected || 'N/A'}, Actual: ${i.actual || 'N/A'})`).join('\n') || "None";
      const expectedStandard = file.paragraphs.find(x => x.matchedLevelId === p.matchedLevelId)?.styleName || p.matchedLevelId;
      
      diagnosticRows.push([
        file.fileName,
        p.paragraphNumber,
        expectedStandard,
        p.validationStatus,
        p.text,
        p.dominantFont,
        p.dominantSize,
        p.dominantBold ? "Bold" : "Regular",
        p.alignment || "LEFT",
        p.indent || 0,
        issuesCombined
      ]);
    });
  });
  const wsDiagnostics = XLSX.utils.aoa_to_sheet([diagnosticHeaders, ...diagnosticRows]);
  wsDiagnostics['!cols'] = [
    { wch: 22 }, { wch: 12 }, { wch: 22 }, { wch: 15 }, { wch: 50 }, 
    { wch: 18 }, { wch: 10 }, { wch: 12 }, { wch: 15 }, { wch: 15 }, { wch: 45 }
  ];
  XLSX.utils.book_append_sheet(wb, wsDiagnostics, "Rule Deviations Logs");

  // Generate Excel stream
  const excelBuffer = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
  return new Blob([excelBuffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
}

/**
 * 2. Generate a beautifully styled, print-ready Word document (.doc formatted HTML)
 */
export function generateWordReport(analysis: ProjectAnalysis, activeRules: StandardRule[]): Blob {
  const complianceRate = ((analysis.summary.passedParagraphs / (analysis.summary.totalParagraphs || 1)) * 100).toFixed(1);
  const dateStr = new Date().toLocaleDateString(undefined, { dateStyle: 'long' });

  // Let's print out guidelines rule definition blocks inside the word metadata section
  const rulesTableRows = activeRules.map(rule => `
    <tr>
      <td style="font-weight: bold; background-color: #f8fafc; width: 140px;">${rule.name}</td>
      <td>${rule.fontName}</td>
      <td>${rule.fontSize} pt</td>
      <td>${rule.bold === null ? 'Any' : rule.bold ? 'Bold' : 'Regular'}</td>
      <td>${rule.alignment || 'Any'}</td>
      <td>${rule.indent !== null ? `${rule.indent} pt` : 'Any'}</td>
    </tr>
  `).join('');

  // Let's create failure logs rows
  let failedParagraphsCount = 0;
  const failureRows: string[] = [];

  analysis.files.forEach(file => {
    file.paragraphs.forEach(p => {
      if (p.validationStatus === 'Fail') {
        failedParagraphsCount++;
        const issuesMarkup = p.issues.map(i => `
          <div style="margin-bottom: 4px; color: #b91c1c; font-weight: 500;">
            &bull; <strong>${i.message}</strong>
            ${i.expected ? `<span style="font-size: 10px; color: #64748b; font-family: monospace; display: block;">Expected: ${i.expected} | Actual: ${i.actual}</span>` : ''}
          </div>
        `).join('');

        const ruleMatched = file.paragraphs.find(x => x.matchedLevelId === p.matchedLevelId)?.styleName || p.matchedLevelId;

        failureRows.push(`
          <tr>
            <td style="font-size: 11px; font-weight: bold; width: 100px;">${file.fileName}</td>
            <td style="font-size: 11px; text-align: center; width: 35px;">${p.paragraphNumber}</td>
            <td style="font-size: 10px; width: 120px;">
              <span style="font-weight:500;">Matched Standard:</span> <span style="background-color: #f1f5f9; padding: 2px 4px; border-radius: 4px; font-family: monospace;">${ruleMatched}</span><br/>
              <span style="font-weight:500;">Detected properties:</span> <span style="font-family: monospace; color: #475569;">${p.dominantFont} / ${p.dominantSize}pt / ${p.dominantBold ? 'Bold' : 'Reg'}</span>
            </td>
            <td style="font-size: 11px; color: #4b5563; font-style: italic;">"${p.text.length > 180 ? p.text.substring(0, 180) + '...' : p.text}"</td>
            <td style="font-size: 11px; width: 180px;">${issuesMarkup}</td>
          </tr>
        `);
      }
    });
  });

  const failureSection = failedParagraphsCount > 0 
    ? `
      <h2>Detailed Verification Failures Log</h2>
      <p style="font-size: 12px; color: #475569; margin-bottom: 12px;">The following lines containing non-compliant styling were automatically flagged by DocuLint validator:</p>
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
          ${failureRows.join('')}
        </tbody>
      </table>
    `
    : `
      <div style="background-color: #ecfdf5; border: 1px solid #10b981; padding: 15px; border-radius: 8px; margin-top: 30px; text-align: center;">
        <h3 style="color: #10b981; margin: 0 0 5px 0;">🎉 Perfect Alignment & Formatting Compliance</h3>
        <p style="color: #064e3b; margin: 0; font-size: 12px;">All scrutinized paragraphs matched your styled templates flawlessly. No deviations to remedy!</p>
      </div>
    `;

  const htmlContent = `
    <html xmlns:o='urn:schemas-microsoft-com:office:office' xmlns:w='urn:schemas-microsoft-com:office:word' xmlns='http://www.w3.org/TR/REC-html40'>
    <head>
      <meta charset="utf-8">
      <title>DocuLint Standards Compliance Summary Report</title>
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
      </style>
    </head>
    <body>
      <div class="header-title">DocuLint v2.4 Assessment Report</div>
      <div class="meta-subtitle">AUTOGENERATED ON ${dateStr.toUpperCase()} | RECTOR COMPLIANCE VALIDATOR</div>

      <table class="metric-container" style="border: 1px solid #e2e8f0; background-color: #f8fafc;">
        <tr>
          <td class="metric-cell" style="border:none;">
            <div class="metric-desc">Scanned Documents</div>
            <div class="metric-val">${analysis.summary.totalFiles} docs</div>
          </td>
          <td class="metric-cell" style="border:none;">
            <div class="metric-desc">Aggregate Paragraphs</div>
            <div class="metric-val">${analysis.summary.totalParagraphs} lines</div>
          </td>
          <td class="metric-cell" style="border:none;">
            <div class="metric-desc">Compliance Score</div>
            <div class="metric-val" style="color: ${Number(complianceRate) >= 90 ? '#10b981' : Number(complianceRate) >= 70 ? '#d97706' : '#ef4444'}">${complianceRate}%</div>
          </td>
          <td class="metric-cell" style="border:none;">
            <div class="metric-desc font-semibold" style="color: #b91c1c;">Rule Breaches</div>
            <div class="metric-val" style="color: #ef4444;">${analysis.summary.totalIssues} issues</div>
          </td>
        </tr>
      </table>

      <h2>Enforced Layout Guidelines Standard Rules</h2>
      <p style="font-size: 12px; color: #475569;">The document layout scans were strictly verified against the following rule matrix definitions:</p>
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
          ${rulesTableRows}
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

      ${failureSection}

      <div style="margin-top: 40px; padding-top: 15px; border-top: 1px solid #e2e8f0; font-family: monospace; font-size: 9px; color: #94a3b8; text-align: center;">
        📄 Generated securely offline entirely within the local browser container. End of DocuLint report.
      </div>
    </body>
    </html>
  `;

  return new Blob([htmlContent], { type: 'application/msword' });
}

/**
 * 3. Render any HTML element as a client-side high-quality PNG screenshot image
 */
export async function captureScreenshot(elementId: string): Promise<string | null> {
  const element = document.getElementById(elementId);
  if (!element) {
    console.error(`captureScreenshot: Element with ID '${elementId}' not found in DOM.`);
    return null;
  }

  try {
    // Scroll element briefly into full view if necessary or set background style for flawless capture
    const backgroundStyle = window.getComputedStyle(element).backgroundColor;
    const hasTransparentBg = backgroundStyle === 'transparent' || backgroundStyle === 'rgba(0, 0, 0, 0)';
    
    const canvas = await html2canvas(element, {
      backgroundColor: hasTransparentBg ? '#f8fafc' : backgroundStyle,
      scale: 2, // Retinal displays quality scaling boost
      logging: false,
      useCORS: true, 
      allowTaint: true,
      onclone: (clonedDoc) => {
        // Optional tweaking of cloned DOM representation (e.g. hiding expand buttons or temporary inputs)
        const clonedElement = clonedDoc.getElementById(elementId);
        if (clonedElement) {
          clonedElement.style.padding = '24px';
          clonedElement.style.borderRadius = '16px';
        }
      }
    });

    return canvas.toDataURL('image/png');
  } catch (err) {
    console.error("captureScreenshot: Canvas rendering failed.", err);
    return null;
  }
}
