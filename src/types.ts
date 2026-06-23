/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

export interface StandardRule {
  id: string;
  name: string;
  fontName: string;
  fontSize: number;
  bold: boolean | null;
  alignment: string | null; // e.g. 'LEFT', 'CENTER', 'RIGHT', 'JUSTIFY', null (meaning any or not enforced)
  indent: number | null; // expected left indent in pt, null (meaning any or not enforced)
}

export interface RunDetail {
  text: string;
  fontName?: string;
  fontSize?: number;
  bold?: boolean;
}

export interface ParagraphAnalysis {
  paragraphNumber: number;
  text: string;
  alignment: string | null; // 'LEFT' | 'CENTER' | 'RIGHT' | 'JUSTIFY' | null
  indent: number | null; // left indent in pt
  styleName: string | null;
  runs: RunDetail[];
  dominantFont: string;
  dominantSize: number;
  dominantBold: boolean;
  matchedLevelId: string; // The standard level matched
  validationStatus: 'Pass' | 'Fail';
  issues: ValidationIssue[];
}

export interface ValidationIssue {
  type: 'fontName' | 'fontSize' | 'bold' | 'alignment' | 'indent' | 'unmatched';
  severity: 'error' | 'warning';
  message: string;
  expected?: string;
  actual?: string;
}

export interface FileAnalysisSummary {
  totalParagraphs: number;
  passed: number;
  failed: number;
  issuesCount: number;
}

export interface FileAnalysis {
  fileName: string;
  fileSize: number;
  paragraphs: ParagraphAnalysis[];
  summary: FileAnalysisSummary;
}

export interface ProjectAnalysis {
  files: FileAnalysis[];
  summary: {
    totalFiles: number;
    totalParagraphs: number;
    passedParagraphs: number;
    failedParagraphs: number;
    totalIssues: number;
  };
}
