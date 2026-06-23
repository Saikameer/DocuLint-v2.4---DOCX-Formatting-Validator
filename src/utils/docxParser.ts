/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import JSZip from 'jszip';
import { RunDetail, ParagraphAnalysis, ValidationIssue, StandardRule } from '../types';

interface DocDefaults {
  fontName: string | null;
  fontSize: number | null;
  bold: boolean | null;
}

interface ResolvedStyle {
  id: string;
  name: string | null;
  basedOn: string | null;
  fontName: string | null;
  fontSize: number | null;
  bold: boolean | null;
  alignment: string | null;
  indent: number | null;
}

// Helper to get element by local name
function findChildByLocalName(elem: Element, localName: string): Element | null {
  for (let i = 0; i < elem.children.length; i++) {
    const child = elem.children[i];
    if (child.localName === localName) {
      return child;
    }
  }
  return null;
}

// Helper to get attribute by local name
function getAttrByLocalName(elem: Element | null, localName: string): string | null {
  if (!elem) return null;
  for (let i = 0; i < elem.attributes.length; i++) {
    const attr = elem.attributes[i];
    if (attr.localName === localName) {
      return attr.value;
    }
  }
  return null;
}

// Extract properties from a w:rPr element
function extractRunPr(rPrElem: Element | null): { fontName: string | null; fontSize: number | null; bold: boolean | null; isBoldTagPresent: boolean; runStyleId: string | null } {
  const result: {
    fontName: string | null;
    fontSize: number | null;
    bold: boolean | null;
    isBoldTagPresent: boolean;
    runStyleId: string | null;
  } = {
    fontName: null,
    fontSize: null,
    bold: null,
    isBoldTagPresent: false,
    runStyleId: null,
  };

  if (!rPrElem) return result;

  // Font
  const rFonts = findChildByLocalName(rPrElem, 'rFonts');
  if (rFonts) {
    const ascii = getAttrByLocalName(rFonts, 'ascii');
    const hAnsi = getAttrByLocalName(rFonts, 'hAnsi');
    const cs = getAttrByLocalName(rFonts, 'cs');
    const eastAsia = getAttrByLocalName(rFonts, 'eastAsia');
    result.fontName = ascii || hAnsi || cs || eastAsia || null;
  }

  // Size
  const sz = findChildByLocalName(rPrElem, 'sz');
  if (sz) {
    const szVal = getAttrByLocalName(sz, 'val');
    if (szVal) {
      const halfPts = parseInt(szVal, 10);
      if (!isNaN(halfPts)) {
        result.fontSize = halfPts / 2;
      }
    }
  }

  // Bold
  const b = findChildByLocalName(rPrElem, 'b');
  if (b) {
    result.isBoldTagPresent = true;
    const bVal = getAttrByLocalName(b, 'val');
    if (bVal === '0' || bVal === 'false' || bVal === 'none') {
      result.bold = false;
    } else {
      result.bold = true;
    }
  }

  // Check bCs (complex script bold) as fallback
  if (result.bold === null) {
    const bCs = findChildByLocalName(rPrElem, 'bCs');
    if (bCs) {
      const bcsVal = getAttrByLocalName(bCs, 'val');
      if (bcsVal === '0' || bcsVal === 'false' || bcsVal === 'none') {
        result.bold = false;
      } else {
        result.bold = true;
      }
    }
  }

  // Style reference
  const rStyle = findChildByLocalName(rPrElem, 'rStyle');
  if (rStyle) {
    result.runStyleId = getAttrByLocalName(rStyle, 'val');
  }

  return result;
}

// Extract properties from a w:pPr element
function extractParaPr(pPrElem: Element | null): { styleId: string | null; alignment: string | null; indent: number | null } {
  const result: { styleId: string | null; alignment: string | null; indent: number | null } = {
    styleId: null,
    alignment: null,
    indent: null,
  };

  if (!pPrElem) return result;

  const pStyle = findChildByLocalName(pPrElem, 'pStyle');
  if (pStyle) {
    result.styleId = getAttrByLocalName(pStyle, 'val');
  }

  const jc = findChildByLocalName(pPrElem, 'jc');
  if (jc) {
    const val = getAttrByLocalName(jc, 'val');
    if (val) {
      result.alignment = val.toUpperCase(); // LEFT, CENTER, RIGHT, BOTH, JUSTIFY, etc.
    }
  }

  const ind = findChildByLocalName(pPrElem, 'ind');
  if (ind) {
    const left = getAttrByLocalName(ind, 'left');
    if (left) {
      const dxa = parseInt(left, 10);
      if (!isNaN(dxa)) {
        result.indent = Math.round((dxa / 20) * 100) / 100; // dxa to pt, rounded to 2 decimal places
      }
    }
  }

  return result;
}

// Parse document defaults and styles from styles.xml
function parseStylesXml(stylesXmlText: string): { defaults: DocDefaults; styles: Record<string, ResolvedStyle> } {
  const defaults: DocDefaults = { fontName: null, fontSize: null, bold: null };
  const styles: Record<string, ResolvedStyle> = {};

  try {
    const parser = new DOMParser();
    const doc = parser.parseFromString(stylesXmlText, 'application/xml');

    // Parse docDefaults
    const docDefaults = findChildByLocalName(doc.documentElement, 'docDefaults');
    if (docDefaults) {
      const rPrDefault = findChildByLocalName(docDefaults, 'rPrDefault');
      if (rPrDefault) {
        const rPr = findChildByLocalName(rPrDefault, 'rPr');
        if (rPr) {
          const parsed = extractRunPr(rPr);
          defaults.fontName = parsed.fontName;
          defaults.fontSize = parsed.fontSize;
          defaults.bold = parsed.bold;
        }
      }
    }

    // Parse individual styles
    const styleElems = doc.documentElement.getElementsByTagNameNS('*', 'style');
    for (let i = 0; i < styleElems.length; i++) {
      const styleElem = styleElems[i];
      const styleId = getAttrByLocalName(styleElem, 'styleId');
      if (!styleId) continue;

      const nameElem = findChildByLocalName(styleElem, 'name');
      const name = getAttrByLocalName(nameElem, 'val');

      const basedOnElem = findChildByLocalName(styleElem, 'basedOn');
      const basedOn = getAttrByLocalName(basedOnElem, 'val');

      const rPr = findChildByLocalName(styleElem, 'rPr');
      const runPrs = extractRunPr(rPr);

      const pPr = findChildByLocalName(styleElem, 'pPr');
      const paraPrs = extractParaPr(pPr);

      styles[styleId] = {
        id: styleId,
        name: name,
        basedOn: basedOn,
        fontName: runPrs.fontName,
        fontSize: runPrs.fontSize,
        bold: runPrs.bold,
        alignment: paraPrs.alignment,
        indent: paraPrs.indent,
      };
    }
  } catch (error) {
    console.error('Error parsing styles.xml:', error);
  }

  return { defaults, styles };
}

// Walk inheritance chains to get full properties
function walkStyleChain<T>(
  styleId: string | null,
  styles: Record<string, ResolvedStyle>,
  getter: (style: ResolvedStyle) => T | null,
  maxDepth = 10
): T | null {
  let currId = styleId;
  let depth = 0;
  while (currId && depth < maxDepth) {
    const style = styles[currId];
    if (!style) break;
    const val = getter(style);
    if (val !== null && val !== undefined) {
      return val;
    }
    currId = style.basedOn;
    depth++;
  }
  return null;
}

// Parse text chunks out of a w:r element
function getRunText(rElem: Element): string {
  const parts: string[] = [];
  for (let i = 0; i < rElem.children.length; i++) {
    const child = rElem.children[i];
    if (child.localName === 't') {
      if (child.textContent) {
        parts.push(child.textContent);
      }
    } else if (child.localName === 'tab') {
      parts.push('\t');
    } else if (child.localName === 'br') {
      parts.push('\n');
    } else if (child.localName === 'sym') {
      const charAttr = getAttrByLocalName(child, 'char');
      if (charAttr) {
        try {
          const code = parseInt(charAttr, 16);
          if (!isNaN(code)) {
            parts.push(String.fromCharCode(code));
          }
        } catch {
          parts.push(`[sym:${charAttr}]`);
        }
      }
    }
  }
  return parts.join('');
}

// Master parsing function for a .docx file
export async function parseDocxFile(file: File, rules: StandardRule[]): Promise<ParagraphAnalysis[]> {
  try {
    const zip = await JSZip.loadAsync(file);

    // Read styles and main document XMLs
    const stylesXmlText = await zip.file('word/styles.xml')?.async('text') || '';
    const docXmlText = await zip.file('word/document.xml')?.async('text') || '';

    if (!docXmlText) {
      throw new Error("Invalid DOCX style: word/document.xml is missing.");
    }

    // Parse styles
    const { defaults: docDefaults, styles: docStyles } = parseStylesXml(stylesXmlText);

    // Parse document XML
    const parser = new DOMParser();
    const doc = parser.parseFromString(docXmlText, 'application/xml');

    // Find all paragraph tags <w:p>
    const pElems = doc.getElementsByTagNameNS('*', 'p');
    const analyses: ParagraphAnalysis[] = [];
    let pCounter = 0;

    for (let i = 0; i < pElems.length; i++) {
      const pElem = pElems[i];

      // Check if this paragraph is inside another nested structure like w:tbl
      // Only parse direct body paragraph blocks (table structure parsed differently if wanted, but standard is sequentially parsing all w:p is robust and captures all text!)
      // Wait, is it inside tab or header/footer? Yes we can find its text and properties.
      
      const textParts: string[] = [];
      const runDetails: RunDetail[] = [];

      // Get paragraph properties w:pPr
      const pPr = findChildByLocalName(pElem, 'pPr');
      const paraProps = extractParaPr(pPr);
      const styleId = paraProps.styleId;
      const explicitAlignment = paraProps.alignment;
      const explicitIndent = paraProps.indent;

      // Resolve paragraph style details
      const styleName = styleId ? (docStyles[styleId]?.name || styleId) : 'Normal';

      // Resolve base paragraph alignment and indentation
      const baseAlignment = explicitAlignment || walkStyleChain(styleId, docStyles, (s) => s.alignment) || 'LEFT';
      const baseIndent = explicitIndent !== null ? explicitIndent : (walkStyleChain(styleId, docStyles, (s) => s.indent) || 0);

      // Find runs w:r in this paragraph
      const rElems = pElem.getElementsByTagNameNS('*', 'r');

      for (let j = 0; j < rElems.length; j++) {
        const rElem = rElems[j];
        const runText = getRunText(rElem);
        if (!runText) continue;

        textParts.push(runText);

        const rPr = findChildByLocalName(rElem, 'rPr');
        const explicitRunProps = extractRunPr(rPr);

        // Resolve run styling based on standard Word resolution rules:
        // Priority: Run Explicit > Run Style > Paragraph Style > Doc Defaults > Fallback
        
        let valFont = explicitRunProps.fontName;
        if (valFont === null && explicitRunProps.runStyleId) {
          valFont = walkStyleChain(explicitRunProps.runStyleId, docStyles, (s) => s.fontName);
        }
        if (valFont === null) {
          valFont = walkStyleChain(styleId, docStyles, (s) => s.fontName);
        }
        if (valFont === null) {
          valFont = docDefaults.fontName;
        }
        const fontName = valFont || 'Calibri'; // Default word font

        let valSize = explicitRunProps.fontSize;
        if (valSize === null && explicitRunProps.runStyleId) {
          valSize = walkStyleChain(explicitRunProps.runStyleId, docStyles, (s) => s.fontSize);
        }
        if (valSize === null) {
          valSize = walkStyleChain(styleId, docStyles, (s) => s.fontSize);
        }
        if (valSize === null) {
          valSize = docDefaults.fontSize;
        }
        const fontSize = valSize || 11; // Default size in pt

        let valBold = explicitRunProps.bold;
        if (valBold === null && explicitRunProps.runStyleId) {
          valBold = walkStyleChain(explicitRunProps.runStyleId, docStyles, (s) => s.bold);
        }
        if (valBold === null) {
          valBold = walkStyleChain(styleId, docStyles, (s) => s.bold);
        }
        if (valBold === null) {
          valBold = docDefaults.bold;
        }
        const bold = valBold || false; // Default bold is false

        runDetails.push({
          text: runText,
          fontName,
          fontSize,
          bold,
        });
      }

      const fullText = textParts.join('');
      if (!fullText.trim()) {
        continue; // Skip empty paragraphs
      }

      pCounter++;

      // Compute dominant properties for the paragraph
      const fontCounts: Record<string, number> = {};
      const sizeCounts: Record<number, number> = {};
      const boldCounts: Record<string, number> = { 'true': 0, 'false': 0 };

      runDetails.forEach(run => {
        const charCount = run.text.length;
        if (run.fontName) {
          fontCounts[run.fontName] = (fontCounts[run.fontName] || 0) + charCount;
        }
        if (run.fontSize) {
          sizeCounts[run.fontSize] = (sizeCounts[run.fontSize] || 0) + charCount;
        }
        const bk = run.bold ? 'true' : 'false';
        boldCounts[bk] = (boldCounts[bk] || 0) + charCount;
      });

      // Get modes (dominant properties)
      let dominantFont = 'Calibri';
      let maxFontCount = -1;
      Object.entries(fontCounts).forEach(([font, count]) => {
        if (count > maxFontCount) {
          maxFontCount = count;
          dominantFont = font;
        }
      });

      let dominantSize = 11;
      let maxSizeCount = -1;
      Object.entries(sizeCounts).forEach(([sizeStr, count]) => {
        const sizeNum = parseFloat(sizeStr);
        if (count > maxSizeCount) {
          maxSizeCount = count;
          dominantSize = sizeNum;
        }
      });

      const dominantBold = boldCounts['true'] >= boldCounts['false'];

      // Perform validation and level matching
      const analysis = runValidation(
        pCounter,
        fullText,
        baseAlignment,
        baseIndent,
        styleName,
        runDetails,
        dominantFont,
        dominantSize,
        dominantBold,
        rules
      );

      analyses.push(analysis);
    }

    return analyses;
  } catch (error) {
    console.error('Failed to parse docx:', error);
    throw error;
  }
}

// Function to match standard rule and build validation issues
export function runValidation(
  pNumber: number,
  text: string,
  alignment: string | null,
  indent: number | null,
  styleName: string | null,
  runs: RunDetail[],
  dominantFont: string,
  dominantSize: number,
  dominantBold: boolean,
  rules: StandardRule[]
): ParagraphAnalysis {
  const issues: ValidationIssue[] = [];

  // Match the closest standard level
  // Let's compute a distance score to each rule, or use cascading heuristics.
  // Cascadinig heuristic:
  // Heading Level 1: expected size is 20, font "Human Black Condensed"
  // Heading Level 2: expected size is 11.5, font "Human Black Condensed"
  // Heading Level 3: expected size is 10.5, font "Human Black Condensed"
  // Text Level Bold: expected size is 10.5, font "Human Bold Condensed", bold True
  // Normal Text: expected size is 10.5, font "Human Light Condensed", bold False
  // Form Name: expected size is 7, font "Human Light Condensed"

  let matchedRule: StandardRule | null = null;
  let minDiff = Infinity;

  for (const r of rules) {
    let diff = 0;
    
    // Weight font size matching heavily
    diff += Math.abs(dominantSize - r.fontSize) * 10;
    
    // Font name match
    if (dominantFont.toLowerCase() !== r.fontName.toLowerCase()) {
      diff += 5; // Penalty for wrong font
    }

    // Bold match if specified
    if (r.bold !== null && dominantBold !== r.bold) {
      diff += 3;
    }

    // Style name hinting (Heading style name heavily weights towards heading rules)
    const isHeadingStyle = styleName?.toLowerCase().includes('heading');
    const isHeadingRule = r.id.startsWith('heading');
    if (isHeadingStyle && !isHeadingRule) {
      diff += 8;
    } else if (!isHeadingStyle && isHeadingRule) {
      // Normal paragraph styled as heading? Keep lower, but if properties align perfectly allow it.
    }

    if (diff < minDiff) {
      minDiff = diff;
      matchedRule = r;
    }
  }

  // Fallback if no rules matched
  const levelId = matchedRule ? matchedRule.id : 'unmatched';
  const levelName = matchedRule ? matchedRule.name : 'Unknown Standard';

  if (!matchedRule) {
    issues.push({
      type: 'unmatched',
      severity: 'error',
      message: `Paragraph formatting does not match any defined standard. (Font: ${dominantFont}, Size: ${dominantSize}pt, Bold: ${dominantBold ? 'Yes' : 'No'})`,
    });
  } else {
    // 1. Verify font name
    if (dominantFont.toLowerCase() !== matchedRule.fontName.toLowerCase()) {
      issues.push({
        type: 'fontName',
        severity: 'error',
        message: `Incorrect Font for ${matchedRule.name}. Expected "${matchedRule.fontName}", but found "${dominantFont}".`,
        expected: matchedRule.fontName,
        actual: dominantFont,
      });
    }

    // 2. Verify font size
    if (dominantSize !== matchedRule.fontSize) {
      issues.push({
        type: 'fontSize',
        severity: 'error',
        message: `Incorrect Font Size for ${matchedRule.name}. Expected ${matchedRule.fontSize}pt, but found ${dominantSize}pt.`,
        expected: `${matchedRule.fontSize}pt`,
        actual: `${dominantSize}pt`,
      });
    }

    // 3. Verify bold setting
    if (matchedRule.bold !== null && dominantBold !== matchedRule.bold) {
      issues.push({
        type: 'bold',
        severity: 'error',
        message: `Incorrect Bold formatting for ${matchedRule.name}. Expected Bold to be ${matchedRule.bold ? 'True' : 'False'}, but found ${dominantBold ? 'True' : 'False'}.`,
        expected: matchedRule.bold ? 'Bold' : 'Normal',
        actual: dominantBold ? 'Bold' : 'Normal',
      });
    }

    // 4. Verify alignment (if defined)
    if (matchedRule.alignment !== null) {
      const cleanAlign = (alignment || 'LEFT').toUpperCase();
      const expectedAlign = matchedRule.alignment.toUpperCase();
      if (cleanAlign !== expectedAlign && !(expectedAlign === 'LEFT' && cleanAlign === 'BOTH')) { // BOTH is justified, usually counts as left-ish
        issues.push({
          type: 'alignment',
          severity: 'warning',
          message: `Alignment Inconsistency for ${matchedRule.name}. Expected "${expectedAlign}", but paragraph is "${cleanAlign}".`,
          expected: expectedAlign,
          actual: cleanAlign,
        });
      }
    }

    // 5. Verify indentation (if defined)
    if (matchedRule.indent !== null && indent !== null) {
      const diffIndent = Math.abs(indent - matchedRule.indent);
      if (diffIndent > 1.0) { // allow 1pt tolerance (e.g. minor variations)
        issues.push({
          type: 'indent',
          severity: 'warning',
          message: `Indentation Issue for ${matchedRule.name}. Expected left indent of ${matchedRule.indent}pt, but found ${indent}pt.`,
          expected: `${matchedRule.indent}pt`,
          actual: `${indent}pt`,
        });
      }
    }

    // 6. Check for "No Bold character" requirement
    // Wait, the client said: "Detect and flag: ... No Bold character"
    // Let's see: if a paragraph matches "Text Level Bold" or a heading level (which is expected to be bold),
    // but the actual text has segments that aren't bold, or does it mean a Bold paragraph should have bold characters?
    // Let's see: if matchedRule is "Text Level Bold", but none of the runs are bold, then we raise an issue. Or if we matched "Text Level Bold" but some runs are not bold, let's flag as a warn!
    if (matchedRule.id === 'boldText') {
      const hasAnyBold = runs.some(r => r.bold);
      if (!hasAnyBold) {
        issues.push({
          type: 'bold',
          severity: 'error',
          message: `No bold runs found in candidate Text Level Bold paragraph.`,
          expected: 'Bold text runs',
          actual: 'None',
        });
      }
    }
  }

  // Let's do general checks too (e.g. alignment inconsistencies, indentation issues)
  // E.g. Check if any text run has mismatching font names or sizes from paragraph dominant
  runs.forEach((run, index) => {
    if (run.fontName && run.fontName.toLowerCase() !== dominantFont.toLowerCase()) {
      issues.push({
        type: 'fontName',
        severity: 'warning',
        message: `Internal Font Mismatch: Run ${index + 1} uses "${run.fontName}", deviating from paragraph's main font "${dominantFont}".`,
        expected: dominantFont,
        actual: run.fontName,
      });
    }
    if (run.fontSize && run.fontSize !== dominantSize) {
      issues.push({
        type: 'fontSize',
        severity: 'warning',
        message: `Internal Size Mismatch: Run ${index + 1} text is ${run.fontSize}pt, deviating from paragraph's main size ${dominantSize}pt.`,
        expected: `${dominantSize}pt`,
        actual: `${run.fontSize}pt`,
      });
    }
  });

  const validationStatus = issues.filter(issue => issue.severity === 'error').length > 0 ? 'Fail' : 'Pass';

  return {
    paragraphNumber: pNumber,
    text,
    alignment: alignment || 'LEFT',
    indent: indent || 0,
    styleName,
    runs,
    dominantFont,
    dominantSize,
    dominantBold,
    matchedLevelId: levelId,
    validationStatus,
    issues,
  };
}

// Default standard rules based on the prompt instructions
export const DEFAULT_STANDARDS: StandardRule[] = [
  {
    id: 'heading1',
    name: 'Level 1 Heading',
    fontName: 'Human Black Condensed',
    fontSize: 20,
    bold: true,
    alignment: 'LEFT',
    indent: 0,
  },
  {
    id: 'heading2',
    name: 'Level 2 Heading',
    fontName: 'Human Black Condensed',
    fontSize: 11.5,
    bold: true,
    alignment: 'LEFT',
    indent: 0,
  },
  {
    id: 'heading3',
    name: 'Level 3 Heading',
    fontName: 'Human Black Condensed',
    fontSize: 10.5,
    bold: true,
    alignment: 'LEFT',
    indent: 0,
  },
  {
    id: 'normal',
    name: 'Normal Text',
    fontName: 'Human Light Condensed',
    fontSize: 10.5,
    bold: false,
    alignment: 'LEFT',
    indent: 0,
  },
  {
    id: 'boldText',
    name: 'Text Level Bold',
    fontName: 'Human Bold Condensed',
    fontSize: 10.5,
    bold: true,
    alignment: 'LEFT',
    indent: 0,
  },
  {
    id: 'formName',
    name: 'Form Name',
    fontName: 'Human Light Condensed',
    fontSize: 7,
    bold: false,
    alignment: 'LEFT',
    indent: 0,
  },
];
