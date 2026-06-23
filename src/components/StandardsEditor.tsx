/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React, { useState } from 'react';
import { StandardRule } from '../types';
import { Settings, RefreshCw, CheckCircle2, Sliders, Plus, Trash2 } from 'lucide-react';

interface StandardsEditorProps {
  standards: StandardRule[];
  onUpdateStandards: (newStandards: StandardRule[]) => void;
  onReset: () => void;
}

export default function StandardsEditor({ standards, onUpdateStandards, onReset }: StandardsEditorProps) {
  const [isOpen, setIsOpen] = useState(false);
  const [editingStandards, setEditingStandards] = useState<StandardRule[]>(standards);
  const [savedMessage, setSavedMessage] = useState(false);

  React.useEffect(() => {
    setEditingStandards(standards);
  }, [standards]);

  const handleChange = (id: string, field: keyof StandardRule, value: any) => {
    const updated = editingStandards.map(r => {
      if (r.id === id) {
        return { ...r, [field]: value };
      }
      return r;
    });
    setEditingStandards(updated);
  };

  const handleAddRule = () => {
    const newRule: StandardRule = {
      id: `custom_${Date.now()}`,
      name: `Custom Rule ${editingStandards.length + 1}`,
      fontName: editingStandards[editingStandards.length - 1]?.fontName || 'Arial',
      fontSize: editingStandards[editingStandards.length - 1]?.fontSize || 11,
      bold: null,
      alignment: null,
      indent: null
    };
    setEditingStandards([...editingStandards, newRule]);
  };

  const handleDeleteRule = (id: string) => {
    const updated = editingStandards.filter(r => r.id !== id);
    setEditingStandards(updated);
  };

  const handleSave = () => {
    onUpdateStandards(editingStandards);
    setSavedMessage(true);
    setTimeout(() => setSavedMessage(false), 3000);
  };

  return (
    <div id="standards_editor_card" class="bg-white rounded-xl border border-slate-200 shadow-sm overflow-hidden mb-8 transition-all duration-300">
      <button
        onClick={() => setIsOpen(!isOpen)}
        class="w-full px-6 py-4 flex items-center justify-between hover:bg-slate-50 border-b border-transparent transition"
        style={{ borderBottomColor: isOpen ? '#f1f5f9' : 'transparent' }}
      >
        <div class="flex items-center gap-3">
          <div class="p-2 bg-indigo-50 text-indigo-600 rounded-lg">
            <Settings id="settings_icon" class="w-5 h-5" />
          </div>
          <div class="text-left">
            <h2 class="text-base font-bold text-slate-800">Formatting Validation Standards</h2>
            <p class="text-xs text-slate-500">Configure target font face, size, and styling rules logic</p>
          </div>
        </div>
        <div class="flex items-center gap-2">
          <span class="text-xs text-slate-400 bg-slate-100 font-mono font-medium px-2 py-0.5 rounded-full">
            {standards.length} Active Rules
          </span>
          <span class="text-xs text-slate-500 font-medium">
            {isOpen ? 'Collapse configuration' : 'Tweak validation parameters'}
          </span>
        </div>
      </button>

      {isOpen && (
        <div class="p-6 bg-slate-50/50 animate-fade-in">
          <div class="overflow-x-auto rounded-lg border border-slate-200 bg-white shadow-sm mb-4">
            <table class="w-full border-collapse text-left text-sm text-slate-500">
              <thead class="bg-slate-50 text-xs font-semibold uppercase text-slate-600 border-b border-slate-200">
                <tr>
                  <th scope="col" class="px-6 py-3">Formatting Category</th>
                  <th scope="col" class="px-6 py-3">Expected Font Family</th>
                  <th scope="col" class="px-6 py-3">Size (pt)</th>
                  <th scope="col" class="px-6 py-3">Bold Setting</th>
                  <th scope="col" class="px-6 py-3">Expected Alignment</th>
                  <th scope="col" class="px-6 py-3">Left Indent (pt)</th>
                  <th scope="col" class="px-6 py-3 text-center">Actions</th>
                </tr>
              </thead>
              <tbody class="divide-y divide-slate-100 border-t border-slate-100 font-sans">
                {editingStandards.map((rule) => (
                  <tr key={rule.id} class="hover:bg-slate-50/70 transition">
                    <td class="px-6 py-2.5">
                      <input
                        type="text"
                        value={rule.name}
                        onChange={(e) => handleChange(rule.id, 'name', e.target.value)}
                        class="w-full px-2 py-1 text-xs font-semibold text-slate-800 bg-transparent border border-transparent rounded hover:border-slate-200 hover:bg-slate-50 focus:bg-white focus:border-indigo-500 outline-none transition"
                      />
                    </td>
                    <td class="px-6 py-2.5">
                      <input
                        type="text"
                        value={rule.fontName}
                        onChange={(e) => handleChange(rule.id, 'fontName', e.target.value)}
                        class="w-full px-3 py-1 text-xs border border-slate-200 rounded outline-none focus:border-indigo-500"
                      />
                    </td>
                    <td class="px-6 py-2.5 w-28">
                      <input
                        type="number"
                        step="0.5"
                        value={rule.fontSize}
                        onChange={(e) => handleChange(rule.id, 'fontSize', parseFloat(e.target.value) || 0)}
                        class="w-full px-3 py-1 text-xs border border-slate-200 rounded outline-none focus:border-indigo-500"
                      />
                    </td>
                    <td class="px-6 py-2.5 w-36">
                      <select
                        value={rule.bold === null ? 'any' : rule.bold ? 'true' : 'false'}
                        onChange={(e) => {
                          const v = e.target.value;
                          handleChange(rule.id, 'bold', v === 'any' ? null : v === 'true');
                        }}
                        class="w-full px-2 py-1 text-xs border border-slate-200 rounded bg-white outline-none focus:border-indigo-500"
                      >
                        <option value="true">Must be Bold</option>
                        <option value="false">Must be Regular</option>
                        <option value="any">Don't Enforce</option>
                      </select>
                    </td>
                    <td class="px-6 py-2.5 w-40">
                      <select
                        value={rule.alignment || 'ANY'}
                        onChange={(e) => {
                          const v = e.target.value;
                          handleChange(rule.id, 'alignment', v === 'ANY' ? null : v);
                        }}
                        class="w-full px-2 py-1 text-xs border border-slate-200 rounded bg-white outline-none focus:border-indigo-500 animate-none"
                      >
                        <option value="ANY">Don't Enforce</option>
                        <option value="LEFT">Left Aligned</option>
                        <option value="CENTER">Center Aligned</option>
                        <option value="RIGHT">Right Aligned</option>
                        <option value="JUSTIFY">Justify (Both)</option>
                      </select>
                    </td>
                    <td class="px-6 py-2.5 w-32">
                      <input
                        type="number"
                        placeholder="0"
                        value={rule.indent === null ? '' : rule.indent}
                        onChange={(e) => {
                          const v = e.target.value;
                          handleChange(rule.id, 'indent', v === '' ? null : parseFloat(v) || 0);
                        }}
                        class="w-full px-3 py-1 text-xs border border-slate-200 rounded outline-none focus:border-indigo-500"
                      />
                    </td>
                    <td class="px-6 py-2.5 text-center w-20">
                      <button
                        onClick={() => handleDeleteRule(rule.id)}
                        disabled={editingStandards.length <= 1}
                        class="text-slate-400 hover:text-rose-600 disabled:text-slate-200 disabled:pointer-events-none p-1 rounded hover:bg-rose-50 transition active:scale-95 cursor-pointer"
                        title="Delete rule"
                      >
                        <Trash2 class="w-4 h-4" />
                      </button>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>

          <div class="flex justify-end mb-4">
            <button
              onClick={handleAddRule}
              class="px-3 py-1.5 bg-indigo-50 hover:bg-indigo-150 text-indigo-600 text-xs font-semibold rounded-lg flex items-center gap-1.5 transition active:scale-95 shadow-sm cursor-pointer"
              title="Add a custom validation standard rule"
            >
              <Plus class="w-3.5 h-3.5" />
              Add Custom Rule
            </button>
          </div>

          <div class="flex items-center justify-between">
            <p class="text-xs text-slate-500 max-w-lg">
              <span class="font-semibold text-slate-600">Note:</span> The system classifies paragraphs matching their closest style properties and flags differences. Edits update the parsing logic instantly.
            </p>
            <div class="flex items-center gap-3">
              <button
                onClick={onReset}
                class="px-4 py-2 text-xs font-semibold text-slate-600 hover:text-slate-800 flex items-center gap-1.5 hover:bg-slate-100 rounded-lg transition"
              >
                <RefreshCw class="w-3.5 h-3.5" />
                Reset Defaults
              </button>
              <button
                onClick={handleSave}
                class="px-5 py-2 text-xs font-semibold text-white bg-indigo-600 hover:bg-indigo-700 rounded-lg flex items-center gap-1.5 shadow-sm hover:shadow transition"
              >
                <CheckCircle2 class="w-3.5 h-3.5" />
                Apply Tweak rules
              </button>
            </div>
          </div>

          {savedMessage && (
            <div class="mt-3 bg-emerald-50 text-emerald-800 text-xs px-4 py-2 rounded-lg border border-emerald-100 text-center animate-fade-in font-medium">
              Standards overridden successfully! Scanned documents will be re-analyzed instantly.
            </div>
          )}
        </div>
      )}
    </div>
  );
}
