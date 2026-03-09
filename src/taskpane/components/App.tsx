/* global Excel */
import React, { useState, useMemo } from 'react';

// --- INLINE SVG ICONS (Replaces lucide-react to prevent version mismatch crashes) ---
const IconActivity = ({ size = 24, className = "" }) => (<svg width={size} height={size} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" className={className}><polyline points="22 12 18 12 15 21 9 3 6 12 2 12"></polyline></svg>);
const IconFileSpreadsheet = ({ size = 24, className = "" }) => (<svg width={size} height={size} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" className={className}><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"></path><polyline points="14 2 14 8 20 8"></polyline><line x1="8" y1="13" x2="16" y2="13"></line><line x1="8" y1="17" x2="16" y2="17"></line><polyline points="10 9 9 9 8 9"></polyline></svg>);
const IconSearch = ({ size = 24, className = "" }) => (<svg width={size} height={size} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" className={className}><circle cx="11" cy="11" r="8"></circle><line x1="21" y1="21" x2="16.65" y2="16.65"></line></svg>);
const IconMaximize2 = ({ size = 24, className = "" }) => (<svg width={size} height={size} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" className={className}><polyline points="15 3 21 3 21 9"></polyline><polyline points="9 21 3 21 3 15"></polyline><line x1="21" y1="3" x2="14" y2="10"></line><line x1="3" y1="21" x2="10" y2="14"></line></svg>);
const IconMinimize2 = ({ size = 24, className = "" }) => (<svg width={size} height={size} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" className={className}><polyline points="4 14 10 14 10 20"></polyline><polyline points="20 10 14 10 14 4"></polyline><line x1="14" y1="10" x2="21" y2="3"></line><line x1="3" y1="21" x2="10" y2="14"></line></svg>);
const IconChevronDown = ({ size = 24, className = "" }) => (<svg width={size} height={size} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" className={className}><polyline points="6 9 12 15 18 9"></polyline></svg>);
const IconChevronRight = ({ size = 24, className = "" }) => (<svg width={size} height={size} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" className={className}><polyline points="9 18 15 12 9 6"></polyline></svg>);
const IconCrosshair = ({ size = 24, className = "" }) => (<svg width={size} height={size} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" className={className}><circle cx="12" cy="12" r="10"></circle><line x1="22" y1="12" x2="18" y2="12"></line><line x1="6" y1="12" x2="2" y2="12"></line><line x1="12" y1="6" x2="12" y2="2"></line><line x1="12" y1="22" x2="12" y2="18"></line></svg>);

// --- HELPER: Convert Column Index (0) to Letter (A) ---
const getColLetter = (colIndex) => {
  let temp, letter = '';
  let col = colIndex + 1;
  while (col > 0) {
    temp = (col - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    col = (col - temp - 1) / 26;
  }
  return letter;
};

// --- HELPER: Regex Engine to identify Hardcoded Values ---
const checkHardcoded = (formula) => {
  if (/showrow|showcolumn|hiderow|hidecolumn/i.test(formula)) return false;
  let cleaned = formula.replace(/"[^"]*"/g, ""); // Remove Strings
  cleaned = cleaned.replace(/'?([^'!]+)'?!/g, ""); // Remove Sheet Refs
  cleaned = cleaned.replace(/\$?[A-Z]{1,3}\$?[1-9][0-9]{0,6}(:\$?[A-Z]{1,3}\$?[1-9][0-9]{0,6})?/gi, ""); // Remove Cell Refs
  cleaned = cleaned.replace(/\b([A-Z][A-Z0-9_\.]*)\s*\(/gi, ""); // Remove Functions
  return /\b\d+(\.\d+)?\b/.test(cleaned); // Check if any rogue numbers are left over
};

// --- TYPESCRIPT INTERFACE ---
export interface AppProps {
  title?: string;
  isOfficeInitialized?: boolean;
}

export default function App(props: AppProps) {
  const [isScanning, setIsScanning] = useState(false);
  const [data, setData] = useState([]);
  const [searchQuery, setSearchQuery] = useState('');
  const [expandedGroups, setExpandedGroups] = useState(new Set());

  // --- THE EXCEL API SCANNER ---
  const handleScan = async () => {
    setIsScanning(true);
    setData([]);
    setExpandedGroups(new Set());

    try {
      await Excel.run(async (context) => {
        const sheets = context.workbook.worksheets;
        sheets.load("items/name");
        await context.sync();

        const groups = {};

        // Loop all sheets
        for (let sheet of sheets.items) {
          const usedRange = sheet.getUsedRangeOrNullObject();
          await context.sync();
          if (usedRange.isNullObject) continue;

          // Find ALL formulas instantly
          const formulaAreas = usedRange.getSpecialCellsOrNullObject(Excel.SpecialCellType.formulas);
          await context.sync();
          if (formulaAreas.isNullObject) continue;

          formulaAreas.areas.load("items");
          await context.sync();

          // Load values for all grouped formula areas
          for (let area of formulaAreas.areas.items) {
            area.load(["formulas", "formulasR1C1", "rowIndex", "columnIndex", "rowCount", "columnCount"]);
          }
          await context.sync();

          // Process the 2D arrays
          for (let area of formulaAreas.areas.items) {
            const r1c1s = area.formulasR1C1;
            const a1s = area.formulas;
            const rOffset = area.rowIndex;
            const cOffset = area.columnIndex;

            for (let r = 0; r < area.rowCount; r++) {
              for (let c = 0; c < area.columnCount; c++) {
                const r1c1 = r1c1s[r][c];
                if (!r1c1) continue;

                const a1 = a1s[r][c];
                const groupKey = `${sheet.name}|${r1c1}`;
                const address = `${getColLetter(cOffset + c)}${rOffset + r + 1}`;

                if (!groups[groupKey]) {
                  groups[groupKey] = {
                    id: groupKey,
                    sheet: sheet.name,
                    count: 1,
                    r1c1: r1c1,
                    a1: a1,
                    hasHardcoded: checkHardcoded(a1),
                    isSpilled: false,
                    addresses: [address]
                  };
                } else {
                  groups[groupKey].count++;
                  // Limit memory: only store up to 50 addresses per group
                  if (groups[groupKey].addresses.length < 50) {
                    groups[groupKey].addresses.push(address);
                  } else if (groups[groupKey].addresses.length === 50) {
                    groups[groupKey].addresses.push(`...and more`);
                  }
                }
              }
            }
          }
        }

        const results = Object.values(groups);
        setData(results);
      });
    } catch (error) {
      console.error("Scan Failed: ", error);
      alert("An error occurred while scanning the workbook.");
    } finally {
      setIsScanning(false);
    }
  };

  // --- THE EXCEL API NAVIGATOR ---
  const handleNavigateToCell = async (sheetName, address) => {
    if (address.includes('more')) return;
    try {
      await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getItem(sheetName);
        const range = sheet.getRange(address);
        range.select();
        await context.sync();
      });
    } catch (error) {
      console.error(error);
    }
  };

  const toggleGroup = (id) => {
    const newExpanded = new Set(expandedGroups);
    if (newExpanded.has(id)) newExpanded.delete(id); else newExpanded.add(id);
    setExpandedGroups(newExpanded);
  };

  const expandAll = () => setExpandedGroups(new Set(data.map(d => d.id)));
  const collapseAll = () => setExpandedGroups(new Set());

  const filteredData = useMemo(() => {
    let filtered = data;
    if (searchQuery.trim() !== '') {
      const q = searchQuery.toLowerCase();
      filtered = data.filter(d => d.a1.toLowerCase().includes(q) || d.sheet.toLowerCase().includes(q));
    }
    return filtered.sort((a, b) => b.count - a.count);
  }, [data, searchQuery]);

  return (
    <div className="flex flex-col h-screen bg-slate-50 font-sans text-slate-800">

      {/* Excel-Style Header */}
      <header className="bg-[#107C41] text-white p-4 shadow-md shrink-0 flex items-center gap-3">
        <IconActivity size={24} className="opacity-90" />
        <div>
          <h1 className="text-lg font-semibold leading-tight">{props.title || 'Auxiliary Pro'}</h1>
          <p className="text-xs text-green-100 opacity-90">Advanced Formula Auditor</p>
        </div>
      </header>

      {/* Action Bar */}
      <div className="p-4 bg-white border-b border-slate-200 shrink-0 space-y-4 shadow-sm z-10">
        <div className="flex gap-3">
          <button
            onClick={handleScan}
            disabled={isScanning}
            className="flex-1 bg-[#107C41] hover:bg-[#0b5c30] text-white font-medium py-2 px-4 rounded transition-colors disabled:opacity-70 disabled:cursor-not-allowed flex items-center justify-center gap-2 shadow-sm"
          >
            {isScanning ? (
              <span className="animate-pulse">Scanning Workbook...</span>
            ) : (
              <>
                <IconFileSpreadsheet size={18} />
                Scan Workbook
              </>
            )}
          </button>
        </div>

        <div className="flex gap-2 items-center">
          <div className="relative flex-1">
            <IconSearch size={16} className="absolute left-3 top-1/2 -translate-y-1/2 text-slate-400" />
            <input
              type="text"
              placeholder="Filter by formula or sheet..."
              value={searchQuery}
              onChange={(e) => setSearchQuery(e.target.value)}
              className="w-full pl-9 pr-3 py-2 bg-slate-50 border border-slate-200 rounded text-sm focus:outline-none focus:ring-2 focus:ring-[#107C41] focus:border-transparent transition-all"
            />
          </div>
          <button onClick={expandAll} className="p-2 text-slate-500 hover:bg-slate-100 rounded border border-slate-200" title="Expand All">
            <IconMaximize2 size={16} />
          </button>
          <button onClick={collapseAll} className="p-2 text-slate-500 hover:bg-slate-100 rounded border border-slate-200" title="Collapse All">
            <IconMinimize2 size={16} />
          </button>
        </div>

        {data.length > 0 && (
          <div className="text-xs text-slate-500 font-medium flex justify-between">
            <span>Showing {filteredData.length} unique formula patterns</span>
          </div>
        )}
      </div>

      {/* List View Area */}
      <div className="flex-1 overflow-y-auto p-2 bg-slate-50">
        {data.length === 0 && !isScanning && (
          <div className="h-full flex flex-col items-center justify-center text-slate-400 p-8 text-center gap-3">
            <IconFileSpreadsheet size={48} className="opacity-20" />
            <p>Click <strong>Scan Workbook</strong> to analyze all formulas and identify hardcoded values.</p>
          </div>
        )}

        <div className="space-y-1 pb-4">
          {filteredData.map((group) => {
            const isExpanded = expandedGroups.has(group.id);
            return (
              <div key={group.id} className="bg-white rounded border border-slate-200 overflow-hidden shadow-sm transition-all hover:border-[#107C41]">
                {/* Parent Row */}
                <div
                  className="flex items-center p-2.5 cursor-pointer hover:bg-green-50/50 transition-colors"
                  onClick={() => toggleGroup(group.id)}
                >
                  <div className="text-slate-400 mr-2 shrink-0">
                    {isExpanded ? <IconChevronDown size={18} /> : <IconChevronRight size={18} />}
                  </div>

                  <div className="flex-1 min-w-0 flex items-center gap-3">
                    <span className="text-xs font-bold text-slate-500 bg-slate-100 px-2 py-0.5 rounded shrink-0">
                      {group.count}x
                    </span>
                    <span className="text-xs font-semibold text-slate-700 w-24 truncate shrink-0" title={group.sheet}>
                      {group.sheet}
                    </span>
                    <div className="flex gap-1 shrink-0">
                      {group.hasHardcoded && (
                        <span className="text-[10px] font-bold text-red-700 bg-red-100 px-1.5 py-0.5 rounded border border-red-200">
                          HARDCODED
                        </span>
                      )}
                    </div>
                    <span className="font-mono text-sm text-slate-700 truncate w-full" title={group.a1}>
                      {group.a1}
                    </span>
                  </div>
                </div>

                {/* Expanded Children */}
                {isExpanded && (
                  <div className="bg-slate-50 border-t border-slate-100 py-1">
                    {group.addresses.map((address, idx) => (
                      <div
                        key={idx}
                        onClick={() => handleNavigateToCell(group.sheet, address)}
                        className={`
                          flex items-center gap-2 py-1.5 px-10 text-xs font-mono cursor-pointer transition-colors
                          ${address.includes('more') ? 'text-slate-400 italic cursor-default' : 'text-slate-600 hover:text-[#107C41] hover:bg-green-100/50'}
                        `}
                      >
                        {!address.includes('more') && <IconCrosshair size={12} className="opacity-50" />}
                        {address}
                      </div>
                    ))}
                  </div>
                )}
              </div>
            );
          })}
        </div>
      </div>
    </div>
  );
}