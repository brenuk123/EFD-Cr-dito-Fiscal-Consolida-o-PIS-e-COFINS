import React, { useMemo, useState } from 'react';
import {
  EFD_TABS,
  ManualAuditState,
  ParsedWorkbook,
  TabConfigMap,
} from '../types/efd';
import {
  formatCurrency,
  getAllCandidateRowsForSheet,
} from '../utils/excelParser';
import { CheckCircle2, Search, X, XCircle } from 'lucide-react';

interface ManualRowModalProps {
  isOpen: boolean;
  onClose: () => void;
  workbook: ParsedWorkbook | null;
  initialTab?: string;
  tabConfig: TabConfigMap;
  blueColors: Set<string>;
  ignoreHidden: boolean;
  manualState: ManualAuditState;
  onToggleRowInclusion: (rowKey: string, hasBlueCell: boolean, currentlyIncluded: boolean) => void;
}

export const ManualRowModal: React.FC<ManualRowModalProps> = ({
  isOpen,
  onClose,
  workbook,
  initialTab,
  tabConfig,
  blueColors,
  ignoreHidden,
  manualState,
  onToggleRowInclusion,
}) => {
  const [selectedTab, setSelectedTab] = useState<string>(
    initialTab && EFD_TABS.includes(initialTab as any) ? initialTab : EFD_TABS[0]
  );
  const [searchQuery, setSearchQuery] = useState('');
  const [filterMode, setFilterMode] = useState<'all' | 'non_blue' | 'manual_changes'>('all');

  const candidateRows = useMemo(() => {
    return getAllCandidateRowsForSheet(
      workbook,
      selectedTab,
      tabConfig,
      blueColors,
      ignoreHidden,
      manualState
    );
  }, [workbook, selectedTab, tabConfig, blueColors, ignoreHidden, manualState]);

  const filteredCandidates = useMemo(() => {
    const q = searchQuery.trim().toLowerCase();
    return candidateRows.filter((row) => {
      if (filterMode === 'non_blue' && row.hasBlueCell) return false;
      if (
        filterMode === 'manual_changes' &&
        !row.isManualIncluded &&
        !row.isExcluded
      ) {
        return false;
      }
      if (q) {
        const hay = `${row.linha} ${row.previewText}`.toLowerCase();
        if (!hay.includes(q)) return false;
      }
      return true;
    });
  }, [candidateRows, searchQuery, filterMode]);

  if (!isOpen || !workbook) return null;

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center bg-slate-900/50 p-4">
      <div className="bg-white border border-slate-200 rounded-lg shadow-xl w-full max-w-5xl max-h-[88vh] flex flex-col">
        {/* Modal Header */}
        <div className="flex items-center justify-between px-6 py-4 border-b border-slate-200">
          <div>
            <h3 className="text-base font-bold text-slate-900">
              Inspecionar Todas as Linhas da Planilha (Inclusão / Exclusão Manual)
            </h3>
            <p className="text-xs text-slate-600 mt-0.5">
              Visualize todas as linhas de dados da aba original — inclusive linhas sem preenchimento azul — para incluir ou excluir registros manualmente na apuração.
            </p>
          </div>
          <button
            type="button"
            onClick={onClose}
            className="p-1.5 text-slate-500 hover:text-slate-900 rounded-md hover:bg-slate-100 transition-colors cursor-pointer"
          >
            <X className="w-5 h-5" />
          </button>
        </div>

        {/* Toolbar */}
        <div className="px-6 py-3 bg-slate-50 border-b border-slate-200 flex flex-wrap items-center justify-between gap-3">
          <div className="flex flex-wrap items-center gap-2">
            <label className="text-xs font-semibold text-slate-700">Aba:</label>
            <select
              value={selectedTab}
              onChange={(e) => setSelectedTab(e.target.value)}
              className="text-xs font-medium bg-white border border-slate-300 rounded-md px-3 py-1.5 text-slate-900 focus:outline-none focus:ring-2 focus:ring-blue-800"
            >
              {EFD_TABS.map((t) => (
                <option key={t} value={t}>
                  {t} ({workbook.sheets[t]?.rows.size || 0} linhas)
                </option>
              ))}
            </select>

            <div className="flex items-center gap-1 p-1 bg-slate-200/70 rounded-md ml-2">
              <button
                type="button"
                onClick={() => setFilterMode('all')}
                className={`px-2.5 py-1 text-xs font-medium rounded transition-colors cursor-pointer whitespace-nowrap shrink-0 ${
                  filterMode === 'all'
                    ? 'bg-white text-slate-900 shadow-2xs'
                    : 'text-slate-600 hover:text-slate-900'
                }`}
              >
                Todas ({candidateRows.length})
              </button>
              <button
                type="button"
                onClick={() => setFilterMode('non_blue')}
                className={`px-2.5 py-1 text-xs font-medium rounded transition-colors cursor-pointer whitespace-nowrap shrink-0 ${
                  filterMode === 'non_blue'
                    ? 'bg-white text-slate-900 shadow-2xs'
                    : 'text-slate-600 hover:text-slate-900'
                }`}
              >
                Sem cor azul ({candidateRows.filter((c) => !c.hasBlueCell).length})
              </button>
              <button
                type="button"
                onClick={() => setFilterMode('manual_changes')}
                className={`px-2.5 py-1 text-xs font-medium rounded transition-colors cursor-pointer whitespace-nowrap shrink-0 ${
                  filterMode === 'manual_changes'
                    ? 'bg-white text-slate-900 shadow-2xs'
                    : 'text-slate-600 hover:text-slate-900'
                }`}
              >
                Ajustadas manualmente (
                {
                  candidateRows.filter(
                    (c) => c.isManualIncluded || c.isExcluded
                  ).length
                }
                )
              </button>
            </div>
          </div>

          <div className="relative">
            <Search className="w-3.5 h-3.5 text-slate-400 absolute left-2.5 top-1/2 -translate-y-1/2" />
            <input
              type="text"
              value={searchQuery}
              onChange={(e) => setSearchQuery(e.target.value)}
              placeholder="Buscar documento, fornecedor, linha..."
              className="pl-8 pr-3 py-1.5 text-xs bg-white border border-slate-300 rounded-md w-64 focus:outline-none focus:ring-2 focus:ring-blue-800"
            />
          </div>
        </div>

        {/* Table Body */}
        <div className="flex-1 overflow-auto p-6">
          {filteredCandidates.length === 0 ? (
            <div className="py-12 text-center text-sm text-slate-500">
              Nenhuma linha encontrada com os filtros atuais nesta aba.
            </div>
          ) : (
            <div className="border border-slate-200 rounded-md overflow-hidden">
              <table className="w-full border-collapse text-left text-xs">
                <thead>
                  <tr className="bg-slate-100 border-b border-slate-200 text-slate-700 font-semibold">
                    <th className="py-2.5 px-3 whitespace-nowrap">Linha Excel</th>
                    <th className="py-2.5 px-3 whitespace-nowrap">Formatação Original</th>
                    <th className="py-2.5 px-3">Resumo dos Dados na Linha</th>
                    <th className="py-2.5 px-3 text-right whitespace-nowrap">Base (R$)</th>
                    <th className="py-2.5 px-3 text-right whitespace-nowrap">PIS (R$)</th>
                    <th className="py-2.5 px-3 text-right whitespace-nowrap">COFINS (R$)</th>
                    <th className="py-2.5 px-3 text-right whitespace-nowrap">Participação no Crédito</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-200">
                  {filteredCandidates.map((row) => (
                    <tr
                      key={row.key}
                      className={`hover:bg-slate-50 ${
                        row.isCurrentlyIncluded ? 'bg-white' : 'bg-slate-50/60 text-slate-500'
                      }`}
                    >
                      <td className="py-2 px-3 font-mono tabular-nums font-medium text-slate-900 whitespace-nowrap">
                        Linha {row.linha}
                      </td>
                      <td className="py-2 px-3 whitespace-nowrap">
                        {row.hasBlueCell ? (
                          <div className="inline-flex items-center gap-1.5 text-blue-900 font-medium">
                            <span
                              className="w-3 h-3 rounded-xs border border-slate-400"
                              style={{
                                backgroundColor: `#${row.matchedBlueColors[0]}`,
                              }}
                            />
                            <span>Linha Azul</span>
                          </div>
                        ) : (
                          <span className="text-slate-500">Sem cor azul</span>
                        )}
                      </td>
                      <td className="py-2 px-3 max-w-md truncate text-slate-700" title={row.previewText}>
                        {row.previewText || '—'}
                      </td>
                      <td className="py-2 px-3 text-right font-mono tabular-nums">
                        {formatCurrency(row.base)}
                      </td>
                      <td className="py-2 px-3 text-right font-mono tabular-nums">
                        {formatCurrency(row.pis)}
                      </td>
                      <td className="py-2 px-3 text-right font-mono tabular-nums">
                        {formatCurrency(row.cof)}
                      </td>
                      <td className="py-2 px-3 text-right whitespace-nowrap">
                        <button
                          type="button"
                          onClick={() =>
                            onToggleRowInclusion(
                              row.key,
                              row.hasBlueCell,
                              row.isCurrentlyIncluded
                            )
                          }
                          className={`inline-flex items-center gap-1.5 px-3 py-1 rounded text-xs font-medium transition-colors cursor-pointer whitespace-nowrap shrink-0 ${
                            row.isCurrentlyIncluded
                              ? 'bg-emerald-50 text-emerald-800 border border-emerald-200 hover:bg-red-50 hover:text-red-800 hover:border-red-200'
                              : 'bg-slate-100 text-slate-700 border border-slate-300 hover:bg-blue-50 hover:text-blue-900 hover:border-blue-300'
                          }`}
                        >
                          {row.isCurrentlyIncluded ? (
                            <>
                              <CheckCircle2 className="w-3.5 h-3.5" />
                              {row.isManualIncluded
                                ? 'Incluída Manualmente (Remover)'
                                : 'Incluída (Excluir)'}
                            </>
                          ) : (
                            <>
                              <XCircle className="w-3.5 h-3.5" />
                              {row.isExcluded
                                ? 'Excluída Manualmente (Restaurar)'
                                : 'Fora da Seleção (Incluir)'}
                            </>
                          )}
                        </button>
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          )}
        </div>

        {/* Modal Footer */}
        <div className="px-6 py-3.5 border-t border-slate-200 bg-slate-50 flex items-center justify-between">
          <div className="text-xs text-slate-600">
            Total na aba <span className="font-semibold text-slate-800">{selectedTab}</span>:{' '}
            <span className="font-mono tabular-nums font-semibold text-slate-900">
              {candidateRows.filter((c) => c.isCurrentlyIncluded).length}
            </span>{' '}
            linhas ativas na consolidação ·{' '}
            <span className="font-mono tabular-nums">
              {candidateRows.filter((c) => c.isManualIncluded).length}
            </span>{' '}
            incluídas manualmente ·{' '}
            <span className="font-mono tabular-nums">
              {candidateRows.filter((c) => c.isExcluded).length}
            </span>{' '}
            excluídas manualmente
          </div>
          <button
            type="button"
            onClick={onClose}
            className="px-4 py-2 text-xs font-semibold text-white bg-slate-900 rounded-md hover:bg-slate-800 transition-colors cursor-pointer whitespace-nowrap shrink-0"
          >
            Concluir Revisão de Linhas
          </button>
        </div>
      </div>
    </div>
  );
};
