import React, { useState } from 'react';
import { ClosingHistoryEntry } from '../types/efd';
import { formatCurrency } from '../utils/excelParser';
import { Calendar, ChevronDown, ChevronUp, Trash2 } from 'lucide-react';

interface HistorySectionProps {
  history: ClosingHistoryEntry[];
  onDeleteEntry: (id: string) => void;
  onClearHistory: () => void;
}

export const HistorySection: React.FC<HistorySectionProps> = ({
  history,
  onDeleteEntry,
  onClearHistory,
}) => {
  const [expandedId, setExpandedId] = useState<string | null>(
    history[0]?.id || null
  );

  return (
    <section className="bg-white border border-slate-200 rounded-lg p-6 space-y-5">
      <div className="flex flex-wrap items-center justify-between gap-4 pb-4 border-b border-slate-200">
        <div>
          <h2 className="text-lg font-bold text-slate-900 tracking-tight">
            04. Histórico Local de Fechamentos Mensais
          </h2>
          <p className="text-sm text-slate-600 mt-0.5">
            Consulte e compare as apurações de créditos de PIS e COFINS salvas neste navegador entre competências.
          </p>
        </div>

        {history.length > 0 && (
          <button
            type="button"
            onClick={onClearHistory}
            className="inline-flex items-center gap-1.5 px-3 py-1.5 text-xs font-medium text-red-700 bg-white border border-slate-200 rounded-md hover:bg-red-50 transition-colors cursor-pointer whitespace-nowrap shrink-0"
          >
            <Trash2 className="w-3.5 h-3.5" />
            Limpar Histórico
          </button>
        )}
      </div>

      {history.length === 0 ? (
        <div className="p-8 bg-slate-50 border border-slate-200 rounded-md text-center space-y-2">
          <p className="text-sm font-medium text-slate-700">
            Nenhum fechamento mensal salvo até o momento.
          </p>
          <p className="text-xs text-slate-500">
            Após consolidar um arquivo Excel, clique em "Salvar no Histórico Mensal" no painel de conferência para registrar a competência.
          </p>
        </div>
      ) : (
        <div className="space-y-3">
          {history.map((entry) => {
            const isExpanded = expandedId === entry.id;
            const formattedDate = new Date(entry.savedAt).toLocaleString(
              'pt-BR',
              {
                dateStyle: 'short',
                timeStyle: 'short',
              }
            );
            return (
              <div
                key={entry.id}
                className="border border-slate-200 rounded-md overflow-hidden"
              >
                <div className="px-4 py-3 bg-slate-50 flex flex-wrap items-center justify-between gap-4">
                  <div className="flex flex-wrap items-center gap-3 text-xs">
                    <span className="inline-flex items-center gap-1.5 font-bold text-slate-900 text-sm font-mono tabular-nums">
                      <Calendar className="w-4 h-4 text-blue-900" />
                      {entry.competencia}
                    </span>
                    <span aria-hidden="true" className="text-slate-300">
                      ·
                    </span>
                    <span className="font-medium text-slate-700">
                      {entry.fileName}
                    </span>
                    <span aria-hidden="true" className="text-slate-300">
                      ·
                    </span>
                    <span className="text-slate-500">Salvo em {formattedDate}</span>
                    {entry.notes && (
                      <>
                        <span aria-hidden="true" className="text-slate-300">
                          ·
                        </span>
                        <span className="text-slate-600 italic">
                          "{entry.notes}"
                        </span>
                      </>
                    )}
                  </div>

                  <div className="flex items-center gap-4">
                    <div className="text-xs font-mono tabular-nums">
                      <span className="text-slate-500">Crédito Total: </span>
                      <span className="font-bold text-slate-900">
                        R$ {formatCurrency(entry.totalCredit)}
                      </span>
                    </div>

                    <button
                      type="button"
                      onClick={() =>
                        setExpandedId(isExpanded ? null : entry.id)
                      }
                      className="inline-flex items-center gap-1 px-2.5 py-1 text-xs font-medium text-slate-700 bg-white border border-slate-300 rounded hover:bg-slate-100 transition-colors cursor-pointer"
                    >
                      {isExpanded ? (
                        <>
                          Ocultar detalhes
                          <ChevronUp className="w-3.5 h-3.5" />
                        </>
                      ) : (
                        <>
                          Ver resumo por aba
                          <ChevronDown className="w-3.5 h-3.5" />
                        </>
                      )}
                    </button>

                    <button
                      type="button"
                      onClick={() => onDeleteEntry(entry.id)}
                      className="p-1.5 text-slate-400 hover:text-red-700 rounded hover:bg-red-50 transition-colors cursor-pointer"
                      title="Remover este fechamento do histórico"
                    >
                      <Trash2 className="w-4 h-4" />
                    </button>
                  </div>
                </div>

                {isExpanded && (
                  <div className="p-4 bg-white border-t border-slate-200 space-y-4">
                    <div className="grid grid-cols-2 sm:grid-cols-6 gap-3 text-xs">
                      <div className="p-2.5 bg-slate-50 border border-slate-200 rounded">
                        <div className="text-slate-500">Registros</div>
                        <div className="font-mono tabular-nums font-bold text-slate-900 text-sm mt-0.5">
                          {entry.totalRecords}
                        </div>
                      </div>
                      <div className="p-2.5 bg-slate-50 border border-slate-200 rounded">
                        <div className="text-slate-500">Base Cálculo (R$)</div>
                        <div className="font-mono tabular-nums font-bold text-slate-900 text-sm mt-0.5">
                          {formatCurrency(entry.totalBase)}
                        </div>
                      </div>
                      <div className="p-2.5 bg-slate-50 border border-slate-200 rounded">
                        <div className="text-slate-500">PIS (R$)</div>
                        <div className="font-mono tabular-nums font-bold text-slate-900 text-sm mt-0.5">
                          {formatCurrency(entry.totalPis)}
                        </div>
                      </div>
                      <div className="p-2.5 bg-slate-50 border border-slate-200 rounded">
                        <div className="text-slate-500">COFINS (R$)</div>
                        <div className="font-mono tabular-nums font-bold text-slate-900 text-sm mt-0.5">
                          {formatCurrency(entry.totalCofins)}
                        </div>
                      </div>
                      <div className="p-2.5 bg-slate-50 border border-slate-200 rounded">
                        <div className="text-slate-500">Divergências Pend.</div>
                        <div className="font-mono tabular-nums font-bold text-slate-900 text-sm mt-0.5">
                          {entry.pendingDivergences} (
                          {entry.justifiedDivergences} just.)
                        </div>
                      </div>
                      <div className="p-2.5 bg-slate-50 border border-slate-200 rounded">
                        <div className="text-slate-500">Duplicados</div>
                        <div className="font-mono tabular-nums font-bold text-slate-900 text-sm mt-0.5">
                          {entry.duplicateCount}
                        </div>
                      </div>
                    </div>

                    <div className="overflow-x-auto border border-slate-200 rounded">
                      <table className="w-full border-collapse text-left text-xs">
                        <thead>
                          <tr className="bg-slate-100 border-b border-slate-200 text-slate-700 font-semibold">
                            <th className="py-2 px-3">Categoria / Aba</th>
                            <th className="py-2 px-2.5 text-right">Registros</th>
                            <th className="py-2 px-2.5 text-right">Base (R$)</th>
                            <th className="py-2 px-2.5 text-right">PIS (R$)</th>
                            <th className="py-2 px-2.5 text-right">COFINS (R$)</th>
                            <th className="py-2 px-2.5 text-right">Total (R$)</th>
                          </tr>
                        </thead>
                        <tbody className="divide-y divide-slate-200">
                          {entry.byTab.map((t) => (
                            <tr key={t.aba} className="hover:bg-slate-50">
                              <td className="py-1.5 px-3 font-medium text-slate-900">
                                {t.aba}
                              </td>
                              <td className="py-1.5 px-2.5 text-right font-mono tabular-nums">
                                {t.n}
                              </td>
                              <td className="py-1.5 px-2.5 text-right font-mono tabular-nums">
                                {formatCurrency(t.base)}
                              </td>
                              <td className="py-1.5 px-2.5 text-right font-mono tabular-nums">
                                {formatCurrency(t.pis)}
                              </td>
                              <td className="py-1.5 px-2.5 text-right font-mono tabular-nums">
                                {formatCurrency(t.cof)}
                              </td>
                              <td className="py-1.5 px-2.5 text-right font-mono tabular-nums font-semibold">
                                {formatCurrency(t.tot)}
                              </td>
                            </tr>
                          ))}
                        </tbody>
                      </table>
                    </div>
                  </div>
                )}
              </div>
            );
          })}
        </div>
      )}
    </section>
  );
};
