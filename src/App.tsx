import React, { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import {
  ClosingHistoryEntry,
  EFD_TABS,
  ManualAuditState,
  ParsedWorkbook,
  SavedProfile,
  TabConfigMap,
  createDefaultTabConfig,
} from './types/efd';
import {
  buildConsolidation,
  computeResumoBlockBreakdown,
  computeResumoTotalComparisonRows,
  computeTabAggregations,
  parseExcelWorkbook,
  parseNum,
} from './utils/excelParser';
import {
  AccountingMapEntry,
  DEFAULT_CTA_DEBITO_COFINS,
  DEFAULT_CTA_DEBITO_PIS,
  buildAccountingDynamics,
  inferCompetenciaFromFileName,
  loadAccountingMap,
  saveAccountingMapToStorage,
} from './utils/accountingDynamics';
import { generateDemoEfdWorkbookBuffer } from './utils/demoWorkbookGenerator';
import { exportConsolidatedExcel } from './utils/excelExporter';
import {
  loadClosingHistory,
  loadSavedProfiles,
  saveClosingHistoryToStorage,
  saveProfilesToStorage,
} from './utils/storage';
import { CriteriaSection } from './components/CriteriaSection';
import { ConferenceSection } from './components/ConferenceSection';
import { AccountingDynamicsSection } from './components/AccountingDynamicsSection';
import { HistorySection } from './components/HistorySection';
import { ManualRowModal } from './components/ManualRowModal';
import {
  CheckCircle2,
  Download,
  FileSpreadsheet,
  FolderOpen,
  Play,
  XCircle,
} from 'lucide-react';

export default function App() {
  const fileInputRef = useRef<HTMLInputElement | null>(null);

  // Core workbook & loading states
  const [workbook, setWorkbook] = useState<ParsedWorkbook | null>(null);
  const [isLoadingFile, setIsLoadingFile] = useState(false);
  const [loadError, setLoadError] = useState<string | null>(null);
  const [isExporting, setIsExporting] = useState(false);

  // Saved profiles & history
  const [profiles, setProfiles] = useState<SavedProfile[]>(() =>
    loadSavedProfiles()
  );
  const [activeProfileId, setActiveProfileId] = useState<string>(
    'default-efd-8-2026'
  );
  const [history, setHistory] = useState<ClosingHistoryEntry[]>(() =>
    loadClosingHistory()
  );

  // Accounting Dynamics state
  const [accountingMap, setAccountingMap] = useState<AccountingMapEntry[]>(() =>
    loadAccountingMap()
  );
  const [competencia, setCompetencia] = useState<string>('08/2026');
  const [ctaDebitoPis, setCtaDebitoPis] = useState<string>(
    DEFAULT_CTA_DEBITO_PIS
  );
  const [ctaDebitoCofins, setCtaDebitoCofins] = useState<string>(
    DEFAULT_CTA_DEBITO_COFINS
  );

  // Criteria & section configurations
  const [blueColors, setBlueColors] = useState<Set<string>>(
    () => new Set(['BDD7EE', '8EA9DB', '9BC2E6', '4472C4', '5B9BD5', 'D9E1F2'])
  );
  const [toleranceInput, setToleranceInput] = useState<string>('0,02');
  const [ignoreHidden, setIgnoreHidden] = useState<boolean>(false);
  const [tabConfig, setTabConfig] = useState<TabConfigMap>(() =>
    createDefaultTabConfig()
  );

  // Manual audit adjustments (manual inclusions, exclusions, justifications)
  const [manualState, setManualState] = useState<ManualAuditState>({
    includedRowKeys: [],
    excludedRowKeys: [],
    justifications: {},
  });

  // Manual row inspection modal state
  const [manualModalOpen, setManualModalOpen] = useState(false);
  const [manualModalTab, setManualModalTab] = useState<string | undefined>(
    undefined
  );

  const tolerance = useMemo(
    () => parseNum(toleranceInput) || 0.02,
    [toleranceInput]
  );

  const handleLoadBufferAsWorkbook = useCallback(
    async (fileName: string, buffer: ArrayBuffer) => {
      setIsLoadingFile(true);
      setLoadError(null);
      try {
        const parsed = await parseExcelWorkbook({
          name: fileName,
          size: buffer.byteLength,
          arrayBuffer: async () => buffer,
        });
        setWorkbook(parsed);
        setCompetencia(inferCompetenciaFromFileName(fileName));
        setBlueColors(new Set(parsed.defaultBlueColors));
        setManualState({
          includedRowKeys: [],
          excludedRowKeys: [],
          justifications: {},
        });
      } catch (err: any) {
        setLoadError(err?.message || 'Erro ao processar a planilha Excel.');
      } finally {
        setIsLoadingFile(false);
      }
    },
    []
  );

  useEffect(() => {
    let mounted = true;
    generateDemoEfdWorkbookBuffer()
      .then((buf) => {
        if (mounted) {
          handleLoadBufferAsWorkbook('EFD 8_2026.xlsx', buf);
        }
      })
      .catch((err) => {
        console.error('Erro ao inicializar planilha exemplo:', err);
      });
    return () => {
      mounted = false;
    };
  }, [handleLoadBufferAsWorkbook]);

  const handleFileUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
    setIsLoadingFile(true);
    setLoadError(null);
    try {
      const parsed = await parseExcelWorkbook(file);
      setWorkbook(parsed);
      setCompetencia(inferCompetenciaFromFileName(file.name));
      setBlueColors(new Set(parsed.defaultBlueColors));
      setManualState({
        includedRowKeys: [],
        excludedRowKeys: [],
        justifications: {},
      });
    } catch (err: any) {
      setLoadError(err?.message || 'Erro ao ler o arquivo Excel selecionado.');
    } finally {
      setIsLoadingFile(false);
    }
  };

  const handleLoadDemoWorkbook = async () => {
    const buf = await generateDemoEfdWorkbookBuffer();
    await handleLoadBufferAsWorkbook('EFD 8_2026.xlsx', buf);
  };

  const handleDownloadDemoWorkbook = async () => {
    const buf = await generateDemoEfdWorkbookBuffer();
    const blob = new Blob([buf], {
      type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = 'EFD 8_2026_Exemplo.xlsx';
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  };

  // Consolidation computation
  const { records: allRecords, usedColumns } = useMemo(() => {
    return buildConsolidation(
      workbook,
      tabConfig,
      blueColors,
      tolerance,
      ignoreHidden,
      manualState
    );
  }, [workbook, tabConfig, blueColors, tolerance, ignoreHidden, manualState]);

  const tabAggregations = useMemo(
    () => computeTabAggregations(allRecords),
    [allRecords]
  );

  const resumoTotalRows = useMemo(
    () => computeResumoTotalComparisonRows(workbook, tabAggregations, tolerance),
    [workbook, tabAggregations, tolerance]
  );

  const resumoBlocks = useMemo(
    () => computeResumoBlockBreakdown(workbook, tabConfig, allRecords),
    [workbook, tabConfig, allRecords]
  );

  const accountingBlocks = useMemo(
    () =>
      buildAccountingDynamics({
        allRecords,
        accountingMap,
        competencia,
        ctaDebitoPis,
        ctaDebitoCofins,
      }),
    [allRecords, accountingMap, competencia, ctaDebitoPis, ctaDebitoCofins]
  );

  const handleChangeAccountingMap = (next: AccountingMapEntry[]) => {
    setAccountingMap(next);
    saveAccountingMapToStorage(next);
  };

  // Handlers for color & profile management
  const handleToggleBlueColor = (hex: string) => {
    setBlueColors((prev) => {
      const next = new Set(prev);
      if (next.has(hex)) next.delete(hex);
      else next.add(hex);
      return next;
    });
  };

  const handleResetAutoBlueColors = () => {
    if (!workbook) return;
    setBlueColors(new Set(workbook.defaultBlueColors));
  };

  const handleResetDefaultConfig = () => {
    setTabConfig(createDefaultTabConfig());
    setToleranceInput('0,02');
    setIgnoreHidden(false);
    setActiveProfileId('default-efd-8-2026');
    if (workbook) {
      setBlueColors(new Set(workbook.defaultBlueColors));
    }
  };

  const handleSelectProfile = (profile: SavedProfile) => {
    setActiveProfileId(profile.id);
    setToleranceInput(
      profile.tolerance.toLocaleString('pt-BR', { minimumFractionDigits: 2 })
    );
    setIgnoreHidden(profile.ignoreHidden);
    setTabConfig(JSON.parse(JSON.stringify(profile.tabConfig)));
    if (profile.blueColors?.length) {
      setBlueColors(new Set(profile.blueColors));
    }
  };

  const handleSaveProfile = (name: string) => {
    const newProfile: SavedProfile = {
      id: `profile-${Date.now()}`,
      name,
      updatedAt: new Date().toISOString(),
      tolerance,
      ignoreHidden,
      blueColors: [...blueColors],
      tabConfig: JSON.parse(JSON.stringify(tabConfig)),
    };
    const nextProfiles = [...profiles, newProfile];
    setProfiles(nextProfiles);
    saveProfilesToStorage(nextProfiles);
    setActiveProfileId(newProfile.id);
  };

  const handleDeleteProfile = (id: string) => {
    if (id === 'default-efd-8-2026') return;
    const nextProfiles = profiles.filter((p) => p.id !== id);
    setProfiles(nextProfiles);
    saveProfilesToStorage(nextProfiles);
    setActiveProfileId('default-efd-8-2026');
  };

  // Handlers for manual row inclusion/exclusion and divergence justification
  const handleToggleRowInclusion = (
    rowKey: string,
    hasBlueCell: boolean,
    currentlyIncluded: boolean
  ) => {
    setManualState((prev) => {
      const inc = new Set(prev.includedRowKeys);
      const exc = new Set(prev.excludedRowKeys);

      if (currentlyIncluded) {
        inc.delete(rowKey);
        if (hasBlueCell) {
          exc.add(rowKey);
        }
      } else {
        exc.delete(rowKey);
        if (!hasBlueCell) {
          inc.add(rowKey);
        }
      }

      return {
        ...prev,
        includedRowKeys: [...inc],
        excludedRowKeys: [...exc],
      };
    });
  };

  const handleSaveJustification = (rowKey: string, note: string) => {
    setManualState((prev) => ({
      ...prev,
      justifications: {
        ...prev.justifications,
        [rowKey]: {
          note,
          updatedAt: new Date().toISOString(),
        },
      },
    }));
  };

  const handleClearJustification = (rowKey: string) => {
    setManualState((prev) => {
      const nextJust = { ...prev.justifications };
      delete nextJust[rowKey];
      return {
        ...prev,
        justifications: nextJust,
      };
    });
  };

  // Export handler
  const handleExportExcel = async () => {
    if (!workbook) return;
    setIsExporting(true);
    try {
      await exportConsolidatedExcel({
        sourceFileName: workbook.fileName,
        allRecords,
        usedColumns,
        tabAggregations,
        resumoTotalRows,
        resumoBlocks,
        accountingBlocks,
        accountingMap,
      });
    } catch (err) {
      console.error('Erro ao exportar planilha:', err);
    } finally {
      setIsExporting(false);
    }
  };

  // History handlers
  const handleSaveToHistory = (comp: string, notes: string) => {
    if (!workbook) return;
    const totals = tabAggregations.reduce(
      (s, a) => ({
        n: s.n + a.n,
        b: s.b + a.base,
        p: s.p + a.pis,
        c: s.c + a.cof,
        v: s.v + a.div,
        j: s.j + a.justified,
        d: s.d + a.dup,
      }),
      { n: 0, b: 0, p: 0, c: 0, v: 0, j: 0, d: 0 }
    );

    const newEntry: ClosingHistoryEntry = {
      id: `hist-${Date.now()}`,
      competencia: comp,
      fileName: workbook.fileName,
      savedAt: new Date().toISOString(),
      totalRecords: totals.n,
      totalBase: totals.b,
      totalPis: totals.p,
      totalCofins: totals.c,
      totalCredit: totals.p + totals.c,
      pendingDivergences: totals.v,
      justifiedDivergences: totals.j,
      duplicateCount: totals.d,
      byTab: tabAggregations,
      notes,
    };

    const nextHistory = [newEntry, ...history];
    setHistory(nextHistory);
    saveClosingHistoryToStorage(nextHistory);
  };

  const handleDeleteHistoryEntry = (id: string) => {
    const nextHistory = history.filter((h) => h.id !== id);
    setHistory(nextHistory);
    saveClosingHistoryToStorage(nextHistory);
  };

  const handleClearHistory = () => {
    setHistory([]);
    saveClosingHistoryToStorage([]);
  };

  return (
    <div className="min-h-screen flex flex-col bg-slate-50 text-slate-900">
      {/* Strict 3-Zone Top Bar Contract */}
      <header className="sticky top-0 z-30 bg-white border-b border-slate-200 px-6 py-3.5 flex items-center justify-between gap-8">
        {/* Zone 1: Brand Wordmark (single text element) */}
        <a
          href="#importacao"
          className="text-lg font-bold tracking-tight text-slate-900 whitespace-nowrap shrink-0"
        >
          EFD Crédito Fiscal
        </a>

        {/* Zone 2: 5 concise single-line navigation links */}
        <nav className="hidden md:flex items-center gap-6 text-sm font-medium text-slate-600">
          <a
            href="#importacao"
            className="hover:text-slate-900 hover:underline underline-offset-4 transition-colors whitespace-nowrap shrink-0"
          >
            Importação
          </a>
          <a
            href="#criterios"
            className="hover:text-slate-900 hover:underline underline-offset-4 transition-colors whitespace-nowrap shrink-0"
          >
            Critérios
          </a>
          <a
            href="#conferencia"
            className="hover:text-slate-900 hover:underline underline-offset-4 transition-colors whitespace-nowrap shrink-0"
          >
            Conferência
          </a>
          <a
            href="#contabil"
            className="hover:text-slate-900 hover:underline underline-offset-4 transition-colors whitespace-nowrap shrink-0"
          >
            Dinâmica Contábil
          </a>
          <a
            href="#historico"
            className="hover:text-slate-900 hover:underline underline-offset-4 transition-colors whitespace-nowrap shrink-0"
          >
            Histórico
          </a>
        </nav>

        {/* Zone 3: 1 Primary Action Button */}
        <div className="flex items-center gap-3 shrink-0">
          <button
            type="button"
            onClick={handleExportExcel}
            disabled={!workbook || isExporting}
            className="px-4 py-2 text-xs font-semibold text-white bg-blue-900 rounded-md hover:bg-blue-800 disabled:opacity-50 transition-colors whitespace-nowrap shrink-0 cursor-pointer"
          >
            {isExporting ? 'Exportando...' : 'Exportar Consolidado .xlsx'}
          </button>
        </div>
      </header>

      {/* Main Content Container */}
      <main className="flex-1 max-w-[1440px] w-full mx-auto px-6 py-6 space-y-6">
        {/* Step 1: Import Excel & Diagnostic Status */}
        <section
          id="importacao"
          className="bg-white border border-slate-200 rounded-lg p-6 space-y-5"
        >
          <div className="flex flex-col lg:flex-row lg:items-center lg:justify-between gap-4">
            <div>
              <h1 className="text-xl font-bold text-slate-900 tracking-tight">
                01. Importação da Planilha Mensal de Créditos de PIS e COFINS
              </h1>
              <p className="text-sm text-slate-600 mt-1">
                Selecione o arquivo <span className="font-mono text-slate-800">.xlsx</span> do fechamento mensal (estrutura EFD Contribuições <span className="font-mono text-slate-800">EFD 8_2026</span>). O sistema inspeciona automaticamente as 8 abas obrigatórias e identifica os registros destacados em azul.
              </p>
            </div>

            <div className="flex flex-wrap items-center gap-2.5 shrink-0">
              <input
                ref={fileInputRef}
                type="file"
                accept=".xlsx"
                onChange={handleFileUpload}
                className="hidden"
              />
              <button
                type="button"
                onClick={() => fileInputRef.current?.click()}
                className="inline-flex items-center gap-2 px-4 py-2.5 text-xs font-semibold text-white bg-blue-900 rounded-md hover:bg-blue-800 transition-colors cursor-pointer whitespace-nowrap shrink-0"
              >
                <FolderOpen className="w-4 h-4" />
                Selecionar Arquivo Excel (.xlsx)
              </button>

              <button
                type="button"
                onClick={handleLoadDemoWorkbook}
                className="inline-flex items-center gap-1.5 px-3.5 py-2.5 text-xs font-medium text-slate-800 bg-slate-100 border border-slate-300 rounded-md hover:bg-slate-200/70 transition-colors cursor-pointer whitespace-nowrap shrink-0"
              >
                <Play className="w-3.5 h-3.5 text-blue-900" />
                Recarregar Planilha EFD 8_2026 (Demo)
              </button>

              <button
                type="button"
                onClick={handleDownloadDemoWorkbook}
                className="inline-flex items-center gap-1.5 px-3.5 py-2.5 text-xs font-medium text-slate-700 bg-white border border-slate-300 rounded-md hover:bg-slate-50 transition-colors cursor-pointer whitespace-nowrap shrink-0"
              >
                <Download className="w-3.5 h-3.5" />
                Baixar Modelo EFD 8_2026.xlsx
              </button>
            </div>
          </div>

          {isLoadingFile && (
            <div className="p-4 bg-blue-50 border border-blue-200 rounded-md text-xs text-blue-900 font-medium">
              Lendo estrutura XML da planilha, textos compartilhados e cores de preenchimento das 8 abas...
            </div>
          )}

          {loadError && (
            <div className="p-4 bg-red-50 border border-red-200 rounded-md text-xs text-red-800 font-medium">
              Erro ao ler o arquivo: {loadError}
            </div>
          )}

          {/* Diagnostic Grid of the 8 Target Tabs */}
          {workbook && (
            <div className="pt-4 border-t border-slate-200 space-y-3">
              <div className="flex flex-wrap items-center justify-between gap-2 text-xs text-slate-600">
                <div className="flex items-center gap-2">
                  <FileSpreadsheet className="w-4 h-4 text-blue-900" />
                  <span>
                    Arquivo ativo:{' '}
                    <strong className="text-slate-900">{workbook.fileName}</strong>
                  </span>
                  <span aria-hidden="true">·</span>
                  <span className="font-mono tabular-nums">
                    {(workbook.fileSize / 1024).toFixed(1)} KB
                  </span>
                  <span aria-hidden="true">·</span>
                  <span>
                    {workbook.importLogs.filter((l) => l.found).length} de{' '}
                    {EFD_TABS.length} abas localizadas
                  </span>
                </div>
                <div className="text-slate-500">
                  Nota: Na primeira versão, a captura baseia-se na cor de preenchimento azul; o app audita o cálculo matemático das alíquotas, duplicidades e contabilização.
                </div>
              </div>

              <div className="grid grid-cols-2 sm:grid-cols-4 lg:grid-cols-8 gap-2.5">
                {workbook.importLogs.map((log) => (
                  <div
                    key={log.tabName}
                    className={`p-3 rounded-md border text-xs ${
                      log.found
                        ? 'bg-slate-50 border-slate-200'
                        : 'bg-red-50/60 border-red-200'
                    }`}
                  >
                    <div className="flex items-center justify-between gap-1">
                      <span className="font-semibold text-slate-900 truncate">
                        {log.tabName}
                      </span>
                      {log.found ? (
                        <CheckCircle2 className="w-3.5 h-3.5 text-emerald-700 shrink-0" />
                      ) : (
                        <XCircle className="w-3.5 h-3.5 text-red-700 shrink-0" />
                      )}
                    </div>
                    <div className="mt-1.5 text-[11px] font-mono tabular-nums text-slate-600">
                      {log.found ? (
                        <>
                          <div>{log.rowCount} linhas lidas</div>
                          <div className="text-blue-900 font-medium">
                            {log.blueCellCount} cél. azuis
                          </div>
                        </>
                      ) : (
                        <span className="text-red-700 font-sans">
                          Aba não encontrada
                        </span>
                      )}
                    </div>
                  </div>
                ))}
              </div>
            </div>
          )}
        </section>

        {/* Step 2: Criteria & Tab Section Rules */}
        <div id="criterios">
          <CriteriaSection
            workbook={workbook}
            blueColors={blueColors}
            onToggleBlueColor={handleToggleBlueColor}
            onResetAutoBlueColors={handleResetAutoBlueColors}
            toleranceInput={toleranceInput}
            onChangeToleranceInput={setToleranceInput}
            ignoreHidden={ignoreHidden}
            onChangeIgnoreHidden={setIgnoreHidden}
            tabConfig={tabConfig}
            onChangeTabConfig={setTabConfig}
            onResetDefaultConfig={handleResetDefaultConfig}
            profiles={profiles}
            activeProfileId={activeProfileId}
            onSelectProfile={handleSelectProfile}
            onSaveProfile={handleSaveProfile}
            onDeleteProfile={handleDeleteProfile}
            onRunConsolidation={() => {
              document
                .getElementById('conferencia')
                ?.scrollIntoView({ behavior: 'smooth' });
            }}
          />
        </div>

        {/* Step 3: Conference, Reconciliation & Export */}
        <div id="conferencia">
          <ConferenceSection
            workbook={workbook}
            allRecords={allRecords}
            usedColumns={usedColumns}
            tabAggregations={tabAggregations}
            resumoTotalRows={resumoTotalRows}
            resumoBlocks={resumoBlocks}
            onToggleRowInclusion={handleToggleRowInclusion}
            onSaveJustification={handleSaveJustification}
            onClearJustification={handleClearJustification}
            onOpenManualRowModal={(tabName) => {
              setManualModalTab(tabName);
              setManualModalOpen(true);
            }}
            onExportExcel={handleExportExcel}
            onSaveToHistory={handleSaveToHistory}
            isExporting={isExporting}
          />
        </div>

        {/* Step 4: Accounting Dynamics (Dinâmica Contábil) */}
        <div id="contabil">
          <AccountingDynamicsSection
            blocks={accountingBlocks}
            accountingMap={accountingMap}
            onChangeAccountingMap={handleChangeAccountingMap}
            competencia={competencia}
            onChangeCompetencia={setCompetencia}
            ctaDebitoPis={ctaDebitoPis}
            onChangeCtaDebitoPis={setCtaDebitoPis}
            ctaDebitoCofins={ctaDebitoCofins}
            onChangeCtaDebitoCofins={setCtaDebitoCofins}
            onExportExcel={handleExportExcel}
            isExporting={isExporting}
          />
        </div>

        {/* Step 5: Monthly Closing History */}
        <div id="historico">
          <HistorySection
            history={history}
            onDeleteEntry={handleDeleteHistoryEntry}
            onClearHistory={handleClearHistory}
          />
        </div>
      </main>

      {/* Clean Footer */}
      <footer className="bg-white border-t border-slate-200 px-6 py-4 mt-8">
        <div className="max-w-[1440px] mx-auto flex flex-wrap items-center justify-between gap-4 text-xs text-slate-500">
          <span>
            EFD Crédito Fiscal — Consolidação Mensal de Créditos de PIS e COFINS
          </span>
          <span>
            Estrutura compatível com EFD Contribuições (Sucata 1 · Uso e consumo · Nfse · 04 e 05 Outros · Locação · Imobilizado · Frete · Resumo · Dinâmica Contábil)
          </span>
        </div>
      </footer>

      {/* Modal for inspecting all rows in a sheet (including non-blue rows) */}
      <ManualRowModal
        isOpen={manualModalOpen}
        onClose={() => setManualModalOpen(false)}
        workbook={workbook}
        initialTab={manualModalTab}
        tabConfig={tabConfig}
        blueColors={blueColors}
        ignoreHidden={ignoreHidden}
        manualState={manualState}
        onToggleRowInclusion={handleToggleRowInclusion}
      />
    </div>
  );
}
