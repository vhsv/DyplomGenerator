import React, { useState, useCallback, useRef, useMemo } from 'react';
import Docxtemplater from 'docxtemplater';
import PizZip from 'pizzip';
import JSZip from 'jszip';
import { saveAs } from 'file-saver';

// ─── Types ────────────────────────────────────────────────────────────────────

type Phase = 'upload' | 'map' | 'generating' | 'done';

interface TemplateField { key: string; }

interface ColumnMapping {
  columnIndex: number | null; // which parsed column maps to this field; null = static value
  staticValue: string;
}

interface GenerateResult {
  row: number;
  name: string;
  ok: boolean;
  error?: string;
}

// ─── Helpers ──────────────────────────────────────────────────────────────────

function extractTags(xmlContent: string): string[] {
  const clean = xmlContent.replace(/<[^>]+>/g, ' ');
  const regex = /\{([^}]+)\}/g;
  const tags = new Set<string>();
  let match;
  while ((match = regex.exec(clean)) !== null) {
    const tag = match[1].trim();
    if (tag && !tag.startsWith('#') && !tag.startsWith('/') && !tag.startsWith('^'))
      tags.add(tag);
  }
  return Array.from(tags);
}

// Detect separator and parse textarea into rows×columns
function parseTextarea(text: string): string[][] {
  const lines = text.split('\n').map(l => l.trimEnd()).filter(l => l.trim() !== '');
  if (!lines.length) return [];
  // detect: tab > semicolon > comma > space
  const sample = lines.slice(0, 5).join('\n');
  let sep = '\t';
  if ((sample.match(/\t/g) || []).length < (sample.match(/;/g) || []).length) sep = ';';
  if (sep !== '\t' && (sample.match(/;/g) || []).length < (sample.match(/,/g) || []).length) sep = ',';
  return lines.map(l => l.split(sep).map(c => c.trim()));
}

function generateDocx(templateData: ArrayBuffer, values: Record<string, string>): Blob {
  const zip = new PizZip(templateData);
  const doc = new Docxtemplater(zip, { paragraphLoop: true, linebreaks: true });
  doc.render(values);
  return doc.getZip().generate({
    type: 'blob',
    mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  });
}

// ─── App ──────────────────────────────────────────────────────────────────────

export default function App() {
  const [phase, setPhase] = useState<Phase>('upload');
  const [isDragging, setIsDragging] = useState(false);
  const [fileName, setFileName] = useState('');
  const [fileData, setFileData] = useState<ArrayBuffer | null>(null);
  const [fields, setFields] = useState<TemplateField[]>([]);
  const [error, setError] = useState<string | null>(null);

  // Bulk input
  const [rawText, setRawText] = useState('');
  const [hasHeader, setHasHeader] = useState(true);
  const [mappings, setMappings] = useState<Record<string, ColumnMapping>>({});
  const [fileNameField, setFileNameField] = useState<string>(''); // which field to use as filename

  // Generation
  const [progress, setProgress] = useState(0);
  const [results, setResults] = useState<GenerateResult[]>([]);
  const [generatedCount, setGeneratedCount] = useState(0);

  const fileInputRef = useRef<HTMLInputElement>(null);

  // Parse preview
  const parsed = useMemo(() => parseTextarea(rawText), [rawText]);
  const headers = useMemo(() => hasHeader && parsed.length > 0 ? parsed[0] : [], [parsed, hasHeader]);
  const dataRows = useMemo(() => hasHeader ? parsed.slice(1) : parsed, [parsed, hasHeader]);
  const colCount = useMemo(() => Math.max(0, ...parsed.map(r => r.length)), [parsed]);

  // ── File loading ────────────────────────────────────────────────────────────

  const processFile = useCallback((file: File) => {
    if (!file.name.endsWith('.docx')) { setError('Wybierz plik .docx'); return; }
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = e.target?.result as ArrayBuffer;
        const zip = new PizZip(data);
        let xml = '';
        Object.keys(zip.files).forEach(n => { if (n.endsWith('.xml')) try { xml += zip.files[n].asText(); } catch {} });
        const tags = extractTags(xml);
        setFields(tags.map(k => ({ key: k })));
        setFileName(file.name);
        setFileData(data);
        // Init mappings: first field → col 0, rest → static empty
        const initMap: Record<string, ColumnMapping> = {};
        tags.forEach((k, i) => { initMap[k] = { columnIndex: i < 1 ? 0 : null, staticValue: '' }; });
        setMappings(initMap);
        setFileNameField(tags[0] || '');
        setPhase('map');
        setError(null);
      } catch { setError('Nie udało się odczytać pliku .docx'); }
    };
    reader.readAsArrayBuffer(file);
  }, []);

  const handleDrop = useCallback((e: React.DragEvent) => {
    e.preventDefault(); setIsDragging(false);
    const f = e.dataTransfer.files[0]; if (f) processFile(f);
  }, [processFile]);

  const handleFileChange = useCallback((e: React.ChangeEvent<HTMLInputElement>) => {
    const f = e.target.files?.[0]; if (f) processFile(f);
  }, [processFile]);

  // ── Mapping helpers ─────────────────────────────────────────────────────────

  const setMapping = useCallback((key: string, upd: Partial<ColumnMapping>) => {
    setMappings(prev => ({ ...prev, [key]: { ...prev[key], ...upd } }));
  }, []);

  // ── Generate all ────────────────────────────────────────────────────────────

  const handleGenerate = useCallback(async () => {
    if (!fileData || dataRows.length === 0) return;
    setPhase('generating');
    setProgress(0);
    setResults([]);

    const outputZip = new JSZip();
    const res: GenerateResult[] = [];

    for (let i = 0; i < dataRows.length; i++) {
      const row = dataRows[i];
      const values: Record<string, string> = {};
      fields.forEach(({ key }) => {
        const m = mappings[key];
        if (!m) { values[key] = ''; return; }
        if (m.columnIndex !== null) {
          values[key] = row[m.columnIndex] ?? '';
        } else {
          values[key] = m.staticValue;
        }
      });

      // Build output filename
      const nameVal = fileNameField && values[fileNameField]
        ? values[fileNameField].replace(/[^\w\s\-_.()]/g, '').trim()
        : `dokument_${i + 1}`;
      const outName = `${nameVal}.docx`;

      try {
        const blob = generateDocx(fileData, values);
        outputZip.file(outName, blob);
        res.push({ row: i + 1, name: outName, ok: true });
      } catch (err: any) {
        res.push({ row: i + 1, name: outName, ok: false, error: err.message });
      }

      setProgress(Math.round(((i + 1) / dataRows.length) * 100));
      // Let UI breathe
      if (i % 10 === 0) await new Promise(r => setTimeout(r, 0));
    }

    setResults(res);
    setGeneratedCount(res.filter(r => r.ok).length);

    const zipBlob = await outputZip.generateAsync({ type: 'blob' });
    saveAs(zipBlob, fileName.replace('.docx', `_masowe_${dataRows.length}szt.zip`));
    setPhase('done');
  }, [fileData, dataRows, fields, mappings, fileNameField, fileName]);

  const handleReset = useCallback(() => {
    setPhase('upload'); setFileName(''); setFileData(null); setFields([]);
    setRawText(''); setMappings({}); setResults([]); setError(null);
    if (fileInputRef.current) fileInputRef.current.value = '';
  }, []);

  // ── Phase labels ────────────────────────────────────────────────────────────

  const phaseIndex = phase === 'upload' ? 0 : phase === 'map' ? 1 : phase === 'generating' ? 2 : 3;

  // ── Render ──────────────────────────────────────────────────────────────────

  return (
    <div className="app">
      <div className="grid-bg" />
      <div className="noise-overlay" />

      <header className="header">
        <div className="logo">
          <svg className="logo-icon" width="30" height="30" viewBox="0 0 32 32" fill="none">
            <rect x="4" y="2" width="18" height="24" rx="2" fill="#1a1a2e" stroke="#4f8ef7" strokeWidth="1.5"/>
            <rect x="16" y="2" width="6" height="6" rx="1" fill="#4f8ef7" opacity="0.4"/>
            <line x1="8" y1="12" x2="18" y2="12" stroke="#4f8ef7" strokeWidth="1.5" strokeLinecap="round"/>
            <line x1="8" y1="16" x2="18" y2="16" stroke="#4f8ef7" strokeWidth="1.5" strokeLinecap="round" opacity="0.6"/>
            <line x1="8" y1="20" x2="14" y2="20" stroke="#4f8ef7" strokeWidth="1.5" strokeLinecap="round" opacity="0.4"/>
            <circle cx="24" cy="24" r="6" fill="#0f0f1a" stroke="#4f8ef7" strokeWidth="1.5"/>
            <line x1="21" y1="24" x2="27" y2="24" stroke="#4f8ef7" strokeWidth="1.5" strokeLinecap="round"/>
            <line x1="24" y1="21" x2="24" y2="27" stroke="#4f8ef7" strokeWidth="1.5" strokeLinecap="round"/>
          </svg>
          <span>DOCX<span className="logo-accent">bulk</span></span>
        </div>
        <p className="header-sub">Masowe generowanie dokumentów Word · lokalnie, bez serwera</p>
      </header>

      <main className="main">
        {error && (
          <div className="error-banner">
            <span>⚠ {error}</span>
            <button onClick={() => setError(null)}>✕</button>
          </div>
        )}

        {/* Steps */}
        <div className="steps-bar">
          {['Wgraj szablon', 'Dane i mapowanie', 'Generuj ZIP'].map((label, i) => (
            <React.Fragment key={i}>
              <div className={`step-node ${phaseIndex > i ? 'done' : phaseIndex === i ? 'active' : ''}`}>
                <span className="step-num">{phaseIndex > i ? '✓' : `0${i + 1}`}</span>
                <span className="step-label">{label}</span>
              </div>
              {i < 2 && <div className={`step-connector ${phaseIndex > i ? 'done' : ''}`} />}
            </React.Fragment>
          ))}
        </div>

        {/* ── UPLOAD ── */}
        {phase === 'upload' && (
          <div className="phase-panel">
            <div
              className={`drop-zone ${isDragging ? 'dragging' : ''}`}
              onDragOver={e => { e.preventDefault(); setIsDragging(true); }}
              onDragLeave={() => setIsDragging(false)}
              onDrop={handleDrop}
              onClick={() => fileInputRef.current?.click()}
            >
              <input ref={fileInputRef} type="file" accept=".docx" onChange={handleFileChange} style={{ display: 'none' }} />
              <div className="drop-visual">
                <svg width="56" height="56" viewBox="0 0 56 56" fill="none">
                  <path d="M10 36 L28 18 L46 36" stroke="#4f8ef7" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round"/>
                  <line x1="28" y1="18" x2="28" y2="44" stroke="#4f8ef7" strokeWidth="2.5" strokeLinecap="round"/>
                  <path d="M14 48 H42" stroke="#4f8ef7" strokeWidth="2" strokeLinecap="round" opacity="0.4"/>
                </svg>
              </div>
              <p className="drop-title">Wgraj szablon .docx</p>
              <p className="drop-hint">Przeciągnij lub kliknij · plik z polami <code>{`{imie}`}</code></p>
            </div>

            <div className="info-panel">
              <h3 className="info-title">Jak działa masowe generowanie?</h3>
              <div className="info-steps">
                <div className="info-step">
                  <span className="info-num">1</span>
                  <div>
                    <strong>Szablon Word z placeholderami</strong>
                    <p>Utwórz dokument z polami <code>{`{imie}`}</code>, <code>{`{nazwisko}`}</code>, <code>{`{kwota}`}</code> itp.</p>
                  </div>
                </div>
                <div className="info-step">
                  <span className="info-num">2</span>
                  <div>
                    <strong>Wklej dane z Excela / Notatnika</strong>
                    <p>Skopiuj kolumny i wklej — obsługuje tabulatory, średniki, przecinki</p>
                  </div>
                </div>
                <div className="info-step">
                  <span className="info-num">3</span>
                  <div>
                    <strong>Mapuj kolumny → pola</strong>
                    <p>Wskaż, która kolumna trafia do którego pola w szablonie</p>
                  </div>
                </div>
                <div className="info-step">
                  <span className="info-num">4</span>
                  <div>
                    <strong>Pobierz ZIP ze wszystkimi plikami</strong>
                    <p>Każdy wiersz = osobny plik .docx spakowany w archiwum ZIP</p>
                  </div>
                </div>
              </div>
            </div>
          </div>
        )}

        {/* ── MAP ── */}
        {phase === 'map' && (
          <div className="phase-panel">
            <div className="file-info-bar">
              <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="#4f8ef7" strokeWidth="2">
                <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/>
                <polyline points="14 2 14 8 20 8"/>
              </svg>
              <span className="file-name">{fileName}</span>
              <span className="field-count-badge">{fields.length} pól wykryto</span>
              <button className="btn-back" style={{ marginLeft: 'auto' }} onClick={handleReset}>← Zmień plik</button>
            </div>

            {/* Textarea input */}
            <div className="section-card">
              <div className="section-header">
                <div>
                  <h2>Dane do wypełnienia</h2>
                  <p className="section-sub">Wklej z Excela, Notatnika — oddzielone tabulatorem, średnikiem lub przecinkiem</p>
                </div>
                <label className="toggle-label">
                  <input
                    type="checkbox"
                    checked={hasHeader}
                    onChange={e => setHasHeader(e.target.checked)}
                  />
                  <span className="toggle-track"><span className="toggle-thumb" /></span>
                  <span>Pierwszy wiersz to nagłówki</span>
                </label>
              </div>
              <textarea
                className="data-textarea"
                placeholder={`Jan\tKowalski\t1000 zł\nAnna\tNowak\t2500 zł\nPiotr\tWiśniewski\t750 zł`}
                value={rawText}
                onChange={e => setRawText(e.target.value)}
                spellCheck={false}
              />
              {parsed.length > 0 && (
                <div className="parse-stats">
                  <span className="stat-pill">
                    <svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"><rect x="3" y="3" width="18" height="18" rx="2"/><line x1="3" y1="9" x2="21" y2="9"/><line x1="9" y1="3" x2="9" y2="21"/></svg>
                    {dataRows.length} wierszy danych
                  </span>
                  <span className="stat-pill">{colCount} kolumn</span>
                  {headers.length > 0 && <span className="stat-pill">Nagłówki: {headers.join(', ')}</span>}
                </div>
              )}
            </div>

            {/* Preview table */}
            {dataRows.length > 0 && (
              <div className="section-card">
                <h2 className="section-h2">Podgląd danych</h2>
                <div className="table-scroll">
                  <table className="preview-table">
                    <thead>
                      <tr>
                        <th>#</th>
                        {Array.from({ length: colCount }, (_, i) => (
                          <th key={i}>
                            <div className="col-header">
                              <span className="col-idx">kol. {i}</span>
                              {headers[i] && <span className="col-name">{headers[i]}</span>}
                            </div>
                          </th>
                        ))}
                      </tr>
                    </thead>
                    <tbody>
                      {dataRows.slice(0, 5).map((row, ri) => (
                        <tr key={ri}>
                          <td className="row-num">{ri + 1}</td>
                          {Array.from({ length: colCount }, (_, ci) => (
                            <td key={ci}>{row[ci] ?? ''}</td>
                          ))}
                        </tr>
                      ))}
                      {dataRows.length > 5 && (
                        <tr className="more-rows">
                          <td colSpan={colCount + 1}>…i {dataRows.length - 5} więcej wierszy</td>
                        </tr>
                      )}
                    </tbody>
                  </table>
                </div>
              </div>
            )}

            {/* Mapping */}
            <div className="section-card">
              <div className="section-header">
                <div>
                  <h2>Mapowanie pól szablonu</h2>
                  <p className="section-sub">Przypisz każdemu polu kolumnę z danych lub wpisz stałą wartość</p>
                </div>
              </div>
              <div className="mapping-grid">
                {fields.map(({ key }) => {
                  const m = mappings[key] ?? { columnIndex: null, staticValue: '' };
                  const useCol = m.columnIndex !== null;
                  return (
                    <div key={key} className="mapping-row">
                      <div className="mapping-field-tag">{`{${key}}`}</div>
                      <div className="mapping-arrow">→</div>
                      <div className="mapping-controls">
                        <div className="mapping-type-switch">
                          <button
                            className={`mtype-btn ${useCol ? 'active' : ''}`}
                            onClick={() => setMapping(key, { columnIndex: 0 })}
                          >Z kolumny</button>
                          <button
                            className={`mtype-btn ${!useCol ? 'active' : ''}`}
                            onClick={() => setMapping(key, { columnIndex: null })}
                          >Stała wartość</button>
                        </div>
                        {useCol ? (
                          <select
                            className="col-select"
                            value={m.columnIndex ?? 0}
                            onChange={e => setMapping(key, { columnIndex: Number(e.target.value) })}
                          >
                            {Array.from({ length: colCount }, (_, i) => (
                              <option key={i} value={i}>
                                Kol. {i}{headers[i] ? ` — ${headers[i]}` : ''}
                                {dataRows[0]?.[i] ? ` (np. "${dataRows[0][i]}")` : ''}
                              </option>
                            ))}
                          </select>
                        ) : (
                          <input
                            className="static-input"
                            type="text"
                            placeholder="Stała wartość dla wszystkich…"
                            value={m.staticValue}
                            onChange={e => setMapping(key, { staticValue: e.target.value })}
                          />
                        )}
                      </div>
                    </div>
                  );
                })}
              </div>
            </div>

            {/* Filename config */}
            <div className="section-card filename-card">
              <h2 className="section-h2">Nazwy plików wyjściowych</h2>
              <p className="section-sub" style={{ marginBottom: 12 }}>Każdy plik będzie nazwany wartością wybranego pola</p>
              <select
                className="col-select"
                value={fileNameField}
                onChange={e => setFileNameField(e.target.value)}
              >
                <option value="">— Numeruj automatycznie (dokument_1.docx) —</option>
                {fields.map(({ key }) => (
                  <option key={key} value={key}>{`{${key}}`} — wartość z pola</option>
                ))}
              </select>
            </div>

            <div className="action-bar">
              <button className="btn-back" onClick={handleReset}>← Wróć</button>
              <button
                className="btn-generate"
                disabled={dataRows.length === 0}
                onClick={handleGenerate}
              >
                <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5">
                  <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/>
                  <polyline points="7 10 12 15 17 10"/>
                  <line x1="12" y1="15" x2="12" y2="3"/>
                </svg>
                Generuj {dataRows.length} dokument{dataRows.length === 1 ? '' : dataRows.length < 5 ? 'y' : 'ów'} → ZIP
              </button>
            </div>
          </div>
        )}

        {/* ── GENERATING ── */}
        {phase === 'generating' && (
          <div className="phase-panel center-panel">
            <div className="gen-spinner">
              <svg width="80" height="80" viewBox="0 0 80 80" fill="none">
                <circle cx="40" cy="40" r="34" stroke="#1e2040" strokeWidth="6"/>
                <path d="M40 6 A34 34 0 0 1 74 40" stroke="#4f8ef7" strokeWidth="6" strokeLinecap="round">
                  <animateTransform attributeName="transform" type="rotate" from="0 40 40" to="360 40 40" dur="0.9s" repeatCount="indefinite"/>
                </path>
              </svg>
            </div>
            <h2 className="gen-title">Generowanie…</h2>
            <div className="progress-bar-wrap">
              <div className="progress-bar-track">
                <div className="progress-bar-fill" style={{ width: `${progress}%` }} />
              </div>
              <span className="progress-label">{progress}%</span>
            </div>
            <p className="gen-sub">Tworzę dokumenty i pakuję do ZIP — proszę czekać</p>
          </div>
        )}

        {/* ── DONE ── */}
        {phase === 'done' && (
          <div className="phase-panel">
            <div className="center-panel" style={{ marginBottom: 24 }}>
              <div className="success-ring">
                <svg width="80" height="80" viewBox="0 0 80 80" fill="none">
                  <circle cx="40" cy="40" r="36" stroke="#4f8ef7" strokeWidth="2" opacity="0.3"/>
                  <circle cx="40" cy="40" r="28" stroke="#4f8ef7" strokeWidth="2"/>
                  <path d="M26 40 L35 49 L54 30" stroke="#4f8ef7" strokeWidth="3" strokeLinecap="round" strokeLinejoin="round"/>
                </svg>
              </div>
              <h2 className="success-title">ZIP pobrany!</h2>
              <p className="success-sub">
                Wygenerowano <strong>{generatedCount}</strong> z <strong>{results.length}</strong> dokumentów.
                {results.some(r => !r.ok) && <> <span style={{ color: 'var(--error)' }}>{results.filter(r => !r.ok).length} błędów.</span></>}
              </p>
              <button className="btn-generate" onClick={handleReset} style={{ marginTop: 20 }}>
                Generuj ponownie / nowy szablon
              </button>
            </div>

            {/* Results log */}
            <div className="section-card">
              <h2 className="section-h2">Log generowania</h2>
              <div className="results-log">
                {results.map(r => (
                  <div key={r.row} className={`log-row ${r.ok ? 'ok' : 'err'}`}>
                    <span className="log-num">#{r.row}</span>
                    <span className="log-name">{r.name}</span>
                    <span className="log-status">{r.ok ? '✓ OK' : `✕ ${r.error}`}</span>
                  </div>
                ))}
              </div>
            </div>
          </div>
        )}
      </main>

      <footer className="footer">
        Działa w 100% lokalnie · Twoje dane nigdy nie opuszczają przeglądarki
      </footer>
    </div>
  );
}
