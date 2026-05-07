import { useState, useRef, useCallback } from "react";
import * as mammoth from "mammoth";
import * as XLSX from "xlsx";
import * as pdfjsLib from "pdfjs-dist";

// Use a CDN worker for pdf.js
pdfjsLib.GlobalWorkerOptions.workerSrc =
  "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js";

const ACCENT = "#f97316";
const BG = "#0c0c0e";
const SURFACE = "#16161a";
const SURFACE2 = "#1e1e24";
const BORDER = "#2a2a32";
const TEXT = "#e8e8f0";
const MUTED = "#6b6b80";

const FILE_TYPES = {
  pdf: { label: "PDF", color: "#ef4444", icon: "📄" },
  docx: { label: "Word", color: "#3b82f6", icon: "📝" },
  doc: { label: "Word", color: "#3b82f6", icon: "📝" },
  xlsx: { label: "Excel", color: "#22c55e", icon: "📊" },
  xls: { label: "Excel", color: "#22c55e", icon: "📊" },
  csv: { label: "CSV", color: "#22c55e", icon: "📋" },
  txt: { label: "Text", color: "#a855f7", icon: "📃" },
  md: { label: "Markdown", color: "#f59e0b", icon: "✍️" },
  json: { label: "JSON", color: "#06b6d4", icon: "🔧" },
  png: { label: "Image", color: "#ec4899", icon: "🖼️" },
  jpg: { label: "Image", color: "#ec4899", icon: "🖼️" },
  jpeg: { label: "Image", color: "#ec4899", icon: "🖼️" },
  gif: { label: "Image", color: "#ec4899", icon: "🖼️" },
  webp: { label: "Image", color: "#ec4899", icon: "🖼️" },
  svg: { label: "SVG", color: "#f97316", icon: "🎨" },
};

function getExt(name) {
  return name.split(".").pop().toLowerCase();
}

function FileBadge({ ext }) {
  const info = FILE_TYPES[ext] || { label: ext.toUpperCase(), color: MUTED, icon: "📁" };
  return (
    <span style={{
      background: info.color + "22",
      color: info.color,
      border: `1px solid ${info.color}44`,
      borderRadius: 6,
      padding: "2px 8px",
      fontSize: 11,
      fontWeight: 700,
      letterSpacing: "0.08em",
      fontFamily: "monospace",
    }}>
      {info.icon} {info.label}
    </span>
  );
}

function PDFViewer({ file }) {
  const [pages, setPages] = useState([]);
  const [loading, setLoading] = useState(true);
  const [numPages, setNumPages] = useState(0);

  useState(() => {
    let cancelled = false;
    (async () => {
      try {
        const ab = await file.arrayBuffer();
        const pdf = await pdfjsLib.getDocument({ data: ab }).promise;
        setNumPages(pdf.numPages);
        const pageImgs = [];
        for (let i = 1; i <= Math.min(pdf.numPages, 20); i++) {
          const page = await pdf.getPage(i);
          const vp = page.getViewport({ scale: 1.5 });
          const canvas = document.createElement("canvas");
          canvas.width = vp.width;
          canvas.height = vp.height;
          await page.render({ canvasContext: canvas.getContext("2d"), viewport: vp }).promise;
          pageImgs.push({ src: canvas.toDataURL(), num: i });
        }
        if (!cancelled) { setPages(pageImgs); setLoading(false); }
      } catch (e) {
        if (!cancelled) setLoading(false);
      }
    })();
    return () => { cancelled = true; };
  }, []);

  if (loading) return <Spinner label="Rendering PDF pages…" />;
  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 16, alignItems: "center" }}>
      {numPages > 20 && (
        <div style={{ color: MUTED, fontSize: 13 }}>Showing first 20 of {numPages} pages</div>
      )}
      {pages.map(p => (
        <div key={p.num} style={{ position: "relative", width: "100%" }}>
          <div style={{ position: "absolute", top: 8, left: 8, background: "#000a", color: "#fff", borderRadius: 4, padding: "2px 8px", fontSize: 11 }}>p.{p.num}</div>
          <img src={p.src} alt={`Page ${p.num}`} style={{ width: "100%", borderRadius: 8, display: "block", border: `1px solid ${BORDER}` }} />
        </div>
      ))}
    </div>
  );
}

function DocxViewer({ file }) {
  const [html, setHtml] = useState("");
  const [loading, setLoading] = useState(true);

  useState(() => {
    let cancelled = false;
    (async () => {
      try {
        const ab = await file.arrayBuffer();
        const result = await mammoth.convertToHtml({ arrayBuffer: ab });
        if (!cancelled) { setHtml(result.value); setLoading(false); }
      } catch { if (!cancelled) setLoading(false); }
    })();
    return () => { cancelled = true; };
  }, []);

  if (loading) return <Spinner label="Parsing document…" />;
  return (
    <div
      style={{
        background: "#fff",
        borderRadius: 8,
        padding: "32px 40px",
        color: "#111",
        lineHeight: 1.7,
        fontSize: 15,
        fontFamily: "'Georgia', serif",
        boxShadow: "0 4px 32px #0008",
      }}
      dangerouslySetInnerHTML={{ __html: html || "<p style='color:#888'>No content found.</p>" }}
    />
  );
}

function ExcelViewer({ file }) {
  const [sheets, setSheets] = useState([]);
  const [active, setActive] = useState(0);
  const [loading, setLoading] = useState(true);

  useState(() => {
    let cancelled = false;
    (async () => {
      try {
        const ab = await file.arrayBuffer();
        const wb = XLSX.read(ab, { type: "array" });
        const parsed = wb.SheetNames.map(name => ({
          name,
          data: XLSX.utils.sheet_to_json(wb.Sheets[name], { header: 1 }),
        }));
        if (!cancelled) { setSheets(parsed); setLoading(false); }
      } catch { if (!cancelled) setLoading(false); }
    })();
    return () => { cancelled = true; };
  }, []);

  if (loading) return <Spinner label="Parsing spreadsheet…" />;
  const sheet = sheets[active];

  return (
    <div>
      {sheets.length > 1 && (
        <div style={{ display: "flex", gap: 8, marginBottom: 16, flexWrap: "wrap" }}>
          {sheets.map((s, i) => (
            <button key={s.name} onClick={() => setActive(i)} style={{
              background: i === active ? ACCENT : SURFACE2,
              color: i === active ? "#fff" : TEXT,
              border: "none", borderRadius: 6, padding: "6px 14px", cursor: "pointer", fontSize: 13, fontWeight: 600
            }}>{s.name}</button>
          ))}
        </div>
      )}
      {sheet && sheet.data.length > 0 ? (
        <div style={{ overflowX: "auto", borderRadius: 8, border: `1px solid ${BORDER}` }}>
          <table style={{ borderCollapse: "collapse", width: "100%", fontSize: 13 }}>
            <thead>
              <tr>
                {(sheet.data[0] || []).map((cell, ci) => (
                  <th key={ci} style={{
                    background: SURFACE2, color: TEXT, padding: "10px 14px",
                    borderBottom: `2px solid ${ACCENT}`, textAlign: "left",
                    fontWeight: 700, whiteSpace: "nowrap", position: "sticky", top: 0
                  }}>{cell ?? ""}</th>
                ))}
              </tr>
            </thead>
            <tbody>
              {sheet.data.slice(1, 200).map((row, ri) => (
                <tr key={ri} style={{ background: ri % 2 === 0 ? SURFACE : SURFACE2 }}>
                  {(sheet.data[0] || []).map((_, ci) => (
                    <td key={ci} style={{ padding: "8px 14px", color: TEXT, borderBottom: `1px solid ${BORDER}`, whiteSpace: "nowrap" }}>
                      {row[ci] ?? ""}
                    </td>
                  ))}
                </tr>
              ))}
            </tbody>
          </table>
          {sheet.data.length > 201 && (
            <div style={{ padding: 12, color: MUTED, fontSize: 12, textAlign: "center" }}>
              Showing first 200 rows of {sheet.data.length - 1}
            </div>
          )}
        </div>
      ) : <div style={{ color: MUTED }}>Empty sheet.</div>}
    </div>
  );
}

function CSVViewer({ file }) {
  const [data, setData] = useState([]);
  const [loading, setLoading] = useState(true);

  useState(() => {
    let cancelled = false;
    file.text().then(text => {
      if (cancelled) return;
      const rows = text.trim().split("\n").map(r => r.split(",").map(c => c.trim().replace(/^"|"$/g, "")));
      setData(rows);
      setLoading(false);
    });
    return () => { cancelled = true; };
  }, []);

  if (loading) return <Spinner label="Parsing CSV…" />;
  if (!data.length) return <div style={{ color: MUTED }}>Empty file.</div>;

  return (
    <div style={{ overflowX: "auto", borderRadius: 8, border: `1px solid ${BORDER}` }}>
      <table style={{ borderCollapse: "collapse", width: "100%", fontSize: 13 }}>
        <thead>
          <tr>{data[0].map((c, i) => (
            <th key={i} style={{ background: SURFACE2, color: TEXT, padding: "10px 14px", borderBottom: `2px solid ${ACCENT}`, textAlign: "left", fontWeight: 700 }}>{c}</th>
          ))}</tr>
        </thead>
        <tbody>
          {data.slice(1, 200).map((row, ri) => (
            <tr key={ri} style={{ background: ri % 2 === 0 ? SURFACE : SURFACE2 }}>
              {row.map((c, ci) => <td key={ci} style={{ padding: "8px 14px", color: TEXT, borderBottom: `1px solid ${BORDER}` }}>{c}</td>)}
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
}

function TextViewer({ file, ext }) {
  const [text, setText] = useState("");
  const [loading, setLoading] = useState(true);

  useState(() => {
    let cancelled = false;
    file.text().then(t => { if (!cancelled) { setText(t); setLoading(false); } });
    return () => { cancelled = true; };
  }, []);

  if (loading) return <Spinner label="Loading…" />;

  if (ext === "json") {
    let formatted = text;
    try { formatted = JSON.stringify(JSON.parse(text), null, 2); } catch {}
    return (
      <pre style={{ background: SURFACE2, borderRadius: 8, padding: 24, color: "#7dd3fc", fontSize: 13, overflowX: "auto", fontFamily: "'Fira Mono', monospace", border: `1px solid ${BORDER}`, lineHeight: 1.6 }}>
        {formatted}
      </pre>
    );
  }

  return (
    <pre style={{ background: SURFACE2, borderRadius: 8, padding: 24, color: TEXT, fontSize: 14, overflowX: "auto", fontFamily: ext === "md" ? "inherit" : "'Fira Mono', monospace", border: `1px solid ${BORDER}`, lineHeight: 1.7, whiteSpace: "pre-wrap", wordBreak: "break-word" }}>
      {text}
    </pre>
  );
}

function ImageViewer({ file }) {
  const [src, setSrc] = useState("");
  useState(() => {
    const url = URL.createObjectURL(file);
    setSrc(url);
    return () => URL.revokeObjectURL(url);
  }, []);
  return (
    <div style={{ textAlign: "center" }}>
      <img src={src} alt={file.name} style={{ maxWidth: "100%", borderRadius: 8, boxShadow: "0 8px 40px #0008", border: `1px solid ${BORDER}` }} />
    </div>
  );
}

function Spinner({ label }) {
  return (
    <div style={{ display: "flex", flexDirection: "column", alignItems: "center", gap: 16, padding: 48 }}>
      <div style={{
        width: 40, height: 40, border: `3px solid ${BORDER}`,
        borderTop: `3px solid ${ACCENT}`, borderRadius: "50%",
        animation: "spin 0.8s linear infinite"
      }} />
      <div style={{ color: MUTED, fontSize: 14 }}>{label}</div>
      <style>{`@keyframes spin { to { transform: rotate(360deg); } }`}</style>
    </div>
  );
}

function FileViewer({ file }) {
  const ext = getExt(file.name);
  if (["png", "jpg", "jpeg", "gif", "webp", "svg"].includes(ext)) return <ImageViewer file={file} />;
  if (ext === "pdf") return <PDFViewer file={file} />;
  if (["docx", "doc"].includes(ext)) return <DocxViewer file={file} />;
  if (["xlsx", "xls"].includes(ext)) return <ExcelViewer file={file} />;
  if (ext === "csv") return <CSVViewer file={file} />;
  if (["txt", "md", "json"].includes(ext)) return <TextViewer file={file} ext={ext} />;
  return (
    <div style={{ textAlign: "center", padding: 48, color: MUTED }}>
      <div style={{ fontSize: 48, marginBottom: 12 }}>🗂️</div>
      <div style={{ fontSize: 16, marginBottom: 8 }}>Preview not available</div>
      <div style={{ fontSize: 13 }}>File type <strong style={{ color: TEXT }}>.{ext}</strong> is not supported for preview.</div>
    </div>
  );
}

export default function App() {
  const [files, setFiles] = useState([]);
  const [active, setActive] = useState(null);
  const [dragging, setDragging] = useState(false);
  const inputRef = useRef();

  const addFiles = useCallback((newFiles) => {
    const arr = Array.from(newFiles);
    setFiles(prev => {
      const merged = [...prev, ...arr.filter(f => !prev.find(p => p.name === f.name && p.size === f.size))];
      if (active === null && merged.length > 0) setActive(0);
      return merged;
    });
    if (active === null && arr.length > 0) setActive(0);
  }, [active]);

  const onDrop = useCallback((e) => {
    e.preventDefault();
    setDragging(false);
    if (e.dataTransfer.files.length) addFiles(e.dataTransfer.files);
  }, [addFiles]);

  const removeFile = (i) => {
    setFiles(prev => {
      const next = prev.filter((_, idx) => idx !== i);
      setActive(next.length === 0 ? null : Math.min(i, next.length - 1));
      return next;
    });
  };

  const currentFile = files[active] ?? null;

  return (
    <div style={{
      minHeight: "100vh",
      background: BG,
      color: TEXT,
      fontFamily: "'DM Sans', 'Segoe UI', sans-serif",
      display: "flex",
      flexDirection: "column",
    }}>
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@400;500;600;700&family=DM+Mono&display=swap');
        * { box-sizing: border-box; margin: 0; padding: 0; }
        ::-webkit-scrollbar { width: 6px; height: 6px; }
        ::-webkit-scrollbar-track { background: ${SURFACE}; }
        ::-webkit-scrollbar-thumb { background: ${BORDER}; border-radius: 3px; }
        ::-webkit-scrollbar-thumb:hover { background: ${MUTED}; }
        body { background: ${BG}; }
        .tab:hover { background: ${SURFACE2} !important; }
        .tab.active-tab { background: ${SURFACE2} !important; border-top: 2px solid ${ACCENT} !important; }
        .drop-zone.drag-over { border-color: ${ACCENT} !important; background: ${ACCENT}11 !important; }
        .close-btn:hover { color: #ef4444 !important; }
        .upload-btn:hover { background: ${ACCENT} !important; color: #fff !important; }
      `}</style>

      {/* Header */}
      <div style={{ borderBottom: `1px solid ${BORDER}`, padding: "16px 28px", display: "flex", alignItems: "center", gap: 16 }}>
        <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
          <div style={{ width: 32, height: 32, borderRadius: 8, background: ACCENT, display: "flex", alignItems: "center", justifyContent: "center", fontSize: 18 }}>📂</div>
          <div>
            <div style={{ fontWeight: 700, fontSize: 17, letterSpacing: "-0.02em" }}>FileViewer</div>
            <div style={{ fontSize: 11, color: MUTED }}>PDF · Word · Excel · CSV · Images · Text</div>
          </div>
        </div>
        <div style={{ flex: 1 }} />
        <button
          className="upload-btn"
          onClick={() => inputRef.current.click()}
          style={{
            background: "transparent", color: ACCENT, border: `1.5px solid ${ACCENT}`,
            borderRadius: 8, padding: "8px 18px", cursor: "pointer", fontSize: 13,
            fontWeight: 600, transition: "all 0.15s", fontFamily: "inherit"
          }}
        >
          + Open Files
        </button>
        <input ref={inputRef} type="file" multiple style={{ display: "none" }}
          accept=".pdf,.doc,.docx,.xls,.xlsx,.csv,.txt,.md,.json,.png,.jpg,.jpeg,.gif,.webp,.svg"
          onChange={e => { if (e.target.files.length) addFiles(e.target.files); e.target.value = ""; }}
        />
      </div>

      <div style={{ display: "flex", flex: 1, minHeight: 0 }}>
        {/* Sidebar */}
        <div style={{
          width: 240, borderRight: `1px solid ${BORDER}`, background: SURFACE,
          display: "flex", flexDirection: "column", overflowY: "auto", flexShrink: 0
        }}>
          {files.length === 0 ? (
            <div style={{ padding: 24, color: MUTED, fontSize: 13, textAlign: "center", marginTop: 24 }}>
              Open files to get started
            </div>
          ) : files.map((f, i) => (
            <div
              key={i}
              onClick={() => setActive(i)}
              style={{
                display: "flex", alignItems: "center", gap: 10, padding: "11px 14px",
                cursor: "pointer", borderBottom: `1px solid ${BORDER}`,
                background: i === active ? SURFACE2 : "transparent",
                borderLeft: i === active ? `3px solid ${ACCENT}` : "3px solid transparent",
                transition: "all 0.1s",
              }}
            >
              <div style={{ fontSize: 18, flexShrink: 0 }}>{FILE_TYPES[getExt(f.name)]?.icon ?? "📁"}</div>
              <div style={{ flex: 1, minWidth: 0 }}>
                <div style={{ fontSize: 13, fontWeight: 500, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{f.name}</div>
                <div style={{ fontSize: 11, color: MUTED }}>{(f.size / 1024).toFixed(1)} KB</div>
              </div>
              <button
                className="close-btn"
                onClick={e => { e.stopPropagation(); removeFile(i); }}
                style={{ background: "none", border: "none", color: MUTED, cursor: "pointer", fontSize: 16, lineHeight: 1, padding: 2, flexShrink: 0, transition: "color 0.15s" }}
              >×</button>
            </div>
          ))}
        </div>

        {/* Main content */}
        <div style={{ flex: 1, display: "flex", flexDirection: "column", overflow: "hidden" }}>
          {currentFile ? (
            <>
              {/* File header */}
              <div style={{ padding: "14px 24px", borderBottom: `1px solid ${BORDER}`, display: "flex", alignItems: "center", gap: 12, background: SURFACE }}>
                <div style={{ fontSize: 22 }}>{FILE_TYPES[getExt(currentFile.name)]?.icon ?? "📁"}</div>
                <div>
                  <div style={{ fontWeight: 600, fontSize: 15 }}>{currentFile.name}</div>
                  <div style={{ fontSize: 12, color: MUTED }}>{(currentFile.size / 1024).toFixed(2)} KB</div>
                </div>
                <FileBadge ext={getExt(currentFile.name)} />
              </div>
              {/* Content area */}
              <div style={{ flex: 1, overflowY: "auto", padding: 28 }}>
                <FileViewer key={currentFile.name + currentFile.size} file={currentFile} />
              </div>
            </>
          ) : (
            // Drop zone
            <div
              className={`drop-zone${dragging ? " drag-over" : ""}`}
              onDragOver={e => { e.preventDefault(); setDragging(true); }}
              onDragLeave={() => setDragging(false)}
              onDrop={onDrop}
              style={{
                flex: 1, display: "flex", flexDirection: "column", alignItems: "center",
                justifyContent: "center", gap: 20, cursor: "pointer",
                border: `2px dashed ${BORDER}`, margin: 32, borderRadius: 16,
                transition: "all 0.2s", padding: 48,
              }}
              onClick={() => inputRef.current.click()}
            >
              <div style={{ fontSize: 56 }}>📂</div>
              <div style={{ textAlign: "center" }}>
                <div style={{ fontSize: 20, fontWeight: 700, marginBottom: 8 }}>Drop files here</div>
                <div style={{ fontSize: 14, color: MUTED, lineHeight: 1.7 }}>
                  or click to browse<br />
                  PDF · Word · Excel · CSV · JSON · Text · Images
                </div>
              </div>
              <div style={{ display: "flex", gap: 8, flexWrap: "wrap", justifyContent: "center" }}>
                {["PDF", "DOCX", "XLSX", "CSV", "TXT", "JSON", "PNG"].map(t => (
                  <span key={t} style={{ background: SURFACE2, border: `1px solid ${BORDER}`, borderRadius: 6, padding: "4px 10px", fontSize: 11, fontWeight: 600, color: MUTED, fontFamily: "monospace" }}>{t}</span>
                ))}
              </div>
            </div>
          )}
        </div>
      </div>
    </div>
  );
                                        }
