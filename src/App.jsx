import { useState, useRef, useCallback, useEffect } from "react";
import * as mammoth from "mammoth";
import * as XLSX from "xlsx";
import * as pdfjsLib from "pdfjs-dist";

pdfjsLib.GlobalWorkerOptions.workerSrc =
  "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js";

const C = {
  bg: "#0a0a0f",
  surface: "#13131a",
  card: "#1a1a24",
  border: "#252530",
  accent: "#f97316",
  accentDim: "#f9731622",
  text: "#f0f0f8",
  muted: "#5a5a70",
  green: "#22c55e",
  blue: "#3b82f6",
  red: "#ef4444",
  purple: "#a855f7",
  yellow: "#f59e0b",
  pink: "#ec4899",
};

const FILE_META = {
  pdf: { label: "PDF", color: C.red, icon: "📄" },
  docx: { label: "WORD", color: C.blue, icon: "📝" },
  doc: { label: "WORD", color: C.blue, icon: "📝" },
  xlsx: { label: "EXCEL", color: C.green, icon: "📊" },
  xls: { label: "EXCEL", color: C.green, icon: "📊" },
  csv: { label: "CSV", color: C.green, icon: "📋" },
  txt: { label: "TXT", color: C.purple, icon: "📃" },
  md: { label: "MD", color: C.yellow, icon: "✍️" },
  json: { label: "JSON", color: "#06b6d4", icon: "🔧" },
  png: { label: "IMG", color: C.pink, icon: "🖼️" },
  jpg: { label: "IMG", color: C.pink, icon: "🖼️" },
  jpeg: { label: "IMG", color: C.pink, icon: "🖼️" },
  gif: { label: "IMG", color: C.pink, icon: "🖼️" },
  webp: { label: "IMG", color: C.pink, icon: "🖼️" },
  svg: { label: "SVG", color: C.accent, icon: "🎨" },
};

const getExt = (n) => n.split(".").pop().toLowerCase();

const fmtSize = (b) =>
  b > 1048576
    ? (b / 1048576).toFixed(1) + " MB"
    : (b / 1024).toFixed(0) + " KB";

function Spinner({ label = "Loading…" }) {
  return (
    <div
      style={{
        display: "flex",
        flexDirection: "column",
        alignItems: "center",
        justifyContent: "center",
        gap: 16,
        padding: 64,
      }}
    >
      <div
        style={{
          width: 44,
          height: 44,
          border: `3px solid ${C.border}`,
          borderTop: `3px solid ${C.accent}`,
          borderRadius: "50%",
          animation: "spin .8s linear infinite",
        }}
      />
      <span style={{ color: C.muted, fontSize: 13 }}>{label}</span>
    </div>
  );
}

function PDFViewer({ file }) {
  const [pages, setPages] = useState([]);
  const [loading, setLoading] = useState(true);

  useEffect(() => {
    let dead = false;

    (async () => {
      try {
        const pdf = await pdfjsLib.getDocument({
          data: await file.arrayBuffer(),
        }).promise;

        const imgs = [];

        for (let i = 1; i <= Math.min(pdf.numPages, 20); i++) {
          const page = await pdf.getPage(i);

          const viewport = page.getViewport({
            scale: window.innerWidth < 500 ? 1.4 : 2,
          });

          const canvas = document.createElement("canvas");

          canvas.width = viewport.width;
          canvas.height = viewport.height;

          await page.render({
            canvasContext: canvas.getContext("2d"),
            viewport,
          }).promise;

          imgs.push({
            src: canvas.toDataURL(),
            num: i,
          });
        }

        if (!dead) {
          setPages(imgs);
          setLoading(false);
        }
      } catch {
        if (!dead) setLoading(false);
      }
    })();

    return () => {
      dead = true;
    };
  }, [file]);

  if (loading) return <Spinner label="Rendering PDF…" />;

  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 12 }}>
      {pages.map((p) => (
        <div key={p.num} style={{ position: "relative" }}>
          <div
            style={{
              position: "absolute",
              top: 10,
              left: 10,
              background: "#000c",
              color: "#fff",
              borderRadius: 6,
              padding: "3px 10px",
              fontSize: 12,
              fontWeight: 700,
              zIndex: 1,
            }}
          >
            p.{p.num}
          </div>

          <img
            src={p.src}
            alt=""
            style={{
              width: "100%",
              borderRadius: 12,
              display: "block",
              border: `1px solid ${C.border}`,
            }}
          />
        </div>
      ))}
    </div>
  );
}

function DocxViewer({ file }) {
  const [html, setHtml] = useState("");
  const [loading, setLoading] = useState(true);

  useEffect(() => {
    let dead = false;

    (async () => {
      try {
        const res = await mammoth.convertToHtml({
          arrayBuffer: await file.arrayBuffer(),
        });

        if (!dead) {
          setHtml(res.value);
          setLoading(false);
        }
      } catch {
        if (!dead) setLoading(false);
      }
    })();

    return () => {
      dead = true;
    };
  }, [file]);

  if (loading) return <Spinner label="Parsing document…" />;

  return (
    <div
      style={{
        background: "#fff",
        borderRadius: 12,
        padding: 20,
        color: "#111",
        lineHeight: 1.8,
        overflowX: "auto",
      }}
      dangerouslySetInnerHTML={{
        __html: html || "<p>Empty document.</p>",
      }}
    />
  );
}

function ExcelViewer({ file }) {
  const [rows, setRows] = useState([]);
  const [loading, setLoading] = useState(true);

  useEffect(() => {
    let dead = false;

    (async () => {
      try {
        const wb = XLSX.read(await file.arrayBuffer(), {
          type: "array",
        });

        const sheet = wb.Sheets[wb.SheetNames[0]];

        const data = XLSX.utils.sheet_to_json(sheet, {
          header: 1,
        });

        if (!dead) {
          setRows(data);
          setLoading(false);
        }
      } catch {
        if (!dead) setLoading(false);
      }
    })();

    return () => {
      dead = true;
    };
  }, [file]);

  if (loading) return <Spinner label="Parsing spreadsheet…" />;

  return (
    <div
      style={{
        overflowX: "auto",
        borderRadius: 10,
        border: `1px solid ${C.border}`,
        WebkitOverflowScrolling: "touch",
      }}
    >
      <table
        style={{
          borderCollapse: "collapse",
          minWidth: "100%",
          fontSize: 13,
        }}
      >
        <tbody>
          {rows.map((row, ri) => (
            <tr key={ri}>
              {row.map((cell, ci) => (
                <td
                  key={ci}
                  style={{
                    padding: "10px 14px",
                    border: `1px solid ${C.border}`,
                    color: C.text,
                    whiteSpace: "nowrap",
                    background: ri === 0 ? C.card : C.surface,
                    fontWeight: ri === 0 ? 700 : 400,
                  }}
                >
                  {cell}
                </td>
              ))}
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
}

function TextViewer({ file }) {
  const [text, setText] = useState("");

  useEffect(() => {
    file.text().then(setText);
  }, [file]);

  return (
    <pre
      style={{
        background: C.card,
        padding: 16,
        borderRadius: 12,
        overflowX: "auto",
        whiteSpace: "pre-wrap",
        wordBreak: "break-word",
        fontSize: 13,
        lineHeight: 1.7,
      }}
    >
      {text}
    </pre>
  );
}

function ImageViewer({ file }) {
  const [src, setSrc] = useState("");

  useEffect(() => {
    const url = URL.createObjectURL(file);

    setSrc(url);

    return () => URL.revokeObjectURL(url);
  }, [file]);

  return (
    <img
      src={src}
      alt=""
      style={{
        width: "100%",
        borderRadius: 12,
        display: "block",
      }}
    />
  );
}

function FileContent({ file }) {
  const ext = getExt(file.name);

  if (["png", "jpg", "jpeg", "gif", "webp", "svg"].includes(ext))
    return <ImageViewer file={file} />;

  if (ext === "pdf") return <PDFViewer file={file} />;

  if (["docx", "doc"].includes(ext))
    return <DocxViewer file={file} />;

  if (["xlsx", "xls", "csv"].includes(ext))
    return <ExcelViewer file={file} />;

  return <TextViewer file={file} />;
}

export default function App() {
  const [files, setFiles] = useState([]);
  const [active, setActive] = useState(null);
  const [view, setView] = useState("home");
  const [dragging, setDragging] = useState(false);

  const inputRef = useRef();

  const addFiles = useCallback((newFiles) => {
    const arr = Array.from(newFiles);

    setFiles((prev) => [
      ...prev,
      ...arr.filter(
        (f) => !prev.find((p) => p.name === f.name && p.size === f.size)
      ),
    ]);

    setView("files");
  }, []);

  const current = active !== null ? files[active] : null;

  return (
    <div
      style={{
        minHeight: "100dvh",
        background: C.bg,
        color: C.text,
        fontFamily: "'DM Sans','Segoe UI',sans-serif",
        width: "100%",
        overflowX: "hidden",
        display: "flex",
        justifyContent: "center",
      }}
    >
      <style>{`
        *{
          box-sizing:border-box;
          margin:0;
          padding:0;
          -webkit-tap-highlight-color:transparent;
        }

        html,body,#root{
          width:100%;
          min-height:100%;
          background:${C.bg};
          overflow-x:hidden;
        }

        body{
          touch-action:manipulation;
          -webkit-font-smoothing:antialiased;
        }

        button{
          font:inherit;
        }

        img{
          max-width:100%;
          height:auto;
        }

        @keyframes spin{
          to{
            transform:rotate(360deg)
          }
        }

        @keyframes fadeUp{
          from{
            opacity:0;
            transform:translateY(20px)
          }
          to{
            opacity:1;
            transform:translateY(0)
          }
        }

        .tap:active{
          transform:scale(.97);
          opacity:.7;
        }

        ::-webkit-scrollbar{
          width:3px;
          height:3px;
        }

        ::-webkit-scrollbar-thumb{
          background:${C.border};
        }
      `}</style>

      <input
        ref={inputRef}
        type="file"
        multiple
        hidden
        onChange={(e) => addFiles(e.target.files)}
      />

      <div
        style={{
          width: "100%",
          maxWidth: 480,
          minHeight: "100dvh",
          display: "flex",
          flexDirection: "column",
        }}
      >
        {view === "home" && (
          <div
            style={{
              flex: 1,
              display: "flex",
              flexDirection: "column",
              padding: 20,
              animation: "fadeUp .3s ease",
            }}
          >
            <div style={{ textAlign: "center", marginTop: 50 }}>
              <div
                style={{
                  width: 90,
                  height: 90,
                  borderRadius: 28,
                  background: `linear-gradient(135deg,${C.accent},#fb923c)`,
                  display: "flex",
                  alignItems: "center",
                  justifyContent: "center",
                  fontSize: 42,
                  margin: "0 auto 22px",
                }}
              >
                📂
              </div>

              <h1
                style={{
                  fontSize: window.innerWidth < 400 ? 24 : 30,
                  fontWeight: 800,
                  marginBottom: 8,
                }}
              >
                FileViewer
              </h1>

              <p
                style={{
                  color: C.muted,
                  fontSize: 14,
                  lineHeight: 1.7,
                }}
              >
                Open and preview files directly on mobile
              </p>
            </div>

            <div
              className="tap"
              onClick={() => inputRef.current.click()}
              onDragOver={(e) => {
                e.preventDefault();
                setDragging(true);
              }}
              onDragLeave={() => setDragging(false)}
              onDrop={(e) => {
                e.preventDefault();
                setDragging(false);

                if (e.dataTransfer.files.length) {
                  addFiles(e.dataTransfer.files);
                }
              }}
              style={{
                marginTop: 40,
                border: `2px dashed ${
                  dragging ? C.accent : C.border
                }`,
                borderRadius: 20,
                padding:
                  window.innerWidth < 400
                    ? "32px 18px"
                    : "44px 24px",
                textAlign: "center",
                background: dragging
                  ? C.accentDim
                  : "transparent",
                cursor: "pointer",
              }}
            >
              <div style={{ fontSize: 42, marginBottom: 12 }}>
                📁
              </div>

              <div
                style={{
                  fontSize: 18,
                  fontWeight: 700,
                  marginBottom: 6,
                }}
              >
                Tap to open files
              </div>

              <div
                style={{
                  color: C.muted,
                  fontSize: 13,
                }}
              >
                PDF · Word · Excel · Images
              </div>
            </div>

            {files.length > 0 && (
              <button
                className="tap"
                onClick={() => setView("files")}
                style={{
                  marginTop: 20,
                  background: C.card,
                  border: `1px solid ${C.border}`,
                  color: C.text,
                  borderRadius: 14,
                  padding: 16,
                  fontWeight: 700,
                  cursor: "pointer",
                }}
              >
                View {files.length} file
                {files.length > 1 ? "s" : ""}
              </button>
            )}
          </div>
        )}

        {view === "files" && (
          <div
            style={{
              flex: 1,
              padding: 16,
              animation: "fadeUp .3s ease",
            }}
          >
            <div
              style={{
                display: "flex",
                alignItems: "center",
                marginBottom: 18,
                gap: 10,
              }}
            >
              <button
                className="tap"
                onClick={() => setView("home")}
                style={{
                  width: 42,
                  height: 42,
                  borderRadius: 12,
                  border: `1px solid ${C.border}`,
                  background: C.card,
                  color: C.text,
                  cursor: "pointer",
                }}
              >
                ←
              </button>

              <div style={{ flex: 1 }}>
                <div
                  style={{
                    fontWeight: 800,
                    fontSize: 20,
                  }}
                >
                  My Files
                </div>

                <div
                  style={{
                    color: C.muted,
                    fontSize: 12,
                  }}
                >
                  {files.length} files
                </div>
              </div>
            </div>

            {files.map((f, i) => {
              const ext = getExt(f.name);
              const meta =
                FILE_META[ext] || {
                  icon: "📁",
                  color: C.muted,
                  label: ext.toUpperCase(),
                };

              return (
                <div
                  key={i}
                  className="tap"
                  onClick={() => {
                    setActive(i);
                    setView("viewer");
                  }}
                  style={{
                    display: "flex",
                    alignItems: "center",
                    gap: 12,
                    padding: 14,
                    background: C.card,
                    borderRadius: 16,
                    marginBottom: 12,
                    border: `1px solid ${C.border}`,
                    cursor: "pointer",
                  }}
                >
                  <div
                    style={{
                      width: 52,
                      height: 52,
                      borderRadius: 14,
                      background: meta.color + "22",
                      display: "flex",
                      alignItems: "center",
                      justifyContent: "center",
                      fontSize: 26,
                      flexShrink: 0,
                    }}
                  >
                    {meta.icon}
                  </div>

                  <div
                    style={{
                      flex: 1,
                      minWidth: 0,
                    }}
                  >
                    <div
                      style={{
                        fontWeight: 600,
                        overflow: "hidden",
                        textOverflow: "ellipsis",
                        whiteSpace: "nowrap",
                      }}
                    >
                      {f.name}
                    </div>

                    <div
                      style={{
                        marginTop: 6,
                        display: "flex",
                        alignItems: "center",
                        gap: 8,
                      }}
                    >
                      <span
                        style={{
                          background: meta.color + "22",
                          color: meta.color,
                          borderRadius: 5,
                          padding: "2px 7px",
                          fontSize: 10,
                          fontWeight: 700,
                        }}
                      >
                        {meta.label}
                      </span>

                      <span
                        style={{
                          color: C.muted,
                          fontSize: 12,
                        }}
                      >
                        {fmtSize(f.size)}
                      </span>
                    </div>
                  </div>
                </div>
              );
            })}
          </div>
        )}

        {view === "viewer" && current && (
          <div
            style={{
              flex: 1,
              display: "flex",
              flexDirection: "column",
            }}
          >
            <div
              style={{
                padding: 14,
                borderBottom: `1px solid ${C.border}`,
                display: "flex",
                alignItems: "center",
                gap: 10,
                background: C.surface,
                position: "sticky",
                top: 0,
                zIndex: 20,
              }}
            >
              <button
                className="tap"
                onClick={() => setView("files")}
                style={{
                  width: 40,
                  height: 40,
                  borderRadius: 10,
                  border: `1px solid ${C.border}`,
                  background: C.card,
                  color: C.text,
                  cursor: "pointer",
                }}
              >
                ←
              </button>

              <div
                style={{
                  flex: 1,
                  minWidth: 0,
                }}
              >
                <div
                  style={{
                    fontWeight: 700,
                    overflow: "hidden",
                    textOverflow: "ellipsis",
                    whiteSpace: "nowrap",
                  }}
                >
                  {current.name}
                </div>

                <div
                  style={{
           
