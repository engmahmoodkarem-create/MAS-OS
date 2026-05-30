import { basename, extname } from "path";
import { memo, type FC, useCallback, useEffect, useRef, useState } from "react";
import StyledQuickView from "components/apps/QuickView/StyledQuickView";
import { type ComponentProcessProps } from "components/system/Apps/RenderComponent";
import { useProcesses } from "contexts/process";
import { useFileSystem } from "contexts/fileSystem";

type ViewType = "image" | "sheet" | "doc" | "slides" | "text" | "pdf" | "unknown";

interface ISheetData {
  name: string;
  rows: string[][];
}

interface ISlideData {
  body: string;
  title: string;
}

const parseDocx = async (buffer: Buffer): Promise<string> => {
  const { default: JSZip } = await import("jszip");
  const zip = await JSZip.loadAsync(buffer);
  const docXmlText = await zip.file("word/document.xml")?.async("text");

  if (!docXmlText) return "<p>Empty Document</p>";

  const parser = new DOMParser();
  const xmlDoc = parser.parseFromString(docXmlText, "application/xml");
  const paragraphs = xmlDoc.querySelectorAll(String.raw`w\:p`);
  let html = "";

  for (const p of paragraphs) {
    let pText = "";
    const textNodes = p.querySelectorAll(String.raw`w\:t`);

    for (const node of textNodes) {
      pText += node.textContent || "";
    }

    if (pText.trim()) {
      const headingEl = p.querySelector(String.raw`w\:pStyle`);
      const isHeading = headingEl?.getAttribute("w:val")?.startsWith("Heading");

      html += isHeading ? `<h2>${pText}</h2>` : `<p>${pText}</p>`;
    }
  }

  return html || "<p>Empty document or unsupported Word layout</p>";
};

const parsePptx = async (buffer: Buffer): Promise<ISlideData[]> => {
  const { default: JSZip } = await import("jszip");
  const zip = await JSZip.loadAsync(buffer);
  const slidesList: ISlideData[] = [];

  const slideFiles = Object.keys(zip.files).filter((path) =>
    path.startsWith("ppt/slides/slide") && path.endsWith(".xml")
  );

  slideFiles.sort((a, b) => {
    const numA = Number((/\d+/).exec(a)?.[0] || 0);
    const numB = Number((/\d+/).exec(b)?.[0] || 0);
    return numA - numB;
  });

  const parser = new DOMParser();

  const slideDataList = await Promise.all(
    slideFiles.map(async (slidePath) => {
      const xmlText = await zip.file(slidePath)?.async("text");
      if (!xmlText) return null; // eslint-disable-line unicorn/no-null

      const xmlDoc = parser.parseFromString(xmlText, "application/xml");
      const textElements = xmlDoc.querySelectorAll(String.raw`a\:t`);
      let title = "";
      const bodyLines: string[] = [];

      for (const el of textElements) {
        const textVal = el.textContent || "";
        if (textVal.trim()) {
          if (!title && textVal.length < 40) {
            title = textVal;
          } else {
            bodyLines.push(textVal);
          }
        }
      }

      return {
        body: bodyLines.join("\n"),
        title: title || `Slide ${slidesList.length + 1}`,
      };
    })
  );

  for (const data of slideDataList) {
    if (data) slidesList.push(data);
  }

  if (slidesList.length === 0) {
    slidesList.push({
      body: "No slide slides found in presentation archive.",
      title: "Presentation",
    });
  }

  return slidesList;
};

const parseXlsx = async (buffer: Buffer): Promise<ISheetData[]> => {
  const XLSX = await import("xlsx");
  const workbook = XLSX.read(buffer, { type: "array" });
  const sheetsList: ISheetData[] = [];

  for (const name of workbook.SheetNames) {
    const worksheet = workbook.Sheets[name];
    const rows = XLSX.utils.sheet_to_json<unknown[]>(worksheet, { header: 1 });
    sheetsList.push({
      name,
      rows: rows.map((r) => r.map((c) => {
        if (c === undefined) return "";
        if (c === null) return "";
        if (typeof c === "object") return JSON.stringify(c);
        if (typeof c === "string") return c;

        return String(c); // eslint-disable-line @typescript-eslint/no-base-to-string
      })),
    });
  }

  return sheetsList;
};

const QuickView: FC<ComponentProcessProps> = ({ id }) => {
  const { processes: { [id]: process } = {} } = useProcesses();
  const { url } = process || {};
  const { readFile } = useFileSystem();

  const [loading, setLoading] = useState(true);
  const [viewType, setViewType] = useState<ViewType>("unknown");
  const [fileName, setFileName] = useState("");
  const [fileUrl, setFileUrl] = useState("");

  // Views Data
  const [sheets, setSheets] = useState<ISheetData[]>([]);
  const [activeSheet, setActiveSheet] = useState(0);
  const [docHTML, setDocHTML] = useState("");
  const [slides, setSlides] = useState<ISlideData[]>([]);
  const [activeSlide, setActiveSlide] = useState(0);
  const [textContent, setTextContent] = useState("");

  const loadTokenRef = useRef(0);

  const loadFile = useCallback(async (): Promise<void> => {
    if (!url) return;
    setLoading(true);
    const token = ++loadTokenRef.current;

    const name = basename(url);
    setFileName(name);

    try {
      const buffer = await readFile(url);
      if (token !== loadTokenRef.current) return;

      const ext = extname(url).toLowerCase();

      // Convert buffer to object URL for native elements
      const blob = new Blob([new Uint8Array(buffer)]);
      const blobUrl = URL.createObjectURL(blob);
      setFileUrl(blobUrl);

      if ([".png", ".jpg", ".jpeg", ".gif", ".svg", ".webp"].includes(ext)) {
        setViewType("image");
      } else if ([".xlsx", ".xls"].includes(ext)) {
        const sheetsData = await parseXlsx(buffer);
        if (token !== loadTokenRef.current) return;
        setSheets(sheetsData);
        setActiveSheet(0);
        setViewType("sheet");
      } else if ([".docx", ".doc"].includes(ext)) {
        if (ext === ".docx") {
          const html = await parseDocx(buffer);
          if (token !== loadTokenRef.current) return;
          setDocHTML(html);
        } else {
          setDocHTML(`<p>Word .doc formats not fully supported. Please view .docx files.</p>`);
        }
        setViewType("doc");
      } else if ([".pptx", ".ppt"].includes(ext)) {
        if (ext === ".pptx") {
          const parsedSlides = await parsePptx(buffer);
          if (token !== loadTokenRef.current) return;
          setSlides(parsedSlides);
          setActiveSlide(0);
        } else {
          setSlides([{ body: "PowerPoint .ppt formats not fully supported. Please view .pptx files.", title: "Unsupported Presentation" }]);
        }
        setViewType("slides");
      } else if ([".pdf"].includes(ext)) {
        setViewType("pdf");
      } else if ([".txt", ".json", ".xml", ".js", ".ts", ".css", ".html", ".md"].includes(ext)) {
        const text = new TextDecoder().decode(buffer);
        setTextContent(text);
        setViewType("text");
      } else {
        setViewType("unknown");
      }
    } catch {
      if (token === loadTokenRef.current) {
        setViewType("unknown");
      }
    } finally {
      if (token === loadTokenRef.current) {
        setLoading(false);
      }
    }
  }, [readFile, url]);

  useEffect(() => {
    loadFile();

    return () => {
      if (fileUrl) {
        URL.revokeObjectURL(fileUrl);
      }
    };
  // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [loadFile]);

  const renderContent = (): React.ReactNode => {
    if (loading) {
      return (
        <div className="quick-loading">
          <div className="spinner" />
          <span>Loading file...</span>
        </div>
      );
    }

    switch (viewType) {
      case "image":
        return (
          <div className="image-container">
            <img alt={fileName} src={fileUrl} />
          </div>
        );

      case "pdf":
        return (
          <div className="iframe-container">
            <iframe src={fileUrl} title={fileName} />
          </div>
        );

      case "text":
        return (
          <pre className="text-container">
            <code>{textContent}</code>
          </pre>
        );

      case "sheet":
        return (
          <div className="sheet-container">
            <div className="sheet-tabs">
              {sheets.map((sheet) => (
                <button
                  key={sheet.name}
                  className={`sheet-tab${activeSheet === sheets.indexOf(sheet) ? " active" : ""}`}
                  onClick={() => setActiveSheet(sheets.indexOf(sheet))}
                  type="button"
                >
                  {sheet.name}
                </button>
              ))}
            </div>
            <div className="sheet-grid-wrapper">
              <table className="sheet-table">
                <tbody>
                  {sheets[activeSheet]?.rows.map((row, rowIndex) => (
                    // eslint-disable-next-line react/no-array-index-key
                    <tr key={rowIndex}>
                      <td>{rowIndex + 1}</td>
                      {row.map((cell, cellIndex) => (
                        // eslint-disable-next-line react/no-array-index-key
                        <td key={cellIndex}>{cell}</td>
                      ))}
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </div>
        );

      case "doc":
        return (
          <div className="doc-container">
            {/* eslint-disable-next-line react/no-danger */}
            <div dangerouslySetInnerHTML={{ __html: docHTML }} className="doc-page" />
          </div>
        );

      case "slides":
        return (
          <div className="slides-container">
            <div className="slides-sidebar">
              {slides.map((slide, index) => (
                <div
                  // eslint-disable-next-line react/no-array-index-key
                  key={index}
                  className={`slide-thumb${activeSlide === index ? " active" : ""}`}
                  onClick={() => setActiveSlide(index)}
                  onKeyDown={(e) => {
                    if (e.key === "Enter" || e.key === " ") {
                      setActiveSlide(index);
                    }
                  }}
                  role="button"
                  tabIndex={0}
                >
                  <div className="thumb-title">{slide.title}</div>
                  <div>Slide {index + 1}</div>
                </div>
              ))}
            </div>
            <div className="slide-main-stage">
              <div className="slide-card">
                <div className="slide-title-text">{slides[activeSlide]?.title}</div>
                <div className="slide-body-content">{slides[activeSlide]?.body}</div>
                <div className="slide-footer">
                  <span>PowerPoint Preview</span>
                  <span>Slide {activeSlide + 1} of {slides.length}</span>
                </div>
              </div>
            </div>
          </div>
        );

      default:
        return (
          <div className="quick-loading">
            <span>Unsupported file preview. Use &quot;Open With&quot; for another application.</span>
          </div>
        );
    }
  };

  return (
    <StyledQuickView>
      <div className="quick-toolbar">
        <div className="file-info">
          <svg fill="currentColor" height="16" viewBox="0 0 16 16" width="16">
            <path d="M4 0h5.5v1H4v14h8V4h-1.5V2.5H9.5V0z" />
          </svg>
          <span>{fileName || "QuickView Preview"}</span>
        </div>
        <div className="controls">
          {viewType === "slides" && (
            <>
              <button
                className="toolbar-btn"
                disabled={activeSlide === 0}
                onClick={() => setActiveSlide((p) => Math.max(0, p - 1))}
                type="button"
              >
                Previous
              </button>
              <button
                className="toolbar-btn"
                disabled={activeSlide === slides.length - 1}
                onClick={() => setActiveSlide((p) => Math.min(slides.length - 1, p + 1))}
                type="button"
              >
                Next
              </button>
            </>
          )}
        </div>
      </div>
      <div className={`quick-viewport${viewType === "text" ? " text-view" : ""}${viewType === "sheet" ? " sheet-view" : ""}`}>
        {renderContent()}
      </div>
    </StyledQuickView>
  );
};

export default memo(QuickView);
