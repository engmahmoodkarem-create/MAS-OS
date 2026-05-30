import { basename, dirname, extname, join } from "path";
import { memo, type FC, useCallback, useEffect, useRef, useState } from "react";
import { Document as Doc, Packer, Paragraph, TextRun } from "docx";
import JSZip from "jszip";
import * as XLSXLib from "xlsx";
import StyledMASOffice from "components/apps/MASOffice/StyledMASOffice";
import { type ComponentProcessProps } from "components/system/Apps/RenderComponent";
import { useProcesses } from "contexts/process";
import { useFileSystem } from "contexts/fileSystem";
import { DESKTOP_PATH } from "utils/constants";
import { blobToBuffer } from "utils/functions";

import "@univerjs/design/lib/index.css";
import "@univerjs/ui/lib/index.css";
import "@univerjs/docs-ui/lib/index.css";
import "@univerjs/slides-ui/lib/index.css";

type OfficeMode = "sheet" | "doc" | "slides";

interface IUniverAPI {
  createUnit: (type: number, data?: Record<string, unknown>) => void;
  dispose: () => void;
  exportDOCXBySnapshotAsync: (snapshot: Record<string, unknown>) => Promise<File | undefined>;
  exportXLSXBySnapshotAsync: (snapshot: Record<string, unknown>) => Promise<File | undefined>;
  getActiveDocument: () => { getId: () => string; getSnapshot: () => Record<string, unknown> } | undefined;
  getActiveWorkbook: () => { getId: () => string; save: () => Record<string, unknown> } | undefined;
  importDOCXToSnapshotAsync: (file: File | string) => Promise<Record<string, unknown> | undefined>;
  importXLSXToSnapshotAsync: (file: File | string) => Promise<Record<string, unknown> | undefined>;
}

interface IUniverPresets {
  IUniverInstanceService: unknown;
  LocaleType: { EN_US: string };
  UniverInstanceType: { UNIVER_DOC: number; UNIVER_SHEET: number; UNIVER_SLIDE: number };
  createUniver: (config: Record<string, unknown>) => {
    univer: {
      __getInjector?: () => { get: (identifier: unknown) => unknown };
      createUnit: (type: number, data?: Record<string, unknown>) => { getUnitId?: () => string } | void;
      dispose: () => void;
    };
    univerAPI: IUniverAPI;
  };
  mergeLocales: (...locales: unknown[]) => Record<string, unknown>;
}

type MenuEntry =
  | { type: "divider" }
  | { action: () => void; label: string; shortcut?: string; type: "item" };

const MODES: { id: OfficeMode; label: string }[] = [
  { id: "sheet", label: "Sheet" },
  { id: "doc", label: "Doc" },
  { id: "slides", label: "Slides" },
];

const SLIDE_DATA: Record<string, unknown> = {
  body: {
    pageOrder: ["page_1", "page_2"],
    pages: {
      page_1: {
        description: "",
        id: "page_1",
        pageBackgroundFill: { rgb: "rgb(0, 120, 212)" },
        pageElements: {
          subtitle1: {
            description: "",
            height: 60, id: "subtitle1", left: 80,
            richText: { cl: { rgb: "rgb(200, 220, 255)" }, fs: 18, text: "Create beautiful presentations" },
            title: "Subtitle",
            top: 300, type: 2, width: 800, zIndex: 2,
          },
          title1: {
            description: "",
            height: 100, id: "title1", left: 80,
            richText: { bl: 1, cl: { rgb: "rgb(255,255,255)" }, fs: 36, text: "Welcome to MAS Office" },
            title: "Title",
            top: 180, type: 2, width: 800, zIndex: 2,
          },
        },
        pageType: 0, title: "Welcome", zIndex: 1,
      },
      page_2: {
        description: "",
        id: "page_2",
        pageBackgroundFill: { rgb: "rgb(255,255,255)" },
        pageElements: {
          body2: {
            description: "",
            height: 300, id: "body2", left: 80,
            richText: { cl: { rgb: "rgb(51,51,51)" }, fs: 13.5,
              text: "Sheet - Create and edit spreadsheets\nDoc - Write and format documents\nSlides - Build presentations" },
            title: "Body",
            top: 160, type: 2, width: 800, zIndex: 2,
          },
          title2: {
            description: "",
            height: 60, id: "title2", left: 80,
            richText: { cl: { rgb: "rgb(0, 120, 212)" }, fs: 27, text: "Getting Started" },
            title: "Title",
            top: 60, type: 2, width: 800, zIndex: 2,
          },
        },
        pageType: 0, title: "Content", zIndex: 1,
      },
    },
  },
  id: "slide_1",
  pageSize: { height: 540, width: 960 },
  title: "Presentation",
};



interface ISlideElement {
  height?: number;
  left?: number;
  richText?: { fs?: number; rich?: { body?: { dataStream?: string } }; text?: string };
  shape?: { text?: string };
  top?: number;
  width?: number;
}

interface ISlidePageDef {
  pageBackgroundFill?: { rgb?: string };
  pageElements?: Record<string, ISlideElement>;
}

interface ISheetCell {
  v?: string | number | boolean | null;
}

interface ISheetSnapshot {
  cellData?: Record<string, Record<string, ISheetCell>>;
  name?: string;
}

interface ISheetWorkbookSnapshot {
  sheets?: Record<string, ISheetSnapshot>;
}

interface IDocSnapshot {
  body?: { dataStream?: string; text?: string };
}

const DEFAULT_SHEET_GRID = [
  ["", "A", "B", "C", "D", "E"],
  ["1", "", "", "", "", ""],
  ["2", "", "", "", "", ""],
  ["3", "", "", "", "", ""],
  ["4", "", "", "", "", ""],
  ["5", "", "", "", "", ""],
  ["6", "", "", "", "", ""],
  ["7", "", "", "", "", ""],
  ["8", "", "", "", "", ""],
  ["9", "", "", "", "", ""],
  ["10", "", "", "", "", ""],
  ["11", "", "", "", "", ""],
];

const DEFAULT_DOC_TEXT = "Untitled document\nStart typing here.";

const isSheetWorkbookSnapshot = (value: unknown): value is ISheetWorkbookSnapshot =>
  Boolean(value && typeof value === "object" && "sheets" in value);

const isDocSnapshot = (value: unknown): value is IDocSnapshot =>
  Boolean(value && typeof value === "object" && "body" in value);

const getColumnLabel = (index: number): string => {
  let value = index + 1;
  let label = "";

  while (value > 0) {
    const remainder = (value - 1) % 26;
    label = String.fromCodePoint(65 + remainder) + label;
    value = Math.floor((value - 1) / 26);
  }

  return label;
};

const snapshotToSheetGrid = (snapshot?: unknown): string[][] => {
  if (!isSheetWorkbookSnapshot(snapshot)) return DEFAULT_SHEET_GRID;

  const [sheet] = Object.values(snapshot.sheets || {});
  const cellData = sheet?.cellData || {};
  const rowIndexes = Object.keys(cellData).map(Number).filter((value) => Number.isFinite(value));

  if (rowIndexes.length === 0) return DEFAULT_SHEET_GRID;

  const maxRow = Math.max(...rowIndexes);
  const maxCol = Math.max(
    0,
    ...rowIndexes.flatMap((rowIndex) =>
      Object.keys(cellData[String(rowIndex)] || {}).map(Number).filter((value) => Number.isFinite(value))),
  );

  const grid = [["", ...Array.from({ length: maxCol + 1 }, (_col, colIndex) => getColumnLabel(colIndex))]];

  for (let row = 0; row <= maxRow; row++) {
    const cells = [String(row + 1)];

    for (let col = 0; col <= maxCol; col++) {
      cells.push(String(cellData[String(row)]?.[String(col)]?.v ?? ""));
    }

    grid.push(cells);
  }

  return grid;
};

const snapshotToDocText = (snapshot?: unknown): string => {
  if (!isDocSnapshot(snapshot)) return DEFAULT_DOC_TEXT;

  const { body } = snapshot;
  const rawText = body?.dataStream || body?.text;

  if (!rawText) return DEFAULT_DOC_TEXT;

  return rawText.replace(/\r/g, "\n").replace(/\n{2,}/g, "\n\n").trim() || DEFAULT_DOC_TEXT;
};

interface ICreatePptxBody {
  pageOrder?: string[];
  pageSize?: { height?: number; width?: number };
  pages?: Record<string, ISlidePageDef>;
}

const escapeXml = (s: string): string =>
  s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;").replace(/'/g, "&apos;");

const getPptEmu = (value: number, total: number): string =>
  String(Math.round((value / total) * 9144000));

const extractSlideText = (el: ISlideElement): string => {
  if (el.richText?.text) return escapeXml(el.richText.text);
  if (el.richText?.rich?.body?.dataStream) {
    return escapeXml(el.richText.rich.body.dataStream.replace(/\r/g, "\n"));
  }
  if (el.shape?.text) return escapeXml(el.shape.text);

  return "";
};

const createPptxBlob = async (slideBody: ICreatePptxBody): Promise<Blob> => {
  const zip = new JSZip();
  const pageOrder = slideBody.pageOrder || [];
  const pages = slideBody.pages || {};
  const slideCount = pageOrder.length || 1;

  // Build slide XMLs
  const slideXmls: string[] = [];

  for (let i = 0; i < slideCount; i++) {
    const pageId = pageOrder[i] || `page_${i + 1}`;
    const page = pages[pageId];
    const textBoxes: string[] = [];

    if (page) {
      const elements = page.pageElements || {};

      for (const el of Object.values(elements)) {
        const text = extractSlideText(el);

        if (text) {
          const leftEmu = getPptEmu(el.left || 0, 960);
          const topEmu = getPptEmu(el.top || 0, 540);
          const widthEmu = getPptEmu(el.width || 800, 960);
          const heightEmu = getPptEmu(el.height || 60, 540);
          const fontSize = el.richText?.fs || 18;

          textBoxes.push(`<p:sp>
  <p:nvSpPr><p:cNvPr id="${i * 100 + textBoxes.length + 1}" name="TextBox${textBoxes.length + 1}"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>
  <p:spPr><a:xfrm><a:off x="${leftEmu}" y="${topEmu}"/><a:ext cx="${widthEmu}" cy="${heightEmu}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>
  <p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr sz="${fontSize * 100}" /><a:t>${text}</a:t></a:r></a:p></p:txBody></p:sp>`);
        }
      }
    }

    slideXmls.push(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld>
    <p:spTree>
      <p:nvGrpSpPr><p:nvGrpSpPrPr/><p:cNvPr id="1" name=""/><p:nvPr/></p:nvGrpSpPr>
      <p:grpSpPr/>
      ${textBoxes.join("\n")}
    </p:spTree>
  </p:cSld>
</p:sld>`);
  }

  // Add files to ZIP
  // [Content_Types].xml
  let contentTypes = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
  <Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>
  <Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>
  <Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>`;

  for (let i = 1; i <= slideCount; i++) {
    contentTypes += `\n  <Override PartName="/ppt/slides/slide${i}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>`;
  }

  contentTypes += `\n  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
</Types>`;

  zip.file("[Content_Types].xml", contentTypes);

  // _rels/.rels
  zip.file("_rels/.rels", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>
</Relationships>`);

  // ppt/presentation.xml
  const slideRels = slideXmls.map((_, i) =>
    `<Relationship Id="rId${i + 2}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide${i + 1}.xml"/>`).join("\n    ");

  zip.file("ppt/presentation.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentation xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst>
  <p:sldIdLst>${slideXmls.map((_, i) => `<p:sldId id="${256 + i}" r:id="rId${i + 2}"/>`).join("")}</p:sldIdLst>
  <p:sldSz cx="${getPptEmu(slideBody.pageSize?.width || 960, 960)}" cy="${getPptEmu(slideBody.pageSize?.height || 540, 540)}"/>
  <p:notesSz cx="6858000" cy="9144000"/>
</p:presentation>`);

  // ppt/_rels/presentation.xml.rels
  zip.file("ppt/_rels/presentation.xml.rels", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster1.xml"/>
  ${slideRels}
</Relationships>`);

  // ppt/slides/ and their rels
  for (let i = 0; i < slideCount; i++) {
    zip.file(`ppt/slides/slide${i + 1}.xml`, slideXmls[i]);
    zip.file(`ppt/slides/_rels/slide${i + 1}.xml.rels`, `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
</Relationships>`);
  }

  // ppt/slideMasters/slideMaster1.xml
  zip.file("ppt/slideMasters/slideMaster1.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld><p:spTree><p:nvGrpSpPr><p:nvGrpSpPrPr/><p:cNvPr id="1" name=""/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/></p:spTree></p:cSld>
  <p:type>sldMaster</p:type>
</p:sldMaster>`);

  // ppt/slideMasters/_rels/slideMaster1.xml.rels
  zip.file("ppt/slideMasters/_rels/slideMaster1.xml.rels", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="../theme/theme1.xml"/>
</Relationships>`);

  // ppt/slideLayouts/slideLayout1.xml
  zip.file("ppt/slideLayouts/slideLayout1.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" type="blank">
  <p:cSld><p:spTree><p:nvGrpSpPr><p:nvGrpSpPrPr/><p:cNvPr id="1" name=""/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/></p:spTree></p:cSld>
</p:sldLayout>`);

  // ppt/slideLayouts/_rels/slideLayout1.xml.rels
  zip.file("ppt/slideLayouts/_rels/slideLayout1.xml.rels", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="../theme/theme1.xml"/>
</Relationships>`);

  // ppt/theme/theme1.xml - minimal theme
  zip.file("ppt/theme/theme1.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Default">
  <a:themeElements>
    <a:clrScheme name="Default"><a:dk1><a:srgbClr val="000000"/></a:dk1><a:lt1><a:srgbClr val="FFFFFF"/></a:lt1><a:dk2><a:srgbClr val="44546A"/></a:dk2><a:lt2><a:srgbClr val="E7E6E6"/></a:lt2><a:accent1><a:srgbClr val="4472C4"/></a:accent1><a:accent2><a:srgbClr val="ED7D31"/></a:accent2><a:accent3><a:srgbClr val="A5A5A5"/></a:accent3><a:accent4><a:srgbClr val="FFC000"/></a:accent4><a:accent5><a:srgbClr val="5B9BD5"/></a:accent5><a:accent6><a:srgbClr val="70AD47"/></a:accent6><a:hlink><a:srgbClr val="0563C1"/></a:hlink><a:folHlink><a:srgbClr val="954F72"/></a:folHlink></a:clrScheme>
    <a:fontScheme name="Default"><a:majorFont><a:latin typeface="Calibri Light"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont><a:minorFont><a:latin typeface="Calibri"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont></a:fontScheme>
    <a:fmtScheme name="Default"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:fillStyleLst><a:lnStyleLst><a:ln w="6350"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln></a:lnStyleLst><a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst><a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:bgFillStyleLst></a:fmtScheme>
  </a:themeElements>
</a:theme>`);

  // docProps/core.xml
  zip.file("docProps/core.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/">
  <dc:creator>MAS Office</dc:creator>
  <dc:title>Presentation</dc:title>
  <dcterms:created xsi:type="dcterms:W3CDTF">${new Date().toISOString()}</dcterms:created>
</cp:coreProperties>`);

  // docProps/app.xml
  zip.file("docProps/app.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties">
  <Application>MAS Office</Application>
  <SlideCount>${slideCount}</SlideCount>
</Properties>`);

  const blob = await zip.generateAsync({ mimeType: "application/vnd.openxmlformats-officedocument.presentationml.presentation", type: "blob" });

  return blob;
};

const waitForDeferredDispose = (): Promise<void> =>
  new Promise((resolve) => {
    window.setTimeout(resolve, 0);
  });

const renderSheetFallback = (ref: React.RefObject<HTMLDivElement | null>, grid: string[][]): React.ReactNode => (
  <div ref={ref} className="sheet-fallback">
    {grid.map((row) => {
      const rowLabel = row[0] || "header";

      return (
        <div key={rowLabel} className="sheet-row">
          {row.map((value, columnIndex) => {
            const columnLabel = columnIndex === 0 ? "row" : getColumnLabel(columnIndex - 1);

            return (
              <div
                key={`${rowLabel}-${columnLabel}`}
                className={columnIndex === 0 ? "sheet-heading" : "sheet-cell"}
                contentEditable={columnIndex !== 0}
                suppressContentEditableWarning
              >
                {value}
              </div>
            );
          })}
        </div>
      );
    })}
  </div>
);

const renderDocFallback = (ref: React.RefObject<HTMLDivElement | null>, text: string): React.ReactNode => (
  <div ref={ref} className="doc-fallback" contentEditable suppressContentEditableWarning>
    {text.split(/\n{2,}/).map((paragraph) => (
      <p key={`${paragraph}-${paragraph.length}`}>{paragraph}</p>
    ))}
  </div>
);

const MASOffice: FC<ComponentProcessProps> = ({ id }) => {
  const [mode, setMode] = useState<OfficeMode>("sheet");
  const [activeMenu, setActiveMenu] = useState<string | undefined>();
  const [activeExportMenu, setActiveExportMenu] = useState(false);
  const [status, setStatus] = useState("Ready");
  const [sheetGrid, setSheetGrid] = useState<string[][]>(DEFAULT_SHEET_GRID);
  const [docText, setDocText] = useState(DEFAULT_DOC_TEXT);
  const containerRef = useRef<HTMLDivElement>(null);
  const sheetFallbackRef = useRef<HTMLDivElement>(null);
  const docFallbackRef = useRef<HTMLDivElement>(null);
  const fileInputRef = useRef<HTMLInputElement>(null);
  const univerRef = useRef<{ destroy: () => void } | undefined>(undefined);
  const apiRef = useRef<IUniverAPI | undefined>(undefined);
  const initTokenRef = useRef(0);
  const pendingTimersRef = useRef<number[]>([]);
  const menuBarRef = useRef<HTMLDivElement>(null);
  const exportBtnRef = useRef<HTMLDivElement>(null);
  const { processes: { [id]: process } = {} } = useProcesses();
  const { url } = process || {};
  const { readFile, updateFolder, writeFile } = useFileSystem();

  const setReadyStatus = useCallback((message: string): void => {
    setStatus(message);
    window.setTimeout(() => setStatus("Ready"), 2500);
  }, []);

  const getSaveTarget = useCallback((fallbackName: string, extension: string): { dir: string; name: string; path: string } => {
    const dir = url ? dirname(url) : DESKTOP_PATH;
    const name = url ? basename(url, extname(url)) : fallbackName;
    const fileName = `${name}.${extension}`;

    return { dir, name: fileName, path: join(dir, fileName) };
  }, [url]);

  const writeDesktopFile = useCallback(async (fallbackName: string, extension: string, data: Buffer | string): Promise<string> => {
    const target = getSaveTarget(fallbackName, extension);

    await writeFile(target.path, data, true);
    updateFolder(target.dir, target.name);
    setReadyStatus(`Saved ${target.name} to Desktop`);

    return target.path;
  }, [getSaveTarget, setReadyStatus, updateFolder, writeFile]);

  const clearPendingTimers = useCallback((): void => {
    pendingTimersRef.current.forEach((timerId) => window.clearTimeout(timerId));
    pendingTimersRef.current = [];
  }, []);

  const scheduleTask = useCallback((task: () => void, delay: number): number => {
    const timerId = window.setTimeout(() => {
      pendingTimersRef.current = pendingTimersRef.current.filter((currentTimerId) => currentTimerId !== timerId);
      task();
    }, delay);

    pendingTimersRef.current.push(timerId);

    return timerId;
  }, []);

  const scheduleIfCurrent = useCallback((token: number, task: () => void, delay: number): number =>
    scheduleTask(() => {
      if (token === initTokenRef.current) task();
    }, delay), [scheduleTask]);

  const invalidatePendingWork = useCallback((): number => {
    initTokenRef.current += 1;
    clearPendingTimers();

    return initTokenRef.current;
  }, [clearPendingTimers]);

  const destroyUniver = useCallback((invalidate = true): void => {
    if (invalidate) {
      invalidatePendingWork();
    }

    if (!univerRef.current) return;

    const { destroy } = univerRef.current;
    univerRef.current = undefined;
    apiRef.current = undefined;

    window.setTimeout(() => {
      try {
        destroy();
      } catch {
        /* empty */
      }
    }, 0);
  }, [invalidatePendingWork]);

  const initSlides = useCallback(async (token: number): Promise<void> => {
    if (!containerRef.current) return;

    const [
      presetsMod,
      { UniverSlidesPlugin },
      { UniverSlidesUIPlugin },
      { UniverRenderEnginePlugin },
      { UniverUIPlugin },
      { UniverDrawingPlugin },
      { UniverDocsPlugin },
      { UniverDocsUIPlugin },
      designEnUS, uiEnUS, docsUIEnUS, slidesUIEnUS,
    ] = await Promise.all([
      import("@univerjs/presets") as Promise<unknown>,
      import("@univerjs/slides") as Promise<{ UniverSlidesPlugin: new () => unknown }>,
      import("@univerjs/slides-ui") as Promise<{ UniverSlidesUIPlugin: new () => unknown }>,
      import("@univerjs/engine-render") as Promise<{ UniverRenderEnginePlugin: new () => unknown }>,
      import("@univerjs/ui") as Promise<{ UniverUIPlugin: new (...args: unknown[]) => unknown }>,
      import("@univerjs/drawing") as Promise<{ UniverDrawingPlugin: new () => unknown }>,
      import("@univerjs/docs") as Promise<{ UniverDocsPlugin: new () => unknown }>,
      import("@univerjs/docs-ui") as Promise<{ UniverDocsUIPlugin: new (...args: unknown[]) => unknown }>,
      import("@univerjs/design/locale/en-US") as Promise<{ default: Record<string, unknown> }>,
      import("@univerjs/ui/locale/en-US") as Promise<{ default: Record<string, unknown> }>,
      import("@univerjs/docs-ui/locale/en-US") as Promise<{ default: Record<string, unknown> }>,
      import("@univerjs/slides-ui/locale/en-US") as Promise<{ default: Record<string, unknown> }>,
    ]);

    const { createUniver, LocaleType, mergeLocales, UniverInstanceType } =
      presetsMod as IUniverPresets;

    if (token !== initTokenRef.current) return;

    const { univer, univerAPI } = createUniver({
      locale: LocaleType.EN_US,
      locales: {
        [LocaleType.EN_US]: mergeLocales(
          designEnUS.default, uiEnUS.default, docsUIEnUS.default, slidesUIEnUS.default,
        ),
      },
      plugins: [
        UniverRenderEnginePlugin,
        [UniverUIPlugin, { container: containerRef.current }],
        UniverDocsPlugin,
        UniverDocsUIPlugin,
        UniverDrawingPlugin,
        UniverSlidesPlugin,
        UniverSlidesUIPlugin,
      ],
    });

    apiRef.current = univerAPI;
    univer.createUnit(UniverInstanceType.UNIVER_SLIDE, SLIDE_DATA);
    univerRef.current = { destroy: () => { try { univer.dispose(); } catch { /* empty */ } } };
  }, []);

  const initSheet = useCallback((token: number, data?: Record<string, unknown>): void => {
    if (!containerRef.current) return;

    if (token !== initTokenRef.current) return;

    apiRef.current = undefined;
    setSheetGrid(snapshotToSheetGrid(data));
    setStatus("Ready");
  }, []);

  const initDoc = useCallback((token: number, data?: Record<string, unknown>): void => {
    if (!containerRef.current) return;

    if (token !== initTokenRef.current) return;

    apiRef.current = undefined;
    setDocText(snapshotToDocText(data));
    setStatus("Ready");
  }, []);

  const initUniver = useCallback(async (officeMode: OfficeMode, _data?: Record<string, unknown>): Promise<void> => {
    const token = invalidatePendingWork();
    destroyUniver(false);
    await waitForDeferredDispose();

    if (token !== initTokenRef.current) return;

    if (!containerRef.current) return;

    try {
      if (officeMode === "slides") {
        await initSlides(token);
      } else if (officeMode === "sheet") {
        initSheet(token, _data);
      } else {
        initDoc(token, _data);
      }
    } catch {
      // init failed silently
    }
  }, [destroyUniver, initDoc, initSheet, initSlides, invalidatePendingWork]);

  useEffect(() => {
    if (containerRef.current) {
      initUniver(mode);
    }
  }, [initUniver, mode]);

  useEffect(() => () => { destroyUniver(); }, [destroyUniver]);

  useEffect(() => {
    if (!url) return;

    const loadFile = async (): Promise<void> => {
      try {
        const buffer = await readFile(url);
        const file = new File([new Uint8Array(buffer)], url.split("/").pop() || "file", { type: "application/octet-stream" });
        const ext = extname(url).toLowerCase();
        const inferredMode: OfficeMode = ext === ".pptx" ? "slides" : ext === ".docx" ? "doc" : "sheet";

        setMode(inferredMode);

        if (ext === ".json") {
          const data = JSON.parse(buffer.toString()) as Record<string, unknown>;
          const body = data.body as Record<string, unknown> | undefined;
          const jsonMode: OfficeMode = body?.pages ? "slides" : data.sheets ? "sheet" : "doc";

          setMode(jsonMode);
          scheduleIfCurrent(initTokenRef.current, () => { initUniver(jsonMode, data).catch(() => { /* empty */ }); }, 100);

          return;
        }

        if (ext === ".xlsx") {
          const token = initTokenRef.current;

          scheduleTask(() => {
            (async (): Promise<void> => {
              if (token !== initTokenRef.current) return;

              try {
                const api = apiRef.current;

                if (api?.importXLSXToSnapshotAsync) {
                  const snapshot = await api.importXLSXToSnapshotAsync(file);

                  if (snapshot) {
                    scheduleIfCurrent(token, () => { initUniver("sheet", snapshot).catch(() => { /* empty */ }); }, 50);

                    return;
                  }
                }
              } catch { /* fallback */ }

              scheduleIfCurrent(token, () => { initUniver("sheet").catch(() => { /* empty */ }); }, 50);
            })().catch(() => { /* empty */ });
          }, 200);

          return;
        }

        if (ext === ".docx") {
          const token = initTokenRef.current;

          scheduleTask(() => {
            (async (): Promise<void> => {
              if (token !== initTokenRef.current) return;

              try {
                const api = apiRef.current;

                if (api?.importDOCXToSnapshotAsync) {
                  const snapshot = await api.importDOCXToSnapshotAsync(file);

                  if (snapshot) {
                    scheduleIfCurrent(token, () => { initUniver("doc", snapshot).catch(() => { /* empty */ }); }, 50);

                    return;
                  }
                }
              } catch { /* fallback */ }

              scheduleIfCurrent(token, () => { initUniver("doc").catch(() => { /* empty */ }); }, 50);
            })().catch(() => { /* empty */ });
          }, 200);

          return;
        }

        if (ext === ".pptx") {
          scheduleIfCurrent(initTokenRef.current, () => { initUniver("slides").catch(() => { /* empty */ }); }, 100);
        }
      } catch {
        // file load failed silently
      }
    };

    loadFile().catch(() => { /* empty */ });
  }, [initUniver, readFile, scheduleIfCurrent, scheduleTask, url]);

  const handleOpenFile = useCallback((): void => {
    fileInputRef.current?.click();
  }, []);

  const handleFileSelected = useCallback(async (e: React.ChangeEvent<HTMLInputElement>): Promise<void> => {
    const file = e.target.files?.[0];

    if (!file) return;

    const ext = file.name.split(".").pop()?.toLowerCase();

    try {
      if (ext === "json") {
        const text = await file.text();
        const data = JSON.parse(text) as Record<string, unknown>;
        const body = data.body as Record<string, unknown> | undefined;

        setActiveMenu(undefined);

        const inferredMode: OfficeMode = body?.pages ? "slides" : data.sheets ? "sheet" : "doc";

        setMode(inferredMode);
        scheduleIfCurrent(initTokenRef.current, () => { initUniver(inferredMode, data).catch(() => { /* empty */ }); }, 50);

        e.target.value = "";

        return;
      }

      if (ext === "xlsx") {
        setActiveMenu(undefined);
        setMode("sheet");

        try {
          const api = apiRef.current;

          if (api?.importXLSXToSnapshotAsync) {
            const snapshot = await api.importXLSXToSnapshotAsync(file);

            if (snapshot) {
              scheduleIfCurrent(initTokenRef.current, () => { initUniver("sheet", snapshot).catch(() => { /* empty */ }); }, 50);
              e.target.value = "";

              return;
            }
          }
        } catch { /* fallback */ }

        scheduleIfCurrent(initTokenRef.current, () => { initUniver("sheet").catch(() => { /* empty */ }); }, 50);
        e.target.value = "";

        return;
      }

      if (ext === "docx") {
        setActiveMenu(undefined);
        setMode("doc");

        try {
          const api = apiRef.current;

          if (api?.importDOCXToSnapshotAsync) {
            const snapshot = await api.importDOCXToSnapshotAsync(file);

            if (snapshot) {
              scheduleIfCurrent(initTokenRef.current, () => { initUniver("doc", snapshot).catch(() => { /* empty */ }); }, 50);
              e.target.value = "";

              return;
            }
          }
        } catch { /* fallback */ }

        scheduleIfCurrent(initTokenRef.current, () => { initUniver("doc").catch(() => { /* empty */ }); }, 50);
        e.target.value = "";

        return;
      }

      if (ext === "pptx") {
        setActiveMenu(undefined);
        setMode("slides");
        scheduleIfCurrent(initTokenRef.current, () => { initUniver("slides").catch(() => { /* empty */ }); }, 50);
      }
    } catch {
      // file load failed silently
    }

    e.target.value = "";
  }, [initUniver, scheduleIfCurrent]);

  const getFallbackSheetData = useCallback((): string[][] => {
    const rows = [...(sheetFallbackRef.current?.querySelectorAll(".sheet-row") || [])];

    return rows.slice(1).map((row) =>
      [...row.querySelectorAll(".sheet-cell")].map((cell) => cell.textContent?.trim() || ""));
  }, []);

  const getFallbackDocText = useCallback((): string => docFallbackRef.current?.textContent?.trim() || "Untitled document", []);

  const saveFallbackSheetJson = useCallback(async (): Promise<void> => {
    const data = getFallbackSheetData();
    const content = JSON.stringify({ sheets: { Sheet1: { data } } }, undefined, 2);

    await writeDesktopFile("Spreadsheet", "json", content);
  }, [getFallbackSheetData, writeDesktopFile]);

  const saveFallbackDocJson = useCallback(async (): Promise<void> => {
    const content = JSON.stringify({ body: { text: getFallbackDocText() } }, undefined, 2);

    await writeDesktopFile("Document", "json", content);
  }, [getFallbackDocText, writeDesktopFile]);

  const handleSave = useCallback(async (): Promise<void> => {
    try {
      if (mode === "sheet") {
        const wb = apiRef.current?.getActiveWorkbook();

        if (wb) {
          const snapshot = wb.save();
          const content = JSON.stringify(snapshot, undefined, 2);

          await writeDesktopFile("Spreadsheet", "json", content);
        } else {
          await saveFallbackSheetJson();
        }
      } else if (mode === "doc") {
        const doc = apiRef.current?.getActiveDocument();

        if (doc) {
          const snapshot = doc.getSnapshot();
          const content = JSON.stringify(snapshot, undefined, 2);

          await writeDesktopFile("Document", "json", content);
        } else {
          await saveFallbackDocJson();
        }
      } else {
        const content = JSON.stringify(SLIDE_DATA, undefined, 2);

        await writeDesktopFile("Presentation", "json", content);
      }
    } catch {
      setReadyStatus("Save failed");
    }

    setActiveMenu(undefined);
  }, [mode, saveFallbackDocJson, saveFallbackSheetJson, setReadyStatus, writeDesktopFile]);

  const handleExportXLSX = useCallback(async (): Promise<void> => {
    try {
      const wb = apiRef.current?.getActiveWorkbook();

      if (wb) {
        const snapshot = wb.save();
        const { sheets } = snapshot as { sheets?: Record<string, {
          cellData?: Record<string, Record<string, { v?: string | number | boolean }>>;
          name?: string;
        }> };

        if (sheets) {
          const workbook = XLSXLib.utils.book_new();

          for (const sheetId of Object.keys(sheets)) {
            const sheet = sheets[sheetId];
            const { cellData, name: sheetName } = sheet;

            if (cellData) {
              const rowKeys = Object.keys(cellData).map(Number).sort((a, b) => a - b);
              let maxCol = 0;

              for (const r of rowKeys) {
                const cols = Object.keys(cellData[String(r)]).map(Number);

                if (cols.length > 0) maxCol = Math.max(maxCol, ...cols);
              }

              const data: (string | number | boolean)[][] = [];

              for (const r of rowKeys) {
                const rowData = cellData[String(r)];

                for (let c = 0; c <= maxCol; c++) {
                  data[r] = data[r] || [];
                  data[r][c] = rowData?.[String(c)]?.v ?? "";
                }
              }

              const ws = XLSXLib.utils.aoa_to_sheet(data);
              XLSXLib.utils.book_append_sheet(workbook, ws, sheetName || "Sheet1");
            }
          }

          const xlsxData: unknown = XLSXLib.write(workbook, { bookType: "xlsx", type: "array" });
          const xlsxBlob = new Blob([xlsxData as BlobPart], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });

          await writeDesktopFile("Spreadsheet", "xlsx", await blobToBuffer(xlsxBlob));
          setActiveExportMenu(false);

          return;
        }
      }

      const fallbackWorkbook = XLSXLib.utils.book_new();
      const worksheet = XLSXLib.utils.aoa_to_sheet(getFallbackSheetData());

      XLSXLib.utils.book_append_sheet(fallbackWorkbook, worksheet, "Sheet1");

      const fallbackData: unknown = XLSXLib.write(fallbackWorkbook, { bookType: "xlsx", type: "array" });
      const fallbackBlob = new Blob([fallbackData as BlobPart], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });

      await writeDesktopFile("Spreadsheet", "xlsx", await blobToBuffer(fallbackBlob));
    } catch {
      setReadyStatus("Export failed");
    }

    setActiveExportMenu(false);
  }, [getFallbackSheetData, setReadyStatus, writeDesktopFile]);

  const handleExportDOCX = useCallback(async (): Promise<void> => {
    try {
      const doc = apiRef.current?.getActiveDocument();

      if (doc) {
        const snapshot = doc.getSnapshot() as {
          body?: { dataStream?: string };
        };
        const dataStream = snapshot?.body?.dataStream;

        if (dataStream) {
          const paragraphs = dataStream
            .split("\r")
            .filter((t) => t.length > 0)
            .map((text) => new Paragraph({ children: [new TextRun({ text })] }));
          const docObj = new Doc({ sections: [{ children: paragraphs }] });
          const docxBlob = await Packer.toBlob(docObj);

          await writeDesktopFile("Document", "docx", await blobToBuffer(docxBlob));
          setActiveExportMenu(false);

          return;
        }
      }

      const fallbackParagraphs = getFallbackDocText().split(/\r?\n/).filter(Boolean)
        .map((text) => new Paragraph({ children: [new TextRun({ text })] }));
      const fallbackDocObj = new Doc({ sections: [{ children: fallbackParagraphs }] });
      const fallbackDocxBlob = await Packer.toBlob(fallbackDocObj);

      await writeDesktopFile("Document", "docx", await blobToBuffer(fallbackDocxBlob));
    } catch {
      setReadyStatus("Export failed");
    }

    setActiveExportMenu(false);
  }, [getFallbackDocText, setReadyStatus, writeDesktopFile]);

  const handleExportPPTX = useCallback(async (): Promise<void> => {
    try {
      const slideBody = SLIDE_DATA.body as { pageOrder?: string[]; pageSize?: { height?: number; width?: number }; pages?: Record<string, ISlidePageDef> } | undefined;
      const pptxBlob = await createPptxBlob(slideBody || {});

      await writeDesktopFile("Presentation", "pptx", await blobToBuffer(pptxBlob));
      setActiveExportMenu(false);
    } catch {
      setReadyStatus("Export failed");
      setActiveExportMenu(false);
    }
  }, [setReadyStatus, writeDesktopFile]);
 
  const handleExportClick = useCallback((): void => {
    if (mode === "sheet") handleExportXLSX();
    else if (mode === "doc") handleExportDOCX();
    else handleExportPPTX();
  }, [handleExportDOCX, handleExportPPTX, handleExportXLSX, mode]);

  const handleExportFormat = useCallback((format: string): void => {
    setActiveExportMenu(false);

    if (format === "xlsx") handleExportXLSX();
    else if (format === "docx") handleExportDOCX();
    else if (format === "pptx") handleExportPPTX();
  }, [handleExportDOCX, handleExportPPTX, handleExportXLSX]);

  const handleNew = useCallback((): void => {
    setActiveMenu(undefined);
    destroyUniver();
    scheduleIfCurrent(initTokenRef.current, () => { initUniver(mode).catch(() => { /* empty */ }); }, 50);
    setReadyStatus("New document");
  }, [destroyUniver, initUniver, mode, scheduleIfCurrent, setReadyStatus]);

  const handleFormat = useCallback((command: "bold" | "italic" | "underline"): void => {
    const editableDocument = document as unknown as { execCommand: (commandId: string) => boolean };

    editableDocument.execCommand(command);
    setReadyStatus(command[0].toUpperCase() + command.slice(1));
  }, [setReadyStatus]);

  const handleMenuAction = useCallback((action: () => void): void => {
    action();
    setActiveMenu(undefined);
  }, []);

  const toggleMenu = useCallback((menu: string): void => {
    setActiveMenu((prev) => (prev === menu ? undefined : menu));
  }, []);

  useEffect(() => {
    const handleClickOutside = (e: MouseEvent): void => {
      if (menuBarRef.current && !menuBarRef.current.contains(e.target as Node)) {
        setActiveMenu(undefined);
      }

      if (exportBtnRef.current && !exportBtnRef.current.contains(e.target as Node)) {
        setActiveExportMenu(false);
      }
    };

    document.addEventListener("mousedown", handleClickOutside);

    return () => document.removeEventListener("mousedown", handleClickOutside);
  }, []);

  useEffect(() => {
    const handleKeyboard = (e: KeyboardEvent): void => {
      if ((e.ctrlKey || e.metaKey) && e.key === "n") {
        e.preventDefault();
        handleNew();
      } else if ((e.ctrlKey || e.metaKey) && e.key === "o") {
        e.preventDefault();
        handleOpenFile();
      } else if ((e.ctrlKey || e.metaKey) && e.key === "s") {
        e.preventDefault();
        handleSave();
      }
    };

    document.addEventListener("keydown", handleKeyboard);

    return () => document.removeEventListener("keydown", handleKeyboard);
  }, [handleNew, handleOpenFile, handleSave]);

  const FILE_MENU: MenuEntry[] = [
    { action: handleNew, label: "New", shortcut: "Ctrl+N", type: "item" },
    { action: handleOpenFile, label: "Open...", shortcut: "Ctrl+O", type: "item" },
    { type: "divider" },
    { action: handleSave, label: "Save", shortcut: "Ctrl+S", type: "item" },
    { type: "divider" },
    { action: handleExportXLSX, label: "Export XLSX", type: "item" },
    { action: handleExportDOCX, label: "Export DOCX", type: "item" },
    { action: handleExportPPTX, label: "Export PPTX", type: "item" },
  ];

  const EDIT_MENU: MenuEntry[] = [
    { action: () => { /* noop */ }, label: "Undo", shortcut: "Ctrl+Z", type: "item" },
    { action: () => { /* noop */ }, label: "Redo", shortcut: "Ctrl+Y", type: "item" },
  ];

  const VIEW_MENU: MenuEntry[] = MODES.map((m) => ({
    action: () => setMode(m.id),
    label: m.label,
    shortcut: m.id === mode ? "\u2713" : undefined,
    type: "item" as const,
  }));

  const HELP_MENU: MenuEntry[] = [
    { action: () => {
      setActiveMenu("About");
    }, label: "About MAS Office", type: "item" },
  ];

  const renderMenu = (label: string, entries: MenuEntry[]): React.ReactNode => (
    <div
      key={label}
      className={`menu-trigger${activeMenu === label ? " active" : ""}`}
      onKeyDown={(e: React.KeyboardEvent) => { if (e.key === "Enter" || e.key === " ") toggleMenu(label); }}
      onMouseDown={() => toggleMenu(label)}
      onMouseEnter={() => { if (activeMenu) setActiveMenu(label); }}
      role="menuitem"
      tabIndex={0}
    >
      {label}
      {activeMenu === label && (
        <div className="menu-dropdown" role="menu">
          {entries.map((entry, index) => (
            entry.type === "divider" ? (
              // eslint-disable-next-line react/no-array-index-key
              <div key={index} className="menu-divider" role="separator" />
            ) : (
              <button
                key={entry.label}
                className="menu-dropdown-item"
                onMouseDown={(evt) => { evt.stopPropagation(); handleMenuAction(entry.action); }}
                type="button"
              >
                <span className="menu-item-label">{entry.label}</span>
                {entry.shortcut && <span className="menu-shortcut">{entry.shortcut}</span>}
              </button>
            )
          ))}
        </div>
      )}
    </div>
  );

  return (
    <StyledMASOffice>
      <div ref={menuBarRef} className="menu-bar" role="menubar">
        {renderMenu("File", FILE_MENU)}
        {renderMenu("Edit", EDIT_MENU)}
        {renderMenu("View", VIEW_MENU)}
        {renderMenu("Help", HELP_MENU)}
      </div>
      <div className="toolbar">
        <button className="toolbar-btn" onClick={handleNew} title="New" type="button">
          <svg fill="currentColor" height="16" viewBox="0 0 16 16" width="16">
            <path d="M2 2h7l3 3v9H2V2zm1 1v10h9V6H8V3H3zm6 .5V5h1.5L9 3.5z" />
          </svg>
        </button>
        <button className="toolbar-btn" onClick={handleOpenFile} title="Open" type="button">
          <svg fill="currentColor" height="16" viewBox="0 0 16 16" width="16">
            <path d="M2 2h5l2 2h5v10H2V2zm1 1v10h12V5H8.5L7 4H3zm2 3h7v1H5V6zm0 2h7v1H5V8zm0 2h5v1H5v-1z" />
          </svg>
        </button>
        <button className="toolbar-btn" onClick={handleSave} title="Save" type="button">
          <svg fill="currentColor" height="16" viewBox="0 0 16 16" width="16">
            <path d="M2 2v12h12V5l-3-3H2zm1 1h3v3H3V3zm4 0h2v2H7V3zm-4 5h10v5H3V8zm1 1v3h8V9H4z" />
          </svg>
        </button>
        <div className="toolbar-divider" />
        <div
          ref={exportBtnRef}
          className={`toolbar-btn${activeExportMenu ? " active" : ""}`}
          onClick={handleExportClick}
          onKeyDown={(e: React.KeyboardEvent) => { if (e.key === "Enter" || e.key === " ") { e.preventDefault(); handleExportClick(); } }}
          role="button"
          tabIndex={0}
          title="Export to Desktop"
        >
          <svg fill="currentColor" height="16" viewBox="0 0 16 16" width="16">
            <path d="M8 1l4 4h-3v6H7V5H4l4-4zM2 12v2h12v-2h1v3H1v-3h1z" />
          </svg>
          {activeExportMenu && (
            <div className="export-dropdown">
              <button className="menu-dropdown-item" onMouseDown={() => handleExportFormat("xlsx")} type="button">
                <span className="menu-item-label">Export as XLSX</span>
              </button>
              <button className="menu-dropdown-item" onMouseDown={() => handleExportFormat("docx")} type="button">
                <span className="menu-item-label">Export as DOCX</span>
              </button>
              <button className="menu-dropdown-item" onMouseDown={() => handleExportFormat("pptx")} type="button">
                <span className="menu-item-label">Export as PPTX</span>
              </button>
            </div>
          )}
        </div>
        <div className="toolbar-divider" />
        <button className="toolbar-btn" onClick={() => handleFormat("bold")} title="Bold" type="button">
          <svg fill="currentColor" height="16" viewBox="0 0 16 16" width="16">
            <path d="M4 2h5a3 3 0 012.5 4.7A3 3 0 0110 14H4V2zm1 5V3h4a2 2 0 010 4H5zm0 1v5h5a2 2 0 100-4H5z" />
          </svg>
        </button>
        <button className="toolbar-btn" onClick={() => handleFormat("italic")} title="Italic" type="button">
          <svg fill="currentColor" height="16" viewBox="0 0 16 16" width="16">
            <path d="M6 2h6v2H9l-2 8h3v2H5v-2h3l2-8H8V2h3v-1H5v1z" />
          </svg>
        </button>
        <button className="toolbar-btn" onClick={() => handleFormat("underline")} title="Underline" type="button">
          <svg fill="currentColor" height="16" viewBox="0 0 16 16" width="16">
            <path d="M4 2v5a4 4 0 008 0V2h-1v5a3 3 0 11-6 0V2H4zM3 13h10v1H3v-1z" />
          </svg>
        </button>
        <div className="toolbar-spacer" />
        <div className="mode-tabs">
          {MODES.map((m) => (
            <button
              key={m.id}
              className={`mode-tab${mode === m.id ? " active" : ""}`}
              onClick={() => setMode(m.id)}
              type="button"
            >
              {m.label}
            </button>
          ))}
        </div>
      </div>
      <div className="univer-container">
        {mode === "slides" && <div ref={containerRef} className="univer-host" />}
        {mode === "sheet" && renderSheetFallback(sheetFallbackRef, sheetGrid)}
        {mode === "doc" && renderDocFallback(docFallbackRef, docText)}
      </div>
      <div className="status-bar">
        <span className="status-section">{MODES.find((m) => m.id === mode)?.label || "Sheet"}</span>
        <span className="status-divider" />
        <span className="status-section">{status}</span>
        <span className="status-right">MAS Office</span>
      </div>
      <input
        ref={fileInputRef}
        accept=".xlsx,.docx,.pptx,.json"
        onChange={handleFileSelected}
        style={{ display: "none" }}
        type="file"
      />
    </StyledMASOffice>
  );
};

export default memo(MASOffice);
