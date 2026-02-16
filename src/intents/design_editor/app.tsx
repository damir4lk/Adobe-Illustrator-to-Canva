import React from "react";
import { Rows, Text, Button, Title } from "@canva/app-ui-kit";
import { upload, findFonts, requestFontSelection } from "@canva/asset";
import type { ImageRef, FontRef } from "@canva/asset";
import { addPage, openDesign } from "@canva/design";

/**
 * ILLUSTRATOR IMPORTER v4.6
 *
 * КЛЮЧЕВЫЕ ИЗМЕНЕНИЯ:
 *   - findFonts() возвращает ПОДМНОЖЕСТВО Canva → агрессивный partial matching
 *   - requestFontSelection() — встроенный пикер Canva со ВСЕМИ шрифтами
 *   - Кэш выбора: fontFamily → FontRef | "use_svg" (сохраняется между артбордами)
 *   - Сначала резолвим ВСЕ уникальные шрифты, потом строим элементы
 */

type TextData = {
  content: string;
  fontFamily: string;
  fontPostScriptName?: string;
  fontSize: number;
  fontWeight: string;
  fontStyle: string;
  color: string;
  alignment: string;
  lineHeight: number;
  letterSpacing: number;
  textDecoration: string;
};

type ClipShape = {
  type: "ellipse" | "rect" | "unknown";
  cornerRadius?: number;
};

type StrokeBounds = {
  x: number; y: number;
  width: number; height: number;
};

type LayoutObject = {
  index: number;
  fileName: string;
  contentFileName?: string;
  strokeFileName?: string;
  strokeBounds?: StrokeBounds;
  type: "svg" | "png" | "text" | "clipMask";
  objectType: string;
  x: number; y: number;
  width: number; height: number;
  zIndex: number;
  opacity: number;
  textData?: TextData;
  textSvgFileName?: string;
  clipShape?: ClipShape;
};

type Layout = {
  version?: string;
  artboardName: string;
  width: number; height: number;
  objects: LayoutObject[];
};

type FontChoice = FontRef | "use_svg";

const FONT_WEIGHT_MAP: Record<string, string> = {
  normal: "normal", thin: "thin", extralight: "extralight",
  light: "light", medium: "medium", semibold: "semibold",
  bold: "bold", ultrabold: "ultrabold", heavy: "heavy", black: "heavy",
};

const TEXT_ALIGN_MAP: Record<string, "start" | "center" | "end" | "justify"> = {
  left: "start", center: "center", right: "end", justify: "justify",
};

const TEXT_WIDTH_PADDING = 1.15;

// ═══ ГЕНЕРАЦИЯ SVG PATH ДЛЯ CANVA FRAME ═══
// Canva требует: 1 M, нет Q, закрывать Z
// viewBox = width x height элемента
function generateShapePath(
  clipShape: ClipShape | undefined,
  w: number,
  h: number
): string | null {
  if (!clipShape) return null;

  const r2 = (v: number) => Math.round(v * 100) / 100;
  const K = 0.5523; // bezier approx for circle

  if (clipShape.type === "ellipse") {
    const cx = w / 2, cy = h / 2;
    const rx = w / 2, ry = h / 2;
    // 4 cubic bezier curves approximating ellipse
    return [
      "M", r2(cx), 0,
      "C", r2(cx + rx * K), 0, r2(w), r2(cy - ry * K), r2(w), r2(cy),
      "C", r2(w), r2(cy + ry * K), r2(cx + rx * K), r2(h), r2(cx), r2(h),
      "C", r2(cx - rx * K), r2(h), 0, r2(cy + ry * K), 0, r2(cy),
      "C", 0, r2(cy - ry * K), r2(cx - rx * K), 0, r2(cx), 0,
      "Z"
    ].join(" ");
  }

  if (clipShape.type === "rect") {
    const cr = clipShape.cornerRadius || 0;
    if (cr <= 0) {
      // Простой прямоугольник
      return "M 0 0 L " + r2(w) + " 0 L " + r2(w) + " " + r2(h) + " L 0 " + r2(h) + " Z";
    }
    // Скруглённые углы через cubic bezier
    const rr = Math.min(cr, w / 2, h / 2); // не больше половины стороны
    const kk = rr * K; // handle distance
    return [
      "M", r2(rr), 0,
      "L", r2(w - rr), 0,
      "C", r2(w - rr + kk), 0, r2(w), r2(rr - kk), r2(w), r2(rr),
      "L", r2(w), r2(h - rr),
      "C", r2(w), r2(h - rr + kk), r2(w - rr + kk), r2(h), r2(w - rr), r2(h),
      "L", r2(rr), r2(h),
      "C", r2(rr - kk), r2(h), 0, r2(h - rr + kk), 0, r2(h - rr),
      "L", 0, r2(rr),
      "C", 0, r2(rr - kk), r2(rr - kk), 0, r2(rr), 0,
      "Z"
    ].join(" ");
  }

  // unknown → null (будет fallback к обрезанной картинке)
  return null;
}

// Нормализация: убрать пробелы, дефисы, подчёркивания
function norm(s: string): string {
  return s.toLowerCase().replace(/[\s\-_]+/g, "");
}

// Убрать суффиксы стиля: "Montserrat Bold Italic" → "montserrat"
function stripStyle(s: string): string {
  return s.toLowerCase()
    .replace(/[\s\-]*(bold|italic|light|medium|thin|heavy|regular|semibold|black|condensed|oblique|book|demi|ultra|extra|narrow)[\s\-]*/gi, " ")
    .replace(/[\s\-_]+/g, "")
    .trim();
}

export function App() {
  const [layouts, setLayouts] = React.useState<Layout[]>([]);
  const [selectedLayoutIndex, setSelectedLayoutIndex] = React.useState(0);
  const [files, setFiles] = React.useState<Map<string, File>>(new Map());
  const [artboardFileMap, setArtboardFileMap] = React.useState<Map<string, Map<string, string>>>(new Map());
  const [status, setStatus] = React.useState("Готов к импорту");
  const [progress, setProgress] = React.useState(0);
  const [isImporting, setIsImporting] = React.useState(false);
  const [importLog, setImportLog] = React.useState<string[]>([]);

  // Кэш из findFonts (подмножество)
  const [fontCacheRaw, setFontCacheRaw] = React.useState<Map<string, FontRef>>(new Map());
  const [fontCacheNorm, setFontCacheNorm] = React.useState<Map<string, FontRef>>(new Map());

  // ═══ КЭШ ВЫБОРА ПОЛЬЗОВАТЕЛЯ (сохраняется между артбордами!) ═══
  const userFontChoicesRef = React.useRef<Map<string, FontChoice>>(new Map());

  // Для диалога выбора шрифта (Promise-based)
  const [pendingFontResolve, setPendingFontResolve] = React.useState<((choice: FontChoice) => void) | null>(null);
  const [pendingFontName, setPendingFontName] = React.useState("");
  const [showFontChoices, setShowFontChoices] = React.useState(false);

  const addLog = (msg: string) => {
    setImportLog((prev) => [...prev.slice(-80), msg]);
  };

  // ═══ Загрузка шрифтов ═══
  const preloadFonts = async () => {
    try {
      const { fonts } = await findFonts();
      const raw = new Map<string, FontRef>();
      const normalized = new Map<string, FontRef>();

      for (const font of fonts) {
        raw.set(font.name.toLowerCase(), font.ref);
        normalized.set(norm(font.name), font.ref);
        // Также без суффиксов стиля
        const stripped = stripStyle(font.name);
        if (stripped && !normalized.has(stripped)) {
          normalized.set(stripped, font.ref);
        }
      }

      setFontCacheRaw(raw);
      setFontCacheNorm(normalized);
      addLog("findFonts: " + String(fonts.length) + " шрифтов (подмножество!)");

      const sample = fonts.slice(0, 8).map((f) => f.name);
      addLog("Примеры: " + sample.join(", "));

      return { raw, normalized };
    } catch (e) {
      addLog("Ошибка findFonts: " + String(e));
      return { raw: new Map<string, FontRef>(), normalized: new Map<string, FontRef>() };
    }
  };

  // ═══════════════════════════════════════════
  // АГРЕССИВНЫЙ ПОИСК ШРИФТА
  // findFonts() возвращает мало шрифтов!
  // Поэтому ищем: exact → normalized → stripped → partial contains
  // ═══════════════════════════════════════════
  const findCanvaFont = (
    fontFamily: string,
    rawC: Map<string, FontRef>,
    normC: Map<string, FontRef>
  ): FontRef | undefined => {
    if (!fontFamily || rawC.size === 0) return undefined;

    const lower = fontFamily.toLowerCase().trim();
    const normalized = norm(fontFamily);
    const stripped = stripStyle(fontFamily);

    // 1. Точное совпадение по lowercase имени
    if (rawC.has(lower)) return rawC.get(lower);

    // 2. Нормализованное (без пробелов/дефисов)
    if (normC.has(normalized)) return normC.get(normalized);

    // 3. Без суффиксов стиля: "Montserrat Bold" → "montserrat"
    if (stripped && stripped !== normalized && normC.has(stripped)) {
      return normC.get(stripped);
    }

    // 4. Partial match: имя из Illustrator СОДЕРЖИТ имя из Canva (или наоборот)
    //    "opensans" contains "opensans" ✓
    //    "arialnarrow" contains "arial" ✓ (минимум 4 символа)
    for (const [canvaLower, ref] of rawC) {
      const canvaNorm = norm(canvaLower);
      if (canvaNorm.length >= 4 && normalized.length >= 4) {
        if (normalized.includes(canvaNorm) || canvaNorm.includes(normalized)) {
          return ref;
        }
      }
    }

    // 5. Stripped partial match
    if (stripped && stripped.length >= 4) {
      for (const [canvaLower, ref] of rawC) {
        const canvaStripped = stripStyle(canvaLower);
        if (canvaStripped && canvaStripped.length >= 4) {
          if (stripped.includes(canvaStripped) || canvaStripped.includes(stripped)) {
            return ref;
          }
        }
      }
    }

    return undefined;
  };

  // ═══ ЗАПРОС У ПОЛЬЗОВАТЕЛЯ ═══
  const askUserForFont = (fontFamily: string): Promise<FontChoice> => {
    return new Promise<FontChoice>((resolve) => {
      setPendingFontName(fontFamily);
      setPendingFontResolve(() => resolve);
    });
  };

  // Пользователь нажал "Выбрать шрифт"
  const handlePickFont = async () => {
    if (!pendingFontResolve) return;
    const resolve = pendingFontResolve;
    const fontName = pendingFontName;

    try {
      const result = await requestFontSelection();
      if (result.type === "completed" && result.font) {
        addLog("✓ Выбран: " + result.font.name + " → для «" + fontName + "»");
        userFontChoicesRef.current.set(fontName.toLowerCase(), result.font.ref);
        setShowFontChoices((prev) => !prev); // trigger re-render for choices display
        resolve(result.font.ref);
      } else {
        // Отмена → SVG
        addLog("✓ Отмена пикера → SVG для «" + fontName + "»");
        userFontChoicesRef.current.set(fontName.toLowerCase(), "use_svg");
        resolve("use_svg");
      }
    } catch (e) {
      addLog("Ошибка пикера: " + String(e));
      resolve("use_svg");
    } finally {
      setPendingFontResolve(null);
      setPendingFontName("");
    }
  };

  // Пользователь нажал "Использовать SVG"
  const handleUseSvg = () => {
    if (!pendingFontResolve) return;
    const resolve = pendingFontResolve;
    const fontName = pendingFontName;

    addLog("✓ SVG для «" + fontName + "»");
    userFontChoicesRef.current.set(fontName.toLowerCase(), "use_svg");
    setShowFontChoices((prev) => !prev);
    resolve("use_svg");
    setPendingFontResolve(null);
    setPendingFontName("");
  };

  // ─── Upload ───
  const uploadFile = async (file: File): Promise<ImageRef | null> => {
    try {
      const dataUrl = await new Promise<string>((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = () => resolve(reader.result as string);
        reader.onerror = reject;
        reader.readAsDataURL(file);
      });
      const mime = file.name.toLowerCase().endsWith(".svg") ? "image/svg+xml" : "image/png";
      const result = await upload({
        type: "image",
        mimeType: mime as "image/svg+xml" | "image/png",
        url: dataUrl, thumbnailUrl: dataUrl,
        aiDisclosure: "none",
      });
      return result.ref;
    } catch (e) {
      return null;
    }
  };

  // ─── Файл по артборду ───
  const getFile = (abName: string, fileName: string): File | undefined => {
    const af = artboardFileMap.get(abName);
    if (af) { const fp = af.get(fileName); if (fp) return files.get(fp); }
    for (const [, m] of artboardFileMap) {
      const fp = m.get(fileName); if (fp) return files.get(fp);
    }
    for (const [p, f] of files) { if (p.endsWith("/" + fileName)) return f; }
    return undefined;
  };

  // ═══ Выбор папки ═══
  const handleFolderSelect = async () => {
    const input = document.createElement("input");
    input.type = "file";
    input.webkitdirectory = true;
    input.multiple = true;

    input.onchange = async (e: any) => {
      const fileList = Array.from(e.target.files) as File[];
      const fpm = new Map<string, File>();
      const abm = new Map<string, Map<string, string>>();

      for (const file of fileList) {
        const rp = (file as any).webkitRelativePath || file.name;
        fpm.set(rp, file);
        const parts = rp.split("/");
        if (parts.length >= 3) {
          const dir = parts[1];
          if (!abm.has(dir)) abm.set(dir, new Map());
          abm.get(dir)!.set(parts[parts.length - 1], rp);
        } else if (parts.length === 2) {
          if (!abm.has("__root__")) abm.set("__root__", new Map());
          abm.get("__root__")!.set(parts[parts.length - 1], rp);
        }
      }

      setFiles(fpm);
      setArtboardFileMap(abm);

      const abNames = Array.from(abm.keys()).filter((k) => k !== "__root__");
      addLog("Папки: " + (abNames.length > 0 ? abNames.join(", ") : "(корень)"));

      for (const [dir, fileMap] of abm) {
        const fnames = Array.from(fileMap.keys());
        addLog("  " + dir + ": " + fnames.filter((f) => f.endsWith("_text.svg")).length + " text SVG, " + fnames.filter((f) => f.endsWith(".svg") && !f.includes("_text")).length + " SVG, " + fnames.filter((f) => f.endsWith(".json") && f !== "layout.json").length + " JSON");
      }

      const layoutFiles = fileList.filter((f) => f.name === "layout.json");
      if (layoutFiles.length === 0) { setStatus("layout.json не найден!"); return; }

      const loaded: Layout[] = [];
      for (const lf of layoutFiles) {
        try { loaded.push(JSON.parse(await lf.text()) as Layout); }
        catch (err) { addLog("Ошибка парсинга: " + String(err)); }
      }
      if (loaded.length === 0) { setStatus("Ошибка layout.json!"); return; }

      setLayouts(loaded);
      setSelectedLayoutIndex(0);
      await preloadFonts();

      if (fileList.some((f) => ((f as any).webkitRelativePath || "").includes("_fonts/"))) {
        addLog("Папка _fonts → загрузите в Canva Brand Kit");
      }

      setStatus(String(loaded.length) + " артборд(ов), " + String(loaded.reduce((s, l) => s + l.objects.length, 0)) + " объектов");
    };
    input.click();
  };

  // ═══════════════════════════════════════════
  // РЕЗОЛВ ШРИФТА: авто → кэш → спросить
  // ═══════════════════════════════════════════
  const resolveFontRef = async (
    fontFamily: string,
    rawC: Map<string, FontRef>,
    normC: Map<string, FontRef>
  ): Promise<FontChoice | undefined> => {

    // 1. Автоматический поиск
    const autoRef = findCanvaFont(fontFamily, rawC, normC);
    if (autoRef) {
      addLog("✓ " + fontFamily + " → авто");
      return autoRef;
    }

    // 2. Кэш пользователя
    const cached = userFontChoicesRef.current.get(fontFamily.toLowerCase());
    if (cached !== undefined) {
      addLog("✓ " + fontFamily + " → кэш (" + (cached === "use_svg" ? "SVG" : "шрифт") + ")");
      return cached;
    }

    // 3. Спрашиваем пользователя
    addLog("? " + fontFamily + " — не найден, жду выбор...");
    setStatus("Шрифт не найден: " + fontFamily);
    const choice = await askUserForFont(fontFamily);
    return choice;
  };

  // ═══════════════════════════════════════════
  // ИМПОРТ АРТБОРДА
  // ═══════════════════════════════════════════
  const handleImportLayout = async (layoutIndex: number) => {
    const layout = layouts[layoutIndex];
    if (!layout) return;

    setIsImporting(true);
    setProgress(0);
    addLog("─── " + layout.artboardName + " " + String(layout.width) + "x" + String(layout.height) + " ───");

    const sorted = [...layout.objects].sort((a, b) => (a.zIndex ?? a.index) - (b.zIndex ?? b.index));
    const opacityMap: Array<{ elementIndex: number; opacity: number }> = [];

    try {
      // ═══ ШАГ 1: Загрузка файлов ═══
      const imageRefs = new Map<string, ImageRef>();
      const toUpload = new Set<string>();

      for (const obj of sorted) {
        if (obj.type === "svg" || obj.type === "png") toUpload.add(obj.fileName);
        if (obj.textSvgFileName) toUpload.add(obj.textSvgFileName);
        if (obj.type === "clipMask") {
          toUpload.add(obj.fileName); // clipped fallback
          if (obj.contentFileName) toUpload.add(obj.contentFileName);
          if (obj.strokeFileName) toUpload.add(obj.strokeFileName);
        }
      }

      const uploadList = Array.from(toUpload);
      addLog("Загрузка " + String(uploadList.length) + " файлов...");

      for (let i = 0; i < uploadList.length; i++) {
        const fname = uploadList[i];
        setStatus("Загрузка " + String(i + 1) + "/" + String(uploadList.length));
        setProgress(Math.round(((i + 1) / uploadList.length) * 50));

        if (!fname) {
          addLog("ПРОПУСК: пустое имя файла");
          continue;
        }

        const file = getFile(layout.artboardName, fname);
        if (!file) {
          addLog("НЕТ ФАЙЛА: " + fname + " (артборд: " + layout.artboardName + ")");
          continue;
        }

        const ref = await uploadFile(file);
        if (ref) {
          imageRefs.set(fname, ref);
        } else {
          addLog("UPLOAD FAIL: " + fname);
        }
        await new Promise((r) => setTimeout(r, 500));
      }

      addLog("Загружено: " + String(imageRefs.size) + "/" + String(uploadList.length));

      // ═══ ШАГ 2: РЕЗОЛВ ВСЕХ ШРИФТОВ (до построения элементов) ═══
      setStatus("Определяю шрифты...");
      setProgress(55);

      const uniqueFonts = new Set<string>();
      for (const obj of sorted) {
        if (obj.type === "text" && obj.textData?.fontFamily) {
          uniqueFonts.add(obj.textData.fontFamily);
        }
      }

      // Резолвим каждый уникальный шрифт
      const resolvedFonts = new Map<string, FontChoice>();

      for (const fontFamily of uniqueFonts) {
        const result = await resolveFontRef(fontFamily, fontCacheRaw, fontCacheNorm);
        if (result) {
          resolvedFonts.set(fontFamily, result);
        }
      }

      addLog("Шрифтов: " + String(uniqueFonts.size) + " уникальных, " +
        String(Array.from(resolvedFonts.values()).filter((v) => v !== "use_svg").length) + " найдено, " +
        String(Array.from(resolvedFonts.values()).filter((v) => v === "use_svg").length) + " SVG");

      // ═══ ШАГ 3: Элементы ═══
      setStatus("Подготовка...");
      setProgress(65);

      const elements: Array<any> = [];
      let svgCount = 0;
      let frameCount = 0;
      let elementIdx = 0;

      for (const obj of sorted) {

        // ═══ CLIPPING MASK → CANVA FRAME ═══
        if (obj.type === "clipMask") {
          const contentFile = obj.contentFileName || obj.fileName;
          const contentRef = imageRefs.get(contentFile);

          if (!contentRef) {
            addLog("ПРОПУСК clip: " + obj.fileName + " (нет contentRef)");
            continue;
          }

          // Генерируем SVG path для формы
          const pathD = generateShapePath(obj.clipShape, obj.width, obj.height);

          if (pathD) {
            // Shape element с image fill = Frame в Canva
            const frameElement: any = {
              type: "shape" as const,
              paths: [{
                d: pathD,
                fill: {
                  dropTarget: true,
                  asset: {
                    type: "image" as const,
                    ref: contentRef,
                  },
                },
              }],
              viewBox: {
                width: obj.width,
                height: obj.height,
                top: 0,
                left: 0,
              },
              top: obj.y,
              left: obj.x,
              width: obj.width,
              height: obj.height,
            };

            // Обводка/тень → группа с правильным позиционированием
            const strokeRef = obj.strokeFileName ? imageRefs.get(obj.strokeFileName) : undefined;
            if (strokeRef && obj.strokeBounds) {
              // strokeBounds = реальный размер stroke PNG (включает тень/обводку за пределами маски)
              const sb = obj.strokeBounds;
              // Группа = размер strokeBounds (чтобы вместить тень/обводку)
              // Frame внутри группы смещён на разницу между strokeBounds и mask bounds
              const frameOffsetX = obj.x - sb.x;
              const frameOffsetY = obj.y - sb.y;

              elements.push({
                type: "group" as const,
                children: [
                  {
                    type: "image" as const,
                    ref: strokeRef,
                    top: 0, left: 0,
                    width: sb.width, height: sb.height,
                    altText: { text: "effects", decorative: true },
                  },
                  { ...frameElement, top: frameOffsetY, left: frameOffsetX,
                    width: obj.width, height: obj.height },
                ],
                top: sb.y, left: sb.x,
                width: sb.width, height: sb.height,
              });
              addLog("⬡ Frame+Effects: " + contentFile);
            } else if (strokeRef) {
              // Fallback без strokeBounds (не должно происходить)
              elements.push({
                type: "group" as const,
                children: [
                  { ...frameElement, top: 0, left: 0 },
                  {
                    type: "image" as const,
                    ref: strokeRef,
                    top: 0, left: 0,
                    width: obj.width, height: obj.height,
                    altText: { text: "stroke", decorative: true },
                  },
                ],
                top: obj.y, left: obj.x,
                width: obj.width, height: obj.height,
              });
              addLog("⬡ Frame+Stroke: " + contentFile);
            } else {
              elements.push(frameElement);
              addLog("⬡ Frame: " + contentFile + " (" + String(obj.width) + "x" + String(obj.height) + ")");
            }
            frameCount++;
          } else {
            // Unknown shape → fallback как обычная картинка (обрезанная)
            const fallbackRef = imageRefs.get(obj.fileName);
            if (fallbackRef) {
              elements.push({
                type: "image", ref: fallbackRef,
                top: obj.y, left: obj.x,
                width: obj.width, height: obj.height,
                altText: { text: "clipped", decorative: false },
              });
              addLog("→ Clip fallback: " + obj.fileName);
            } else {
              addLog("ПРОПУСК: " + obj.fileName);
              continue;
            }
          }

          if (obj.opacity !== undefined && obj.opacity < 1) {
            opacityMap.push({ elementIndex: elementIdx, opacity: obj.opacity });
          }
          elementIdx++;
          continue;
        }

        if (obj.type === "text" && obj.textData) {
          const td = obj.textData;
          const fontResult = resolvedFonts.get(td.fontFamily);

          if (fontResult && fontResult !== "use_svg") {
            // ═══ ШРИФТ НАЙДЕН / ВЫБРАН ═══
            elements.push({
              type: "text",
              children: [td.content],
              top: obj.y, left: obj.x,
              width: Math.ceil(obj.width * TEXT_WIDTH_PADDING),
              fontSize: Math.max(1, Math.min(100, Math.round(td.fontSize))),
              color: td.color || "#000000",
              fontWeight: FONT_WEIGHT_MAP[td.fontWeight?.toLowerCase()] || "normal",
              fontStyle: td.fontStyle === "italic" ? "italic" : "normal",
              textAlign: TEXT_ALIGN_MAP[td.alignment] || "start",
              decoration: td.textDecoration === "underline" ? "underline" : "none",
              fontRef: fontResult,
            });
          } else if (obj.textSvgFileName && imageRefs.has(obj.textSvgFileName)) {
            // ═══ SVG FALLBACK ═══
            elements.push({
              type: "image",
              ref: imageRefs.get(obj.textSvgFileName)!,
              top: obj.y, left: obj.x,
              width: obj.width, height: obj.height,
              altText: { text: td.content.substring(0, 30), decorative: false },
            });
            svgCount++;
          } else {
            // ═══ НИ ШРИФТ НИ SVG ═══
            elements.push({
              type: "text",
              children: [td.content],
              top: obj.y, left: obj.x,
              width: Math.ceil(obj.width * TEXT_WIDTH_PADDING),
              fontSize: Math.max(1, Math.min(100, Math.round(td.fontSize))),
              color: td.color || "#000000",
              fontWeight: FONT_WEIGHT_MAP[td.fontWeight?.toLowerCase()] || "normal",
              fontStyle: td.fontStyle === "italic" ? "italic" : "normal",
              textAlign: TEXT_ALIGN_MAP[td.alignment] || "start",
            });
          }
        } else {
          // ═══ Картинка ═══
          const ref = imageRefs.get(obj.fileName);
          if (!ref) { addLog("ПРОПУСК: " + obj.fileName + " (нет ref)"); continue; }

          elements.push({
            type: "image",
            ref, top: obj.y, left: obj.x,
            width: obj.width, height: obj.height,
            altText: { text: obj.fileName, decorative: false },
          });
        }

        // Запоминаем opacity для openDesign
        if (obj.opacity !== undefined && obj.opacity < 1) {
          opacityMap.push({ elementIndex: elementIdx, opacity: obj.opacity });
        }
        elementIdx++;
      }

      if (svgCount > 0) addLog(String(svgCount) + " текст(ов) как SVG");
      if (frameCount > 0) addLog("⬡ " + String(frameCount) + " фреймов");

      // ═══ ШАГ 4: addPage ═══
      setStatus("Создание страницы...");
      setProgress(80);
      addLog("addPage: " + String(elements.length) + " элементов");

      let delay = 3000;
      for (let att = 1; att <= 5; att++) {
        try {
          await addPage({
            title: layout.artboardName,
            dimensions: { width: layout.width, height: layout.height },
            elements,
          });
          break;
        } catch (err: any) {
          const s = String(err);
          if ((s.includes("rate_limit") || s.includes("RATE_LIMITED")) && att < 5) {
            addLog("Rate limit → " + String(delay / 1000) + "с...");
            await new Promise((r) => setTimeout(r, delay));
            delay *= 2;
          } else throw err;
        }
      }

      addLog("Страница создана!");

      // ═══ ШАГ 5: ПРОЗРАЧНОСТЬ через openDesign() ═══
      if (opacityMap.length > 0) {
        addLog("Применяю прозрачность (" + String(opacityMap.length) + " элементов)...");
        setStatus("Настройка прозрачности...");

        try {
          await new Promise((r) => setTimeout(r, 1000));

          await openDesign(
            { type: "current_page" },
            async (session) => {
              if (session.page.type !== "absolute" || session.page.locked) {
                addLog("Страница locked/не absolute — пропуск opacity");
                return;
              }

              const pageElements = session.page.elements.toArray();
              addLog("openDesign: " + String(pageElements.length) + " элементов");

              for (const { elementIndex, opacity } of opacityMap) {
                if (elementIndex < pageElements.length) {
                  const el = pageElements[elementIndex];
                  if (el && el.type !== "unsupported" && !el.locked) {
                    const transparency = 1 - opacity;
                    el.transparency = transparency;
                    addLog("  [" + String(elementIndex) + "] transparency=" + String(Math.round(transparency * 100)) + "%");
                  }
                }
              }

              await session.sync();
              addLog("Прозрачность применена!");
            }
          );
        } catch (opErr) {
          addLog("openDesign ошибка: " + String(opErr));
          addLog("Прозрачность выставьте вручную:");
          for (const { elementIndex, opacity } of opacityMap) {
            addLog("  Элемент #" + String(elementIndex) + ": " + String(Math.round(opacity * 100)) + "%");
          }
        }
      }

      setProgress(100);
      setStatus("Готово: " + layout.artboardName);

    } catch (error) {
      setStatus("Ошибка");
      addLog("ОШИБКА: " + String(error));
      console.error(error);
    } finally {
      setIsImporting(false);
    }
  };

  // ═══ Импорт всех ═══
  const handleImportAll = async () => {
    setIsImporting(true);
    setImportLog([]);
    addLog("Импорт " + String(layouts.length) + " артбордов...");
    for (let i = 0; i < layouts.length; i++) {
      setSelectedLayoutIndex(i);
      await handleImportLayout(i);
      if (i < layouts.length - 1) {
        addLog("Пауза 5с...");
        await new Promise((r) => setTimeout(r, 5000));
      }
    }
    setStatus(String(layouts.length) + " страниц!");
    setIsImporting(false);
  };

  const cur = layouts.length > 0 ? layouts[selectedLayoutIndex] : null;
  const choicesArr = Array.from(userFontChoicesRef.current.entries());

  return (
    <div style={{ height: "100%", overflowY: "auto", padding: "16px" }}>
      <Rows spacing="2u">
        <Title size="medium">Illustrator Importer v5.0</Title>

        {layouts.length === 0 && (
          <Button variant="primary" onClick={handleFolderSelect} disabled={isImporting}>
            Выбрать папку
          </Button>
        )}

        {layouts.length > 1 && (
          <div style={{ background: "#e8f0fe", borderRadius: "8px", padding: "12px", margin: "4px 0" }}>
            <Text size="small"><strong>{`Артборды (${String(layouts.length)}):`}</strong></Text>
            {layouts.map((l, idx) => (
              <div key={idx} onClick={() => setSelectedLayoutIndex(idx)} style={{
                padding: "8px", margin: "4px 0", borderRadius: "4px", cursor: "pointer",
                background: idx === selectedLayoutIndex ? "#1a73e8" : "#fff",
                border: "1px solid #dadce0",
              }}>
                <span style={{ fontSize: "12px", color: idx === selectedLayoutIndex ? "#fff" : "#1a1a1a" }}>
                  {`${String(idx + 1)}. ${l.artboardName} (${String(l.width)}x${String(l.height)}, ${String(l.objects.length)} obj)`}
                </span>
              </div>
            ))}
          </div>
        )}

        {/* ═══ ДИАЛОГ ВЫБОРА ШРИФТА ═══ */}
        {pendingFontResolve && (
          <div style={{
            background: "#fff3cd", border: "2px solid #ffc107",
            borderRadius: "8px", padding: "16px", margin: "4px 0",
          }}>
            <Text size="small">
              <strong>Шрифт не найден: «{pendingFontName}»</strong>
            </Text>
            <Text size="xsmall">
              Выберите замену из библиотеки Canva или импортируйте как SVG-картинку (сохранит внешний вид, но текст нельзя будет редактировать).
            </Text>
            <div style={{ display: "flex", gap: "8px", marginTop: "12px" }}>
              <Button variant="primary" onClick={handlePickFont}>
                Выбрать шрифт
              </Button>
              <Button variant="secondary" onClick={handleUseSvg}>
                SVG (картинка)
              </Button>
            </div>
          </div>
        )}

        {cur && (
          <>
            <div style={{ background: "#f0f0f0", borderRadius: "8px", padding: "12px" }}>
              <Text size="small"><strong>{cur.artboardName}</strong></Text>
              <Text size="small">{`${String(cur.width)}x${String(cur.height)}px, ${String(cur.objects.length)} объектов`}</Text>
              {cur.objects.some((o) => o.opacity < 1) && (
                <Text size="small">{`С прозрачностью: ${String(cur.objects.filter((o) => o.opacity < 1).length)}`}</Text>
              )}
              {cur.objects.some((o) => o.type === ("clipMask" as any)) && (
                <Text size="small">{`⬡ Фреймов: ${String(cur.objects.filter((o) => o.type === ("clipMask" as any)).length)}`}</Text>
              )}
            </div>

            {isImporting && (
              <div style={{ width: "100%", height: "8px", background: "#e0e0e0", borderRadius: "4px", overflow: "hidden" }}>
                <div style={{
                  height: "100%", background: "linear-gradient(90deg,#6366f1,#8b5cf6)",
                  transition: "width 0.3s", borderRadius: "4px", width: `${String(progress)}%`,
                }} />
              </div>
            )}

            <Text size="small">{status}</Text>

            {!pendingFontResolve && (
              <>
                <Button variant="primary" onClick={() => handleImportLayout(selectedLayoutIndex)} disabled={isImporting}>
                  {isImporting ? "Импортирую..." : layouts.length > 1
                    ? `Импортировать "${cur.artboardName}"` : "Импортировать"}
                </Button>

                {layouts.length > 1 && !isImporting && (
                  <Button variant="secondary" onClick={handleImportAll}>
                    {`Все (${String(layouts.length)} стр.)`}
                  </Button>
                )}
              </>
            )}

            {/* Кэш выбранных замен */}
            {choicesArr.length > 0 && !isImporting && (
              <div style={{ background: "#e8f5e9", borderRadius: "8px", padding: "10px", margin: "4px 0" }}>
                <Text size="xsmall"><strong>Замены шрифтов (кэш):</strong></Text>
                {choicesArr.map(([font, choice], idx) => (
                  <Text key={idx} size="xsmall">
                    {font} → {choice === "use_svg" ? "SVG" : "✓ выбран"}
                  </Text>
                ))}
                <div style={{ marginTop: "6px" }}>
                  <Button variant="tertiary" onClick={() => {
                    userFontChoicesRef.current = new Map();
                    setShowFontChoices((p) => !p);
                    addLog("Кэш шрифтов очищен");
                  }}>
                    Сбросить замены
                  </Button>
                </div>
              </div>
            )}

            {importLog.length > 0 && (
              <div style={{ background: "#1e1e1e", borderRadius: "8px", padding: "10px", maxHeight: "250px", overflowY: "auto" }}>
                {importLog.map((msg, idx) => (
                  <div key={idx} style={{
                    fontSize: "11px",
                    color: msg.includes("ОШИБКА") || msg.includes("FAIL") || msg.includes("НЕТ")
                      ? "#ff6b6b"
                      : msg.startsWith("✓") ? "#69db7c"
                      : msg.startsWith("⬡") ? "#a78bfa"
                      : msg.startsWith("?") ? "#ffd93d"
                      : msg.includes("→") ? "#74c0fc"
                      : "#e0e0e0",
                    lineHeight: "1.5", fontFamily: "monospace",
                  }}>
                    {msg}
                  </div>
                ))}
              </div>
            )}

            {!isImporting && !pendingFontResolve && (
              <Button variant="secondary" onClick={() => {
                setLayouts([]); setSelectedLayoutIndex(0); setFiles(new Map());
                setArtboardFileMap(new Map()); setStatus("Готов"); setProgress(0);
                setImportLog([]); userFontChoicesRef.current = new Map();
              }}>
                Другая папка
              </Button>
            )}
          </>
        )}

        <div style={{ background: "#f0f7ff", borderLeft: "4px solid #3b82f6", borderRadius: "4px", padding: "12px", marginTop: "8px" }}>
          <Text size="xsmall"><strong>v5.0:</strong></Text>
          <Text size="xsmall">1. JSX: SortObjectsToArtboards.jsx → экспорт</Text>
          <Text size="xsmall">2. _fonts → Brand Kit, шрифт не найден → пикер/SVG</Text>
          <Text size="xsmall">3. Clipping Mask → Frame (замена фото!)</Text>
          <Text size="xsmall">4. Обводка/тень → группа с фреймом</Text>
          <Text size="xsmall">5. Масштаб экспорта, прозрачность авто</Text>
        </div>
      </Rows>
    </div>
  );
}