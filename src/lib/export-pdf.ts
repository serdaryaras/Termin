import { toPng } from "html-to-image";
import { jsPDF } from "jspdf";

/** CSS px → mm (96 dpi) */
function pxToMm(px: number): number {
  return (px * 25.4) / 96;
}

/**
 * Tailwind v4 / modern CSS (oklch, color-mix) SVG foreignObject içinde
 * bozulduğu için yakalamadan önce hesaplanmış RGB değerlerini inline yazar.
 */
function inlineComputedColors(root: HTMLElement): () => void {
  const touched: { el: HTMLElement; cssText: string }[] = [];
  const nodes = [root, ...Array.from(root.querySelectorAll<HTMLElement>("*"))];

  for (const el of nodes) {
    touched.push({ el, cssText: el.style.cssText });
    const cs = getComputedStyle(el);
    el.style.setProperty("color", cs.color);
    el.style.setProperty("background-color", cs.backgroundColor);
    el.style.setProperty("border-top-color", cs.borderTopColor);
    el.style.setProperty("border-right-color", cs.borderRightColor);
    el.style.setProperty("border-bottom-color", cs.borderBottomColor);
    el.style.setProperty("border-left-color", cs.borderLeftColor);
    el.style.setProperty("outline-color", cs.outlineColor);
    el.style.setProperty("text-decoration-color", cs.textDecorationColor);
    el.style.setProperty("caret-color", cs.caretColor);
    el.style.setProperty("column-rule-color", cs.columnRuleColor);
    if (cs.boxShadow && cs.boxShadow !== "none") {
      el.style.setProperty("box-shadow", cs.boxShadow);
    }
    if (cs.backgroundImage && cs.backgroundImage !== "none") {
      el.style.setProperty("background-image", cs.backgroundImage);
    }
  }

  return () => {
    for (const { el, cssText } of touched) {
      el.style.cssText = cssText;
    }
  };
}

/**
 * Gantt alanını tek sayfalık özel boyutlu PDF olarak indirir.
 * Kağıt boyutu, yakalanan görselin en/boy oranına göre ayarlanır.
 */
export async function exportElementToPdf(
  element: HTMLElement,
  fileName: string
): Promise<void> {
  const width = Math.max(element.scrollWidth, element.offsetWidth, 1);
  const height = Math.max(element.scrollHeight, element.offsetHeight, 1);

  const maxEdge = 8192;
  const scale = Math.min(2, maxEdge / width, maxEdge / height);

  const restore = inlineComputedColors(element);
  let dataUrl: string;
  try {
    dataUrl = await toPng(element, {
      cacheBust: true,
      pixelRatio: scale,
      backgroundColor: "#ffffff",
      width,
      height,
      style: {
        transform: "none",
        overflow: "visible",
      },
      filter: (node) => {
        if (!(node instanceof HTMLElement)) return true;
        return !node.classList.contains("no-print");
      },
    });
  } catch (err) {
    const msg = err instanceof Error ? err.message : String(err);
    throw new Error(`Görsel yakalanamadı: ${msg}`);
  } finally {
    restore();
  }

  if (!dataUrl || dataUrl === "data:," ) {
    throw new Error("Görsel boş geldi; Gantt alanını kontrol edin.");
  }

  const img = new Image();
  await new Promise<void>((resolve, reject) => {
    img.onload = () => resolve();
    img.onerror = () => reject(new Error("PDF görseli yüklenemedi."));
    img.src = dataUrl;
  });

  const contentW = pxToMm(img.width / scale);
  const contentH = pxToMm(img.height / scale);
  const margin = 6;
  const pageW = Math.max(contentW + margin * 2, 50);
  const pageH = Math.max(contentH + margin * 2, 50);

  const maxMm = 5000;
  const fit = Math.min(1, maxMm / pageW, maxMm / pageH);
  const finalPageW = pageW * fit;
  const finalPageH = pageH * fit;
  const finalContentW = contentW * fit;
  const finalContentH = contentH * fit;
  const finalMargin = margin * fit;

  try {
    const pdf = new jsPDF({
      orientation: finalPageW >= finalPageH ? "landscape" : "portrait",
      unit: "mm",
      format: [finalPageW, finalPageH],
      compress: true,
    });

    pdf.addImage(
      dataUrl,
      "PNG",
      finalMargin,
      finalMargin,
      finalContentW,
      finalContentH,
      undefined,
      "FAST"
    );
    pdf.save(fileName.endsWith(".pdf") ? fileName : `${fileName}.pdf`);
  } catch (err) {
    const msg = err instanceof Error ? err.message : String(err);
    throw new Error(`PDF yazılamadı: ${msg}`);
  }
}
