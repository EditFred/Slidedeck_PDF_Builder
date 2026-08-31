const SLIDE_W = 13.333;
const SLIDE_H = 7.5;
const WHITE = "FFFFFF";
const BLACK = "000000";
const RED = "D92D20";

const deckState = {
  title: "",
  slides: [
    {
      id: "slide-1",
      titleOverride: "",
      english: "",
      spanish: "",
    },
  ],
};

const els = {};

window.addEventListener("DOMContentLoaded", () => {
  bindElements();
  bindEvents();
  render();
});

function bindElements() {
  ["deckTitle", "slideCount", "slidesForm", "addSlideBtn", "exportPdfBtn", "exportEspPdfBtn", "exportPptxBtn"].forEach(
    (id) => {
      els[id] = document.getElementById(id);
    },
  );
}

function bindEvents() {
  els.deckTitle.addEventListener("input", (event) => {
    deckState.title = event.target.value;
    updateBoundTitlePlaceholders();
  });
  els.addSlideBtn.addEventListener("click", addSlide);
  els.exportPdfBtn.addEventListener("click", () => exportPdf("english"));
  els.exportEspPdfBtn.addEventListener("click", () => exportPdf("spanish"));
  els.exportPptxBtn.addEventListener("click", exportPptx);
}

function addSlide() {
  deckState.slides.push({
    id: createId(),
    titleOverride: "",
    english: "",
    spanish: "",
  });
  render();
}

function deleteSlide(id) {
  if (deckState.slides.length === 1) return;
  deckState.slides = deckState.slides.filter((slide) => slide.id !== id);
  render();
}

function updateSlide(id, patch) {
  const slide = deckState.slides.find((item) => item.id === id);
  if (!slide) return;
  Object.assign(slide, patch);
}

function render() {
  els.deckTitle.value = deckState.title;
  els.slideCount.textContent = `${deckState.slides.length} slide${deckState.slides.length === 1 ? "" : "s"}`;
  els.slidesForm.innerHTML = "";

  deckState.slides.forEach((slide, index) => {
    const card = document.createElement("article");
    card.className = "slide-card";
    card.innerHTML = `
      <div class="slide-card-header">
        <div class="slide-number">Slide ${index + 1}</div>
        <button class="danger" type="button" ${deckState.slides.length === 1 ? "disabled" : ""}>Delete</button>
      </div>
      <label>
        Title
        <input
          class="title-input"
          type="text"
          autocomplete="off"
          placeholder="${escapeAttr(deckState.title)}"
          value="${escapeAttr(slide.titleOverride)}"
        />
      </label>
      <div class="language-grid">
        <label>
          English text
          <textarea class="english-input" placeholder="English text goes here.">${escapeHtml(slide.english)}</textarea>
        </label>
        <label>
          Spanish text
          <textarea class="spanish-input" placeholder="El texto en espanol va aqui.">${escapeHtml(slide.spanish)}</textarea>
        </label>
      </div>
    `;

    card.querySelector(".danger").addEventListener("click", () => deleteSlide(slide.id));
    card.querySelector(".title-input").addEventListener("input", (event) => {
      updateSlide(slide.id, { titleOverride: event.target.value });
    });
    card.querySelector(".english-input").addEventListener("input", (event) => {
      updateSlide(slide.id, { english: event.target.value });
    });
    card.querySelector(".english-input").addEventListener("keydown", handleTextAreaKeydown);
    card.querySelector(".spanish-input").addEventListener("input", (event) => {
      updateSlide(slide.id, { spanish: event.target.value });
    });
    card.querySelector(".spanish-input").addEventListener("keydown", handleTextAreaKeydown);

    els.slidesForm.appendChild(card);
  });
}

function updateBoundTitlePlaceholders() {
  document.querySelectorAll(".title-input").forEach((input) => {
    input.placeholder = deckState.title;
  });
}

async function exportPptx() {
  const PptxGen = window.PptxGenJS || window.pptxgen;
  if (!PptxGen) {
    alert("The PowerPoint export tools are still loading. Try again in a moment.");
    return;
  }

  const pptx = new PptxGen();
  pptx.defineLayout({ name: "PROPRESENTER_1920_1080", width: SLIDE_W, height: SLIDE_H });
  pptx.layout = "PROPRESENTER_1920_1080";
  pptx.author = "ProPresenter Slide Builder";
  pptx.subject = deckState.title;
  pptx.title = deckState.title;
  pptx.lang = "en-US";
  pptx.theme = {
    headFontFace: "Arial",
    bodyFontFace: "Arial",
    lang: "en-US",
  };

  deckState.slides.forEach((slide) => {
    const title = slideTitle(slide);
    const page = pptx.addSlide();
    page.background = { color: WHITE, transparency: 100 };
    page.addText(title || " ", {
      x: 0.6,
      y: 0.45,
      w: 12.15,
      h: 1.1,
      margin: 0,
      fontFace: "Arial",
      fontSize: 80,
      bold: true,
      color: WHITE,
      fit: "shrink",
      align: "center",
      valign: "mid",
    });
    page.addText(formatBulletsForExport(slide.english) || " ", {
      x: 0.72,
      y: 2.05,
      w: 5.8,
      h: 4.9,
      margin: 0,
      fontFace: "Arial",
      fontSize: 44,
      color: WHITE,
      fit: "shrink",
      breakLine: false,
      valign: "top",
    });
    page.addText(formatBulletsForExport(slide.spanish) || " ", {
      x: 6.82,
      y: 2.05,
      w: 5.8,
      h: 4.9,
      margin: 0,
      fontFace: "Arial",
      fontSize: 44,
      color: WHITE,
      fit: "shrink",
      breakLine: false,
      valign: "top",
    });
  });

  await pptx.writeFile({ fileName: `${fileSafe(deckState.title)}.pptx` });
}

async function exportPdf(language) {
  if (!window.jspdf) {
    alert("The PDF export tools are still loading. Try again in a moment.");
    return;
  }

  setBusy(true);
  try {
    const { jsPDF } = window.jspdf;
    const pdf = new jsPDF({ orientation: "portrait", unit: "in", format: "letter" });
    const marginX = 0.75;
    const maxWidth = 7;
    let y = 0.8;

    pdf.setTextColor(...hexToRgb(BLACK));
    pdf.setFont("helvetica", "bold");
    pdf.setFontSize(22);
    pdf.text(deckState.title || "Title here", marginX, y, { maxWidth });
    y += 0.65;

    deckState.slides.forEach((slide, index) => {
      const customTitle = slide.titleOverride.trim();
      y = ensurePdfSpace(pdf, y, customTitle ? 1.35 : 1.1);

      pdf.setTextColor(...hexToRgb(RED));
      pdf.setFont("helvetica", "bold");
      pdf.setFontSize(13);
      pdf.text(`Slide ${index + 1}`, marginX, y);
      y += 0.3;

      if (customTitle) {
        pdf.setTextColor(...hexToRgb(BLACK));
        pdf.setFont("helvetica", "bold");
        pdf.setFontSize(12);
        const titleLines = pdf.splitTextToSize(customTitle, maxWidth);
        pdf.text(titleLines, marginX, y);
        y += titleLines.length * 0.2 + 0.12;
      }

      pdf.setTextColor(...hexToRgb(BLACK));
      pdf.setFont("helvetica", "normal");
      pdf.setFontSize(12);
      y = drawPdfTextBlock(pdf, formatBulletsForExport(slide[language]) || " ", marginX, y, maxWidth);
      y += 0.38;
    });

    const suffix = language === "spanish" ? "-ESP" : "";
    pdf.save(`${fileSafe(deckState.title)}${suffix}.pdf`);
  } finally {
    setBusy(false);
  }
}

function hexToRgb(hex) {
  return [0, 2, 4].map((start) => parseInt(hex.slice(start, start + 2), 16));
}

function drawPdfTextBlock(pdf, text, x, y, maxWidth) {
  const lineHeight = 0.2;
  const bulletGap = 0.18;

  String(text)
    .split("\n")
    .forEach((rawLine) => {
      const bullet = rawLine.match(/^(\s*)([•◦])\s+(.*)$/);
      const indentLevel = bullet ? Math.min(Math.floor(bullet[1].length / 2), 1) : 0;
      const marker = bullet?.[2];
      const content = bullet ? bullet[3] : rawLine;
      const textX = bullet ? x + indentLevel * 0.35 + bulletGap : x;
      const availableWidth = bullet ? maxWidth - indentLevel * 0.35 - bulletGap : maxWidth;
      const lines = pdf.splitTextToSize(content || " ", availableWidth);

      lines.forEach((line, lineIndex) => {
        y = ensurePdfSpace(pdf, y, lineHeight);

        if (marker && lineIndex === 0) {
          const bulletX = x + indentLevel * 0.35 + 0.04;
          const bulletY = y - 0.035;
          pdf.setFillColor(...hexToRgb(BLACK));
          pdf.circle(bulletX, bulletY, marker === "◦" ? 0.025 : 0.035, "F");
        }

        pdf.setTextColor(...hexToRgb(BLACK));
        pdf.text(line, textX, y);
        y += lineHeight;
      });
    });

  return y;
}

function handleTextAreaKeydown(event) {
  if (event.key === "*" && isAtLineStart(event.target)) {
    event.preventDefault();
    replaceSelection(event.target, "• ");
    event.target.dispatchEvent(new Event("input", { bubbles: true }));
    return;
  }

  if (event.key === "Tab" && isBulletLine(event.target)) {
    event.preventDefault();
    event.shiftKey ? outdentBulletLine(event.target) : indentBulletLine(event.target);
    event.target.dispatchEvent(new Event("input", { bubbles: true }));
    return;
  }

  if (event.key === "Enter" && isBulletLine(event.target)) {
    handleBulletEnter(event);
    return;
  }

  if (event.key === "Backspace" && isAtSubBulletContentStart(event.target)) {
    event.preventDefault();
    outdentBulletLine(event.target);
    event.target.dispatchEvent(new Event("input", { bubbles: true }));
  }
}

function isAtLineStart(textarea) {
  const lineStart = textarea.value.lastIndexOf("\n", textarea.selectionStart - 1) + 1;
  return textarea.value.slice(lineStart, textarea.selectionStart).trim() === "";
}

function isBulletLine(textarea) {
  return /^\s*[•◦]\s/.test(currentLine(textarea).text);
}

function isAtSubBulletContentStart(textarea) {
  if (textarea.selectionStart !== textarea.selectionEnd) return false;
  const line = currentLine(textarea);
  const marker = line.text.match(/^\s*◦\s/);
  return Boolean(marker) && textarea.selectionStart === line.start + marker[0].length;
}

function indentBulletLine(textarea) {
  const line = currentLine(textarea);
  const nextLine = line.text.replace(/^\s*[•◦]\s/, "  ◦ ");
  replaceRange(textarea, line.start, line.end, nextLine);
}

function outdentBulletLine(textarea) {
  const line = currentLine(textarea);
  const nextLine = line.text.replace(/^\s*◦\s/, "• ");
  replaceRange(textarea, line.start, line.end, nextLine);
}

function handleBulletEnter(event) {
  const textarea = event.target;
  const line = currentLine(textarea);
  const marker = line.text.match(/^(\s*[•◦]\s)/)?.[1] ?? "• ";
  const content = line.text.replace(/^\s*[•◦]\s/, "").trim();

  event.preventDefault();
  if (!content) {
    replaceRange(textarea, line.start, line.end, "");
  } else {
    replaceSelection(textarea, `\n${marker}`);
  }
  textarea.dispatchEvent(new Event("input", { bubbles: true }));
}

function currentLine(textarea) {
  const value = textarea.value;
  const start = value.lastIndexOf("\n", textarea.selectionStart - 1) + 1;
  const nextBreak = value.indexOf("\n", textarea.selectionStart);
  const end = nextBreak === -1 ? value.length : nextBreak;
  return { start, end, text: value.slice(start, end) };
}

function replaceSelection(textarea, text) {
  replaceRange(textarea, textarea.selectionStart, textarea.selectionEnd, text);
}

function replaceRange(textarea, start, end, text) {
  const before = textarea.value.slice(0, start);
  const after = textarea.value.slice(end);
  textarea.value = `${before}${text}${after}`;
  textarea.selectionStart = start + text.length;
  textarea.selectionEnd = start + text.length;
}

function formatBulletsForExport(text) {
  return String(text ?? "")
    .split("\n")
    .map((line) => line.replace(/^(\s*)\*\s+/, "$1• "))
    .join("\n");
}

function ensurePdfSpace(pdf, y, neededHeight) {
  if (y + neededHeight <= 10.35) return y;
  pdf.addPage("letter", "portrait");
  return 0.75;
}

function slideTitle(slide) {
  return slide.titleOverride.trim() || deckState.title.trim();
}

function setBusy(isBusy) {
  document.body.classList.toggle("busy", isBusy);
}

function fileSafe(value) {
  return (value || "slide-deck").trim().replace(/[^a-z0-9-_]+/gi, "-").replace(/^-+|-+$/g, "") || "slide-deck";
}

function createId() {
  if (crypto?.randomUUID) return crypto.randomUUID();
  return `slide-${Date.now()}-${Math.random().toString(16).slice(2)}`;
}

function escapeHtml(value) {
  return String(value ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;");
}

function escapeAttr(value) {
  return escapeHtml(value).replaceAll('"', "&quot;");
}
