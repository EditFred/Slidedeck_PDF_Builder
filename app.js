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
  ["deckTitle", "slideCount", "slidesForm", "addSlideBtn", "exportPdfBtn", "exportPptxBtn"].forEach((id) => {
    els[id] = document.getElementById(id);
  });
}

function bindEvents() {
  els.deckTitle.addEventListener("input", (event) => {
    deckState.title = event.target.value;
    updateBoundTitlePlaceholders();
  });
  els.addSlideBtn.addEventListener("click", addSlide);
  els.exportPdfBtn.addEventListener("click", exportPdf);
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
    card.querySelector(".spanish-input").addEventListener("input", (event) => {
      updateSlide(slide.id, { spanish: event.target.value });
    });

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
    page.addText(slide.english || " ", {
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
    page.addText(slide.spanish || " ", {
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

async function exportPdf() {
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
      const lines = pdf.splitTextToSize(slide.english || " ", maxWidth);
      pdf.text(lines, marginX, y);
      y += lines.length * 0.2 + 0.38;
    });

    pdf.save(`${fileSafe(deckState.title)}.pdf`);
  } finally {
    setBusy(false);
  }
}

function hexToRgb(hex) {
  return [0, 2, 4].map((start) => parseInt(hex.slice(start, start + 2), 16));
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
