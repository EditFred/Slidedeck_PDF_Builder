# ProPresenter Slide Builder

A lightweight single-page form for creating 16:9 decks that can be exported as PDF or PowerPoint for ProPresenter import workflows.

## Use

Open `index.html` directly, or serve the folder with any static server.

```bash
python3 -m http.server 5173
```

Then visit:

```text
http://127.0.0.1:5173
```

## Hosting

This is a static site. GitHub Pages can host the folder as-is.

The centered main title field uses `Title here` as placeholder text and sets the default title for every slide. Each slide can still override that title when needed.

Each slide has three fields:

- Title
- English text
- Spanish text

The exports apply the slide formatting:

- Title: 80 pt
- English text: 44 pt, left side
- Spanish text: 44 pt, right side
- PPTX text color: white
- PPTX size: 16:9, equivalent to 1920x1080
- PPTX background: fully transparent background requested; some PowerPoint viewers still display transparent slide canvases as white
- PDF: main title, then red slide labels, optional custom slide titles, and black English text only

The export buttons use browser-loaded CDN libraries:

- `jsPDF`
- `PptxGenJS`

If you need fully offline exports later, bundle those dependencies locally instead of loading them from CDN.
