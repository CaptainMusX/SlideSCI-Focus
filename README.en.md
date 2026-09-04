<div align="center">

# SlideSCI Focus

**A focused PowerPoint add-in for scientific figure workflows** (VSTO, Windows)

</div>

> **SlideSCI Focus** is a community-modified version of [Achuan-2/SlideSCI](https://github.com/Achuan-2/SlideSCI):
> it focuses on scientific figure workflows, removes the AI assistant and shape-library sidebars,
> adds a "Magnified Inset" feature, and includes stability hardening and an installer pipeline rewrite.

## ✨ Features

## 🧭 Ribbon tabs

- **SciFigure**: image arrangement, captions and panel labels, magnified insets, formatting, sizing, and export
- **SciStudio**: Markdown, LaTeX, code, text formatting, selection, and general tools

### 🔍 Magnified Inset (journal figure style)
Create the classic "zoom-in" panel used in Cell/Nature-style figures from microscopy/EM images:

- **"Insert selection box"**: places a square selection box at the center of the selected image;
  drag/resize it freely to frame the exact region to magnify
- **"Generate magnified image"**: opens a settings dialog and generates the result in one click:
  - Target size: same as original image / specified magnification factor / custom width (cm)
  - Connector style: **journal funnel** (box bottom corners → inset top corners, parallel
    non-corner lines) or **crossed X** (four corners crossed)
  - Connector/frame line: color (including custom), weight, solid/dashed
  - Live preview of the resulting dimensions and the effective magnification; warns when the
    box aspect ratio does not match the target frame (stretch warning)
- **Snapped connectors**: line endpoints are glued to the box corners and the inset corners
  (via invisible corner anchors) using native PowerPoint connection points — moving the image
  or the box later re-stretches the lines automatically, no manual rework
- **Pixel-lossless**: only vector operations are used (duplicate, crop, scale, connect, group);
  no file export, no image re-encoding — the original pixels stay untouched
- Re-generating automatically replaces the old inset and lines while keeping your box

### 🖼️ Image tools
- **Auto image arrangement**: 3 sort modes × 3 layout modes (column-width grid / uniform height /
  waterfall), with per-row/column spacing and image width & height in cm
- **Batch captions** for images (above/below, font/size/offset/centering/auto-group)
- **Batch subfigure labels** (A, a, A), 1, Ⅰ, ① … 14 templates, offsets, bold, auto numbering)
- **Export slide / export original image / copy large image**: batch high-DPI image & PDF export;
  original images are read directly from the package (OpenXML) without quality loss

### 📋 Format & layout
- Copy/paste **shape format**, **text format** (font/color/size/effects), **group format**
- Copy/paste **position** (9 anchor points, multi-select), **swap positions**
- Copy/paste **width, height, picture crop** (multi-select)
- **Vertical/horizontal center** relative to the first selected object (Illustrator-style),
  and **spacing/distribution**

### 📝 Markdown / LaTeX / code
- **Insert Markdown**: paste a whole note as slides in order — headings, lists (with hanging
  indent), task lists, tables, block quotes, inline styles, inline & block math
- **Textbox to rich text**: convert a textbox containing Markdown into rich text in place
- **Insert LaTeX text**: native PowerPoint equations (OMML), good for simple formulas
- **Insert LaTeX SVG**: MathJax-based conversion for complex formulas (requires Node.js, see below)
- **Insert code block**: syntax highlighting for 8 languages (matlab/python/r/js/html/css/
  csharp/fortran), black/white background toggle

## 🪟 Requirements

- Windows + Microsoft PowerPoint (VSTO add-in)
- WPS is installable but does not support Markdown/LaTeX insertion (may freeze)
- macOS is not supported

## 📥 Installation

1. Download the installer from the [Releases](https://github.com/CaptainMusX/SlideSCI-Focus/releases) page
2. **Quit PowerPoint first**, then run the installer
3. Dependencies: .NET Framework 4.7.2 and Microsoft Visual Studio 2010 Tools for Office Runtime
   (the installer prompts automatically)

> If the add-in does not appear: Developer → COM Add-ins → check `CaptainMusX.SlideSCI.Focus`.
> If it reports "Unhandled error", install the runtime dependencies above and restart PowerPoint.

## 🔧 Development & build

- Visual Studio + Visual Studio Tools For Office; open `SlideSCI-Focus.sln`
- Build the installer from the command line: `pwsh -File .\build\Build-Installer.ps1`
  (see `build/README.md`)
- For "Insert LaTeX SVG", a Node.js runtime is required:
  ```
  cd <addin-dir>/latex-converter
  npm install
  ```
  or bundle MathJax into the installer with `-BundleLatexRuntime` (not bundled by default)

## ❓ FAQ

- **Add buttons to the Quick Access Toolbar?** Right-click a button → Add to Quick Access Toolbar
- **LaTeX renders incorrectly?** Use "Insert LaTeX text" for simple formulas and
  "Insert LaTeX SVG" for complex ones
- **Connectors no longer follow the image?** They are glued to the box and the inset corners;
  if you deleted the selection box, re-insert it and regenerate

## 📄 License & notice

- This project is a modified version of [Achuan-2/SlideSCI](https://github.com/Achuan-2/SlideSCI) (AGPL-3.0),
  modified by CaptainMusX; see `CHANGELOG.md` for dates and `NOTICE.md` for attribution
- The combined project is released under GNU AGPL-3.0. When distributing binaries, provide the corresponding
  source code and build instructions; this project is not an official release by the original author

## 💬 Feedback

- GitHub Issues: https://github.com/CaptainMusX/SlideSCI-Focus/issues
