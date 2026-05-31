# 📊 Cells — Premium Client-Side Spreadsheet Editor

<div align="center">

![Cairo English/Arabic Typography](https://img.shields.io/badge/Cairo%20Font-Bilingual-06b6d4?style=for-the-badge)
![100% Client-Side](https://img.shields.io/badge/Security-100%25%20Private%20Offline-10b981?style=for-the-badge)
![Auto-Recovery Timelines](https://img.shields.io/badge/Data%20Safety-Auto--Recovery-3b82f6?style=for-the-badge)
![Excel Selection Parity](https://img.shields.io/badge/Parity-Microsoft%20Excel-8b5cf6?style=for-the-badge)

<p align="center">
  A high-end, responsive spreadsheet web application built entirely on client-side technologies. Operating with zero server-side dependencies, <strong>Cells</strong> guarantees absolute private workflows, storing all sheet contents, typography configurations, custom tab colors, and histories directly on your local machine.
</p>

</div>

---

## ✨ Primary Features

### 🏁 Excel-Parity Selection & Keyboard Navigation
- **Marching Ants Clipboard**: Copying ranges (`Ctrl + C`) paints a custom pulsing green dashed marching ants outline. Clears instantly on edit, deselect, or `Escape`.
- **Excel Selection Contiguity**: Full mouse dragging, `Shift + Click` range anchors, and `Shift + Arrow` selections are fully supported with visual boundaries.
- **F2 Inline Focus**: Tapping `F2` focuses the inline cell editor, positioning the cursor right at the end of the text. Navigating and typing Alphanumeric keys immediately overwrites cell values.

### 📋 Custom Spatial Clipboard Tools & Series Autofill
- **Autofill series handle**: Dragging the cell bottom-right crosshair autofills repeating patterns, steps (linear regression for numbers like `10, 20` -> `30, 40`), and copies styles or formulas.
- **Spatial copy-paste operations**: Split, vertically copy, or horizontally copy values by space separators or delimiters. Pasting repeating blocks into large ranges distributes patterns automatically.

### 🌐 Cross-Sheet Formulas & Autocomplete Suggestions
- **Cross-Sheet calculations**: Formulas resolve sheets case-insensitively (e.g. `=Sheet2!A1+sHeEt3!B5`).
- **Safety checks**: Supports recursive formulas, circular reference protection (`#CIRC!`), and evaluation depth recursion limits (`#LIMIT!`) to prevent thread locks.
- **Suggestions panel**: Toggling `=` inside the formula bar reveals a glassmorphic autocompleter. Injects formulas and focuses parameters selection ranges (`A1:A10`) instantly.

### 🎨 Curated Swatches, Themes, & Bilingual Typography
- **Bilingual Font Cairo**: Cairo Google Font matches Arabic and Latin text beautifully within spreadsheet cells. Supports custom local system font loading and font sizing (`10px` to `24px`).
- **Curated swatches**: Right-clicking workbook tabs opens swatches to customize color strips. Features sheet duplication, step-by-step movement, and confirmation-guarded sheet sweeps.
- **LTR/RTL Directions**: Toggling per-sheet direction pivots layouts, row headers, resize margins, and autofill handles seamlessly.

---

## 🚀 Newly Added Advanced Systems

### ⌨️ Interactive Keyboard Shortcuts Modal
* Discover and learn high-speed grid keystrokes using a beautiful glassmorphic modal box.
* **Filter Search Bar**: Type keywords (e.g. "copy") to immediately hide unrelated shortcuts in real-time.
* **Global Access**: Press `?` (Shift + `/`) when not editing cells or click `❓` in the header actions.

### 🔍 Find and Replace Sliding Panel
* Slide-out panel matching Chart Studio (accordion-style).
* **Advanced Filters**: Match Case, Entire Cell, and Search All Sheets.
* **Double Highlights**: Matching cells light up with amber outlines (`.search-match`), and the active index hit glows in pulsing emerald (`.search-match-active`).
* **Instant Click Focus**: Matches render in a scrollable sidebar list. Clicking a row automatically switches sheets, scroll-focuses (`scrollIntoView`), and selects the matching cell.

### 🛡️ Workbook Auto-Recovery & Backup Timelines
* **Intelligent Auto-Saves**: Saves complete workbook state snapshots to `localStorage` every 30 seconds, skipping saves if state hashes match (saves disk space).
* **timeline recovery list**: View recovery timestamps, sheet count, and occupied cells.
* **Manual Snapshots**: Capture custom-named saves (e.g. *"Before CSV parsing"*).
* **FIFO limits**: Rotating queues keep 10 automated and 5 manual saves.
* **Durable restores**: confirmation prompts hot-swap states and write to history so restores are undoable.

---

## ⌨️ Keystroke Commands Guide

| Keystroke | Scope | Visual Action |
| :--- | :--- | :--- |
| `Tab` / `Shift+Tab` | Grid Selection | Focuses cell Right / Left |
| `Enter` / `Shift+Enter` | Grid Selection | Focuses cell Down / Up |
| `← ↑ ↓ →` (Arrows) | Grid Selection | Navigates selection anchor |
| `Shift + Arrows` | Grid Selection | Expands selection range outline |
| `Ctrl + A` | Grid Selection | Selects entire worksheet grid |
| `Ctrl + C` | Clipboard | Copies selected range ( Marching Ants ) |
| `Ctrl + V` | Clipboard | Pastes formulas (offset relative) / Styles |
| `F2` | Cell Editor | Starts inline cell editing |
| `Esc` | Cell Editor / Selection | Cancels edits / Clears copied ants borders |
| `Ctrl + Z` / `Ctrl + Y` | History State | Triggers Undo / Redo action |
| `?` (Shift + `/`) | Shortcuts Modal | Toggles Shortcuts Help Overlay |
| `Ctrl + F` | Search Panel | Toggles Find & Replace Sidebar |
| `Ctrl + S` | Backups Panel | Focuses Manual Snapshot save form |

---

## 🏃 Running the Application

Since **Cells** is 100% client-side and server-free, you can run the app offline on your machine with zero configuration:

1. **Direct Double-Click**:
   Open **index.html** in any modern web browser (Google Chrome, Firefox, Microsoft Edge, Safari) directly from your filesystem.
2. **Local Dev Server**:
   Alternatively, serve it locally using a simple HTTP server:
   ```bash
   npx serve .
   ```
   Or double-click and open the local file directly to enjoy lightning-fast, high-end private spreadsheets!
