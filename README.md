# Keyboard Jockey

A keyboard-driven mouse navigation and window management utility for Windows. Keyboard Jockey lives in the system tray and provides a full-screen overlay grid that lets you move the mouse, click, switch windows, and scroll — all without touching the mouse.

## Getting Started

Launch `KeyboardJockey.exe`. It minimizes to the system tray immediately. Press **Ctrl+Alt+M** to activate the grid overlay. Right-click the tray icon for options or to exit.

## Features

### Grid Navigation

<img width="1913" height="1129" alt="image" src="https://github.com/user-attachments/assets/beb9e7cd-05cd-476d-a7d6-33c78f2cdc2e" />

Press **Ctrl+Alt+M** to show a full-screen overlay grid. Each cell is labeled with a short letter code. Type the letters to move the mouse to that cell, then press **Enter** to click. Hold **Ctrl+Enter** for a right-click.

If you need further accuracy, each cell also contains a 3×3 sub-grid labeled **a–h** (around the center). After typing a cell code, type one more letter to move the mouse to a specific sub-position within that cell.

<img width="1940" height="1136" alt="image" src="https://github.com/user-attachments/assets/542c9972-1d19-43c4-a9b9-d696486bf2f9" />

If the content is too obscured by the grid, hold down **Shift** to make your desktop more visible.

### Arrow Key Fine-Tuning

Once the grid is visible, use the **arrow keys** to nudge the mouse:

- **Arrow keys**: Move 10 pixels
- **Shift+Arrow**: Move 1 pixel (precision mode)
- **Ctrl+Arrow**: Move 50 pixels (fast mode)

The grid fades to semi-transparent during arrow key movement so you can see what's underneath.

### Window Switching (TAB Mode)

Press **Tab** while the grid is showing to cycle through open application windows. Each window is highlighted with a red border and a title label showing its position in the list. Unlike traditional Windows Alt+Tab behaviour (which cycles through by order of last use), windows are sorted by visible area (most visible first).

<img width="1913" height="1180" alt="image" src="https://github.com/user-attachments/assets/8a674c16-8a85-4009-b95b-31b6ff27e8ef" />

- **Tab**: Next window
- **Shift+Tab**: Previous window
- **Enter**: Activate the highlighted window
- **Escape**: Cancel and close the overlay

### Window Search (Type-to-Select)

After pressing Tab, you can search for windows by typing part of their title:

- After a brief pause in TAB mode, or by pressing **\***, all windows (including fully occluded ones) are shown with red highlight boxes
- Start **typing** to filter windows by substring match (case-insensitive). For example, typing "out" would match "Outlook", "About", etc.
- **Enter** activates the highlighted window. If your search narrows to a single match, Enter focuses it immediately
- **Backspace** removes the last character from the search
- The search resets after a brief pause, returning to the all-windows view
- **Tab/Shift+Tab** still cycles through the filtered results

### Cursor Hide & Reveal

Keyboard Jockey automatically hides your mouse cursor once you start typing, so it doesn't obscure what you're trying to type. Moving the mouse at any time will bring the cursor back. This is similar to the built-in Windows "Hide pointer while typing" setting, except it works more consistently — for example, in web browsers and Electron-based applications.

- On reappear, the cursor plays a **shrink animation** (large → normal size) so you can easily spot where it is

### Scroll Pass-Through (PgUp/PgDn)

Press **Page Up** or **Page Down** while the grid is showing to scroll the content under the mouse cursor without dismissing the overlay:

- The grid becomes fully transparent, and scroll wheel events are sent to the window under the cursor
- You can press PgUp/PgDn repeatedly to keep scrolling
- **Moving the mouse** or **pressing any other key** exits scroll mode and closes the overlay

### Per-Monitor DPI Awareness

Keyboard Jockey is per-monitor DPI aware (V2). On multi-monitor setups with different scaling factors, the overlay, grid coordinates, and window highlight rects all use physical pixel coordinates for accurate alignment. Font sizes and line widths scale proportionally to screen dimensions.

## Building

Requires Visual Studio 2022 with the C++ desktop development workload.

```powershell
.\build.ps1
```

The output is `x64\Release\KeyboardJockey.exe`.

## Keyboard Reference

| Key | Context | Action |
|-----|---------|--------|
| **Ctrl+Alt+M** | Global | Toggle grid overlay |
| **a–z** | Grid mode | Type cell label to move mouse |
| **Enter** | Grid mode | Left-click |
| **Ctrl+Enter** | Grid mode | Right-click |
| **Arrow keys** | Grid mode | Nudge mouse (10px) |
| **Shift+Arrow** | Grid mode | Nudge mouse (1px) |
| **Ctrl+Arrow** | Grid mode | Nudge mouse (50px) |
| **Space** | Grid mode | Hide cursor and close grid |
| **Tab** | Grid mode | Enter window cycling mode |
| **Shift+Tab** | TAB mode | Previous window |
| **Enter** | TAB mode | Activate highlighted window |
| **a–z** | TAB mode | Filter windows by title substring |
| **\*** | TAB mode | Show all windows (including hidden) |
| **Backspace** | TAB mode | Delete last search character |
| **PgUp/PgDn** | Grid mode | Scroll content under cursor |
| **Escape** | Any mode | Close overlay |

