// KeyboardJockey.cpp - A keyboard-driven mouse navigation utility
// Minimizes to tray, shows grid overlay on Ctrl+Alt+M, type letters to move mouse

#define WIN32_LEAN_AND_MEAN
#define OEMRESOURCE  // Required for OCR_* cursor constants

#include <windows.h>
#include <shellapi.h>
#include <ShellScalingApi.h>
#pragma comment(lib, "Shcore.lib")
#include <vector>
#include <string>
#include <map>
#include <cstdlib>
#include <cmath>
#include <algorithm>
#include <thread>

// Resource IDs
#define IDI_KEYBOARDJOCKEY 101
#define IDM_EXIT 1001
#define IDM_SHOW 1002
#define IDM_PALETTE 1003

// Constants
#define WM_TRAYICON (WM_USER + 1)
#define HOTKEY_ID_SHOW_GRID 1
#define TARGET_CELL_SIZE_DIP 86  // Target cell size in device-independent pixels (at 96 DPI)
#define TIMER_ID_RESET 1
#define TIMER_ID_TAB_TEXT 2
#define RESET_TIMEOUT_MS 3000
#define TAB_TEXT_TIMEOUT_MS 4000
#define GRID_ALPHA 160           // Default grid overlay opacity (0-255)
#define MOUSE_MOVE_ALPHA 0       // Overlay fully invisible during arrow-key mouse movement
#define SHIFT_PEEK_ALPHA 51      // 80% transparent peek when Shift held in typing mode
#define ACTIVATION_DELAY_MS 50   // Brief sleep before activating a window
#define DEFAULT_DPI 96           // Standard Windows DPI baseline
#define MAIN_FONT_HEIGHT_PCT 80  // Main label font height as % of sub-cell height
#define MAIN_FONT_WIDTH_DIV 5    // Main label font width = cellW / this
#define MIN_MAIN_FONT_SIZE (-8)  // Floor for main label font
#define SUB_FONT_HEIGHT_PCT 60   // Sub-label font height as % of sub-cell height
#define MIN_SUB_FONT_SIZE (-6)   // Floor for sub-label font
#define DT_CENTERED (DT_CENTER | DT_VCENTER | DT_SINGLELINE)

static const wchar_t* GRID_FONT_NAME = L"Segoe UI Variable Display";
static const wchar_t* SUB_LABELS = L"abcdefgh";
static const DWORD CURSOR_IDS[] = { OCR_NORMAL, OCR_IBEAM, OCR_HAND, OCR_CROSS,
                                    OCR_SIZEALL, OCR_SIZENWSE, OCR_SIZENESW, OCR_SIZEWE, OCR_SIZENS };

// Centralized color palette – all colors generated from a single base hue
struct Palette {
    // Base grid
    COLORREF background;              // Overall background fill
    COLORREF cellBgEven;              // Checkerboard cell fill (even)
    COLORREF cellBgOdd;               // Checkerboard cell fill (odd)
    COLORREF gridLine;                // Major cell border lines
    COLORREF subGridLine;             // 3×3 sub-grid lines inside each cell
    COLORREF mainLabelText;           // Main 3-letter label text
    COLORREF subLabelText;            // Sub-label text (a–h)

    // Typing – fully matched cell
    COLORREF matchCellBg;             // Background of the matched cell
    COLORREF matchGridLine;           // Sub-grid lines on the matched cell
    COLORREF matchLabelText;          // Main label on the matched cell
    COLORREF matchSubLabelText;       // Sub-labels on the matched cell
    COLORREF matchSubHighlightBg;     // Highlighted sub-cell background
    COLORREF matchSubHighlightText;   // Highlighted sub-cell text

    // Typing – partial match
    COLORREF partialMatchBg;          // Partially matched cell background
    COLORREF partialMatchText;        // Text on partially matched cell

    // Typing – non-match (dimmed)
    COLORREF dimBg;                   // Dimmed non-matching cell background
    COLORREF dimText;                 // Dimmed non-matching cell text

    // Window highlight (TAB mode) reuses: mainLabelText, gridLine,
    //   matchCellBg, cellBgEven, matchLabelText
    // Minimized panel reuses: background, gridLine, mainLabelText,
    //   subLabelText, matchSubHighlightBg, matchSubHighlightText
};

// --- HSL → RGB conversion for generative palette ---
static COLORREF hsl(float h, float s, float l) {
    h = fmodf(h, 360.0f);
    if (h < 0.0f) h += 360.0f;
    float c = (1.0f - fabsf(2.0f * l - 1.0f)) * s;
    float x = c * (1.0f - fabsf(fmodf(h / 60.0f, 2.0f) - 1.0f));
    float m = l - c / 2.0f;
    float r, g, b;
    if      (h < 60.0f)  { r = c; g = x; b = 0; }
    else if (h < 120.0f) { r = x; g = c; b = 0; }
    else if (h < 180.0f) { r = 0; g = c; b = x; }
    else if (h < 240.0f) { r = 0; g = x; b = c; }
    else if (h < 300.0f) { r = x; g = 0; b = c; }
    else                  { r = c; g = 0; b = x; }
    return RGB(
        max(0, min(255, (int)((r + m) * 255.0f + 0.5f))),
        max(0, min(255, (int)((g + m) * 255.0f + 0.5f))),
        max(0, min(255, (int)((b + m) * 255.0f + 0.5f))));
}

// Color wheel:  0°=Red  30°=Orange  60°=Yellow  120°=Green
//               180°=Cyan  210°=Azure  240°=Blue  270°=Purple  300°=Magenta
//
// ► Change BASE_HUE to shift the entire UI to your favorite color.
#define BASE_HUE_DEFAULT 30.0f  // 30 = woodsy amber (default)
                                  // Try: 270 = purple, 210 = ocean blue, 0 = crimson,
                                  //      160 = teal, 340 = rose, 60 = golden

static float g_baseHue = BASE_HUE_DEFAULT;  // Current hue – changed at runtime by palette picker

static Palette GeneratePalette(float H) {
    float A = H + 90.0f;  // accent hue – 90° offset for natural contrast

    Palette p;
    //                               Hue          Sat    Light
    // -- base grid ------------------------------------------------
    p.background            = hsl(H,         0.40f, 0.04f);  // very dark base
    p.cellBgEven            = hsl(H,         0.40f, 0.12f);  // dark base tint
    p.cellBgOdd             = hsl(A,         0.35f, 0.12f);  // dark accent (checker)
    p.gridLine              = hsl(H,         0.25f, 0.32f);  // medium base
    p.subGridLine           = hsl(H + 45,    0.20f, 0.25f);  // muted mid-tone
    p.mainLabelText         = hsl(H + 10,    0.65f, 0.65f);  // bright warm label
    p.subLabelText          = hsl(A - 20,    0.30f, 0.58f);  // medium accent

    // -- typing: full match ---------------------------------------
    p.matchCellBg           = hsl(A,         0.45f, 0.20f);  // rich accent bg
    p.matchGridLine         = hsl(A,         0.45f, 0.33f);  // bright accent lines
    p.matchLabelText        = hsl(H,         0.20f, 0.90f);  // near-white base tint
    p.matchSubLabelText     = hsl(A,         0.35f, 0.72f);  // light accent
    p.matchSubHighlightBg   = hsl(A,         0.55f, 0.33f);  // vivid accent
    p.matchSubHighlightText = hsl(H,         0.10f, 0.95f);  // near-white

    // -- typing: partial match ------------------------------------
    p.partialMatchBg        = hsl(A,         0.35f, 0.12f);  // subtle accent
    p.partialMatchText      = hsl(A,         0.45f, 0.75f);  // bright accent

    // -- typing: non-match (dimmed) -------------------------------
    p.dimBg                 = hsl(H,         0.30f, 0.04f);  // fade to background
    p.dimText               = hsl(H,         0.20f, 0.25f);  // muted base

    return p;
}

static Palette g_palette = GeneratePalette(BASE_HUE_DEFAULT);

// Global variables
HINSTANCE g_hInstance;
HWND g_hMainWnd = NULL;
HWND g_hOverlayWnd = NULL;
NOTIFYICONDATA g_nid = {};
bool g_bGridVisible = false;
bool g_bMouseMoveMode = false;  // True when using arrow keys
bool g_bCursorHidden = false;   // True when cursor is hidden via Space
HHOOK g_hMouseHook = NULL;      // Low-level mouse hook
HHOOK g_hKeyboardHook = NULL;   // Low-level keyboard hook for hiding cursor on typing
bool g_bCursorAnimating = false; // True during cursor shrink animation
HCURSOR g_hSavedArrow = NULL;    // Saved copy of default arrow cursor for animation
bool g_bScrollMode = false;       // True when in PgUp/PgDn scroll-through mode
HHOOK g_hScrollMouseHook = NULL;  // Mouse hook for detecting movement in scroll mode
bool g_bTabTextMode = false;      // True when in TAB "select by text" mode (all windows shown)
std::wstring g_typedChars;
std::map<std::wstring, POINT> g_gridMap;

// Cached base grid bitmap (rendered once at startup)
HBITMAP g_hGridBitmap = NULL;
int g_gridBitmapW = 0;
int g_gridBitmapH = 0;

// Virtual screen bounds (all monitors combined)
struct VirtualScreenBounds {
    int left, top, width, height;
};

inline VirtualScreenBounds GetVirtualScreenBounds() {
    return {
        GetSystemMetrics(SM_XVIRTUALSCREEN),
        GetSystemMetrics(SM_YVIRTUALSCREEN),
        GetSystemMetrics(SM_CXVIRTUALSCREEN),
        GetSystemMetrics(SM_CYVIRTUALSCREEN)
    };
}

// Per-monitor info
struct MonitorInfo {
    HMONITOR hMonitor;
    RECT rcMonitor;
    UINT dpiX, dpiY;
    wchar_t prefix;  // First letter of labels on this monitor ('a', 'b', ...)
};
std::vector<MonitorInfo> g_monitors;

// Grid cell structure
struct GridCell {
    RECT rect;
    std::wstring label;  // 3-letter label: monitor prefix + 2-char cell code
    POINT center;
    POINT subPoints[9];  // 3x3 sub-grid points (0-8, center is 4)
    int gridRow, gridCol; // Position in the monitor grid (for checkerboard)
};

std::vector<GridCell> g_cells;

// Window highlight (TAB cycling) state
struct AppWindow {
    HWND hwnd;
    RECT rect;
    std::wstring title;
    int visibleArea;  // Pixels of visible (unoccluded) area
};
std::vector<AppWindow> g_appWindows;
std::vector<AppWindow> g_allAppWindows;  // All enumerated windows (including 0 visible area)
std::vector<AppWindow> g_minimizedWindows;  // Minimized windows
std::vector<AppWindow> g_allMinimizedWindows;  // All minimized (before search filter)
int g_highlightIndex = -1;  // -1 = no highlight active
std::wstring g_tabSearchStr;  // Substring search in TAB mode

HWND g_hPaletteWnd = NULL;  // Palette picker window

// Forward declarations
LRESULT CALLBACK MainWndProc(HWND, UINT, WPARAM, LPARAM);
LRESULT CALLBACK OverlayWndProc(HWND, UINT, WPARAM, LPARAM);
LRESULT CALLBACK PaletteWndProc(HWND, UINT, WPARAM, LPARAM);
void CreateTrayIcon(HWND hWnd);
void RemoveTrayIcon();
void ShowContextMenu(HWND hWnd);
void ShowPaletteWindow();
void ApplyHue(float hue);
void CreateOverlayWindow();
void ShowGrid();
void HideGrid();
void BuildGridCells();
void RenderBaseGridBitmap();
void PaintGrid(HDC hdc);
void ProcessTypedChar(wchar_t ch);
void MoveMouse(POINT pt);
void SendClick(bool rightClick);
std::wstring GenerateLabel(wchar_t monitorPrefix, int index);
void FilterAppWindowsBySearch();
void HideCursor();
void RestoreCursor();
void MoveMouseByArrowKey(int dx, int dy, int moveAmount);
void CreateGridFonts(int sh, int cellW, HFONT* outMain, HFONT* outSub);
LRESULT CALLBACK LowLevelMouseProc(int nCode, WPARAM wParam, LPARAM lParam);
LRESULT CALLBACK LowLevelKeyboardProc(int nCode, WPARAM wParam, LPARAM lParam);
void InstallGlobalKeyboardHook();
void UninstallGlobalKeyboardHook();
void EnumerateAppWindows();
void CycleHighlight(bool forward);
void ExitScrollMode();
LRESULT CALLBACK ScrollMouseProc(int nCode, WPARAM wParam, LPARAM lParam);

// Force restore cursors - called on exit/crash
void ForceRestoreCursors() {
    g_bCursorAnimating = false;
    SystemParametersInfo(SPI_SETCURSORS, 0, NULL, 0);
}

// Unhandled exception filter - restore cursor on crash
LONG WINAPI CrashHandler(EXCEPTION_POINTERS* pExceptionInfo) {
    ForceRestoreCursors();
    return EXCEPTION_CONTINUE_SEARCH;
}

// Low-level mouse hook for scroll mode - detect any mouse movement
LRESULT CALLBACK ScrollMouseProc(int nCode, WPARAM wParam, LPARAM lParam) {
    if (nCode >= 0 && g_bScrollMode && wParam == WM_MOUSEMOVE) {
        ExitScrollMode();
    }
    return CallNextHookEx(g_hScrollMouseHook, nCode, wParam, lParam);
}

void ExitScrollMode() {
    if (!g_bScrollMode) return;
    g_bScrollMode = false;
    // Remove input-transparent flag so overlay receives input again
    LONG_PTR exStyle = GetWindowLongPtr(g_hOverlayWnd, GWL_EXSTYLE);
    SetWindowLongPtr(g_hOverlayWnd, GWL_EXSTYLE, exStyle & ~WS_EX_TRANSPARENT);
    if (g_hScrollMouseHook) {
        UnhookWindowsHookEx(g_hScrollMouseHook);
        g_hScrollMouseHook = NULL;
    }
    HideGrid();
}

// Low-level mouse hook callback
LRESULT CALLBACK LowLevelMouseProc(int nCode, WPARAM wParam, LPARAM lParam) {
    if (nCode >= 0 && g_bCursorHidden) {
        // Mouse moved - restore cursor
        if (wParam == WM_MOUSEMOVE) {
            RestoreCursor();
        }
    }
    return CallNextHookEx(g_hMouseHook, nCode, wParam, lParam);
}

// Hide the cursor system-wide
void HideCursor() {
    if (g_bCursorHidden) return;
    
    // Cancel any ongoing animation and wait for it to see the flag
    if (g_bCursorAnimating) {
        g_bCursorAnimating = false;
        Sleep(100);  // Give animation thread time to exit
    }
    
    // Create a blank cursor
    BYTE andMask[128];
    BYTE xorMask[128];
    memset(andMask, 0xFF, sizeof(andMask));  // AND mask all 1s = transparent
    memset(xorMask, 0x00, sizeof(xorMask));  // XOR mask all 0s = no change
    HCURSOR hBlankCursor = CreateCursor(g_hInstance, 0, 0, 32, 32, andMask, xorMask);
    
    // Copy and set for all standard cursor types
    for (DWORD id : CURSOR_IDS) {
        HCURSOR hCopy = CopyCursor(hBlankCursor);
        SetSystemCursor(hCopy, id);
    }
    
    DestroyCursor(hBlankCursor);
    
    // Install low-level mouse hook to detect movement
    g_hMouseHook = SetWindowsHookEx(WH_MOUSE_LL, LowLevelMouseProc, g_hInstance, 0);
    
    g_bCursorHidden = true;
}

// Set all system cursors to a scaled version of the saved arrow cursor
// Scale the saved cursor to a new size using DrawIconEx for proper transparency
HCURSOR CreateScaledCursor(HCURSOR hOriginal, int targetSize) {
    if (!hOriginal) return NULL;
    
    // Get original cursor info for hotspot
    ICONINFO iiOrig;
    if (!GetIconInfo(hOriginal, &iiOrig)) return NULL;
    
    BITMAP bm;
    GetObject(iiOrig.hbmMask, sizeof(bm), &bm);
    int origW = bm.bmWidth;
    int origH = iiOrig.hbmColor ? bm.bmHeight : bm.bmHeight / 2;
    
    // Calculate scaled hotspot
    int hotX = (int)((float)iiOrig.xHotspot / origW * targetSize);
    int hotY = (int)((float)iiOrig.yHotspot / origH * targetSize);
    
    // Clean up ICONINFO bitmaps
    DeleteObject(iiOrig.hbmColor);
    DeleteObject(iiOrig.hbmMask);
    
    HDC hdcScreen = GetDC(NULL);
    
    // Create a 32-bit ARGB bitmap for the color (supports per-pixel alpha)
    BITMAPINFO bmi = {};
    bmi.bmiHeader.biSize = sizeof(BITMAPINFOHEADER);
    bmi.bmiHeader.biWidth = targetSize;
    bmi.bmiHeader.biHeight = -targetSize;  // Top-down
    bmi.bmiHeader.biPlanes = 1;
    bmi.bmiHeader.biBitCount = 32;
    bmi.bmiHeader.biCompression = BI_RGB;
    
    void* pBits = NULL;
    HBITMAP hbmColor = CreateDIBSection(hdcScreen, &bmi, DIB_RGB_COLORS, &pBits, NULL, 0);
    
    // Create mask bitmap
    HBITMAP hbmMask = CreateBitmap(targetSize, targetSize, 1, 1, NULL);
    
    HDC hdcColor = CreateCompatibleDC(hdcScreen);
    HDC hdcMask = CreateCompatibleDC(hdcScreen);
    
    // Draw color: clear to transparent black, then draw icon
    SelectObject(hdcColor, hbmColor);
    memset(pBits, 0, targetSize * targetSize * 4);  // Fully transparent
    DrawIconEx(hdcColor, 0, 0, hOriginal, targetSize, targetSize, 0, NULL, DI_NORMAL);
    
    // Draw mask: white = transparent, black = opaque
    SelectObject(hdcMask, hbmMask);
    RECT rcMask = { 0, 0, targetSize, targetSize };
    FillRect(hdcMask, &rcMask, (HBRUSH)GetStockObject(WHITE_BRUSH));
    DrawIconEx(hdcMask, 0, 0, hOriginal, targetSize, targetSize, 0, NULL, DI_MASK);
    
    DeleteDC(hdcColor);
    DeleteDC(hdcMask);
    ReleaseDC(NULL, hdcScreen);
    
    // Create the scaled cursor
    ICONINFO iiNew;
    iiNew.fIcon = FALSE;
    iiNew.xHotspot = hotX;
    iiNew.yHotspot = hotY;
    iiNew.hbmMask = hbmMask;
    iiNew.hbmColor = hbmColor;
    
    HCURSOR hResult = CreateIconIndirect(&iiNew);
    DeleteObject(hbmColor);
    DeleteObject(hbmMask);
    
    return hResult;
}

void SetScaledCursors(int size) {
    if (!g_hSavedArrow) return;
    
    HCURSOR hScaled = CreateScaledCursor(g_hSavedArrow, size);
    if (!hScaled) return;
    
    for (DWORD id : CURSOR_IDS) {
        HCURSOR hCopy = CopyCursor(hScaled);
        SetSystemCursor(hCopy, id);
    }
    DestroyCursor(hScaled);
}

// Animate cursor from large to normal size over ~1 second
void AnimateCursorRestore() {
    g_bCursorAnimating = true;
    
    const int startSize = 128;   // Start big
    const int endSize = 32;      // Normal cursor size
    const int steps = 15;
    const int delayMs = 500 / steps;  // ~33ms per step
    
    // Set the first frame immediately (large cursor)
    SetScaledCursors(startSize);
    
    for (int i = 1; i <= steps; i++) {
        if (!g_bCursorAnimating) break;  // Cancelled
        
        Sleep(delayMs);
        
        if (!g_bCursorAnimating) break;  // Check again after sleep
        
        // Ease-out: fast shrink at start, slow at end
        float t = (float)i / (float)steps;
        float eased = 1.0f - (1.0f - t) * (1.0f - t);  // quadratic ease-out
        int size = startSize - (int)((startSize - endSize) * eased);
        if (size < endSize) size = endSize;
        
        SetScaledCursors(size);
    }
    
    // Final step: restore to true system defaults (only if not cancelled)
    if (g_bCursorAnimating) {
        SystemParametersInfo(SPI_SETCURSORS, 0, NULL, 0);
    }
    g_bCursorAnimating = false;
}

// Restore the cursor to system defaults
void RestoreCursor() {
    if (!g_bCursorHidden) return;
    
    // Unhook mouse first
    if (g_hMouseHook) {
        UnhookWindowsHookEx(g_hMouseHook);
        g_hMouseHook = NULL;
    }
    
    g_bCursorHidden = false;
    
    // Start animated cursor restore in a background thread
    std::thread(AnimateCursorRestore).detach();
}

// Check if a key is a typing key (not a modifier, function key, etc.)
bool IsTypingKey(DWORD vkCode) {
    // Ignore modifiers
    if (vkCode == VK_SHIFT || vkCode == VK_CONTROL || vkCode == VK_MENU ||
        vkCode == VK_LSHIFT || vkCode == VK_RSHIFT ||
        vkCode == VK_LCONTROL || vkCode == VK_RCONTROL ||
        vkCode == VK_LMENU || vkCode == VK_RMENU ||
        vkCode == VK_LWIN || vkCode == VK_RWIN) {
        return false;
    }
    
    // Ignore function keys
    if (vkCode >= VK_F1 && vkCode <= VK_F24) {
        return false;
    }
    
    // Ignore navigation keys that typically involve mouse use
    if (vkCode == VK_PRINT || vkCode == VK_SNAPSHOT ||
        vkCode == VK_PAUSE || vkCode == VK_CAPITAL ||
        vkCode == VK_NUMLOCK || vkCode == VK_SCROLL) {
        return false;
    }
    
    // Consider everything else as typing
    return true;
}

// Low-level keyboard hook callback - hide cursor on typing
LRESULT CALLBACK LowLevelKeyboardProc(int nCode, WPARAM wParam, LPARAM lParam) {
    if (nCode >= 0 && !g_bGridVisible && !g_bMouseMoveMode) {
        // Only on key down events, and not when our grid is active
        if (wParam == WM_KEYDOWN || wParam == WM_SYSKEYDOWN) {
            KBDLLHOOKSTRUCT* pKb = (KBDLLHOOKSTRUCT*)lParam;
            if (IsTypingKey(pKb->vkCode)) {
                // Hide cursor when typing
                if (!g_bCursorHidden) {
                    HideCursor();
                }
            }
        }
    }
    return CallNextHookEx(g_hKeyboardHook, nCode, wParam, lParam);
}

// Install global keyboard hook for cursor hiding while typing
void InstallGlobalKeyboardHook() {
    if (g_hKeyboardHook) return;  // Already installed
    g_hKeyboardHook = SetWindowsHookEx(WH_KEYBOARD_LL, LowLevelKeyboardProc, g_hInstance, 0);
}

// Uninstall global keyboard hook
void UninstallGlobalKeyboardHook() {
    if (g_hKeyboardHook) {
        UnhookWindowsHookEx(g_hKeyboardHook);
        g_hKeyboardHook = NULL;
    }
}

// Monitor enumeration callback for DPI-aware per-monitor grids
BOOL CALLBACK GridMonitorEnumProc(HMONITOR hMon, HDC hdcMon, LPRECT lprcMonitor, LPARAM lParam) {
    MonitorInfo mi;
    mi.hMonitor = hMon;
    mi.rcMonitor = *lprcMonitor;
    HRESULT hr = GetDpiForMonitor(hMon, MDT_EFFECTIVE_DPI, &mi.dpiX, &mi.dpiY);
    if (FAILED(hr)) { mi.dpiX = DEFAULT_DPI; mi.dpiY = DEFAULT_DPI; }
    auto* vec = reinterpret_cast<std::vector<MonitorInfo>*>(lParam);
    mi.prefix = L'a' + (wchar_t)vec->size();
    vec->push_back(mi);
    return TRUE;
}

// Generate 3-letter label: monitor prefix + 2-char cell code
std::wstring GenerateLabel(wchar_t monitorPrefix, int index) {
    std::wstring label;
    label += monitorPrefix;
    int first = index / 26;
    int second = index % 26;
    if (first < 26) {
        label += static_cast<wchar_t>(L'a' + first);
        label += static_cast<wchar_t>(L'a' + second);
    }
    return label;
}

// Build grid cells per monitor with DPI-aware sizing
void BuildGridCells() {
    g_cells.clear();
    g_gridMap.clear();
    g_monitors.clear();
    
    // Enumerate all monitors with DPI info
    EnumDisplayMonitors(NULL, NULL, GridMonitorEnumProc, reinterpret_cast<LPARAM>(&g_monitors));
    
    // Build grid independently for each monitor
    for (const auto& mon : g_monitors) {
        int monWidth = mon.rcMonitor.right - mon.rcMonitor.left;
        int monHeight = mon.rcMonitor.bottom - mon.rcMonitor.top;
        
        // Scale target cell size by this monitor's DPI
        int targetCellPx = TARGET_CELL_SIZE_DIP * (int)mon.dpiX / DEFAULT_DPI;
        
        int gridCols = max(1, monWidth / targetCellPx);
        int gridRows = max(1, monHeight / targetCellPx);
        // Cap to 676 cells per monitor (2-char codes aa-zz)
        while (gridCols * gridRows > 676) {
            if (gridCols > gridRows) gridCols--; else gridRows--;
        }
        
        int cellWidth = monWidth / gridCols;
        int cellHeight = monHeight / gridRows;
        int subWidth = cellWidth / 3;
        int subHeight = cellHeight / 3;
        
        int index = 0;
        for (int row = 0; row < gridRows; row++) {
            for (int col = 0; col < gridCols; col++) {
                GridCell cell;
                cell.rect.left = mon.rcMonitor.left + col * cellWidth;
                cell.rect.top = mon.rcMonitor.top + row * cellHeight;
                cell.rect.right = cell.rect.left + cellWidth;
                cell.rect.bottom = cell.rect.top + cellHeight;
                cell.center.x = cell.rect.left + cellWidth / 2;
                cell.center.y = cell.rect.top + cellHeight / 2;
                cell.label = GenerateLabel(mon.prefix, index);
                cell.gridRow = row;
                cell.gridCol = col;
                
                // Calculate 3x3 sub-grid points
                for (int sy = 0; sy < 3; sy++) {
                    for (int sx = 0; sx < 3; sx++) {
                        int subIdx = sy * 3 + sx;
                        cell.subPoints[subIdx].x = cell.rect.left + sx * subWidth + subWidth / 2;
                        cell.subPoints[subIdx].y = cell.rect.top + sy * subHeight + subHeight / 2;
                    }
                }
                
                g_cells.push_back(cell);
                g_gridMap[cell.label] = cell.center;
                index++;
            }
        }
    }
}

// Render the static base grid (lines, labels, sub-labels) to a cached bitmap
void RenderBaseGridBitmap() {
    if (g_hGridBitmap) { DeleteObject(g_hGridBitmap); g_hGridBitmap = NULL; }
    
    auto vs = GetVirtualScreenBounds();
    int virtualWidth = vs.width;
    int virtualHeight = vs.height;
    int virtualLeft = vs.left;
    int virtualTop = vs.top;
    g_gridBitmapW = virtualWidth;
    g_gridBitmapH = virtualHeight;
    
    HDC hdcScreen = GetDC(NULL);
    HDC hdcMem = CreateCompatibleDC(hdcScreen);
    g_hGridBitmap = CreateCompatibleBitmap(hdcScreen, virtualWidth, virtualHeight);
    SelectObject(hdcMem, g_hGridBitmap);
    
    // Background fill
    HBRUSH hBrushBg = CreateSolidBrush(g_palette.background);
    RECT rcFull = { 0, 0, virtualWidth, virtualHeight };
    FillRect(hdcMem, &rcFull, hBrushBg);
    DeleteObject(hBrushBg);
    
    // Checkerboard cell backgrounds
    HBRUSH hBrushEven = CreateSolidBrush(g_palette.cellBgEven);
    HBRUSH hBrushOdd  = CreateSolidBrush(g_palette.cellBgOdd);
    for (const auto& cell : g_cells) {
        RECT adj;
        adj.left   = cell.rect.left   - virtualLeft;
        adj.top    = cell.rect.top    - virtualTop;
        adj.right  = cell.rect.right  - virtualLeft;
        adj.bottom = cell.rect.bottom - virtualTop;
        bool isEven = ((cell.gridRow + cell.gridCol) % 2 == 0);
        FillRect(hdcMem, &adj, isEven ? hBrushEven : hBrushOdd);
    }
    DeleteObject(hBrushEven);
    DeleteObject(hBrushOdd);
    
    // Grid lines
    int gridPenWidth = max(1, virtualHeight / 800);
    HPEN hPen = CreatePen(PS_SOLID, gridPenWidth, g_palette.gridLine);
    HPEN hSubPen = CreatePen(PS_SOLID, max(1, gridPenWidth / 2), g_palette.subGridLine);
    HPEN hOldPen = (HPEN)SelectObject(hdcMem, hPen);
    
    for (const auto& cell : g_cells) {
        int sw = (cell.rect.right - cell.rect.left) / 3;
        int sh = (cell.rect.bottom - cell.rect.top) / 3;
        
        RECT adj;
        adj.left = cell.rect.left - virtualLeft;
        adj.top = cell.rect.top - virtualTop;
        adj.right = cell.rect.right - virtualLeft;
        adj.bottom = cell.rect.bottom - virtualTop;
        
        SelectObject(hdcMem, hPen);
        MoveToEx(hdcMem, adj.left, adj.top, NULL);
        LineTo(hdcMem, adj.right, adj.top);
        LineTo(hdcMem, adj.right, adj.bottom);
        LineTo(hdcMem, adj.left, adj.bottom);
        LineTo(hdcMem, adj.left, adj.top);
        
        SelectObject(hdcMem, hSubPen);
        MoveToEx(hdcMem, adj.left + sw, adj.top, NULL);
        LineTo(hdcMem, adj.left + sw, adj.bottom);
        MoveToEx(hdcMem, adj.left + sw * 2, adj.top, NULL);
        LineTo(hdcMem, adj.left + sw * 2, adj.bottom);
        MoveToEx(hdcMem, adj.left, adj.top + sh, NULL);
        LineTo(hdcMem, adj.right, adj.top + sh);
        MoveToEx(hdcMem, adj.left, adj.top + sh * 2, NULL);
        LineTo(hdcMem, adj.right, adj.top + sh * 2);
    }
    
    SelectObject(hdcMem, hOldPen);
    DeleteObject(hPen);
    DeleteObject(hSubPen);
    
    // Draw labels
    SetBkMode(hdcMem, TRANSPARENT);
    SetTextColor(hdcMem, g_palette.mainLabelText);
    
    int lastSubH = 0;
    HFONT hFont = NULL, hSubFont = NULL;
    HFONT hOldFont = NULL;
    
    for (const auto& cell : g_cells) {
        int sw = (cell.rect.right - cell.rect.left) / 3;
        int sh = (cell.rect.bottom - cell.rect.top) / 3;
        
        if (sh != lastSubH) {
            if (hOldFont) { SelectObject(hdcMem, hOldFont); hOldFont = NULL; }
            if (hFont) DeleteObject(hFont);
            if (hSubFont) DeleteObject(hSubFont);
            
            int cellW = sw * 3;
            CreateGridFonts(sh, cellW, &hFont, &hSubFont);
            
            hOldFont = (HFONT)SelectObject(hdcMem, hFont);
            lastSubH = sh;
        }
        
        RECT adj;
        adj.left = cell.rect.left - virtualLeft;
        adj.top = cell.rect.top - virtualTop;
        adj.right = cell.rect.right - virtualLeft;
        adj.bottom = cell.rect.bottom - virtualTop;
        
        // Center label
        SelectObject(hdcMem, hFont);
        SetTextColor(hdcMem, g_palette.mainLabelText);
        DrawText(hdcMem, cell.label.c_str(), -1, &adj, DT_CENTERED);
        
        // Sub-labels
        SelectObject(hdcMem, hSubFont);
        SetTextColor(hdcMem, g_palette.subLabelText);
        int subLabelIdx = 0;
        for (int sy = 0; sy < 3; sy++) {
            for (int sx = 0; sx < 3; sx++) {
                if (sx == 1 && sy == 1) continue;
                RECT subRect;
                subRect.left = adj.left + sx * sw;
                subRect.top = adj.top + sy * sh;
                subRect.right = subRect.left + sw;
                subRect.bottom = subRect.top + sh;
                wchar_t subLabel[2] = { SUB_LABELS[subLabelIdx], 0 };
                DrawText(hdcMem, subLabel, 1, &subRect, DT_CENTERED);
                subLabelIdx++;
            }
        }
    }
    
    if (hOldFont) SelectObject(hdcMem, hOldFont);
    if (hFont) DeleteObject(hFont);
    if (hSubFont) DeleteObject(hSubFont);
    
    DeleteDC(hdcMem);
    ReleaseDC(NULL, hdcScreen);
}

// Paint the grid overlay
void PaintGrid(HDC hdc) {
    // Get virtual screen bounds
    auto vs = GetVirtualScreenBounds();
    int virtualLeft = vs.left;
    int virtualTop = vs.top;
    int virtualWidth = vs.width;
    int virtualHeight = vs.height;
    
    // Semi-transparent background
    HBRUSH hBrushBg = CreateSolidBrush(g_palette.background);
    RECT rcFull = { 0, 0, virtualWidth, virtualHeight };
    FillRect(hdc, &rcFull, hBrushBg);
    DeleteObject(hBrushBg);
    
    // In scroll mode, paint only black (LWA_COLORKEY makes it fully transparent)
    if (g_bScrollMode) return;
    
    // If in window highlight or text select mode, skip drawing the grid entirely
    bool highlightMode = (g_highlightIndex >= 0 || g_bTabTextMode || !g_tabSearchStr.empty());
    
  if (!highlightMode) {
    // Blit cached base grid
    if (g_hGridBitmap) {
        HDC hdcGrid = CreateCompatibleDC(hdc);
        SelectObject(hdcGrid, g_hGridBitmap);
        BitBlt(hdc, 0, 0, g_gridBitmapW, g_gridBitmapH, hdcGrid, 0, 0, SRCCOPY);
        DeleteDC(hdcGrid);
    }
    
    // Overlay dynamic highlights for typed chars
    if (!g_typedChars.empty()) {
        SetBkMode(hdc, TRANSPARENT);
        
        int lastSubH = 0;
        HFONT hFont = NULL, hSubFont = NULL;
        HFONT hOldFont = NULL;
        
        for (const auto& cell : g_cells) {
            int sw = (cell.rect.right - cell.rect.left) / 3;
            int sh = (cell.rect.bottom - cell.rect.top) / 3;
            
            if (sh != lastSubH) {
                if (hOldFont) { SelectObject(hdc, hOldFont); hOldFont = NULL; }
                if (hFont) DeleteObject(hFont);
                if (hSubFont) DeleteObject(hSubFont);
                int cellW = sw * 3;
                CreateGridFonts(sh, cellW, &hFont, &hSubFont);
                hOldFont = (HFONT)SelectObject(hdc, hFont);
                lastSubH = sh;
            }
            
            RECT adjusted;
            adjusted.left = cell.rect.left - virtualLeft;
            adjusted.top = cell.rect.top - virtualTop;
            adjusted.right = cell.rect.right - virtualLeft;
            adjusted.bottom = cell.rect.bottom - virtualTop;
            
            bool isMatch = false;
            bool isPartialMatch = false;
            if (g_typedChars.length() >= 3) {
                if (cell.label == g_typedChars.substr(0, 3)) isMatch = true;
            } else {
                if (cell.label.substr(0, g_typedChars.length()) == g_typedChars) isPartialMatch = true;
            }
            
            if (!isMatch && !isPartialMatch) {
                // Dim non-matching cells
                HBRUSH hDim = CreateSolidBrush(g_palette.dimBg);
                FillRect(hdc, &adjusted, hDim);
                DeleteObject(hDim);
                SelectObject(hdc, hFont);
                SetTextColor(hdc, g_palette.dimText);
                DrawText(hdc, cell.label.c_str(), -1, &adjusted, DT_CENTERED);
                continue;
            }
            
            if (isMatch) {
                HBRUSH hHighlight = CreateSolidBrush(g_palette.matchCellBg);
                FillRect(hdc, &adjusted, hHighlight);
                DeleteObject(hHighlight);
                
                HPEN hSubPenLight = CreatePen(PS_SOLID, 1, g_palette.matchGridLine);
                HPEN hOldPen = (HPEN)SelectObject(hdc, hSubPenLight);
                MoveToEx(hdc, adjusted.left + sw, adjusted.top, NULL);
                LineTo(hdc, adjusted.left + sw, adjusted.bottom);
                MoveToEx(hdc, adjusted.left + sw * 2, adjusted.top, NULL);
                LineTo(hdc, adjusted.left + sw * 2, adjusted.bottom);
                MoveToEx(hdc, adjusted.left, adjusted.top + sh, NULL);
                LineTo(hdc, adjusted.right, adjusted.top + sh);
                MoveToEx(hdc, adjusted.left, adjusted.top + sh * 2, NULL);
                LineTo(hdc, adjusted.right, adjusted.top + sh * 2);
                SelectObject(hdc, hOldPen);
                DeleteObject(hSubPenLight);
                
                SelectObject(hdc, hFont);
                SetTextColor(hdc, g_palette.matchLabelText);
                DrawText(hdc, cell.label.c_str(), -1, &adjusted, DT_CENTERED);
                
                // Draw sub-labels on matched cell
                SelectObject(hdc, hSubFont);
                int subLabelIdx = 0;
                for (int sy = 0; sy < 3; sy++) {
                    for (int sx = 0; sx < 3; sx++) {
                        if (sx == 1 && sy == 1) continue;
                        RECT subRect;
                        subRect.left = adjusted.left + sx * sw;
                        subRect.top = adjusted.top + sy * sh;
                        subRect.right = subRect.left + sw;
                        subRect.bottom = subRect.top + sh;
                        
                        if (g_typedChars.length() == 4) {
                            wchar_t subChar = g_typedChars[3];
                            if (subChar >= L'a' && subChar <= L'h' && (subChar - L'a') == subLabelIdx) {
                                HBRUSH hSubHi = CreateSolidBrush(g_palette.matchSubHighlightBg);
                                FillRect(hdc, &subRect, hSubHi);
                                DeleteObject(hSubHi);
                                SetTextColor(hdc, g_palette.matchSubHighlightText);
                            } else {
                                SetTextColor(hdc, g_palette.matchSubLabelText);
                            }
                        } else {
                            SetTextColor(hdc, g_palette.matchSubLabelText);
                        }
                        wchar_t sl[2] = { SUB_LABELS[subLabelIdx], 0 };
                        DrawText(hdc, sl, 1, &subRect, DT_CENTERED);
                        subLabelIdx++;
                    }
                }
            } else {
                // Partial match - subtle green tint so user can still see underneath
                HBRUSH hPartial = CreateSolidBrush(g_palette.partialMatchBg);
                FillRect(hdc, &adjusted, hPartial);
                DeleteObject(hPartial);
                SelectObject(hdc, hFont);
                SetTextColor(hdc, g_palette.partialMatchText);
                DrawText(hdc, cell.label.c_str(), -1, &adjusted, DT_CENTERED);
            }
        }
        
        if (hOldFont) SelectObject(hdc, hOldFont);
        if (hFont) DeleteObject(hFont);
        if (hSubFont) DeleteObject(hSubFont);
    }
  } // end if (!highlightMode)
    if (g_highlightIndex >= 0 && !g_appWindows.empty()) {
        // When search is active or in text mode, highlight ALL matching windows
        // When just cycling (no search), highlight only the current one
        bool showAll = !g_tabSearchStr.empty() || g_bTabTextMode;
        int startIdx = showAll ? 0 : g_highlightIndex;
        int endIdx = showAll ? (int)g_appWindows.size() : g_highlightIndex + 1;
        
        HFONT hLabelFont = CreateFont(-(max(12, virtualHeight / 80)), 0, 0, 0, FW_BOLD, FALSE, FALSE, FALSE,
            DEFAULT_CHARSET, OUT_TT_PRECIS, CLIP_DEFAULT_PRECIS,
            CLEARTYPE_NATURAL_QUALITY, DEFAULT_PITCH | FF_DONTCARE, GRID_FONT_NAME);
        
        for (int idx = startIdx; idx < endIdx && idx < (int)g_appWindows.size(); idx++) {
            const AppWindow& aw = g_appWindows[idx];
            
            // Convert screen coords to overlay window coords
            RECT hr;
            hr.left = aw.rect.left - virtualLeft;
            hr.top = aw.rect.top - virtualTop;
            hr.right = aw.rect.right - virtualLeft;
            hr.bottom = aw.rect.bottom - virtualTop;
            
        // Draw thick red border (scaled to screen)
            int thickness = max(2, virtualHeight / 400);
            bool isCurrent = (idx == g_highlightIndex);
            HBRUSH hBorderBrush = CreateSolidBrush(isCurrent ? g_palette.mainLabelText : g_palette.gridLine);
            
            // Top edge
            RECT edge = { hr.left, hr.top, hr.right, hr.top + thickness };
            FillRect(hdc, &edge, hBorderBrush);
            // Bottom edge
            edge = { hr.left, hr.bottom - thickness, hr.right, hr.bottom };
            FillRect(hdc, &edge, hBorderBrush);
            // Left edge
            edge = { hr.left, hr.top, hr.left + thickness, hr.bottom };
            FillRect(hdc, &edge, hBorderBrush);
            // Right edge
            edge = { hr.right - thickness, hr.top, hr.right, hr.bottom };
            FillRect(hdc, &edge, hBorderBrush);
            
            DeleteObject(hBorderBrush);
            
            // Draw window title + index label
            HFONT hPrevFont = (HFONT)SelectObject(hdc, hLabelFont);
            
            wchar_t labelBuf[300];
            swprintf_s(labelBuf, L" [%d/%d] %s ", idx + 1, (int)g_appWindows.size(), aw.title.c_str());
            
            SIZE textSize;
            GetTextExtentPoint32(hdc, labelBuf, (int)wcslen(labelBuf), &textSize);
            
            int labelX = hr.left;
            int labelY = hr.top - textSize.cy - 6;
            if (labelY < 0) labelY = hr.top + thickness;
            
            RECT labelBg = { labelX, labelY, labelX + textSize.cx + 8, labelY + textSize.cy + 8 };
            HBRUSH hLabelBg = CreateSolidBrush(isCurrent ? g_palette.matchCellBg : g_palette.cellBgEven);
            FillRect(hdc, &labelBg, hLabelBg);
            DeleteObject(hLabelBg);
            
            SetTextColor(hdc, g_palette.matchLabelText);
            SetBkMode(hdc, TRANSPARENT);
            RECT labelRect = { labelX + 4, labelY + 2, labelBg.right, labelBg.bottom };
            DrawText(hdc, labelBuf, -1, &labelRect, DT_LEFT | DT_SINGLELINE | DT_NOPREFIX);
            
            SelectObject(hdc, hPrevFont);
        }
        
        DeleteObject(hLabelFont);
    }
    
    // Draw minimized windows panel (bottom-right of primary monitor) - only in text search mode
    if ((g_bTabTextMode || !g_tabSearchStr.empty()) && !g_minimizedWindows.empty()) {
        // Get primary monitor work area
        MONITORINFO mi = { sizeof(mi) };
        GetMonitorInfo(MonitorFromPoint({ 0, 0 }, MONITOR_DEFAULTTOPRIMARY), &mi);
        RECT workArea = mi.rcWork;
        
        int panelPad = max(8, virtualHeight / 200);
        int lineH = max(18, virtualHeight / 60);
        int titleH = lineH + panelPad;
        int maxItems = min((int)g_minimizedWindows.size(), 20);  // cap to 20 rows
        int panelH = titleH + maxItems * lineH + panelPad * 2;
        int panelW = max(250, virtualWidth / 5);
        
        // Position: bottom-right of primary monitor work area
        int panelX = (workArea.right - virtualLeft) - panelW - panelPad;
        int panelY = (workArea.bottom - virtualTop) - panelH - panelPad;
        
        RECT panelRect = { panelX, panelY, panelX + panelW, panelY + panelH };
        
        // Panel background
        HBRUSH hPanelBg = CreateSolidBrush(g_palette.cellBgEven);
        FillRect(hdc, &panelRect, hPanelBg);
        DeleteObject(hPanelBg);
        
        // Panel border
        int borderT = max(1, virtualHeight / 500);
        HBRUSH hPanelBorder = CreateSolidBrush(g_palette.gridLine);
        RECT be;
        be = { panelRect.left, panelRect.top, panelRect.right, panelRect.top + borderT };
        FillRect(hdc, &be, hPanelBorder);
        be = { panelRect.left, panelRect.bottom - borderT, panelRect.right, panelRect.bottom };
        FillRect(hdc, &be, hPanelBorder);
        be = { panelRect.left, panelRect.top, panelRect.left + borderT, panelRect.bottom };
        FillRect(hdc, &be, hPanelBorder);
        be = { panelRect.right - borderT, panelRect.top, panelRect.right, panelRect.bottom };
        FillRect(hdc, &be, hPanelBorder);
        DeleteObject(hPanelBorder);
        
        // Title font (bold)
        HFONT hTitleFont = CreateFont(-(lineH * 80 / 100), 0, 0, 0, FW_BOLD, FALSE, FALSE, FALSE,
            DEFAULT_CHARSET, OUT_TT_PRECIS, CLIP_DEFAULT_PRECIS,
            CLEARTYPE_NATURAL_QUALITY, DEFAULT_PITCH | FF_DONTCARE, GRID_FONT_NAME);
        // Item font (normal)
        HFONT hItemFont = CreateFont(-(lineH * 70 / 100), 0, 0, 0, FW_NORMAL, FALSE, FALSE, FALSE,
            DEFAULT_CHARSET, OUT_TT_PRECIS, CLIP_DEFAULT_PRECIS,
            CLEARTYPE_NATURAL_QUALITY, DEFAULT_PITCH | FF_DONTCARE, GRID_FONT_NAME);
        
        SetBkMode(hdc, TRANSPARENT);
        
        // Draw title
        HFONT hPrevFont = (HFONT)SelectObject(hdc, hTitleFont);
        SetTextColor(hdc, g_palette.mainLabelText);
        RECT titleRect = { panelX + panelPad, panelY + panelPad, panelX + panelW - panelPad, panelY + titleH };
        DrawText(hdc, L"Minimized applications", -1, &titleRect, DT_LEFT | DT_SINGLELINE | DT_END_ELLIPSIS);
        
        // Draw items
        SelectObject(hdc, hItemFont);
        int totalNormal = (int)g_appWindows.size();
        for (int i = 0; i < maxItems; i++) {
            int itemY = panelY + titleH + i * lineH;
            RECT itemRect = { panelX + panelPad, itemY, panelX + panelW - panelPad, itemY + lineH };
            
            // Check if this minimized item is the current highlight
            bool isCurrent = (g_highlightIndex >= totalNormal && (g_highlightIndex - totalNormal) == i);
            
            if (isCurrent) {
                RECT hlRect = { panelX + borderT, itemY, panelX + panelW - borderT, itemY + lineH };
                HBRUSH hItemHl = CreateSolidBrush(g_palette.matchSubHighlightBg);
                FillRect(hdc, &hlRect, hItemHl);
                DeleteObject(hItemHl);
                SetTextColor(hdc, g_palette.matchSubHighlightText);
            } else {
                SetTextColor(hdc, g_palette.subLabelText);
            }
            
            wchar_t itemBuf[300];
            swprintf_s(itemBuf, L" %s", g_minimizedWindows[i].title.c_str());
            DrawText(hdc, itemBuf, -1, &itemRect, DT_LEFT | DT_SINGLELINE | DT_END_ELLIPSIS | DT_NOPREFIX);
        }
        
        SelectObject(hdc, hPrevFont);
        DeleteObject(hTitleFont);
        DeleteObject(hItemFont);
    }
}

// Enumerate visible application windows in Z-order (front to back)
BOOL CALLBACK EnumWindowsProc(HWND hwnd, LPARAM lParam) {
    // Skip our own windows
    if (hwnd == g_hOverlayWnd || hwnd == g_hMainWnd) return TRUE;
    
    // Must be visible
    if (!IsWindowVisible(hwnd)) return TRUE;
    
    // Must have a non-empty title
    wchar_t title[256] = {};
    GetWindowText(hwnd, title, 256);
    if (wcslen(title) == 0) return TRUE;
    
    // Skip tool windows and other non-app windows
    LONG exStyle = GetWindowLong(hwnd, GWL_EXSTYLE);
    if (exStyle & WS_EX_TOOLWINDOW) return TRUE;
    
    // Collect minimized windows separately
    if (IsIconic(hwnd)) {
        AppWindow aw;
        aw.hwnd = hwnd;
        GetWindowRect(hwnd, &aw.rect);
        aw.title = title;
        aw.visibleArea = 0;
        g_minimizedWindows.push_back(aw);
        return TRUE;
    }
    
    // Must have some area
    RECT rc;
    GetWindowRect(hwnd, &rc);
    if (rc.right - rc.left <= 0 || rc.bottom - rc.top <= 0) return TRUE;
    
    // Skip windows that are completely off-screen (e.g. cloaked UWP)
    BOOL cloaked = FALSE;
    // DwmGetWindowAttribute not linked; use owner check instead
    HWND owner = GetWindow(hwnd, GW_OWNER);
    if (owner != NULL) {
        // Skip owned windows (popups/dialogs) - we want top-level app windows
        // But some apps use owned windows, so only skip if also WS_EX_TOOLWINDOW
    }
    
    AppWindow aw;
    aw.hwnd = hwnd;
    aw.rect = rc;
    aw.title = title;
    
    auto* vec = reinterpret_cast<std::vector<AppWindow>*>(lParam);
    vec->push_back(aw);
    return TRUE;
}

void EnumerateAppWindows() {
    g_appWindows.clear();
    g_minimizedWindows.clear();
    EnumWindows(EnumWindowsProc, reinterpret_cast<LPARAM>(&g_appWindows));
    // EnumWindows returns in Z-order (front to back)
    
    // Calculate visible area for each window by subtracting regions of windows above it
    for (int i = 0; i < (int)g_appWindows.size(); i++) {
        const RECT& r = g_appWindows[i].rect;
        HRGN hRgn = CreateRectRgn(r.left, r.top, r.right, r.bottom);
        
        // Subtract all windows that are above this one (indices 0..i-1 are higher Z-order)
        for (int j = 0; j < i; j++) {
            const RECT& above = g_appWindows[j].rect;
            HRGN hAbove = CreateRectRgn(above.left, above.top, above.right, above.bottom);
            CombineRgn(hRgn, hRgn, hAbove, RGN_DIFF);
            DeleteObject(hAbove);
        }
        
        // Get the bounding box area of the remaining visible region
        // Use GetRegionData for exact pixel area
        DWORD size = GetRegionData(hRgn, 0, NULL);
        int visArea = 0;
        if (size > 0) {
            RGNDATA* pData = (RGNDATA*)malloc(size);
            if (pData && GetRegionData(hRgn, size, pData)) {
                RECT* pRects = (RECT*)pData->Buffer;
                for (DWORD k = 0; k < pData->rdh.nCount; k++) {
                    visArea += (pRects[k].right - pRects[k].left) * (pRects[k].bottom - pRects[k].top);
                }
            }
            free(pData);
        }
        g_appWindows[i].visibleArea = visArea;
        DeleteObject(hRgn);
    }
    
    // Sort by visible area descending (most visible windows first)
    std::sort(g_appWindows.begin(), g_appWindows.end(),
        [](const AppWindow& a, const AppWindow& b) {
            return a.visibleArea > b.visibleArea;
        });
    
    // Store all windows (including 0 visible area) for search
    g_allAppWindows = g_appWindows;
    
    // Store all minimized windows for search
    g_allMinimizedWindows = g_minimizedWindows;
    
    // Remove windows with 0 visible area from the cycling list
    g_appWindows.erase(
        std::remove_if(g_appWindows.begin(), g_appWindows.end(),
            [](const AppWindow& w) { return w.visibleArea <= 0; }),
        g_appWindows.end());
}

void FilterAppWindowsBySearch() {
    if (g_tabSearchStr.empty()) {
        // No search active - restore original visible-area-filtered list
        // Re-run enumeration to get fresh state
        return;
    }
    
    // Case-insensitive substring search across ALL windows (including occluded)
    std::wstring searchLower = g_tabSearchStr;
    for (auto& c : searchLower) c = towlower(c);
    
    g_appWindows.clear();
    for (const auto& aw : g_allAppWindows) {
        std::wstring titleLower = aw.title;
        for (auto& c : titleLower) c = towlower(c);
        if (titleLower.find(searchLower) != std::wstring::npos) {
            g_appWindows.push_back(aw);
        }
    }
    
    // Also filter minimized windows
    g_minimizedWindows.clear();
    for (const auto& aw : g_allMinimizedWindows) {
        std::wstring titleLower = aw.title;
        for (auto& c : titleLower) c = towlower(c);
        if (titleLower.find(searchLower) != std::wstring::npos) {
            g_minimizedWindows.push_back(aw);
        }
    }
    
    // If we have matches, highlight the first one
    if (!g_appWindows.empty()) {
        g_highlightIndex = 0;
    } else if (!g_minimizedWindows.empty()) {
        g_highlightIndex = 0;  // will point into minimized range (0 normal + i)
    } else {
        g_highlightIndex = -1;
    }
}

void CycleHighlight(bool forward) {
    if (g_appWindows.empty() && g_minimizedWindows.empty()) {
        EnumerateAppWindows();
    }
    int total = (int)g_appWindows.size() + (int)g_minimizedWindows.size();
    if (total == 0) return;
    
    if (forward) {
        g_highlightIndex++;
        if (g_highlightIndex >= total) {
            g_highlightIndex = 0;
        }
    } else {
        g_highlightIndex--;
        if (g_highlightIndex < 0) {
            g_highlightIndex = total - 1;
        }
    }
    
    // Make black background transparent via color key so red highlights stay vivid
    SetLayeredWindowAttributes(g_hOverlayWnd, g_palette.background, 0, LWA_COLORKEY);
    
    // Reset the tab-to-text timer on each TAB press
    g_bTabTextMode = false;
    KillTimer(g_hOverlayWnd, TIMER_ID_TAB_TEXT);
    SetTimer(g_hOverlayWnd, TIMER_ID_TAB_TEXT, TAB_TEXT_TIMEOUT_MS, NULL);
    
    InvalidateRect(g_hOverlayWnd, NULL, TRUE);
}

// Create overlay window
void CreateOverlayWindow() {
    if (g_hOverlayWnd) return;
    
    // Get virtual screen bounds (all monitors)
    auto vs = GetVirtualScreenBounds();
    int virtualLeft = vs.left;
    int virtualTop = vs.top;
    int virtualWidth = vs.width;
    int virtualHeight = vs.height;
    
    g_hOverlayWnd = CreateWindowEx(
        WS_EX_TOPMOST | WS_EX_TOOLWINDOW | WS_EX_LAYERED,
        L"KeyboardJockeyOverlay",
        L"Grid Overlay",
        WS_POPUP,
        virtualLeft, virtualTop, virtualWidth, virtualHeight,
        NULL, NULL, g_hInstance, NULL
    );
    
    // Set transparency
    SetLayeredWindowAttributes(g_hOverlayWnd, 0, GRID_ALPHA, LWA_ALPHA);
}

// Show the grid overlay
void ShowGrid() {
    if (g_bGridVisible) return;
    
    // Restore cursor in case it was hidden by typing (e.g., pressing hotkey)
    // Do it instantly without animation since we're showing the grid
    if (g_bCursorHidden) {
        g_bCursorAnimating = false;
        if (g_hMouseHook) {
            UnhookWindowsHookEx(g_hMouseHook);
            g_hMouseHook = NULL;
        }
        SystemParametersInfo(SPI_SETCURSORS, 0, NULL, 0);
        g_bCursorHidden = false;
    }
    
    g_typedChars.clear();
    g_bMouseMoveMode = false;  // Reset mouse move mode
    CreateOverlayWindow();
    
    // Reset transparency to default (in case it was reduced in a previous session)
    SetLayeredWindowAttributes(g_hOverlayWnd, 0, GRID_ALPHA, LWA_ALPHA);
    
    ShowWindow(g_hOverlayWnd, SW_SHOW);
    SetForegroundWindow(g_hOverlayWnd);
    SetFocus(g_hOverlayWnd);
    
    g_bGridVisible = true;
    InvalidateRect(g_hOverlayWnd, NULL, TRUE);
}

// Hide the grid overlay
void HideGrid() {
    if (!g_bGridVisible) return;
    
    ShowWindow(g_hOverlayWnd, SW_HIDE);
    g_bGridVisible = false;
    g_bMouseMoveMode = false;
    g_typedChars.clear();
    g_appWindows.clear();
    g_allAppWindows.clear();
    g_minimizedWindows.clear();
    g_allMinimizedWindows.clear();
    g_highlightIndex = -1;
    g_tabSearchStr.clear();
    g_bTabTextMode = false;
    KillTimer(g_hOverlayWnd, TIMER_ID_TAB_TEXT);
    // Clean up scroll mode if active
    if (g_bScrollMode) {
        g_bScrollMode = false;
        if (g_hScrollMouseHook) {
            UnhookWindowsHookEx(g_hScrollMouseHook);
            g_hScrollMouseHook = NULL;
        }
    }
}

// Move mouse to a point
void MoveMouse(POINT pt) {
    SetCursorPos(pt.x, pt.y);
}

// Send a mouse click
void SendClick(bool rightClick) {
    INPUT input[2] = {};
    
    input[0].type = INPUT_MOUSE;
    input[0].mi.dwFlags = rightClick ? MOUSEEVENTF_RIGHTDOWN : MOUSEEVENTF_LEFTDOWN;
    
    input[1].type = INPUT_MOUSE;
    input[1].mi.dwFlags = rightClick ? MOUSEEVENTF_RIGHTUP : MOUSEEVENTF_LEFTUP;
    
    SendInput(2, input, sizeof(INPUT));
}

// Get sub-grid point index from character a-h
// Layout: a b c
//         d X e
//         f g h
// Returns the subPoints index (0-8, skipping 4 which is center)
int GetSubPointIndex(wchar_t ch) {
    // a=0, b=1, c=2, d=3, e=5, f=6, g=7, h=8 (4 is center, skipped)
    if (ch >= L'a' && ch <= L'd') {
        return ch - L'a';  // 0-3
    } else if (ch >= L'e' && ch <= L'h') {
        return ch - L'a' + 1;  // 5-8 (skip index 4)
    }
    return 4;  // center as fallback
}

// Create main and sub-label fonts for a given sub-cell height and cell width
void CreateGridFonts(int sh, int cellW, HFONT* outMain, HFONT* outSub) {
    int fromHeight = sh * MAIN_FONT_HEIGHT_PCT / 100;
    int fromWidth = cellW / MAIN_FONT_WIDTH_DIV;
    int mainFontSize = -min(fromHeight, fromWidth);
    if (mainFontSize > MIN_MAIN_FONT_SIZE) mainFontSize = MIN_MAIN_FONT_SIZE;
    *outMain = CreateFont(mainFontSize, 0, 0, 0, FW_MEDIUM, FALSE, FALSE, FALSE,
        DEFAULT_CHARSET, OUT_TT_PRECIS, CLIP_DEFAULT_PRECIS,
        CLEARTYPE_NATURAL_QUALITY, DEFAULT_PITCH | FF_DONTCARE, GRID_FONT_NAME);
    int subFontSize = -(sh * SUB_FONT_HEIGHT_PCT / 100);
    if (subFontSize > MIN_SUB_FONT_SIZE) subFontSize = MIN_SUB_FONT_SIZE;
    *outSub = CreateFont(subFontSize, 0, 0, 0, FW_NORMAL, FALSE, FALSE, FALSE,
        DEFAULT_CHARSET, OUT_TT_PRECIS, CLIP_DEFAULT_PRECIS,
        CLEARTYPE_NATURAL_QUALITY, DEFAULT_PITCH | FF_DONTCARE, GRID_FONT_NAME);
}

// Move cursor by arrow key offset and enter mouse-move mode
void MoveMouseByArrowKey(int dx, int dy, int moveAmount) {
    POINT pt;
    GetCursorPos(&pt);
    pt.x += dx;
    pt.y += dy;
    SetCursorPos(pt.x, pt.y);
    if (!g_bMouseMoveMode) {
        g_bMouseMoveMode = true;
        SetLayeredWindowAttributes(g_hOverlayWnd, 0, MOUSE_MOVE_ALPHA, LWA_ALPHA);
        SetCursor(LoadCursor(NULL, IDC_ARROW));
    }
}

// Process typed character
void ProcessTypedChar(wchar_t ch) {
    if (!g_bGridVisible) return;
    
    // Convert to lowercase
    ch = towlower(ch);
    
    if (ch >= L'a' && ch <= L'z') {
        g_typedChars += ch;
        
        // Exit window highlight mode if active
        if (g_highlightIndex >= 0) {
            g_highlightIndex = -1;
            g_appWindows.clear();
            SetLayeredWindowAttributes(g_hOverlayWnd, 0, GRID_ALPHA, LWA_ALPHA);
            InvalidateRect(g_hOverlayWnd, NULL, TRUE);
        }
        
        // Restore default opacity when typing
        if (g_bMouseMoveMode) {
            g_bMouseMoveMode = false;
            SetLayeredWindowAttributes(g_hOverlayWnd, 0, GRID_ALPHA, LWA_ALPHA);
        }
        
        // Reset the timeout timer
        if (g_hOverlayWnd) {
            KillTimer(g_hOverlayWnd, TIMER_ID_RESET);
            SetTimer(g_hOverlayWnd, TIMER_ID_RESET, RESET_TIMEOUT_MS, NULL);
        }
        
        // Check if we have 4 characters (sub-grid selection)
        if (g_typedChars.length() == 4) {
            std::wstring threeChar = g_typedChars.substr(0, 3);
            wchar_t subChar = g_typedChars[3];
            
            // Find the cell
            for (const auto& cell : g_cells) {
                if (cell.label == threeChar) {
                    if (subChar >= L'a' && subChar <= L'h') {
                        int subIdx = GetSubPointIndex(subChar);
                        MoveMouse(cell.subPoints[subIdx]);
                    } else {
                        // Invalid sub-char, move to center
                        MoveMouse(cell.center);
                    }
                    break;
                }
            }
            // Clear for next selection
            g_typedChars.clear();
        }
        // If exactly 3 chars, move to cell center immediately
        else if (g_typedChars.length() == 3) {
            auto it = g_gridMap.find(g_typedChars);
            if (it != g_gridMap.end()) {
                MoveMouse(it->second);
            }
        }
        
        InvalidateRect(g_hOverlayWnd, NULL, TRUE);
    }
}

// Create tray icon
void CreateTrayIcon(HWND hWnd) {
    g_nid.cbSize = sizeof(NOTIFYICONDATA);
    g_nid.hWnd = hWnd;
    g_nid.uID = 1;
    g_nid.uFlags = NIF_ICON | NIF_MESSAGE | NIF_TIP;
    g_nid.uCallbackMessage = WM_TRAYICON;
    g_nid.hIcon = LoadIcon(g_hInstance, MAKEINTRESOURCE(IDI_KEYBOARDJOCKEY));
    wcscpy_s(g_nid.szTip, L"Keyboard Jockey - Ctrl+Alt+M to show grid");
    
    Shell_NotifyIcon(NIM_ADD, &g_nid);
}

// Remove tray icon
void RemoveTrayIcon() {
    Shell_NotifyIcon(NIM_DELETE, &g_nid);
}

// ========================================================================
// Palette picker window
// ========================================================================

// DPI-scaled layout for the palette window — all values computed at runtime
struct PalLayout {
    int winW, winH;
    int hueBarX, hueBarY, hueBarW, hueBarH;
    int markerH;
    int previewX, previewY, previewW, previewH;
    int btnW, btnH, btnY, btnOkX, btnCancelX;
    int fontLabel, fontSmall;   // font sizes (negative = point)
    float dpiScale;
};

static PalLayout g_palLayout = {};

static PalLayout ComputePalLayout() {
    // Get DPI of primary monitor
    UINT dpiX = DEFAULT_DPI, dpiY = DEFAULT_DPI;
    HMONITOR hMon = MonitorFromPoint({ 0, 0 }, MONITOR_DEFAULTTOPRIMARY);
    GetDpiForMonitor(hMon, MDT_EFFECTIVE_DPI, &dpiX, &dpiY);
    float s = (float)dpiX / 96.0f;

    PalLayout L;
    L.dpiScale  = s;
    L.winW      = (int)(620 * s);
    L.winH      = (int)(560 * s);
    int pad     = (int)(20 * s);

    L.hueBarX   = pad;
    L.hueBarY   = pad;
    L.hueBarW   = L.winW - pad * 2;
    L.hueBarH   = (int)(36 * s);
    L.markerH   = (int)(10 * s);

    L.btnH      = (int)(32 * s);
    L.btnW      = (int)(90 * s);
    L.btnY      = L.winH - pad - L.btnH;  // bottom area
    L.btnOkX    = L.winW - pad - L.btnW * 2 - (int)(10 * s);
    L.btnCancelX = L.winW - pad - L.btnW;

    L.previewX  = pad;
    L.previewY  = L.hueBarY + L.hueBarH + L.markerH + (int)(16 * s);
    L.previewW  = L.winW - pad * 2;
    L.previewH  = L.btnY - L.previewY - (int)(8 * s);

    L.fontLabel = -(int)(14 * s);
    L.fontSmall = -(int)(11 * s);

    return L;
}

#define IDC_PAL_OK      2001
#define IDC_PAL_CANCEL  2002

static bool g_bDraggingHue = false;
static float g_hueBeforeEdit = 0.0f;   // saved on dialog open for Cancel
static HWND g_hBtnOk = NULL;
static HWND g_hBtnCancel = NULL;
static HBITMAP g_hHueBarBitmap = NULL;  // cached rainbow strip (static, never changes)

// Map x pixel inside the hue bar to hue 0..360
static float HueBarPixelToHue(int x) {
    const PalLayout& L = g_palLayout;
    float t = (float)(x - L.hueBarX) / (float)L.hueBarW;
    if (t < 0.0f) t = 0.0f;
    if (t > 1.0f) t = 1.0f;
    return t * 360.0f;
}

static int HueToHueBarPixel(float hue) {
    const PalLayout& L = g_palLayout;
    return L.hueBarX + (int)(hue / 360.0f * L.hueBarW);
}

// Apply a new hue: regenerate palette, optionally rebuild grid bitmap, repaint
void ApplyHue(float hue) {
    g_baseHue = hue;
    if (g_baseHue < 0.0f)   g_baseHue = 0.0f;
    if (g_baseHue > 359.9f) g_baseHue = 359.9f;
    g_palette = GeneratePalette(g_baseHue);
    // Defer the expensive grid bitmap rebuild unless we're not dragging
    if (!g_bDraggingHue) {
        RenderBaseGridBitmap();
        if (g_bGridVisible) {
            InvalidateRect(g_hOverlayWnd, NULL, TRUE);
        }
    }
}

// Build the static hue rainbow bitmap (called once when palette window opens)
static void BuildHueBarBitmap() {
    if (g_hHueBarBitmap) { DeleteObject(g_hHueBarBitmap); g_hHueBarBitmap = NULL; }
    const PalLayout& L = g_palLayout;
    HDC hdcScreen = GetDC(NULL);
    HDC hdcMem = CreateCompatibleDC(hdcScreen);
    g_hHueBarBitmap = CreateCompatibleBitmap(hdcScreen, L.hueBarW, L.hueBarH);
    SelectObject(hdcMem, g_hHueBarBitmap);
    for (int x = 0; x < L.hueBarW; x++) {
        float h = (float)x / (float)L.hueBarW * 360.0f;
        COLORREF c = hsl(h, 0.85f, 0.50f);
        // SetPixelV for single-pixel columns is faster than brush create/fill/delete
        for (int y = 0; y < L.hueBarH; y++)
            SetPixelV(hdcMem, x, y, c);
    }
    DeleteDC(hdcMem);
    ReleaseDC(NULL, hdcScreen);
}

// Draw the horizontal hue rainbow bar (blits cached bitmap + draws marker)
static void PaintHueBar(HDC hdc) {
    const PalLayout& L = g_palLayout;
    // Blit cached rainbow
    if (g_hHueBarBitmap) {
        HDC hdcBmp = CreateCompatibleDC(hdc);
        SelectObject(hdcBmp, g_hHueBarBitmap);
        BitBlt(hdc, L.hueBarX, L.hueBarY, L.hueBarW, L.hueBarH, hdcBmp, 0, 0, SRCCOPY);
        DeleteDC(hdcBmp);
    }
    // Border
    HPEN hPen = CreatePen(PS_SOLID, 1, RGB(80, 80, 80));
    HPEN hOld = (HPEN)SelectObject(hdc, hPen);
    HBRUSH hNull = (HBRUSH)GetStockObject(NULL_BRUSH);
    HBRUSH hOldBr = (HBRUSH)SelectObject(hdc, hNull);
    Rectangle(hdc, L.hueBarX - 1, L.hueBarY - 1,
              L.hueBarX + L.hueBarW + 1, L.hueBarY + L.hueBarH + 1);
    SelectObject(hdc, hOldBr);
    SelectObject(hdc, hOld);
    DeleteObject(hPen);

    // Marker triangle below bar at current hue
    int triH = L.markerH;
    int mx = HueToHueBarPixel(g_baseHue);
    int triTop = L.hueBarY + L.hueBarH + 2;
    POINT tri[3] = {
        { mx, triTop },
        { mx - triH * 3 / 4, triTop + triH },
        { mx + triH * 3 / 4, triTop + triH }
    };
    HBRUSH hMarker = CreateSolidBrush(RGB(255, 255, 255));
    HPEN hMarkerPen = CreatePen(PS_SOLID, 1, RGB(40, 40, 40));
    SelectObject(hdc, hMarker);
    SelectObject(hdc, hMarkerPen);
    Polygon(hdc, tri, 3);
    SelectObject(hdc, hOld);
    DeleteObject(hMarker);
    DeleteObject(hMarkerPen);
}

// Draw a miniature preview of the grid + window highlight using the current palette
static void PaintPreview(HDC hdc) {
    const PalLayout& L = g_palLayout;
    int px = L.previewX, py = L.previewY;
    int pw = L.previewW, ph = L.previewH;
    float s = L.dpiScale;
    const Palette& P = g_palette;

    // Background for the whole preview area
    HBRUSH hBg = CreateSolidBrush(RGB(20, 20, 20));
    RECT rcBg = { px, py, px + pw, py + ph };
    FillRect(hdc, &rcBg, hBg);
    DeleteObject(hBg);

    // ---- Grid preview (left half) ----
    int pad = (int)(8 * s);
    int headerH = (int)(22 * s);
    int gridX = px + pad, gridY = py + headerH + pad;
    int gridW = pw / 2 - pad * 2, gridH = ph / 2 - headerH - pad;
    int cols = 4, rows = 3;
    int cellW = gridW / cols, cellH = gridH / rows;

    // Background fill
    HBRUSH hGridBg = CreateSolidBrush(P.background);
    RECT rcGrid = { gridX, gridY, gridX + cols * cellW, gridY + rows * cellH };
    FillRect(hdc, &rcGrid, hGridBg);
    DeleteObject(hGridBg);

    // DPI-scaled fonts (cached – only recreated when layout changes)
    static HFONT hSmallFont = NULL;
    static HFONT hTinyFont = NULL;
    static int sCachedLabelSz = 0, sCachedSmallSz = 0;
    if (!hSmallFont || sCachedLabelSz != L.fontLabel) {
        if (hSmallFont) DeleteObject(hSmallFont);
        hSmallFont = CreateFont(L.fontLabel, 0, 0, 0, FW_BOLD, FALSE, FALSE, FALSE,
            DEFAULT_CHARSET, OUT_TT_PRECIS, CLIP_DEFAULT_PRECIS,
            CLEARTYPE_NATURAL_QUALITY, DEFAULT_PITCH | FF_DONTCARE, GRID_FONT_NAME);
        sCachedLabelSz = L.fontLabel;
    }
    if (!hTinyFont || sCachedSmallSz != L.fontSmall) {
        if (hTinyFont) DeleteObject(hTinyFont);
        hTinyFont = CreateFont(L.fontSmall, 0, 0, 0, FW_NORMAL, FALSE, FALSE, FALSE,
            DEFAULT_CHARSET, OUT_TT_PRECIS, CLIP_DEFAULT_PRECIS,
            CLEARTYPE_NATURAL_QUALITY, DEFAULT_PITCH | FF_DONTCARE, GRID_FONT_NAME);
        sCachedSmallSz = L.fontSmall;
    }
    HFONT hOldFont = (HFONT)SelectObject(hdc, hSmallFont);
    SetBkMode(hdc, TRANSPARENT);
    SetTextColor(hdc, RGB(200, 200, 200));
    RECT rcLabel = { gridX, py + pad / 2, gridX + gridW, py + headerH };
    DrawText(hdc, L"Grid View", -1, &rcLabel, DT_LEFT | DT_SINGLELINE);

    // Checkerboard cells + grid lines
    HBRUSH hEven = CreateSolidBrush(P.cellBgEven);
    HBRUSH hOdd  = CreateSolidBrush(P.cellBgOdd);
    const wchar_t* demoLabels[] = {
        L"aaa", L"aab", L"aac", L"aad",
        L"aae", L"aaf", L"aag", L"aah",
        L"aai", L"aaj", L"aak", L"aal"
    };

    for (int r = 0; r < rows; r++) {
        for (int c = 0; c < cols; c++) {
            RECT rc = { gridX + c * cellW, gridY + r * cellH,
                        gridX + (c + 1) * cellW, gridY + (r + 1) * cellH };
            FillRect(hdc, &rc, ((r + c) % 2 == 0) ? hEven : hOdd);

            // Grid line border
            HPEN hPen = CreatePen(PS_SOLID, 1, P.gridLine);
            HPEN hOldPen = (HPEN)SelectObject(hdc, hPen);
            HBRUSH hNullBr = (HBRUSH)GetStockObject(NULL_BRUSH);
            HBRUSH hOldBr2 = (HBRUSH)SelectObject(hdc, hNullBr);
            Rectangle(hdc, rc.left, rc.top, rc.right, rc.bottom);
            SelectObject(hdc, hOldBr2);
            SelectObject(hdc, hOldPen);
            DeleteObject(hPen);

            // Sub-grid lines
            int sw = cellW / 3, sh = cellH / 3;
            HPEN hSub = CreatePen(PS_SOLID, 1, P.subGridLine);
            SelectObject(hdc, hSub);
            MoveToEx(hdc, rc.left + sw, rc.top, NULL); LineTo(hdc, rc.left + sw, rc.bottom);
            MoveToEx(hdc, rc.left + sw * 2, rc.top, NULL); LineTo(hdc, rc.left + sw * 2, rc.bottom);
            MoveToEx(hdc, rc.left, rc.top + sh, NULL); LineTo(hdc, rc.right, rc.top + sh);
            MoveToEx(hdc, rc.left, rc.top + sh * 2, NULL); LineTo(hdc, rc.right, rc.top + sh * 2);
            SelectObject(hdc, (HPEN)GetStockObject(BLACK_PEN));
            DeleteObject(hSub);

            // Cell label
            SelectObject(hdc, hTinyFont);
            int idx = r * cols + c;
            SetTextColor(hdc, P.mainLabelText);
            DrawText(hdc, demoLabels[idx], -1, &rc, DT_CENTERED);

            // Sub-labels
            const wchar_t* subs = L"abcdefgh";
            int si = 0;
            for (int sy = 0; sy < 3; sy++) {
                for (int sx = 0; sx < 3; sx++) {
                    if (sx == 1 && sy == 1) continue;
                    RECT sr = { rc.left + sx * sw, rc.top + sy * sh,
                                rc.left + (sx + 1) * sw, rc.top + (sy + 1) * sh };
                    SetTextColor(hdc, P.subLabelText);
                    wchar_t sl[2] = { subs[si], 0 };
                    DrawText(hdc, sl, 1, &sr, DT_CENTERED);
                    si++;
                }
            }
        }
    }
    DeleteObject(hEven);
    DeleteObject(hOdd);

    // ---- Typing preview (right half) - show match/partial/dim ----
    int typX = px + pw / 2 + pad, typY = py + headerH + pad;
    int typW = pw / 2 - pad * 2, typH = ph / 2 - headerH - pad;
    int tCols = 3, tRows = 3;
    int tCellW = typW / tCols, tCellH = typH / tRows;

    // Label
    SelectObject(hdc, hSmallFont);
    SetTextColor(hdc, RGB(200, 200, 200));
    RECT rcTypLabel = { typX, py + pad / 2, typX + typW, py + headerH };
    DrawText(hdc, L"Typing Match", -1, &rcTypLabel, DT_LEFT | DT_SINGLELINE);

    // Simulate: cell [1,1] = full match, [0,x] = partial, rest = dim
    for (int r = 0; r < tRows; r++) {
        for (int c = 0; c < tCols; c++) {
            RECT rc = { typX + c * tCellW, typY + r * tCellH,
                        typX + (c + 1) * tCellW, typY + (r + 1) * tCellH };
            bool isMatch = (r == 1 && c == 1);
            bool isPartial = (r == 0);

            COLORREF bg, fg;
            if (isMatch) {
                bg = P.matchCellBg; fg = P.matchLabelText;
            } else if (isPartial) {
                bg = P.partialMatchBg; fg = P.partialMatchText;
            } else {
                bg = P.dimBg; fg = P.dimText;
            }

            HBRUSH hCellBg = CreateSolidBrush(bg);
            FillRect(hdc, &rc, hCellBg);
            DeleteObject(hCellBg);

            // Grid lines
            HPEN hPen = CreatePen(PS_SOLID, 1, isMatch ? P.matchGridLine : P.gridLine);
            HPEN hOldPen = (HPEN)SelectObject(hdc, hPen);
            HBRUSH hNullBr = (HBRUSH)GetStockObject(NULL_BRUSH);
            SelectObject(hdc, hNullBr);
            Rectangle(hdc, rc.left, rc.top, rc.right, rc.bottom);
            SelectObject(hdc, hOldPen);
            DeleteObject(hPen);

            // Label
            SelectObject(hdc, hTinyFont);
            SetTextColor(hdc, fg);
            const wchar_t* cellLabel = isMatch ? L"aaf" : (isPartial ? L"aab" : L"abz");
            DrawText(hdc, cellLabel, -1, &rc, DT_CENTERED);

            // On the matched cell, draw sub-labels + one highlighted
            if (isMatch) {
                int sw2 = tCellW / 3, sh2 = tCellH / 3;
                HPEN hSub = CreatePen(PS_SOLID, 1, P.matchGridLine);
                SelectObject(hdc, hSub);
                MoveToEx(hdc, rc.left + sw2, rc.top, NULL); LineTo(hdc, rc.left + sw2, rc.bottom);
                MoveToEx(hdc, rc.left + sw2 * 2, rc.top, NULL); LineTo(hdc, rc.left + sw2 * 2, rc.bottom);
                MoveToEx(hdc, rc.left, rc.top + sh2, NULL); LineTo(hdc, rc.right, rc.top + sh2);
                MoveToEx(hdc, rc.left, rc.top + sh2 * 2, NULL); LineTo(hdc, rc.right, rc.top + sh2 * 2);
                SelectObject(hdc, (HPEN)GetStockObject(BLACK_PEN));
                DeleteObject(hSub);

                int si = 0;
                for (int sy = 0; sy < 3; sy++) {
                    for (int sx = 0; sx < 3; sx++) {
                        if (sx == 1 && sy == 1) continue;
                        RECT sr = { rc.left + sx * sw2, rc.top + sy * sh2,
                                    rc.left + (sx + 1) * sw2, rc.top + (sy + 1) * sh2 };
                        if (si == 3) {
                            HBRUSH hHl = CreateSolidBrush(P.matchSubHighlightBg);
                            FillRect(hdc, &sr, hHl);
                            DeleteObject(hHl);
                            SetTextColor(hdc, P.matchSubHighlightText);
                        } else {
                            SetTextColor(hdc, P.matchSubLabelText);
                        }
                        wchar_t sl[2] = { L"abcdefgh"[si], 0 };
                        DrawText(hdc, sl, 1, &sr, DT_CENTERED);
                        si++;
                    }
                }
            }
        }
    }

    // ---- Window highlight preview (bottom half) ----
    int winY = py + ph / 2 + headerH + pad;
    int winH = ph / 2 - headerH - pad * 2;
    int winW = pw - pad * 2;
    int winX = px + pad;

    SelectObject(hdc, hSmallFont);
    SetTextColor(hdc, RGB(200, 200, 200));
    RECT rcWinLabel = { winX, py + ph / 2 + pad / 2, winX + winW, py + ph / 2 + headerH };
    DrawText(hdc, L"Window Highlight (TAB mode)", -1, &rcWinLabel, DT_LEFT | DT_SINGLELINE);

    // Dark background
    HBRUSH hWinBg = CreateSolidBrush(P.background);
    RECT rcWin = { winX, winY, winX + winW, winY + winH };
    FillRect(hdc, &rcWin, hWinBg);
    DeleteObject(hWinBg);

    // Fake windows
    struct FakeWin { RECT r; const wchar_t* title; bool current; };
    int fw1 = winW * 55 / 100, fh1 = winH * 70 / 100;
    int fw2 = winW * 45 / 100, fh2 = winH * 60 / 100;
    FakeWin fakes[] = {
        { { winX + (int)(10*s), winY + (int)(24*s), winX + (int)(10*s) + fw1, winY + (int)(24*s) + fh1 }, L"[1/2] Visual Studio Code", true },
        { { winX + winW - fw2 - (int)(10*s), winY + (int)(12*s), winX + winW - (int)(10*s), winY + (int)(12*s) + fh2 }, L"[2/2] Firefox", false }
    };

    int thick = max(2, (int)(3 * s));
    for (const auto& fw : fakes) {
        COLORREF borderCol = fw.current ? P.mainLabelText : P.gridLine;
        HBRUSH hBorder = CreateSolidBrush(borderCol);

        RECT e;
        e = { fw.r.left, fw.r.top, fw.r.right, fw.r.top + thick };
        FillRect(hdc, &e, hBorder);
        e = { fw.r.left, fw.r.bottom - thick, fw.r.right, fw.r.bottom };
        FillRect(hdc, &e, hBorder);
        e = { fw.r.left, fw.r.top, fw.r.left + thick, fw.r.bottom };
        FillRect(hdc, &e, hBorder);
        e = { fw.r.right - thick, fw.r.top, fw.r.right, fw.r.bottom };
        FillRect(hdc, &e, hBorder);
        DeleteObject(hBorder);

        // Label above the window
        SIZE ts;
        GetTextExtentPoint32(hdc, fw.title, (int)wcslen(fw.title), &ts);
        RECT lblBg = { fw.r.left, fw.r.top - ts.cy - (int)(6*s),
                       fw.r.left + ts.cx + (int)(8*s), fw.r.top - 2 };
        HBRUSH hLbl = CreateSolidBrush(fw.current ? P.matchCellBg : P.cellBgEven);
        FillRect(hdc, &lblBg, hLbl);
        DeleteObject(hLbl);

        SelectObject(hdc, hTinyFont);
        SetTextColor(hdc, P.matchLabelText);
        RECT lblRc = { lblBg.left + (int)(4*s), lblBg.top + (int)(2*s), lblBg.right, lblBg.bottom };
        DrawText(hdc, fw.title, -1, &lblRc, DT_LEFT | DT_SINGLELINE | DT_NOPREFIX);
    }

    // Minimized panel preview
    int mpW = winW / 3, mpH = winH - (int)(10*s);
    int mpX = winX + winW - mpW - (int)(4*s), mpY = winY + (int)(4*s);
    RECT rcPanel = { mpX, mpY, mpX + mpW, mpY + mpH };
    HBRUSH hPanelBg = CreateSolidBrush(P.background);
    FillRect(hdc, &rcPanel, hPanelBg);
    DeleteObject(hPanelBg);

    // Panel border
    HBRUSH hPanelBorder = CreateSolidBrush(P.gridLine);
    RECT be;
    be = { rcPanel.left, rcPanel.top, rcPanel.right, rcPanel.top + 1 };
    FillRect(hdc, &be, hPanelBorder);
    be = { rcPanel.left, rcPanel.bottom - 1, rcPanel.right, rcPanel.bottom };
    FillRect(hdc, &be, hPanelBorder);
    be = { rcPanel.left, rcPanel.top, rcPanel.left + 1, rcPanel.bottom };
    FillRect(hdc, &be, hPanelBorder);
    be = { rcPanel.right - 1, rcPanel.top, rcPanel.right, rcPanel.bottom };
    FillRect(hdc, &be, hPanelBorder);
    DeleteObject(hPanelBorder);

    // Panel title
    SetTextColor(hdc, P.mainLabelText);
    SelectObject(hdc, hTinyFont);
    int itemH = (int)(18 * s);
    RECT rcPanelTitle = { mpX + (int)(4*s), mpY + (int)(3*s), mpX + mpW - (int)(4*s), mpY + itemH };
    DrawText(hdc, L"Minimized", -1, &rcPanelTitle, DT_LEFT | DT_SINGLELINE);

    // Panel items
    const wchar_t* items[] = { L"Notepad", L"Calculator", L"Slack" };
    int itemTop = mpY + itemH + (int)(2*s);
    for (int i = 0; i < 3; i++) {
        int iy = itemTop + i * itemH;
        RECT ir = { mpX + (int)(4*s), iy, mpX + mpW - (int)(4*s), iy + itemH };
        if (i == 0) {
            RECT hlr = { mpX + 1, iy, mpX + mpW - 1, iy + itemH };
            HBRUSH hHl = CreateSolidBrush(P.matchSubHighlightBg);
            FillRect(hdc, &hlr, hHl);
            DeleteObject(hHl);
            SetTextColor(hdc, P.matchSubHighlightText);
        } else {
            SetTextColor(hdc, P.subLabelText);
        }
        DrawText(hdc, items[i], -1, &ir, DT_LEFT | DT_SINGLELINE | DT_NOPREFIX);
    }

    // Hue value label at bottom of preview
    SelectObject(hdc, hSmallFont);
    SetTextColor(hdc, RGB(200, 200, 200));
    wchar_t hueBuf[64];
    swprintf_s(hueBuf, L"Hue: %.0f\u00b0", g_baseHue);
    RECT rcHue = { px, py + ph - (int)(22*s), px + pw, py + ph };
    DrawText(hdc, hueBuf, -1, &rcHue, DT_CENTER | DT_SINGLELINE);

    // Cleanup (fonts are cached statics, don't delete here)
    SelectObject(hdc, hOldFont);
}

// --- Registry persistence for user settings ---
static const wchar_t* REG_KEY = L"Software\\KeyboardJockey";

static void SaveHueToRegistry(float hue) {
    HKEY hKey;
    if (RegCreateKeyEx(HKEY_CURRENT_USER, REG_KEY, 0, NULL,
            REG_OPTION_NON_VOLATILE, KEY_WRITE, NULL, &hKey, NULL) == ERROR_SUCCESS) {
        DWORD val = (DWORD)(hue * 100.0f);  // store as fixed-point (2 decimals)
        RegSetValueEx(hKey, L"BaseHue", 0, REG_DWORD, (BYTE*)&val, sizeof(val));
        RegCloseKey(hKey);
    }
}

static float LoadHueFromRegistry() {
    HKEY hKey;
    if (RegOpenKeyEx(HKEY_CURRENT_USER, REG_KEY, 0, KEY_READ, &hKey) == ERROR_SUCCESS) {
        DWORD val = 0, sz = sizeof(val), type = 0;
        if (RegQueryValueEx(hKey, L"BaseHue", NULL, &type, (BYTE*)&val, &sz) == ERROR_SUCCESS
            && type == REG_DWORD) {
            RegCloseKey(hKey);
            float h = (float)val / 100.0f;
            if (h >= 0.0f && h < 360.0f) return h;
        } else {
            RegCloseKey(hKey);
        }
    }
    return BASE_HUE_DEFAULT;  // no saved value
}

// Palette window procedure
LRESULT CALLBACK PaletteWndProc(HWND hWnd, UINT message, WPARAM wParam, LPARAM lParam) {
    switch (message) {
    case WM_CREATE: {
        const PalLayout& L = g_palLayout;
        // Create OK and Cancel buttons
        HFONT hBtnFont = CreateFont(L.fontLabel, 0, 0, 0, FW_NORMAL, FALSE, FALSE, FALSE,
            DEFAULT_CHARSET, OUT_TT_PRECIS, CLIP_DEFAULT_PRECIS,
            CLEARTYPE_NATURAL_QUALITY, DEFAULT_PITCH | FF_DONTCARE, GRID_FONT_NAME);
        g_hBtnOk = CreateWindowEx(0, L"BUTTON", L"OK",
            WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON,
            L.btnOkX, L.btnY, L.btnW, L.btnH,
            hWnd, (HMENU)IDC_PAL_OK, g_hInstance, NULL);
        g_hBtnCancel = CreateWindowEx(0, L"BUTTON", L"Cancel",
            WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON,
            L.btnCancelX, L.btnY, L.btnW, L.btnH,
            hWnd, (HMENU)IDC_PAL_CANCEL, g_hInstance, NULL);
        SendMessage(g_hBtnOk, WM_SETFONT, (WPARAM)hBtnFont, TRUE);
        SendMessage(g_hBtnCancel, WM_SETFONT, (WPARAM)hBtnFont, TRUE);
        // hBtnFont intentionally not deleted — owned by the buttons for their lifetime
        return 0;
    }

    case WM_PAINT: {
        PAINTSTRUCT ps;
        HDC hdc = BeginPaint(hWnd, &ps);

        // Double-buffer
        RECT rc;
        GetClientRect(hWnd, &rc);
        int w = rc.right, h = rc.bottom;
        HDC hdcMem = CreateCompatibleDC(hdc);
        HBITMAP hbm = CreateCompatibleBitmap(hdc, w, h);
        HBITMAP hOldBm = (HBITMAP)SelectObject(hdcMem, hbm);

        // Fill with dark gray
        HBRUSH hBg = CreateSolidBrush(RGB(30, 30, 30));
        FillRect(hdcMem, &rc, hBg);
        DeleteObject(hBg);

        PaintHueBar(hdcMem);
        PaintPreview(hdcMem);

        BitBlt(hdc, 0, 0, w, h, hdcMem, 0, 0, SRCCOPY);
        SelectObject(hdcMem, hOldBm);
        DeleteObject(hbm);
        DeleteDC(hdcMem);

        EndPaint(hWnd, &ps);
        return 0;
    }

    case WM_LBUTTONDOWN: {
        const PalLayout& L = g_palLayout;
        int mx = LOWORD(lParam), my = HIWORD(lParam);
        if (mx >= L.hueBarX && mx <= L.hueBarX + L.hueBarW &&
            my >= L.hueBarY - (int)(4*L.dpiScale) &&
            my <= L.hueBarY + L.hueBarH + L.markerH + (int)(8*L.dpiScale)) {
            g_bDraggingHue = true;
            SetCapture(hWnd);
            float newHue = HueBarPixelToHue(mx);
            ApplyHue(newHue);
            InvalidateRect(hWnd, NULL, FALSE);
        }
        return 0;
    }

    case WM_MOUSEMOVE: {
        if (g_bDraggingHue) {
            int mx = (short)LOWORD(lParam);
            float newHue = HueBarPixelToHue(mx);
            ApplyHue(newHue);
            InvalidateRect(hWnd, NULL, FALSE);
        }
        return 0;
    }

    case WM_LBUTTONUP:
        if (g_bDraggingHue) {
            g_bDraggingHue = false;
            ReleaseCapture();
            // Now do the deferred expensive rebuild
            RenderBaseGridBitmap();
            if (g_bGridVisible) {
                InvalidateRect(g_hOverlayWnd, NULL, TRUE);
            }
        }
        return 0;

    case WM_COMMAND:
        switch (LOWORD(wParam)) {
        case IDC_PAL_OK:
            // Accept — save to registry and close
            SaveHueToRegistry(g_baseHue);
            DestroyWindow(hWnd);
            break;
        case IDC_PAL_CANCEL:
            // Revert to the hue we had when the dialog opened
            ApplyHue(g_hueBeforeEdit);
            DestroyWindow(hWnd);
            break;
        }
        return 0;

    case WM_CLOSE:
        // Closing via X button = Cancel
        ApplyHue(g_hueBeforeEdit);
        DestroyWindow(hWnd);
        return 0;

    case WM_DESTROY:
        g_hBtnOk = NULL;
        g_hBtnCancel = NULL;
        g_hPaletteWnd = NULL;
        if (g_hHueBarBitmap) { DeleteObject(g_hHueBarBitmap); g_hHueBarBitmap = NULL; }
        return 0;

    default:
        return DefWindowProc(hWnd, message, wParam, lParam);
    }
    return 0;
}

void ShowPaletteWindow() {
    // If already open, bring to front
    if (g_hPaletteWnd && IsWindow(g_hPaletteWnd)) {
        SetForegroundWindow(g_hPaletteWnd);
        return;
    }

    // Save current hue for Cancel
    g_hueBeforeEdit = g_baseHue;

    // Compute DPI-aware layout
    g_palLayout = ComputePalLayout();
    const PalLayout& L = g_palLayout;

    // Build cached hue bar bitmap
    BuildHueBarBitmap();

    // Compute outer window size from desired client area
    DWORD style = WS_OVERLAPPED | WS_CAPTION | WS_SYSMENU | WS_CLIPCHILDREN;
    RECT rcDesired = { 0, 0, L.winW, L.winH };
    AdjustWindowRectEx(&rcDesired, style, FALSE, WS_EX_APPWINDOW);
    int outerW = rcDesired.right - rcDesired.left;
    int outerH = rcDesired.bottom - rcDesired.top;

    // Center on primary monitor
    int sx = GetSystemMetrics(SM_CXSCREEN);
    int sy = GetSystemMetrics(SM_CYSCREEN);
    int wx = (sx - outerW) / 2;
    int wy = (sy - outerH) / 2;

    g_hPaletteWnd = CreateWindowEx(
        WS_EX_APPWINDOW,
        L"KeyboardJockeyPalette",
        L"Keyboard Jockey \u2013 Palette",
        style,
        wx, wy, outerW, outerH,
        NULL, NULL, g_hInstance, NULL
    );

    ShowWindow(g_hPaletteWnd, SW_SHOW);
    UpdateWindow(g_hPaletteWnd);
}

// Show context menu
void ShowContextMenu(HWND hWnd) {
    POINT pt;
    GetCursorPos(&pt);
    
    HMENU hMenu = CreatePopupMenu();
    AppendMenu(hMenu, MF_STRING, IDM_SHOW, L"Show Grid (Ctrl+Alt+M)");
    AppendMenu(hMenu, MF_STRING, IDM_PALETTE, L"Palette...");
    AppendMenu(hMenu, MF_SEPARATOR, 0, NULL);
    AppendMenu(hMenu, MF_STRING, IDM_EXIT, L"Exit");
    
    SetForegroundWindow(hWnd);
    TrackPopupMenu(hMenu, TPM_RIGHTBUTTON, pt.x, pt.y, 0, hWnd, NULL);
    DestroyMenu(hMenu);
}

// Overlay window procedure
LRESULT CALLBACK OverlayWndProc(HWND hWnd, UINT message, WPARAM wParam, LPARAM lParam) {
    switch (message) {
    case WM_ERASEBKGND:
        return 1;  // Skip erase — we paint the full surface ourselves
    
    case WM_SETCURSOR:
        // Show normal arrow cursor during mouse-move and scroll modes instead of cross
        if (g_bMouseMoveMode || g_bScrollMode) {
            SetCursor(LoadCursor(NULL, IDC_ARROW));
            return TRUE;
        }
        return DefWindowProc(hWnd, message, wParam, lParam);
    
    case WM_PAINT: {
        PAINTSTRUCT ps;
        HDC hdc = BeginPaint(hWnd, &ps);
        
        // Double-buffer: paint to off-screen bitmap, then blit
        RECT rc;
        GetClientRect(hWnd, &rc);
        int w = rc.right - rc.left;
        int h = rc.bottom - rc.top;
        HDC hdcMem = CreateCompatibleDC(hdc);
        HBITMAP hbmMem = CreateCompatibleBitmap(hdc, w, h);
        HBITMAP hbmOld = (HBITMAP)SelectObject(hdcMem, hbmMem);
        
        PaintGrid(hdcMem);
        
        BitBlt(hdc, 0, 0, w, h, hdcMem, 0, 0, SRCCOPY);
        
        SelectObject(hdcMem, hbmOld);
        DeleteObject(hbmMem);
        DeleteDC(hdcMem);
        
        EndPaint(hWnd, &ps);
        return 0;
    }
    
    case WM_KEYDOWN: {
        // In scroll mode, only PgUp/PgDn/Escape are allowed; anything else exits
        if (g_bScrollMode && wParam != VK_PRIOR && wParam != VK_NEXT && wParam != VK_ESCAPE) {
            ExitScrollMode();
            return 0;
        }
        
        // In TAB/text mode, only allow relevant keys; block grid navigation
        bool inTabMode = (g_highlightIndex >= 0 || g_bTabTextMode || !g_tabSearchStr.empty());
        if (inTabMode) {
            switch (wParam) {
            case VK_ESCAPE:
                HideGrid();
                return 0;
            case VK_RETURN: {
                // Activate highlighted window (or single search match)
                int totalNormal = (int)g_appWindows.size();
                HWND targetWnd = NULL;
                if (g_highlightIndex >= 0 && g_highlightIndex < totalNormal) {
                    targetWnd = g_appWindows[g_highlightIndex].hwnd;
                } else if (g_highlightIndex >= totalNormal && (g_highlightIndex - totalNormal) < (int)g_minimizedWindows.size()) {
                    targetWnd = g_minimizedWindows[g_highlightIndex - totalNormal].hwnd;
                }
                if (targetWnd) {
                    HideGrid();
                    Sleep(ACTIVATION_DELAY_MS);
                    SetForegroundWindow(targetWnd);
                    if (IsIconic(targetWnd)) {
                        ShowWindow(targetWnd, SW_RESTORE);
                    }
                }
                }
                return 0;
            case VK_TAB: {
                // If in text search mode, TAB exits back to normal cycling
                if (g_bTabTextMode || !g_tabSearchStr.empty()) {
                    g_bTabTextMode = false;
                    g_tabSearchStr.clear();
                    // Restore to visible-only windows for normal cycling
                    g_appWindows.clear();
                    for (const auto& aw : g_allAppWindows) {
                        if (aw.visibleArea > 0) g_appWindows.push_back(aw);
                    }
                    g_minimizedWindows.clear();  // hide minimized panel
                    if (!g_appWindows.empty()) g_highlightIndex = 0;
                    else g_highlightIndex = -1;
                    SetLayeredWindowAttributes(g_hOverlayWnd, g_palette.background, 0, LWA_COLORKEY);
                    KillTimer(hWnd, TIMER_ID_TAB_TEXT);
                    SetTimer(hWnd, TIMER_ID_TAB_TEXT, TAB_TEXT_TIMEOUT_MS, NULL);
                    InvalidateRect(hWnd, NULL, TRUE);
                } else {
                    bool shiftTab = (GetKeyState(VK_SHIFT) & 0x8000) != 0;
                    CycleHighlight(!shiftTab);
                }
                return 0;
            }
            case VK_BACK:
                if (!g_tabSearchStr.empty()) {
                    g_tabSearchStr.pop_back();
                    if (g_tabSearchStr.empty()) {
                        g_appWindows.clear();
                        if (g_bTabTextMode) {
                            g_appWindows = g_allAppWindows;
                        } else {
                            for (const auto& aw : g_allAppWindows) {
                                if (aw.visibleArea > 0) g_appWindows.push_back(aw);
                            }
                        }
                        g_minimizedWindows = g_allMinimizedWindows;
                        if (!g_appWindows.empty()) g_highlightIndex = 0;
                        else if (!g_minimizedWindows.empty()) g_highlightIndex = 0;
                        else g_highlightIndex = -1;
                    } else {
                        FilterAppWindowsBySearch();
                    }
                    InvalidateRect(hWnd, NULL, TRUE);
                }
                return 0;
            default:
                // Block all other keys (arrows, space, etc.) in TAB mode
                return 0;
            }
        }
        
        // Check modifiers for movement amount
        bool shiftHeld = (GetKeyState(VK_SHIFT) & 0x8000) != 0;
        bool ctrlHeld = (GetKeyState(VK_CONTROL) & 0x8000) != 0;
        int moveAmount = shiftHeld ? 1 : (ctrlHeld ? 50 : 10);
        
        switch (wParam) {
        case VK_ESCAPE:
            HideGrid();
            break;
        case VK_LEFT:
            MoveMouseByArrowKey(-moveAmount, 0, moveAmount);
            break;
        case VK_RIGHT:
            MoveMouseByArrowKey(moveAmount, 0, moveAmount);
            break;
        case VK_UP:
            MoveMouseByArrowKey(0, -moveAmount, moveAmount);
            break;
        case VK_DOWN:
            MoveMouseByArrowKey(0, moveAmount, moveAmount);
            break;
        case VK_SHIFT:
            // Make UI subtle when Shift is pressed (skip in highlight mode)
            if (g_highlightIndex < 0) {
                BYTE peekAlpha = g_bMouseMoveMode ? MOUSE_MOVE_ALPHA : SHIFT_PEEK_ALPHA;
                SetLayeredWindowAttributes(g_hOverlayWnd, 0, peekAlpha, LWA_ALPHA);
            }
            break;
        case VK_SPACE: {
            // Hide the cursor and close the grid
            HideGrid();
            HideCursor();
            break;
        }
        case VK_RETURN: {
            // If in window highlight mode, activate the highlighted window
            HWND targetWnd2 = NULL;
            int totalNormal2 = (int)g_appWindows.size();
            if (g_highlightIndex >= 0 && g_highlightIndex < totalNormal2) {
                targetWnd2 = g_appWindows[g_highlightIndex].hwnd;
            } else if (g_highlightIndex >= totalNormal2 && (g_highlightIndex - totalNormal2) < (int)g_minimizedWindows.size()) {
                targetWnd2 = g_minimizedWindows[g_highlightIndex - totalNormal2].hwnd;
            }
            if (targetWnd2) {
                HideGrid();
                Sleep(ACTIVATION_DELAY_MS);
                // Activate the window
                SetForegroundWindow(targetWnd2);
                // If minimized, restore it
                if (IsIconic(targetWnd2)) {
                    ShowWindow(targetWnd2, SW_RESTORE);
                }
                break;
            }
            // If we have 3 chars typed, move to center before clicking
            if (g_typedChars.length() >= 3) {
                std::wstring threeChar = g_typedChars.substr(0, 3);
                auto it = g_gridMap.find(threeChar);
                if (it != g_gridMap.end()) {
                    MoveMouse(it->second);
                }
            }
            HideGrid();
            // Small delay to let grid disappear
            Sleep(ACTIVATION_DELAY_MS);
            // Check if Ctrl is held for right-click
            if (GetKeyState(VK_CONTROL) & 0x8000) {
                SendClick(true);  // Right click
            } else {
                SendClick(false); // Left click
            }
            break;
        }
        case VK_TAB: {
            bool shiftTab = (GetKeyState(VK_SHIFT) & 0x8000) != 0;
            CycleHighlight(!shiftTab);
            break;
        }
        case VK_BACK:
            // In TAB search mode, delete from search string
            if (g_highlightIndex >= 0 && !g_tabSearchStr.empty()) {
                g_tabSearchStr.pop_back();
                if (g_tabSearchStr.empty()) {
                    // Restore to cycling mode with all visible windows
                    g_appWindows.clear();
                    for (const auto& aw : g_allAppWindows) {
                        if (aw.visibleArea > 0) g_appWindows.push_back(aw);
                    }
                    g_minimizedWindows = g_allMinimizedWindows;
                    if (!g_appWindows.empty()) g_highlightIndex = 0;
                    else if (!g_minimizedWindows.empty()) g_highlightIndex = 0;
                    else g_highlightIndex = -1;
                } else {
                    FilterAppWindowsBySearch();
                }
                InvalidateRect(hWnd, NULL, TRUE);
            } else if (!g_typedChars.empty()) {
                g_typedChars.pop_back();
                InvalidateRect(hWnd, NULL, TRUE);
            }
            break;
        case VK_PRIOR:   // PgUp
        case VK_NEXT: {  // PgDn
            // Enter scroll mode: make overlay invisible, send wheel to window under cursor
            if (!g_bScrollMode) {
                g_bScrollMode = true;
                // Make overlay visually transparent via color key
                SetLayeredWindowAttributes(g_hOverlayWnd, g_palette.background, 0, LWA_COLORKEY);
                InvalidateRect(g_hOverlayWnd, NULL, TRUE);
                UpdateWindow(g_hOverlayWnd);
                SetCursor(LoadCursor(NULL, IDC_ARROW));
                // Make overlay transparent to ALL input (mouse clicks, wheel, etc.)
                LONG_PTR exStyle = GetWindowLongPtr(g_hOverlayWnd, GWL_EXSTYLE);
                SetWindowLongPtr(g_hOverlayWnd, GWL_EXSTYLE, exStyle | WS_EX_TRANSPARENT);
                // Install mouse hook to detect movement
                g_hScrollMouseHook = SetWindowsHookEx(WH_MOUSE_LL, ScrollMouseProc, g_hInstance, 0);
            }
            // Simulate mouse wheel scroll (passes through to window under cursor)
            {
                INPUT input = {};
                input.type = INPUT_MOUSE;
                input.mi.dwFlags = MOUSEEVENTF_WHEEL;
                input.mi.mouseData = (wParam == VK_PRIOR) ? (DWORD)(3 * WHEEL_DELTA) : (DWORD)(-3 * WHEEL_DELTA);
                SendInput(1, &input, sizeof(INPUT));
            }
            break;
        }
        }
        return 0;
    }
    
    case WM_CHAR: {
        wchar_t ch = (wchar_t)wParam;
        // Any typing in scroll mode exits it
        if (g_bScrollMode) {
            ExitScrollMode();
            return 0;
        }
        if (ch == L'*' && (g_highlightIndex >= 0 || !g_tabSearchStr.empty()) && !g_bTabTextMode) {
            // Enter "all windows" mode immediately
            g_bTabTextMode = true;
            g_tabSearchStr.clear();
            KillTimer(hWnd, TIMER_ID_TAB_TEXT);
            g_appWindows = g_allAppWindows;
            g_minimizedWindows = g_allMinimizedWindows;
            if (!g_appWindows.empty()) g_highlightIndex = 0;
            else if (!g_minimizedWindows.empty()) g_highlightIndex = 0;
            // Combine color-key + alpha: background pixels = invisible, panel = semi-transparent
            SetLayeredWindowAttributes(g_hOverlayWnd, g_palette.background, GRID_ALPHA, LWA_COLORKEY | LWA_ALPHA);
            InvalidateRect(hWnd, NULL, TRUE);
            return 0;
        }
        if (ch >= L'a' && ch <= L'z') {
            // In TAB highlight/text mode, typing filters windows by substring
            if (g_highlightIndex >= 0 || g_bTabTextMode || !g_tabSearchStr.empty()) {
                g_tabSearchStr += ch;
                FilterAppWindowsBySearch();
                // Reset the tab text timeout timer (don't use TIMER_ID_RESET for search)
                KillTimer(hWnd, TIMER_ID_TAB_TEXT);
                SetTimer(hWnd, TIMER_ID_TAB_TEXT, TAB_TEXT_TIMEOUT_MS, NULL);
                InvalidateRect(hWnd, NULL, TRUE);
            } else {
                ProcessTypedChar(ch);
            }
        }
        return 0;
    }
    
    case WM_KEYUP: {
        // Restore UI when Shift is released (skip in highlight mode)
        if (wParam == VK_SHIFT && g_highlightIndex < 0) {
            // Restore to mouse move mode transparency or full opacity
            if (g_bMouseMoveMode) {
                SetLayeredWindowAttributes(g_hOverlayWnd, 0, MOUSE_MOVE_ALPHA, LWA_ALPHA);
            } else {
                SetLayeredWindowAttributes(g_hOverlayWnd, 0, GRID_ALPHA, LWA_ALPHA);
            }
        }
        return 0;
    }
    
    case WM_KILLFOCUS:
        // Hide grid when losing focus
        HideGrid();
        return 0;
    
    case WM_TIMER:
        if (wParam == TIMER_ID_RESET) {
            KillTimer(hWnd, TIMER_ID_RESET);
            // In TAB search mode, clear search but stay in highlight mode
            if ((g_highlightIndex >= 0 || g_bTabTextMode) && !g_tabSearchStr.empty()) {
                g_tabSearchStr.clear();
                // Restore all windows (if in text mode) or visible-area ones
                g_appWindows.clear();
                if (g_bTabTextMode) {
                    g_appWindows = g_allAppWindows;
                } else {
                    for (const auto& aw : g_allAppWindows) {
                        if (aw.visibleArea > 0) g_appWindows.push_back(aw);
                    }
                }
                g_minimizedWindows = g_allMinimizedWindows;
                if (!g_appWindows.empty()) g_highlightIndex = 0;
                else if (!g_minimizedWindows.empty()) g_highlightIndex = 0;
                else g_highlightIndex = -1;
                InvalidateRect(hWnd, NULL, TRUE);
            } else if (!g_typedChars.empty()) {
                g_typedChars.clear();
                InvalidateRect(hWnd, NULL, TRUE);
            }
        }
        else if (wParam == TIMER_ID_TAB_TEXT) {
            KillTimer(hWnd, TIMER_ID_TAB_TEXT);
            // Enter "select by text" mode - show ALL windows highlighted
            if (g_highlightIndex >= 0 && !g_bTabTextMode) {
                g_bTabTextMode = true;
                g_tabSearchStr.clear();
                // Switch to all windows (including fully occluded)
                g_appWindows = g_allAppWindows;
                g_minimizedWindows = g_allMinimizedWindows;
                if (!g_appWindows.empty()) g_highlightIndex = 0;
                else if (!g_minimizedWindows.empty()) g_highlightIndex = 0;
                // Combine color-key + alpha: background pixels = invisible, panel = semi-transparent
                SetLayeredWindowAttributes(g_hOverlayWnd, g_palette.background, GRID_ALPHA, LWA_COLORKEY | LWA_ALPHA);
                InvalidateRect(hWnd, NULL, TRUE);
            }
        }
        return 0;
    
    default:
        return DefWindowProc(hWnd, message, wParam, lParam);
    }
}

// Main window procedure
LRESULT CALLBACK MainWndProc(HWND hWnd, UINT message, WPARAM wParam, LPARAM lParam) {
    switch (message) {
    case WM_CREATE:
        CreateTrayIcon(hWnd);
        // Register global hotkey: Ctrl+Alt+M
        if (!RegisterHotKey(hWnd, HOTKEY_ID_SHOW_GRID, MOD_CONTROL | MOD_ALT, 'M')) {
            MessageBox(hWnd, L"Failed to register hotkey Ctrl+Alt+M", L"Error", MB_ICONERROR);
        }
        // Pre-build grid cells and cache base grid bitmap
        BuildGridCells();
        RenderBaseGridBitmap();
        // Install global keyboard hook for cursor hiding while typing
        InstallGlobalKeyboardHook();
        return 0;
    
    case WM_HOTKEY:
        if (wParam == HOTKEY_ID_SHOW_GRID) {
            if (g_bGridVisible) {
                HideGrid();
            } else {
                ShowGrid();
            }
        }
        return 0;
    
    case WM_TRAYICON:
        switch (LOWORD(lParam)) {
        case WM_RBUTTONUP:
            ShowContextMenu(hWnd);
            break;
        case WM_LBUTTONDBLCLK:
            ShowGrid();
            break;
        }
        return 0;
    
    case WM_COMMAND:
        switch (LOWORD(wParam)) {
        case IDM_EXIT:
            DestroyWindow(hWnd);
            break;
        case IDM_SHOW:
            ShowGrid();
            break;
        case IDM_PALETTE:
            ShowPaletteWindow();
            break;
        }
        return 0;
    
    case WM_DESTROY:
        UnregisterHotKey(hWnd, HOTKEY_ID_SHOW_GRID);
        RemoveTrayIcon();
        if (g_hGridBitmap) { DeleteObject(g_hGridBitmap); g_hGridBitmap = NULL; }
        UninstallGlobalKeyboardHook();  // Remove keyboard hook
        RestoreCursor();  // Make sure cursor is restored on exit
        if (g_hOverlayWnd) {
            DestroyWindow(g_hOverlayWnd);
        }
        PostQuitMessage(0);
        return 0;
    
    default:
        return DefWindowProc(hWnd, message, wParam, lParam);
    }
}

// Entry point
int WINAPI wWinMain(HINSTANCE hInstance, HINSTANCE hPrevInstance, LPWSTR lpCmdLine, int nCmdShow) {
    // Enable per-monitor DPI awareness for correct coordinates on mixed-DPI setups
    SetProcessDpiAwarenessContext(DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE_V2);
    
    g_hInstance = hInstance;

    // Load saved hue from registry and apply it
    g_baseHue = LoadHueFromRegistry();
    g_palette = GeneratePalette(g_baseHue);
    
    // Save a copy of the default arrow cursor before we ever modify system cursors
    HCURSOR hArrow = LoadCursor(NULL, IDC_ARROW);
    if (hArrow) {
        g_hSavedArrow = CopyCursor(hArrow);
    }
    
    // Register cleanup handlers for crash/force-kill scenarios
    atexit(ForceRestoreCursors);
    SetUnhandledExceptionFilter(CrashHandler);
    
    // Register main window class
    WNDCLASSEX wcMain = {};
    wcMain.cbSize = sizeof(WNDCLASSEX);
    wcMain.lpfnWndProc = MainWndProc;
    wcMain.hInstance = hInstance;
    wcMain.hIcon = LoadIcon(hInstance, MAKEINTRESOURCE(IDI_KEYBOARDJOCKEY));
    wcMain.hCursor = LoadCursor(NULL, IDC_ARROW);
    wcMain.hbrBackground = (HBRUSH)(COLOR_WINDOW + 1);
    wcMain.lpszClassName = L"KeyboardJockeyMain";
    wcMain.hIconSm = LoadIcon(hInstance, MAKEINTRESOURCE(IDI_KEYBOARDJOCKEY));
    
    if (!RegisterClassEx(&wcMain)) {
        MessageBox(NULL, L"Failed to register main window class", L"Error", MB_ICONERROR);
        return 1;
    }
    
    // Register overlay window class
    WNDCLASSEX wcOverlay = {};
    wcOverlay.cbSize = sizeof(WNDCLASSEX);
    wcOverlay.lpfnWndProc = OverlayWndProc;
    wcOverlay.hInstance = hInstance;
    wcOverlay.hCursor = LoadCursor(NULL, IDC_CROSS);
    wcOverlay.hbrBackground = NULL;
    wcOverlay.lpszClassName = L"KeyboardJockeyOverlay";
    
    if (!RegisterClassEx(&wcOverlay)) {
        MessageBox(NULL, L"Failed to register overlay window class", L"Error", MB_ICONERROR);
        return 1;
    }
    
    // Register palette picker window class
    WNDCLASSEX wcPalette = {};
    wcPalette.cbSize = sizeof(WNDCLASSEX);
    wcPalette.lpfnWndProc = PaletteWndProc;
    wcPalette.hInstance = hInstance;
    wcPalette.hCursor = LoadCursor(NULL, IDC_ARROW);
    wcPalette.hbrBackground = (HBRUSH)(COLOR_WINDOW + 1);
    wcPalette.lpszClassName = L"KeyboardJockeyPalette";
    wcPalette.hIcon = LoadIcon(hInstance, MAKEINTRESOURCE(IDI_KEYBOARDJOCKEY));
    wcPalette.hIconSm = LoadIcon(hInstance, MAKEINTRESOURCE(IDI_KEYBOARDJOCKEY));
    RegisterClassEx(&wcPalette);
    
    // Create main window (hidden)
    g_hMainWnd = CreateWindowEx(
        0,
        L"KeyboardJockeyMain",
        L"Keyboard Jockey",
        WS_OVERLAPPEDWINDOW,
        CW_USEDEFAULT, CW_USEDEFAULT, 400, 300,
        NULL, NULL, hInstance, NULL
    );
    
    if (!g_hMainWnd) {
        MessageBox(NULL, L"Failed to create main window", L"Error", MB_ICONERROR);
        return 1;
    }
    
    // Don't show the main window - it minimizes to tray immediately
    // ShowWindow(g_hMainWnd, SW_HIDE);
    
    // Message loop
    MSG msg;
    while (GetMessage(&msg, NULL, 0, 0)) {
        TranslateMessage(&msg);
        DispatchMessage(&msg);
    }
    
    return (int)msg.wParam;
}
