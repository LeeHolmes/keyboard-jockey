// KeyboardJockey.cpp - A keyboard-driven mouse navigation utility
// Minimizes to tray, shows grid overlay on Ctrl+Alt+M, type letters to move mouse

#define WIN32_LEAN_AND_MEAN
#define OEMRESOURCE  // Required for OCR_* cursor constants

#include <windows.h>
#include <shellapi.h>
#include <vector>
#include <string>
#include <map>
#include <cstdlib>
#include <algorithm>
#include <thread>

// Resource IDs
#define IDI_KEYBOARDJOCKEY 101
#define IDM_EXIT 1001
#define IDM_SHOW 1002

// Constants
#define WM_TRAYICON (WM_USER + 1)
#define HOTKEY_ID_SHOW_GRID 1
#define GRID_COLS 32
#define GRID_ROWS 12
#define TIMER_ID_RESET 1
#define TIMER_ID_TAB_TEXT 2
#define RESET_TIMEOUT_MS 1000
#define TAB_TEXT_TIMEOUT_MS 2000

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

// Grid cell structure
struct GridCell {
    RECT rect;
    std::wstring label;
    POINT center;
    POINT subPoints[9];  // 3x3 sub-grid points (0-8, center is 4)
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
int g_highlightIndex = -1;  // -1 = no highlight active
std::wstring g_tabSearchStr;  // Substring search in TAB mode

// Forward declarations
LRESULT CALLBACK MainWndProc(HWND, UINT, WPARAM, LPARAM);
LRESULT CALLBACK OverlayWndProc(HWND, UINT, WPARAM, LPARAM);
void CreateTrayIcon(HWND hWnd);
void RemoveTrayIcon();
void ShowContextMenu(HWND hWnd);
void CreateOverlayWindow();
void ShowGrid();
void HideGrid();
void BuildGridCells();
void PaintGrid(HDC hdc);
void ProcessTypedChar(wchar_t ch);
void MoveMouse(POINT pt);
void SendClick(bool rightClick);
std::wstring GenerateLabel(int index);
void FilterAppWindowsBySearch();
void HideCursor();
void RestoreCursor();
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
    HCURSOR hCopy;
    hCopy = CopyCursor(hBlankCursor);
    SetSystemCursor(hCopy, OCR_NORMAL);
    hCopy = CopyCursor(hBlankCursor);
    SetSystemCursor(hCopy, OCR_IBEAM);
    hCopy = CopyCursor(hBlankCursor);
    SetSystemCursor(hCopy, OCR_HAND);
    hCopy = CopyCursor(hBlankCursor);
    SetSystemCursor(hCopy, OCR_CROSS);
    hCopy = CopyCursor(hBlankCursor);
    SetSystemCursor(hCopy, OCR_SIZEALL);
    hCopy = CopyCursor(hBlankCursor);
    SetSystemCursor(hCopy, OCR_SIZENWSE);
    hCopy = CopyCursor(hBlankCursor);
    SetSystemCursor(hCopy, OCR_SIZENESW);
    hCopy = CopyCursor(hBlankCursor);
    SetSystemCursor(hCopy, OCR_SIZEWE);
    hCopy = CopyCursor(hBlankCursor);
    SetSystemCursor(hCopy, OCR_SIZENS);
    
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
    
    DWORD cursorIDs[] = { OCR_NORMAL, OCR_IBEAM, OCR_HAND, OCR_CROSS,
                          OCR_SIZEALL, OCR_SIZENWSE, OCR_SIZENESW, OCR_SIZEWE, OCR_SIZENS };
    for (DWORD id : cursorIDs) {
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

// Generate label like "aa", "ab", "ac", ... "ba", "bb", etc.
std::wstring GenerateLabel(int index) {
    std::wstring label;
    // Use 2-character combinations: aa-zz gives us 676 combinations
    int first = index / 26;
    int second = index % 26;
    
    if (first < 26) {
        label += static_cast<wchar_t>(L'a' + first);
        label += static_cast<wchar_t>(L'a' + second);
    }
    return label;
}

// Build grid cells for all monitors
void BuildGridCells() {
    g_cells.clear();
    g_gridMap.clear();
    
    // Get virtual screen bounds (all monitors)
    int virtualLeft = GetSystemMetrics(SM_XVIRTUALSCREEN);
    int virtualTop = GetSystemMetrics(SM_YVIRTUALSCREEN);
    int virtualWidth = GetSystemMetrics(SM_CXVIRTUALSCREEN);
    int virtualHeight = GetSystemMetrics(SM_CYVIRTUALSCREEN);
    
    // Calculate cell size
    int cellWidth = virtualWidth / GRID_COLS;
    int cellHeight = virtualHeight / GRID_ROWS;
    
    // Sub-cell size (1/3 of cell)
    int subWidth = cellWidth / 3;
    int subHeight = cellHeight / 3;
    
    int index = 0;
    for (int row = 0; row < GRID_ROWS; row++) {
        for (int col = 0; col < GRID_COLS; col++) {
            GridCell cell;
            cell.rect.left = virtualLeft + col * cellWidth;
            cell.rect.top = virtualTop + row * cellHeight;
            cell.rect.right = cell.rect.left + cellWidth;
            cell.rect.bottom = cell.rect.top + cellHeight;
            cell.center.x = cell.rect.left + cellWidth / 2;
            cell.center.y = cell.rect.top + cellHeight / 2;
            cell.label = GenerateLabel(index);
            
            // Calculate 3x3 sub-grid points
            // Layout: a b c
            //         d X e
            //         f g h
            // Where X is center (index 4)
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

// Paint the grid overlay
void PaintGrid(HDC hdc) {
    // Get virtual screen bounds
    int virtualLeft = GetSystemMetrics(SM_XVIRTUALSCREEN);
    int virtualTop = GetSystemMetrics(SM_YVIRTUALSCREEN);
    int virtualWidth = GetSystemMetrics(SM_CXVIRTUALSCREEN);
    int virtualHeight = GetSystemMetrics(SM_CYVIRTUALSCREEN);
    
    // Semi-transparent background
    HBRUSH hBrushBg = CreateSolidBrush(RGB(0, 0, 0));
    RECT rcFull = { 0, 0, virtualWidth, virtualHeight };
    FillRect(hdc, &rcFull, hBrushBg);
    DeleteObject(hBrushBg);
    
    // In scroll mode, paint only black (LWA_COLORKEY makes it fully transparent)
    if (g_bScrollMode) return;
    
    // If in window highlight or text select mode, skip drawing the grid entirely
    bool highlightMode = (g_highlightIndex >= 0 || g_bTabTextMode || !g_tabSearchStr.empty());
    
    // Sub-grid labels: a-h around center (positions 0-3, 5-8, skipping 4 which is center)
    const wchar_t* subLabels = L"abcdefgh";
    
    // Calculate cell dimensions
    int cellWidth = virtualWidth / GRID_COLS;
    int cellHeight = virtualHeight / GRID_ROWS;
    int subWidth = cellWidth / 3;
    int subHeight = cellHeight / 3;
    
  if (!highlightMode) {
    // Grid lines (scale pen widths)
    int gridPenWidth = max(1, virtualHeight / 800);
    HPEN hPen = CreatePen(PS_SOLID, gridPenWidth, RGB(120, 120, 120));
    HPEN hSubPen = CreatePen(PS_SOLID, max(1, gridPenWidth / 2), RGB(60, 80, 100));  // Subtle sub-grid lines
    HPEN hOldPen = (HPEN)SelectObject(hdc, hPen);

    for (const auto& cell : g_cells) {
        // Adjust coordinates for window position
        RECT adjusted;
        adjusted.left = cell.rect.left - virtualLeft;
        adjusted.top = cell.rect.top - virtualTop;
        adjusted.right = cell.rect.right - virtualLeft;
        adjusted.bottom = cell.rect.bottom - virtualTop;
        
        // Draw cell border
        SelectObject(hdc, hPen);
        MoveToEx(hdc, adjusted.left, adjusted.top, NULL);
        LineTo(hdc, adjusted.right, adjusted.top);
        LineTo(hdc, adjusted.right, adjusted.bottom);
        LineTo(hdc, adjusted.left, adjusted.bottom);
        LineTo(hdc, adjusted.left, adjusted.top);
        
        // Draw subtle 3x3 sub-grid lines
        SelectObject(hdc, hSubPen);
        // Vertical sub-lines
        MoveToEx(hdc, adjusted.left + subWidth, adjusted.top, NULL);
        LineTo(hdc, adjusted.left + subWidth, adjusted.bottom);
        MoveToEx(hdc, adjusted.left + subWidth * 2, adjusted.top, NULL);
        LineTo(hdc, adjusted.left + subWidth * 2, adjusted.bottom);
        // Horizontal sub-lines
        MoveToEx(hdc, adjusted.left, adjusted.top + subHeight, NULL);
        LineTo(hdc, adjusted.right, adjusted.top + subHeight);
        MoveToEx(hdc, adjusted.left, adjusted.top + subHeight * 2, NULL);
        LineTo(hdc, adjusted.right, adjusted.top + subHeight * 2);
    }
    
    SelectObject(hdc, hOldPen);
    DeleteObject(hPen);
    DeleteObject(hSubPen);
    
    // Draw labels
    SetBkMode(hdc, TRANSPARENT);
    
    // Font for main labels (center) - scale to ~30% of cell height
    int mainFontSize = -(cellHeight * 30 / 100);
    if (mainFontSize > -8) mainFontSize = -8;  // minimum readable size
    HFONT hFont = CreateFont(mainFontSize, 0, 0, 0, FW_MEDIUM, FALSE, FALSE, FALSE,
        DEFAULT_CHARSET, OUT_TT_PRECIS, CLIP_DEFAULT_PRECIS,
        CLEARTYPE_NATURAL_QUALITY, DEFAULT_PITCH | FF_DONTCARE, L"Segoe UI Variable Display");
    
    // Font for sub-labels - scale to ~15% of cell height
    int subFontSize = -(cellHeight * 15 / 100);
    if (subFontSize > -6) subFontSize = -6;
    HFONT hSubFont = CreateFont(subFontSize, 0, 0, 0, FW_NORMAL, FALSE, FALSE, FALSE,
        DEFAULT_CHARSET, OUT_TT_PRECIS, CLIP_DEFAULT_PRECIS,
        CLEARTYPE_NATURAL_QUALITY, DEFAULT_PITCH | FF_DONTCARE, L"Segoe UI Variable Display");
    
    HFONT hOldFont = (HFONT)SelectObject(hdc, hFont);
    
    for (const auto& cell : g_cells) {
        RECT adjusted;
        adjusted.left = cell.rect.left - virtualLeft;
        adjusted.top = cell.rect.top - virtualTop;
        adjusted.right = cell.rect.right - virtualLeft;
        adjusted.bottom = cell.rect.bottom - virtualTop;
        
        // Check if this cell matches typed chars
        bool isMatch = false;
        bool isPartialMatch = false;
        
        if (g_typedChars.length() >= 2) {
            std::wstring twoChar = g_typedChars.substr(0, 2);
            if (cell.label == twoChar) {
                isMatch = true;
            }
        } else if (g_typedChars.length() == 1) {
            if (cell.label[0] == g_typedChars[0]) {
                isPartialMatch = true;
            }
        }
        
        // Highlight matching cell
        if (isMatch) {
            HBRUSH hHighlight = CreateSolidBrush(RGB(0, 100, 0));
            FillRect(hdc, &adjusted, hHighlight);
            DeleteObject(hHighlight);
            
            // Redraw sub-grid lines on highlighted cell
            HPEN hSubPenLight = CreatePen(PS_SOLID, 1, RGB(0, 150, 0));
            SelectObject(hdc, hSubPenLight);
            MoveToEx(hdc, adjusted.left + subWidth, adjusted.top, NULL);
            LineTo(hdc, adjusted.left + subWidth, adjusted.bottom);
            MoveToEx(hdc, adjusted.left + subWidth * 2, adjusted.top, NULL);
            LineTo(hdc, adjusted.left + subWidth * 2, adjusted.bottom);
            MoveToEx(hdc, adjusted.left, adjusted.top + subHeight, NULL);
            LineTo(hdc, adjusted.right, adjusted.top + subHeight);
            MoveToEx(hdc, adjusted.left, adjusted.top + subHeight * 2, NULL);
            LineTo(hdc, adjusted.right, adjusted.top + subHeight * 2);
            DeleteObject(hSubPenLight);
        }
        
        // Draw center label (main 2-letter code)
        int descPad = cellHeight / 10;  // descender padding
        RECT centerRect;
        centerRect.left = adjusted.left + subWidth;
        centerRect.top = adjusted.top + subHeight;
        centerRect.right = adjusted.left + subWidth * 2;
        centerRect.bottom = adjusted.top + subHeight * 2 + descPad;
        
        SelectObject(hdc, hFont);
        if (isMatch) {
            SetTextColor(hdc, RGB(255, 255, 255));
        } else if (isPartialMatch) {
            SetTextColor(hdc, RGB(150, 230, 255));
        } else if (!g_typedChars.empty()) {
            SetTextColor(hdc, RGB(100, 100, 100));
        } else {
            SetTextColor(hdc, RGB(0, 200, 255));
        }
        
        DrawText(hdc, cell.label.c_str(), -1, &centerRect, 
            DT_CENTER | DT_VCENTER | DT_SINGLELINE);
        
        // Draw sub-labels a-h around the center
        SelectObject(hdc, hSubFont);
        if (isMatch) {
            SetTextColor(hdc, RGB(200, 255, 200));
        } else {
            SetTextColor(hdc, RGB(140, 180, 220));
        }
        
        int subLabelIdx = 0;
        for (int sy = 0; sy < 3; sy++) {
            for (int sx = 0; sx < 3; sx++) {
                if (sx == 1 && sy == 1) continue;  // Skip center
                
                RECT subRect;
                subRect.left = adjusted.left + sx * subWidth;
                subRect.top = adjusted.top + sy * subHeight;
                subRect.right = subRect.left + subWidth;
                subRect.bottom = subRect.top + subHeight + descPad;
                
                // Highlight the selected sub-cell if 3 chars typed
                if (isMatch && g_typedChars.length() == 3) {
                    wchar_t subChar = g_typedChars[2];
                    if (subChar >= L'a' && subChar <= L'h') {
                        int targetSubIdx = subChar - L'a';
                        if (targetSubIdx == subLabelIdx) {
                            HBRUSH hSubHighlight = CreateSolidBrush(RGB(0, 150, 50));
                            FillRect(hdc, &subRect, hSubHighlight);
                            DeleteObject(hSubHighlight);
                            SetTextColor(hdc, RGB(255, 255, 255));
                        }
                    }
                }
                
                wchar_t subLabel[2] = { subLabels[subLabelIdx], 0 };
                DrawText(hdc, subLabel, 1, &subRect, 
                    DT_CENTER | DT_VCENTER | DT_SINGLELINE);
                
                if (isMatch) {
                    SetTextColor(hdc, RGB(200, 255, 200));
                }
                
                subLabelIdx++;
            }
        }
    }
    
    SelectObject(hdc, hOldFont);
    DeleteObject(hFont);
    DeleteObject(hSubFont);
  } // end if (!highlightMode)
    if (g_highlightIndex >= 0 && !g_appWindows.empty()) {
        // When search is active or in text mode, highlight ALL matching windows
        // When just cycling (no search), highlight only the current one
        bool showAll = !g_tabSearchStr.empty() || g_bTabTextMode;
        int startIdx = showAll ? 0 : g_highlightIndex;
        int endIdx = showAll ? (int)g_appWindows.size() : g_highlightIndex + 1;
        
        HFONT hLabelFont = CreateFont(-(max(12, virtualHeight / 80)), 0, 0, 0, FW_BOLD, FALSE, FALSE, FALSE,
            DEFAULT_CHARSET, OUT_TT_PRECIS, CLIP_DEFAULT_PRECIS,
            CLEARTYPE_NATURAL_QUALITY, DEFAULT_PITCH | FF_DONTCARE, L"Segoe UI Variable Display");
        
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
            HBRUSH hBorderBrush = CreateSolidBrush(isCurrent ? RGB(255, 0, 0) : RGB(255, 100, 100));
            
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
            HBRUSH hLabelBg = CreateSolidBrush(isCurrent ? RGB(200, 0, 0) : RGB(150, 50, 50));
            FillRect(hdc, &labelBg, hLabelBg);
            DeleteObject(hLabelBg);
            
            SetTextColor(hdc, RGB(255, 255, 255));
            SetBkMode(hdc, TRANSPARENT);
            RECT labelRect = { labelX + 4, labelY + 2, labelBg.right, labelBg.bottom };
            DrawText(hdc, labelBuf, -1, &labelRect, DT_LEFT | DT_SINGLELINE | DT_NOPREFIX);
            
            SelectObject(hdc, hPrevFont);
        }
        
        DeleteObject(hLabelFont);
    }
}

// Enumerate visible application windows in Z-order (front to back)
BOOL CALLBACK EnumWindowsProc(HWND hwnd, LPARAM lParam) {
    // Skip our own windows
    if (hwnd == g_hOverlayWnd || hwnd == g_hMainWnd) return TRUE;
    
    // Must be visible
    if (!IsWindowVisible(hwnd)) return TRUE;
    
    // Skip minimized windows
    if (IsIconic(hwnd)) return TRUE;
    
    // Must have a non-empty title
    wchar_t title[256] = {};
    GetWindowText(hwnd, title, 256);
    if (wcslen(title) == 0) return TRUE;
    
    // Skip tool windows and other non-app windows
    LONG exStyle = GetWindowLong(hwnd, GWL_EXSTYLE);
    if (exStyle & WS_EX_TOOLWINDOW) return TRUE;
    
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
    
    // If we have matches, highlight the first one
    if (!g_appWindows.empty()) {
        g_highlightIndex = 0;
    } else {
        g_highlightIndex = -1;
    }
}

void CycleHighlight(bool forward) {
    if (g_appWindows.empty()) {
        EnumerateAppWindows();
    }
    if (g_appWindows.empty()) return;
    
    if (forward) {
        g_highlightIndex++;
        if (g_highlightIndex >= (int)g_appWindows.size()) {
            g_highlightIndex = 0;
        }
    } else {
        g_highlightIndex--;
        if (g_highlightIndex < 0) {
            g_highlightIndex = (int)g_appWindows.size() - 1;
        }
    }
    
    // Make black background transparent via color key so red highlights stay vivid
    SetLayeredWindowAttributes(g_hOverlayWnd, RGB(0, 0, 0), 0, LWA_COLORKEY);
    
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
    int virtualLeft = GetSystemMetrics(SM_XVIRTUALSCREEN);
    int virtualTop = GetSystemMetrics(SM_YVIRTUALSCREEN);
    int virtualWidth = GetSystemMetrics(SM_CXVIRTUALSCREEN);
    int virtualHeight = GetSystemMetrics(SM_CYVIRTUALSCREEN);
    
    g_hOverlayWnd = CreateWindowEx(
        WS_EX_TOPMOST | WS_EX_TOOLWINDOW | WS_EX_LAYERED,
        L"KeyboardJockeyOverlay",
        L"Grid Overlay",
        WS_POPUP,
        virtualLeft, virtualTop, virtualWidth, virtualHeight,
        NULL, NULL, g_hInstance, NULL
    );
    
    // Set transparency (200 out of 255)
    SetLayeredWindowAttributes(g_hOverlayWnd, 0, 200, LWA_ALPHA);
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
    BuildGridCells();
    CreateOverlayWindow();
    
    // Reset transparency to full opacity (in case it was reduced in a previous session)
    SetLayeredWindowAttributes(g_hOverlayWnd, 0, 200, LWA_ALPHA);
    
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
            SetLayeredWindowAttributes(g_hOverlayWnd, 0, 200, LWA_ALPHA);
            InvalidateRect(g_hOverlayWnd, NULL, TRUE);
        }
        
        // Restore full opacity when typing
        if (g_bMouseMoveMode) {
            g_bMouseMoveMode = false;
            SetLayeredWindowAttributes(g_hOverlayWnd, 0, 200, LWA_ALPHA);
        }
        
        // Reset the timeout timer
        if (g_hOverlayWnd) {
            KillTimer(g_hOverlayWnd, TIMER_ID_RESET);
            SetTimer(g_hOverlayWnd, TIMER_ID_RESET, RESET_TIMEOUT_MS, NULL);
        }
        
        // Check if we have 3 characters (sub-grid selection)
        if (g_typedChars.length() == 3) {
            std::wstring twoChar = g_typedChars.substr(0, 2);
            wchar_t subChar = g_typedChars[2];
            
            // Find the cell
            for (const auto& cell : g_cells) {
                if (cell.label == twoChar) {
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
        // If exactly 2 chars, move to center immediately
        else if (g_typedChars.length() == 2) {
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

// Show context menu
void ShowContextMenu(HWND hWnd) {
    POINT pt;
    GetCursorPos(&pt);
    
    HMENU hMenu = CreatePopupMenu();
    AppendMenu(hMenu, MF_STRING, IDM_SHOW, L"Show Grid (Ctrl+Alt+M)");
    AppendMenu(hMenu, MF_SEPARATOR, 0, NULL);
    AppendMenu(hMenu, MF_STRING, IDM_EXIT, L"Exit");
    
    SetForegroundWindow(hWnd);
    TrackPopupMenu(hMenu, TPM_RIGHTBUTTON, pt.x, pt.y, 0, hWnd, NULL);
    DestroyMenu(hMenu);
}

// Overlay window procedure
LRESULT CALLBACK OverlayWndProc(HWND hWnd, UINT message, WPARAM wParam, LPARAM lParam) {
    switch (message) {
    case WM_PAINT: {
        PAINTSTRUCT ps;
        HDC hdc = BeginPaint(hWnd, &ps);
        PaintGrid(hdc);
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
            case VK_RETURN:
                // Activate highlighted window (or single search match)
                if (g_highlightIndex >= 0 && g_highlightIndex < (int)g_appWindows.size()) {
                    HWND targetWnd = g_appWindows[g_highlightIndex].hwnd;
                    HideGrid();
                    Sleep(50);
                    SetForegroundWindow(targetWnd);
                    if (IsIconic(targetWnd)) {
                        ShowWindow(targetWnd, SW_RESTORE);
                    }
                }
                return 0;
            case VK_TAB: {
                bool shiftTab = (GetKeyState(VK_SHIFT) & 0x8000) != 0;
                CycleHighlight(!shiftTab);
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
                        if (!g_appWindows.empty()) g_highlightIndex = 0;
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
        case VK_LEFT: {
            POINT pt;
            GetCursorPos(&pt);
            pt.x -= moveAmount;
            SetCursorPos(pt.x, pt.y);
            // Go 50% more transparent for mouse movement mode
            if (!g_bMouseMoveMode) {
                g_bMouseMoveMode = true;
                SetLayeredWindowAttributes(g_hOverlayWnd, 0, 100, LWA_ALPHA);
            }
            break;
        }
        case VK_RIGHT: {
            POINT pt;
            GetCursorPos(&pt);
            pt.x += moveAmount;
            SetCursorPos(pt.x, pt.y);
            if (!g_bMouseMoveMode) {
                g_bMouseMoveMode = true;
                SetLayeredWindowAttributes(g_hOverlayWnd, 0, 100, LWA_ALPHA);
            }
            break;
        }
        case VK_UP: {
            POINT pt;
            GetCursorPos(&pt);
            pt.y -= moveAmount;
            SetCursorPos(pt.x, pt.y);
            if (!g_bMouseMoveMode) {
                g_bMouseMoveMode = true;
                SetLayeredWindowAttributes(g_hOverlayWnd, 0, 100, LWA_ALPHA);
            }
            break;
        }
        case VK_DOWN: {
            POINT pt;
            GetCursorPos(&pt);
            pt.y += moveAmount;
            SetCursorPos(pt.x, pt.y);
            if (!g_bMouseMoveMode) {
                g_bMouseMoveMode = true;
                SetLayeredWindowAttributes(g_hOverlayWnd, 0, 100, LWA_ALPHA);
            }
            break;
        }
        case VK_SHIFT:
            // Make UI subtle when Shift is pressed (skip in highlight mode)
            if (g_highlightIndex < 0) {
                SetLayeredWindowAttributes(g_hOverlayWnd, 0, 100, LWA_ALPHA);
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
            if (g_highlightIndex >= 0 && g_highlightIndex < (int)g_appWindows.size()) {
                HWND targetWnd = g_appWindows[g_highlightIndex].hwnd;
                HideGrid();
                Sleep(50);
                // Activate the window
                SetForegroundWindow(targetWnd);
                // If minimized, restore it
                if (IsIconic(targetWnd)) {
                    ShowWindow(targetWnd, SW_RESTORE);
                }
                break;
            }
            // If we have 2 chars typed, move to center before clicking
            if (g_typedChars.length() >= 2) {
                std::wstring twoChar = g_typedChars.substr(0, 2);
                auto it = g_gridMap.find(twoChar);
                if (it != g_gridMap.end()) {
                    MoveMouse(it->second);
                }
            }
            HideGrid();
            // Small delay to let grid disappear
            Sleep(50);
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
                    if (!g_appWindows.empty()) g_highlightIndex = 0;
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
            // Enter scroll mode: make overlay invisible, send key to window under cursor
            if (!g_bScrollMode) {
                g_bScrollMode = true;
                // Make overlay fully transparent but keep it for focus
                SetLayeredWindowAttributes(g_hOverlayWnd, RGB(0, 0, 0), 0, LWA_COLORKEY);
                InvalidateRect(g_hOverlayWnd, NULL, TRUE);
                // Install mouse hook to detect movement
                g_hScrollMouseHook = SetWindowsHookEx(WH_MOUSE_LL, ScrollMouseProc, g_hInstance, 0);
            }
            // Simulate mouse wheel scroll (page-sized: 3 * WHEEL_DELTA)
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
            if (!g_appWindows.empty()) g_highlightIndex = 0;
            SetLayeredWindowAttributes(g_hOverlayWnd, RGB(0, 0, 0), 0, LWA_COLORKEY);
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
                SetLayeredWindowAttributes(g_hOverlayWnd, 0, 100, LWA_ALPHA);
            } else {
                SetLayeredWindowAttributes(g_hOverlayWnd, 0, 200, LWA_ALPHA);
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
                if (!g_appWindows.empty()) g_highlightIndex = 0;
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
                if (!g_appWindows.empty()) g_highlightIndex = 0;
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
        }
        return 0;
    
    case WM_DESTROY:
        UnregisterHotKey(hWnd, HOTKEY_ID_SHOW_GRID);
        RemoveTrayIcon();
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
    wcOverlay.hbrBackground = (HBRUSH)GetStockObject(BLACK_BRUSH);
    wcOverlay.lpszClassName = L"KeyboardJockeyOverlay";
    
    if (!RegisterClassEx(&wcOverlay)) {
        MessageBox(NULL, L"Failed to register overlay window class", L"Error", MB_ICONERROR);
        return 1;
    }
    
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
