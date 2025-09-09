#define NOMINMAX
#define UNICODE
#define _UNICODE

#include <algorithm>

#include "photomesh_parser.hpp"
#include "realitymesh_parser.hpp"
#include "excel_writer.hpp"
#include "single_sheet_writer.hpp"
#include "models.hpp"
#include "util_time.hpp"

#include <windows.h>
#include <shellapi.h>
#include <shobjidl.h>
#include <commctrl.h>
#include <dwmapi.h>
#include <uxtheme.h>
#include <vssym32.h>
#include <filesystem>
#include <string>
#include <vector>
#include <fstream>
#include <sstream>

#pragma comment(lib, "Comctl32.lib")
#pragma comment(lib, "Ole32.lib")
#pragma comment(lib, "Shell32.lib")
#pragma comment(lib, "Uuid.lib")
#pragma comment(lib, "Dwmapi.lib")
#pragma comment(lib, "UxTheme.lib")
// optional for GradientFill (if you use gradient backgrounds):
// #include <msimg32.h>
// #pragma comment(lib, "Msimg32.lib")

namespace fs = std::filesystem;

// ---------- Theming ----------
struct Theme {
    bool    dark = false;
    COLORREF bg;         // main background
    COLORREF panel;      // panels / drop zone
    COLORREF border;     // borders / separators
    COLORREF text;       // normal text
    COLORREF subtext;    // dim text
    COLORREF accent;     // accent
    HBRUSH  hbrBg = nullptr;
    HBRUSH  hbrPanel = nullptr;
};
static Theme gTheme;

// Theme override tri-state
enum class ThemeMode { FollowSystem, ForceDark, ForceLight };
static ThemeMode gThemeMode = ThemeMode::FollowSystem;

// Fullscreen state
static bool  gFullscreen = false;
static RECT  gRestoreRect{};
static DWORD gRestoreStyle = 0, gRestoreExStyle = 0;

// Query system dark mode (AppsUseLightTheme = 0 => dark)
static bool query_system_dark() {
    HKEY hKey;
    DWORD val = 1, sz = sizeof(val);
    if (RegOpenKeyExW(HKEY_CURRENT_USER,
        L"Software\\Microsoft\\Windows\\CurrentVersion\\Themes\\Personalize",
        0, KEY_READ, &hKey) == ERROR_SUCCESS)
    {
        RegQueryValueExW(hKey, L"AppsUseLightTheme", nullptr, nullptr, (LPBYTE)&val, &sz);
        RegCloseKey(hKey);
    }
    return val == 0;
}

static void free_theme_brushes() {
    if (gTheme.hbrBg)    { DeleteObject(gTheme.hbrBg);    gTheme.hbrBg = nullptr; }
    if (gTheme.hbrPanel) { DeleteObject(gTheme.hbrPanel); gTheme.hbrPanel = nullptr; }
}

static void load_theme(bool forceDark = false, bool forceLight = false) {
    free_theme_brushes();
    // New: honor global override first
    if (gThemeMode == ThemeMode::ForceDark)      { forceDark = true; forceLight = false; }
    else if (gThemeMode == ThemeMode::ForceLight){ forceDark = false; forceLight = true; }

    gTheme.dark = forceDark ? true : (forceLight ? false : query_system_dark());

    if (gTheme.dark) {
        gTheme.bg     = RGB(32, 32, 36);
        gTheme.panel  = RGB(45, 45, 50);
        gTheme.border = RGB(70, 70, 75);
        gTheme.text   = RGB(235,235,235);
        gTheme.subtext= RGB(170,170,170);
        gTheme.accent = RGB(0, 120, 215); // Windows blue
    } else {
        gTheme.bg     = RGB(250,250,250);
        gTheme.panel  = RGB(244,244,244);
        gTheme.border = RGB(200,200,200);
        gTheme.text   = RGB(25,25,25);
        gTheme.subtext= RGB(100,100,100);
        gTheme.accent = RGB(0, 120, 215);
    }
    gTheme.hbrBg    = CreateSolidBrush(gTheme.bg);
    gTheme.hbrPanel = CreateSolidBrush(gTheme.panel);
}

// Try enable Mica/Tabbed (Windows 11); safe no‑op on earlier versions
static void apply_backdrop(HWND hwnd) {
#ifndef DWMWA_SYSTEMBACKDROP_TYPE
#define DWMWA_SYSTEMBACKDROP_TYPE 38
#endif
#ifndef DWMSBT_MAINWINDOW
#define DWMSBT_MAINWINDOW 2
#define DWMSBT_TABBEDWINDOW 3
#endif
#ifndef DWMWA_USE_IMMERSIVE_DARK_MODE
#define DWMWA_USE_IMMERSIVE_DARK_MODE 20
#endif
    BOOL useDark = gTheme.dark ? TRUE : FALSE;
    DwmSetWindowAttribute(hwnd, DWMWA_USE_IMMERSIVE_DARK_MODE, &useDark, sizeof(useDark));
    int backdrop = DWMSBT_MAINWINDOW; // subtle Mica
    DwmSetWindowAttribute(hwnd, DWMWA_SYSTEMBACKDROP_TYPE, &backdrop, sizeof(backdrop));
}

// Opt‑in per‑monitor DPI (sharp text)
static void enable_per_monitor_dpi() {
    // On older SDKs this may be unavailable; ignore failure
    SetProcessDpiAwarenessContext(DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE_V2);
}

// ---------------------- small helpers ----------------------
static std::string narrow(const std::wstring& ws) {
    if (ws.empty()) return {};
    int len = ::WideCharToMultiByte(CP_UTF8, 0, ws.c_str(), (int)ws.size(), nullptr, 0, nullptr, nullptr);
    std::string out(len, '\0');
    ::WideCharToMultiByte(CP_UTF8, 0, ws.c_str(), (int)ws.size(), out.data(), len, nullptr, nullptr);
    return out;
}
static std::wstring widen(const std::string& s) {
    if (s.empty()) return {};
    int len = ::MultiByteToWideChar(CP_UTF8, 0, s.c_str(), (int)s.size(), nullptr, 0);
    std::wstring out(len, L'\0');
    ::MultiByteToWideChar(CP_UTF8, 0, s.c_str(), (int)s.size(), out.data(), len);
    return out;
}
static std::string exe_dir() {
    wchar_t buf[MAX_PATH]; ::GetModuleFileNameW(nullptr, buf, MAX_PATH);
    fs::path p(buf); return p.parent_path().string();
}
static void ensure_dir(const std::string& d) {
    std::error_code ec; fs::create_directories(fs::path(d), ec);
}

// ---------------------- classification ----------------------
enum class LogKind { PhotoMesh, RealityMesh, Unknown };

static LogKind detect_log_type(const std::string& path) {
    std::ifstream in(path);
    if (!in) return LogKind::Unknown;
    std::string line; size_t limit = 4000;
    while (limit-- && std::getline(in, line)) {
        // PhotoMesh markers
        if (line.find("SLDEFAULT=>") != std::string::npos ||
            line.find("\"MachineName\"") != std::string::npos ||
            line.find("\"MsgTime\"")    != std::string::npos) {
            return LogKind::PhotoMesh;
        }
        // RealityMesh markers
        if (line.find("Time to run TT project:") != std::string::npos ||
            line.find("-command_file")           != std::string::npos ||
            line.find("Process completed with exit code:") != std::string::npos ||
            line.find("Converted offset:")       != std::string::npos) {
            return LogKind::RealityMesh;
        }
    }
    return LogKind::Unknown;
}

// ---------------------- UI state ----------------------
enum class Mode { MasterOnly, ReportOnly, Both };
struct UIState {
    std::vector<std::string> dropPaths;
    Mode mode = Mode::Both;
    std::string outputsDir = exe_dir() + "/EXCEL OUTPUTS";
};

static UIState g;

// ---------------------- pipeline ----------------------
static void run_pipeline(HWND hwndLog) {
    SendMessage(hwndLog, EM_SETSEL, (WPARAM)-1, (LPARAM)-1);
    auto append = [&](const std::string& s){
        std::wstring w = widen(s + "\r\n");
        SendMessage(hwndLog, EM_REPLACESEL, FALSE, (LPARAM)w.c_str());
    };

    if (g.dropPaths.empty()) { append("[!] No logs to process."); return; }

    std::vector<std::string> pmFiles, rmFiles;
    for (auto& p : g.dropPaths) {
        switch (detect_log_type(p)) {
            case LogKind::PhotoMesh:   pmFiles.push_back(p); break;
            case LogKind::RealityMesh: rmFiles.push_back(p); break;
            default: // unknown -> try both parsers safely
                pmFiles.push_back(p); rmFiles.push_back(p); break;
        }
    }

    append("[*] Parsing logs...");
    std::vector<PhotoMeshRow> pm_rows;
    for (auto &p : pmFiles) pm_rows.push_back(parse_photomesh(p));  // safe to parse; row stays mostly empty if unmatched
    std::vector<RealityMeshRow> rm_rows;
    for (auto &p : rmFiles) rm_rows.push_back(parse_realitymesh(p));

    // Filter out rows that are completely empty (no success, no machine, no duration)
    auto pm_end = std::remove_if(pm_rows.begin(), pm_rows.end(), [](const PhotoMeshRow& r){
        return r.machine.empty() && r.duration.empty() && r.success.empty() && r.exportType.empty();
    });
    pm_rows.erase(pm_end, pm_rows.end());
    auto rm_end = std::remove_if(rm_rows.begin(), rm_rows.end(), [](const RealityMeshRow& r){
        return r.machine.empty() && r.duration.empty() && r.success.empty() && r.exportType.empty();
    });
    rm_rows.erase(rm_end, rm_rows.end());

    ensure_dir(g.outputsDir);

    // Master single-sheet (append + rebuild)
    if (g.mode == Mode::MasterOnly || g.mode == Mode::Both) {
        auto unified = excel::unify(pm_rows, rm_rows);
        excel::append_to_master_and_rebuild_xlsx(g.outputsDir, unified);
        append("[+] Updated master: " + g.outputsDir + "/All_Exports.xlsx");
    }

    // Per-run report (multi-sheet)
    if (g.mode == Mode::ReportOnly || g.mode == Mode::Both) {
        std::vector<SummaryRow> summary;
        for (const auto &r : pm_rows) summary.push_back(make_summary(r));
        for (const auto &r : rm_rows) summary.push_back(make_summary(r));
        const std::string out = g.outputsDir + "/Report.xlsx";
        excel::write_workbook(out, pm_rows, rm_rows, summary);
        append("[+] Wrote per-run report: " + out);
    }

    append("[OK] Done.");
    g.dropPaths.clear();
}

// ---------------------- UI creation ----------------------
#define IDC_RAD_MASTER   1001
#define IDC_RAD_REPORT   1002
#define IDC_RAD_BOTH     1003
#define IDC_BTN_CHOOSE   1004
#define IDC_BTN_RUN      1005
#define IDC_BTN_OPEN     1006
#define IDC_EDIT_LOG     1007
#define IDC_BTN_SETTINGS 1008
#define IDM_SET_THEME_SYS   3001
#define IDM_SET_THEME_DARK  3002
#define IDM_SET_THEME_LIGHT 3003
#define IDM_SET_FULLSCREEN  3004

static HFONT g_hFontTitle = nullptr;
static HFONT g_hFontBase  = nullptr;

static void create_fonts() {
    LOGFONTW lf{}; lf.lfHeight = -18; lf.lfWeight = FW_SEMIBOLD; wcscpy_s(lf.lfFaceName, L"Segoe UI");
    g_hFontTitle = CreateFontIndirectW(&lf);
    lf = {}; lf.lfHeight = -15; wcscpy_s(lf.lfFaceName, L"Segoe UI");
    g_hFontBase = CreateFontIndirectW(&lf);
}

static void paint_dropzone(HDC hdc, RECT rc) {
    // Panel fill
    FillRect(hdc, &rc, gTheme.hbrPanel);

    // Rounded border
    HGDIOBJ oldBrush = SelectObject(hdc, GetStockObject(NULL_BRUSH));
    HPEN pen = CreatePen(PS_SOLID, 2, gTheme.accent);
    HGDIOBJ oldPen = SelectObject(hdc, pen);
    RoundRect(hdc, rc.left, rc.top, rc.right, rc.bottom, 18, 18);
    SelectObject(hdc, oldPen); DeleteObject(pen);
    SelectObject(hdc, oldBrush);

    // Title
    SetBkMode(hdc, TRANSPARENT);
    SetTextColor(hdc, gTheme.text);
    HFONT old = (HFONT)SelectObject(hdc, g_hFontTitle);
    const wchar_t* title = L"\xE8E5  Drop your logs here"; // MDL2 'OpenFile' glyph
    DrawTextW(hdc, title, -1, &rc, DT_CENTER | DT_VCENTER | DT_SINGLELINE);
    SelectObject(hdc, old);
}

static void layout_controls(HWND h,
                            HWND hRadMaster, HWND hRadReport, HWND hRadBoth,
                            HWND hBtnChoose, HWND hBtnRun, HWND hBtnOpen, HWND hBtnSettings,
                            HWND hEditLog, RECT& dropRc) {
    RECT rc; GetClientRect(h, &rc);

    const int MARGIN  = 48;
    const int GUTTER  = 16;
    const int lineH   = 34;
    const int btnH    = lineH + 4;

    // Absolute minimums so nothing gets cut off
    const int MIN_DROP_H    = 160;
    const int MIN_CONSOLE_H = 160;

    // Radios in a row
    const int headerY = 20;
    MoveWindow(hRadMaster, MARGIN,                 headerY, 200, lineH, TRUE);
    MoveWindow(hRadReport, MARGIN + 200 + GUTTER,  headerY, 160, lineH, TRUE);
    MoveWindow(hRadBoth,   MARGIN + 200 + GUTTER + 160 + GUTTER, headerY, 180, lineH, TRUE);

    // Top of drop zone
    int top = 84;

    // Available vertical pixels below 'top' down to bottom margin
    int avail = rc.bottom - top - MARGIN;
    if (avail < 0) avail = 0;

    // We need: dropH + 24 + btnH + 24 + consoleH  <= avail
    // Start with drop zone as 45% of available, then clamp so console can have MIN_CONSOLE_H
    int dropPreferred = (avail * 45) / 100;
    int minRequiredBelowDrop = 24 + btnH + 24 + MIN_CONSOLE_H;
    int dropH = dropPreferred;

    // clamp dropH to [MIN_DROP_H, avail - minRequiredBelowDrop]
    if (dropH < MIN_DROP_H) dropH = MIN_DROP_H;
    if (dropH > avail - minRequiredBelowDrop) dropH = avail - minRequiredBelowDrop;
    if (dropH < MIN_DROP_H) dropH = std::max(80, avail - minRequiredBelowDrop); // last resort

    // Lay out drop zone
    dropRc.left   = MARGIN;
    dropRc.top    = top;
    dropRc.right  = rc.right - MARGIN;
    dropRc.bottom = top + std::max(0, dropH);

    // Buttons row
    int by = dropRc.bottom + 24;
    int x  = MARGIN;

    int wChoose = 220;
    int wRun    = 120;
    int wOpen   = 260;
    int wSett   = 140;

    MoveWindow(hBtnChoose,   x,          by, wChoose, btnH, TRUE);
    x += wChoose + GUTTER;
    MoveWindow(hBtnRun,      x,          by, wRun,    btnH, TRUE);
    x += wRun + GUTTER;
    MoveWindow(hBtnOpen,     x,          by, wOpen,   btnH, TRUE);
    // Settings anchored to right margin
    MoveWindow(hBtnSettings, rc.right - MARGIN - wSett, by, wSett, btnH, TRUE);

    // Console fills the rest
    int logTop = by + 24;
    int logH   = rc.bottom - logTop - MARGIN;

    // If console would be too small, steal space from drop zone and recompute
    if (logH < MIN_CONSOLE_H) {
        int need = MIN_CONSOLE_H - logH;
        // Shrink drop zone, but not below MIN_DROP_H
        int currentDropH = dropRc.bottom - dropRc.top;
        int newDropH = std::max(MIN_DROP_H, currentDropH - need);

        dropRc.bottom = dropRc.top + newDropH;

        // Recompute positions with the updated dropRc
        by = dropRc.bottom + 24;
        x  = MARGIN;
        MoveWindow(hBtnChoose,   x,          by, wChoose, btnH, TRUE);
        x += wChoose + GUTTER;
        MoveWindow(hBtnRun,      x,          by, wRun,    btnH, TRUE);
        x += wRun + GUTTER;
        MoveWindow(hBtnOpen,     x,          by, wOpen,   btnH, TRUE);
        MoveWindow(hBtnSettings, rc.right - MARGIN - wSett, by, wSett, btnH, TRUE);

        logTop = by + 24;
        logH   = std::max(MIN_CONSOLE_H, rc.bottom - logTop - MARGIN);
    }

    MoveWindow(hEditLog, MARGIN, logTop, rc.right - 2 * MARGIN, std::max(0, logH), TRUE);
}

static void apply_fullscreen(HWND h, bool enable) {
    if (enable == gFullscreen) return;
    gFullscreen = enable;

    if (enable) {
        // Save current placement/style
        gRestoreStyle   = GetWindowLong(h, GWL_STYLE);
        gRestoreExStyle = GetWindowLong(h, GWL_EXSTYLE);
        GetWindowRect(h, &gRestoreRect);

        // Switch to borderless fullscreen on current monitor
        MONITORINFO mi{ sizeof(mi) };
        GetMonitorInfo(MonitorFromWindow(h, MONITOR_DEFAULTTONEAREST), &mi);

        SetWindowLong(h, GWL_STYLE, (gRestoreStyle & ~(WS_OVERLAPPED | WS_CAPTION |
                         WS_SYSMENU | WS_MINIMIZEBOX | WS_MAXIMIZEBOX | WS_THICKFRAME)) | WS_POPUP);
        SetWindowLong(h, GWL_EXSTYLE, gRestoreExStyle);
        SetWindowPos(h, HWND_TOP, mi.rcMonitor.left, mi.rcMonitor.top,
                     mi.rcMonitor.right - mi.rcMonitor.left,
                     mi.rcMonitor.bottom - mi.rcMonitor.top,
                     SWP_FRAMECHANGED | SWP_SHOWWINDOW);
    } else {
        // Restore windowed style/size
        SetWindowLong(h, GWL_STYLE, gRestoreStyle);
        SetWindowLong(h, GWL_EXSTYLE, gRestoreExStyle);
        SetWindowPos(h, nullptr, gRestoreRect.left, gRestoreRect.top,
                     gRestoreRect.right - gRestoreRect.left,
                     gRestoreRect.bottom - gRestoreRect.top,
                     SWP_FRAMECHANGED | SWP_NOZORDER | SWP_SHOWWINDOW);
    }

    // Re-apply backdrop + redraw
    apply_backdrop(h);
    InvalidateRect(h, nullptr, TRUE);
}

static HMENU build_settings_menu() {
    HMENU m = CreatePopupMenu();
    AppendMenuW(m, MF_STRING | (gThemeMode == ThemeMode::FollowSystem ? MF_CHECKED : 0),
                IDM_SET_THEME_SYS, L"Follow system");
    AppendMenuW(m, MF_STRING | (gThemeMode == ThemeMode::ForceDark    ? MF_CHECKED : 0),
                IDM_SET_THEME_DARK, L"Dark");
    AppendMenuW(m, MF_STRING | (gThemeMode == ThemeMode::ForceLight   ? MF_CHECKED : 0),
                IDM_SET_THEME_LIGHT, L"Light");
    AppendMenuW(m, MF_SEPARATOR, 0, nullptr);
    AppendMenuW(m, MF_STRING | (gFullscreen ? MF_CHECKED : 0),
                IDM_SET_FULLSCREEN, L"Fullscreen (F11)");
    return m;
}

// ---------------------- Win32 Window Proc ----------------------
static LRESULT CALLBACK WndProc(HWND h, UINT msg, WPARAM w, LPARAM l) {
    static HWND hRadMaster, hRadReport, hRadBoth, hBtnChoose, hBtnRun, hBtnOpen, hBtnSettings, hEditLog;
    static RECT dropRc{40, 80, 760, 280};

    switch (msg) {
        case WM_CTLCOLORDLG:
        case WM_ERASEBKGND:
            return (LRESULT)gTheme.hbrBg;

        case WM_CTLCOLORSTATIC:
        case WM_CTLCOLOREDIT:
        case WM_CTLCOLORBTN: {
            HDC hdc = (HDC)w;
            SetTextColor(hdc, gTheme.text);
            SetBkColor(hdc, gTheme.bg);
            return (LRESULT)gTheme.hbrBg;
        }

        case WM_SETTINGCHANGE:
            if (gThemeMode == ThemeMode::FollowSystem &&
                l && !lstrcmpiW((LPCWSTR)l, L"ImmersiveColorSet")) {
                load_theme(); apply_backdrop(h);
                SetClassLongPtr(h, GCLP_HBRBACKGROUND, (LONG_PTR)gTheme.hbrBg);
                InvalidateRect(h, nullptr, TRUE);
            }
            return 0;

        case WM_CREATE: {
            create_fonts();
            // radios
            hRadMaster = CreateWindowW(L"BUTTON", L"Append to Master",
                WS_CHILD|WS_VISIBLE|BS_AUTORADIOBUTTON, 40, 20, 180, 24, h, (HMENU)IDC_RAD_MASTER, nullptr, nullptr);
            hRadReport = CreateWindowW(L"BUTTON", L"Run Report",
                WS_CHILD|WS_VISIBLE|BS_AUTORADIOBUTTON, 230, 20, 140, 24, h, (HMENU)IDC_RAD_REPORT, nullptr, nullptr);
            hRadBoth   = CreateWindowW(L"BUTTON", L"Both (Default)",
                WS_CHILD|WS_VISIBLE|BS_AUTORADIOBUTTON, 380, 20, 160, 24, h, (HMENU)IDC_RAD_BOTH, nullptr, nullptr);
            SendMessage(hRadBoth, BM_SETCHECK, BST_CHECKED, 0);

            // buttons
            hBtnChoose = CreateWindowW(L"BUTTON", L"Choose Logs…", WS_CHILD|WS_VISIBLE, 40, 320, 140, 28, h, (HMENU)IDC_BTN_CHOOSE, nullptr, nullptr);
            hBtnRun    = CreateWindowW(L"BUTTON", L"Run",          WS_CHILD|WS_VISIBLE | BS_DEFPUSHBUTTON, 190,320, 80, 28, h, (HMENU)IDC_BTN_RUN, nullptr, nullptr);
            hBtnOpen   = CreateWindowW(L"BUTTON", L"Open Output Folder", WS_CHILD|WS_VISIBLE, 280, 320, 200, 28, h, (HMENU)IDC_BTN_OPEN, nullptr, nullptr);
            hBtnSettings = CreateWindowW(L"BUTTON", L"Settings…", WS_CHILD|WS_VISIBLE,
                                        520, 320, 120, 28, h, (HMENU)IDC_BTN_SETTINGS, nullptr, nullptr);

            // status log (read-only, multiline)
            hEditLog = CreateWindowW(L"EDIT", L"", WS_CHILD|WS_VISIBLE|WS_VSCROLL|ES_MULTILINE|ES_READONLY,
                                     40, 360, 760, 180, h, (HMENU)IDC_EDIT_LOG, nullptr, nullptr);

            SendMessage(hRadMaster, WM_SETFONT, (WPARAM)g_hFontBase, TRUE);
            SendMessage(hRadReport, WM_SETFONT, (WPARAM)g_hFontBase, TRUE);
            SendMessage(hRadBoth,   WM_SETFONT, (WPARAM)g_hFontBase, TRUE);
            SendMessage(hBtnChoose, WM_SETFONT, (WPARAM)g_hFontBase, TRUE);
            SendMessage(hBtnRun,    WM_SETFONT, (WPARAM)g_hFontBase, TRUE);
            SendMessage(hBtnOpen,     WM_SETFONT, (WPARAM)g_hFontBase, TRUE);
            SendMessage(hBtnSettings, WM_SETFONT, (WPARAM)g_hFontBase, TRUE);
            SendMessage(hEditLog,     WM_SETFONT, (WPARAM)g_hFontBase, TRUE);

            // Let common controls adopt Explorer theme
            SetWindowTheme(hRadMaster, L"Explorer", nullptr);
            SetWindowTheme(hRadReport, L"Explorer", nullptr);
            SetWindowTheme(hRadBoth,   L"Explorer", nullptr);
            SetWindowTheme(hBtnChoose, L"Explorer", nullptr);
            SetWindowTheme(hBtnRun,    L"Explorer", nullptr);
            SetWindowTheme(hBtnOpen,     L"Explorer", nullptr);
            SetWindowTheme(hBtnSettings, L"Explorer", nullptr);
            SetWindowTheme(hEditLog,     L"Explorer", nullptr);

            DragAcceptFiles(h, TRUE);

            layout_controls(h, hRadMaster,hRadReport,hRadBoth,
                            hBtnChoose,hBtnRun,hBtnOpen,hBtnSettings,
                            hEditLog, dropRc);
            return 0;
        }
        case WM_SIZE:
            layout_controls(h, hRadMaster,hRadReport,hRadBoth,
                            hBtnChoose,hBtnRun,hBtnOpen,hBtnSettings,
                            hEditLog, dropRc);
            return 0;
        case WM_DPICHANGED: {
            const RECT* pr = reinterpret_cast<const RECT*>(l);
            SetWindowPos(h, nullptr, pr->left, pr->top,
                         pr->right - pr->left, pr->bottom - pr->top,
                         SWP_NOZORDER | SWP_NOACTIVATE);
            return 0;
        }
        case WM_PAINT: {
            PAINTSTRUCT ps; HDC hdc = BeginPaint(h, &ps);

            // Header separator
            RECT client; GetClientRect(h, &client);
            HPEN sep = CreatePen(PS_SOLID, 1, gTheme.border);
            HGDIOBJ oldPen = SelectObject(hdc, sep);
            MoveToEx(hdc, 40, 56, nullptr); LineTo(hdc, client.right-40, 56);
            SelectObject(hdc, oldPen); DeleteObject(sep);

            // Drop zone
            paint_dropzone(hdc, dropRc);

            EndPaint(h, &ps);
            return 0;
        }
        case WM_DROPFILES: {
            HDROP hd = (HDROP)w;
            UINT count = DragQueryFileW(hd, 0xFFFFFFFF, nullptr, 0);
            for (UINT i=0;i<count;++i) {
                wchar_t path[MAX_PATH]; DragQueryFileW(hd, i, path, MAX_PATH);
                g.dropPaths.push_back(narrow(path));
            }
            DragFinish(hd);
            // Auto-run on drop
            run_pipeline(GetDlgItem(h, IDC_EDIT_LOG));
            InvalidateRect(h, &dropRc, FALSE);
            return 0;
        }
        case WM_COMMAND: {
            switch (LOWORD(w)) {
                case IDC_RAD_MASTER: g.mode = Mode::MasterOnly; break;
                case IDC_RAD_REPORT: g.mode = Mode::ReportOnly; break;
                case IDC_RAD_BOTH:   g.mode = Mode::Both;       break;
                case IDC_BTN_CHOOSE: {
                    // Use standard picker
                    IFileOpenDialog* dlg = nullptr;
                    if (SUCCEEDED(CoCreateInstance(CLSID_FileOpenDialog, nullptr, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&dlg)))) {
                        DWORD opts=0; dlg->GetOptions(&opts); dlg->SetOptions(opts | FOS_ALLOWMULTISELECT);
                        COMDLG_FILTERSPEC filters[] = { { L"Log files (*.log;*.txt)", L"*.log;*.txt" }, { L"All files (*.*)", L"*.*" } };
                        dlg->SetFileTypes(2, filters);
                        if (SUCCEEDED(dlg->Show(nullptr))) {
                            IShellItemArray* items = nullptr;
                            if (SUCCEEDED(dlg->GetResults(&items))) {
                                DWORD n=0; items->GetCount(&n);
                                for (DWORD i=0;i<n;++i) {
                                    IShellItem* it=nullptr; if (SUCCEEDED(items->GetItemAt(i, &it))) {
                                        PWSTR psz=nullptr; if (SUCCEEDED(it->GetDisplayName(SIGDN_FILESYSPATH, &psz)) && psz) {
                                            g.dropPaths.push_back(narrow(psz)); CoTaskMemFree(psz);
                                        }
                                        it->Release();
                                    }
                                }
                                items->Release();
                            }
                        }
                        dlg->Release();
                    }
                    break;
                }
                case IDC_BTN_RUN: {
                    run_pipeline(GetDlgItem(h, IDC_EDIT_LOG));
                    break;
                }
                case IDC_BTN_OPEN: {
                    ensure_dir(g.outputsDir);
                    ::ShellExecuteW(nullptr, L"open", widen(g.outputsDir).c_str(), nullptr, nullptr, SW_SHOWNORMAL);
                    break;
                }
                case IDC_BTN_SETTINGS: {
                    HMENU m = build_settings_menu();
                    RECT r; GetWindowRect(hBtnSettings, &r);
                    TrackPopupMenu(m, TPM_LEFTALIGN | TPM_TOPALIGN,
                                   r.left, r.bottom, 0, h, nullptr);
                    DestroyMenu(m);
                    return 0;
                }
                case IDM_SET_THEME_SYS:
                    gThemeMode = ThemeMode::FollowSystem;
                    load_theme(); apply_backdrop(h);
                    InvalidateRect(h, nullptr, TRUE);
                    return 0;
                case IDM_SET_THEME_DARK:
                    gThemeMode = ThemeMode::ForceDark;
                    load_theme(true, false); apply_backdrop(h);
                    InvalidateRect(h, nullptr, TRUE);
                    return 0;
                case IDM_SET_THEME_LIGHT:
                    gThemeMode = ThemeMode::ForceLight;
                    load_theme(false, true); apply_backdrop(h);
                    InvalidateRect(h, nullptr, TRUE);
                    return 0;
                case IDM_SET_FULLSCREEN:
                    apply_fullscreen(h, !gFullscreen);
                    return 0;
            }
            return 0;
        }
        case WM_KEYDOWN:
            if (w == VK_F11) { apply_fullscreen(h, !gFullscreen); return 0; }
            break;
        case WM_DESTROY: {
            if (g_hFontTitle) DeleteObject(g_hFontTitle);
            if (g_hFontBase)  DeleteObject(g_hFontBase);
            free_theme_brushes();
            PostQuitMessage(0);
            return 0;
        }
    }
    return DefWindowProc(h, msg, w, l);
}

int APIENTRY wWinMain(HINSTANCE hInst, HINSTANCE, LPWSTR, int nCmd) {
    enable_per_monitor_dpi();
    load_theme(); // picks current system mode

    CoInitializeEx(nullptr, COINIT_APARTMENTTHREADED | COINIT_DISABLE_OLE1DDE);
    INITCOMMONCONTROLSEX icc{sizeof(icc), ICC_STANDARD_CLASSES}; InitCommonControlsEx(&icc);

    const wchar_t* cls = L"LogToExcel.GUI";
    WNDCLASSEXW wc{sizeof(wc)}; wc.hInstance = hInst; wc.lpszClassName = cls;
    wc.hCursor = LoadCursor(nullptr, IDC_ARROW);
    wc.hbrBackground = gTheme.hbrBg;  // theme background
    wc.lpfnWndProc = WndProc;
    RegisterClassExW(&wc);

    // Pick a size that fits the current work area (taskbar excluded), then center
    RECT wa; SystemParametersInfoW(SPI_GETWORKAREA, 0, &wa, 0);
    int workW = wa.right  - wa.left;
    int workH = wa.bottom - wa.top;

    // Target ~80% of work area, but don't go below a usable minimum
    int targetW = std::max(1100, (workW * 80) / 100);
    int targetH = std::max( 720, (workH * 75) / 100);

    RECT wr{0,0,targetW,targetH};
    AdjustWindowRectEx(&wr, WS_OVERLAPPED | WS_CAPTION | WS_SYSMENU | WS_MINIMIZEBOX, FALSE, 0);
    int winW = wr.right - wr.left;
    int winH = wr.bottom - wr.top;

    // Center within work area
    int posX = wa.left + (workW - winW) / 2;
    int posY = wa.top  + (workH - winH) / 2;

    HWND hWnd = CreateWindowExW(0, cls, L"Log to Excel",
                                WS_OVERLAPPED | WS_CAPTION | WS_SYSMENU | WS_MINIMIZEBOX,
                                posX, posY, winW, winH,
                                nullptr, nullptr, hInst, nullptr);
    apply_backdrop(hWnd);
    ShowWindow(hWnd, nCmd); UpdateWindow(hWnd);

    MSG msg;
    while (GetMessage(&msg, nullptr, 0, 0) > 0) {
        TranslateMessage(&msg); DispatchMessage(&msg);
    }
    CoUninitialize();
    return 0;
}
