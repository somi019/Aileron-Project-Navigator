#include "framework.h"
#include "ProjectNavigator.h"
#include "Resource.h"
#include <shobjidl.h> 
#include <shellapi.h>
#include <string>
#include <map>
#include <vector>
#include <fstream>
#include <shlwapi.h>

#define MAX_LOADSTRING 100
#define ID_BTN_ROOT      1001
#define ID_BTN_ADD_PATH  1002
#define IDM_FILE_NEW     2000
#define IDM_FILE_LOAD    2001
#define IDM_FILE_SAVE    2002
#define IDM_FILE_SAVE_AS 2004
#define IDM_SET_ROOT     2003
#define IDM_REMOVE_BTN   2005

// Global Variables
HINSTANCE hInst;
HWND g_hMainWnd = NULL; // main window handle
HMENU g_hMenu = NULL; // main menu handle
HWND g_hRootBtn = NULL; // track Root button so it can be destroyed
HFONT g_hMatrixFont = NULL; // Consolas-like font for UI
HBRUSH g_hBgBrush = NULL; // dark background brush
COLORREF g_lightGreen = RGB(180, 255, 180);
COLORREF g_grey = RGB(180, 180, 180);
bool g_removeMode = false; // when true, next button click removes that button
WCHAR szTitle[MAX_LOADSTRING] = L"Project Navigator";
WCHAR szWindowClass[MAX_LOADSTRING] = L"PROJECT_NAV_CLASS";
std::wstring g_configFilePath = L""; // path to currently bound config file

std::wstring g_rootPath = L"";
std::map<HWND, std::wstring> g_buttonMap; // Maps button handle to its folder path
struct PathData { std::wstring name; std::wstring path; };
std::vector<PathData> g_customData; // Helper for saving/loading

// Forward Declarations
ATOM MyRegisterClass(HINSTANCE);
BOOL InitInstance(HINSTANCE, int);
LRESULT CALLBACK WndProc(HWND, UINT, WPARAM, LPARAM);
LRESULT CALLBACK AddPathDlgProc(HWND, UINT, WPARAM, LPARAM);
void PickFolder(HWND, std::wstring&, bool allowFiles = false);
void RefreshUI(HWND);
void CenterWindow(HWND hwnd);
void UpdateMenuState();

int APIENTRY wWinMain(HINSTANCE hInstance, HINSTANCE hPrev, LPWSTR lpCmd, int nCmdShow) {
    CoInitializeEx(NULL, COINIT_APARTMENTTHREADED);
    MyRegisterClass(hInstance);
    if (!InitInstance(hInstance, nCmdShow)) return FALSE;
    MSG msg;
    while (GetMessage(&msg, nullptr, 0, 0)) {
        TranslateMessage(&msg);
        DispatchMessage(&msg);
    }
    CoUninitialize();
    return (int)msg.wParam;
}

ATOM MyRegisterClass(HINSTANCE hInstance) {
    WNDCLASSEXW wcex = { sizeof(WNDCLASSEX) };
    wcex.style = CS_HREDRAW | CS_VREDRAW;
    wcex.lpfnWndProc = WndProc;
    wcex.hInstance = hInstance;
    wcex.hCursor = LoadCursor(nullptr, IDC_ARROW);
    wcex.hbrBackground = (HBRUSH)(COLOR_WINDOW + 1);
    wcex.lpszClassName = szWindowClass;
    wcex.hIcon = LoadIconW(hInstance, MAKEINTRESOURCE(IDI_ICON1));
    wcex.hIconSm = LoadIconW(hInstance, MAKEINTRESOURCE(IDI_SMALL));
    return RegisterClassExW(&wcex);
}

BOOL InitInstance(HINSTANCE hInstance, int nCmdShow) {
    hInst = hInstance;
    int screenW = GetSystemMetrics(SM_CXSCREEN);
    HWND hWnd = CreateWindowW(szWindowClass, szTitle, WS_OVERLAPPEDWINDOW,
        screenW - 600, 50, 550, 150, nullptr, nullptr, hInstance, nullptr);

    if (!hWnd) return FALSE;

    g_hMainWnd = hWnd; // store handle for dialogs to post to

    // Ensure the small icon appears in the top-left corner of the titlebar.
    // Also set the big icon (used by taskbar) to match.
    SendMessageW(hWnd, WM_SETICON, ICON_BIG, (LPARAM)LoadIconW(hInstance, MAKEINTRESOURCE(IDI_ICON1)));
    SendMessageW(hWnd, WM_SETICON, ICON_SMALL, (LPARAM)LoadIconW(hInstance, MAKEINTRESOURCE(IDI_SMALL)));

    // Create Matrix-style font and background brush
    g_hMatrixFont = CreateFontW(
        16, 0, 0, 0, FW_NORMAL,
        FALSE, FALSE, FALSE, DEFAULT_CHARSET,
        OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS, CLEARTYPE_QUALITY,
        FIXED_PITCH | FF_DONTCARE,
        L"Consolas");
    // Use light grey background for overall app
    g_hBgBrush = CreateSolidBrush(g_grey);

    HMENU hMenu = CreateMenu();
    HMENU hFile = CreatePopupMenu();
    AppendMenu(hFile, MF_STRING, IDM_FILE_NEW, L"Create New");
    AppendMenu(hFile, MF_SEPARATOR, 0, NULL);
    AppendMenu(hFile, MF_STRING, IDM_FILE_LOAD, L"Load");
    AppendMenu(hFile, MF_STRING, IDM_FILE_SAVE_AS, L"Save As...");
    AppendMenu(hFile, MF_STRING, IDM_FILE_SAVE, L"Save");
    AppendMenu(hMenu, MF_POPUP, (UINT_PTR)hFile, L"File");

    HMENU hSettings = CreatePopupMenu();
    AppendMenu(hSettings, MF_STRING, IDM_SET_ROOT, L"Set Root Folder");
    AppendMenu(hMenu, MF_POPUP, (UINT_PTR)hSettings, L"Settings");

    // New top-level Add menu with Add path item (replaces on-window '+' button)
    HMENU hAdd = CreatePopupMenu();
    AppendMenu(hAdd, MF_STRING, ID_BTN_ADD_PATH, L"Add File/Folder");
    AppendMenu(hMenu, MF_POPUP, (UINT_PTR)hAdd, L"Add");
    
    // Top-level Remove menu with single action
    HMENU hRemove = CreatePopupMenu();
    AppendMenu(hRemove, MF_STRING, IDM_REMOVE_BTN, L"Remove Button");
    AppendMenu(hMenu, MF_POPUP, (UINT_PTR)hRemove, L"Remove");

    SetMenu(hWnd, hMenu);
    g_hMenu = hMenu;

    ShowWindow(hWnd, nCmdShow);
    UpdateWindow(hWnd);

    UpdateMenuState();
    return TRUE;
}

void UpdateMenuState() {
    if (!g_hMenu) return;
    BOOL canSave = !g_configFilePath.empty();
    EnableMenuItem(g_hMenu, IDM_FILE_SAVE, MF_BYCOMMAND | (canSave ? MF_ENABLED : MF_GRAYED));

    // Enable Remove when there are any custom buttons or root exists
    BOOL hasAny = !g_customData.empty() || !g_rootPath.empty();
    EnableMenuItem(g_hMenu, IDM_REMOVE_BTN, MF_BYCOMMAND | (hasAny ? MF_ENABLED : MF_GRAYED));
    if (g_hMainWnd) DrawMenuBar(g_hMainWnd);
}

void RefreshUI(HWND hWnd) {
    // Destroy existing Root button if present
    if (g_hRootBtn) {
        DestroyWindow(g_hRootBtn);
        g_hRootBtn = NULL;
    }

    // 1. Clear existing dynamic buttons from the map and the screen
    for (auto const& [hBtn, path] : g_buttonMap) {
        DestroyWindow(hBtn);
    }
    g_buttonMap.clear();

    // Get client area to layout buttons and potentially resize window vertically
    RECT rcClient;
    GetClientRect(hWnd, &rcClient);
    int clientW = rcClient.right - rcClient.left;
    int clientH = rcClient.bottom - rcClient.top;

    int marginLeft = 10;
    int x = marginLeft;
    int y = 20;
    int btnW = 100;
    int btnH = 35;
    int hSpacing = 10;
    int vSpacing = 10;

    // 2. Create Root Button only if root path is set
    if (!g_rootPath.empty()) {
        // If root button doesn't fit horizontally, move to next row
        if (x + btnW > clientW) {
            x = marginLeft;
            y += btnH + vSpacing;
        }

        HWND hRoot = CreateWindowW(L"BUTTON", L"Root", WS_VISIBLE | WS_CHILD | BS_PUSHBUTTON,
            x, y, btnW, btnH, hWnd, (HMENU)ID_BTN_ROOT, hInst, NULL);
        g_hRootBtn = hRoot;
        if (g_hMatrixFont) SendMessageW(hRoot, WM_SETFONT, (WPARAM)g_hMatrixFont, TRUE);
        x += btnW + hSpacing;
    }

    // 3. Create buttons for custom paths stored in g_customData
    for (const auto& item : g_customData) {
        // If next button would overflow horizontally, wrap to next row
        if (x + btnW > clientW) {
            x = marginLeft;
            y += btnH + vSpacing;

            // If we've grown past the client height, resize the window to accommodate new row
            if (y + btnH > clientH) {
                RECT rcWindow;
                GetWindowRect(hWnd, &rcWindow);
                int winW = rcWindow.right - rcWindow.left;
                int winH = rcWindow.bottom - rcWindow.top;

                // Increase window height by the required amount (use client coordinates delta)
                int extraNeeded = (y + btnH) - clientH + vSpacing;
                int newWinH = winH + extraNeeded;

                // Resize window (keep position)
                SetWindowPos(hWnd, NULL, 0, 0, winW, newWinH, SWP_NOMOVE | SWP_NOZORDER);

                // Update client size for further layout
                GetClientRect(hWnd, &rcClient);
                clientH = rcClient.bottom - rcClient.top;
            }
        }

        HWND hNewBtn = CreateWindowW(L"BUTTON", item.name.c_str(), WS_VISIBLE | WS_CHILD | BS_PUSHBUTTON,
            x, y, btnW, btnH, hWnd, NULL, hInst, NULL);
        if (g_hMatrixFont) SendMessageW(hNewBtn, WM_SETFONT, (WPARAM)g_hMatrixFont, TRUE);
        g_buttonMap[hNewBtn] = item.path; // Map the handle to the path
        x += btnW + hSpacing;
    }
}

LRESULT CALLBACK WndProc(HWND hWnd, UINT message, WPARAM wParam, LPARAM lParam) {
    switch (message) {
    case WM_CREATE:
        RefreshUI(hWnd);
        break;

    case WM_COMMAND: {
        HWND hCtrl = (HWND)lParam; // Handle of the control that sent the message

        // Handle the static buttons by ID
        int id = LOWORD(wParam);
        if (id == ID_BTN_ROOT) {
            if (g_removeMode) {
                // remove root
                g_rootPath.clear();
                g_removeMode = false;
                UpdateMenuState();
                RefreshUI(hWnd);
            } else {
                if (g_rootPath.empty()) PickFolder(hWnd, g_rootPath, false);
                else ShellExecuteW(NULL, L"open", g_rootPath.c_str(), NULL, NULL, SW_SHOWDEFAULT);
            }
        }
        else if (id == IDM_FILE_NEW) {
            // Clear current project state and refresh UI
            g_rootPath.clear();
            g_customData.clear();
            g_configFilePath.clear();
            UpdateMenuState();
            RefreshUI(hWnd);
        }
        else if (id == ID_BTN_ADD_PATH) {
            // Same behavior as before: open add-path dialog when menu item selected
            WNDCLASSW wc = { 0 };
            wc.lpfnWndProc = AddPathDlgProc;
            wc.hInstance = hInst;
            wc.hbrBackground = (HBRUSH)(COLOR_BTNFACE + 1);
            wc.lpszClassName = L"AddPathForm";

            WNDCLASSW existing = { 0 };
            if (!GetClassInfoW(hInst, wc.lpszClassName, &existing)) {
                RegisterClassW(&wc);
            }

            HWND hDlg = CreateWindowW(L"AddPathForm", L"Add Custom Path", WS_OVERLAPPED | WS_CAPTION | WS_SYSMENU,
                CW_USEDEFAULT, CW_USEDEFAULT, 350, 180, NULL, NULL, hInst, NULL);

            if (hDlg) {
                CenterWindow(hDlg);
                ShowWindow(hDlg, SW_SHOW);
                UpdateWindow(hDlg);
            }
        }
        else if (id == IDM_REMOVE_BTN) {
            // Ask user to confirm removal mode; allow Cancel to abort
            int rv = MessageBox(hWnd, L"Click a button to remove it.\nPress Cancel to abort.", L"Remove Button", MB_OKCANCEL | MB_ICONINFORMATION);
            if (rv == IDOK) {
                g_removeMode = true;
            } else {
                g_removeMode = false;
            }
        }
        else if (id == IDM_SET_ROOT) {
            // Open the AddPathForm in "root mode" (lpCreateParams == (void*)1)
            WNDCLASSW wcRoot = { 0 };
            wcRoot.lpfnWndProc = AddPathDlgProc;
            wcRoot.hInstance = hInst;
            wcRoot.hbrBackground = (HBRUSH)(COLOR_BTNFACE + 1);
            wcRoot.lpszClassName = L"AddPathForm";
            WNDCLASSW existingRoot = { 0 };
            if (!GetClassInfoW(hInst, wcRoot.lpszClassName, &existingRoot)) {
                RegisterClassW(&wcRoot);
            }

            HWND hRootDlg = CreateWindowW(L"AddPathForm", L"Set Root Folder", WS_OVERLAPPED | WS_CAPTION | WS_SYSMENU,
                CW_USEDEFAULT, CW_USEDEFAULT, 420, 160, NULL, NULL, hInst, (LPVOID)1);

            if (hRootDlg) {
                CenterWindow(hRootDlg);
                ShowWindow(hRootDlg, SW_SHOW);
                UpdateWindow(hRootDlg);
            }
        }
        else if (id == IDM_FILE_SAVE_AS) {
            // Ensure configs folder exists
            CreateDirectoryW(L"configs", NULL);

            IFileSaveDialog* pSave;
            if (SUCCEEDED(CoCreateInstance(CLSID_FileSaveDialog, NULL, CLSCTX_ALL, IID_PPV_ARGS(&pSave)))) {
                // filter for .config
                COMDLG_FILTERSPEC spec = { L"Config Files (*.config)", L"*.config" };
                pSave->SetFileTypes(1, &spec);
                pSave->SetDefaultExtension(L"config");

                // set default folder to configs (relative)
                PWSTR pwszFolder = NULL;
                WCHAR fullPath[MAX_PATH];
                if (GetFullPathNameW(L"configs", MAX_PATH, fullPath, NULL)) {
                    IShellItem* psiFolder = NULL;
                    if (SUCCEEDED(SHCreateItemFromParsingName(fullPath, NULL, IID_PPV_ARGS(&psiFolder)))) {
                        pSave->SetDefaultFolder(psiFolder);
                        psiFolder->Release();
                    }
                }

                if (SUCCEEDED(pSave->Show(hWnd))) {
                    IShellItem* psiResult = NULL;
                    if (SUCCEEDED(pSave->GetResult(&psiResult))) {
                        PWSTR pszFilePath = NULL;
                        if (SUCCEEDED(psiResult->GetDisplayName(SIGDN_FILESYSPATH, &pszFilePath))) {
                            // Write config
                            std::wofstream f(pszFilePath);
                            f << g_rootPath << L"\n";
                            for (auto& item : g_customData) f << item.name << L"|" << item.path << L"\n";
                            f.close();

                            g_configFilePath = pszFilePath; // bind
                            CoTaskMemFree(pszFilePath);
                            UpdateMenuState();
                        }
                        psiResult->Release();
                    }
                }
                pSave->Release();
            }
        }
        else if (id == IDM_FILE_SAVE) {
            if (g_configFilePath.empty()) {
                MessageBox(hWnd, L"No configuration file selected. Use 'Save As...' to choose a file.", L"Save", MB_OK);
            } else {
                std::wofstream f(g_configFilePath);
                f << g_rootPath << L"\n";
                for (auto& item : g_customData) f << item.name << L"|" << item.path << L"\n";
                f.close();
            }
        }
        else if (id == IDM_FILE_LOAD) {
            // Show open dialog filtered to .config files and bind to selected file
            IFileOpenDialog* pOpen;
            if (SUCCEEDED(CoCreateInstance(CLSID_FileOpenDialog, NULL, CLSCTX_ALL, IID_PPV_ARGS(&pOpen)))) {
                COMDLG_FILTERSPEC spec = { L"Config Files (*.config)", L"*.config" };
                pOpen->SetFileTypes(1, &spec);
                pOpen->SetDefaultExtension(L"config");
                DWORD dwOptions;
                if (SUCCEEDED(pOpen->GetOptions(&dwOptions))) {
                    pOpen->SetOptions(dwOptions | FOS_FORCEFILESYSTEM | FOS_PATHMUSTEXIST | FOS_FILEMUSTEXIST);
                }

                if (SUCCEEDED(pOpen->Show(hWnd))) {
                    IShellItem* psiResult = NULL;
                    if (SUCCEEDED(pOpen->GetResult(&psiResult))) {
                        PWSTR pszFilePath = NULL;
                        if (SUCCEEDED(psiResult->GetDisplayName(SIGDN_FILESYSPATH, &pszFilePath))) {
                            // Load file and bind
                            std::wifstream f(pszFilePath);
                            if (f) {
                                g_customData.clear();
                                std::getline(f, g_rootPath);
                                std::wstring line;
                                while (std::getline(f, line)) {
                                    size_t sep = line.find(L'|');
                                    if (sep != std::wstring::npos)
                                        g_customData.push_back({ line.substr(0, sep), line.substr(sep + 1) });
                                }
                                g_configFilePath = pszFilePath;
                                UpdateMenuState();
                                RefreshUI(hWnd);
                            }
                            CoTaskMemFree(pszFilePath);
                        }
                        psiResult->Release();
                    }
                }
                pOpen->Release();
            }
        }
        // Handle Dynamic Buttons using the std::map
        else if (g_buttonMap.count(hCtrl)) {
            if (g_removeMode) {
                // remove this custom button
                std::wstring path = g_buttonMap[hCtrl];
                // remove from g_customData by path
                for (auto it = g_customData.begin(); it != g_customData.end(); ++it) {
                    if (it->path == path) { g_customData.erase(it); break; }
                }
                // remove from map and destroy
                g_buttonMap.erase(hCtrl);
                DestroyWindow(hCtrl);
                g_removeMode = false;
                UpdateMenuState();
                RefreshUI(hWnd);
            } else {
                ShellExecuteW(NULL, L"open", g_buttonMap[hCtrl].c_str(), NULL, NULL, SW_SHOWDEFAULT);
            }
        }
    } break;

    case WM_USER + 1: // Message from Form to refresh buttons
        RefreshUI(hWnd);
        break;

    case WM_ERASEBKGND: {
        if (g_hBgBrush) {
            HDC hdc = (HDC)wParam;
            RECT rc; GetClientRect(hWnd, &rc);
            FillRect(hdc, &rc, g_hBgBrush);
            return 1;
        }
        break;
    }

    case WM_CTLCOLORBTN: {
        HDC hdc = (HDC)wParam;
        SetTextColor(hdc, g_lightGreen);
        SetBkMode(hdc, TRANSPARENT);
        return (LRESULT)g_hBgBrush;
    }
    case WM_CTLCOLOREDIT:
    case WM_CTLCOLORSTATIC: {
        HDC hdc = (HDC)wParam;
        SetTextColor(hdc, g_lightGreen);
        SetBkMode(hdc, TRANSPARENT);
        return (LRESULT)g_hBgBrush;
    }

    case WM_DESTROY:
        if (g_hMatrixFont) { DeleteObject(g_hMatrixFont); g_hMatrixFont = NULL; }
        if (g_hBgBrush) { DeleteObject(g_hBgBrush); g_hBgBrush = NULL; }
        PostQuitMessage(0);
        break;
    default: return DefWindowProc(hWnd, message, wParam, lParam);
    }
    return 0;
}

LRESULT CALLBACK AddPathDlgProc(HWND hWnd, UINT msg, WPARAM wp, LPARAM lp) {
    static HWND hNameEdit;
    static HWND hPathEdit;
    static HWND hBtnPickFolder;
    static HWND hBtnPickFile;
    static HWND hBtnSave;
    static std::wstring selectedPath = L"";
    bool isRootMode = false;

    switch (msg) {
    case WM_CREATE: {
        CREATESTRUCTW* pcs = (CREATESTRUCTW*)lp;
        isRootMode = pcs && pcs->lpCreateParams != NULL;
        // Store mode in window user data for later
        SetWindowLongPtr(hWnd, GWLP_USERDATA, (LONG_PTR)(isRootMode ? 1 : 0));

        if (isRootMode) {
            // Root selection dialog: expand path edit to full width and no label
            hPathEdit = CreateWindowW(L"EDIT", L"", WS_VISIBLE | WS_CHILD | WS_BORDER | ES_READONLY, 10, 10, 360, 25, hWnd, NULL, hInst, NULL);
            hBtnPickFolder = CreateWindowW(L"BUTTON", L"Choose Folder", WS_VISIBLE | WS_CHILD, 10, 50, 120, 30, hWnd, (HMENU)1, hInst, NULL);
            hBtnSave = CreateWindowW(L"BUTTON", L"Save", WS_VISIBLE | WS_CHILD, 150, 50, 100, 30, hWnd, (HMENU)2, hInst, NULL);

            // Apply font if available
            if (g_hMatrixFont) {
                SendMessageW(hPathEdit, WM_SETFONT, (WPARAM)g_hMatrixFont, TRUE);
                SendMessageW(hBtnPickFolder, WM_SETFONT, (WPARAM)g_hMatrixFont, TRUE);
                SendMessageW(hBtnSave, WM_SETFONT, (WPARAM)g_hMatrixFont, TRUE);
            }
        } else {
            // Normal add-path dialog: label + name edit and three buttons in one row
            hNameEdit = CreateWindowW(L"EDIT", L"", WS_VISIBLE | WS_CHILD | WS_BORDER, 110, 20, 200, 25, hWnd, NULL, hInst, NULL);

            // Place three buttons on the same row
            hBtnPickFolder = CreateWindowW(L"BUTTON", L"Pick Folder", WS_VISIBLE | WS_CHILD, 10, 60, 100, 30, hWnd, (HMENU)1, hInst, NULL);
            hBtnPickFile = CreateWindowW(L"BUTTON", L"Pick File", WS_VISIBLE | WS_CHILD, 120, 60, 100, 30, hWnd, (HMENU)3, hInst, NULL);
            hBtnSave = CreateWindowW(L"BUTTON", L"Save", WS_VISIBLE | WS_CHILD, 230, 60, 100, 30, hWnd, (HMENU)2, hInst, NULL);

            // Apply font if available
            if (g_hMatrixFont) {
                SendMessageW(hNameEdit, WM_SETFONT, (WPARAM)g_hMatrixFont, TRUE);
                SendMessageW(hBtnPickFolder, WM_SETFONT, (WPARAM)g_hMatrixFont, TRUE);
                SendMessageW(hBtnPickFile, WM_SETFONT, (WPARAM)g_hMatrixFont, TRUE);
                SendMessageW(hBtnSave, WM_SETFONT, (WPARAM)g_hMatrixFont, TRUE);
            }
        }
        break;
    }

    case WM_ERASEBKGND: {
        if (g_hBgBrush) {
            HDC hdc = (HDC)wp;
            RECT rc; GetClientRect(hWnd, &rc);
            FillRect(hdc, &rc, g_hBgBrush);
            return 1;
        }
        break;
    }

    case WM_CTLCOLORBTN: {
        HDC hdc = (HDC)wp;
        SetTextColor(hdc, RGB(0,0,0));
        SetBkMode(hdc, TRANSPARENT);
        return (LRESULT)g_hBgBrush;
    }
    case WM_CTLCOLOREDIT:
    case WM_CTLCOLORSTATIC: {
        HDC hdc = (HDC)wp;
        SetTextColor(hdc, RGB(0,0,0));
        SetBkMode(hdc, TRANSPARENT);
        return (LRESULT)g_hBgBrush;
    }

    case WM_COMMAND: {
         isRootMode = GetWindowLongPtr(hWnd, GWLP_USERDATA) != 0;
         if (isRootMode) {
             if (LOWORD(wp) == 1) {
                 // Choose folder for root
                 PickFolder(hWnd, selectedPath, false);
                 if (!selectedPath.empty() && hPathEdit) SetWindowTextW(hPathEdit, selectedPath.c_str());
             }
             if (LOWORD(wp) == 2) {
                 // Save root
                 if (!selectedPath.empty()) {
                     g_rootPath = selectedPath;
                     selectedPath.clear();
                     UpdateMenuState();
                     if (g_hMainWnd) PostMessage(g_hMainWnd, WM_USER + 1, 0, 0);
                     DestroyWindow(hWnd);
                 } else {
                     MessageBox(hWnd, L"Please choose a folder before saving.", L"No folder", MB_OK | MB_ICONWARNING);
                 }
             }
         } else {
             if (LOWORD(wp) == 1) {
                 // Pick a folder
                 PickFolder(hWnd, selectedPath, false);
             }
             if (LOWORD(wp) == 3) {
                 // Pick a file
                 PickFolder(hWnd, selectedPath, true);
             }
             if (LOWORD(wp) == 2) {
                 wchar_t name[100]; GetWindowTextW(hNameEdit, name, 100);
                 if (!selectedPath.empty() && wcslen(name) > 0) {
                    g_customData.push_back({ name, selectedPath });
                    selectedPath = L""; // reset static
                    UpdateMenuState();
                    if (g_hMainWnd) PostMessage(g_hMainWnd, WM_USER + 1, 0, 0);
                    DestroyWindow(hWnd);
                 } else {
                     MessageBox(hWnd, L"Please enter a name and choose a path.", L"Incomplete", MB_OK | MB_ICONWARNING);
                 }
             }
         }
         break;
     }

     case WM_DESTROY:
         // nothing special
         break;

     default: return DefWindowProc(hWnd, msg, wp, lp);
     }
     return 0;
 }

void PickFolder(HWND hWnd, std::wstring& outPath, bool allowFiles) {
    IFileOpenDialog* pDlg;
    if (SUCCEEDED(CoCreateInstance(CLSID_FileOpenDialog, NULL, CLSCTX_ALL, IID_PPV_ARGS(&pDlg)))) {
        DWORD dwOpt; pDlg->GetOptions(&dwOpt);
        if (allowFiles) {
            // Do not set FOS_PICKFOLDERS so user can pick files. Keep file system items only.
            pDlg->SetOptions(dwOpt | FOS_FORCEFILESYSTEM | FOS_FILEMUSTEXIST);
        } else {
            pDlg->SetOptions(dwOpt | FOS_PICKFOLDERS);
        }
        if (SUCCEEDED(pDlg->Show(hWnd))) {
            IShellItem* pItem;
            if (SUCCEEDED(pDlg->GetResult(&pItem))) {
                PWSTR psz;
                if (SUCCEEDED(pItem->GetDisplayName(SIGDN_FILESYSPATH, &psz))) {
                    outPath = psz;
                    CoTaskMemFree(psz);
                }
                pItem->Release();
            }
        }
        pDlg->Release();
    }
}

void CenterWindow(HWND hwnd) {
    RECT rc = { 0 };
    GetWindowRect(hwnd, &rc);
    int w = rc.right - rc.left;
    int h = rc.bottom - rc.top;
    int sx = GetSystemMetrics(SM_CXSCREEN);
    int sy = GetSystemMetrics(SM_CYSCREEN);
    int x = (sx - w) / 2;
    int y = (sy - h) / 2;
    SetWindowPos(hwnd, NULL, x, y, 0, 0, SWP_NOSIZE | SWP_NOZORDER);
}