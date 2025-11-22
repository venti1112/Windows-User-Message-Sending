#include <windows.h>
#include <shellapi.h>
#include <shlwapi.h>
#include <string>
#include <vector>
#include <tlhelp32.h>
#include <strsafe.h>

#ifndef WINVER
#define WINVER 0x0600
#endif

#ifndef _WIN32_WINNT
#define _WIN32_WINNT 0x0600
#endif

#ifndef _WIN32_WINDOWS
#define _WIN32_WINDOWS 0x0600
#endif

#ifndef _WIN32_IE
#define _WIN32_IE 0x0700
#endif

#pragma comment(lib, "shell32.lib")
#pragma comment(lib, "shlwapi.lib")
#pragma comment(linker,"\"/manifestdependency:type='win32' name='Microsoft.Windows.Common-Controls' version='6.0.0.0' processorArchitecture='*' publicKeyToken='6595b64144ccf1df' language='*'\"")

// 控件ID
#define IDC_USERNAME_EDIT   1001
#define IDC_MESSAGE_EDIT    1002
#define IDC_SEND_BUTTON     1003

// 函数前置声明
std::string GetLastErrorAsString();
bool IsProcessSuccessful(DWORD exitCode, const std::string& output);

// 全局变量
HWND g_hWndUsername = NULL;
HWND g_hWndMessage = NULL;
HWND g_hWndSendBtn = NULL;
bool g_isSending = false;
HICON g_hAppIcon = NULL;  // 应用程序图标句柄
HICON g_hAppSmallIcon = NULL;  // 小图标句柄

// 函数声明
LRESULT CALLBACK WndProc(HWND, UINT, WPARAM, LPARAM);
void CreateControls(HWND hWnd);
bool IsAdmin();
bool RequestElevation();
DWORD WINAPI SendThreadProc(LPVOID lpParam);
std::string GetCommandOutput(const std::string& command, DWORD* exitCode = nullptr);
void UpdateUIState(bool isSending);
void ShowResultMessage(bool isSuccess, const std::string& output);
void LoadApplicationIcons(HINSTANCE hInstance);
void SetWindowIcons(HWND hWnd);

// 主函数
int APIENTRY WinMain(_In_ HINSTANCE hInstance, _In_opt_ HINSTANCE hPrevInstance, _In_ LPSTR lpCmdLine, _In_ int nCmdShow) {
    UNREFERENCED_PARAMETER(hPrevInstance);
    UNREFERENCED_PARAMETER(lpCmdLine);

    // 检查是否需要提权
    if (!IsAdmin()) {
        if (!RequestElevation()) {
            MessageBoxA(NULL, "无法获取管理员权限，程序将退出", "Windows用户消息发送器", MB_ICONERROR);
            return 1;
        }
        return 0;
    }

    // 加载应用程序图标
    LoadApplicationIcons(hInstance);

    // 注册窗口类
    const char CLASS_NAME[] = "MsgSenderWindowClass";

    WNDCLASSEX wc = {};
    wc.cbSize = sizeof(WNDCLASSEX);
    wc.style = CS_HREDRAW | CS_VREDRAW;
    wc.lpfnWndProc = WndProc;
    wc.hInstance = hInstance;
    wc.hCursor = LoadCursor(nullptr, IDC_ARROW);
    wc.hbrBackground = (HBRUSH)(COLOR_WINDOW + 1);
    wc.lpszClassName = CLASS_NAME;

    // 使用应用程序图标
    wc.hIcon = g_hAppIcon ? g_hAppIcon : LoadIcon(NULL, IDI_APPLICATION);
    wc.hIconSm = g_hAppSmallIcon ? g_hAppSmallIcon : LoadIcon(NULL, IDI_APPLICATION);

    if (!RegisterClassEx(&wc)) {
        MessageBoxA(NULL, "窗口类注册失败!", "Windows用户消息发送器", MB_ICONERROR);
        return 1;
    }

    // 创建窗口
    HWND hWnd = CreateWindowEx(
        0,
        CLASS_NAME,
        "Windows用户消息发送器",
        WS_OVERLAPPEDWINDOW & ~(WS_SIZEBOX | WS_MAXIMIZEBOX),
        CW_USEDEFAULT, CW_USEDEFAULT, 500, 430,
        nullptr,
        nullptr,
        hInstance,
        nullptr
    );

    if (!hWnd) {
        MessageBoxA(NULL, "窗口创建失败!", "Windows用户消息发送器", MB_ICONERROR);
        return 1;
    }

    // 设置窗口图标
    SetWindowIcons(hWnd);

    // 显示窗口
    ShowWindow(hWnd, nCmdShow);
    UpdateWindow(hWnd);

    // 消息循环
    MSG msg = {};
    while (GetMessage(&msg, nullptr, 0, 0)) {
        if (!IsDialogMessage(hWnd, &msg)) {
            TranslateMessage(&msg);
            DispatchMessage(&msg);
        }
    }

    // 清理图标资源
    if (g_hAppIcon) DestroyIcon(g_hAppIcon);
    if (g_hAppSmallIcon) DestroyIcon(g_hAppSmallIcon);

    return (int)msg.wParam;
}

LRESULT CALLBACK WndProc(HWND hWnd, UINT message, WPARAM wParam, LPARAM lParam) {
    switch (message) {
    case WM_CREATE: {
        CreateControls(hWnd);
        break;
    }
    case WM_COMMAND: {
        int wmId = LOWORD(wParam);
        switch (wmId) {
        case IDC_SEND_BUTTON:
            if (!g_isSending) {
                // 创建新线程发送消息，避免阻塞UI
                CreateThread(NULL, 0, SendThreadProc, hWnd, 0, NULL);
            }
            break;
        }
        break;
    }
    case WM_DESTROY:
        PostQuitMessage(0);
        break;
    default:
        return DefWindowProc(hWnd, message, wParam, lParam);
    }
    return 0;
}

void CreateControls(HWND hWnd) {
    // 用户名标签和输入框
    CreateWindowA("STATIC", "目标用户名:", WS_VISIBLE | WS_CHILD, 20, 20, 100, 20, hWnd, NULL, NULL, NULL);
    g_hWndUsername = CreateWindowA("EDIT", "", WS_VISIBLE | WS_CHILD | WS_BORDER, 130, 20, 300, 20, hWnd, (HMENU)IDC_USERNAME_EDIT, NULL, NULL);

    // 消息内容标签和输入框
    CreateWindowA("STATIC", "消息内容:", WS_VISIBLE | WS_CHILD, 20, 50, 100, 20, hWnd, NULL, NULL, NULL);
    g_hWndMessage = CreateWindowA("EDIT", "",
        WS_VISIBLE | WS_CHILD | WS_BORDER | WS_VSCROLL | ES_MULTILINE | ES_AUTOVSCROLL | ES_WANTRETURN,
        20, 75, 460, 250, hWnd, (HMENU)IDC_MESSAGE_EDIT, NULL, NULL);
    SendMessageA(g_hWndMessage, EM_SETLIMITTEXT, 255, 0);

    // 设置默认字体
    NONCLIENTMETRICS ncm = { sizeof(NONCLIENTMETRICS) };
    SystemParametersInfoA(SPI_GETNONCLIENTMETRICS, sizeof(NONCLIENTMETRICS), &ncm, 0);
    HFONT hFont = CreateFontIndirectA(&ncm.lfMessageFont);
    SendMessageA(g_hWndUsername, WM_SETFONT, (WPARAM)hFont, TRUE);
    SendMessageA(g_hWndMessage, WM_SETFONT, (WPARAM)hFont, TRUE);

    // 发送按钮
    g_hWndSendBtn = CreateWindowA("BUTTON", "发送消息", WS_VISIBLE | WS_CHILD | WS_TABSTOP, 200, 340, 100, 30, hWnd, (HMENU)IDC_SEND_BUTTON, NULL, NULL);
    SendMessageA(g_hWndSendBtn, WM_SETFONT, (WPARAM)hFont, TRUE);
}

bool IsAdmin() {
    BOOL fIsAdmin = FALSE;
    PSID pAdministratorsGroup = NULL;

    SID_IDENTIFIER_AUTHORITY NtAuthority = SECURITY_NT_AUTHORITY;
    if (AllocateAndInitializeSid(&NtAuthority, 2, SECURITY_BUILTIN_DOMAIN_RID, DOMAIN_ALIAS_RID_ADMINS,
        0, 0, 0, 0, 0, 0, &pAdministratorsGroup)) {
        if (!CheckTokenMembership(NULL, pAdministratorsGroup, &fIsAdmin)) {
            fIsAdmin = FALSE;
        }
        FreeSid(pAdministratorsGroup);
    }

    return (fIsAdmin == TRUE);
}

bool RequestElevation() {
    SHELLEXECUTEINFO shExecInfo = { sizeof(SHELLEXECUTEINFO) };

    char szPath[MAX_PATH];
    if (GetModuleFileNameA(NULL, szPath, ARRAYSIZE(szPath)) == 0) {
        return false;
    }

    shExecInfo.lpVerb = "runas";
    shExecInfo.lpFile = szPath;
    shExecInfo.nShow = SW_NORMAL;
    shExecInfo.fMask = 0x00000040;

    if (!ShellExecuteExA(&shExecInfo) || shExecInfo.hProcess == NULL) {
        return false;
    }

    CloseHandle(shExecInfo.hProcess);
    return true;
}

DWORD WINAPI SendThreadProc(LPVOID lpParam) {
    HWND hWnd = (HWND)lpParam;

    // 更新UI状态为发送中
    UpdateUIState(true);

    // 获取输入内容
    char username[256] = { 0 };
    // 增加消息缓冲区大小以支持长文本
    const int MESSAGE_BUFFER_SIZE = 16384;
    char* message = new char[MESSAGE_BUFFER_SIZE];
    memset(message, 0, MESSAGE_BUFFER_SIZE);

    GetWindowTextA(g_hWndUsername, username, ARRAYSIZE(username));
    GetWindowTextA(g_hWndMessage, message, MESSAGE_BUFFER_SIZE);

    // 检查输入
    if (strlen(username) == 0) {
        MessageBoxA(hWnd, "请输入目标用户名!", "Windows用户消息发送器", MB_ICONWARNING);
        UpdateUIState(false);
        delete[] message;
        return 1;
    }

    if (strlen(message) == 0) {
        MessageBoxA(hWnd, "请输入消息内容!", "Windows用户消息发送器", MB_ICONWARNING);
        UpdateUIState(false);
        delete[] message;
        return 1;
    }

    if (strlen(message) > 255) {
        MessageBoxA(hWnd, "消息内容过长!", "Windows用户消息发送器", MB_ICONWARNING);
        UpdateUIState(false);
        delete[] message;
        return 1;
    }

    // 构建命令（确保正确处理包含引号和换行的消息）
    std::string escapedMessage = message;
    // 将消息中的双引号转义
    size_t pos = 0;
    while ((pos = escapedMessage.find('"', pos)) != std::string::npos) {
        escapedMessage.replace(pos, 1, "\\\"");
        pos += 2;
    }

    std::string command = "msg " + std::string(username) + " \"" + escapedMessage + "\"";

    // 执行命令并获取输出和退出码
    DWORD exitCode = 0;
    std::string output = GetCommandOutput(command, &exitCode);

    // 显示结果消息（成功或失败）
    ShowResultMessage(IsProcessSuccessful(exitCode, output), output);

    // 更新UI状态为完成
    UpdateUIState(false);

    delete[] message;
    return 0;
}

std::string GetCommandOutput(const std::string& command, DWORD* exitCode) {
    SECURITY_ATTRIBUTES sa = { sizeof(SECURITY_ATTRIBUTES) };
    sa.bInheritHandle = TRUE;
    sa.lpSecurityDescriptor = NULL;

    HANDLE hRead = NULL;
    HANDLE hWrite = NULL;

    // 创建管道
    if (!CreatePipe(&hRead, &hWrite, &sa, 0)) {
        if (exitCode) *exitCode = 1;
        return "创建管道失败";
    }

    // 设置标准输出重定向
    STARTUPINFO si = { sizeof(STARTUPINFO) };
    si.dwFlags = STARTF_USESTDHANDLES | STARTF_USESHOWWINDOW;
    si.hStdOutput = hWrite;
    si.hStdError = hWrite;
    si.wShowWindow = SW_HIDE;

    // 创建进程
    PROCESS_INFORMATION pi = { 0 };

    // 需要临时缓冲区来修改命令字符串（因为CreateProcessA需要非const char*）
    char* cmd = new char[command.length() + 1];
    strcpy_s(cmd, command.length() + 1, command.c_str());

    BOOL result = CreateProcessA(
        NULL,
        cmd,
        NULL,
        NULL,
        TRUE,  // 继承句柄
        CREATE_NO_WINDOW,
        NULL,
        NULL,
        &si,
        &pi
    );

    delete[] cmd;

    if (!result) {
        CloseHandle(hRead);
        CloseHandle(hWrite);
        if (exitCode) *exitCode = GetLastError();
        return "创建进程失败: " + GetLastErrorAsString();
    }

    // 关闭写入端，因为子进程不需要它
    CloseHandle(hWrite);

    // 读取输出
    std::string output;
    char buffer[1024];
    DWORD bytesRead;

    while (ReadFile(hRead, buffer, sizeof(buffer) - 1, &bytesRead, NULL) && bytesRead != 0) {
        buffer[bytesRead] = '\0';
        output += buffer;
    }

    // 等待进程结束并获取退出码
    WaitForSingleObject(pi.hProcess, INFINITE);
    if (exitCode) {
        GetExitCodeProcess(pi.hProcess, exitCode);
    }

    // 清理
    CloseHandle(pi.hProcess);
    CloseHandle(pi.hThread);
    CloseHandle(hRead);

    return output;
}

bool IsProcessSuccessful(DWORD exitCode, const std::string& output) {
    // msg命令成功时通常返回0
    if (exitCode == 0) {
        return true;
    }

    // 检查输出中是否包含成功信息（某些系统可能不同）
    if (output.find("成功") != std::string::npos ||
        output.find("sent") != std::string::npos) {
        return true;
    }

    return false;
}

void ShowResultMessage(bool isSuccess, const std::string& output) {
    if (isSuccess) {
        MessageBoxA(NULL, "消息已成功发送!", "Windows用户消息发送器", MB_ICONINFORMATION | MB_OK);
    }
    else {
        std::string errorMsg = "消息发送失败!\n\n详细信息:\n" + output;
        MessageBoxA(NULL, errorMsg.c_str(), "Windows用户消息发送器", MB_ICONERROR | MB_OK);
    }
}

void UpdateUIState(bool isSending) {
    g_isSending = isSending;

    // 更新发送按钮状态
    EnableWindow(g_hWndSendBtn, !isSending);
    SetWindowTextA(g_hWndSendBtn, isSending ? "发送中..." : "发送消息");
}

std::string GetLastErrorAsString() {
    DWORD errorMessageID = GetLastError();
    if (errorMessageID == 0) {
        return "没有错误信息";
    }

    LPSTR messageBuffer = nullptr;
    size_t size = FormatMessageA(
        FORMAT_MESSAGE_ALLOCATE_BUFFER | FORMAT_MESSAGE_FROM_SYSTEM | FORMAT_MESSAGE_IGNORE_INSERTS,
        NULL,
        errorMessageID,
        MAKELANGID(LANG_NEUTRAL, SUBLANG_DEFAULT),
        (LPSTR)&messageBuffer,
        0,
        NULL
    );

    std::string message(messageBuffer, size);
    LocalFree(messageBuffer);
    return message;
}

// 加载应用程序图标
void LoadApplicationIcons(HINSTANCE hInstance) {
    char szExePath[MAX_PATH];
    if (GetModuleFileNameA(NULL, szExePath, ARRAYSIZE(szExePath)) > 0) {
        // 从可执行文件中提取图标
        g_hAppIcon = ExtractIconA(hInstance, szExePath, 0);

        // 如果提取成功，获取小图标
        if (g_hAppIcon && g_hAppIcon != (HICON)ERROR_FILE_NOT_FOUND) {
            int smIconX = GetSystemMetrics(SM_CXSMICON);
            int smIconY = GetSystemMetrics(SM_CYSMICON);

            // 创建小尺寸图标
            g_hAppSmallIcon = (HICON)CopyImage(
                g_hAppIcon,
                IMAGE_ICON,
                smIconX,
                smIconY,
                LR_COPYFROMRESOURCE
            );

            // 如果复制失败，使用原始图标
            if (!g_hAppSmallIcon) {
                g_hAppSmallIcon = g_hAppIcon;
            }
        }
    }

    // 如果没有成功加载图标，使用系统默认图标
    if (!g_hAppIcon || g_hAppIcon == (HICON)ERROR_FILE_NOT_FOUND) {
        g_hAppIcon = LoadIcon(NULL, IDI_APPLICATION);
        g_hAppSmallIcon = LoadIcon(NULL, IDI_APPLICATION);
    }
}

// 设置窗口图标
void SetWindowIcons(HWND hWnd) {
    if (g_hAppIcon) {
        SendMessageA(hWnd, WM_SETICON, ICON_BIG, (LPARAM)g_hAppIcon);
    }
    if (g_hAppSmallIcon) {
        SendMessageA(hWnd, WM_SETICON, ICON_SMALL, (LPARAM)g_hAppSmallIcon);
    }
}