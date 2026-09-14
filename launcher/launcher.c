/* Tiny Windows launcher: starts bin\BantaneShiftOptimizer.exe with the
 * launcher's own directory as the working directory, so files/, input/,
 * output/ stay at the top level while all runtime DLLs live in bin/. */
#include <windows.h>
#include <stdlib.h>

#define TARGET L"\\bin\\BantaneShiftOptimizer.exe"

static wchar_t *get_own_dir(void)
{
    DWORD cap = MAX_PATH;
    for (;;) {
        wchar_t *buf = (wchar_t *)malloc(cap * sizeof(wchar_t));
        if (!buf) return NULL;
        DWORD n = GetModuleFileNameW(NULL, buf, cap);
        if (n == 0) { free(buf); return NULL; }
        if (n < cap - 1) {
            wchar_t *slash = wcsrchr(buf, L'\\');
            if (slash) *slash = L'\0';
            return buf;
        }
        free(buf);
        cap *= 2;
        if (cap > 65536) return NULL;
    }
}

int WINAPI wWinMain(HINSTANCE hInst, HINSTANCE hPrev, PWSTR cmdline, int show)
{
    STARTUPINFOW si;
    PROCESS_INFORMATION pi;

    wchar_t *dir = get_own_dir();
    if (!dir) return 1;

    size_t len = wcslen(dir) + wcslen(TARGET) + 3; /* quotes + NUL */
    wchar_t *exe = (wchar_t *)malloc(len * sizeof(wchar_t));
    if (!exe) { free(dir); return 1; }
    wcscpy_s(exe, len, L"\"");
    wcscat_s(exe, len, dir);
    wcscat_s(exe, len, TARGET);
    wcscat_s(exe, len, L"\"");

    ZeroMemory(&si, sizeof(si));
    si.cb = sizeof(si);
    ZeroMemory(&pi, sizeof(pi));

    BOOL ok = CreateProcessW(NULL, exe, NULL, NULL, FALSE, 0, NULL, dir, &si, &pi);
    if (!ok) {
        MessageBoxW(NULL, L"bin\\BantaneShiftOptimizer.exe が見つかりません。\nフォルダ構成を確認してください。",
                    L"BantaneShiftOptimizer", MB_ICONERROR | MB_OK);
    } else {
        CloseHandle(pi.hThread);
        CloseHandle(pi.hProcess);
    }
    free(exe);
    free(dir);
    return ok ? 0 : 1;
}
