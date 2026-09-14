/* Tiny Windows launcher: starts bin\BantaneShiftOptimizer.exe with the
 * launcher's own directory as the working directory, so files/, input/,
 * output/ stay at the top level while all runtime DLLs live in bin/. */
#include <windows.h>

int WINAPI wWinMain(HINSTANCE hInst, HINSTANCE hPrev, PWSTR cmdline, int show)
{
    wchar_t dir[MAX_PATH];
    wchar_t exe[MAX_PATH];
    STARTUPINFOW si;
    PROCESS_INFORMATION pi;

    if (!GetModuleFileNameW(NULL, dir, MAX_PATH)) return 1;
    wchar_t *slash = wcsrchr(dir, L'\\');
    if (slash) *slash = L'\0';

    if (swprintf(exe, MAX_PATH, L"\"%s\\bin\\BantaneShiftOptimizer.exe\"", dir) < 0) return 1;

    ZeroMemory(&si, sizeof(si));
    si.cb = sizeof(si);
    ZeroMemory(&pi, sizeof(pi));

    if (!CreateProcessW(NULL, exe, NULL, NULL, FALSE, 0, NULL, dir, &si, &pi)) {
        MessageBoxW(NULL, L"bin\\BantaneShiftOptimizer.exe が見つかりません。\nフォルダ構成を確認してください。",
                    L"BantaneShiftOptimizer", MB_ICONERROR | MB_OK);
        return 1;
    }
    CloseHandle(pi.hThread);
    CloseHandle(pi.hProcess);
    return 0;
}
