//
// Practices on Speech Recognition.
// Using Microsoft Speech API (SAPI) 5.3
//
// y.s.n@live.com, 2014-09
//
#include "stdafx.h"
#include "WaveToText.h"
#include "resource.h"


BOOL CALLBACK MainDiagProc(HWND hWnd, UINT Message, WPARAM wParam, LPARAM lParam);

HINSTANCE g_hInst = NULL;
CWaveToText wtt;

void OnButtonBrowse(HWND hWnd)
{
    TCHAR lpszFileName[MAX_PATH] = { 0 };
    OPENFILENAME ofn;
    ZeroMemory(&ofn, sizeof(ofn));
    ofn.lStructSize = sizeof(ofn);
    ofn.hwndOwner = hWnd;
    ofn.lpstrFile = lpszFileName;
    ofn.nMaxFile = MAX_PATH;
    ofn.lpstrFilter = _T("wave file\0*.wav\0All\0*.*\0");
    ofn.nFilterIndex = 1;
    ofn.Flags = OFN_PATHMUSTEXIST | OFN_FILEMUSTEXIST;
    if (GetOpenFileName(&ofn)) {
        SetDlgItemText(hWnd, IDC_EDIT_INPUT_FILE, ofn.lpstrFile);
    }
}

bool VerifyFileExisting(LPCTSTR lpszFile)
{
    WIN32_FIND_DATA Win32_Find_Data;
    HANDLE hFindFile = FindFirstFile(lpszFile, &Win32_Find_Data);
    if (INVALID_HANDLE_VALUE == hFindFile) {
        return false;
    } else {
        FindClose(hFindFile);
        return true;
    }
}

void SpeechRecognitionCallback(WPARAM wParam, LPARAM lParam, LPCTSTR lpszText)
{
    HWND hWnd = reinterpret_cast<HWND>(wParam);
    ::SetDlgItemText(hWnd, IDC_RICHEDIT_RESULT, lpszText);
}

void LaunchRecognition(HWND hWnd)
{
    TCHAR szFile[MAX_PATH] = { 0 };
    ::GetDlgItemText(hWnd, IDC_EDIT_INPUT_FILE, szFile, MAX_PATH);

    if (!VerifyFileExisting(szFile)) {
        ::MessageBox(hWnd, _T("No valid file was found."), _T("Error"), MB_OK);
        return;
    }

    wtt.SetInputWaveFile(szFile);
    wtt.SetRecognitionCallback(SpeechRecognitionCallback, reinterpret_cast<WPARAM>(hWnd), 0);
    if (wtt.Start() != 0) {
        ::MessageBox(hWnd, _T("Fail to convert audio to text."), _T("Error"), MB_OK);
        return;
    }
}

INT_PTR CALLBACK AboutDialogProc(HWND hDlg, UINT message, WPARAM wParam, LPARAM lParam)
{
    UNREFERENCED_PARAMETER(lParam);
    switch (message) {
    case WM_INITDIALOG:
        return (INT_PTR)TRUE;
    case WM_COMMAND:
        if (LOWORD(wParam) == IDOK || LOWORD(wParam) == IDCANCEL) {
            EndDialog(hDlg, LOWORD(wParam));
            return (INT_PTR)TRUE;
        }
        break;
    }
    return (INT_PTR)FALSE;
}

void ShowAboutDialogBox(HWND hWnd)
{
    DialogBox(g_hInst, MAKEINTRESOURCE(IDD_ABOUTBOX), hWnd, AboutDialogProc);
}

BOOL CALLBACK MainDiagProc(HWND hWnd, UINT message, WPARAM wParam, LPARAM lParam)
{
    switch (message) {
    case WM_INITDIALOG:
        break;
    case WM_COMMAND:
        switch (LOWORD(wParam)) {
        case IDC_BUTTON_BROWSE:
            OnButtonBrowse(hWnd);
            break;
        case IDC_BUTTON_START:
            LaunchRecognition(hWnd);
            break;
        case IDC_BUTTON_STOP:
            break;
        case IDC_BUTTON_ABOUT:
            ShowAboutDialogBox(hWnd);
            break;
        }
        break;
    case WM_CLOSE:
        EndDialog(hWnd, 0);
        break;
    default:
        return FALSE;
    }
    return TRUE;
}

int WINAPI WinMain(HINSTANCE hInstance, HINSTANCE hPrevInstance, LPSTR lpCmdLine, int nShowCmd)
{
    g_hInst = hInstance;

    HMODULE hRichEdit2 = LoadLibrary(_T("Riched20.dll"));
    if (!hRichEdit2) {
        ::MessageBox(NULL, _T("Fail to load Riched20.dll. Please install rich edit 2.0."), _T("Error"), MB_OK);
        return -1;
    }

    DialogBox(hInstance, MAKEINTRESOURCE(IDD_DIALOG1), NULL, MainDiagProc);

    FreeLibrary(hRichEdit2);
    return 0;
}
