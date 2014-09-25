//
// Practices on Speech Recognition.
// Using Microsoft Speech API (SAPI) 5.3
//
// y.s.n@live.com, 2014-09
//
#include "stdafx.h"
#include "resource.h"
#include "WaveToText.h"


#define WM_NOTIFY_ICON          (WM_USER + 1)

HINSTANCE           g_hInst = NULL;
NOTIFYICONDATA      g_nid = { 0 };
CWaveToText         g_wtt;
BOOL                g_bChildWndShown = FALSE;

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

void AppendRecognizedText(HWND hWnd, LPCTSTR lpszText)
{
    GETTEXTLENGTHEX gtlex;
    gtlex.codepage = 1200;
    gtlex.flags = GTL_DEFAULT;

    int nCurLen;
    nCurLen = (int)SendMessage(hWnd, EM_GETTEXTLENGTHEX, (WPARAM)&gtlex, 0);

    CHARRANGE cr;
    cr.cpMin = nCurLen;
    cr.cpMax = -1;

    SendMessage(hWnd, EM_EXSETSEL, 0, (LPARAM)&cr);
    SendMessage(hWnd, EM_REPLACESEL, (WPARAM)&cr, (LPARAM)lpszText);
    SendMessage(hWnd, WM_VSCROLL, (WPARAM)SB_BOTTOM, 0);
}

void EnableDialogItem(HWND hDlg, UINT nIdDlgItem, BOOL bEnable)
{
    HWND hWndItem = GetDlgItem(hDlg, nIdDlgItem);
    if (hWndItem) {
        EnableWindow(hWndItem, bEnable);
    }
}

void RecognitionNotifyStarted(HWND hWnd)
{
    EnableDialogItem(hWnd, IDC_BUTTON_STOP, TRUE);
    SetDlgItemText(hWnd, IDC_STATIC_LOG, _T("Recognition has started"));
}

void RecognitionNotifyEnded(HWND hWnd)
{
    SetDlgItemText(hWnd, IDC_STATIC_LOG, _T("Recognition has stopped"));
    EnableDialogItem(hWnd, IDC_BUTTON_START, TRUE);
    EnableDialogItem(hWnd, IDC_BUTTON_STOP, FALSE);
    SetDlgItemText(hWnd, IDC_STATIC_LOG, _T(""));
}

void RecognitionNotifySuccess(HWND hWnd, CONST VOID* pData)
{
    SetDlgItemText(hWnd, IDC_STATIC_LOG, _T("Recognition is processing ..."));

    HWND hRichEdit = ::GetDlgItem(hWnd, IDC_RICHEDIT_RESULT);
    if (hRichEdit) {
        AppendRecognizedText(hRichEdit, reinterpret_cast<LPCTSTR>(pData));
        AppendRecognizedText(hRichEdit, _T("\n"));
    }
}

void SpeechRecognitionCallback(WPARAM wParam, LPARAM lParam, LPRECOG_NOTIFY_DATA lpNotifyData)
{
    HWND hWnd = reinterpret_cast<HWND>(wParam);

    switch (lpNotifyData->event) {
    case RECOG_STARTED:
        RecognitionNotifyStarted(hWnd);
        break;
    case RECOG_ENDED:
        RecognitionNotifyEnded(hWnd);
        break;
    case RECOG_SUCCESS:
        RecognitionNotifySuccess(hWnd, lpNotifyData->data);
        break;
    }
}

void LaunchRecognition(HWND hWnd)
{
    TCHAR szFile[MAX_PATH] = { 0 };
    ::GetDlgItemText(hWnd, IDC_EDIT_INPUT_FILE, szFile, MAX_PATH);

    if (!VerifyFileExisting(szFile)) {
        ::MessageBox(hWnd, _T("No valid file was found."), _T("Error"), MB_OK);
        return;
    }

    EnableDialogItem(hWnd, IDC_BUTTON_START, FALSE);

    g_wtt.SetInputWaveFile(szFile);
    g_wtt.SetRecognitionCallback(SpeechRecognitionCallback, reinterpret_cast<WPARAM>(hWnd), 0);
    if (g_wtt.Start() != 0) {
        ::MessageBox(hWnd, _T("Fail to convert audio to text."), _T("Error"), MB_OK);
        EnableDialogItem(hWnd, IDC_BUTTON_START, TRUE);
        return;
    } else {
        SetDlgItemText(hWnd, IDC_RICHEDIT_RESULT, _T(""));
    }
}

void StopRecognition(HWND hWnd)
{
    EnableDialogItem(hWnd, IDC_BUTTON_STOP, FALSE);
    SetDlgItemText(hWnd, IDC_STATIC_LOG, _T("Recognition engine is stopping ..."));
    g_wtt.Stop();
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

void InitRicheditCtrl(HWND hWnd)
{
    HWND hRichEdit = GetDlgItem(hWnd, IDC_RICHEDIT_RESULT);
    if (!hRichEdit) {
        return;
    }

    SendMessage(hRichEdit, WM_SETFONT, (WPARAM)GetStockObject(DEFAULT_GUI_FONT), (LPARAM)FALSE);
    SendMessage(hRichEdit, EM_SETEVENTMASK, 0, (LPARAM)(ENM_KEYEVENTS | ENM_MOUSEEVENTS));
}

void OnInitMainDialog(HWND hWnd)
{
    CONST TCHAR* lpszText = _T("Please specify a wave-formatted audio file to proceed.");
    SetDlgItemText(hWnd, IDC_EDIT_INPUT_FILE, lpszText);
    EnableDialogItem(hWnd, IDC_BUTTON_START, TRUE);
    EnableDialogItem(hWnd, IDC_BUTTON_STOP, FALSE);

    InitRicheditCtrl(hWnd);

    // {{ Shell Notification Icon
    ZeroMemory(&g_nid, sizeof(NOTIFYICONDATA));
    g_nid.cbSize = sizeof(NOTIFYICONDATA);
    g_nid.hWnd = hWnd;
    g_nid.hIcon = ::LoadIcon(g_hInst, MAKEINTRESOURCE(IDI_SPEECHTOTEXT));
    g_nid.uCallbackMessage = WM_NOTIFY_ICON;
    g_nid.uTimeout = 3000;
    g_nid.uFlags = NIF_ICON | NIF_TIP | NIF_MESSAGE;
    _tcscpy_s(g_nid.szTip, _T("SpeechToText\nConvert wave audio to text messages"));
    _tcscpy_s(g_nid.szInfoTitle, _T("SpeechToText"));
    _tcscpy_s(g_nid.szInfo, _T("Convert wave audio to text messages."));
    Shell_NotifyIcon(NIM_ADD, &g_nid);
    // }}
}

void ShowOrHideMainWindow(HWND hWnd)
{
    if (IsIconic(hWnd)) {
        ShowWindow(hWnd, SW_RESTORE);
        SetForegroundWindow(hWnd);
    } else if (IsWindowVisible(hWnd) && !g_bChildWndShown) {
        ShowWindow(hWnd, SW_HIDE);
    } else {
        ShowWindow(hWnd, SW_SHOW);
        SetForegroundWindow(hWnd);
    }
}

void OnResultRicheditNotifyRButtonDown(HWND hWnd, WPARAM wParam, LPARAM lParam)
{
    UNREFERENCED_PARAMETER(wParam);
    HWND hRichEdit = GetDlgItem(hWnd, IDC_RICHEDIT_RESULT);
    if (!hRichEdit) {
        return;
    }

    HMENU hMenu = LoadMenu(NULL, MAKEINTRESOURCE(IDC_SPEECHTOTEXT));
    if (hMenu == NULL) {
        return;
    }

    HMENU hMenuPopup = GetSubMenu(hMenu, 0);
    if (hMenuPopup == NULL) {
        return;
    }

    POINT pt = { LOWORD(lParam), HIWORD(lParam) };
    ClientToScreen(hRichEdit, &pt);

    GETTEXTLENGTHEX gtlex;
    gtlex.codepage = 1200;
    gtlex.flags = GTL_DEFAULT;
    int nCurLen;
    nCurLen = (int)SendMessage(hRichEdit, EM_GETTEXTLENGTHEX, (WPARAM)&gtlex, 0);
    if (nCurLen <= 0) {
        EnableMenuItem(hMenuPopup, IDM_RESULT_SELECTALL, MF_BYCOMMAND | MF_GRAYED);
    }

    CHARRANGE cr = { 0 };
    SendMessage(hRichEdit, EM_EXGETSEL, 0, (LPARAM)&cr);
    if (cr.cpMin == cr.cpMax) {
        EnableMenuItem(hMenuPopup, IDM_RESULT_COPY, MF_BYCOMMAND | MF_GRAYED);
    }

    TrackPopupMenuEx(hMenuPopup, TPM_LEFTALIGN | TPM_LEFTBUTTON, pt.x, pt.y, hWnd, NULL);
    DestroyMenu(hMenuPopup);
}

void OnRichEditResultClear(HWND hWnd)
{
    SetDlgItemText(hWnd, IDC_RICHEDIT_RESULT, _T(""));
}

void OnRichEditResultCopy(HWND hWnd)
{
    SendDlgItemMessage(hWnd, IDC_RICHEDIT_RESULT, WM_COPY, 0, 0);
}

void OnRichEditResultSelectAll(HWND hWnd)
{
    CHARRANGE cr;
    cr.cpMin = 0;
    cr.cpMax = -1;
    SendDlgItemMessage(hWnd, IDC_RICHEDIT_RESULT, EM_EXSETSEL, 0, (LPARAM)&cr);
}

void OnResultRicheditNotifyKeyUp(HWND hWnd, WPARAM wParam, LPARAM lParam)
{
    UNREFERENCED_PARAMETER(lParam);
    UINT nVKey = (UINT)wParam;
    switch (nVKey) {
    case VK_DELETE:
        OnRichEditResultClear(hWnd);
        break;
    }
}

void OnResultRicheditNotify(HWND hWnd, WPARAM wParam, LPARAM lParam)
{
    UNREFERENCED_PARAMETER(wParam);
    NMHDR* nmhdr = (NMHDR*)lParam;
    MSGFILTER* msgFilter = NULL;

    switch (nmhdr->code) {
    case EN_MSGFILTER:
        msgFilter = (MSGFILTER*)lParam;
        switch (msgFilter->msg) {
        case WM_RBUTTONDOWN:
            OnResultRicheditNotifyRButtonDown(hWnd, msgFilter->wParam, msgFilter->lParam);
            break;

        case WM_KEYDOWN:
            OnResultRicheditNotifyKeyUp(hWnd, msgFilter->wParam, msgFilter->lParam);
            break;
        }
        break;
    }
}

void OnCtrlNotify(HWND hWnd, WPARAM wParam, LPARAM lParam)
{
    NMHDR* nmhdr = (NMHDR*)lParam;
    switch (nmhdr->idFrom) {
    case IDC_RICHEDIT_RESULT:
        OnResultRicheditNotify(hWnd, wParam, lParam);
        break;
    }
}

BOOL CALLBACK MainDiagProc(HWND hWnd, UINT message, WPARAM wParam, LPARAM lParam)
{
    switch (message) {
    case WM_NOTIFY:
        OnCtrlNotify(hWnd, wParam, lParam);
        break;

    case WM_INITDIALOG:
        OnInitMainDialog(hWnd);
        break;

    case WM_NOTIFY_ICON:
        switch (static_cast<UINT>(lParam)) {
        case WM_LBUTTONDOWN:
            ShowOrHideMainWindow(hWnd);
            break;
        case WM_RBUTTONDOWN:
            break;
        }
        break;

    case WM_COMMAND:
        switch (LOWORD(wParam)) {
        case IDC_BUTTON_BROWSE:
            g_bChildWndShown = TRUE;
            OnButtonBrowse(hWnd);
            g_bChildWndShown = FALSE;
            break;
        case IDC_BUTTON_START:
            LaunchRecognition(hWnd);
            break;
        case IDC_BUTTON_STOP:
            StopRecognition(hWnd);
            break;
        case IDC_BUTTON_ABOUT:
            g_bChildWndShown = TRUE;
            ShowAboutDialogBox(hWnd);
            g_bChildWndShown = FALSE;
            break;
        case IDM_RESULT_COPY:
            OnRichEditResultCopy(hWnd);
            break;
        case IDM_RESULT_CLEAR:
            OnRichEditResultClear(hWnd);
            break;
        case IDM_RESULT_SELECTALL:
            OnRichEditResultSelectAll(hWnd);
            break;
        }
        break;

    case WM_CLOSE:
        Shell_NotifyIcon(NIM_DELETE, &g_nid);
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
