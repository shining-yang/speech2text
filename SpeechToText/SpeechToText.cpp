//
// Practices on Speech Recognition.
// Using Microsoft Speech API (SAPI) 5.3
//
// y.s.n@live.com, 2014-09
//
#include "stdafx.h"
#include "resource.h"
#define WM_RECOEVENT            WM_USER + 1
#define SAMPLE_WAV_FILE         _T("sample.wav")

BOOL CALLBACK DlgProc(HWND hWnd, UINT Message, WPARAM wParam, LPARAM lParam);
void LaunchRecognition(HWND hWnd);
void HandleEvent(HWND hWnd);
WCHAR *ExtractInput(CSpEvent& event);
void CleanupSAPI();

CComPtr<ISpRecognizer> g_cpEngine;
CComPtr<ISpRecoContext> g_cpRecoCtx;
CComPtr<ISpRecoGrammar> g_cpRecoGrammar;

int WINAPI WinMain(HINSTANCE hInstance, HINSTANCE hPrevInstance, LPSTR lpCmdLine, int nShowCmd)
{
    DialogBox(hInstance, MAKEINTRESOURCE(IDD_DIALOG1), NULL, DlgProc);
    return 0;
}

BOOL CALLBACK DlgProc(HWND hWnd, UINT Message, WPARAM wParam, LPARAM lParam)
{
    switch (Message) {
    case WM_RECOEVENT:
        HandleEvent(hWnd);
        break;
    case WM_COMMAND:
        switch (LOWORD(wParam)) {
        case ID_START_RECOG:
            LaunchRecognition(hWnd);
            break;
        }
        break;
    case WM_CLOSE:
        CleanupSAPI();
        EndDialog(hWnd, 0);
        break;
    default:
        return FALSE;
    }
    return TRUE;
}

void LaunchRecognition(HWND hWnd)
{
    HRESULT hr = S_OK;

    if (FAILED(::CoInitialize(NULL))) {
        throw std::string("Unable to initialize COM objects");
    }

    CComPtr<ISpStream> spInputStream;
    hr = spInputStream.CoCreateInstance(CLSID_SpStream);
    if (FAILED(hr)) {
        throw std::string("Fail to create ISpStream");
    }

    CSpStreamFormat inputFormat;
    hr = inputFormat.AssignFormat(SPSF_22kHz16BitStereo);
    if (FAILED(hr)) {
        throw std::string("Fail to AssignFormat");
    }

    hr = spInputStream->BindToFile(
        SAMPLE_WAV_FILE,
        SPFM_OPEN_READONLY,
        &inputFormat.FormatId(),
        inputFormat.WaveFormatExPtr(),
        SPFEI(SPEI_SOUND_START) | SPFEI(SPEI_SOUND_END) |
        SPFEI(SPEI_PHRASE_START) | SPFEI(SPEI_RECOGNITION));
    if (FAILED(hr)) {
        throw std::string("Fail to BindToFile");
    }


    ULONGLONG ullGramId = 1;
    hr = g_cpEngine.CoCreateInstance(CLSID_SpInprocRecognizer);
    if (FAILED(hr)) {
        throw std::string("Unable to create recognition engine");
    }

    hr = g_cpEngine->SetInput(spInputStream, TRUE);
    if (FAILED(hr)) {
        throw std::string("Unable to SetInput");
    }

    hr = g_cpEngine->CreateRecoContext(&g_cpRecoCtx);
    if (FAILED(hr)) {
        throw std::string("Failed command recognition");
    }

    hr = g_cpRecoCtx->SetNotifyWindowMessage(hWnd, WM_RECOEVENT, 0, 0);
    if (FAILED(hr)) {
        throw std::string("Unable to select notification window");
    }

    const ULONGLONG ullInterest =
        SPFEI(SPEI_SOUND_START) | SPFEI(SPEI_SOUND_END) |
        SPFEI(SPEI_PHRASE_START) | SPFEI(SPEI_RECOGNITION) |
        SPFEI(SPEI_FALSE_RECOGNITION) | SPFEI(SPEI_HYPOTHESIS) |
        SPFEI(SPEI_INTERFERENCE) | SPFEI(SPEI_RECO_OTHER_CONTEXT) |
        SPFEI(SPEI_REQUEST_UI) | SPFEI(SPEI_RECO_STATE_CHANGE) |
        SPFEI(SPEI_PROPERTY_NUM_CHANGE) | SPFEI(SPEI_PROPERTY_STRING_CHANGE);

    hr = g_cpRecoCtx->SetInterest(ullInterest, ullInterest);
    if (FAILED(hr)) {
        throw std::string("Failed to create interest");
    }

    hr = g_cpRecoCtx->CreateGrammar(0, &g_cpRecoGrammar);
    if (FAILED(hr)) {
        throw std::string("Unable to create grammar");
    }

    hr = g_cpRecoGrammar->LoadDictation(0, SPLO_STATIC);
    if (FAILED(hr)) {
        throw std::string("Failed to load dictation");
    }

    hr = g_cpRecoGrammar->SetDictationState(SPRS_ACTIVE);
    if (FAILED(hr)) {
        throw std::string("Failed setting dictation state");
    }
}

void HandleEvent(HWND hWnd)
{
    CSpEvent event;
    WCHAR  *pwszText;

    static bool bFinishProcessing = false;
    if (bFinishProcessing) {
        return;
    }

    // Loop processing events while there are any in the queue
    while (event.GetFrom(g_cpRecoCtx) == S_OK) {
        switch (event.eEventId) {
        case SPEI_RECOGNITION:
            pwszText = ExtractInput(event);
            SetDlgItemTextW(hWnd, IDC_EDIT1, pwszText);
            break;
        case SPEI_FALSE_RECOGNITION:
            break;
        case SPEI_END_SR_STREAM:
            bFinishProcessing = true;
            break;
        }
    }
}

WCHAR *ExtractInput(CSpEvent& event)
{
    HRESULT                   hr = S_OK;
    CComPtr<ISpRecoResult>    cpRecoResult;
    SPPHRASE                  *pPhrase;
    WCHAR                     *pwszText = NULL;

    cpRecoResult = event.RecoResult();

    hr = cpRecoResult->GetPhrase(&pPhrase);
    if (SUCCEEDED(hr)) {
        if (event.eEventId == SPEI_FALSE_RECOGNITION) {
            pwszText = L"False recognition";
        } else {
            // Get the phrase's entire text string, including replacements.
            hr = cpRecoResult->GetText(SP_GETWHOLEPHRASE, SP_GETWHOLEPHRASE, TRUE, &pwszText, NULL);
        }
    }
    CoTaskMemFree(pPhrase);
    return pwszText;
}

void CleanupSAPI()
{
    if (g_cpRecoGrammar) {
        g_cpRecoGrammar.Release();
    }
    if (g_cpRecoCtx) {
        g_cpRecoCtx->SetNotifySink(NULL);
        g_cpRecoCtx.Release();
    }
    if (g_cpEngine) {
        g_cpEngine.Release();
    }
    CoUninitialize();
}