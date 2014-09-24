//
// File: WaveToText
//
#include "stdafx.h"
#include "WaveToText.h"


CWaveToText::CWaveToText()
{
    Init();
}

CWaveToText::~CWaveToText()
{
    CleanUp();
}

void CWaveToText::SetInputWaveFile(LPCTSTR lpszFile)
{
    if (lpszFile) {
        StringCchCopy(m_strFileName, MAX_PATH, lpszFile);
    }
}

void CWaveToText::SetRecognitionCallback(RECOGNITIONFUNC callback, WPARAM wParam, LPARAM lParam)
{
    m_fnRecognitionCallback = callback;
    m_wpCallback = wParam;
    m_lpCallback = lParam;
}

int CWaveToText::Start()
{
    if (_tcslen(m_strFileName) == 0) {
        return -1;
    }

    CleanUp();

    HRESULT hr = S_OK;
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
        m_strFileName,
        SPFM_OPEN_READONLY,
        &inputFormat.FormatId(),
        inputFormat.WaveFormatExPtr(),
        SPFEI(SPEI_SOUND_START) | SPFEI(SPEI_SOUND_END) |
        SPFEI(SPEI_PHRASE_START) | SPFEI(SPEI_RECOGNITION));
    if (FAILED(hr)) {
        throw std::string("Fail to BindToFile");
    }

    ULONGLONG ullGramId = 1;
    hr = m_recognizer.CoCreateInstance(CLSID_SpInprocRecognizer);
    if (FAILED(hr)) {
        throw std::string("Unable to create recognition engine");
    }

    hr = m_recognizer->SetInput(spInputStream, TRUE);
    if (FAILED(hr)) {
        throw std::string("Unable to SetInput");
    }

    hr = m_recognizer->CreateRecoContext(&m_context);
    if (FAILED(hr)) {
        throw std::string("Failed command recognition");
    }

    hr = m_context->SetNotifyCallbackFunction(CWaveToText::NotifyCallback, reinterpret_cast<WPARAM>(this), 0);
    if (FAILED(hr)) {
        throw std::string("Unable to set notify callback function");
    }

    const ULONGLONG ullInterest =
        SPFEI(SPEI_SOUND_START) | SPFEI(SPEI_SOUND_END) |
        SPFEI(SPEI_PHRASE_START) | SPFEI(SPEI_RECOGNITION) |
        SPFEI(SPEI_FALSE_RECOGNITION) | SPFEI(SPEI_HYPOTHESIS) |
        SPFEI(SPEI_INTERFERENCE) | SPFEI(SPEI_RECO_OTHER_CONTEXT) |
        SPFEI(SPEI_REQUEST_UI) | SPFEI(SPEI_RECO_STATE_CHANGE) |
        SPFEI(SPEI_PROPERTY_NUM_CHANGE) | SPFEI(SPEI_PROPERTY_STRING_CHANGE);

    hr = m_context->SetInterest(ullInterest, ullInterest);
    if (FAILED(hr)) {
        throw std::string("Failed to create interest");
    }

    hr = m_context->CreateGrammar(0, &m_grammer);
    if (FAILED(hr)) {
        throw std::string("Unable to create grammar");
    }

    hr = m_grammer->LoadDictation(0, SPLO_STATIC);
    if (FAILED(hr)) {
        throw std::string("Failed to load dictation");
    }

    hr = m_grammer->SetDictationState(SPRS_ACTIVE);
    if (FAILED(hr)) {
        throw std::string("Failed setting dictation state");
    }

    return 0;
}

int CWaveToText::Stop()
{
    return 0;
}

void CWaveToText::Init()
{
    m_bFinishProcessing = FALSE;
    ZeroMemory(m_strFileName, sizeof(m_strFileName));

    if (FAILED(::CoInitialize(NULL))) {
        throw std::exception("Fail to initialize COM.");
    }
}

void CWaveToText::CleanUp()
{
    if (m_grammer) {
        m_grammer.Release();
    }

    if (m_context) {
        m_context->SetNotifySink(NULL);
        m_context.Release();
    }

    if (m_recognizer) {
        m_recognizer.Release();
    }

    m_bFinishProcessing = FALSE;
}

LPCTSTR CWaveToText::ExtractInput(CSpEvent& event)
{
    SPPHRASE    *pPhrase;
    TCHAR       *pwszText = NULL;

    CComPtr<ISpRecoResult> cpRecoResult = event.RecoResult();

    HRESULT hr = cpRecoResult->GetPhrase(&pPhrase);
    if (SUCCEEDED(hr)) {
        if (event.eEventId == SPEI_FALSE_RECOGNITION) {
            pwszText = _T("False recognition");
        } else {
            // Get the phrase's entire text string, including replacements.
            hr = cpRecoResult->GetText(SP_GETWHOLEPHRASE, SP_GETWHOLEPHRASE, TRUE, &pwszText, NULL);
        }
    }

    CoTaskMemFree(pPhrase);
    return pwszText;
}

void __stdcall CWaveToText::NotifyCallback(WPARAM wParam, LPARAM lParam)
{
    CWaveToText* pObj = reinterpret_cast<CWaveToText*>(wParam);

    CSpEvent event;
    LPCTSTR  lpszText;

    if (pObj->m_bFinishProcessing) {
        return;
    }

    // Loop processing events while there are any in the queue
    while (event.GetFrom(pObj->m_context) == S_OK) {
        switch (event.eEventId) {
        case SPEI_RECOGNITION:
            lpszText = pObj->ExtractInput(event);
            pObj->m_fnRecognitionCallback(pObj->m_wpCallback, pObj->m_lpCallback, lpszText);
            break;
        case SPEI_FALSE_RECOGNITION:
            break;
        case SPEI_END_SR_STREAM:
            pObj->m_bFinishProcessing = true;
            break;
        }
    }
}
