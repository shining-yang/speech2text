//
// File: WaveToText.h
//
// Declaration for class CWaveToText.
//
// Convert wave audio file to text messages using Microsoft SAPI.
//
// Shining Yang <y.s.n@live.com>, 2014-09-24
//
#pragma once

#include <sphelper.h>
#include <string>

typedef enum RECOG_NOTIFY_EVENT_t
{
    RECOG_STARTED,
    RECOG_ENDED,
    RECOG_SUCCESS,
} RECOG_NOTIFY_EVENT;

typedef struct RECOG_NOTIFY_DATA_t
{
    RECOG_NOTIFY_EVENT event;
    CONST VOID* data;
} RECOG_NOTIFY_DATA, *LPRECOG_NOTIFY_DATA;

typedef void(*PFUNC_RECOG_NOTIFY)(WPARAM wParam, LPARAM lParam, LPRECOG_NOTIFY_DATA lpNotifyData);

class CWaveToText
{
public:
    CWaveToText();
    ~CWaveToText();

public:
    void SetRecognitionCallback(PFUNC_RECOG_NOTIFY callback, WPARAM wParam, LPARAM lParam);
    void SetInputWaveFile(LPCTSTR lpszFile);
    int Start();
    int Stop();

protected:
    void Init();
    void CleanUp();
    LPCTSTR ExtractInput(CSpEvent& event);
    static void __stdcall NotifyCallback(WPARAM wParam, LPARAM lParam);

private:
    CComPtr<ISpRecognizer>      m_recognizer;
    CComPtr<ISpRecoContext>     m_context;
    CComPtr<ISpRecoGrammar>     m_grammer;
    PFUNC_RECOG_NOTIFY m_fnRecognitionCallback;
    TCHAR   m_strFileName[MAX_PATH];
    WPARAM  m_wpCallback;
    LPARAM  m_lpCallback;
};

