//
// File: WaveToText.h
//
// Declaration for class CWaveToText.
// Convert speech which saved in wave format audio file to text.
//
#pragma once

#include <sapi.h>
#include <string>

typedef void (*RECOGNITIONFUNC)(WPARAM wParam, LPARAM lParam, LPCTSTR lpszText);

class CWaveToText
{
public:
    CWaveToText();
    ~CWaveToText();

public:
    void SetInputWaveFile(LPCTSTR lpszFile);
    void SetRecognitionCallback(RECOGNITIONFUNC callback, WPARAM wParam, LPARAM lParam);
    int Start();
    int Stop();

protected:
    void Init();
    void CleanUp();
    LPCTSTR ExtractInput(CSpEvent& event);
    static void __stdcall NotifyCallback(WPARAM wParam, LPARAM lParam);

private:
    TCHAR   m_strFileName[MAX_PATH];
    RECOGNITIONFUNC m_fnRecognitionCallback;
    WPARAM  m_wpCallback;
    LPARAM  m_lpCallback;
    CComPtr<ISpRecognizer>      m_recognizer;
    CComPtr<ISpRecoContext>     m_context;
    CComPtr<ISpRecoGrammar>     m_grammer;
    BOOL m_bFinishProcessing;
};

