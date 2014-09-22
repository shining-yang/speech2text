//
// File: WaveToText.h
//
// Declaration for class CWaveToText.
// Convert speech which saved in wave format audio file to text.
//
#pragma once

class CWaveToText
{
public:
    CWaveToText();
    ~CWaveToText();

private:
    CComPtr<ISpStream>          m_stream;
    CComPtr<ISpRecognizer>      m_recognizer;
    CComPtr<ISpRecoContext>     m_context;
    CComPtr<ISpRecoGrammar>     m_grammer;
    CComPtr<ISpRecoResult>      m_result;
};

