// ----------------------------------------------------------------------------------------
//                              COPYRIGHT NOTICE
// ----------------------------------------------------------------------------------------
//
// The Source Code Store LLC
// ACTIVEGANTT SCHEDULER COMPONENT FOR C++ - ActiveGanttVC
// ActiveX Control
// Copyright (c) 2002-2017 The Source Code Store LLC
//
// All Rights Reserved. No parts of this file may be reproduced, modified or transmitted 
// in any form or by any means without the written permission of the author.
//
// ----------------------------------------------------------------------------------------
#include "stdafx.h"
#include "Afxtempl.h"
#include "ActiveGanttVC.h"
#include "NumericFormat.h"
#include <math.h>
#ifdef _UNICODE
#include <afxpriv.h>        // For A2W & W2A conversions
#endif

#ifdef _DEBUG
#undef THIS_FILE
static char THIS_FILE[]=__FILE__;
#define new DEBUG_NEW
#endif

// Format characters definition
#define FRM_OPEN_BRACKET	_T('[')
#define FRM_CLOSE_BRACKET	_T(']')
#define FRM_PURGE_SIGN		_T('~')
#define FRM_DEC_POINT	    _T('.')
#define FRM_SECTION_DELIM	_T(';')

#define FRM_DIGIT           _T('#')
#define FRM_DIGIT_ZERO      _T('0')
#define FRM_DIGIT_BALNKS    _T('?')
#define FRM_THOUSAND_SEP    _T(',')
#define FRM_PERSENT			_T('%')

#define OUTCHAR_BLANK          _T(' ')
#define INTL_GROUP_DIGIT_LEN   3
#define INTL_NEGATIVE_STR      _T("-")


class xAuxCfmtCondition
{
public:
	xAuxCfmtCondition();
	xAuxCfmtCondition(int nSign, double nLimit);
	BOOL    Satisfied(double nValue) const;
private:
	double m_nLimit;
	int    m_nSign;
};

typedef CArray<xAuxCfmtCondition, xAuxCfmtCondition&> xAuxCfmtConditionsArray;

//_________________________________________________________________
//
// xAuxCNumFrmNode private helper class defenition
//
class xAuxCNumFrmNode
{
public:
    xAuxCNumFrmNode();

    LPCTSTR PrintSimpleNumFormat(double nOut, int nForceDec) const;
    LPCTSTR PrintNumFormat(double nVal, CDC* pDC, LPRECT lpRect, UINT nFormat, int nForceDec, BOOL bFull) const;
    LPCTSTR PrintNumFormatETO(double nVal, CDC* pDC, int x, int y, UINT nOptions, LPCRECT lpRect, int nForceDec, BOOL bFull) const;

	BOOL Satisfied(double nValue) const;
	BOOL IsConditional() const;

    BOOL SetFormat(const CString& sFormat);
	void SetCondition(int nSign, double nLimit);

    BOOL IsEmpty() const { return m_sFormat.IsEmpty(); };
	void Reset();

private:
	BOOL GobbleText1(CString& sFrm);
	void ProcessPower(LPCTSTR pBuff);
	static void ProcessNumFlags(LPCTSTR pBuff, int& nFlags);
	BOOL GobbleInt(CString& sFrm);
	BOOL ExtractColorOrCondition(CString& sFmt);
	BOOL ExtractPercent(CString& sFmt);
	BOOL ExtractPurgeSign(CString& sFmt);

	BOOL RecognizedAsColor(const CString& sItem);
	BOOL RecognizedAsCondition(const CString& sItem);

	void BeforePrint(double nVal, CDC* pDC, LPCRECT lpRect, int nForceDec, BOOL bFull) const;
	void AfterPrint(CDC* pDC, LPCRECT lpRect) const;

    static BOOL WinFormat(LPTSTR szOutput, int nOutBufferLen, double nNumb, int nDec, 
                          BOOL bThousandSepareted, BOOL bIntZeroObligated);

    enum EDigitFlags
    {
        DF_DIGIT        = 1,
        DF_DIGIT_ZERO   = 2,
        DF_DIGIT_BALNKS = 4,
        DF_THOUSANDS    = 8,
    };

    enum ECommonFlags
    {
        CF_USE_COLOR    = 1,    
        CF_USE_POWER    = 2,
        CF_USE_PERCENT  = 4,
        CF_PURGE_SIGN   = 8,
    };

    CString     m_sFormat;
    double      m_nPowered;
    int         m_nDec;
    int         m_nLeftFlags;
    int         m_nRightFlags;
    int         m_nCommonFlags;
    COLORREF    m_Color;

    CString    m_sText1;
    CString    m_sText2;

	xAuxCfmtConditionsArray m_arConditions;

    static TCHAR m_pTmpBuffer[MAX_NFORMAT_LEN + 1];
    static BOOL  m_bLocaleInitialized;
    static CString m_sThousand;
    static CString m_sDecimal;
};

//_________________________________________________________________
//
TCHAR   xAuxCNumFrmNode::m_pTmpBuffer[MAX_NFORMAT_LEN + 1];
BOOL    xAuxCNumFrmNode::m_bLocaleInitialized = FALSE;
CString xAuxCNumFrmNode::m_sThousand;
CString xAuxCNumFrmNode::m_sDecimal;

//_________________________________________________________________
//
// Enumeration of comparison operations
// and array of corresponding strings
//
enum ESigns
{
    SGN_EQUAL,
    SGN_NOEQUAL,
    SGN_LESS,
    SGN_GREATER,
    SGN_LESS_OR_EQUAL_1,
    SGN_LESS_OR_EQUAL_2,
    SGN_GREATER_OR_EQUAL_1,
    SGN_GREATER_OR_EQUAL_2,
};

static LPCTSTR szComparisonSigns[] = 
{
	_T("="),
	_T("#"),
	_T("<"),
	_T(">"),
	_T("<="),
	_T("=<"),
	_T(">="),
	_T("=>"),
	NULL,
};

//_________________________________________________________________
//
// Helper functions
//
inline static BOOL NonZero(double nValue)
{
    const double nMinBound = -1.0e-100;
    const double nMaxBound = 1.0e-100;
    return nValue > nMaxBound || nValue < nMinBound;
}

static CString Extract(CString& str, int nStart, int nFinish)
{
    CString sRet;
    if(nStart <= nFinish && nStart < str.GetLength())
    {
        sRet = str.Mid(nStart, nFinish - nStart + 1);
        str  = str.Mid(0, nStart) + str.Mid(nFinish + 1);
    }
    return sRet;
}

static CString GetLine(CString& str, TCHAR pchDelim)
{
    CString sRetStr;
    int     pos;
    pos = str.Find(pchDelim);
    if(pos != -1)
    {
        sRetStr = str.Left(pos);
        str = str.Mid(pos + 1);
    }
    else
    {
        sRetStr = str;
        str.Empty();
    }
    return sRetStr;
}

//_________________________________________________________________
//
// Constructor without initialization 
// is used only in CArray::SetSize() 
// (CArray::ConstructElements call).
//
xAuxCfmtCondition::xAuxCfmtCondition()
{
}

xAuxCfmtCondition::xAuxCfmtCondition(int nSign, double nLimit)
{
	m_nSign  = nSign;
	m_nLimit = nLimit;
}

BOOL xAuxCfmtCondition::Satisfied(double nValue) const
{
	BOOL bResult;
	switch(m_nSign)
	{
		case SGN_EQUAL:
			bResult = !NonZero(m_nLimit - nValue);
			break;
		case SGN_NOEQUAL:
			bResult = NonZero(m_nLimit - nValue);
			break;
		case SGN_LESS:
			bResult = nValue < m_nLimit && NonZero(m_nLimit - nValue);
			break;
		case SGN_GREATER:
			bResult = nValue > m_nLimit && NonZero(m_nLimit - nValue);
			break;
		case SGN_LESS_OR_EQUAL_1:
		case SGN_LESS_OR_EQUAL_2:
			bResult = nValue < m_nLimit || !NonZero(m_nLimit - nValue);
			break;
		case SGN_GREATER_OR_EQUAL_1:
		case SGN_GREATER_OR_EQUAL_2:
			bResult = nValue > m_nLimit || !NonZero(m_nLimit - nValue);
			break;
		default:
			ASSERT(FALSE);
	}
	return bResult;
}

//_________________________________________________________________
//
xAuxCNumFrmNode::xAuxCNumFrmNode()
{
    Reset();
}

BOOL xAuxCNumFrmNode::SetFormat(const CString& sFormat)
{
    if(sFormat != m_sFormat)
    {
        Reset();

        if( sFormat.IsEmpty() )
            return TRUE;

        if( sFormat.GetLength() > MAX_NFORMAT_LEN )
            return FALSE;

        CString sFrm = sFormat;

        if(!ExtractColorOrCondition(sFrm))  return FALSE;
        if(!ExtractPercent(sFrm))			return FALSE;
        if(!ExtractPurgeSign(sFrm))			return FALSE;

        if(!GobbleText1(sFrm))  return FALSE;
        if(!GobbleInt(sFrm))    return FALSE;

        m_sText2 = sFrm;

        m_sFormat = sFormat;
    }
    return TRUE;
}

void xAuxCNumFrmNode::Reset()
{
    m_sFormat.Empty();
    m_nPowered = 1.0;
    m_nDec           = 0;
    m_nLeftFlags     = 0;
    m_nRightFlags    = 0;
    m_nCommonFlags   = 0;
    m_Color = RGB(0,0,0);
}

BOOL xAuxCNumFrmNode::Satisfied(double nValue) const
{
    if(m_arConditions.GetSize())
	{
		for(int i = 0; i < m_arConditions.GetSize(); i++)
		{
            if(!m_arConditions[i].Satisfied(nValue))
				return FALSE;
		}
	}
	return TRUE;
}

void xAuxCNumFrmNode::SetCondition(int nSign, double nLimit)
{
	xAuxCfmtCondition cond(nSign, nLimit);
	m_arConditions.Add(cond);
}

BOOL xAuxCNumFrmNode::IsConditional() const
{
	return m_arConditions.GetSize() ? TRUE : FALSE;
}

BOOL xAuxCNumFrmNode::ExtractColorOrCondition(CString& sFmt)
{
    int  pos_open;
    int  pos_close;
	CString sItem;
    while( (pos_open = sFmt.Find(FRM_OPEN_BRACKET))  != -1 && 
           (pos_close = sFmt.Find(FRM_CLOSE_BRACKET)) != -1 && 
            pos_open < pos_close                 )
	{
	    sItem = Extract(sFmt, pos_open, pos_close);
		if(!RecognizedAsColor(sItem))
			RecognizedAsCondition(sItem);
	}
	return TRUE;
}

BOOL xAuxCNumFrmNode::RecognizedAsColor(const CString& sItem)
{
    static struct FmtColorMap 
    {
        LPCTSTR pColorName;
        COLORREF cColor;
    }pColorMap[] =
	{
		{_T("[Green]"),RGB(0,255,0)},
		{_T("[Red]"),RGB(255,0,0)},
		{ _T("[Blue]"),RGB(0,0,255)},
		{ _T("[Magenta]"),RGB(255,0,255)},
		{ _T("[Cyan]"),RGB(0,255,255)},
		{ _T("[Yellow]"),RGB(255,255,0)},
		{ _T("[White]"),RGB(255,255,255)},
		{ _T("[Black]"),RGB(0,0,0)},
		{NULL,0}
	};


    FmtColorMap* pMap = pColorMap;
    while(pMap->pColorName)
    {
        if(sItem.CompareNoCase(pMap->pColorName) == 0)
        {
            m_nCommonFlags |= CF_USE_COLOR;
            m_Color = pMap->cColor;
            return TRUE;
        }
        pMap++;
    }
    return FALSE;
}

BOOL xAuxCNumFrmNode::RecognizedAsCondition(const CString& sItem)
{
	CString	sTmp;	
	CString	sSign;	
	int		i;

	sTmp = sItem.Mid(1, sItem.GetLength() - 2);

	for(i = 0; i < sTmp.GetLength() && !(_istdigit(sTmp[i])) && sTmp[i] != _T('-'); i++)
		;

	if(i > 0 && i < sTmp.GetLength())
	{
		sSign = sTmp.Mid(0, i);
		sSign.TrimLeft();
		sSign.TrimRight();

		// look
		for(int j = 0; szComparisonSigns[j] != NULL; j++)
		{
			if(szComparisonSigns[j] == sSign)
			{
                double nValue;
#ifdef _UNICODE
                USES_CONVERSION;
                LPCSTR  pAStr;
                LPCTSTR pWStr = ((LPCTSTR)sTmp) + i;
                pAStr = W2CA(pWStr);
                nValue = atof(pAStr);
#else
                nValue = atof(((LPCTSTR)sTmp) + i);
#endif

				xAuxCfmtCondition cond(j, nValue);
				m_arConditions.Add(cond);
				return TRUE;
			}
		}
	}
	return FALSE;
}

BOOL xAuxCNumFrmNode::ExtractPercent(CString& sFmt)
{
    int pos = sFmt.Find(FRM_PERSENT);
    if(pos != -1)
    {
        Extract(sFmt, pos, pos);
        m_nCommonFlags |= CF_USE_PERCENT;
        m_nPowered = m_nPowered * 100;
    }
    return TRUE;
}

BOOL xAuxCNumFrmNode::ExtractPurgeSign(CString& sFmt)
{
    int pos = sFmt.Find(FRM_PURGE_SIGN);
    if(pos != -1)
    {
        Extract(sFmt, pos, pos);
        m_nCommonFlags |= CF_PURGE_SIGN;
    }
    return TRUE;
}

BOOL xAuxCNumFrmNode::GobbleText1(CString& sFrm)
{
    // Declaration of format string for sscanf
    // _T("%[^#0?,.]")
    static const TCHAR szSScanfFormat[] = 
    {   _T('%'), _T('['), _T('^'), 
        FRM_DIGIT, FRM_DIGIT_ZERO, FRM_DIGIT_BALNKS,
        FRM_THOUSAND_SEP, FRM_DEC_POINT, 
        _T(']'), 0 };

    m_pTmpBuffer[0] = 0;

#if _MSC_VER <= 1310
	_stscanf(sFrm, szSScanfFormat, m_pTmpBuffer);
#else
	_stscanf_s(sFrm, szSScanfFormat, m_pTmpBuffer, sizeof(sFrm));
#endif

    
    m_sText1 = m_pTmpBuffer;
    if(m_sText1.GetLength())
        sFrm = sFrm.Mid(m_sText1.GetLength());
    return TRUE;
}

BOOL xAuxCNumFrmNode::GobbleInt(CString& sFrm)
{
    // Declaration of format string for sscanf  1
    // _T("%[#0?,]")
    static const TCHAR szSScanfFormat1[] = 
    {   _T('%'), _T('['), 
        FRM_DIGIT, FRM_DIGIT_ZERO, FRM_DIGIT_BALNKS,
        FRM_THOUSAND_SEP,
        _T(']'), 0 };

    // Declaration of format string for sscanf  2
    // _T("%[#0?]")
    static const TCHAR szSScanfFormat2[] = 
    {   _T('%'), _T('['),
        FRM_DIGIT, FRM_DIGIT_ZERO, FRM_DIGIT_BALNKS,
        _T(']'), 0 };

    // Declaration of format string for sscanf  3
    // _T("%[,]")
    static const TCHAR szSScanfFormat3[] = 
    {   _T('%'), _T('['),
        FRM_THOUSAND_SEP,
        _T(']'), 0 };

    int nLen;

    m_nDec = 0;
    m_pTmpBuffer[0] = 0;


#if _MSC_VER <= 1310
	_stscanf(sFrm, szSScanfFormat1, m_pTmpBuffer);      // INTEGER PART
#else
	_stscanf_s(sFrm, szSScanfFormat1, m_pTmpBuffer, sizeof(sFrm));      // INTEGER PART
#endif

    

    if(nLen = lstrlen(m_pTmpBuffer))
        sFrm = sFrm.Mid(nLen);

    ProcessNumFlags(m_pTmpBuffer, m_nLeftFlags);

    if(sFrm.GetLength() && sFrm[0] == FRM_DEC_POINT)
    {
        sFrm = sFrm.Mid(1);
        m_pTmpBuffer[0] = 0;


#if _MSC_VER <= 1310
	_stscanf(sFrm, szSScanfFormat2, m_pTmpBuffer);  // DEC PART
#else
	_stscanf_s(sFrm, szSScanfFormat2, m_pTmpBuffer, sizeof(sFrm));  // DEC PART
#endif

        
        ProcessNumFlags(m_pTmpBuffer, m_nRightFlags);
        m_nDec = lstrlen(m_pTmpBuffer);

        if(m_nDec)
            sFrm = sFrm.Mid(m_nDec);

        m_pTmpBuffer[0] = 0;

#if _MSC_VER <= 1310
	_stscanf(sFrm, szSScanfFormat3, m_pTmpBuffer);  // POWER PART
#else
	_stscanf_s(sFrm, szSScanfFormat3, m_pTmpBuffer, sizeof(sFrm));  // POWER PART
#endif


        

        if(nLen = lstrlen(m_pTmpBuffer))
            sFrm = sFrm.Mid(nLen);
    }
    ProcessPower(m_pTmpBuffer);

    return TRUE;    
}

void xAuxCNumFrmNode::ProcessNumFlags(LPCTSTR pBuff, int& nFlags)
{
    if(_tcschr(pBuff, FRM_DIGIT))           nFlags |= DF_DIGIT;
    if(_tcschr(pBuff, FRM_DIGIT_ZERO))      nFlags |= DF_DIGIT_ZERO;
    if(_tcschr(pBuff, FRM_DIGIT_BALNKS))    nFlags |= DF_DIGIT_BALNKS;
    if(_tcschr(pBuff, FRM_THOUSAND_SEP))    nFlags |= DF_THOUSANDS;
}

void xAuxCNumFrmNode::ProcessPower(LPCTSTR pBuff)
{
    int nLen = lstrlen(pBuff);
    int i;
    if(nLen)
    {
        int nPowers = 0;
        for(i = nLen - 1; i >= 0; i--)
        {
            if(pBuff[i] == FRM_THOUSAND_SEP)
                nPowers++;
            else
                break;
        }
        if(nPowers)
        {
            m_nCommonFlags |= CF_USE_POWER;
            m_nPowered = m_nPowered / pow(10.0, (double)(nPowers * INTL_GROUP_DIGIT_LEN));
        }
    }
}

static COLORREF colOldColor;

void xAuxCNumFrmNode::BeforePrint(double nVal, CDC* pDC, LPCRECT lpRect, int nForceDec, BOOL bFull) const
{
    int         nLenBuffer;
    int         nNumDec;

    m_pTmpBuffer[0] = 0;
    
    if(nForceDec != -1)
        nNumDec = nForceDec;
    else
        nNumDec = m_nDec;

    if(bFull)   // BEFORE MAIN NUMBER
    {
        if(m_nCommonFlags & (CF_USE_POWER | CF_USE_PERCENT))
            nVal = nVal * m_nPowered;

        lstrcpy(m_pTmpBuffer, m_sText1);
    }

    if( (m_nLeftFlags  & (DF_DIGIT | DF_DIGIT_BALNKS | DF_DIGIT_ZERO ) ) |
        (m_nRightFlags & (DF_DIGIT | DF_DIGIT_BALNKS | DF_DIGIT_ZERO ) )   )
    {
        if(bFull)
		{
            WinFormat(m_pTmpBuffer + m_sText1.GetLength(), 
                MAX_NFORMAT_LEN - m_sText1.GetLength() - m_sText2.GetLength() - 1,
                nVal, nNumDec, m_nLeftFlags & DF_THOUSANDS,
                m_nLeftFlags & DF_DIGIT_ZERO);

			if( (m_nCommonFlags & CF_PURGE_SIGN) && nVal < 0.0 )
			{
				PTCHAR pch;
				pch =  _tcsstr(m_pTmpBuffer + m_sText1.GetLength(), INTL_NEGATIVE_STR);
				if(pch)
					memmove(pch, pch + 1, (lstrlen(pch + 1) + 1) * sizeof(TCHAR));
			}
		}
        else
            WinFormat(m_pTmpBuffer, MAX_NFORMAT_LEN - 1,
                nVal, nNumDec, m_nLeftFlags & DF_THOUSANDS, m_nLeftFlags & DF_DIGIT_ZERO );
    }

    if(m_nDec)
    {
        int i;
        if(m_nRightFlags & DF_DIGIT)
        {
            nLenBuffer = lstrlen(m_pTmpBuffer);
            for(i = nLenBuffer - 1; i >= 0; i--)
            {
                if(m_pTmpBuffer[i] == _T('0'))
                    m_pTmpBuffer[i] = 0;
                else
                    break;
            }
        }
        //if(m_nRightFlags & DF_DIGIT_ZERO)
        //{ /* Nothing to do */ }

        if(m_nRightFlags & DF_DIGIT_BALNKS)
        {
            nLenBuffer = lstrlen(m_pTmpBuffer);
            for(i = nLenBuffer - 1; i >= 0; i--)
            {
                if(m_pTmpBuffer[i] == _T('0'))
                    m_pTmpBuffer[i] = OUTCHAR_BLANK;
                else
                    break;
            }
        }
    }
    
    if(bFull)   // AFTER MAIN NUMBER
    {
        lstrcat(m_pTmpBuffer, m_sText2);

        if( m_nCommonFlags & CF_USE_PERCENT )
        {
            int nCurrLength = lstrlen(m_pTmpBuffer);
            m_pTmpBuffer[nCurrLength] = FRM_PERSENT;
            m_pTmpBuffer[nCurrLength + 1] = _T('\0');
        }
    }

    if(pDC && lpRect && (m_nCommonFlags & CF_USE_COLOR))
        colOldColor = pDC->SetTextColor(m_Color);
}

void xAuxCNumFrmNode::AfterPrint(CDC* pDC, LPCRECT lpRect) const
{
    if(pDC && lpRect && (m_nCommonFlags & CF_USE_COLOR))
        pDC->SetTextColor(colOldColor);
}

LPCTSTR xAuxCNumFrmNode::PrintNumFormat(double nVal, CDC* pDC, LPRECT lpRect, UINT nFormat, int nForceDec, BOOL bFull) const
{
	BeforePrint(nVal, pDC, lpRect, nForceDec, bFull);
    if(pDC && lpRect)
        pDC->DrawText(m_pTmpBuffer, -1, lpRect, nFormat);

    AfterPrint(pDC, lpRect);
    return (LPCTSTR)m_pTmpBuffer;
}

LPCTSTR xAuxCNumFrmNode::PrintNumFormatETO(double nVal, CDC* pDC, int x, int y, UINT nOptions, LPCRECT lpRect, int nForceDec, BOOL bFull) const
{
	BeforePrint(nVal, pDC, lpRect, nForceDec, bFull);
    if(pDC && lpRect)
		pDC->ExtTextOut(x, y, nOptions, lpRect, m_pTmpBuffer, lstrlen(m_pTmpBuffer), NULL );

    AfterPrint(pDC, lpRect);
    return (LPCTSTR)m_pTmpBuffer;
}


LPCTSTR xAuxCNumFrmNode::PrintSimpleNumFormat(double nOut, int nForceDec) const
{
    CString sSprintfFormat;

    m_pTmpBuffer[0] = 0;
    
    sSprintfFormat.Format(_T("%%16.%dlf"), 
        nForceDec == -1 ? m_nDec : nForceDec);

	#if _MSC_VER <= 1310
		_stprintf(m_pTmpBuffer, sSprintfFormat, nOut);
	#else
		_stprintf_s(m_pTmpBuffer, sSprintfFormat, nOut);
	#endif
    return (LPCTSTR)m_pTmpBuffer;
}

BOOL xAuxCNumFrmNode::WinFormat(LPTSTR szOutput, int nOutBufferLen, double nNumb, int nDec, BOOL bThousandSepareted, BOOL bIntZeroObligated)
{
    const int nBufferLen = 16;
    
    if(!m_bLocaleInitialized)
    {
        // Locale initialzation
        if(!GetLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_STHOUSAND, m_sThousand.GetBuffer(nBufferLen), nBufferLen))
            m_sThousand = _T(" ");
        m_sThousand.ReleaseBuffer();
        
        if(!GetLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_SDECIMAL, m_sDecimal.GetBuffer(nBufferLen), nBufferLen))
            m_sDecimal = _T(".");
        m_sDecimal.ReleaseBuffer();
        
        m_bLocaleInitialized = TRUE;
    }
    
    int         nPosDecimal;
    int         nSign;
    
    LPTSTR      pfcvtBuffer = NULL;
    LPTSTR      pPos;
    
    if( fabs(nNumb) > 1.0e+030 )
    {
        lstrcpy(szOutput, _T("Overflow"));
        return FALSE;
    }
    
	#ifdef _UNICODE
		#if _MSC_VER <= 1300
			LPCSTR pAStr = fcvt(nNumb, nDec, &nPosDecimal, &nSign );
			pfcvtBuffer = A2W(pAStr);
		#else
			char * pAStr = new char[15];
			_fcvt_s(pAStr, 15, nNumb, nDec, &nPosDecimal, &nSign );
			USES_CONVERSION;
			pfcvtBuffer = A2W(pAStr);
			delete [] pAStr;
			pAStr = NULL;
		#endif
	#else
		#if _MSC_VER <= 1310
			pfcvtBuffer = fcvt(nNumb, nDec, &nPosDecimal, &nSign );
		#else
			_fcvt_s(pfcvtBuffer, 50, nNumb, nDec, &nPosDecimal, &nSign );
		#endif
	#endif
    
    int nfcvtLen = lstrlen(pfcvtBuffer);
    
    // Calculating approximate length of result
    int nEstimatedLength = 
        (nSign ? lstrlen(INTL_NEGATIVE_STR) : 0) +
        max(nfcvtLen, nDec + 1) +
        (nDec ? m_sDecimal.GetLength() : 0) +
        (bThousandSepareted ? ((nPosDecimal / INTL_GROUP_DIGIT_LEN) * m_sThousand.GetLength()) : 0);
    
    if(nEstimatedLength > nOutBufferLen)
    {
        lstrcpy(szOutput, _T("Overflow"));
        return FALSE;
    }
    
    pPos = szOutput;
    if(nSign && NonZero(nNumb))
    {
        lstrcpy(pPos, INTL_NEGATIVE_STR);
        pPos += lstrlen(INTL_NEGATIVE_STR);
    }
    
    // zero length
    if(nPosDecimal <= 0 && bIntZeroObligated)
        *pPos++ = _T('0');

    if(nPosDecimal < 0 && nDec > 0)
    {
        lstrcpy(pPos, m_sDecimal);
        pPos += m_sDecimal.GetLength();
        for(int pos = nPosDecimal; pos < 0 && (pos - nPosDecimal) < nDec; pos++)
            *pPos++ = _T('0');
    }

    for(int i = 0; i < nfcvtLen; i++)
    {
        // Check for thousand separator insertion
        if( i > 0 && i < nPosDecimal && 
            bThousandSepareted && 
            (nPosDecimal - i) % INTL_GROUP_DIGIT_LEN == 0)
        {
            lstrcpy(pPos, m_sThousand);
            pPos += m_sThousand.GetLength();
        }
        // Check for decimal separator insertion
        else if(i == nPosDecimal)
        {
            lstrcpy(pPos, m_sDecimal);
            pPos += m_sDecimal.GetLength();
        }
        *pPos++ = pfcvtBuffer[i];
    }
    *pPos = _T('\0');
    return TRUE;
}

//_________________________________________________________________
//
xAuxCNumericFormat::xAuxCNumericFormat()
{
    m_pNodes        = NULL;
    m_nNodeCount    = 0;
}

xAuxCNumericFormat::xAuxCNumericFormat(const CString& sFormat)
{
    m_pNodes        = NULL;
    m_nNodeCount    = 0;

    SetFormat(sFormat);
}

xAuxCNumericFormat::~xAuxCNumericFormat()
{
    if(m_pNodes)
	{
        delete [] m_pNodes;
		m_pNodes = NULL;
	}
}

CString xAuxCNumericFormat::GetFormat() const
{
    return m_sFormat;
}

BOOL xAuxCNumericFormat::SetFormat(const CString& sFormat)
{
    if(sFormat.IsEmpty())
        return FALSE;

    if(sFormat != m_sFormat)
    {
        CString sFrm = sFormat;

        // Check for old style format (sprintf style),
        // and automatic convertion into new style if 
        // nessesory.
        if(sFrm[0] == _T('%') && sFrm.Right(2) == _T("lf"))
        {
            int pos = sFrm.Find(_T('.'));
            int nNumDec;
            if(pos != -1)
            {
                nNumDec = _ttoi(&(((LPCTSTR)sFrm)[pos + 1]));

                sFrm = CString(FRM_DIGIT) + FRM_THOUSAND_SEP;
                sFrm += CString(FRM_DIGIT, INTL_GROUP_DIGIT_LEN);

                if(nNumDec)
                {
                    sFrm += FRM_DEC_POINT;
                    sFrm += CString(FRM_DIGIT_ZERO, nNumDec);
                }
            }
        }

        CString sNode;
		int  i;

		// Calculate number of sections in new format string
		m_nNodeCount = 0;

		for(i = 0; i < sFrm.GetLength(); i++)
		{
			if(sFrm[i] == FRM_SECTION_DELIM) 
				m_nNodeCount++;
		}
		if(sFrm[sFrm.GetLength() - 1] != FRM_SECTION_DELIM)
			m_nNodeCount++;

        // Obligatory clean up old content of array
        if(m_pNodes)
		{
            delete [] m_pNodes;
			m_pNodes = NULL;
		}

        m_pNodes = new xAuxCNumFrmNode[m_nNodeCount];
        if(!m_pNodes)
            return FALSE;

		BOOL bConditionalFormat = FALSE;
		for(i = 0; i < m_nNodeCount; i++)
		{
	        sNode = GetLine(sFrm, FRM_SECTION_DELIM);
		    if( !(m_pNodes[i].SetFormat(sNode)) )
				return FALSE;

			if(m_pNodes[i].IsConditional())
				bConditionalFormat = TRUE;
		}

		// Set up default conditions (if requered)
		if(!bConditionalFormat && m_nNodeCount > 1)
		{
			if(m_nNodeCount == 2)
			{
				m_pNodes[0].SetCondition(SGN_GREATER_OR_EQUAL_1, 0.0);
			}
			else /* m_nNodeCount >= 3 */
			{
				m_pNodes[0].SetCondition(SGN_GREATER, 0.0);
				m_pNodes[1].SetCondition(SGN_LESS,    0.0);
			}
		}

        m_sFormat = sFormat;
    }
    return TRUE;
}

LPCTSTR xAuxCNumericFormat::SPrintNumber(double nVal, BOOL bFull) const
{
    LPCTSTR lpszBuffer = NULL; 
    if(bFull)
        lpszBuffer = PrintNumber(nVal, NULL, NULL, 0, -1);
    else
        lpszBuffer = m_pNodes[0].PrintNumFormat(nVal, NULL, NULL, 0, -1, bFull);
    return lpszBuffer;
}

LPCTSTR xAuxCNumericFormat::SPrintNumberSimple(double nOut, int nForceDec) const
{
    LPCTSTR lpszBuffer = NULL; 
    if(!m_sFormat.IsEmpty())
    {
        lpszBuffer = m_pNodes[0].PrintSimpleNumFormat(nOut, nForceDec);
    }
    else
        TRACE(_T("xAuxCNumericFormat: Attempt to print on noninitialized format !\n"));
    return lpszBuffer;
}

LPCTSTR xAuxCNumericFormat::PrintNumber(double nVal, CDC* pDC, LPRECT lpRect, UINT nFormat, int nForceDec) const
{
    LPCTSTR lpszBuffer = NULL; 
    if(!m_sFormat.IsEmpty() && m_pNodes != NULL)
    {
		for(int i = 0; i < m_nNodeCount; i++)
		{
			if(m_pNodes[i].Satisfied(nVal))
			{
	            lpszBuffer = m_pNodes[i].PrintNumFormat(nVal, pDC, lpRect, nFormat, nForceDec, TRUE);
				return lpszBuffer;
			}
		}
    }
    else
        TRACE(_T("xAuxCNumericFormat: Attempt to print by uninitialized format !\n"));
    return lpszBuffer;
}

LPCTSTR xAuxCNumericFormat::PrintNumberETO(double nVal, CDC* pDC, int x, int y, UINT nOptions, LPCRECT lpRect, int nForceDec) const
{
    LPCTSTR lpszBuffer = NULL; 
    if(!m_sFormat.IsEmpty() && m_pNodes != NULL)
    {
		for(int i = 0; i < m_nNodeCount; i++)
		{
			if(m_pNodes[i].Satisfied(nVal))
			{
	            lpszBuffer = m_pNodes[i].PrintNumFormatETO(nVal, pDC, x, y, nOptions, lpRect, nForceDec, TRUE);
				return lpszBuffer;
			}
		}
    }
    else
        TRACE(_T("xAuxCNumericFormat: Attempt to print by uninitialized format !\n"));
    return lpszBuffer;
}


CArchive& operator<<(CArchive& ar, const xAuxCNumericFormat& frmt)
{ 
    return ar << frmt.m_sFormat; 
}

CArchive& operator>>(CArchive& ar, xAuxCNumericFormat& frmt)
{
    CString sFormat; 
    ar >> sFormat; 
    frmt.SetFormat(sFormat); 
    return ar; 
}

BOOL xAuxCNumericFormat::LoadString(UINT nRes)
{
    CString sStr;
    if(sStr.LoadString(nRes))
        return SetFormat(sStr); 
    else
        return FALSE;
}