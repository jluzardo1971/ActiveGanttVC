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
#if !defined(AFX_NUMERICFORMAT_H__04FED3F4_E71F_4401_BAB9_D2F7DED1911F__INCLUDED_)
#define AFX_NUMERICFORMAT_H__04FED3F4_E71F_4401_BAB9_D2F7DED1911F__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif



#define MAX_NFORMAT_LEN 128

class xAuxCNumFrmNode;

class xAuxCNumericFormat  
{
public:
    xAuxCNumericFormat();
    xAuxCNumericFormat(const CString& sFormat);
	~xAuxCNumericFormat();

    BOOL LoadString(UINT nRes);
    BOOL SetFormat(const CString& sFormat);
    CString GetFormat() const;

    LPCTSTR PrintNumber(double nVal, CDC* pDC, LPRECT lpRect, UINT nFormat = DT_LEFT, int nForceDec = -1) const;
    LPCTSTR PrintNumberETO(double nVal, CDC* pDC, int x, int y, UINT nOptions, LPCRECT lpRect, int nForceDec = -1) const;
    LPCTSTR SPrintNumber(double nVal, BOOL bFull = FALSE) const;
    LPCTSTR SPrintNumberSimple(double nVal, int nForceDec = -1) const;

	friend CArchive& operator<<(CArchive& ar, const xAuxCNumericFormat& frmt);
	friend CArchive& operator>>(CArchive& ar, xAuxCNumericFormat& frmt);

private:
    xAuxCNumFrmNode*    m_pNodes;
    int             m_nNodeCount;
    CString		    m_sFormat;

};

#endif // !defined(AFX_NUMERICFORMAT_H__04FED3F4_E71F_4401_BAB9_D2F7DED1911F__INCLUDED_)