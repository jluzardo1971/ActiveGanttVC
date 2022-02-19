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
#pragma once


class fAbout : public CDialog
{
public:
	fAbout(LPCTSTR sVersion, CWnd* pParent = NULL);

	//{{AFX_DATA(fAbout)
	enum { IDD = IDD_ABOUT };
	//}}AFX_DATA


	//{{AFX_VIRTUAL(fAbout)
	protected:
	virtual void DoDataExchange(CDataExchange* pDX);
	//}}AFX_VIRTUAL

protected:
	//{{AFX_MSG(fAbout)
	afx_msg void OnLButtonDown(UINT nFlags, CPoint point);
	virtual BOOL OnInitDialog();
	afx_msg void OnMouseMove(UINT nFlags, CPoint point);
	afx_msg void OnPaint();
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
private:
	CString m_sVersion;
	HCURSOR mp_hBrowserCursor;
	HICON m_hIcon;
	RECT mp_udtURL;
	RECT mp_udtStore;
	RECT mp_udtSupport;
	RECT mp_udtSales;
};