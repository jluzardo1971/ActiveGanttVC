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
#include "afxwin.h"

class fPrintDialog;

// fPrintPreview dialog

class fPrintPreview : public CDialog
{
	DECLARE_DYNAMIC(fPrintPreview)

public:
	fPrintPreview(CWnd* pParent = NULL);   // standard constructor
	virtual ~fPrintPreview();

	fPrintDialog* mp_oParent;
	CToolTipToolBar oToolBar1;
	CStatic lblPages;
	CBitmap oToolBarBitmap;
	CFont mp_oFont;
	

// Dialog Data
	enum { IDD = IDD_FPRINTPREVIEW };

protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support

	DECLARE_MESSAGE_MAP()
public:
	virtual BOOL OnInitDialog();
	afx_msg void OnSize(UINT nType, int cx, int cy);
	void PositionScrollBars();

	CScrollBar m_HScroll;
	CScrollBar m_VScroll;
	int mp_lHSB_SmallChange;
	int mp_lHSB_LargeChange;
	int mp_lVSB_SmallChange;
	int mp_lVSB_LargeChange;


	LONG mp_lPage;
	LONG mp_lColumn;
	LONG mp_lRow;
	FLOAT mp_fScale;



	int mp_lHScroll_Value;
	int mp_lVScroll_Value;

	afx_msg void OnPaint();
	CButton cmdScrollBarsSeparator;

	void mp_UpdatePageNumber();
	afx_msg void OnHScroll(UINT nSBCode, UINT nPos, CScrollBar* pScrollBar);
	afx_msg void OnVScroll(UINT nSBCode, UINT nPos, CScrollBar* pScrollBar);
	afx_msg int OnCreate(LPCREATESTRUCT lpCreateStruct);
	afx_msg BOOL OnEraseBkgnd(CDC* pDC);
	void ToolBar_Click(UINT nID);
};