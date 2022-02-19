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


// fTemplates dialog

class fTemplates : public CDialog
{
	DECLARE_DYNAMIC(fTemplates)

public:
	fTemplates(CWnd* pParent = NULL);   // standard constructor
	virtual ~fTemplates();

	CActiveGanttVCCtl	ActiveGanttVCCtl1;
	CToolTipToolBar oToolBar1;
	CBitmap oToolBarBitmap;
	CArray<CString,CString> mp_aControlTemplates;
	CArray<CString,CString> mp_aObjectTemplates;


	CFont mp_oFont;
	CComboBox cboControlTemplates;
	CComboBox cboObjectTemplates;
	CStatic lblControlTemplates;
	CStatic lblObjectTemplates;
	void mp_SetStyles();
	void mp_CaptionRow(CString sRowKey, CString sObject, CString sStyle);

//ToolBar Messages
public:
	void cmdCopy();
	void cboControlTemplates_SelChange();
	void cboObjectTemplates_SelChange();

// Dialog Data
	enum { IDD = IDD_FTEMPLATES };

protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support

	DECLARE_MESSAGE_MAP()
public:
	virtual BOOL OnInitDialog();
	afx_msg void OnSize(UINT nType, int cx, int cy);
	afx_msg int OnCreate(LPCREATESTRUCT lpCreateStruct);
};