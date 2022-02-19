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



// fLoadXML dialog

class fLoadXML : public CDialog
{
	DECLARE_DYNAMIC(fLoadXML)

public:
	fLoadXML(CWnd* pParent = NULL);   // standard constructor
	virtual ~fLoadXML();
	HICON m_hIcon;

	CActiveGanttVCCtl	ActiveGanttVCCtl1;
	BOOL bLoaded;

	CToolTipToolBar oToolBar1;
	CBitmap oToolBarBitmap;

	void cmdLoadXML();
	void cmdSaveXML();

	void mnuLoadXML_Click();
	void mnuSaveXML_Click();
	void mnuClose_Click();

	void SaveXML();
	void LoadXML();



// Dialog Data
	enum { IDD = IDD_FLOADXML };

protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support

	DECLARE_MESSAGE_MAP()
public:
	afx_msg int OnCreate(LPCREATESTRUCT lpCreateStruct);
	virtual BOOL OnInitDialog();
	afx_msg void OnSize(UINT nType, int cx, int cy);
	DECLARE_EVENTSINK_MAP()
	void CustomTierDrawActiveganttvcctl1(LPDISPATCH e);
};