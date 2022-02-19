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
#include "MSP2003Wrappers.h"


// fMSProject11 dialog

class fMSProject11 : public CDialog
{
	DECLARE_DYNAMIC(fMSProject11)

public:
	fMSProject11(CWnd* pParent = NULL);   // standard constructor
	virtual ~fMSProject11();

	CActiveGanttVCCtl ActiveGanttVCCtl1;
	MSP2003::CMSP2003Ctl MSP2003Ctl1;
	HICON m_hIcon;
	CString mp_sFontName;
	COleDateTime mp_dtStartDate;
	COleDateTime mp_dtEndDate;

	CToolTipToolBar oToolBar1;
	CBitmap oToolBarBitmap;

	void ToolBar_Click(UINT nID);
	void cmdSaveXML();
	void cmdLoadXML();
	void cmdPrint_Click();
	void cmdIndent_Click();
	void cmdZoomIn_Click();
	void cmdZoomOut_Click();

	void mnuLoadXML_Click();
	void mnuSaveXML_Click();
	void mnuPrint_Click();
	void mnuClose_Click();

	void InitializeAG();
	void MP11_To_AG();
	ActiveGanttVC::E_CONSTRAINTTYPE GetAGPredecessorType(MSP2003::E_TYPE_3 MPPredecessorType);
	void AGSetStartDate(COleDateTime dtDate);
	

// Dialog Data
	enum { IDD = IDD_FMSPROJECT11 };

protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support

	DECLARE_MESSAGE_MAP()
public:
	afx_msg int OnCreate(LPCREATESTRUCT lpCreateStruct);
	virtual BOOL OnInitDialog();
	afx_msg void OnSize(UINT nType, int cx, int cy);
	DECLARE_EVENTSINK_MAP()
	void CustomTierDrawActiveganttvcctl1(LPDISPATCH e);
	void ControlMouseWheelActiveganttvcctl1(LPDISPATCH e);
	afx_msg void OnClose();
};