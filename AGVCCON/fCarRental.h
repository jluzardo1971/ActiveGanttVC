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

class fCarRental : public CDialog
{
	DECLARE_DYNAMIC(fCarRental)

public:
	fCarRental(CWnd* pParent = NULL);   
	virtual ~fCarRental();

	CActiveGanttVCCtl	ActiveGanttVCCtl1;

	CFont mp_oFont;
	CComboBox drpMode;

	HICON m_hIcon;
	
	int mp_lZoom;
	HPE_ADDMODE mp_yAddMode;
	int mp_lDeleteTask;
	CDatabase mp_oConn;

	CToolTipToolBar oToolBar1;
	CBitmap oToolBarBitmap;
	void OnToolBar(UINT nID);

	CString mp_sEditRowKey;
	CString mp_sEditTaskKey;
	CBrush oBrushReservation;
	CBrush oBrushRental;
	CBrush oBrushMaintenance;
	

	void Access_LoadRowsAndTasks();
	int GetZoom();
	void SetZoom(int NewValue);
	HPE_ADDMODE GetMode();
	CString GetAddModeStyle();
	CString GetDescription(LONG lCarTypeID);
	void CalculateRate(CclsTask* oTask);
	DOUBLE GetStateTax(CclsTask* oTask, CString* sState);

//ToolBar Messages
public:
	void cmdSaveXML();
	void cmdLoadXML();
	void cmdPrint_Click();
	void cmdZoomIn_Click();
	void cmdZoomOut_Click();
	void cmdAddVehicle_Click();
	void cmdAddBranch_Click();
	void cmdBack2_Click();
	void cmdBack1_Click();
	void cmdBack0_Click();
	void cmdFwd0_Click();
	void cmdFwd1_Click();
	void cmdFwd2_Click();
	void cmdHelp_Click();

	void mnuEditRow_Click();
	void mnuDeleteRow_Click();

	void mnuEditTask_Click();
	void mnuDeleteTask_Click();
	void mnuConvertToRental_Click();

//Menu Messages
public:
	void mnuSaveXML_Click();
	void mnuLoadXML_Click();
	void mnuPrint_Click();
	void mnuClose_Click();

//Common Menu/Toolbar
public:
	void SaveXML();
	void LoadXML();
	void Print();


public:



	enum { IDD = IDD_FCARRENTAL };

protected:
	virtual void DoDataExchange(CDataExchange* pDX);    

	DECLARE_MESSAGE_MAP()
public:
	virtual BOOL OnInitDialog();
	afx_msg void OnSize(UINT nType, int cx, int cy);
	afx_msg int OnCreate(LPCREATESTRUCT lpCreateStruct);
	DECLARE_EVENTSINK_MAP()
	void CustomTierDrawActiveganttvcctl1(LPDISPATCH e);
	void ObjectAddedActiveganttvcctl1(LPDISPATCH e);
	void ControlMouseDownActiveganttvcctl1(LPDISPATCH e);
	afx_msg void OnClose();
	void CompleteObjectChangeActiveganttvcctl1(LPDISPATCH e);
	void ControlMouseWheelActiveganttvcctl1(LPDISPATCH e);
	void ControlDrawActiveganttvcctl1();
};