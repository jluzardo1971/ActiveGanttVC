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
#include "afxdb.h"

class fWBSProject : public CDialog
{
	DECLARE_DYNAMIC(fWBSProject)

public:
	fWBSProject(CWnd* pParent = NULL);   
	virtual ~fWBSProject();




// Member Variables
public:

	CActiveGanttVCCtl	ActiveGanttVCCtl1;
	HICON m_hIcon;
	CString mp_sFontName;
	COleDateTime mp_dtStartDate;
	COleDateTime mp_dtEndDate;
	CToolTipToolBar oToolBar1;
	CBitmap oToolBarBitmap;
	CDatabase mp_oConn;
	BOOL mp_bBluePercentagesVisible;
	BOOL mp_bGreenPercentagesVisible;
	BOOL mp_bRedPercentagesVisible;
	
	
// Form Events & Properties
public:
	afx_msg int OnCreate(LPCREATESTRUCT lpCreateStruct);
	virtual BOOL OnInitDialog();
	afx_msg void OnSize(UINT nType, int cx, int cy);

// ActiveGantt Messages
public:
	void CompleteObjectChangeActiveganttvcctl1(LPDISPATCH e);
	void ControlMouseDownActiveganttvcctl1(LPDISPATCH e);
	void ControlMouseWheelActiveganttvcctl1(LPDISPATCH e);
	void NodeCheckedActiveganttvcctl1(LPDISPATCH e);
	void ObjectAddedActiveganttvcctl1(LPDISPATCH e);
	void ToolTipOnMouseHoverActiveganttvcctl1(LPDISPATCH e);
	void OnMouseHoverToolTipDrawActiveganttvcctl1(LPDISPATCH e);
	void ToolTipOnMouseMoveActiveganttvcctl1(LPDISPATCH e);
	void OnMouseMoveToolTipDrawActiveganttvcctl1(LPDISPATCH e);
	void TaskToolTipCalculateDim(CToolTipEventArgs* oE);
	void TaskToolTipDraw(CToolTipEventArgs* oE);

//Operations
	void InsertPredecessor(LONG PredecessorObjectIndex);
	void UpdateTask(LONG Index);
	void UpdateSummary(CclsNode* oNode);
	CclsTask GetTaskByRowKey(CString sRowKey);
	CclsPercentage GetPercentageByTaskKey(CString sTaskKey);
	void SetTaskGridColumns(CclsTask* oTask);

//ToolBar Messages
public:
	void cmdLoadXML();
	void cmdSaveXML();
	void cmdPrint_Click();
	void cmdZoomIn_Click();
	void cmdZoomOut_Click();
	void cmdBluePercentages_Click();
	void cmdGreenPercentages_Click();
	void cmdRedPercentages_Click();
	void cmdProperties_Click();
	void cmdCheckPredecessors_Click();
	void cmdToolTips_Click();
	void cmdHelp_Click();

//Menu Messages
public:
	void mnuSaveXML_Click();
	void mnuLoadXML_Click();
	void mnuPrint_Click();
	void mnuClose_Click();
	void mnuCheckboxes_Click();
	void mnuPictures_Click();
	void mnuFullColumnSelect_Click();

//Common Menu/Toolbar
public:
	void SaveXML();
	void LoadXML();
	void Print();

//Data Loading
public:
	void Access_LoadTasks();
	void AddRow_Task(int lID, int lDepth, CString sTaskType, CString sDescription, COleDateTime dtStartDate, COleDateTime dtEndDate, float fPercentCompleted, BOOL bSummary, BOOL bHasTasks);
	void AddPredecessor(int lPredecessorID, int lSuccessorID, int lType, int lLagFactor, int yLagInterval);



	enum { IDD = IDD_FWBSPROJECT };

protected:
	virtual void DoDataExchange(CDataExchange* pDX);    

	DECLARE_MESSAGE_MAP()
public:
	DECLARE_EVENTSINK_MAP()
	void EndTextEditActiveganttvcctl1(LPDISPATCH e);
	void PagePrintActiveganttvcctl1(LPDISPATCH e);
	afx_msg void OnClose();
};