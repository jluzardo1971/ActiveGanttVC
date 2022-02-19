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

class clsGDIGraphics;

class ToolTipEventArgs : public CCmdTarget
{
	DECLARE_DYNCREATE(ToolTipEventArgs)

public:
	ToolTipEventArgs(void);
	virtual ~ToolTipEventArgs();
	virtual void OnFinalRelease();

// Member Variables
public:
	LONG mp_lInitialRowIndex;				
	LONG mp_lFinalRowIndex;				
	LONG mp_lTaskIndex;				
	LONG mp_lMilestoneIndex;				
	LONG mp_lPercentageIndex;				
	LONG mp_lRowIndex;				
	LONG mp_lCellIndex;				
	LONG mp_lColumnIndex;				
	COleDateTime mp_dtInitialStartDate;				
	COleDateTime mp_dtInitialEndDate;				
	COleDateTime mp_dtStartDate;				
	COleDateTime mp_dtEndDate;				
	LONG mp_lXStart;				
	LONG mp_lXEnd;				
	LONG mp_yOperation;				
	LONG mp_yEventTarget;				
	CString mp_sTaskPosition;				
	CString mp_sPredecessorPosition;				
	LONG mp_lX;				
	LONG mp_lY;				
	LONG mp_lX1;				
	LONG mp_lY1;				
	LONG mp_lX2;				
	LONG mp_lY2;				
	BOOL mp_bCustomDraw;				
	clsGDIGraphics mp_oGraphics;				
	LONG mp_yToolTipType;				

//Internal Properties (Extensions to ODL)
public:
	LONG GetInitialRowIndex(void);
	void SetInitialRowIndex(LONG Value);
	LONG GetFinalRowIndex(void);
	void SetFinalRowIndex(LONG Value);
	LONG GetTaskIndex(void);
	void SetTaskIndex(LONG Value);
	LONG GetMilestoneIndex(void);
	void SetMilestoneIndex(LONG Value);
	LONG GetPercentageIndex(void);
	void SetPercentageIndex(LONG Value);
	LONG GetRowIndex(void);
	void SetRowIndex(LONG Value);
	LONG GetCellIndex(void);
	void SetCellIndex(LONG Value);
	LONG GetColumnIndex(void);
	void SetColumnIndex(LONG Value);
	COleDateTime GetInitialStartDate(void);
	void SetInitialStartDate(COleDateTime Value);
	COleDateTime GetInitialEndDate(void);
	void SetInitialEndDate(COleDateTime Value);
	COleDateTime GetStartDate(void);
	void SetStartDate(COleDateTime Value);
	COleDateTime GetEndDate(void);
	void SetEndDate(COleDateTime Value);
	LONG GetXStart(void);
	void SetXStart(LONG Value);
	LONG GetXEnd(void);
	void SetXEnd(LONG Value);
	LONG GetOperation(void);
	void SetOperation(LONG Value);
	LONG GetEventTarget(void);
	void SetEventTarget(LONG Value);
	CString GetTaskPosition(void);
	void SetTaskPosition(CString Value);
	CString GetPredecessorPosition(void);
	void SetPredecessorPosition(CString Value);
	LONG GetX(void);
	void SetX(LONG Value);
	LONG GetY(void);
	void SetY(LONG Value);
	LONG GetX1(void);
	void SetX1(LONG Value);
	LONG GetY1(void);
	void SetY1(LONG Value);
	LONG GetX2(void);
	void SetX2(LONG Value);
	LONG GetY2(void);
	void SetY2(LONG Value);
	BOOL GetCustomDraw(void);
	void SetCustomDraw(BOOL Value);
	clsGDIGraphics* GetGraphics(void);
	LONG GetToolTipType(void);
	void SetToolTipType(LONG Value);

//Internal Properties
public:

//Internal Methods
public:
	void Clear(void);

protected:
	DECLARE_DISPATCH_MAP()
	DECLARE_MESSAGE_MAP()
	DECLARE_OLECREATE(ToolTipEventArgs)
	DECLARE_INTERFACE_MAP()

LONG odl_GetInitialRowIndex(void)
{
	
	return GetInitialRowIndex();
}

void odl_SetInitialRowIndex(LONG Value)
{
	
	SetInitialRowIndex(Value);
}

LONG odl_GetFinalRowIndex(void)
{
	
	return GetFinalRowIndex();
}

void odl_SetFinalRowIndex(LONG Value)
{
	
	SetFinalRowIndex(Value);
}

LONG odl_GetTaskIndex(void)
{
	
	return GetTaskIndex();
}

void odl_SetTaskIndex(LONG Value)
{
	
	SetTaskIndex(Value);
}

LONG odl_GetMilestoneIndex(void)
{
	
	return GetMilestoneIndex();
}

void odl_SetMilestoneIndex(LONG Value)
{
	
	SetMilestoneIndex(Value);
}

LONG odl_GetPercentageIndex(void)
{
	
	return GetPercentageIndex();
}

void odl_SetPercentageIndex(LONG Value)
{
	
	SetPercentageIndex(Value);
}

LONG odl_GetRowIndex(void)
{
	
	return GetRowIndex();
}

void odl_SetRowIndex(LONG Value)
{
	
	SetRowIndex(Value);
}

LONG odl_GetCellIndex(void)
{
	
	return GetCellIndex();
}

void odl_SetCellIndex(LONG Value)
{
	
	SetCellIndex(Value);
}

LONG odl_GetColumnIndex(void)
{
	
	return GetColumnIndex();
}

void odl_SetColumnIndex(LONG Value)
{
	
	SetColumnIndex(Value);
}

DATE odl_GetInitialStartDate(void)
{
	
	return mp_dtInitialStartDate;
}

void odl_SetInitialStartDate(DATE Value)
{
	SetInitialStartDate(Value);
}

DATE odl_GetInitialEndDate(void)
{
	
	return mp_dtInitialEndDate;
}

void odl_SetInitialEndDate(DATE Value)
{
	SetInitialEndDate(Value);
}

DATE odl_GetStartDate(void)
{
	
	return mp_dtStartDate;
}

void odl_SetStartDate(DATE Value)
{
	SetStartDate(Value);
}

DATE odl_GetEndDate(void)
{
	
	return mp_dtEndDate;
}

void odl_SetEndDate(DATE Value)
{
	SetEndDate(Value);
}

LONG odl_GetXStart(void)
{
	
	return GetXStart();
}

void odl_SetXStart(LONG Value)
{
	
	SetXStart(Value);
}

LONG odl_GetXEnd(void)
{
	
	return GetXEnd();
}

void odl_SetXEnd(LONG Value)
{
	
	SetXEnd(Value);
}

LONG odl_GetOperation(void)
{
	
	return GetOperation();
}

void odl_SetOperation(LONG Value)
{
	
	SetOperation(Value);
}

LONG odl_GetEventTarget(void)
{
	
	return GetEventTarget();
}

void odl_SetEventTarget(LONG Value)
{
	
	SetEventTarget(Value);
}

BSTR odl_GetTaskPosition(void)
{
	
	return GetTaskPosition().AllocSysString();
}

void odl_SetTaskPosition(LPCTSTR Value)
{
	
	SetTaskPosition(Value);
}

BSTR odl_GetPredecessorPosition(void)
{
	
	return GetPredecessorPosition().AllocSysString();
}

void odl_SetPredecessorPosition(LPCTSTR Value)
{
	
	SetPredecessorPosition(Value);
}

LONG odl_GetX(void)
{
	
	return GetX();
}

void odl_SetX(LONG Value)
{
	
	SetX(Value);
}

LONG odl_GetY(void)
{
	
	return GetY();
}

void odl_SetY(LONG Value)
{
	
	SetY(Value);
}

LONG odl_GetX1(void)
{
	
	return GetX1();
}

void odl_SetX1(LONG Value)
{
	
	SetX1(Value);
}

LONG odl_GetY1(void)
{
	
	return GetY1();
}

void odl_SetY1(LONG Value)
{
	
	SetY1(Value);
}

LONG odl_GetX2(void)
{
	
	return GetX2();
}

void odl_SetX2(LONG Value)
{
	
	SetX2(Value);
}

LONG odl_GetY2(void)
{
	
	return GetY2();
}

void odl_SetY2(LONG Value)
{
	
	SetY2(Value);
}

BOOL odl_GetCustomDraw(void)
{
	
	return GetCustomDraw();
}

void odl_SetCustomDraw(BOOL Value)
{
	
	
	SetCustomDraw(Value);
}

IDispatch* odl_GetGraphics(void)
{
	
	return GetGraphics()->GetIDispatch(TRUE);
}

LONG odl_GetToolTipType(void)
{
	
	return GetToolTipType();
}

void odl_SetToolTipType(LONG Value)
{
	
	SetToolTipType(Value);
}

};