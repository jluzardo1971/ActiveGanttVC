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

class clsTimeLine;
class clsGrid;

class clsClientArea : public CCmdTarget
{
	DECLARE_DYNCREATE(clsClientArea)
	clsClientArea();

public:
	CActiveGanttVCCtl* mp_oControl;

	clsClientArea(CActiveGanttVCCtl* oControl, clsTimeLine* oTimeLine);
	virtual ~clsClientArea();
	virtual void OnFinalRelease();

// Member Variables
public:
	clsTimeLine* mp_oTimeLine;
	BOOL mp_bDetectConflicts;
	LONG mp_lMilestoneSelectionOffset;
	LONG mp_lPredecessorSelectionOffset;
	LONG mp_lTaskBorderSelectionOffset;
	LONG mp_lLastVisibleRow;
	CString mp_sToolTipFormat;
	BOOL mp_bToolTipsVisible;				
	LONG mp_lTop;				
	LONG mp_lBottom;				
	LONG mp_lLeft;				
	LONG mp_lRight;				
	LONG mp_lWidth;				
	LONG mp_lHeight;				
	clsGrid* mp_oGrid;				


//Internal Properties (Extensions to ODL)
public:
	BOOL GetDetectConflicts(void);
	void SetDetectConflicts(BOOL Value);
	LONG GetMilestoneSelectionOffset(void);
	void SetMilestoneSelectionOffset(LONG Value);
	LONG GetFirstVisibleRow(void);
	void SetFirstVisibleRow(LONG Value);
	LONG GetLastVisibleRow(void);
	CString GetToolTipFormat(void);
	void SetToolTipFormat(CString Value);
	BOOL GetToolTipsVisible(void);
	void SetToolTipsVisible(BOOL Value);
	LONG GetTop(void);
	LONG GetBottom(void);
	LONG GetLeft(void);
	LONG GetRight(void);
	LONG GetWidth(void);
	LONG GetHeight(void);
	clsGrid* GetGrid(void);

	LONG GetPredecessorSelectionOffset(void);
	void SetPredecessorSelectionOffset(LONG Value);

	LONG GetTaskBorderSelectionOffset(void);
	void SetTaskBorderSelectionOffset(LONG Value);

//Internal Properties
public:
	void Setf_LastVisibleRow(LONG Value);

//Internal Methods
public:
	void Draw(void);
	CString GetXML(void);
	void SetXML(CString sXML);
	void Clear();
	void Clone(clsClientArea* oClone);

protected:
	DECLARE_DISPATCH_MAP()
	DECLARE_MESSAGE_MAP()
	DECLARE_OLECREATE(clsClientArea)
	DECLARE_INTERFACE_MAP()


BOOL odl_GetDetectConflicts(void)
{
	
	return GetDetectConflicts();
}

void odl_SetDetectConflicts(BOOL Value)
{
	
	
	SetDetectConflicts(Value);
}

LONG odl_GetMilestoneSelectionOffset(void)
{
	
	return GetMilestoneSelectionOffset();
}

void odl_SetMilestoneSelectionOffset(LONG Value)
{
	
	SetMilestoneSelectionOffset(Value);
}

LONG odl_GetPredecessorSelectionOffset(void)
{
	
	return GetPredecessorSelectionOffset();
}

void odl_SetPredecessorSelectionOffset(LONG Value)
{
	
	SetPredecessorSelectionOffset(Value);
}

LONG odl_GetTaskBorderSelectionOffset(void)
{
	
	return GetTaskBorderSelectionOffset();
}

void odl_SetTaskBorderSelectionOffset(LONG Value)
{
	
	SetTaskBorderSelectionOffset(Value);
}

LONG odl_GetFirstVisibleRow(void)
{
	
	return GetFirstVisibleRow();
}

void odl_SetFirstVisibleRow(LONG Value)
{
	
	SetFirstVisibleRow(Value);
}

LONG odl_GetLastVisibleRow(void)
{
	
	return GetLastVisibleRow();
}

BSTR odl_GetToolTipFormat(void)
{
	
	return GetToolTipFormat().AllocSysString();
}

void odl_SetToolTipFormat(LPCTSTR Value)
{
	
	SetToolTipFormat(Value);
}

BOOL odl_GetToolTipsVisible(void)
{
	
	return GetToolTipsVisible();
}

void odl_SetToolTipsVisible(BOOL Value)
{
	
	
	SetToolTipsVisible(Value);
}

LONG odl_GetTop(void)
{
	
	return GetTop();
}

LONG odl_GetBottom(void)
{
	
	return GetBottom();
}

LONG odl_GetLeft(void)
{
	
	return GetLeft();
}

LONG odl_GetRight(void)
{
	
	return GetRight();
}

LONG odl_GetWidth(void)
{
	
	return GetWidth();
}

LONG odl_GetHeight(void)
{
	
	return GetHeight();
}

IDispatch* odl_GetGrid(void)
{
	
	return GetGrid()->GetIDispatch(TRUE);
}

BSTR odl_GetXML(void)
{
	
	return GetXML().AllocSysString();
}

void odl_SetXML(LPCTSTR sXML)
{
	
	SetXML(sXML);
}


};