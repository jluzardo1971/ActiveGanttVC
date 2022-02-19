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

class clsLayer;
class clsImage;
class clsRow;
class clsStyle;
class clsPredecessors;

class clsTask : public clsItemBase
{
	DECLARE_DYNCREATE(clsTask)
	clsTask();

public:
	CActiveGanttVCCtl* mp_oControl;

	clsTask(CActiveGanttVCCtl* oControl);
	virtual ~clsTask();
	virtual void OnFinalRelease();

// Member Variables
public:
	BOOL mp_bAllowStretchLeft;
	BOOL mp_bAllowStretchRight;
	COleDateTime mp_dtEndDate;
	COleDateTime mp_dtStartDate;
	CString mp_sText;
	CString mp_sLayerIndex;
	clsImage* mp_oImage;
	CString mp_sRowKey;
	CString mp_sStyleIndex;
	clsStyle* mp_oStyle;
	CString mp_sTag;
	LONG mp_yAllowedMovement;
	BOOL mp_bVisible;
	clsRow* mp_oRow;
	clsLayer* mp_oLayer;
	BOOL mp_bIncomingPredecessors;
	BOOL mp_bOutgoingPredecessors;
	CString mp_sImageTag;
	LONG mp_lTextLeft;
    LONG mp_lTextTop;
    LONG mp_lTextRight;
    LONG mp_lTextBottom;
	BOOL mp_bAllowTextEdit;
	BOOL mp_bWarning;
	CString mp_sWarningStyleIndex;
	clsStyle* mp_oWarningStyle;
    LONG mp_yTaskType;
    LONG mp_yDurationInterval;
    LONG mp_lDurationFactor;

//Internal Properties (Extensions to ODL)
public:
	CString GetKey(void);
	void SetKey(CString Value);
	BOOL GetIncomingPredecessors(void);
	void SetIncomingPredecessors(BOOL Value);
	BOOL GetOutgoingPredecessors(void);
	void SetOutgoingPredecessors(BOOL Value);
	BOOL GetAllowStretchLeft(void);
	void SetAllowStretchLeft(BOOL Value);
	BOOL GetAllowStretchRight(void);
	void SetAllowStretchRight(BOOL Value);
	CString GetText(void);
	void SetText(CString Value);
	CString GetLayerIndex(void);
	void SetLayerIndex(CString Value);
	clsLayer* GetLayer(void);
	clsImage* GetImage(void);
	CString GetRowKey(void);
	void SetRowKey(CString Value);
	clsRow* GetRow(void);
	CString GetStyleIndex(void);
	void SetStyleIndex(CString Value);
	clsStyle* GetStyle(void);
	CString GetTag(void);
	void SetTag(CString Value);
	LONG GetAllowedMovement(void);
	void SetAllowedMovement(LONG Value);
	LONG GetLeftTrim(void);
	LONG GetRightTrim(void);
	COleDateTime GetStartDate(void);
	void SetStartDate(COleDateTime Value);
	LONG GetLeft(void);
	LONG GetRight(void);
	COleDateTime GetEndDate(void);
	void SetEndDate(COleDateTime Value);
	LONG GetTop(void);
	LONG GetBottom(void);
	BOOL GetVisible(void);
	void SetVisible(BOOL Value);
	LONG GetType(void);
	LONG GetIndex(void);
	void SetIndex(LONG Value);
	CString GetImageTag(void);
	void SetImageTag(CString Value);
	BOOL GetAllowTextEdit(void);
	void SetAllowTextEdit(BOOL Value);
	BOOL GetWarning(void);
	CString GetWarningStyleIndex(void);
	void SetWarningStyleIndex(CString Value);
	clsStyle* GetWarningStyle(void);
	LONG GetTaskType(void);
	void SetTaskType(LONG Value);
	LONG GetDurationInterval(void);
	void SetDurationInterval(LONG Value);
	LONG GetDurationFactor(void);
	void SetDurationFactor(LONG Value);

//Internal Properties
public:
	BOOL Getf_bLeftVisible(void);
	BOOL Getf_bRightVisible(void);
	BOOL GetInsideVisibleTimeLineArea(void);
	E_CLIENTAREAVISIBILITY GetClientAreaVisiblity(void);

//Internal Methods
public:
	BOOL InConflict(void);
	CString GetXML(void);
	void SetXML(CString sXML);
	void mp_GetDuration(void);

protected:
	DECLARE_DISPATCH_MAP()
	DECLARE_MESSAGE_MAP()
	DECLARE_OLECREATE(clsTask)
	DECLARE_INTERFACE_MAP()

BOOL odl_GetAllowTextEdit(void)
{
	
	return GetAllowTextEdit();
}

void odl_SetAllowTextEdit(BOOL Value)
{
	
	
	SetAllowTextEdit(Value);
}

LONG odl_GetIndex(void)
{
	
	return GetIndex();
}

BSTR odl_GetKey(void)
{
	
	return GetKey().AllocSysString();
}

void odl_SetKey(LPCTSTR Value)
{
	
	SetKey(Value);
}

BOOL odl_GetIncomingPredecessors(void)
{
	
	return GetIncomingPredecessors();
}

void odl_SetIncomingPredecessors(BOOL Value)
{
	
	
	SetIncomingPredecessors(Value);
}

BOOL odl_GetOutgoingPredecessors(void)
{
	
	return GetOutgoingPredecessors();
}

void odl_SetOutgoingPredecessors(BOOL Value)
{
	
	
	SetOutgoingPredecessors(Value);
}

BOOL odl_GetAllowStretchLeft(void)
{
	
	return GetAllowStretchLeft();
}

void odl_SetAllowStretchLeft(BOOL Value)
{
	
	
	SetAllowStretchLeft(Value);
}

BOOL odl_GetAllowStretchRight(void)
{
	
	return GetAllowStretchRight();
}

void odl_SetAllowStretchRight(BOOL Value)
{
	
	
	SetAllowStretchRight(Value);
}

BSTR odl_GetText(void)
{
	
	return GetText().AllocSysString();
}

void odl_SetText(LPCTSTR Value)
{
	
	SetText(Value);
}

BSTR odl_GetLayerIndex(void)
{
	
	return GetLayerIndex().AllocSysString();
}

void odl_SetLayerIndex(LPCTSTR Value)
{
	
	SetLayerIndex(Value);
}

IDispatch* odl_GetLayer(void)
{
	
	return GetLayer()->GetIDispatch(TRUE);
}

IDispatch* odl_GetImage(void)
{
	
	return GetImage()->GetIDispatch(TRUE);
}

BSTR odl_GetRowKey(void)
{
	
	return GetRowKey().AllocSysString();
}

void odl_SetRowKey(LPCTSTR Value)
{
	
	SetRowKey(Value);
}

IDispatch* odl_GetRow(void)
{
	
	return GetRow()->GetIDispatch(TRUE);
}

BSTR odl_GetStyleIndex(void)
{
	
	return GetStyleIndex().AllocSysString();
}

void odl_SetStyleIndex(LPCTSTR Value)
{
	
	SetStyleIndex(Value);
}

IDispatch* odl_GetStyle(void)
{
	
	return GetStyle()->GetIDispatch(TRUE);
}

BSTR odl_GetTag(void)
{
	
	return GetTag().AllocSysString();
}

void odl_SetTag(LPCTSTR Value)
{
	
	SetTag(Value);
}

LONG odl_GetAllowedMovement(void)
{
	
	return GetAllowedMovement();
}

void odl_SetAllowedMovement(LONG Value)
{
	
	SetAllowedMovement(Value);
}

LONG odl_GetLeftTrim(void)
{
	
	return GetLeftTrim();
}

LONG odl_GetRightTrim(void)
{
	
	return GetRightTrim();
}

DATE odl_GetStartDate(void)
{
	
	return GetStartDate();
}

void odl_SetStartDate(DATE Value)
{
	SetStartDate(Value);
}

LONG odl_GetLeft(void)
{
	
	return GetLeft();
}

LONG odl_GetRight(void)
{
	
	return GetRight();
}

DATE odl_GetEndDate(void)
{	
	return GetEndDate();
}

void odl_SetEndDate(DATE Value)
{
	SetEndDate(Value);
}

LONG odl_GetTop(void)
{
	
	return GetTop();
}

LONG odl_GetBottom(void)
{
	
	return GetBottom();
}

BOOL odl_GetVisible(void)
{
	
	return GetVisible();
}

void odl_SetVisible(BOOL Value)
{
	
	
	SetVisible(Value);
}

LONG odl_GetType(void)
{
	
	return GetType();
}

BSTR odl_GetImageTag(void)
{
	
	return GetImageTag().AllocSysString();
}

void odl_SetImageTag(LPCTSTR Value)
{
	
	SetImageTag(Value);
}

BOOL odl_InConflict(void)
{
	
	return InConflict();
}

BSTR odl_GetXML(void)
{
	
	return GetXML().AllocSysString();
}

void odl_SetXML(LPCTSTR sXML)
{
	
	SetXML(sXML);
}

BOOL odl_GetWarning(void)
{
	
	return GetWarning();
}

BSTR odl_GetWarningStyleIndex(void)
{
	
	return GetWarningStyleIndex().AllocSysString();
}

void odl_SetWarningStyleIndex(LPCTSTR Value)
{
	
	SetWarningStyleIndex(Value);
}

IDispatch* odl_GetWarningStyle(void)
{
	
	return GetWarningStyle()->GetIDispatch(TRUE);
}

LONG odl_GetTaskType(void)
{
	
	return GetTaskType();
}

void odl_SetTaskType(LONG Value)
{
	
	SetTaskType(Value);
}

LONG odl_GetDurationInterval(void)
{
	
	return GetDurationInterval();
}

void odl_SetDurationInterval(LONG Value)
{
	
	SetDurationInterval(Value);
}

LONG odl_GetDurationFactor(void)
{
	
	return GetDurationFactor();
}

void odl_SetDurationFactor(LONG Value)
{
	
	SetDurationFactor(Value);
}

};