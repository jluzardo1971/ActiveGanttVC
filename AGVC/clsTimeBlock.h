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

class clsStyle;

class clsTimeBlock : public clsItemBase
{
	DECLARE_DYNCREATE(clsTimeBlock)
	clsTimeBlock();

public:
	CActiveGanttVCCtl* mp_oControl;

	clsTimeBlock(CActiveGanttVCCtl* oControl);
	virtual ~clsTimeBlock();
	virtual void OnFinalRelease();

// Member Variables
public:
	CString mp_sStyleIndex;
	CString mp_sTag;
	BOOL mp_bGenerateConflict;
	BOOL mp_bVisible;
	clsStyle* mp_oStyle;
	LONG mp_yTimeBlockType;
	LONG mp_yRecurringType;			
	LONG mp_lLeftTrim;				
	LONG mp_lRightTrim;				
	LONG mp_lLeft;				
	LONG mp_lTop;				
	LONG mp_lRight;				
	LONG mp_lBottom;
    LONG mp_yBaseWeekDay;
    COleDateTime mp_dtBaseDate;
    LONG mp_yDurationInterval;
    LONG mp_lDurationFactor;
    BOOL mp_bNonWorking;

//Internal Properties (Extensions to ODL)
public:
	CString GetKey(void);
	void SetKey(CString Value);
	LONG GetTimeBlockType(void);
	void SetTimeBlockType(LONG Value);
	LONG GetRecurringType(void);
	void SetRecurringType(LONG Value);
	COleDateTime GetEndDate(void);
	COleDateTime GetStartDate(void);
	CString GetStyleIndex(void);
	void SetStyleIndex(CString Value);
	clsStyle* GetStyle(void);
	CString GetTag(void);
	void SetTag(CString Value);
	BOOL GetGenerateConflict(void);
	void SetGenerateConflict(BOOL Value);
	LONG GetLeftTrim(void);
	LONG GetRightTrim(void);
	LONG GetLeft(void);
	LONG GetTop(void);
	LONG GetRight(void);
	LONG GetBottom(void);
	BOOL GetVisible(void);
	void SetVisible(BOOL Value);
	LONG GetIndex(void);
	void SetIndex(LONG Value);
	COleDateTime GetBaseDate(void);
	void SetBaseDate(COleDateTime Value);
	LONG GetBaseWeekDay(void);
	void SetBaseWeekDay(LONG Value);
	LONG GetDurationInterval(void);
	void SetDurationInterval(LONG Value);
	LONG GetDurationFactor(void);
	void SetDurationFactor(LONG Value);
	BOOL GetNonWorking(void);
	void SetNonWorking(BOOL Value);

//Internal Properties
public:
	BOOL Getf_Visible(void);
	void Setf_Visible(BOOL Value);

//Internal Methods
public:
	CString GetXML(void);
	void SetXML(CString sXML);

protected:
	DECLARE_DISPATCH_MAP()
	DECLARE_MESSAGE_MAP()
	DECLARE_OLECREATE(clsTimeBlock)
	DECLARE_INTERFACE_MAP()

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

LONG odl_GetTimeBlockType(void)
{
	
	return GetTimeBlockType();
}

void odl_SetTimeBlockType(LONG Value)
{
	
	SetTimeBlockType(Value);
}

LONG odl_GetRecurringType(void)
{
	
	return GetRecurringType();
}

void odl_SetRecurringType(LONG Value)
{
	
	SetRecurringType(Value);
}

DATE odl_GetEndDate(void)
{
	
	return GetEndDate();
}

DATE odl_GetStartDate(void)
{
	
	return GetStartDate();
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

BOOL odl_GetGenerateConflict(void)
{
	
	return GetGenerateConflict();
}

void odl_SetGenerateConflict(BOOL Value)
{
	
	
	SetGenerateConflict(Value);
}

LONG odl_GetLeftTrim(void)
{
	
	return GetLeftTrim();
}

LONG odl_GetRightTrim(void)
{
	
	return GetRightTrim();
}

LONG odl_GetLeft(void)
{
	
	return GetLeft();
}

LONG odl_GetTop(void)
{
	
	return GetTop();
}

LONG odl_GetRight(void)
{
	
	return GetRight();
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

DATE odl_GetBaseDate(void)
{
	
	return mp_dtBaseDate;
}

void odl_SetBaseDate(DATE Value)
{
	SetBaseDate(Value);
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

BOOL odl_GetNonWorking(void)
{
	
	return GetNonWorking();
}

void odl_SetNonWorking(BOOL Value)
{
	
	SetNonWorking(Value);
}

BSTR odl_GetXML(void)
{
	
	return GetXML().AllocSysString();
}

void odl_SetXML(LPCTSTR sXML)
{
	
	SetXML(sXML);
}

LONG odl_GetBaseWeekDay(void)
{
	
	return GetBaseWeekDay();
}

void odl_SetBaseWeekDay(LONG Value)
{
	
	SetBaseWeekDay(Value);
}

};