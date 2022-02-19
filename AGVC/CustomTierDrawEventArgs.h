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

class CustomTierDrawEventArgs : public CCmdTarget
{
	DECLARE_DYNCREATE(CustomTierDrawEventArgs)

public:
	CustomTierDrawEventArgs(void);
	virtual ~CustomTierDrawEventArgs();
	virtual void OnFinalRelease();

// Member Variables
public:
	CString mp_sText;				
	BOOL mp_bCustomDraw;				
	CString mp_sStyleIndex;				
	LONG mp_yTierPosition;				
	COleDateTime mp_dtStartDate;				
	COleDateTime mp_dtEndDate;			
	LONG mp_lLeft;				
	LONG mp_lTop;				
	LONG mp_lRight;				
	LONG mp_lBottom;				
	LONG mp_lLeftTrim;				
	LONG mp_lRightTrim;				
	clsGDIGraphics mp_oGraphics;				
	LONG mp_yInterval;				
	LONG mp_lFactor;

//Internal Properties (Extensions to ODL)
public:
	CString GetText(void);
	void SetText(CString Value);
	BOOL GetCustomDraw(void);
	void SetCustomDraw(BOOL Value);
	CString GetStyleIndex(void);
	void SetStyleIndex(CString Value);
	LONG GetTierPosition(void);
	void SetTierPosition(LONG Value);
	COleDateTime GetStartDate(void);
	void SetStartDate(COleDateTime Value);
	COleDateTime GetEndDate(void);
	void SetEndDate(COleDateTime Value);
	LONG GetLeft(void);
	void SetLeft(LONG Value);
	LONG GetTop(void);
	void SetTop(LONG Value);
	LONG GetRight(void);
	void SetRight(LONG Value);
	LONG GetBottom(void);
	void SetBottom(LONG Value);
	LONG GetLeftTrim(void);
	void SetLeftTrim(LONG Value);
	LONG GetRightTrim(void);
	void SetRightTrim(LONG Value);
	clsGDIGraphics* GetGraphics(void);
	LONG GetInterval(void);
	void SetInterval(LONG Value);
	LONG GetFactor(void);
	void SetFactor(LONG Value);

//Internal Properties
public:

//Internal Methods
public:
	void Clear(void);

protected:
	DECLARE_DISPATCH_MAP()
	DECLARE_MESSAGE_MAP()
	DECLARE_OLECREATE(CustomTierDrawEventArgs)
	DECLARE_INTERFACE_MAP()

BSTR odl_GetText(void)
{
	
	return GetText().AllocSysString();
}

void odl_SetText(LPCTSTR Value)
{
	
	SetText(Value);
}

BOOL odl_GetCustomDraw(void)
{
	
	return GetCustomDraw();
}

void odl_SetCustomDraw(BOOL Value)
{
	
	
	SetCustomDraw(Value);
}

BSTR odl_GetStyleIndex(void)
{
	
	return GetStyleIndex().AllocSysString();
}

void odl_SetStyleIndex(LPCTSTR Value)
{
	
	SetStyleIndex(Value);
}

LONG odl_GetTierPosition(void)
{
	
	return GetTierPosition();
}

void odl_SetTierPosition(LONG Value)
{
	
	SetTierPosition(Value);
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

LONG odl_GetLeft(void)
{
	
	return GetLeft();
}

void odl_SetLeft(LONG Value)
{
	
	SetLeft(Value);
}

LONG odl_GetTop(void)
{
	
	return GetTop();
}

void odl_SetTop(LONG Value)
{
	
	SetTop(Value);
}

LONG odl_GetRight(void)
{
	
	return GetRight();
}

void odl_SetRight(LONG Value)
{
	
	SetRight(Value);
}

LONG odl_GetBottom(void)
{
	
	return GetBottom();
}

void odl_SetBottom(LONG Value)
{
	
	SetBottom(Value);
}

LONG odl_GetLeftTrim(void)
{
	
	return GetLeftTrim();
}

void odl_SetLeftTrim(LONG Value)
{
	
	SetLeftTrim(Value);
}

LONG odl_GetRightTrim(void)
{
	
	return GetRightTrim();
}

void odl_SetRightTrim(LONG Value)
{
	
	SetRightTrim(Value);
}

IDispatch* odl_GetGraphics(void)
{
	
	return GetGraphics()->GetIDispatch(TRUE);
}

LONG odl_GetInterval(void)
{
	
	return GetInterval();
}

void odl_SetInterval(LONG Value)
{
	
	SetInterval(Value);
}

LONG odl_GetFactor(void)
{
	
	return GetFactor();
}

void odl_SetFactor(LONG Value)
{
	
	SetFactor(Value);
}



};