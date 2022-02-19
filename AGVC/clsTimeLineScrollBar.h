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

class clsHScrollBarTemplate;

class clsTimeLineScrollBar : public CCmdTarget
{
	DECLARE_DYNCREATE(clsTimeLineScrollBar)
	clsTimeLineScrollBar();

public:
	CActiveGanttVCCtl* mp_oControl;

	clsTimeLineScrollBar(CActiveGanttVCCtl* oControl);
	virtual ~clsTimeLineScrollBar();
	virtual void OnFinalRelease();

// Member Variables
public:
	COleDateTime mp_dtStartDate;
	LONG mp_yInterval;
	LONG mp_lFactor;
	BOOL mp_bVisible;
	clsHScrollBarTemplate* mp_oScrollBar;




//Internal Properties (Extensions to ODL)
public:
	LONG GetInterval(void);
	void SetInterval(LONG Value);
	LONG GetFactor(void);
	void SetFactor(LONG Value);
	LONG GetValue(void);
	void SetValue(LONG Value);
	BOOL GetEnabled(void);
	void SetEnabled(BOOL Value);
	BOOL mf_Visible(void);
	BOOL GetVisible(void);
	void SetVisible(BOOL Value);
	LONG GetLargeChange(void);
	void SetLargeChange(LONG Value);
	LONG GetMax(void);
	void SetMax(LONG Value);
	LONG GetSmallChange(void);
	void SetSmallChange(LONG Value);
	COleDateTime GetStartDate(void);
	void SetStartDate(COleDateTime Value);
	clsHScrollBarTemplate* GetScrollBar(void);

//Internal Properties
public:
	LONG GetState(void);
	void SetState(LONG Value);
	LONG GetWidth(void);
	void SetWidth(LONG Value);
	LONG GetHeight(void);
	void SetHeight(LONG Value);
	LONG GetLeft(void);
	void SetLeft(LONG Value);
	LONG GetTop(void);
	void SetTop(LONG Value);

//Internal Methods
public:
	void Finalize(void);
	CString GetXML(void);
	void SetXML(CString sXML);
	void oHScrollBar_ValueChanged(LONG Offset);
	void Position(void);
	void Clear();
	void Clone(clsTimeLineScrollBar* oClone);

protected:
	DECLARE_DISPATCH_MAP()
	DECLARE_MESSAGE_MAP()
	DECLARE_OLECREATE(clsTimeLineScrollBar)
	DECLARE_INTERFACE_MAP()

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

LONG odl_GetValue(void)
{
	
	return GetValue();
}

void odl_SetValue(LONG Value)
{
	
	SetValue(Value);
}

BOOL odl_GetEnabled(void)
{
	
	return GetEnabled();
}

void odl_SetEnabled(BOOL Value)
{
	
	
	SetEnabled(Value);
}

BOOL odl_GetVisible(void)
{
	
	return GetVisible();
}

void odl_SetVisible(BOOL Value)
{
	
	
	SetVisible(Value);
}

LONG odl_GetLargeChange(void)
{
	
	return GetLargeChange();
}

void odl_SetLargeChange(LONG Value)
{
	
	SetLargeChange(Value);
}

LONG odl_GetMax(void)
{
	
	return GetMax();
}

void odl_SetMax(LONG Value)
{
	
	SetMax(Value);
}

LONG odl_GetSmallChange(void)
{
	
	return GetSmallChange();
}

void odl_SetSmallChange(LONG Value)
{
	
	SetSmallChange(Value);
}

DATE odl_GetStartDate(void)
{
	
	return mp_dtStartDate;
}

void odl_SetStartDate(DATE Value)
{
	SetStartDate(Value);
}

IDispatch* odl_GetScrollBar(void)
{
	
	return GetScrollBar()->GetIDispatch(TRUE);
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