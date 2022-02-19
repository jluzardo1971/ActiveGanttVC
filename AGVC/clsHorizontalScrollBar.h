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

class clsHorizontalScrollBar : public CCmdTarget
{
	DECLARE_DYNCREATE(clsHorizontalScrollBar)
	clsHorizontalScrollBar();

public:
	CActiveGanttVCCtl* mp_oControl;

	clsHorizontalScrollBar(CActiveGanttVCCtl* oControl);
	virtual ~clsHorizontalScrollBar();
	virtual void OnFinalRelease();

// Member Variables
public:
	BOOL mp_bVisible;
	LONG mp_lMin;				
	LONG mp_lMax;				
	LONG mp_lValue;				
	LONG mp_lSmallChange;				
	LONG mp_lLargeChange;				
	clsHScrollBarTemplate* mp_oScrollBar;				

	LONG mp_yState;				
	LONG mp_lWidth;				
	LONG mp_lHeight;				
	LONG mp_lLeft;				
	LONG mp_lTop;				

//Internal Properties (Extensions to ODL)
public:
	LONG GetMin(void);
	LONG GetMax(void);
	LONG GetValue(void);
	void SetValue(LONG Value);
	BOOL mf_Visible(void);
	BOOL GetVisible(void);
	void SetVisible(BOOL Value);
	LONG GetSmallChange(void);
	void SetSmallChange(LONG Value);
	LONG GetLargeChange(void);
	void SetLargeChange(LONG Value);
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
	CString GetXML(void);
	void SetXML(CString sXML);

//Internal Methods
public:
	void Update(void);
	void Reset(void);
	void Position(void);
	void oHScrollBar_ValueChanged(LONG Offset);

protected:
	DECLARE_DISPATCH_MAP()
	DECLARE_MESSAGE_MAP()
	DECLARE_OLECREATE(clsHorizontalScrollBar)
	DECLARE_INTERFACE_MAP()


LONG odl_GetMin(void)
{
	
	return GetMin();
}

LONG odl_GetMax(void)
{
	
	return GetMax();
}

LONG odl_GetValue(void)
{
	
	return GetValue();
}

void odl_SetValue(LONG Value)
{
	
	SetValue(Value);
}

BOOL odl_GetVisible(void)
{
	
	return GetVisible();
}

void odl_SetVisible(BOOL Value)
{
	
	
	SetVisible(Value);
}

LONG odl_GetSmallChange(void)
{
	
	return GetSmallChange();
}

void odl_SetSmallChange(LONG Value)
{
	
	SetSmallChange(Value);
}

LONG odl_GetLargeChange(void)
{
	
	return GetLargeChange();
}

void odl_SetLargeChange(LONG Value)
{
	
	SetLargeChange(Value);
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