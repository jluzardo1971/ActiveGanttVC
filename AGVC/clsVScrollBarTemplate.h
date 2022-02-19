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


class clsVScrollBarTemplate : public CCmdTarget
{
	DECLARE_DYNCREATE(clsVScrollBarTemplate)
	clsVScrollBarTemplate();

	enum E_BUTTON
	{
		BTN_NONE = 0,
		BTN_UP = 1,
		BTN_DOWN = 2,
		BTN_UPLCHANGE = 3,
		BTN_DOWNLCHANGE = 4,
		BTN_BUTTON = 5,
	}E_BUTTON;

public:
	CActiveGanttVCCtl* mp_oControl;

	clsVScrollBarTemplate(CActiveGanttVCCtl* oControl);
	virtual ~clsVScrollBarTemplate();
	virtual void OnFinalRelease();

// Member Variables
public:
	LONG mp_lSmallChange;
	LONG mp_lLargeChange;
	LONG mp_lValue;
	LONG mp_lMin;
	LONG mp_lMax;
	LONG ButtonX1;
	LONG ButtonX2;
	LONG ButtonY1;
	LONG ButtonY2;
	BOOL Visible;
	BOOL mp_bEnabled;
	LONG Height;
	LONG Width;
	LONG Top;
	LONG Left;
	BOOL mp_bMouseDown;
	LONG mp_iMouseYPosition;
	LONG mp_iTimerInterval;
	LONG mp_yState;
	LONG mp_lValueBuff;
	LONG mp_lTimerInterval;				
	LONG mp_iButton;
	E_SCROLLBAR mp_yType;

	CString mp_sStyleIndex;
	clsStyle* mp_oStyle;
	clsButtonState* mp_oArrowButtons;
	clsButtonState* mp_oThumbButton;

//Internal Properties (Extensions to ODL)
public:
	LONG GetTimerInterval(void);
	void SetTimerInterval(LONG Value);
	CString GetStyleIndex(void);
	void SetStyleIndex(CString Value);
	clsStyle* GetStyle(void);
	clsButtonState* GetArrowButtons(void);
	clsButtonState* GetThumbButton(void);

//Internal Properties
public:
	BOOL GetEnabled(void);
	void SetEnabled(BOOL Value);
	LONG GetValue(void);
	void SetValue(LONG Value);
	LONG GetMin(void);
	void SetMin(LONG Value);
	LONG GetMax(void);
	void SetMax(LONG Value);
	LONG GetSmallChange(void);
	void SetSmallChange(LONG Value);
	LONG GetLargeChange(void);
	void SetLargeChange(LONG Value);
	LONG GetState(void);
	void SetState(LONG Value);

//Internal Methods
public:
	CString GetXML(void);
	void SetXML(CString sXML);

//Internal Methods
public:
	void Initialize(E_SCROLLBAR yType);
	void Draw(void);
	void OnMouseHover(LONG X,LONG Y);
	void OnMouseDown(LONG X,LONG Y);
	void OnMouseMove(LONG X,LONG Y);
	void OnMouseUp(void);
	BOOL OverControl(LONG X,LONG Y);
	void ConvertToRelativeCoords(LONG& X,LONG& Y);
	BOOL ScrollBarClick(LONG X,LONG Y);
	void SmallChangeUp(void);
	void SmallChangeDown(void);
	void LargeChangeUp(void);
	void LargeChangeDown(void);
	void mp_ValueChanged(void);
	void mp_oTimer_Tick(void);
	void CalculateV(void);
	LONG InvCalculateV(LONG Y);

protected:
	DECLARE_DISPATCH_MAP()
	DECLARE_MESSAGE_MAP()
	DECLARE_OLECREATE(clsVScrollBarTemplate)
	DECLARE_INTERFACE_MAP()

LONG odl_GetTimerInterval(void)
{
	
	return GetTimerInterval();
}

void odl_SetTimerInterval(LONG Value)
{
	
	SetTimerInterval(Value);
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

IDispatch* odl_GetArrowButtons(void)
{
	
	return GetArrowButtons()->GetIDispatch(TRUE);
}

IDispatch* odl_GetThumbButton(void)
{
	
	return GetThumbButton()->GetIDispatch(TRUE);
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