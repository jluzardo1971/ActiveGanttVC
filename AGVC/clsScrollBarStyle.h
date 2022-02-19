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


// clsScrollBarStyle command target

class clsScrollBarStyle : public CCmdTarget
{
	DECLARE_DYNCREATE(clsScrollBarStyle)
	clsScrollBarStyle();

public:
	CActiveGanttVCCtl* mp_oControl;

	clsScrollBarStyle(CActiveGanttVCCtl* oControl);
	virtual ~clsScrollBarStyle();
	virtual void OnFinalRelease();

// Member Variables
public:
    Gdiplus::Color mp_clrArrowColor;
    Gdiplus::Color mp_clrDropShadowArrowColor;
    BOOL mp_bDropShadow;
    LONG mp_lLeftX;
    LONG mp_lLeftY;
    LONG mp_lUpX;
    LONG mp_lUpY;
    LONG mp_lRightX;
    LONG mp_lRightY;
    LONG mp_lDownX;
    LONG mp_lDownY;
    LONG mp_lDropShadowLeftX;
    LONG mp_lDropShadowLeftY;
    LONG mp_lDropShadowUpX;
    LONG mp_lDropShadowUpY;
    LONG mp_lDropShadowRightX;
    LONG mp_lDropShadowRightY;
    LONG mp_lDropShadowDownX;
    LONG mp_lDropShadowDownY;
    LONG mp_lArrowSize;


//Internal Properties (Extensions to ODL)
public:
	Gdiplus::Color GetArrowColor(void);
	void SetArrowColor(Gdiplus::Color Value);
	Gdiplus::Color GetDropShadowArrowColor(void);
	void SetDropShadowArrowColor(Gdiplus::Color Value);
	BOOL GetDropShadow(void);
	void SetDropShadow(BOOL Value);
	LONG GetLeftX(void);
	void SetLeftX(LONG Value);
	LONG GetLeftY(void);
	void SetLeftY(LONG Value);
	LONG GetUpX(void);
	void SetUpX(LONG Value);
	LONG GetUpY(void);
	void SetUpY(LONG Value);
	LONG GetRightX(void);
	void SetRightX(LONG Value);
	LONG GetRightY(void);
	void SetRightY(LONG Value);
	LONG GetDownX(void);
	void SetDownX(LONG Value);
	LONG GetDownY(void);
	void SetDownY(LONG Value);
	LONG GetDropShadowLeftX(void);
	void SetDropShadowLeftX(LONG Value);
	LONG GetDropShadowLeftY(void);
	void SetDropShadowLeftY(LONG Value);
	LONG GetDropShadowUpX(void);
	void SetDropShadowUpX(LONG Value);
	LONG GetDropShadowUpY(void);
	void SetDropShadowUpY(LONG Value);
	LONG GetDropShadowRightX(void);
	void SetDropShadowRightX(LONG Value);
	LONG GetDropShadowRightY(void);
	void SetDropShadowRightY(LONG Value);
	LONG GetDropShadowDownX(void);
	void SetDropShadowDownX(LONG Value);
	LONG GetDropShadowDownY(void);
	void SetDropShadowDownY(LONG Value);
	LONG GetArrowSize(void);
	void SetArrowSize(LONG Value);

//Internal Methods
public:
	CString GetXML(void);
	void SetXML(CString sXML);
	void Clear();
	void Clone(clsScrollBarStyle* oClone);

protected:
	DECLARE_MESSAGE_MAP()
	DECLARE_OLECREATE(clsScrollBarStyle)
	DECLARE_DISPATCH_MAP()
	DECLARE_INTERFACE_MAP()

OLE_COLOR odl_GetArrowColor(void)
{
	
	return (OLE_COLOR) GetArrowColor().ToCOLORREF();
}

void odl_SetArrowColor(OLE_COLOR Value)
{
	
	Gdiplus::Color oColor;
	oColor.SetFromCOLORREF(g_TranslateColor(Value));
	SetArrowColor(oColor);
}

OLE_COLOR odl_GetDropShadowArrowColor(void)
{
	
	return (OLE_COLOR) GetDropShadowArrowColor().ToCOLORREF();
}

void odl_SetDropShadowArrowColor(OLE_COLOR Value)
{
	
	Gdiplus::Color oColor;
	oColor.SetFromCOLORREF(g_TranslateColor(Value));
	SetDropShadowArrowColor(oColor);
}

BOOL odl_GetDropShadow(void)
{
	
	return GetDropShadow();
}

void odl_SetDropShadow(BOOL Value)
{
	
	SetDropShadow(Value);
}

LONG odl_GetLeftX(void)
{
	
	return GetLeftX();
}

void odl_SetLeftX(LONG Value)
{
	
	SetLeftX(Value);
}

LONG odl_GetLeftY(void)
{
	
	return GetLeftY();
}

void odl_SetLeftY(LONG Value)
{
	
	SetLeftY(Value);
}

LONG odl_GetUpX(void)
{
	
	return GetUpX();
}

void odl_SetUpX(LONG Value)
{
	
	SetUpX(Value);
}

LONG odl_GetUpY(void)
{
	
	return GetUpY();
}

void odl_SetUpY(LONG Value)
{
	
	SetUpY(Value);
}

LONG odl_GetRightX(void)
{
	
	return GetRightX();
}

void odl_SetRightX(LONG Value)
{
	
	SetRightX(Value);
}

LONG odl_GetRightY(void)
{
	
	return GetRightY();
}

void odl_SetRightY(LONG Value)
{
	
	SetRightY(Value);
}

LONG odl_GetDownX(void)
{
	
	return GetDownX();
}

void odl_SetDownX(LONG Value)
{
	
	SetDownX(Value);
}

LONG odl_GetDownY(void)
{
	
	return GetDownY();
}

void odl_SetDownY(LONG Value)
{
	
	SetDownY(Value);
}

LONG odl_GetDropShadowLeftX(void)
{
	
	return GetDropShadowLeftX();
}

void odl_SetDropShadowLeftX(LONG Value)
{
	
	SetDropShadowLeftX(Value);
}

LONG odl_GetDropShadowLeftY(void)
{
	
	return GetDropShadowLeftY();
}

void odl_SetDropShadowLeftY(LONG Value)
{
	
	SetDropShadowLeftY(Value);
}

LONG odl_GetDropShadowUpX(void)
{
	
	return GetDropShadowUpX();
}

void odl_SetDropShadowUpX(LONG Value)
{
	
	SetDropShadowUpX(Value);
}

LONG odl_GetDropShadowUpY(void)
{
	
	return GetDropShadowUpY();
}

void odl_SetDropShadowUpY(LONG Value)
{
	
	SetDropShadowUpY(Value);
}

LONG odl_GetDropShadowRightX(void)
{
	
	return GetDropShadowRightX();
}

void odl_SetDropShadowRightX(LONG Value)
{
	
	SetDropShadowRightX(Value);
}

LONG odl_GetDropShadowRightY(void)
{
	
	return GetDropShadowRightY();
}

void odl_SetDropShadowRightY(LONG Value)
{
	
	SetDropShadowRightY(Value);
}

LONG odl_GetDropShadowDownX(void)
{
	
	return GetDropShadowDownX();
}

void odl_SetDropShadowDownX(LONG Value)
{
	
	SetDropShadowDownX(Value);
}

LONG odl_GetDropShadowDownY(void)
{
	
	return GetDropShadowDownY();
}

void odl_SetDropShadowDownY(LONG Value)
{
	
	SetDropShadowDownY(Value);
}

LONG odl_GetArrowSize(void)
{
	
	return GetArrowSize();
}

void odl_SetArrowSize(LONG Value)
{
	
	SetArrowSize(Value);
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