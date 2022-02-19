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


// clsSelectionRectangleStyle command target

class clsSelectionRectangleStyle : public CCmdTarget
{
	DECLARE_DYNCREATE(clsSelectionRectangleStyle)
	clsSelectionRectangleStyle();

public:
	CActiveGanttVCCtl* mp_oControl;

	clsSelectionRectangleStyle(CActiveGanttVCCtl* oControl);
	virtual ~clsSelectionRectangleStyle();
	virtual void OnFinalRelease();

// Member Variables
public:
	LONG mp_lOffsetBottom;
	LONG mp_lOffsetLeft;
	LONG mp_lOffsetRight;
	LONG mp_lOffsetTop;
	BOOL mp_bVisible;
    LONG mp_yMode;
    LONG mp_lBorderWidth;
	Gdiplus::Color mp_clrColor;

//Internal Properties (Extensions to ODL)
public:
	LONG GetOffsetBottom(void);
	void SetOffsetBottom(LONG Value);
	LONG GetOffsetLeft(void);
	void SetOffsetLeft(LONG Value);
	LONG GetOffsetRight(void);
	void SetOffsetRight(LONG Value);
	LONG GetOffsetTop(void);
	void SetOffsetTop(LONG Value);
	BOOL GetVisible(void);
	void SetVisible(BOOL Value);
	LONG GetMode(void);
	void SetMode(LONG Value);
	LONG GetBorderWidth(void);
	void SetBorderWidth(LONG Value);
	Gdiplus::Color GetColor(void);
	void SetColor(Gdiplus::Color Value);

//Internal Methods
public:
	CString GetXML(void);
	void SetXML(CString sXML);
	void Clear();
	void Clone(clsSelectionRectangleStyle* oClone);

protected:
	DECLARE_MESSAGE_MAP()
	DECLARE_OLECREATE(clsSelectionRectangleStyle)
	DECLARE_DISPATCH_MAP()
	DECLARE_INTERFACE_MAP()

LONG odl_GetOffsetBottom(void)
{
	
	return GetOffsetBottom();
}

void odl_SetOffsetBottom(LONG Value)
{
	
	SetOffsetBottom(Value);
}

LONG odl_GetOffsetLeft(void)
{
	
	return GetOffsetLeft();
}

void odl_SetOffsetLeft(LONG Value)
{
	
	SetOffsetLeft(Value);
}

LONG odl_GetOffsetRight(void)
{
	
	return GetOffsetRight();
}

void odl_SetOffsetRight(LONG Value)
{
	
	SetOffsetRight(Value);
}

LONG odl_GetOffsetTop(void)
{
	
	return GetOffsetTop();
}

void odl_SetOffsetTop(LONG Value)
{
	
	SetOffsetTop(Value);
}

BOOL odl_GetVisible(void)
{
	
	return GetVisible();
}

void odl_SetVisible(BOOL Value)
{
	
	
	SetVisible(Value);
}

LONG odl_GetMode(void)
{
	
	return GetMode();
}

void odl_SetMode(LONG Value)
{
	
	SetMode(Value);
}

LONG odl_GetBorderWidth(void)
{
	
	return GetBorderWidth();
}

void odl_SetBorderWidth(LONG Value)
{
	
	SetBorderWidth(Value);
}

OLE_COLOR odl_GetColor(void)
{
	
	return (OLE_COLOR) GetColor().ToCOLORREF();
}

void odl_SetColor(OLE_COLOR Value)
{
	
	Gdiplus::Color oColor;
	oColor.SetFromCOLORREF(g_TranslateColor(Value));
	SetColor(oColor);
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