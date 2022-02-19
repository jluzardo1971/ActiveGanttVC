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

class clsFont;
class clsImage;

class clsToolTip : public CCmdTarget
{
	DECLARE_DYNCREATE(clsToolTip)
	clsToolTip();

public:
	CActiveGanttVCCtl* mp_oControl;

	clsToolTip(CActiveGanttVCCtl* oControl);
	virtual ~clsToolTip();
	virtual void OnFinalRelease();

// Member Variables
public:
	LONG mp_lLeft;
	LONG mp_lTop;
	LONG mp_lWidth;
	LONG mp_lHeight;
	CString mp_sText;
	BOOL mp_bVisible;
	BOOL mp_bBackupDCActive;
	clsFont mp_oFont;
	LONG mp_lBackupLeft;
	LONG mp_lBackupTop;
	LONG mp_lBackupRight;
	LONG mp_lBackupBottom;
	BOOL mp_bAutomaticSizing;
	Gdiplus::Color mp_clrBackColor;
	Gdiplus::Color mp_clrForeColor;
	Gdiplus::Color mp_clrBorderColor;
	Gdiplus::Graphics* mp_oToolTipGraphics;
	LONG mp_lX;
	LONG mp_lY;

//Internal Properties (Extensions to ODL)
public:
	clsFont* GetFont(void);
	Gdiplus::Color GetBackColor(void);
	void SetBackColor(Gdiplus::Color Value);
	Gdiplus::Color GetForeColor(void);
	void SetForeColor(Gdiplus::Color Value);
	Gdiplus::Color GetBorderColor(void);
	void SetBorderColor(Gdiplus::Color Value);
	CString GetText(void);
	void SetText(CString Value);
	BOOL GetAutomaticSizing(void);
	void SetAutomaticSizing(BOOL Value);
	LONG GetLeft(void);
	void SetLeft(LONG Value);
	LONG GetRight(void);
	LONG GetTop(void);
	void SetTop(LONG Value);
	LONG GetBottom(void);
	LONG GetWidth(void);
	void SetWidth(LONG Value);
	LONG GetHeight(void);
	void SetHeight(LONG Value);
	BOOL GetVisible(void);
	void SetVisible(BOOL Value);

//Internal Properties
public:


//Internal Methods
public:
	void RestoreBackupDC(void);

protected:
	DECLARE_DISPATCH_MAP()
	DECLARE_MESSAGE_MAP()
	DECLARE_OLECREATE(clsToolTip)
	DECLARE_INTERFACE_MAP()

IDispatch* odl_GetFont(void)
{
	
	return GetFont()->GetIDispatch(TRUE);
}

OLE_COLOR odl_GetBackColor(void)
{
	
	return (OLE_COLOR) GetBackColor().ToCOLORREF();
}

void odl_SetBackColor(OLE_COLOR Value)
{
	
	Gdiplus::Color oColor;
	oColor.SetFromCOLORREF(g_TranslateColor(Value));
	SetBackColor(oColor);
}

OLE_COLOR odl_GetForeColor(void)
{
	
	return (OLE_COLOR) GetForeColor().ToCOLORREF();
}

void odl_SetForeColor(OLE_COLOR Value)
{
	
	Gdiplus::Color oColor;
	oColor.SetFromCOLORREF(g_TranslateColor(Value));
	SetForeColor(oColor);
}

OLE_COLOR odl_GetBorderColor(void)
{
	
	return (OLE_COLOR) GetBorderColor().ToCOLORREF();
}

void odl_SetBorderColor(OLE_COLOR Value)
{
	
	Gdiplus::Color oColor;
	oColor.SetFromCOLORREF(g_TranslateColor(Value));
	SetBorderColor(oColor);
}

BSTR odl_GetText(void)
{
	
	return GetText().AllocSysString();
}

void odl_SetText(LPCTSTR Value)
{
	
	SetText(Value);
}

BOOL odl_GetAutomaticSizing(void)
{
	
	return GetAutomaticSizing();
}

void odl_SetAutomaticSizing(BOOL Value)
{
	
	
	SetAutomaticSizing(Value);
}

LONG odl_GetLeft(void)
{
	
	return GetLeft();
}

void odl_SetLeft(LONG Value)
{
	
	SetLeft(Value);
}

LONG odl_GetRight(void)
{
	
	return GetRight();
}

LONG odl_GetTop(void)
{
	
	return GetTop();
}

void odl_SetTop(LONG Value)
{
	
	SetTop(Value);
}

LONG odl_GetBottom(void)
{
	
	return GetBottom();
}

LONG odl_GetWidth(void)
{
	
	return GetWidth();
}

void odl_SetWidth(LONG Value)
{
	
	SetWidth(Value);
}

LONG odl_GetHeight(void)
{
	
	return GetHeight();
}

void odl_SetHeight(LONG Value)
{
	
	SetHeight(Value);
}

BOOL odl_GetVisible(void)
{
	
	return GetVisible();
}

void odl_SetVisible(BOOL Value)
{
	
	
	SetVisible(Value);
}


};