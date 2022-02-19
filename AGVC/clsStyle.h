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
class clsTaskStyle;
class clsMilestoneStyle;
class clsPredecessorStyle;
class clsTextFlags;
class clsCustomBorderStyle;

class clsStyle : public clsItemBase
{
	DECLARE_DYNCREATE(clsStyle)
	clsStyle();

public:
	CActiveGanttVCCtl* mp_oControl;

	clsStyle(CActiveGanttVCCtl* oControl);
	virtual ~clsStyle();
	virtual void OnFinalRelease();

// Member Variables
public:
	BOOL mp_bTextVisible;
	LONG mp_yBackgroundMode;
	BOOL mp_bClipText;
	BOOL mp_bDrawTextInVisibleArea;
	BOOL mp_bUseMask;
	Gdiplus::Color mp_clrBackColor;
	Gdiplus::Color mp_clrForeColor;
	Gdiplus::Color mp_clrBorderColor;
	Gdiplus::Color mp_clrEndGradientColor;
	Gdiplus::Color mp_clrStartGradientColor;
	clsFont mp_oFont;
	LONG mp_lPatternFactor;
	LONG mp_lTextXMargin;
	LONG mp_lTextYMargin;
	LONG mp_lOffsetBottom;
	LONG mp_lOffsetTop;
	LONG mp_lImageXMargin;
	LONG mp_lImageYMargin;
	CString mp_sTag;
	LONG mp_lBorderWidth;
	LONG mp_yAppearance;
	LONG mp_yPattern;
	LONG mp_yBorderStyle;
	LONG mp_yButtonStyle;
	LONG mp_yTextAlignmentHorizontal;
	LONG mp_yTextAlignmentVertical;
	LONG mp_yTextPlacement;
	LONG mp_yGradientFillMode;
	LONG mp_yImageAlignmentHorizontal;
	LONG mp_yImageAlignmentVertical;
	LONG mp_yPlacement;
	LONG mp_yFillMode;
	LONG mp_yHatchStyle;
	Gdiplus::Color mp_clrHatchBackColor;
	Gdiplus::Color mp_clrHatchForeColor;
	Gdiplus::Color mp_clrTextEditBackColor;
	Gdiplus::Color mp_clrTextEditForeColor;
	clsTaskStyle* mp_oTaskStyle;				
	clsMilestoneStyle* mp_oMilestoneStyle;				
	clsPredecessorStyle* mp_oPredecessorStyle;				
	clsTextFlags* mp_oTextFlags;				
	clsCustomBorderStyle* mp_oCustomBorderStyle;
    clsScrollBarStyle* mp_oScrollBarStyle;
    clsSelectionRectangleStyle* mp_oSelectionRectangleStyle;
    clsButtonBorderStyle* mp_oButtonBorderStyle;

//Internal Properties (Extensions to ODL)
public:

	Gdiplus::Color GetTextEditBackColor(void);
	void SetTextEditBackColor(Gdiplus::Color Value);
	Gdiplus::Color GetTextEditForeColor(void);
	void SetTextEditForeColor(Gdiplus::Color Value);
	Gdiplus::Color GetHatchBackColor(void);
	void SetHatchBackColor(Gdiplus::Color Value);
	Gdiplus::Color GetHatchForeColor(void);
	void SetHatchForeColor(Gdiplus::Color Value);
	CString GetKey(void);
	void SetKey(CString Value);
	LONG GetAppearance(void);
	void SetAppearance(LONG Value);
	Gdiplus::Color GetBackColor(void);
	void SetBackColor(Gdiplus::Color Value);
	LONG GetPattern(void);
	void SetPattern(LONG Value);
	Gdiplus::Color GetBorderColor(void);
	void SetBorderColor(Gdiplus::Color Value);
	LONG GetBorderStyle(void);
	void SetBorderStyle(LONG Value);
	LONG GetButtonStyle(void);
	void SetButtonStyle(LONG Value);
	LONG GetTextAlignmentHorizontal(void);
	void SetTextAlignmentHorizontal(LONG Value);
	LONG GetTextAlignmentVertical(void);
	void SetTextAlignmentVertical(LONG Value);
	BOOL GetTextVisible(void);
	void SetTextVisible(BOOL Value);
	LONG GetTextXMargin(void);
	void SetTextXMargin(LONG Value);
	LONG GetTextYMargin(void);
	void SetTextYMargin(LONG Value);
	BOOL GetClipText(void);
	void SetClipText(BOOL Value);
	clsFont* GetFont(void);
	Gdiplus::Color GetForeColor(void);
	void SetForeColor(Gdiplus::Color Value);
	LONG GetOffsetBottom(void);
	void SetOffsetBottom(LONG Value);
	LONG GetOffsetTop(void);
	void SetOffsetTop(LONG Value);
	LONG GetImageAlignmentHorizontal(void);
	void SetImageAlignmentHorizontal(LONG Value);
	LONG GetImageAlignmentVertical(void);
	void SetImageAlignmentVertical(LONG Value);
	LONG GetImageXMargin(void);
	void SetImageXMargin(LONG Value);
	LONG GetImageYMargin(void);
	void SetImageYMargin(LONG Value);
	LONG GetPlacement(void);
	void SetPlacement(LONG Value);
	BOOL GetUseMask(void);
	void SetUseMask(BOOL Value);
	LONG GetTextPlacement(void);
	void SetTextPlacement(LONG Value);
	LONG GetPatternFactor(void);
	void SetPatternFactor(LONG Value);
	LONG GetHatchStyle(void);
	void SetHatchStyle(LONG Value);
	Gdiplus::Color GetStartGradientColor(void);
	void SetStartGradientColor(Gdiplus::Color Value);
	Gdiplus::Color GetEndGradientColor(void);
	void SetEndGradientColor(Gdiplus::Color Value);
	LONG GetGradientFillMode(void);
	void SetGradientFillMode(LONG Value);
	LONG GetFillMode(void);
	void SetFillMode(LONG Value);
	LONG GetBackgroundMode(void);
	void SetBackgroundMode(LONG Value);
	BOOL GetDrawTextInVisibleArea(void);
	void SetDrawTextInVisibleArea(BOOL Value);
	CString GetTag(void);
	void SetTag(CString Value);
	LONG GetBorderWidth(void);
	void SetBorderWidth(LONG Value);
	LONG GetIndex(void);
	void SetIndex(LONG Value);
	clsTaskStyle* GetTaskStyle(void);
	clsMilestoneStyle* GetMilestoneStyle(void);
	clsPredecessorStyle* GetPredecessorStyle(void);
	clsTextFlags* GetTextFlags(void);
	clsCustomBorderStyle* GetCustomBorderStyle(void);
	clsScrollBarStyle* GetScrollBarStyle(void);
	clsSelectionRectangleStyle* GetSelectionRectangleStyle(void);
	clsButtonBorderStyle* GetButtonBorderStyle(void);

//Internal Properties
public:

//Internal Methods
public:
	CString GetXML(void);
	void SetXML(CString sXML);
	void Clear(void);
	clsStyle* Clone(CString Key);

protected:
	DECLARE_DISPATCH_MAP()
	DECLARE_MESSAGE_MAP()
	DECLARE_OLECREATE(clsStyle)
	DECLARE_INTERFACE_MAP()


OLE_COLOR odl_GetTextEditBackColor(void)
{
	
	return (OLE_COLOR) GetTextEditBackColor().ToCOLORREF();
}

void odl_SetTextEditBackColor(OLE_COLOR Value)
{
	
	Gdiplus::Color oColor;
	oColor.SetFromCOLORREF(g_TranslateColor(Value));
	SetTextEditBackColor(oColor);
}

OLE_COLOR odl_GetTextEditForeColor(void)
{
	
	return (OLE_COLOR) GetTextEditForeColor().ToCOLORREF();
}

void odl_SetTextEditForeColor(OLE_COLOR Value)
{
	
	Gdiplus::Color oColor;
	oColor.SetFromCOLORREF(g_TranslateColor(Value));
	SetTextEditForeColor(oColor);
}

OLE_COLOR odl_GetHatchBackColor(void)
{
	
	return (OLE_COLOR) GetHatchBackColor().ToCOLORREF();
}

void odl_SetHatchBackColor(OLE_COLOR Value)
{
	
	Gdiplus::Color oColor;
	oColor.SetFromCOLORREF(g_TranslateColor(Value));
	SetHatchBackColor(oColor);
}

LONG odl_GetIndex(void)
{
	
	return GetIndex();
}

OLE_COLOR odl_GetHatchForeColor(void)
{
	
	return (OLE_COLOR) GetHatchForeColor().ToCOLORREF();
}

void odl_SetHatchForeColor(OLE_COLOR Value)
{
	
	Gdiplus::Color oColor;
	oColor.SetFromCOLORREF(g_TranslateColor(Value));
	SetHatchForeColor(oColor);
}

BSTR odl_GetKey(void)
{
	
	return GetKey().AllocSysString();
}

void odl_SetKey(LPCTSTR Value)
{
	
	SetKey(Value);
}

LONG odl_GetAppearance(void)
{
	
	return GetAppearance();
}

void odl_SetAppearance(LONG Value)
{
	
	SetAppearance(Value);
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

LONG odl_GetPattern(void)
{
	
	return GetPattern();
}

void odl_SetPattern(LONG Value)
{
	
	SetPattern(Value);
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

LONG odl_GetBorderStyle(void)
{
	
	return GetBorderStyle();
}

void odl_SetBorderStyle(LONG Value)
{
	
	SetBorderStyle(Value);
}

LONG odl_GetButtonStyle(void)
{
	
	return GetButtonStyle();
}

void odl_SetButtonStyle(LONG Value)
{
	
	SetButtonStyle(Value);
}

LONG odl_GetTextAlignmentHorizontal(void)
{
	
	return GetTextAlignmentHorizontal();
}

void odl_SetTextAlignmentHorizontal(LONG Value)
{
	
	SetTextAlignmentHorizontal(Value);
}

LONG odl_GetTextAlignmentVertical(void)
{
	
	return GetTextAlignmentVertical();
}

void odl_SetTextAlignmentVertical(LONG Value)
{
	
	SetTextAlignmentVertical(Value);
}

BOOL odl_GetTextVisible(void)
{
	
	return GetTextVisible();
}

void odl_SetTextVisible(BOOL Value)
{
	
	
	SetTextVisible(Value);
}

LONG odl_GetTextXMargin(void)
{
	
	return GetTextXMargin();
}

void odl_SetTextXMargin(LONG Value)
{
	
	SetTextXMargin(Value);
}

LONG odl_GetTextYMargin(void)
{
	
	return GetTextYMargin();
}

void odl_SetTextYMargin(LONG Value)
{
	
	SetTextYMargin(Value);
}

BOOL odl_GetClipText(void)
{
	
	return GetClipText();
}

void odl_SetClipText(BOOL Value)
{
	
	
	SetClipText(Value);
}

IDispatch* odl_GetFont(void)
{
	
	return GetFont()->GetIDispatch(TRUE);
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

LONG odl_GetOffsetBottom(void)
{
	
	return GetOffsetBottom();
}

void odl_SetOffsetBottom(LONG Value)
{
	
	SetOffsetBottom(Value);
}

LONG odl_GetOffsetTop(void)
{
	
	return GetOffsetTop();
}

void odl_SetOffsetTop(LONG Value)
{
	
	SetOffsetTop(Value);
}

LONG odl_GetImageAlignmentHorizontal(void)
{
	
	return GetImageAlignmentHorizontal();
}

void odl_SetImageAlignmentHorizontal(LONG Value)
{
	
	SetImageAlignmentHorizontal(Value);
}

LONG odl_GetImageAlignmentVertical(void)
{
	
	return GetImageAlignmentVertical();
}

void odl_SetImageAlignmentVertical(LONG Value)
{
	
	SetImageAlignmentVertical(Value);
}

LONG odl_GetImageXMargin(void)
{
	
	return GetImageXMargin();
}

void odl_SetImageXMargin(LONG Value)
{
	
	SetImageXMargin(Value);
}

LONG odl_GetImageYMargin(void)
{
	
	return GetImageYMargin();
}

void odl_SetImageYMargin(LONG Value)
{
	
	SetImageYMargin(Value);
}

LONG odl_GetPlacement(void)
{
	
	return GetPlacement();
}

void odl_SetPlacement(LONG Value)
{
	
	SetPlacement(Value);
}

BOOL odl_GetUseMask(void)
{
	
	return GetUseMask();
}

void odl_SetUseMask(BOOL Value)
{
	
	
	SetUseMask(Value);
}

LONG odl_GetTextPlacement(void)
{
	
	return GetTextPlacement();
}

void odl_SetTextPlacement(LONG Value)
{
	
	SetTextPlacement(Value);
}

LONG odl_GetPatternFactor(void)
{
	
	return GetPatternFactor();
}

void odl_SetPatternFactor(LONG Value)
{
	
	SetPatternFactor(Value);
}

LONG odl_GetHatchStyle(void)
{
	
	return GetHatchStyle();
}

void odl_SetHatchStyle(LONG Value)
{
	
	SetHatchStyle(Value);
}

OLE_COLOR odl_GetStartGradientColor(void)
{
	
	return (OLE_COLOR) GetStartGradientColor().ToCOLORREF();
}

void odl_SetStartGradientColor(OLE_COLOR Value)
{
	
	Gdiplus::Color oColor;
	oColor.SetFromCOLORREF(g_TranslateColor(Value));
	SetStartGradientColor(oColor);
}

OLE_COLOR odl_GetEndGradientColor(void)
{
	
	return (OLE_COLOR) GetEndGradientColor().ToCOLORREF();
}

void odl_SetEndGradientColor(OLE_COLOR Value)
{
	
	Gdiplus::Color oColor;
	oColor.SetFromCOLORREF(g_TranslateColor(Value));
	SetEndGradientColor(oColor);
}

LONG odl_GetGradientFillMode(void)
{
	
	return GetGradientFillMode();
}

void odl_SetGradientFillMode(LONG Value)
{
	
	SetGradientFillMode(Value);
}

LONG odl_GetFillMode(void)
{
	
	return GetFillMode();
}

void odl_SetFillMode(LONG Value)
{
	
	SetFillMode(Value);
}

LONG odl_GetBackgroundMode(void)
{
	
	return GetBackgroundMode();
}

void odl_SetBackgroundMode(LONG Value)
{
	
	SetBackgroundMode(Value);
}

BOOL odl_GetDrawTextInVisibleArea(void)
{
	
	return GetDrawTextInVisibleArea();
}

void odl_SetDrawTextInVisibleArea(BOOL Value)
{
	
	
	SetDrawTextInVisibleArea(Value);
}

BSTR odl_GetTag(void)
{
	
	return GetTag().AllocSysString();
}

void odl_SetTag(LPCTSTR Value)
{
	
	SetTag(Value);
}

LONG odl_GetBorderWidth(void)
{
	
	return GetBorderWidth();
}

void odl_SetBorderWidth(LONG Value)
{
	
	SetBorderWidth(Value);
}

IDispatch* odl_GetTaskStyle(void)
{
	
	return GetTaskStyle()->GetIDispatch(TRUE);
}

IDispatch* odl_GetMilestoneStyle(void)
{
	
	return GetMilestoneStyle()->GetIDispatch(TRUE);
}

IDispatch* odl_GetPredecessorStyle(void)
{
	
	return GetPredecessorStyle()->GetIDispatch(TRUE);
}

IDispatch* odl_GetTextFlags(void)
{
	
	return GetTextFlags()->GetIDispatch(TRUE);
}

IDispatch* odl_GetCustomBorderStyle(void)
{
	
	return GetCustomBorderStyle()->GetIDispatch(TRUE);
}

IDispatch* odl_GetScrollBarStyle(void)
{
	
	return GetScrollBarStyle()->GetIDispatch(TRUE);
}

IDispatch* odl_GetSelectionRectangleStyle(void)
{
	
	return GetSelectionRectangleStyle()->GetIDispatch(TRUE);
}

IDispatch* odl_GetButtonBorderStyle(void)
{
	
	return GetButtonBorderStyle()->GetIDispatch(TRUE);
}

BSTR odl_GetXML(void)
{
	
	return GetXML().AllocSysString();
}

void odl_SetXML(LPCTSTR sXML)
{
	
	SetXML(sXML);
}

void odl_Clear(void)
{
	
	Clear();
}

IDispatch* odl_Clone(LPCTSTR Key)
{
	
	return Clone(Key)->GetIDispatch(TRUE);
}


};