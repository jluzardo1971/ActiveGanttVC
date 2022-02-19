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
#include "stdafx.h"
#include "ActiveGanttVC.h"
#include "ActiveGanttVCCtl.h"
#include "clsXML.h"
#include "clsStyle.h"



IMPLEMENT_DYNCREATE(clsStyle, clsItemBase)

//{FE222487-24C2-4331-8281-9D85D504E2EC}
static const IID IID_IclsStyle = { 0xFE222487, 0x24C2, 0x4331, { 0x82, 0x81, 0x9D, 0x85, 0xD5, 0x04, 0xE2, 0xEC} };

//{E767BE2D-405F-4783-B7F0-DBB9AAFEFA02}
IMPLEMENT_OLECREATE_FLAGS(clsStyle, "AGVC.clsStyle", afxRegApartmentThreading, 0xE767BE2D, 0x405F, 0x4783, 0xB7, 0xF0, 0xDB, 0xB9, 0xAA, 0xFE, 0xFA, 0x02)

BEGIN_DISPATCH_MAP(clsStyle, clsItemBase)
	DISP_PROPERTY_EX_ID(clsStyle, "HatchBackColor", 1, odl_GetHatchBackColor, odl_SetHatchBackColor, VT_COLOR) 
	DISP_PROPERTY_EX_ID(clsStyle, "HatchForeColor", 2, odl_GetHatchForeColor, odl_SetHatchForeColor, VT_COLOR) 
	DISP_PROPERTY_EX_ID(clsStyle, "Key", 3, odl_GetKey, odl_SetKey, VT_BSTR) 
	DISP_PROPERTY_EX_ID(clsStyle, "Appearance", 4, odl_GetAppearance, odl_SetAppearance, VT_I4) 
	DISP_PROPERTY_EX_ID(clsStyle, "BackColor", 5, odl_GetBackColor, odl_SetBackColor, VT_COLOR) 
	DISP_PROPERTY_EX_ID(clsStyle, "Pattern", 6, odl_GetPattern, odl_SetPattern, VT_I4) 
	DISP_PROPERTY_EX_ID(clsStyle, "BorderColor", 7, odl_GetBorderColor, odl_SetBorderColor, VT_COLOR) 
	DISP_PROPERTY_EX_ID(clsStyle, "BorderStyle", 8, odl_GetBorderStyle, odl_SetBorderStyle, VT_I4) 
	DISP_PROPERTY_EX_ID(clsStyle, "ButtonStyle", 9, odl_GetButtonStyle, odl_SetButtonStyle, VT_I4) 
	DISP_PROPERTY_EX_ID(clsStyle, "TextAlignmentHorizontal", 10, odl_GetTextAlignmentHorizontal, odl_SetTextAlignmentHorizontal, VT_I4) 
	DISP_PROPERTY_EX_ID(clsStyle, "TextAlignmentVertical", 11, odl_GetTextAlignmentVertical, odl_SetTextAlignmentVertical, VT_I4) 
	DISP_PROPERTY_EX_ID(clsStyle, "TextVisible", 12, odl_GetTextVisible, odl_SetTextVisible, VT_BOOL) 
	DISP_PROPERTY_EX_ID(clsStyle, "TextXMargin", 13, odl_GetTextXMargin, odl_SetTextXMargin, VT_I4) 
	DISP_PROPERTY_EX_ID(clsStyle, "TextYMargin", 14, odl_GetTextYMargin, odl_SetTextYMargin, VT_I4) 
	DISP_PROPERTY_EX_ID(clsStyle, "ClipText", 15, odl_GetClipText, odl_SetClipText, VT_BOOL) 
	DISP_PROPERTY_EX_ID(clsStyle, "Font", 16, odl_GetFont, SetNotSupported, VT_DISPATCH) 
	DISP_PROPERTY_EX_ID(clsStyle, "ForeColor", 17, odl_GetForeColor, odl_SetForeColor, VT_COLOR) 
	DISP_PROPERTY_EX_ID(clsStyle, "OffsetBottom", 18, odl_GetOffsetBottom, odl_SetOffsetBottom, VT_I4) 
	DISP_PROPERTY_EX_ID(clsStyle, "OffsetTop", 19, odl_GetOffsetTop, odl_SetOffsetTop, VT_I4) 
	DISP_PROPERTY_EX_ID(clsStyle, "ImageAlignmentHorizontal", 20, odl_GetImageAlignmentHorizontal, odl_SetImageAlignmentHorizontal, VT_I4) 
	DISP_PROPERTY_EX_ID(clsStyle, "ImageAlignmentVertical", 21, odl_GetImageAlignmentVertical, odl_SetImageAlignmentVertical, VT_I4) 
	DISP_PROPERTY_EX_ID(clsStyle, "ImageXMargin", 22, odl_GetImageXMargin, odl_SetImageXMargin, VT_I4) 
	DISP_PROPERTY_EX_ID(clsStyle, "ImageYMargin", 23, odl_GetImageYMargin, odl_SetImageYMargin, VT_I4) 
	DISP_PROPERTY_EX_ID(clsStyle, "Placement", 24, odl_GetPlacement, odl_SetPlacement, VT_I4) 
	DISP_PROPERTY_EX_ID(clsStyle, "UseMask", 25, odl_GetUseMask, odl_SetUseMask, VT_BOOL) 
	DISP_PROPERTY_EX_ID(clsStyle, "TextPlacement", 26, odl_GetTextPlacement, odl_SetTextPlacement, VT_I4) 
	DISP_PROPERTY_EX_ID(clsStyle, "PatternFactor", 27, odl_GetPatternFactor, odl_SetPatternFactor, VT_I4) 
	DISP_PROPERTY_EX_ID(clsStyle, "HatchStyle", 28, odl_GetHatchStyle, odl_SetHatchStyle, VT_I4) 
	DISP_PROPERTY_EX_ID(clsStyle, "StartGradientColor", 29, odl_GetStartGradientColor, odl_SetStartGradientColor, VT_COLOR) 
	DISP_PROPERTY_EX_ID(clsStyle, "EndGradientColor", 30, odl_GetEndGradientColor, odl_SetEndGradientColor, VT_COLOR) 
	DISP_PROPERTY_EX_ID(clsStyle, "GradientFillMode", 31, odl_GetGradientFillMode, odl_SetGradientFillMode, VT_I4) 
	DISP_PROPERTY_EX_ID(clsStyle, "FillMode", 32, odl_GetFillMode, odl_SetFillMode, VT_I4) 
	DISP_PROPERTY_EX_ID(clsStyle, "BackgroundMode", 33, odl_GetBackgroundMode, odl_SetBackgroundMode, VT_I4) 
	DISP_PROPERTY_EX_ID(clsStyle, "DrawTextInVisibleArea", 34, odl_GetDrawTextInVisibleArea, odl_SetDrawTextInVisibleArea, VT_BOOL) 
	DISP_PROPERTY_EX_ID(clsStyle, "Tag", 35, odl_GetTag, odl_SetTag, VT_BSTR)  
	DISP_PROPERTY_EX_ID(clsStyle, "TaskStyle", 36, odl_GetTaskStyle, SetNotSupported, VT_DISPATCH) 
	DISP_PROPERTY_EX_ID(clsStyle, "MilestoneStyle", 37, odl_GetMilestoneStyle, SetNotSupported, VT_DISPATCH) 
	DISP_PROPERTY_EX_ID(clsStyle, "PredecessorStyle", 38, odl_GetPredecessorStyle, SetNotSupported, VT_DISPATCH) 
	DISP_PROPERTY_EX_ID(clsStyle, "TextFlags", 39, odl_GetTextFlags, SetNotSupported, VT_DISPATCH) 
	DISP_PROPERTY_EX_ID(clsStyle, "CustomBorderStyle", 40, odl_GetCustomBorderStyle, SetNotSupported, VT_DISPATCH) 
	DISP_FUNCTION_ID(clsStyle, "GetXML", 41, odl_GetXML, VT_BSTR, VTS_NONE) 
	DISP_FUNCTION_ID(clsStyle, "SetXML", 42, odl_SetXML, VT_EMPTY, VTS_BSTR)
	DISP_PROPERTY_EX_ID(clsStyle, "Index", 43, odl_GetIndex, SetNotSupported, VT_I4)
	DISP_PROPERTY_EX_ID(clsStyle, "BorderWidth", 44, odl_GetBorderWidth, odl_SetBorderWidth, VT_I4)
	DISP_PROPERTY_EX_ID(clsStyle, "ScrollBarStyle", 45, odl_GetScrollBarStyle, SetNotSupported, VT_DISPATCH)
	DISP_PROPERTY_EX_ID(clsStyle, "SelectionRectangleStyle", 46, odl_GetSelectionRectangleStyle, SetNotSupported, VT_DISPATCH)
	DISP_PROPERTY_EX_ID(clsStyle, "ButtonBorderStyle", 47, odl_GetButtonBorderStyle, SetNotSupported, VT_DISPATCH)
	DISP_PROPERTY_EX_ID(clsStyle, "TextEditBackColor", 48, odl_GetTextEditBackColor, odl_SetTextEditBackColor, VT_COLOR) 
	DISP_PROPERTY_EX_ID(clsStyle, "TextEditForeColor", 49, odl_GetTextEditForeColor, odl_SetTextEditForeColor, VT_COLOR)
	DISP_FUNCTION_ID(clsStyle, "Clear", 50, odl_Clear, VT_EMPTY, VTS_NONE)
	DISP_FUNCTION_ID(clsStyle, "Clone", 51, odl_Clone, VT_DISPATCH, VTS_BSTR)
END_DISPATCH_MAP()




BEGIN_INTERFACE_MAP(clsStyle, clsItemBase)
	INTERFACE_PART(clsStyle, IID_IclsStyle, Dispatch)
END_INTERFACE_MAP()

BEGIN_MESSAGE_MAP(clsStyle, clsItemBase)
END_MESSAGE_MAP()

clsStyle::clsStyle()
{
	AfxMessageBox(_T("Invalid constructor has been invoked for clsStyle. Please inform support@sourcecodestore.com of this error"), MB_OK | MB_ICONSTOP | MB_APPLMODAL);
}

clsStyle::clsStyle(CActiveGanttVCCtl* oControl)
{
	EnableAutomation();
	AfxOleLockApp();
	mp_oControl = oControl;
	mp_oTaskStyle = new clsTaskStyle(mp_oControl);
	mp_oMilestoneStyle = new clsMilestoneStyle(mp_oControl);
	mp_oPredecessorStyle = new clsPredecessorStyle(mp_oControl);
	mp_oTextFlags = new clsTextFlags(mp_oControl);
	mp_oCustomBorderStyle = new clsCustomBorderStyle(mp_oControl);
	mp_oScrollBarStyle = new clsScrollBarStyle(mp_oControl);
    mp_oSelectionRectangleStyle = new clsSelectionRectangleStyle(mp_oControl);
    mp_oButtonBorderStyle = new clsButtonBorderStyle(mp_oControl);
	Clear();
}

clsStyle::~clsStyle()
{
	delete mp_oTaskStyle;
	mp_oTaskStyle = NULL;
	delete mp_oMilestoneStyle;
	mp_oMilestoneStyle = NULL;
	delete mp_oPredecessorStyle;
	mp_oPredecessorStyle = NULL;
	delete mp_oTextFlags;
	mp_oTextFlags = NULL;
	delete mp_oCustomBorderStyle;
	mp_oCustomBorderStyle = NULL;
	delete mp_oScrollBarStyle;
	mp_oScrollBarStyle = NULL;
    delete mp_oSelectionRectangleStyle;
	mp_oSelectionRectangleStyle = NULL;
    delete mp_oButtonBorderStyle;
	mp_oButtonBorderStyle = NULL;
	AfxOleUnlockApp();
}

void clsStyle::OnFinalRelease()
{
	clsItemBase::OnFinalRelease();
}

Gdiplus::Color clsStyle::GetTextEditBackColor(void)
{
	return mp_clrTextEditBackColor;
}

void clsStyle::SetTextEditBackColor(Gdiplus::Color Value)
{
	mp_clrTextEditBackColor = Value;
}

Gdiplus::Color clsStyle::GetTextEditForeColor(void)
{
	return mp_clrTextEditForeColor;
}

void clsStyle::SetTextEditForeColor(Gdiplus::Color Value)
{
	mp_clrTextEditForeColor = Value;
}

Gdiplus::Color clsStyle::GetHatchBackColor(void)
{
	return mp_clrHatchBackColor;
}

void clsStyle::SetHatchBackColor(Gdiplus::Color Value)
{
	mp_clrHatchBackColor = Value;
}

Gdiplus::Color clsStyle::GetHatchForeColor(void)
{
	return mp_clrHatchForeColor;
}

void clsStyle::SetHatchForeColor(Gdiplus::Color Value)
{
	mp_clrHatchForeColor = Value;
}

CString clsStyle::GetKey(void)
{
	return mp_sKey;
}

void clsStyle::SetKey(CString Value)
{
	mp_sKey = Value;
}

LONG clsStyle::GetAppearance(void)
{
	return mp_yAppearance;
}

void clsStyle::SetAppearance(LONG Value)
{
	mp_yAppearance = Value;
}

Gdiplus::Color clsStyle::GetBackColor(void)
{
	return mp_clrBackColor;
}

void clsStyle::SetBackColor(Gdiplus::Color Value)
{
	mp_clrBackColor = Value;
}

LONG clsStyle::GetPattern(void)
{
	return mp_yPattern;
}

void clsStyle::SetPattern(LONG Value)
{
	mp_yPattern = Value;
}

Gdiplus::Color clsStyle::GetBorderColor(void)
{
	return mp_clrBorderColor;
}

void clsStyle::SetBorderColor(Gdiplus::Color Value)
{
	mp_clrBorderColor = Value;
}

LONG clsStyle::GetBorderStyle(void)
{
	return mp_yBorderStyle;
}

void clsStyle::SetBorderStyle(LONG Value)
{
	mp_yBorderStyle = Value;
}

LONG clsStyle::GetButtonStyle(void)
{
	return mp_yButtonStyle;
}

void clsStyle::SetButtonStyle(LONG Value)
{
	mp_yButtonStyle = Value;
}

LONG clsStyle::GetTextAlignmentHorizontal(void)
{
	return mp_yTextAlignmentHorizontal;
}

void clsStyle::SetTextAlignmentHorizontal(LONG Value)
{
	mp_yTextAlignmentHorizontal = Value;
}

LONG clsStyle::GetTextAlignmentVertical(void)
{
	return mp_yTextAlignmentVertical;
}

void clsStyle::SetTextAlignmentVertical(LONG Value)
{
	mp_yTextAlignmentVertical = Value;
}

BOOL clsStyle::GetTextVisible(void)
{
	return mp_bTextVisible;
}

void clsStyle::SetTextVisible(BOOL Value)
{
	mp_bTextVisible = Value;
}

LONG clsStyle::GetTextXMargin(void)
{
	return mp_lTextXMargin;
}

void clsStyle::SetTextXMargin(LONG Value)
{
	mp_lTextXMargin = Value;
}

LONG clsStyle::GetTextYMargin(void)
{
	return mp_lTextYMargin;
}

void clsStyle::SetTextYMargin(LONG Value)
{
	mp_lTextYMargin = Value;
}

BOOL clsStyle::GetClipText(void)
{
	return mp_bClipText;
}

void clsStyle::SetClipText(BOOL Value)
{
	mp_bClipText = Value;
}

clsFont* clsStyle::GetFont(void)
{
	return &mp_oFont;
}

Gdiplus::Color clsStyle::GetForeColor(void)
{
	return mp_clrForeColor;
}

void clsStyle::SetForeColor(Gdiplus::Color Value)
{
	mp_clrForeColor = Value;
}

LONG clsStyle::GetOffsetBottom(void)
{
	return mp_lOffsetBottom;
}

void clsStyle::SetOffsetBottom(LONG Value)
{
	mp_lOffsetBottom = Value;
}

LONG clsStyle::GetOffsetTop(void)
{
	return mp_lOffsetTop;
}

void clsStyle::SetOffsetTop(LONG Value)
{
	mp_lOffsetTop = Value;
}

LONG clsStyle::GetImageAlignmentHorizontal(void)
{
	return mp_yImageAlignmentHorizontal;
}

void clsStyle::SetImageAlignmentHorizontal(LONG Value)
{
	mp_yImageAlignmentHorizontal = Value;
}

LONG clsStyle::GetImageAlignmentVertical(void)
{
	return mp_yImageAlignmentVertical;
}

void clsStyle::SetImageAlignmentVertical(LONG Value)
{
	mp_yImageAlignmentVertical = Value;
}

LONG clsStyle::GetImageXMargin(void)
{
	return mp_lImageXMargin;
}

void clsStyle::SetImageXMargin(LONG Value)
{
	mp_lImageXMargin = Value;
}

LONG clsStyle::GetImageYMargin(void)
{
	return mp_lImageYMargin;
}

void clsStyle::SetImageYMargin(LONG Value)
{
	mp_lImageYMargin = Value;
}

LONG clsStyle::GetPlacement(void)
{
	return mp_yPlacement;
}

void clsStyle::SetPlacement(LONG Value)
{
	mp_yPlacement = Value;
}

BOOL clsStyle::GetUseMask(void)
{
	return mp_bUseMask;
}

void clsStyle::SetUseMask(BOOL Value)
{
	mp_bUseMask = Value;
}

LONG clsStyle::GetTextPlacement(void)
{
	return mp_yTextPlacement;
}

void clsStyle::SetTextPlacement(LONG Value)
{
	mp_yTextPlacement = Value;
}

LONG clsStyle::GetPatternFactor(void)
{
	return mp_lPatternFactor;
}

void clsStyle::SetPatternFactor(LONG Value)
{
	mp_lPatternFactor = Value;
}

LONG clsStyle::GetHatchStyle(void)
{
	return mp_yHatchStyle;
}

void clsStyle::SetHatchStyle(LONG Value)
{
	mp_yHatchStyle = Value;
}

Gdiplus::Color clsStyle::GetStartGradientColor(void)
{
	return mp_clrStartGradientColor;
}

void clsStyle::SetStartGradientColor(Gdiplus::Color Value)
{
	mp_clrStartGradientColor = Value;
}

Gdiplus::Color clsStyle::GetEndGradientColor(void)
{
	return mp_clrEndGradientColor;
}

void clsStyle::SetEndGradientColor(Gdiplus::Color Value)
{
	mp_clrEndGradientColor = Value;
}

LONG clsStyle::GetGradientFillMode(void)
{
	return mp_yGradientFillMode;
}

void clsStyle::SetGradientFillMode(LONG Value)
{
	mp_yGradientFillMode = Value;
}

LONG clsStyle::GetFillMode(void)
{
	return mp_yFillMode;
}

void clsStyle::SetFillMode(LONG Value)
{
	mp_yFillMode = Value;
}

LONG clsStyle::GetBackgroundMode(void)
{
	return mp_yBackgroundMode;
}

void clsStyle::SetBackgroundMode(LONG Value)
{
	mp_yBackgroundMode = Value;
}

BOOL clsStyle::GetDrawTextInVisibleArea(void)
{
	return mp_bDrawTextInVisibleArea;
}

void clsStyle::SetDrawTextInVisibleArea(BOOL Value)
{
	mp_bDrawTextInVisibleArea = Value;
}

CString clsStyle::GetTag(void)
{
	return mp_sTag;
}

void clsStyle::SetTag(CString Value)
{
	mp_sTag = Value;
}

LONG clsStyle::GetBorderWidth(void)
{
	return mp_lBorderWidth;
}

void clsStyle::SetBorderWidth(LONG Value)
{
	mp_lBorderWidth = Value;
}


LONG clsStyle::GetIndex(void)
{
	return mp_lIndex;
}

void clsStyle::SetIndex(LONG Value)
{
	mp_lIndex = Value;
}

clsTaskStyle* clsStyle::GetTaskStyle(void)
{
	return mp_oTaskStyle;
}

clsMilestoneStyle* clsStyle::GetMilestoneStyle(void)
{
	return mp_oMilestoneStyle;
}

clsPredecessorStyle* clsStyle::GetPredecessorStyle(void)
{
	return mp_oPredecessorStyle;
}

clsTextFlags* clsStyle::GetTextFlags(void)
{
	return mp_oTextFlags;
}

clsCustomBorderStyle* clsStyle::GetCustomBorderStyle(void)
{
	return mp_oCustomBorderStyle;
}

clsScrollBarStyle* clsStyle::GetScrollBarStyle(void)
{
	return mp_oScrollBarStyle;
}

clsSelectionRectangleStyle* clsStyle::GetSelectionRectangleStyle(void)
{
	return mp_oSelectionRectangleStyle;
}

clsButtonBorderStyle* clsStyle::GetButtonBorderStyle(void)
{
	return mp_oButtonBorderStyle;
}

CString clsStyle::GetXML(void)
{
	clsXML oXML(mp_oControl, "Style");
	oXML.InitializeWriter();
	oXML.WriteProperty("Appearance", mp_yAppearance);
	oXML.WriteProperty("BackColor", mp_clrBackColor);
	oXML.WriteProperty("BackgroundMode", mp_yBackgroundMode);
	oXML.WriteProperty("BorderColor", mp_clrBorderColor);
	oXML.WriteProperty("BorderStyle", mp_yBorderStyle);
	oXML.WriteProperty("ButtonStyle", mp_yButtonStyle);
	oXML.WriteProperty("ClipText", mp_bClipText);
	oXML.WriteProperty("DrawTextInVisibleArea", mp_bDrawTextInVisibleArea);
	oXML.WriteProperty("EndGradientColor", mp_clrEndGradientColor);
	oXML.WriteProperty("FillMode", mp_yFillMode);
	oXML.WriteProperty("Font", mp_oFont);
	oXML.WriteProperty("ForeColor", mp_clrForeColor);
	oXML.WriteProperty("GradientFillMode", mp_yGradientFillMode);
	oXML.WriteProperty("HatchBackColor", mp_clrHatchBackColor);
	oXML.WriteProperty("HatchForeColor", mp_clrHatchForeColor);
	oXML.WriteProperty("HatchStyle", mp_yHatchStyle);
	oXML.WriteProperty("ImageAlignmentHorizontal", mp_yImageAlignmentHorizontal);
	oXML.WriteProperty("ImageAlignmentVertical", mp_yImageAlignmentVertical);
	oXML.WriteProperty("ImageXMargin", mp_lImageXMargin);
	oXML.WriteProperty("ImageYMargin", mp_lImageYMargin);
	oXML.WriteProperty("Key", mp_sKey);
	oXML.WriteProperty("OffsetBottom", mp_lOffsetBottom);
	oXML.WriteProperty("OffsetTop", mp_lOffsetTop);
	oXML.WriteProperty("Pattern", mp_yPattern);
	oXML.WriteProperty("PatternFactor", mp_lPatternFactor);
	oXML.WriteProperty("Placement", mp_yPlacement);
	oXML.WriteProperty("StartGradientColor", mp_clrStartGradientColor);
	oXML.WriteProperty("Tag", mp_sTag);
	oXML.WriteProperty("BorderWidth", mp_lBorderWidth);
	oXML.WriteProperty("TextAlignmentHorizontal", mp_yTextAlignmentHorizontal);
	oXML.WriteProperty("TextAlignmentVertical", mp_yTextAlignmentVertical);
	oXML.WriteProperty("TextPlacement", mp_yTextPlacement);
	oXML.WriteProperty("TextVisible", mp_bTextVisible);
	oXML.WriteProperty("TextXMargin", mp_lTextXMargin);
	oXML.WriteProperty("TextYMargin", mp_lTextYMargin);
	oXML.WriteProperty("UseMask", mp_bUseMask);
	oXML.WriteProperty("TextEditBackColor", mp_clrTextEditBackColor);
	oXML.WriteProperty("TextEditForeColor", mp_clrTextEditForeColor);
	oXML.WriteObject(mp_oCustomBorderStyle->GetXML());
	oXML.WriteObject(mp_oMilestoneStyle->GetXML());
	oXML.WriteObject(mp_oPredecessorStyle->GetXML());
	oXML.WriteObject(mp_oTaskStyle->GetXML());
	oXML.WriteObject(mp_oTextFlags->GetXML());
	oXML.WriteObject(mp_oScrollBarStyle->GetXML());
	oXML.WriteObject(mp_oSelectionRectangleStyle->GetXML());
	oXML.WriteObject(mp_oButtonBorderStyle->GetXML());
	return oXML.GetXML();
}

void clsStyle::SetXML(CString sXML)
{
	clsXML oXML(mp_oControl, "Style");
	oXML.SetXML(sXML);
	oXML.InitializeReader();
	oXML.ReadProperty("Appearance", mp_yAppearance);
	oXML.ReadProperty("BackColor", mp_clrBackColor);
	oXML.ReadProperty("BackgroundMode", mp_yBackgroundMode);
	oXML.ReadProperty("BorderColor", mp_clrBorderColor);
	oXML.ReadProperty("BorderStyle", mp_yBorderStyle);
	oXML.ReadProperty("ButtonStyle", mp_yButtonStyle);
	oXML.ReadProperty("ClipText", mp_bClipText);
	oXML.ReadProperty("DrawTextInVisibleArea", mp_bDrawTextInVisibleArea);
	oXML.ReadProperty("EndGradientColor", mp_clrEndGradientColor);
	oXML.ReadProperty("FillMode", mp_yFillMode);
	oXML.ReadProperty("Font", mp_oFont);
	oXML.ReadProperty("ForeColor", mp_clrForeColor);
	oXML.ReadProperty("GradientFillMode", mp_yGradientFillMode);
	oXML.ReadProperty("HatchBackColor", mp_clrHatchBackColor);
	oXML.ReadProperty("HatchForeColor", mp_clrHatchForeColor);
	oXML.ReadProperty("HatchStyle", mp_yHatchStyle);
	oXML.ReadProperty("ImageAlignmentHorizontal", mp_yImageAlignmentHorizontal);
	oXML.ReadProperty("ImageAlignmentVertical", mp_yImageAlignmentVertical);
	oXML.ReadProperty("ImageXMargin", mp_lImageXMargin);
	oXML.ReadProperty("ImageYMargin", mp_lImageYMargin);
	oXML.ReadProperty("Key", mp_sKey);
	oXML.ReadProperty("OffsetBottom", mp_lOffsetBottom);
	oXML.ReadProperty("OffsetTop", mp_lOffsetTop);
	oXML.ReadProperty("Pattern", mp_yPattern);
	oXML.ReadProperty("PatternFactor", mp_lPatternFactor);
	oXML.ReadProperty("Placement", mp_yPlacement);
	oXML.ReadProperty("StartGradientColor", mp_clrStartGradientColor);
	oXML.ReadProperty("Tag", mp_sTag);
	oXML.ReadProperty("BorderWidth", mp_lBorderWidth);
	oXML.ReadProperty("TextAlignmentHorizontal", mp_yTextAlignmentHorizontal);
	oXML.ReadProperty("TextAlignmentVertical", mp_yTextAlignmentVertical);
	oXML.ReadProperty("TextPlacement", mp_yTextPlacement);
	oXML.ReadProperty("TextVisible", mp_bTextVisible);
	oXML.ReadProperty("TextXMargin", mp_lTextXMargin);
	oXML.ReadProperty("TextYMargin", mp_lTextYMargin);
	oXML.ReadProperty("UseMask", mp_bUseMask);
	oXML.ReadProperty("TextEditBackColor", mp_clrTextEditBackColor);
	oXML.ReadProperty("TextEditForeColor", mp_clrTextEditForeColor);
	mp_oCustomBorderStyle->SetXML(oXML.ReadObject("CustomBorderStyle"));
	mp_oMilestoneStyle->SetXML(oXML.ReadObject("MilestoneStyle"));
	mp_oPredecessorStyle->SetXML(oXML.ReadObject("PredecessorStyle"));
	mp_oTaskStyle->SetXML(oXML.ReadObject("TaskStyle"));
	mp_oTextFlags->SetXML(oXML.ReadObject("TextFlags"));
	mp_oScrollBarStyle->SetXML(oXML.ReadObject("ScrollBarStyle"));
	mp_oSelectionRectangleStyle->SetXML(oXML.ReadObject("SelectionRectangleStyle"));
	mp_oButtonBorderStyle->SetXML(oXML.ReadObject("ButtonBorderStyle"));
}

void clsStyle::Clear(void)
{
	mp_bTextVisible = TRUE;
	mp_bClipText = TRUE;
	mp_bDrawTextInVisibleArea = FALSE;
	mp_bUseMask = TRUE;
	mp_clrBackColor = Color::MakeARGB(255, 192, 192, 192);
	mp_clrBorderColor = Color::Black;
	mp_clrEndGradientColor = Color::Black;
	mp_clrForeColor = Color::Black;
	mp_clrStartGradientColor = Color::Black;
	mp_lPatternFactor = 5;
	mp_lTextXMargin = 0;
	mp_lTextYMargin = 0;
	mp_lOffsetBottom = 10;
	mp_lOffsetTop = 10;
	mp_lImageXMargin = 3;
	mp_lImageYMargin = 3;
	mp_sTag = "";
	mp_lBorderWidth = 1;
	mp_yAppearance = SA_RAISED;
	mp_yPattern = FP_DOWNWARDDIAGONAL;
	mp_yBorderStyle = SBR_NONE;
	mp_yButtonStyle = BT_NORMALWINDOWS;
	mp_yTextAlignmentHorizontal = HAL_CENTER;
	mp_yTextAlignmentVertical = VAL_CENTER;
	mp_yTextPlacement = SCP_OBJECTEXTENTSPLACEMENT;
	mp_yGradientFillMode = GDT_HORIZONTAL;
	mp_yImageAlignmentHorizontal = HAL_CENTER;
	mp_yImageAlignmentVertical = VAL_CENTER;
	mp_yPlacement = PLC_ROWEXTENTSPLACEMENT;
	mp_yFillMode = FM_COMPLETELYFILLED;
	mp_oControl->GetConfiguration()->GetDefaultFont()->Clone(&mp_oFont);
	mp_yBackgroundMode = FP_SOLID;
	mp_yHatchStyle = HS_HORIZONTAL;
	mp_clrHatchBackColor = Color::White;
	mp_clrHatchForeColor = Color::Black;
	mp_clrTextEditBackColor = Color::White;
	mp_clrTextEditForeColor = Color::Black;
    GetCustomBorderStyle()->Clear();
    GetMilestoneStyle()->Clear();
    GetPredecessorStyle()->Clear();
    GetTaskStyle()->Clear();
    GetTextFlags()->Clear();
    GetScrollBarStyle()->Clear();
    GetSelectionRectangleStyle()->Clear();
    GetButtonBorderStyle()->Clear();
}

clsStyle* clsStyle::Clone(CString Key)
{
    clsStyle* oClone;
    oClone = mp_oControl->GetStyles()->Add(Key);
    oClone->SetAppearance(mp_yAppearance);
    oClone->SetBackColor(mp_clrBackColor);
    oClone->SetBackgroundMode(mp_yBackgroundMode);
    oClone->SetBorderColor(mp_clrBorderColor);
    oClone->SetBorderStyle(mp_yBorderStyle);
    oClone->SetButtonStyle(mp_yButtonStyle);
    oClone->SetClipText(mp_bClipText);
    oClone->SetDrawTextInVisibleArea(mp_bDrawTextInVisibleArea);
    oClone->SetEndGradientColor(mp_clrEndGradientColor);
    oClone->SetFillMode(mp_yFillMode);
    mp_oFont.Clone(&oClone->mp_oFont);
    oClone->SetForeColor(mp_clrForeColor);
    oClone->SetGradientFillMode(mp_yGradientFillMode);
    oClone->SetHatchBackColor(mp_clrHatchBackColor);
    oClone->SetHatchForeColor(mp_clrHatchForeColor);
    oClone->SetHatchStyle(mp_yHatchStyle);
    oClone->SetImageAlignmentHorizontal(mp_yImageAlignmentHorizontal);
    oClone->SetImageAlignmentVertical(mp_yImageAlignmentVertical);
    oClone->SetImageXMargin(mp_lImageXMargin);
    oClone->SetImageYMargin(mp_lImageYMargin);
    //Key
    oClone->SetOffsetBottom(mp_lOffsetBottom);
    oClone->SetOffsetTop(mp_lOffsetTop);
    oClone->SetPattern(mp_yPattern);
    oClone->SetPatternFactor(mp_lPatternFactor);
    oClone->SetPlacement(mp_yPlacement);
    oClone->SetStartGradientColor(mp_clrStartGradientColor);
    oClone->SetTag(mp_sTag);
    oClone->SetBorderWidth(mp_lBorderWidth);
    oClone->SetTextAlignmentHorizontal(mp_yTextAlignmentHorizontal);
    oClone->SetTextAlignmentVertical(mp_yTextAlignmentVertical);
    oClone->SetTextPlacement(mp_yTextPlacement);
    oClone->SetTextVisible(mp_bTextVisible);
    oClone->SetTextXMargin(mp_lTextXMargin);
    oClone->SetTextYMargin(mp_lTextYMargin);
    oClone->SetUseMask(mp_bUseMask);
    oClone->SetTextEditBackColor(mp_clrTextEditBackColor);
    oClone->SetTextEditForeColor(mp_clrTextEditForeColor);
    GetCustomBorderStyle()->Clone(oClone->GetCustomBorderStyle());
    GetMilestoneStyle()->Clone(oClone->GetMilestoneStyle());
    GetPredecessorStyle()->Clone(oClone->GetPredecessorStyle());
    GetTaskStyle()->Clone(oClone->GetTaskStyle());
    GetTextFlags()->Clone(oClone->GetTextFlags());
    GetScrollBarStyle()->Clone(oClone->GetScrollBarStyle());
    GetSelectionRectangleStyle()->Clone(oClone->GetSelectionRectangleStyle());
    GetButtonBorderStyle()->Clone(oClone->GetButtonBorderStyle());
    return oClone;
}