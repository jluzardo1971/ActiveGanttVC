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
#include "clsSelectionRectangleStyle.h"


// clsSelectionRectangleStyle

IMPLEMENT_DYNCREATE(clsSelectionRectangleStyle, CCmdTarget)

clsSelectionRectangleStyle::~clsSelectionRectangleStyle()
{
	// To terminate the application when all objects created with
	// 	with OLE automation, the destructor calls AfxOleUnlockApp.
	
	AfxOleUnlockApp();
}


void clsSelectionRectangleStyle::OnFinalRelease()
{
	// When the last reference for an automation object is released
	// OnFinalRelease is called.  The base class will automatically
	// deletes the object.  Add additional cleanup required for your
	// object before calling the base class.

	CCmdTarget::OnFinalRelease();
}


BEGIN_MESSAGE_MAP(clsSelectionRectangleStyle, CCmdTarget)
END_MESSAGE_MAP()


BEGIN_DISPATCH_MAP(clsSelectionRectangleStyle, CCmdTarget)
	DISP_PROPERTY_EX_ID(clsSelectionRectangleStyle, "OffsetBottom", 1, odl_GetOffsetBottom, odl_SetOffsetBottom, VT_I4) 
	DISP_PROPERTY_EX_ID(clsSelectionRectangleStyle, "OffsetLeft", 2, odl_GetOffsetLeft, odl_SetOffsetLeft, VT_I4) 
	DISP_PROPERTY_EX_ID(clsSelectionRectangleStyle, "OffsetRight", 3, odl_GetOffsetRight, odl_SetOffsetRight, VT_I4) 
	DISP_PROPERTY_EX_ID(clsSelectionRectangleStyle, "OffsetTop", 4, odl_GetOffsetTop, odl_SetOffsetTop, VT_I4) 
	DISP_PROPERTY_EX_ID(clsSelectionRectangleStyle, "Visible", 5, odl_GetVisible, odl_SetVisible, VT_BOOL)
	DISP_PROPERTY_EX_ID(clsSelectionRectangleStyle, "Mode", 6, odl_GetMode, odl_SetMode, VT_I4)
	DISP_PROPERTY_EX_ID(clsSelectionRectangleStyle, "BorderWidth", 7, odl_GetBorderWidth, odl_SetBorderWidth, VT_I4)
	DISP_PROPERTY_EX_ID(clsSelectionRectangleStyle, "Color", 8, odl_GetColor, odl_SetColor, VT_COLOR)
	DISP_FUNCTION_ID(clsSelectionRectangleStyle, "GetXML", 9, odl_GetXML, VT_BSTR, VTS_NONE) 
	DISP_FUNCTION_ID(clsSelectionRectangleStyle, "SetXML", 10, odl_SetXML, VT_EMPTY, VTS_BSTR)
END_DISPATCH_MAP()

// Note: we add support for IID_IclsSelectionRectangleStyle to support typesafe binding
//  from VBA.  This IID must match the GUID that is attached to the 
//  dispinterface in the .IDL file.

//{30732742-F035-474C-A5C0-0E41E73F2DF5}
static const IID IID_IclsSelectionRectangleStyle = { 0x30732742, 0xF035, 0x474C, { 0xA5, 0xC0, 0x0E, 0x41, 0xE7, 0x3F, 0x2D, 0xF5} };


BEGIN_INTERFACE_MAP(clsSelectionRectangleStyle, CCmdTarget)
	INTERFACE_PART(clsSelectionRectangleStyle, IID_IclsSelectionRectangleStyle, Dispatch)
END_INTERFACE_MAP()

//{4839E0DA-1418-42E6-8638-8EA94E09970A}
IMPLEMENT_OLECREATE_FLAGS(clsSelectionRectangleStyle, "AGVC.clsSelectionRectangleStyle", afxRegApartmentThreading, 0x4839E0DA, 0x1418, 0x42E6, 0x86, 0x38, 0x8E, 0xA9, 0x4E, 0x09, 0x97, 0x0A)
                            

// clsSelectionRectangleStyle message handlers

clsSelectionRectangleStyle::clsSelectionRectangleStyle()
{
	AfxMessageBox(_T("Invalid constructor has been invoked for clsSelectionRectangleStyle. Please inform support@sourcecodestore.com of this error"), MB_OK | MB_ICONSTOP | MB_APPLMODAL);
}

clsSelectionRectangleStyle::clsSelectionRectangleStyle(CActiveGanttVCCtl* oControl)
{
	EnableAutomation();
	AfxOleLockApp();
	mp_oControl = oControl;
	Clear();
}

LONG clsSelectionRectangleStyle::GetOffsetBottom(void)
{
	return mp_lOffsetBottom;
}

void clsSelectionRectangleStyle::SetOffsetBottom(LONG Value)
{
	mp_lOffsetBottom = Value;
}

LONG clsSelectionRectangleStyle::GetOffsetLeft(void)
{
	return mp_lOffsetLeft;
}

void clsSelectionRectangleStyle::SetOffsetLeft(LONG Value)
{
	mp_lOffsetLeft = Value;
}

LONG clsSelectionRectangleStyle::GetOffsetRight(void)
{
	return mp_lOffsetRight;
}

void clsSelectionRectangleStyle::SetOffsetRight(LONG Value)
{
	mp_lOffsetRight = Value;
}

LONG clsSelectionRectangleStyle::GetOffsetTop(void)
{
	return mp_lOffsetTop;
}

void clsSelectionRectangleStyle::SetOffsetTop(LONG Value)
{
	mp_lOffsetTop = Value;
}

BOOL clsSelectionRectangleStyle::GetVisible(void)
{
	return mp_bVisible;
}

void clsSelectionRectangleStyle::SetVisible(BOOL Value)
{
	mp_bVisible = Value;
}

LONG clsSelectionRectangleStyle::GetMode(void)
{
	return mp_yMode;
}

void clsSelectionRectangleStyle::SetMode(LONG Value)
{
	mp_yMode = Value;
}

LONG clsSelectionRectangleStyle::GetBorderWidth(void)
{
	return mp_lBorderWidth;
}

void clsSelectionRectangleStyle::SetBorderWidth(LONG Value)
{
	mp_lBorderWidth = Value;
}

Gdiplus::Color clsSelectionRectangleStyle::GetColor(void)
{
	return mp_clrColor;
}

void clsSelectionRectangleStyle::SetColor(Gdiplus::Color Value)
{
	mp_clrColor = Value;
}

CString clsSelectionRectangleStyle::GetXML(void)
{
	clsXML oXML(mp_oControl, "SelectionRectangleStyle");
	oXML.InitializeWriter();
    oXML.WriteProperty("OffsetBottom", mp_lOffsetBottom);
    oXML.WriteProperty("OffsetLeft", mp_lOffsetLeft);
    oXML.WriteProperty("OffsetRight", mp_lOffsetRight);
    oXML.WriteProperty("OffsetTop", mp_lOffsetTop);
    oXML.WriteProperty("Visible", mp_bVisible);
    oXML.WriteProperty("Mode", mp_yMode);
    oXML.WriteProperty("BorderWidth", mp_lBorderWidth);
    oXML.WriteProperty("Color", mp_clrColor);
	return oXML.GetXML();
}

void clsSelectionRectangleStyle::SetXML(CString sXML)
{
	clsXML oXML(mp_oControl, "SelectionRectangleStyle");
	oXML.SetXML(sXML);
	oXML.InitializeReader();
    oXML.ReadProperty("OffsetBottom", mp_lOffsetBottom);
    oXML.ReadProperty("OffsetLeft", mp_lOffsetLeft);
    oXML.ReadProperty("OffsetRight", mp_lOffsetRight);
    oXML.ReadProperty("OffsetTop", mp_lOffsetTop);
    oXML.ReadProperty("Visible", mp_bVisible);
    oXML.ReadProperty("Mode", mp_yMode);
    oXML.ReadProperty("BorderWidth", mp_lBorderWidth);
    oXML.ReadProperty("Color", mp_clrColor);
}

void clsSelectionRectangleStyle::Clear()
{
	mp_lOffsetBottom = 3;
	mp_lOffsetLeft = 3;
	mp_lOffsetRight = 3;
	mp_lOffsetTop = 3;
	mp_bVisible = TRUE;
    mp_yMode = SRM_DOTTED;
    mp_lBorderWidth = 1;
	mp_clrColor = Color::Black;
}

void clsSelectionRectangleStyle::Clone(clsSelectionRectangleStyle* oClone)
{
    oClone->SetOffsetBottom(mp_lOffsetBottom);
    oClone->SetOffsetLeft(mp_lOffsetLeft);
    oClone->SetOffsetRight(mp_lOffsetRight);
    oClone->SetOffsetTop(mp_lOffsetTop);
    oClone->SetVisible(mp_bVisible);
    oClone->SetMode(mp_yMode);
    oClone->SetBorderWidth(mp_lBorderWidth);
    oClone->SetColor(mp_clrColor);
}