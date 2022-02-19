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
#include "clsScrollBarStyle.h"


// clsScrollBarStyle

IMPLEMENT_DYNCREATE(clsScrollBarStyle, CCmdTarget)


clsScrollBarStyle::~clsScrollBarStyle()
{
	// To terminate the application when all objects created with
	// 	with OLE automation, the destructor calls AfxOleUnlockApp.
	
	AfxOleUnlockApp();
}


void clsScrollBarStyle::OnFinalRelease()
{
	// When the last reference for an automation object is released
	// OnFinalRelease is called.  The base class will automatically
	// deletes the object.  Add additional cleanup required for your
	// object before calling the base class.

	CCmdTarget::OnFinalRelease();
}


BEGIN_MESSAGE_MAP(clsScrollBarStyle, CCmdTarget)
END_MESSAGE_MAP()

BEGIN_DISPATCH_MAP(clsScrollBarStyle, CCmdTarget)
	DISP_PROPERTY_EX_ID(clsScrollBarStyle, "ArrowColor", 1, odl_GetArrowColor, odl_SetArrowColor, VT_COLOR) 
	DISP_PROPERTY_EX_ID(clsScrollBarStyle, "DropShadowArrowColor", 2, odl_GetDropShadowArrowColor, odl_SetDropShadowArrowColor, VT_COLOR)
	DISP_PROPERTY_EX_ID(clsScrollBarStyle, "DropShadow", 3, odl_GetDropShadow, odl_SetDropShadow, VT_BOOL)
	DISP_PROPERTY_EX_ID(clsScrollBarStyle, "LeftX", 4, odl_GetLeftX, odl_SetLeftX, VT_I4)
	DISP_PROPERTY_EX_ID(clsScrollBarStyle, "LeftY", 5, odl_GetLeftY, odl_SetLeftY, VT_I4)
	DISP_PROPERTY_EX_ID(clsScrollBarStyle, "UpX", 6, odl_GetUpX, odl_SetUpX, VT_I4)
	DISP_PROPERTY_EX_ID(clsScrollBarStyle, "UpY", 7, odl_GetUpY, odl_SetUpY, VT_I4)
	DISP_PROPERTY_EX_ID(clsScrollBarStyle, "RightX", 8, odl_GetRightX, odl_SetRightX, VT_I4)
	DISP_PROPERTY_EX_ID(clsScrollBarStyle, "RightY", 9, odl_GetRightY, odl_SetRightY, VT_I4)
	DISP_PROPERTY_EX_ID(clsScrollBarStyle, "DownX", 10, odl_GetDownX, odl_SetDownX, VT_I4)
	DISP_PROPERTY_EX_ID(clsScrollBarStyle, "DownY", 11, odl_GetDownY, odl_SetDownY, VT_I4)
	DISP_PROPERTY_EX_ID(clsScrollBarStyle, "DropShadowLeftX", 12, odl_GetDropShadowLeftX, odl_SetDropShadowLeftX, VT_I4)
	DISP_PROPERTY_EX_ID(clsScrollBarStyle, "DropShadowLeftY", 13, odl_GetDropShadowLeftY, odl_SetDropShadowLeftY, VT_I4)
	DISP_PROPERTY_EX_ID(clsScrollBarStyle, "DropShadowUpX", 14, odl_GetDropShadowUpX, odl_SetDropShadowUpX, VT_I4)
	DISP_PROPERTY_EX_ID(clsScrollBarStyle, "DropShadowUpY", 15, odl_GetDropShadowUpY, odl_SetDropShadowUpY, VT_I4)
	DISP_PROPERTY_EX_ID(clsScrollBarStyle, "DropShadowRightX", 16, odl_GetDropShadowRightX, odl_SetDropShadowRightX, VT_I4)
	DISP_PROPERTY_EX_ID(clsScrollBarStyle, "DropShadowRightY", 17, odl_GetDropShadowRightY, odl_SetDropShadowRightY, VT_I4)
	DISP_PROPERTY_EX_ID(clsScrollBarStyle, "DropShadowDownX", 18, odl_GetDropShadowDownX, odl_SetDropShadowDownX, VT_I4)
	DISP_PROPERTY_EX_ID(clsScrollBarStyle, "DropShadowDownY", 19, odl_GetDropShadowDownY, odl_SetDropShadowDownY, VT_I4)
	DISP_PROPERTY_EX_ID(clsScrollBarStyle, "ArrowSize", 20, odl_GetArrowSize, odl_SetArrowSize, VT_I4)
	DISP_FUNCTION_ID(clsScrollBarStyle, "GetXML", 21, odl_GetXML, VT_BSTR, VTS_NONE) 
	DISP_FUNCTION_ID(clsScrollBarStyle, "SetXML", 22, odl_SetXML, VT_EMPTY, VTS_BSTR)
END_DISPATCH_MAP()

// Note: we add support for IID_IclsScrollBarStyle to support typesafe binding
//  from VBA.  This IID must match the GUID that is attached to the 
//  dispinterface in the .IDL file.

//{9163905E-8DDC-4137-8819-29FC868B5A36}
static const IID IID_IclsScrollBarStyle = { 0x9163905E, 0x8DDC, 0x4137, { 0x88, 0x19, 0x29, 0xFC, 0x86, 0x8B, 0x5A, 0x36} };


BEGIN_INTERFACE_MAP(clsScrollBarStyle, CCmdTarget)
	INTERFACE_PART(clsScrollBarStyle, IID_IclsScrollBarStyle, Dispatch)
END_INTERFACE_MAP()

//{5417354B-38CC-46C6-853E-C31CBE10AFCF}
IMPLEMENT_OLECREATE_FLAGS(clsScrollBarStyle, "AGVC.clsScrollBarStyle", afxRegApartmentThreading, 0x5417354B, 0x38CC, 0x46C6, 0x85, 0x3E, 0xC3, 0x1C, 0xBE, 0x10, 0xAF, 0xCF)

// clsScrollBarStyle message handlers

clsScrollBarStyle::clsScrollBarStyle()
{
	AfxMessageBox(_T("Invalid constructor has been invoked for clsScrollBarStyle. Please inform support@sourcecodestore.com of this error"), MB_OK | MB_ICONSTOP | MB_APPLMODAL);
}

clsScrollBarStyle::clsScrollBarStyle(CActiveGanttVCCtl* oControl)
{
	EnableAutomation();
	AfxOleLockApp();
	mp_oControl = oControl;
	Clear();
}

void clsScrollBarStyle::SetArrowColor(Gdiplus::Color Value)
{
	mp_clrArrowColor = Value;
}

Gdiplus::Color clsScrollBarStyle::GetArrowColor(void)
{
	return mp_clrArrowColor;
}

void clsScrollBarStyle::SetDropShadowArrowColor(Gdiplus::Color Value)
{
	mp_clrDropShadowArrowColor = Value;
}

Gdiplus::Color clsScrollBarStyle::GetDropShadowArrowColor(void)
{
	return mp_clrDropShadowArrowColor;
}

BOOL clsScrollBarStyle::GetDropShadow(void)
{
	return mp_bDropShadow;
}

void clsScrollBarStyle::SetDropShadow(BOOL Value)
{
	mp_bDropShadow = Value;
}

LONG clsScrollBarStyle::GetLeftX(void)
{
	return mp_lLeftX;
}

void clsScrollBarStyle::SetLeftX(LONG Value)
{
	mp_lLeftX = Value;
}

LONG clsScrollBarStyle::GetLeftY(void)
{
	return mp_lLeftY;
}

void clsScrollBarStyle::SetLeftY(LONG Value)
{
	mp_lLeftY = Value;
}

LONG clsScrollBarStyle::GetUpX(void)
{
	return mp_lUpX;
}

void clsScrollBarStyle::SetUpX(LONG Value)
{
	mp_lUpX = Value;
}

LONG clsScrollBarStyle::GetUpY(void)
{
	return mp_lUpY;
}

void clsScrollBarStyle::SetUpY(LONG Value)
{
	mp_lUpY = Value;
}

LONG clsScrollBarStyle::GetRightX(void)
{
	return mp_lRightX;
}

void clsScrollBarStyle::SetRightX(LONG Value)
{
	mp_lRightX = Value;
}

LONG clsScrollBarStyle::GetRightY(void)
{
	return mp_lRightY;
}

void clsScrollBarStyle::SetRightY(LONG Value)
{
	mp_lRightY = Value;
}

LONG clsScrollBarStyle::GetDownX(void)
{
	return mp_lDownX;
}

void clsScrollBarStyle::SetDownX(LONG Value)
{
	mp_lDownX = Value;
}

LONG clsScrollBarStyle::GetDownY(void)
{
	return mp_lDownY;
}

void clsScrollBarStyle::SetDownY(LONG Value)
{
	mp_lDownY = Value;
}

LONG clsScrollBarStyle::GetDropShadowLeftX(void)
{
	return mp_lDropShadowLeftX;
}

void clsScrollBarStyle::SetDropShadowLeftX(LONG Value)
{
	mp_lDropShadowLeftX = Value;
}

LONG clsScrollBarStyle::GetDropShadowLeftY(void)
{
	return mp_lDropShadowLeftY;
}

void clsScrollBarStyle::SetDropShadowLeftY(LONG Value)
{
	mp_lDropShadowLeftY = Value;
}

LONG clsScrollBarStyle::GetDropShadowUpX(void)
{
	return mp_lDropShadowUpX;
}

void clsScrollBarStyle::SetDropShadowUpX(LONG Value)
{
	mp_lDropShadowUpX = Value;
}

LONG clsScrollBarStyle::GetDropShadowUpY(void)
{
	return mp_lDropShadowUpY;
}

void clsScrollBarStyle::SetDropShadowUpY(LONG Value)
{
	mp_lDropShadowUpY = Value;
}

LONG clsScrollBarStyle::GetDropShadowRightX(void)
{
	return mp_lDropShadowRightX;
}

void clsScrollBarStyle::SetDropShadowRightX(LONG Value)
{
	mp_lDropShadowRightX = Value;
}

LONG clsScrollBarStyle::GetDropShadowRightY(void)
{
	return mp_lDropShadowRightY;
}

void clsScrollBarStyle::SetDropShadowRightY(LONG Value)
{
	mp_lDropShadowRightY = Value;
}

LONG clsScrollBarStyle::GetDropShadowDownX(void)
{
	return mp_lDropShadowDownX;
}

void clsScrollBarStyle::SetDropShadowDownX(LONG Value)
{
	mp_lDropShadowDownX = Value;
}

LONG clsScrollBarStyle::GetDropShadowDownY(void)
{
	return mp_lDropShadowDownY;
}

void clsScrollBarStyle::SetDropShadowDownY(LONG Value)
{
	mp_lDropShadowDownY = Value;
}

LONG clsScrollBarStyle::GetArrowSize(void)
{
	return mp_lArrowSize;
}

void clsScrollBarStyle::SetArrowSize(LONG Value)
{
	mp_lArrowSize = Value;
}

CString clsScrollBarStyle::GetXML(void)
{
	clsXML oXML(mp_oControl, "ScrollBarStyle");
	oXML.InitializeWriter();
    oXML.WriteProperty("ArrowColor", mp_clrArrowColor);
    oXML.WriteProperty("DropShadowArrowColor", mp_clrDropShadowArrowColor);
    oXML.WriteProperty("DropShadow", mp_bDropShadow);
    oXML.WriteProperty("ArrowSize", mp_lArrowSize);
    oXML.WriteProperty("LeftX", mp_lLeftX);
    oXML.WriteProperty("LeftY", mp_lLeftY);
    oXML.WriteProperty("UpX", mp_lUpX);
    oXML.WriteProperty("UpY", mp_lUpY);
    oXML.WriteProperty("RightX", mp_lRightX);
    oXML.WriteProperty("RightY", mp_lRightY);
    oXML.WriteProperty("DownX", mp_lDownX);
    oXML.WriteProperty("DownY", mp_lDownY);
    oXML.WriteProperty("DropShadowLeftX", mp_lDropShadowLeftX);
    oXML.WriteProperty("DropShadowLeftY", mp_lDropShadowLeftY);
    oXML.WriteProperty("DropShadowUpX", mp_lDropShadowUpX);
    oXML.WriteProperty("DropShadowUpY", mp_lDropShadowUpY);
    oXML.WriteProperty("DropShadowRightX", mp_lDropShadowRightX);
    oXML.WriteProperty("DropShadowRightY", mp_lDropShadowRightY);
    oXML.WriteProperty("DropShadowDownX", mp_lDropShadowDownX);
    oXML.WriteProperty("DropShadowDownY", mp_lDropShadowDownY);
	return oXML.GetXML();
}

void clsScrollBarStyle::SetXML(CString sXML)
{
	clsXML oXML(mp_oControl, "ScrollBarStyle");
	oXML.SetXML(sXML);
	oXML.InitializeReader();
    oXML.ReadProperty("ArrowColor", mp_clrArrowColor);
    oXML.ReadProperty("DropShadowArrowColor", mp_clrDropShadowArrowColor);
    oXML.ReadProperty("DropShadow", mp_bDropShadow);
    oXML.ReadProperty("ArrowSize", mp_lArrowSize);
    oXML.ReadProperty("LeftX", mp_lLeftX);
    oXML.ReadProperty("LeftY", mp_lLeftY);
    oXML.ReadProperty("UpX", mp_lUpX);
    oXML.ReadProperty("UpY", mp_lUpY);
    oXML.ReadProperty("RightX", mp_lRightX);
    oXML.ReadProperty("RightY", mp_lRightY);
    oXML.ReadProperty("DownX", mp_lDownX);
    oXML.ReadProperty("DownY", mp_lDownY);
    oXML.ReadProperty("DropShadowLeftX", mp_lDropShadowLeftX);
    oXML.ReadProperty("DropShadowLeftY", mp_lDropShadowLeftY);
	oXML.ReadProperty("DropShadowUpX", mp_lDropShadowUpX);
    oXML.ReadProperty("DropShadowUpY", mp_lDropShadowUpY);
    oXML.ReadProperty("DropShadowRightX", mp_lDropShadowRightX);
    oXML.ReadProperty("DropShadowRightY", mp_lDropShadowRightY);
    oXML.ReadProperty("DropShadowDownX", mp_lDropShadowDownX);
    oXML.ReadProperty("DropShadowDownY", mp_lDropShadowDownY);
}

void clsScrollBarStyle::Clear()
{
    mp_clrArrowColor = Color::Black;
    mp_clrDropShadowArrowColor = Color::MakeARGB(255, 192, 192, 192);
    mp_bDropShadow = FALSE;
    mp_lArrowSize = 3;
    mp_lLeftX = 6;
    mp_lLeftY = 8;
    mp_lUpX = 8;
    mp_lUpY = 6;
    mp_lRightX = 10;
    mp_lRightY = 8;
    mp_lDownX = 8;
    mp_lDownY = 10;
    mp_lDropShadowLeftX = 5;
    mp_lDropShadowLeftY = 7;
    mp_lDropShadowUpX = 7;
    mp_lDropShadowUpY = 5;
    mp_lDropShadowRightX = 9;
    mp_lDropShadowRightY = 7;
    mp_lDropShadowDownX = 7;
    mp_lDropShadowDownY = 9;
}

void clsScrollBarStyle::Clone(clsScrollBarStyle* oClone)
{
    oClone->SetArrowColor(mp_clrArrowColor);
    oClone->SetDropShadowArrowColor(mp_clrDropShadowArrowColor);
    oClone->SetDropShadow(mp_bDropShadow);
    oClone->SetArrowSize(mp_lArrowSize);
    oClone->SetLeftX(mp_lLeftX);
    oClone->SetLeftY(mp_lLeftY);
    oClone->SetUpX(mp_lUpX);
    oClone->SetUpY(mp_lUpY);
    oClone->SetRightX(mp_lRightX);
    oClone->SetRightY(mp_lRightY);
    oClone->SetDownX(mp_lDownX);
    oClone->SetDownY(mp_lDownY);
    oClone->SetDropShadowLeftX(mp_lDropShadowLeftX);
    oClone->SetDropShadowLeftY(mp_lDropShadowLeftY);
    oClone->SetDropShadowUpX(mp_lDropShadowUpX);
    oClone->SetDropShadowUpY(mp_lDropShadowUpY);
    oClone->SetDropShadowRightX(mp_lDropShadowRightX);
    oClone->SetDropShadowRightY(mp_lDropShadowRightY);
    oClone->SetDropShadowDownX(mp_lDropShadowDownX);
    oClone->SetDropShadowDownY(mp_lDropShadowDownY);
}