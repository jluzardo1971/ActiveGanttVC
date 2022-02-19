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
#include "clsCustomBorderStyle.h"

IMPLEMENT_DYNCREATE(clsCustomBorderStyle, CCmdTarget)

//{CF1309C0-A15C-4230-B30B-76600DE745D4}
static const IID IID_IclsCustomBorderStyle = { 0xCF1309C0, 0xA15C, 0x4230, { 0xB3, 0x0B, 0x76, 0x60, 0x0D, 0xE7, 0x45, 0xD4} };

//{98D4BC15-AFD4-4570-8A3B-2ED32A6F4B6D}
IMPLEMENT_OLECREATE_FLAGS(clsCustomBorderStyle, "AGVC.clsCustomBorderStyle", afxRegApartmentThreading, 0x98D4BC15, 0xAFD4, 0x4570, 0x8A, 0x3B, 0x2E, 0xD3, 0x2A, 0x6F, 0x4B, 0x6D)

BEGIN_DISPATCH_MAP(clsCustomBorderStyle, CCmdTarget)
	DISP_PROPERTY_EX_ID(clsCustomBorderStyle, "Left", 1, odl_GetLeft, odl_SetLeft, VT_BOOL) 
	DISP_PROPERTY_EX_ID(clsCustomBorderStyle, "Top", 2, odl_GetTop, odl_SetTop, VT_BOOL) 
	DISP_PROPERTY_EX_ID(clsCustomBorderStyle, "Right", 3, odl_GetRight, odl_SetRight, VT_BOOL) 
	DISP_PROPERTY_EX_ID(clsCustomBorderStyle, "Bottom", 4, odl_GetBottom, odl_SetBottom, VT_BOOL) 
	DISP_FUNCTION_ID(clsCustomBorderStyle, "GetXML", 5, odl_GetXML, VT_BSTR, VTS_NONE) 
	DISP_FUNCTION_ID(clsCustomBorderStyle, "SetXML", 6, odl_SetXML, VT_EMPTY, VTS_BSTR) 
END_DISPATCH_MAP()

BEGIN_INTERFACE_MAP(clsCustomBorderStyle, CCmdTarget)
	INTERFACE_PART(clsCustomBorderStyle, IID_IclsCustomBorderStyle, Dispatch)
END_INTERFACE_MAP()

BEGIN_MESSAGE_MAP(clsCustomBorderStyle, CCmdTarget)
END_MESSAGE_MAP()

clsCustomBorderStyle::clsCustomBorderStyle()
{
	AfxMessageBox(_T("Invalid constructor has been invoked for clsCustomBorderStyle. Please inform support@sourcecodestore.com of this error"), MB_OK | MB_ICONSTOP | MB_APPLMODAL);
}

clsCustomBorderStyle::clsCustomBorderStyle(CActiveGanttVCCtl* oControl)
{
	EnableAutomation();
	AfxOleLockApp();
	mp_oControl = oControl;
	Clear();
}

clsCustomBorderStyle::~clsCustomBorderStyle()
{
	AfxOleUnlockApp();
}

void clsCustomBorderStyle::OnFinalRelease()
{
	CCmdTarget::OnFinalRelease();
}

BOOL clsCustomBorderStyle::GetLeft(void)
{
	return mp_bLeft;
}

void clsCustomBorderStyle::SetLeft(BOOL Value)
{
	mp_bLeft = Value;
}

BOOL clsCustomBorderStyle::GetTop(void)
{
	return mp_bTop;
}

void clsCustomBorderStyle::SetTop(BOOL Value)
{
	mp_bTop = Value;
}

BOOL clsCustomBorderStyle::GetRight(void)
{
	return mp_bRight;
}

void clsCustomBorderStyle::SetRight(BOOL Value)
{
	mp_bRight = Value;
}

BOOL clsCustomBorderStyle::GetBottom(void)
{
	return mp_bBottom;
}

void clsCustomBorderStyle::SetBottom(BOOL Value)
{
	mp_bBottom = Value;
}

CString clsCustomBorderStyle::GetXML(void)
{
	clsXML oXML(mp_oControl, "CustomBorderStyle");
	oXML.InitializeWriter();
	oXML.WriteProperty("Bottom", mp_bBottom);
	oXML.WriteProperty("Left", mp_bLeft);
	oXML.WriteProperty("Right", mp_bRight);
	oXML.WriteProperty("Top", mp_bTop);
	return oXML.GetXML();
}

void clsCustomBorderStyle::SetXML(CString sXML)
{
	clsXML oXML(mp_oControl, "CustomBorderStyle");
	oXML.SetXML(sXML);
	oXML.InitializeReader();
	oXML.ReadProperty("Bottom", mp_bBottom);
	oXML.ReadProperty("Left", mp_bLeft);
	oXML.ReadProperty("Right", mp_bRight);
	oXML.ReadProperty("Top", mp_bTop);
}

void clsCustomBorderStyle::Clear()
{
	mp_bTop = TRUE;
	mp_bBottom = TRUE;
	mp_bLeft = TRUE;
	mp_bRight = TRUE;
}

void clsCustomBorderStyle::Clone(clsCustomBorderStyle* oClone)
{
    oClone->SetBottom(mp_bBottom);
    oClone->SetLeft(mp_bLeft);
    oClone->SetRight(mp_bRight);
    oClone->SetTop(mp_bTop);
}