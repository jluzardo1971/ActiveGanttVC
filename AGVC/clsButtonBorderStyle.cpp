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
#include "clsButtonBorderStyle.h"


// clsButtonBorderStyle

IMPLEMENT_DYNCREATE(clsButtonBorderStyle, CCmdTarget)

clsButtonBorderStyle::clsButtonBorderStyle()
{
	AfxMessageBox(_T("Invalid constructor has been invoked for clsButtonBorderStyle. Please inform support@sourcecodestore.com of this error"), MB_OK | MB_ICONSTOP | MB_APPLMODAL);
}


clsButtonBorderStyle::clsButtonBorderStyle(CActiveGanttVCCtl* oControl)
{
	EnableAutomation();
	AfxOleLockApp();
	mp_oControl = oControl;
	Clear();
}

clsButtonBorderStyle::~clsButtonBorderStyle()
{

	AfxOleUnlockApp();
}


void clsButtonBorderStyle::OnFinalRelease()
{
	CCmdTarget::OnFinalRelease();
}


BEGIN_MESSAGE_MAP(clsButtonBorderStyle, CCmdTarget)
END_MESSAGE_MAP()


BEGIN_DISPATCH_MAP(clsButtonBorderStyle, CCmdTarget)
	DISP_PROPERTY_EX_ID(clsButtonBorderStyle, "RaisedExteriorLeftTopColor", 1, odl_GetRaisedExteriorLeftTopColor, odl_SetRaisedExteriorLeftTopColor, VT_COLOR)
	DISP_PROPERTY_EX_ID(clsButtonBorderStyle, "RaisedInteriorLeftTopColor", 2, odl_GetRaisedInteriorLeftTopColor, odl_SetRaisedInteriorLeftTopColor, VT_COLOR)
	DISP_PROPERTY_EX_ID(clsButtonBorderStyle, "RaisedExteriorRightBottomColor", 3, odl_GetRaisedExteriorRightBottomColor, odl_SetRaisedExteriorRightBottomColor, VT_COLOR)
	DISP_PROPERTY_EX_ID(clsButtonBorderStyle, "RaisedInteriorRightBottomColor", 4, odl_GetRaisedInteriorRightBottomColor, odl_SetRaisedInteriorRightBottomColor, VT_COLOR)
	DISP_PROPERTY_EX_ID(clsButtonBorderStyle, "SunkenExteriorLeftTopColor", 5, odl_GetSunkenExteriorLeftTopColor, odl_SetSunkenExteriorLeftTopColor, VT_COLOR)
	DISP_PROPERTY_EX_ID(clsButtonBorderStyle, "SunkenInteriorLeftTopColor", 6, odl_GetSunkenInteriorLeftTopColor, odl_SetSunkenInteriorLeftTopColor, VT_COLOR)
	DISP_PROPERTY_EX_ID(clsButtonBorderStyle, "SunkenExteriorRightBottomColor", 7, odl_GetSunkenExteriorRightBottomColor, odl_SetSunkenExteriorRightBottomColor, VT_COLOR)
	DISP_PROPERTY_EX_ID(clsButtonBorderStyle, "SunkenInteriorRightBottomColor", 8, odl_GetSunkenInteriorRightBottomColor, odl_SetSunkenInteriorRightBottomColor, VT_COLOR)
	DISP_FUNCTION_ID(clsButtonBorderStyle, "GetXML", 9, odl_GetXML, VT_BSTR, VTS_NONE) 
	DISP_FUNCTION_ID(clsButtonBorderStyle, "SetXML", 10, odl_SetXML, VT_EMPTY, VTS_BSTR)
END_DISPATCH_MAP()

//{B62D37E0-B9ED-4B6D-9FF9-802205853D85}
static const IID IID_IclsButtonBorderStyle = { 0xB62D37E0, 0xB9ED, 0x4B6D, { 0x9F, 0xF9, 0x80, 0x22, 0x05, 0x85, 0x3D, 0x85} };


BEGIN_INTERFACE_MAP(clsButtonBorderStyle, CCmdTarget)
	INTERFACE_PART(clsButtonBorderStyle, IID_IclsButtonBorderStyle, Dispatch)
END_INTERFACE_MAP()

//{2A19248E-5EE7-40F2-A555-EDDB84449793}
IMPLEMENT_OLECREATE_FLAGS(clsButtonBorderStyle, "AGVC.clsButtonBorderStyle", afxRegApartmentThreading, 0x2A19248E, 0x5EE7, 0x40F2, 0xA5, 0x55, 0xED, 0xDB, 0x84, 0x44, 0x97, 0x93)








Gdiplus::Color clsButtonBorderStyle::GetRaisedExteriorLeftTopColor(void)
{
	return mp_clrRaisedExteriorLeftTopColor;
}

void clsButtonBorderStyle::SetRaisedExteriorLeftTopColor(Gdiplus::Color Value)
{
	mp_clrRaisedExteriorLeftTopColor = Value;
}

Gdiplus::Color clsButtonBorderStyle::GetRaisedInteriorLeftTopColor(void)
{
	return mp_clrRaisedInteriorLeftTopColor;
}

void clsButtonBorderStyle::SetRaisedInteriorLeftTopColor(Gdiplus::Color Value)
{
	mp_clrRaisedInteriorLeftTopColor = Value;
}

Gdiplus::Color clsButtonBorderStyle::GetRaisedExteriorRightBottomColor(void)
{
	return mp_clrRaisedExteriorRightBottomColor;
}

void clsButtonBorderStyle::SetRaisedExteriorRightBottomColor(Gdiplus::Color Value)
{
	mp_clrRaisedExteriorRightBottomColor = Value;
}

Gdiplus::Color clsButtonBorderStyle::GetRaisedInteriorRightBottomColor(void)
{
	return mp_clrRaisedInteriorRightBottomColor;
}

void clsButtonBorderStyle::SetRaisedInteriorRightBottomColor(Gdiplus::Color Value)
{
	mp_clrRaisedInteriorRightBottomColor = Value;
}

Gdiplus::Color clsButtonBorderStyle::GetSunkenExteriorLeftTopColor(void)
{
	return mp_clrSunkenExteriorLeftTopColor;
}

void clsButtonBorderStyle::SetSunkenExteriorLeftTopColor(Gdiplus::Color Value)
{
	mp_clrSunkenExteriorLeftTopColor = Value;
}

Gdiplus::Color clsButtonBorderStyle::GetSunkenInteriorLeftTopColor(void)
{
	return mp_clrSunkenInteriorLeftTopColor;
}

void clsButtonBorderStyle::SetSunkenInteriorLeftTopColor(Gdiplus::Color Value)
{
	mp_clrSunkenInteriorLeftTopColor = Value;
}

Gdiplus::Color clsButtonBorderStyle::GetSunkenExteriorRightBottomColor(void)
{
	return mp_clrSunkenExteriorRightBottomColor;
}

void clsButtonBorderStyle::SetSunkenExteriorRightBottomColor(Gdiplus::Color Value)
{
	mp_clrSunkenExteriorRightBottomColor = Value;
}

Gdiplus::Color clsButtonBorderStyle::GetSunkenInteriorRightBottomColor(void)
{
	return mp_clrSunkenInteriorRightBottomColor;
}

void clsButtonBorderStyle::SetSunkenInteriorRightBottomColor(Gdiplus::Color Value)
{
	mp_clrSunkenInteriorRightBottomColor = Value;
}

CString clsButtonBorderStyle::GetXML(void)
{
	clsXML oXML(mp_oControl, "ButtonBorderStyle");
	oXML.InitializeWriter();
	oXML.WriteProperty("RaisedExteriorLeftTopColor", mp_clrRaisedExteriorLeftTopColor);
	oXML.WriteProperty("RaisedInteriorLeftTopColor", mp_clrRaisedInteriorLeftTopColor);
	oXML.WriteProperty("RaisedExteriorRightBottomColor", mp_clrRaisedExteriorRightBottomColor);
	oXML.WriteProperty("RaisedInteriorRightBottomColor", mp_clrRaisedInteriorRightBottomColor);
	oXML.WriteProperty("SunkenExteriorLeftTopColor", mp_clrSunkenExteriorLeftTopColor);
	oXML.WriteProperty("SunkenInteriorLeftTopColor", mp_clrSunkenInteriorLeftTopColor);
	oXML.WriteProperty("SunkenExteriorRightBottomColor", mp_clrSunkenExteriorRightBottomColor);
	oXML.WriteProperty("SunkenInteriorRightBottomColor", mp_clrSunkenInteriorRightBottomColor);
	return oXML.GetXML();
}

void clsButtonBorderStyle::SetXML(CString sXML)
{
	clsXML oXML(mp_oControl, "ButtonBorderStyle");
	oXML.SetXML(sXML);
	oXML.InitializeReader();
	oXML.ReadProperty("RaisedExteriorLeftTopColor", mp_clrRaisedExteriorLeftTopColor);
	oXML.ReadProperty("RaisedInteriorLeftTopColor", mp_clrRaisedInteriorLeftTopColor);
	oXML.ReadProperty("RaisedExteriorRightBottomColor", mp_clrRaisedExteriorRightBottomColor);
	oXML.ReadProperty("RaisedInteriorRightBottomColor", mp_clrRaisedInteriorRightBottomColor);
	oXML.ReadProperty("SunkenExteriorLeftTopColor", mp_clrSunkenExteriorLeftTopColor);
	oXML.ReadProperty("SunkenInteriorLeftTopColor", mp_clrSunkenInteriorLeftTopColor);
	oXML.ReadProperty("SunkenExteriorRightBottomColor", mp_clrSunkenExteriorRightBottomColor);
	oXML.ReadProperty("SunkenInteriorRightBottomColor", mp_clrSunkenInteriorRightBottomColor);
}

void clsButtonBorderStyle::Clear()
{
    mp_clrRaisedExteriorLeftTopColor = Color::MakeARGB(255, 240, 240, 240);
    mp_clrRaisedInteriorLeftTopColor = Color::MakeARGB(255, 192, 192, 192);
	mp_clrRaisedExteriorRightBottomColor = Color::Gray;
    mp_clrRaisedInteriorRightBottomColor = Color::MakeARGB(255, 64, 64, 64);
	mp_clrSunkenExteriorLeftTopColor = Color::Gray;
    mp_clrSunkenInteriorLeftTopColor = Color::MakeARGB(255, 64, 64, 64);
    mp_clrSunkenExteriorRightBottomColor = Color::MakeARGB(255, 240, 240, 240);
    mp_clrSunkenInteriorRightBottomColor = Color::MakeARGB(255, 192, 192, 192);
}

void clsButtonBorderStyle::Clone(clsButtonBorderStyle* oClone)
{
    oClone->SetRaisedExteriorLeftTopColor(mp_clrRaisedExteriorLeftTopColor);
    oClone->SetRaisedInteriorLeftTopColor(mp_clrRaisedInteriorLeftTopColor);
    oClone->SetRaisedExteriorRightBottomColor(mp_clrRaisedExteriorRightBottomColor);
    oClone->SetRaisedInteriorRightBottomColor(mp_clrRaisedInteriorRightBottomColor);
    oClone->SetSunkenExteriorLeftTopColor(mp_clrSunkenExteriorLeftTopColor);
    oClone->SetSunkenInteriorLeftTopColor(mp_clrSunkenInteriorLeftTopColor);
    oClone->SetSunkenExteriorRightBottomColor(mp_clrSunkenExteriorRightBottomColor);
    oClone->SetSunkenInteriorRightBottomColor(mp_clrSunkenInteriorRightBottomColor);
}