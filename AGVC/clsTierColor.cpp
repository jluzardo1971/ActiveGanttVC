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
#include "clsTierColor.h"

IMPLEMENT_DYNCREATE(clsTierColor, clsItemBase)

//{6CBB6C0A-2495-42DE-87CE-A54462A39E0F}
static const IID IID_IclsTierColor = { 0x6CBB6C0A, 0x2495, 0x42DE, { 0x87, 0xCE, 0xA5, 0x44, 0x62, 0xA3, 0x9E, 0x0F} };

//{11AA8015-748C-4F1D-8372-8035C2EB4EE8}
IMPLEMENT_OLECREATE_FLAGS(clsTierColor, "AGVC.clsTierColor", afxRegApartmentThreading, 0x11AA8015, 0x748C, 0x4F1D, 0x83, 0x72, 0x80, 0x35, 0xC2, 0xEB, 0x4E, 0xE8)

BEGIN_DISPATCH_MAP(clsTierColor, clsItemBase)
	DISP_PROPERTY_EX_ID(clsTierColor, "Key", 1, odl_GetKey, odl_SetKey, VT_BSTR) 
	DISP_PROPERTY_EX_ID(clsTierColor, "ForeColor", 2, odl_GetForeColor, odl_SetForeColor, VT_COLOR) 
	DISP_PROPERTY_EX_ID(clsTierColor, "BackColor", 3, odl_GetBackColor, odl_SetBackColor, VT_COLOR) 
	DISP_FUNCTION_ID(clsTierColor, "GetXML", 4, odl_GetXML, VT_BSTR, VTS_NONE) 
	DISP_FUNCTION_ID(clsTierColor, "SetXML", 5, odl_SetXML, VT_EMPTY, VTS_BSTR)
	DISP_PROPERTY_EX_ID(clsTierColor, "Index", 6, odl_GetIndex, SetNotSupported, VT_I4)
	DISP_PROPERTY_EX_ID(clsTierColor, "StartGradientColor", 7, odl_GetStartGradientColor, odl_SetStartGradientColor, VT_COLOR) 
	DISP_PROPERTY_EX_ID(clsTierColor, "EndGradientColor", 8, odl_GetEndGradientColor, odl_SetEndGradientColor, VT_COLOR) 
	DISP_PROPERTY_EX_ID(clsTierColor, "HatchBackColor", 9, odl_GetHatchBackColor, odl_SetHatchBackColor, VT_COLOR) 
	DISP_PROPERTY_EX_ID(clsTierColor, "HatchForeColor", 10, odl_GetHatchForeColor, odl_SetHatchForeColor, VT_COLOR) 
END_DISPATCH_MAP()

BEGIN_INTERFACE_MAP(clsTierColor, clsItemBase)
	INTERFACE_PART(clsTierColor, IID_IclsTierColor, Dispatch)
END_INTERFACE_MAP()

BEGIN_MESSAGE_MAP(clsTierColor, clsItemBase)
END_MESSAGE_MAP()

clsTierColor::clsTierColor(CActiveGanttVCCtl* oControl, clsTierColors* oTierColors)
{
	EnableAutomation();
	AfxOleLockApp();
	mp_oControl = oControl;
	mp_clsTierColors = oTierColors;
}

clsTierColor::clsTierColor()
{
	AfxMessageBox(_T("Invalid constructor has been invoked for clsTierColor. Please inform support@sourcecodestore.com of this error"), MB_OK | MB_ICONSTOP | MB_APPLMODAL);
}

clsTierColor::~clsTierColor()
{
	AfxOleUnlockApp();
}

void clsTierColor::OnFinalRelease()
{
	clsItemBase::OnFinalRelease();
}



CString clsTierColor::GetKey(void)
{
	return mp_sKey;
}

void clsTierColor::SetKey(CString Value)
{
	mp_clsTierColors->mp_oCollection->mp_SetKey(&mp_sKey, Value, TIERCOLORS_SET_KEY);
}

Gdiplus::Color clsTierColor::GetForeColor(void)
{
	return mp_clrForeColor;
}

void clsTierColor::SetForeColor(Gdiplus::Color Value)
{
	mp_clrForeColor = Value;
}

Gdiplus::Color clsTierColor::GetBackColor(void)
{
	return mp_clrBackColor;
}

void clsTierColor::SetBackColor(Gdiplus::Color Value)
{
	mp_clrBackColor = Value;
}

Gdiplus::Color clsTierColor::GetStartGradientColor(void)
{
	return mp_clrStartGradientColor;
}

void clsTierColor::SetStartGradientColor(Gdiplus::Color Value)
{
	mp_clrStartGradientColor = Value;
}

Gdiplus::Color clsTierColor::GetEndGradientColor(void)
{
	return mp_clrEndGradientColor;
}

void clsTierColor::SetEndGradientColor(Gdiplus::Color Value)
{
	mp_clrEndGradientColor = Value;
}

Gdiplus::Color clsTierColor::GetHatchBackColor(void)
{
	return mp_clrHatchBackColor;
}

void clsTierColor::SetHatchBackColor(Gdiplus::Color Value)
{
	mp_clrHatchBackColor = Value;
}

Gdiplus::Color clsTierColor::GetHatchForeColor(void)
{
	return mp_clrHatchForeColor;
}

void clsTierColor::SetHatchForeColor(Gdiplus::Color Value)
{
	mp_clrHatchForeColor = Value;
}

LONG clsTierColor::GetIndex(void)
{
	return mp_lIndex;
}

void clsTierColor::SetIndex(LONG Value)
{
	mp_lIndex = Value;
}

void clsTierColor::Finalize(void)
{
}

CString clsTierColor::GetXML(void)
{
	clsXML oXML(mp_oControl, "TierColor");
	oXML.InitializeWriter();
	oXML.WriteProperty("BackColor", mp_clrBackColor);
	oXML.WriteProperty("ForeColor", mp_clrForeColor);
	oXML.WriteProperty("StartGradientColor", mp_clrStartGradientColor);
	oXML.WriteProperty("EndGradientColor", mp_clrEndGradientColor);
	oXML.WriteProperty("HatchBackColor", mp_clrHatchBackColor);
	oXML.WriteProperty("HatchForeColor", mp_clrHatchForeColor);
	oXML.WriteProperty("Key", mp_sKey);
	return oXML.GetXML();
}

void clsTierColor::SetXML(CString sXML)
{
	clsXML oXML(mp_oControl, "TierColor");
	oXML.SetXML(sXML);
	oXML.InitializeReader();
	oXML.ReadProperty("BackColor", mp_clrBackColor);
	oXML.ReadProperty("ForeColor", mp_clrForeColor);
	oXML.ReadProperty("StartGradientColor", mp_clrStartGradientColor);
	oXML.ReadProperty("EndGradientColor", mp_clrEndGradientColor);
	oXML.ReadProperty("HatchBackColor", mp_clrHatchBackColor);
	oXML.ReadProperty("HatchForeColor", mp_clrHatchForeColor);
	oXML.ReadProperty("Key", mp_sKey);
}

void clsTierColor::Clone(clsTierColor* oClone)
{
    oClone->SetBackColor(mp_clrBackColor);
    oClone->SetForeColor(mp_clrForeColor);
    oClone->SetStartGradientColor(mp_clrStartGradientColor);
    oClone->SetEndGradientColor(mp_clrEndGradientColor);
    oClone->SetHatchBackColor(mp_clrHatchBackColor);
    oClone->SetHatchForeColor(mp_clrHatchForeColor);
}