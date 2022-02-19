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
#include "clsTickMark.h"

IMPLEMENT_DYNCREATE(clsTickMark, clsItemBase)

//{D8894736-C3AF-49E0-8B4B-1B7BD946E5B3}
static const IID IID_IclsTickMark = { 0xD8894736, 0xC3AF, 0x49E0, { 0x8B, 0x4B, 0x1B, 0x7B, 0xD9, 0x46, 0xE5, 0xB3} };

//{C3ECAABF-A4AA-40A4-9C88-1DE80C827B61}
IMPLEMENT_OLECREATE_FLAGS(clsTickMark, "AGVC.clsTickMark", afxRegApartmentThreading, 0xC3ECAABF, 0xA4AA, 0x40A4, 0x9C, 0x88, 0x1D, 0xE8, 0x0C, 0x82, 0x7B, 0x61)

BEGIN_DISPATCH_MAP(clsTickMark, clsItemBase)
	DISP_PROPERTY_EX_ID(clsTickMark, "Key", 1, odl_GetKey, odl_SetKey, VT_BSTR) 
	DISP_PROPERTY_EX_ID(clsTickMark, "DisplayText", 2, odl_GetDisplayText, odl_SetDisplayText, VT_BOOL) 
	DISP_PROPERTY_EX_ID(clsTickMark, "TextFormat", 3, odl_GetTextFormat, odl_SetTextFormat, VT_BSTR) 
	DISP_PROPERTY_EX_ID(clsTickMark, "Tag", 4, odl_GetTag, odl_SetTag, VT_BSTR) 
	DISP_PROPERTY_EX_ID(clsTickMark, "Interval", 5, odl_GetInterval, odl_SetInterval, VT_I4) 
	DISP_PROPERTY_EX_ID(clsTickMark, "TickMarkType", 6, odl_GetTickMarkType, odl_SetTickMarkType, VT_I4) 
	DISP_FUNCTION_ID(clsTickMark, "GetXML", 7, odl_GetXML, VT_BSTR, VTS_NONE) 
	DISP_FUNCTION_ID(clsTickMark, "SetXML", 8, odl_SetXML, VT_EMPTY, VTS_BSTR)
	DISP_PROPERTY_EX_ID(clsTickMark, "Index", 9, odl_GetIndex, SetNotSupported, VT_I4)
	DISP_PROPERTY_EX_ID(clsTickMark, "Factor", 10, odl_GetFactor, odl_SetFactor, VT_I4) 
END_DISPATCH_MAP()

BEGIN_INTERFACE_MAP(clsTickMark, clsItemBase)
	INTERFACE_PART(clsTickMark, IID_IclsTickMark, Dispatch)
END_INTERFACE_MAP()

BEGIN_MESSAGE_MAP(clsTickMark, clsItemBase)
END_MESSAGE_MAP()

clsTickMark::clsTickMark(CActiveGanttVCCtl* oControl, clsTickMarks* oTickMarks)
{
	EnableAutomation();
	AfxOleLockApp();
	mp_oControl = oControl;
	mp_clsTickMarks = oTickMarks;
}

clsTickMark::clsTickMark()
{
	AfxMessageBox(_T("Invalid constructor has been invoked for clsTickMark. Please inform support@sourcecodestore.com of this error"), MB_OK | MB_ICONSTOP | MB_APPLMODAL);
}

clsTickMark::~clsTickMark()
{
	AfxOleUnlockApp();
}

void clsTickMark::OnFinalRelease()
{
	clsItemBase::OnFinalRelease();
}

CString clsTickMark::GetKey(void)
{
	return mp_sKey;
}

void clsTickMark::SetKey(CString Value)
{
	mp_clsTickMarks->mp_oCollection->mp_SetKey(&mp_sKey, Value, TICKMARKS_SET_KEY);
}

BOOL clsTickMark::GetDisplayText(void)
{
	return mp_bDisplayText;
}

void clsTickMark::SetDisplayText(BOOL Value)
{
	mp_bDisplayText = Value;
}

CString clsTickMark::GetTextFormat(void)
{
	return mp_sTextFormat;
}

void clsTickMark::SetTextFormat(CString Value)
{
	mp_sTextFormat = Value;
}

CString clsTickMark::GetTag(void)
{
	return mp_sTag;
}

void clsTickMark::SetTag(CString Value)
{
	mp_sTag = Value;
}

LONG clsTickMark::GetInterval(void)
{
	return mp_yInterval;
}

void clsTickMark::SetInterval(LONG Value)
{
	mp_yInterval = Value;
}

LONG clsTickMark::GetFactor(void)
{
	return mp_lFactor;
}

void clsTickMark::SetFactor(LONG Value)
{
	mp_lFactor = Value;
}

LONG clsTickMark::GetTickMarkType(void)
{
	return mp_yTickMarkType;
}

void clsTickMark::SetTickMarkType(LONG Value)
{
	mp_yTickMarkType = Value;
}

LONG clsTickMark::GetIndex(void)
{
	return mp_lIndex;
}

void clsTickMark::SetIndex(LONG Value)
{
	mp_lIndex = Value;
}

CString clsTickMark::GetXML(void)
{
	clsXML oXML(mp_oControl, "TickMark");
	oXML.InitializeWriter();
	oXML.WriteProperty("DisplayText", mp_bDisplayText);
	oXML.WriteProperty("Interval", mp_yInterval);
	oXML.WriteProperty("Factor", mp_lFactor);
	oXML.WriteProperty("Key", mp_sKey);
	oXML.WriteProperty("Tag", mp_sTag);
	oXML.WriteProperty("TextFormat", mp_sTextFormat);
	oXML.WriteProperty("TickMarkType", mp_yTickMarkType);
	return oXML.GetXML();
}

void clsTickMark::SetXML(CString sXML)
{
	clsXML oXML(mp_oControl, "TickMark");
	oXML.SetXML(sXML);
	oXML.InitializeReader();
	oXML.ReadProperty("DisplayText", mp_bDisplayText);
	oXML.ReadProperty("Interval", mp_yInterval);
	oXML.ReadProperty("Factor", mp_lFactor);
	oXML.ReadProperty("Key", mp_sKey);
	oXML.ReadProperty("Tag", mp_sTag);
	oXML.ReadProperty("TextFormat", mp_sTextFormat);
	oXML.ReadProperty("TickMarkType", mp_yTickMarkType);
}

void clsTickMark::Clear()
{
	mp_bDisplayText = FALSE;
	mp_sTextFormat = "";
	mp_sTag = "";
	mp_yInterval = IL_SECOND;
	mp_lFactor = 1;
	mp_yTickMarkType = TLT_SMALL;
}

void clsTickMark::Clone(clsTickMark* oClone)
{
    oClone->SetDisplayText(mp_bDisplayText);
    oClone->SetInterval(mp_yInterval);
    oClone->SetFactor(mp_lFactor);
    oClone->SetTag(mp_sTag);
    oClone->SetTextFormat(mp_sTextFormat);
    oClone->SetTickMarkType(mp_yTickMarkType);
}