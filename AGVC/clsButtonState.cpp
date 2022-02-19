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
#include "clsButtonState.h"




// clsButtonState

IMPLEMENT_DYNCREATE(clsButtonState, CCmdTarget)


clsButtonState::clsButtonState()
{
	AfxMessageBox(_T("Invalid constructor has been invoked for clsButtonState. Please inform support@sourcecodestore.com of this error"), MB_OK | MB_ICONSTOP | MB_APPLMODAL);
}

clsButtonState::clsButtonState(CActiveGanttVCCtl* oControl, CString sType)
{
	EnableAutomation();
	AfxOleLockApp();
	mp_oControl = oControl;
	mp_sType = sType;
	Clear();
}

clsButtonState::~clsButtonState()
{
	AfxOleUnlockApp();
}


void clsButtonState::OnFinalRelease()
{
	CCmdTarget::OnFinalRelease();
}


BEGIN_MESSAGE_MAP(clsButtonState, CCmdTarget)
END_MESSAGE_MAP()

BEGIN_DISPATCH_MAP(clsButtonState, CCmdTarget)
	DISP_PROPERTY_EX_ID(clsButtonState, "NormalStyleIndex", 1, odl_GetNormalStyleIndex, odl_SetNormalStyleIndex, VT_BSTR) 
	DISP_PROPERTY_EX_ID(clsButtonState, "NormalStyle", 2, odl_GetNormalStyle, SetNotSupported, VT_DISPATCH)
	DISP_PROPERTY_EX_ID(clsButtonState, "PressedStyleIndex", 3, odl_GetPressedStyleIndex, odl_SetPressedStyleIndex, VT_BSTR) 
	DISP_PROPERTY_EX_ID(clsButtonState, "PressedStyle", 4, odl_GetPressedStyle, SetNotSupported, VT_DISPATCH) 
	DISP_PROPERTY_EX_ID(clsButtonState, "DisabledStyleIndex", 5, odl_GetDisabledStyleIndex, odl_SetDisabledStyleIndex, VT_BSTR) 
	DISP_PROPERTY_EX_ID(clsButtonState, "DisabledStyle", 6, odl_GetDisabledStyle, SetNotSupported, VT_DISPATCH)
	DISP_FUNCTION_ID(clsButtonState, "GetXML", 7, odl_GetXML, VT_BSTR, VTS_NONE) 
	DISP_FUNCTION_ID(clsButtonState, "SetXML", 8, odl_SetXML, VT_EMPTY, VTS_BSTR)
END_DISPATCH_MAP()

//{1718BDB6-A1CA-4893-87CC-6A512FD8F76D}
static const IID IID_IclsButtonState = { 0x1718BDB6, 0xA1CA, 0x4893, { 0x87, 0xCC, 0x6A, 0x51, 0x2F, 0xD8, 0xF7, 0x6D} };
                                 

BEGIN_INTERFACE_MAP(clsButtonState, CCmdTarget)
	INTERFACE_PART(clsButtonState, IID_IclsButtonState, Dispatch)
END_INTERFACE_MAP()

//{B9D44F50-705A-45A8-B00B-C4FF8BD2B9FB}
IMPLEMENT_OLECREATE_FLAGS(clsButtonState, "AGVC.clsButtonState", afxRegApartmentThreading, 0xB9D44F50, 0x705A, 0x45A8, 0xB0, 0x0B, 0xC4, 0xFF, 0x8B, 0xD2, 0xB9, 0xFB)



// clsButtonState message handlers



CString clsButtonState::GetNormalStyleIndex(void)
{
	if (mp_sNormalStyleIndex == "DS_SB_NORMAL") 
	{
		return "";
	}
	else 
	{
		return mp_sNormalStyleIndex;
	}
}

void clsButtonState::SetNormalStyleIndex(CString Value)
{
	Value = Value.Trim();
	if (Value.GetLength() == 0) {Value = "DS_SB_NORMAL";} 
	mp_sNormalStyleIndex = Value;
	mp_oNormalStyle = mp_oControl->GetStyles()->FItem(Value);
}

clsStyle* clsButtonState::GetNormalStyle(void)
{
	return mp_oNormalStyle;
}

CString clsButtonState::GetPressedStyleIndex(void)
{
	if (mp_sPressedStyleIndex == "DS_SB_PRESSED") 
	{
		return "";
	}
	else 
	{
		return mp_sPressedStyleIndex;
	}
}

void clsButtonState::SetPressedStyleIndex(CString Value)
{
	Value = Value.Trim();
	if (Value.GetLength() == 0) {Value = "DS_SB_PRESSED";} 
	mp_sPressedStyleIndex = Value;
	mp_oPressedStyle = mp_oControl->GetStyles()->FItem(Value);
}

clsStyle* clsButtonState::GetPressedStyle(void)
{
	return mp_oPressedStyle;
}

CString clsButtonState::GetDisabledStyleIndex(void)
{
	if (mp_sDisabledStyleIndex == "DS_SB_DISABLED") 
	{
		return "";
	}
	else 
	{
		return mp_sDisabledStyleIndex;
	}
}

void clsButtonState::SetDisabledStyleIndex(CString Value)
{
	Value = Value.Trim();
	if (Value.GetLength() == 0) {Value = "DS_SB_DISABLED";} 
	mp_sDisabledStyleIndex = Value;
	mp_oDisabledStyle = mp_oControl->GetStyles()->FItem(Value);
}

clsStyle* clsButtonState::GetDisabledStyle(void)
{
	return mp_oDisabledStyle;
}

CString clsButtonState::GetXML(void)
{
	clsXML oXML(mp_oControl, mp_sType + "ButtonState");
	oXML.InitializeWriter();
	oXML.WriteProperty("NormalStyleIndex", mp_sNormalStyleIndex);
	oXML.WriteProperty("PressedStyleIndex", mp_sPressedStyleIndex);
	oXML.WriteProperty("DisabledStyleIndex", mp_sDisabledStyleIndex);
	return oXML.GetXML();
}

void clsButtonState::SetXML(CString sXML)
{
	clsXML oXML(mp_oControl, mp_sType + "ButtonState");
	oXML.SetXML(sXML);
	oXML.InitializeReader();
	oXML.ReadProperty("NormalStyleIndex", mp_sNormalStyleIndex);
	SetNormalStyleIndex(mp_sNormalStyleIndex);
	oXML.ReadProperty("PressedStyleIndex", mp_sPressedStyleIndex);
	SetPressedStyleIndex(mp_sPressedStyleIndex);
	oXML.ReadProperty("DisabledStyleIndex", mp_sDisabledStyleIndex);
	SetDisabledStyleIndex(mp_sDisabledStyleIndex);
}

void clsButtonState::Clear()
{

	mp_sNormalStyleIndex = "DS_SB_NORMAL";
	mp_oNormalStyle = mp_oControl->GetStyles()->FItem("DS_SB_NORMAL");
	mp_sPressedStyleIndex = "DS_SB_PRESSED";
	mp_oPressedStyle = mp_oControl->GetStyles()->FItem("DS_SB_PRESSED");
	mp_sDisabledStyleIndex = "DS_SB_DISABLED";
	mp_oDisabledStyle = mp_oControl->GetStyles()->FItem("DS_SB_DISABLED");
}

void clsButtonState::Clone(clsButtonState* oClone)
{
    oClone->SetNormalStyleIndex(mp_sNormalStyleIndex);
    oClone->SetPressedStyleIndex(mp_sPressedStyleIndex);
    oClone->SetDisabledStyleIndex(mp_sDisabledStyleIndex);
	TRACE("Cloning mp_oControl: %x \r\n", oClone->mp_oControl);
}