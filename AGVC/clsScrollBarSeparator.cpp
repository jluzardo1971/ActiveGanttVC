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
#include "clsScrollBarSeparator.h"




// clsScrollBarSeparator

IMPLEMENT_DYNCREATE(clsScrollBarSeparator, CCmdTarget)


clsScrollBarSeparator::~clsScrollBarSeparator()
{
	AfxOleUnlockApp();
}


void clsScrollBarSeparator::OnFinalRelease()
{
	CCmdTarget::OnFinalRelease();
}


BEGIN_MESSAGE_MAP(clsScrollBarSeparator, CCmdTarget)
END_MESSAGE_MAP()


BEGIN_DISPATCH_MAP(clsScrollBarSeparator, CCmdTarget)
	DISP_PROPERTY_EX_ID(clsScrollBarSeparator, "StyleIndex", 1, odl_GetStyleIndex, odl_SetStyleIndex, VT_BSTR) 
	DISP_PROPERTY_EX_ID(clsScrollBarSeparator, "Style", 2, odl_GetStyle, SetNotSupported, VT_DISPATCH)
	DISP_FUNCTION_ID(clsScrollBarSeparator, "GetXML", 3, odl_GetXML, VT_BSTR, VTS_NONE) 
	DISP_FUNCTION_ID(clsScrollBarSeparator, "SetXML", 4, odl_SetXML, VT_EMPTY, VTS_BSTR)
END_DISPATCH_MAP()

//{BE2D822E-8865-49C8-B605-F2F6903F47AB}
static const IID IID_IclsScrollBarSeparator = { 0xBE2D822E, 0x8865, 0x49C8, { 0xB6, 0x05, 0xF2, 0xF6, 0x90, 0x3F, 0x47, 0xAB} };



BEGIN_INTERFACE_MAP(clsScrollBarSeparator, CCmdTarget)
	INTERFACE_PART(clsScrollBarSeparator, IID_IclsScrollBarSeparator, Dispatch)
END_INTERFACE_MAP()

//{FDDCD8B2-195D-40D2-A844-F9EDDAD6C58E}
IMPLEMENT_OLECREATE_FLAGS(clsScrollBarSeparator, "AGVC.clsScrollBarSeparator", afxRegApartmentThreading, 0xFDDCD8B2, 0x195D, 0x40D2, 0xA8, 0x44, 0xF9, 0xED, 0xDA, 0xD6, 0xC5, 0x8E)


// clsScrollBarSeparator message handlers

clsScrollBarSeparator::clsScrollBarSeparator()
{
	AfxMessageBox(_T("Invalid constructor has been invoked for clsScrollBarSeparator. Please inform support@sourcecodestore.com of this error"), MB_OK | MB_ICONSTOP | MB_APPLMODAL);
}

clsScrollBarSeparator::clsScrollBarSeparator(CActiveGanttVCCtl* oControl)
{
	EnableAutomation();
	AfxOleLockApp();
	mp_oControl = oControl;
	mp_sStyleIndex = "DS_SB_SEPARATOR";
	mp_oStyle = mp_oControl->GetStyles()->FItem("DS_SB_SEPARATOR");
}

CString clsScrollBarSeparator::GetStyleIndex(void)
{
	if (mp_sStyleIndex == "DS_SB_SEPARATOR") 
	{
		return "";
	}
	else 
	{
		return mp_sStyleIndex;
	}
}

void clsScrollBarSeparator::SetStyleIndex(CString Value)
{
	Value = Value.Trim();
	if (Value.GetLength() == 0) {Value = "DS_SB_SEPARATOR";} 
	mp_sStyleIndex = Value;
	mp_oStyle = mp_oControl->GetStyles()->FItem(Value);
}

clsStyle* clsScrollBarSeparator::GetStyle(void)
{
	return mp_oStyle;
}

CString clsScrollBarSeparator::GetXML(void)
{
	clsXML oXML(mp_oControl, "ScrollBarSeparator");
	oXML.InitializeWriter();
	oXML.WriteProperty("StyleIndex", mp_sStyleIndex);
	return oXML.GetXML();
}

void clsScrollBarSeparator::SetXML(CString sXML)
{
	clsXML oXML(mp_oControl, "ScrollBarSeparator");
	oXML.SetXML(sXML);
	oXML.InitializeReader();
	oXML.ReadProperty("StyleIndex", mp_sStyleIndex);
	SetStyleIndex(mp_sStyleIndex);
}