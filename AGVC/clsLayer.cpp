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
#include "clsLayer.h"



IMPLEMENT_DYNCREATE(clsLayer, clsItemBase)

//{3003C1B0-4FD8-49EC-B77E-AFFF7853C74C}
static const IID IID_IclsLayer = { 0x3003C1B0, 0x4FD8, 0x49EC, { 0xB7, 0x7E, 0xAF, 0xFF, 0x78, 0x53, 0xC7, 0x4C} };

//{6322FA01-EF11-4A92-9CC6-2C2C72C4FBC6}
IMPLEMENT_OLECREATE_FLAGS(clsLayer, "AGVC.clsLayer", afxRegApartmentThreading, 0x6322FA01, 0xEF11, 0x4A92, 0x9C, 0xC6, 0x2C, 0x2C, 0x72, 0xC4, 0xFB, 0xC6)

BEGIN_DISPATCH_MAP(clsLayer, clsItemBase)
	DISP_PROPERTY_EX_ID(clsLayer, "Key", 1, odl_GetKey, odl_SetKey, VT_BSTR) 
	DISP_PROPERTY_EX_ID(clsLayer, "Visible", 2, odl_GetVisible, odl_SetVisible, VT_BOOL) 
	DISP_PROPERTY_EX_ID(clsLayer, "Tag", 3, odl_GetTag, odl_SetTag, VT_BSTR) 
	DISP_FUNCTION_ID(clsLayer, "GetXML", 4, odl_GetXML, VT_BSTR, VTS_NONE) 
	DISP_FUNCTION_ID(clsLayer, "SetXML", 5, odl_SetXML, VT_EMPTY, VTS_BSTR)
	DISP_PROPERTY_EX_ID(clsLayer, "Index", 6, odl_GetIndex, SetNotSupported, VT_I4)
END_DISPATCH_MAP()

BEGIN_INTERFACE_MAP(clsLayer, clsItemBase)
	INTERFACE_PART(clsLayer, IID_IclsLayer, Dispatch)
END_INTERFACE_MAP()

BEGIN_MESSAGE_MAP(clsLayer, clsItemBase)
END_MESSAGE_MAP()

clsLayer::clsLayer()
{
	AfxMessageBox(_T("Invalid constructor has been invoked for clsLayer. Please inform support@sourcecodestore.com of this error"), MB_OK | MB_ICONSTOP | MB_APPLMODAL);
}

clsLayer::clsLayer(CActiveGanttVCCtl* oControl)
{
	EnableAutomation();
	AfxOleLockApp();
	mp_oControl = oControl;
	mp_bVisible = TRUE;
	mp_sTag = "";
}

clsLayer::~clsLayer()
{
	AfxOleUnlockApp();
}

void clsLayer::OnFinalRelease()
{
	clsItemBase::OnFinalRelease();
}

CString clsLayer::GetKey(void)
{
	return mp_sKey;
}

void clsLayer::SetKey(CString Value)
{
	mp_oControl->GetLayers()->mp_oCollection->mp_SetKey(&mp_sKey, Value, LAYERS_SET_KEY);
}

BOOL clsLayer::GetVisible(void)
{
	return mp_bVisible;
}

void clsLayer::SetVisible(BOOL Value)
{
	mp_bVisible = Value;
}

CString clsLayer::GetTag(void)
{
	return mp_sTag;
}

void clsLayer::SetTag(CString Value)
{
	mp_sTag = Value;
}

LONG clsLayer::GetIndex(void)
{
	return mp_lIndex;
}

void clsLayer::SetIndex(LONG Value)
{
	mp_lIndex = Value;
}

void clsLayer::Finalize(void)
{
}

CString clsLayer::GetXML(void)
{
	clsXML oXML(mp_oControl, "Layer");
	oXML.InitializeWriter();
	oXML.WriteProperty("Key", mp_sKey);
	oXML.WriteProperty("Visible", mp_bVisible);
	oXML.WriteProperty("Tag", mp_sTag);
	return oXML.GetXML();
}

void clsLayer::SetXML(CString sXML)
{
	clsXML oXML(mp_oControl, "Layer");
	oXML.SetXML(sXML);
	oXML.InitializeReader();
	oXML.ReadProperty("Key", mp_sKey);
	oXML.ReadProperty("Visible", mp_bVisible);
	oXML.ReadProperty("Tag", mp_sTag);
}