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
#include "clsXML.h"
#include "WBSMasks.h"

IMPLEMENT_DYNCREATE(WBSMasks, CCmdTarget)

//{2FDCAA1D-13AB-4684-920E-7D92D5334A9D}
static const IID IID_IWBSMasks = { 0x2FDCAA1D, 0x13AB, 0x4684, { 0x92, 0x0E, 0x7D, 0x92, 0xD5, 0x33, 0x4A, 0x9D} };

//{2BD579A8-7FE3-4D26-879F-A762E6531B20}
IMPLEMENT_OLECREATE_FLAGS(WBSMasks, "MSP2003.WBSMasks", afxRegApartmentThreading, 0x2BD579A8, 0x7FE3, 0x4D26, 0x87, 0x9F, 0xA7, 0x62, 0xE6, 0x53, 0x1B, 0x20)

BEGIN_DISPATCH_MAP(WBSMasks, CCmdTarget)
DISP_PROPERTY_EX_ID(WBSMasks, "Count", 1, odl_GetCount, SetNotSupported, VT_I4)
DISP_PROPERTY_PARAM_ID(WBSMasks, "Item", 2, odl_Item, SetNotSupported, VT_DISPATCH, VTS_BSTR)
DISP_FUNCTION_ID(WBSMasks, "Add", 3, odl_Add, VT_DISPATCH, VTS_NONE)
DISP_FUNCTION_ID(WBSMasks, "Clear", 4, odl_Clear, VT_EMPTY, VTS_NONE)
DISP_FUNCTION_ID(WBSMasks, "Remove", 5, odl_Remove, VT_EMPTY, VTS_BSTR)
DISP_FUNCTION_ID(WBSMasks, "IsNull", 6, IsNull, VT_BOOL, VTS_NONE)
DISP_FUNCTION_ID(WBSMasks, "Initialize", 7, Initialize, VT_EMPTY, VTS_NONE)
DISP_FUNCTION_ID(WBSMasks, "SetXML", 8, odl_SetXML, VT_EMPTY, VTS_BSTR)
DISP_FUNCTION_ID(WBSMasks, "GetXML", 9, odl_GetXML, VT_BSTR, VTS_NONE)
DISP_DEFVALUE(WBSMasks, "Item")
END_DISPATCH_MAP()

BEGIN_INTERFACE_MAP(WBSMasks, CCmdTarget)
	INTERFACE_PART(WBSMasks, IID_IWBSMasks, Dispatch)
END_INTERFACE_MAP()

BEGIN_MESSAGE_MAP(WBSMasks, CCmdTarget)
END_MESSAGE_MAP()


WBSMasks::WBSMasks()
{
	EnableAutomation();
	AfxOleLockApp();
	InitVars();
}

void WBSMasks::Initialize(void)
{
	InitVars();
}

void WBSMasks::InitVars(void)
{
	mp_oCollection = new clsCollectionBase("WBSMask");
}

WBSMasks::~WBSMasks()
{
	delete mp_oCollection;
	AfxOleUnlockApp();
}

void WBSMasks::OnFinalRelease()
{
	CCmdTarget::OnFinalRelease();
}

LONG WBSMasks::odl_GetCount(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetCount();
}

LONG WBSMasks::GetCount(void)
{
	return mp_oCollection->m_lCount();
}

IDispatch* WBSMasks::odl_Item(LPCTSTR Index)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return Item(Index)->GetIDispatch(TRUE);
}

IDispatch* WBSMasks::odl_Add(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return Add()->GetIDispatch(TRUE);
}

void WBSMasks::odl_Clear(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	Clear();
}

void WBSMasks::odl_Remove(LPCTSTR Index)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	Remove(Index);
}

BOOL WBSMasks::IsNull(void)
{
	BOOL bReturn = TRUE;
	if (GetCount() > 0)
	{
		bReturn = FALSE;
	}
	return bReturn;
}

WBSMask* WBSMasks::Item(CString Index)
{
	WBSMask *oWBSMask;
	oWBSMask = (WBSMask*)mp_oCollection->m_oItem(Index, MP_ITEM_1, MP_ITEM_2, MP_ITEM_3, MP_ITEM_4);
	return oWBSMask;
}

WBSMask* WBSMasks::Add()
{
	mp_oCollection->SetAddMode(TRUE);
	WBSMask* oWBSMask = new WBSMask();
	oWBSMask->mp_oCollection = mp_oCollection;
	mp_oCollection->m_Add(oWBSMask, _T(""), MP_ADD_1, MP_ADD_2, FALSE, MP_ADD_3);
	return oWBSMask;
}

void WBSMasks::Clear(void)
{
	mp_oCollection->m_Clear();
}

void WBSMasks::Remove(CString Index)
{
	mp_oCollection->m_Remove(Index, MP_REMOVE_1, MP_REMOVE_2, MP_REMOVE_3, MP_REMOVE_4);
}

BSTR WBSMasks::odl_GetXML(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetXML().AllocSysString();
}

void WBSMasks::odl_SetXML(LPCTSTR sXML)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetXML(sXML);
}

CString WBSMasks::GetXML(void)
{
	if (IsNull() == TRUE)
	{
		return "<WBSMasks/>";
	}
	LONG lIndex;
	WBSMask* oWBSMask;
	clsXML oXML("WBSMasks");
	oXML.InitializeWriter();
	for (lIndex = 1; lIndex <= GetCount(); lIndex++)
	{
		oWBSMask = (WBSMask*) mp_oCollection->m_oReturnArrayElement(lIndex);
		oXML.WriteObject(oWBSMask->GetXML());
	}
	return oXML.GetXML();
}

void WBSMasks::SetXML(CString sXML)
{
	LONG lIndex;
	clsXML oXML("WBSMasks");
	oXML.SetSupportOptional(TRUE);
	oXML.SetXML(sXML);
	oXML.InitializeReader();
	mp_oCollection->m_Clear();
	for (lIndex = 1;lIndex <= oXML.ReadCollectionCount();lIndex++)
	{
		WBSMask* oWBSMask = new WBSMask();
		oWBSMask->SetXML(oXML.ReadCollectionObject(lIndex));
		mp_oCollection->SetAddMode(TRUE);
		CString sKey = _T("");
		oWBSMask->mp_oCollection = mp_oCollection;
		mp_oCollection->m_Add(oWBSMask, sKey, MP_ADD_1, MP_ADD_2, FALSE, MP_ADD_3);
	}
}