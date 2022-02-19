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
#include "Resources.h"

IMPLEMENT_DYNCREATE(Resources, CCmdTarget)

//{24264016-4EDC-42CB-9116-984F2F4B082A}
static const IID IID_IResources = { 0x24264016, 0x4EDC, 0x42CB, { 0x91, 0x16, 0x98, 0x4F, 0x2F, 0x4B, 0x08, 0x2A} };

//{9304C900-E16C-40E3-BB0E-638131063793}
IMPLEMENT_OLECREATE_FLAGS(Resources, "MSP2010.Resources", afxRegApartmentThreading, 0x9304C900, 0xE16C, 0x40E3, 0xBB, 0x0E, 0x63, 0x81, 0x31, 0x06, 0x37, 0x93)

BEGIN_DISPATCH_MAP(Resources, CCmdTarget)
DISP_PROPERTY_EX_ID(Resources, "Count", 1, odl_GetCount, SetNotSupported, VT_I4)
DISP_PROPERTY_PARAM_ID(Resources, "Item", 2, odl_Item, SetNotSupported, VT_DISPATCH, VTS_BSTR)
DISP_FUNCTION_ID(Resources, "Add", 3, odl_Add, VT_DISPATCH, VTS_NONE)
DISP_FUNCTION_ID(Resources, "Clear", 4, odl_Clear, VT_EMPTY, VTS_NONE)
DISP_FUNCTION_ID(Resources, "Remove", 5, odl_Remove, VT_EMPTY, VTS_BSTR)
DISP_FUNCTION_ID(Resources, "IsNull", 6, IsNull, VT_BOOL, VTS_NONE)
DISP_FUNCTION_ID(Resources, "Initialize", 7, Initialize, VT_EMPTY, VTS_NONE)
DISP_FUNCTION_ID(Resources, "SetXML", 8, odl_SetXML, VT_EMPTY, VTS_BSTR)
DISP_FUNCTION_ID(Resources, "GetXML", 9, odl_GetXML, VT_BSTR, VTS_NONE)
DISP_DEFVALUE(Resources, "Item")
END_DISPATCH_MAP()

BEGIN_INTERFACE_MAP(Resources, CCmdTarget)
	INTERFACE_PART(Resources, IID_IResources, Dispatch)
END_INTERFACE_MAP()

BEGIN_MESSAGE_MAP(Resources, CCmdTarget)
END_MESSAGE_MAP()


Resources::Resources()
{
	EnableAutomation();
	AfxOleLockApp();
	InitVars();
}

void Resources::Initialize(void)
{
	InitVars();
}

void Resources::InitVars(void)
{
	mp_oCollection = new clsCollectionBase("Resource");
}

Resources::~Resources()
{
	delete mp_oCollection;
	AfxOleUnlockApp();
}

void Resources::OnFinalRelease()
{
	CCmdTarget::OnFinalRelease();
}

LONG Resources::odl_GetCount(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetCount();
}

LONG Resources::GetCount(void)
{
	return mp_oCollection->m_lCount();
}

IDispatch* Resources::odl_Item(LPCTSTR Index)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return Item(Index)->GetIDispatch(TRUE);
}

IDispatch* Resources::odl_Add(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return Add()->GetIDispatch(TRUE);
}

void Resources::odl_Clear(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	Clear();
}

void Resources::odl_Remove(LPCTSTR Index)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	Remove(Index);
}

BOOL Resources::IsNull(void)
{
	BOOL bReturn = TRUE;
	if (GetCount() > 0)
	{
		bReturn = FALSE;
	}
	return bReturn;
}

Resource* Resources::Item(CString Index)
{
	Resource *oResource;
	oResource = (Resource*)mp_oCollection->m_oItem(Index, MP_ITEM_1, MP_ITEM_2, MP_ITEM_3, MP_ITEM_4);
	return oResource;
}

Resource* Resources::Add()
{
	mp_oCollection->SetAddMode(TRUE);
	Resource* oResource = new Resource();
	oResource->mp_oCollection = mp_oCollection;
	mp_oCollection->m_Add(oResource, _T(""), MP_ADD_1, MP_ADD_2, FALSE, MP_ADD_3);
	return oResource;
}

void Resources::Clear(void)
{
	mp_oCollection->m_Clear();
}

void Resources::Remove(CString Index)
{
	mp_oCollection->m_Remove(Index, MP_REMOVE_1, MP_REMOVE_2, MP_REMOVE_3, MP_REMOVE_4);
}

BSTR Resources::odl_GetXML(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetXML().AllocSysString();
}

void Resources::odl_SetXML(LPCTSTR sXML)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetXML(sXML);
}

CString Resources::GetXML(void)
{
	if (IsNull() == TRUE)
	{
		return "<Resources/>";
	}
	LONG lIndex;
	Resource* oResource;
	clsXML oXML("Resources");
	oXML.InitializeWriter();
	for (lIndex = 1; lIndex <= GetCount(); lIndex++)
	{
		oResource = (Resource*) mp_oCollection->m_oReturnArrayElement(lIndex);
		oXML.WriteObject(oResource->GetXML());
	}
	return oXML.GetXML();
}

void Resources::SetXML(CString sXML)
{
	LONG lIndex;
	clsXML oXML("Resources");
	oXML.SetSupportOptional(TRUE);
	oXML.SetXML(sXML);
	oXML.InitializeReader();
	mp_oCollection->m_Clear();
	for (lIndex = 1;lIndex <= oXML.ReadCollectionCount();lIndex++)
	{
		Resource* oResource = new Resource();
		oResource->SetXML(oXML.ReadCollectionObject(lIndex));
		mp_oCollection->SetAddMode(TRUE);
		CString sKey = _T("");
		sKey.Format(_T("%d"), oResource->GetUID());
		sKey = _T("K") + sKey;
		oResource->mp_oCollection = mp_oCollection;
		oResource->SetKey(sKey);
		mp_oCollection->m_Add(oResource, sKey, MP_ADD_1, MP_ADD_2, FALSE, MP_ADD_3);
	}
}