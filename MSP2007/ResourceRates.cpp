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
#include "ResourceRates.h"

IMPLEMENT_DYNCREATE(ResourceRates, CCmdTarget)

//{ABF5C8D3-00F4-4083-87F2-B45B91751695}
static const IID IID_IResourceRates = { 0xABF5C8D3, 0x00F4, 0x4083, { 0x87, 0xF2, 0xB4, 0x5B, 0x91, 0x75, 0x16, 0x95} };

//{F37BA1A1-5DA9-42EB-87D7-2B87D71502FD}
IMPLEMENT_OLECREATE_FLAGS(ResourceRates, "MSP2007.ResourceRates", afxRegApartmentThreading, 0xF37BA1A1, 0x5DA9, 0x42EB, 0x87, 0xD7, 0x2B, 0x87, 0xD7, 0x15, 0x02, 0xFD)

BEGIN_DISPATCH_MAP(ResourceRates, CCmdTarget)
DISP_PROPERTY_EX_ID(ResourceRates, "Count", 1, odl_GetCount, SetNotSupported, VT_I4)
DISP_PROPERTY_PARAM_ID(ResourceRates, "Item", 2, odl_Item, SetNotSupported, VT_DISPATCH, VTS_BSTR)
DISP_FUNCTION_ID(ResourceRates, "Add", 3, odl_Add, VT_DISPATCH, VTS_NONE)
DISP_FUNCTION_ID(ResourceRates, "Clear", 4, odl_Clear, VT_EMPTY, VTS_NONE)
DISP_FUNCTION_ID(ResourceRates, "Remove", 5, odl_Remove, VT_EMPTY, VTS_BSTR)
DISP_FUNCTION_ID(ResourceRates, "IsNull", 6, IsNull, VT_BOOL, VTS_NONE)
DISP_FUNCTION_ID(ResourceRates, "Initialize", 7, Initialize, VT_EMPTY, VTS_NONE)
DISP_FUNCTION_ID(ResourceRates, "SetXML", 8, odl_SetXML, VT_EMPTY, VTS_BSTR)
DISP_FUNCTION_ID(ResourceRates, "GetXML", 9, odl_GetXML, VT_BSTR, VTS_NONE)
DISP_DEFVALUE(ResourceRates, "Item")
END_DISPATCH_MAP()

BEGIN_INTERFACE_MAP(ResourceRates, CCmdTarget)
	INTERFACE_PART(ResourceRates, IID_IResourceRates, Dispatch)
END_INTERFACE_MAP()

BEGIN_MESSAGE_MAP(ResourceRates, CCmdTarget)
END_MESSAGE_MAP()


ResourceRates::ResourceRates()
{
	EnableAutomation();
	AfxOleLockApp();
	InitVars();
}

void ResourceRates::Initialize(void)
{
	InitVars();
}

void ResourceRates::InitVars(void)
{
	mp_oCollection = new clsCollectionBase("ResourceRate");
}

ResourceRates::~ResourceRates()
{
	delete mp_oCollection;
	AfxOleUnlockApp();
}

void ResourceRates::OnFinalRelease()
{
	CCmdTarget::OnFinalRelease();
}

LONG ResourceRates::odl_GetCount(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetCount();
}

LONG ResourceRates::GetCount(void)
{
	return mp_oCollection->m_lCount();
}

IDispatch* ResourceRates::odl_Item(LPCTSTR Index)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return Item(Index)->GetIDispatch(TRUE);
}

IDispatch* ResourceRates::odl_Add(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return Add()->GetIDispatch(TRUE);
}

void ResourceRates::odl_Clear(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	Clear();
}

void ResourceRates::odl_Remove(LPCTSTR Index)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	Remove(Index);
}

BOOL ResourceRates::IsNull(void)
{
	BOOL bReturn = TRUE;
	if (GetCount() > 0)
	{
		bReturn = FALSE;
	}
	return bReturn;
}

ResourceRate* ResourceRates::Item(CString Index)
{
	ResourceRate *oResourceRate;
	oResourceRate = (ResourceRate*)mp_oCollection->m_oItem(Index, MP_ITEM_1, MP_ITEM_2, MP_ITEM_3, MP_ITEM_4);
	return oResourceRate;
}

ResourceRate* ResourceRates::Add()
{
	mp_oCollection->SetAddMode(TRUE);
	ResourceRate* oResourceRate = new ResourceRate();
	oResourceRate->mp_oCollection = mp_oCollection;
	mp_oCollection->m_Add(oResourceRate, _T(""), MP_ADD_1, MP_ADD_2, FALSE, MP_ADD_3);
	return oResourceRate;
}

void ResourceRates::Clear(void)
{
	mp_oCollection->m_Clear();
}

void ResourceRates::Remove(CString Index)
{
	mp_oCollection->m_Remove(Index, MP_REMOVE_1, MP_REMOVE_2, MP_REMOVE_3, MP_REMOVE_4);
}

BSTR ResourceRates::odl_GetXML(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetXML().AllocSysString();
}

void ResourceRates::odl_SetXML(LPCTSTR sXML)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetXML(sXML);
}

CString ResourceRates::GetXML(void)
{
	if (IsNull() == TRUE)
	{
		return "<Rates/>";
	}
	LONG lIndex;
	ResourceRate* oResourceRate;
	clsXML oXML("Rates");
	oXML.InitializeWriter();
	for (lIndex = 1; lIndex <= GetCount(); lIndex++)
	{
		oResourceRate = (ResourceRate*) mp_oCollection->m_oReturnArrayElement(lIndex);
		oXML.WriteObject(oResourceRate->GetXML());
	}
	return oXML.GetXML();
}

void ResourceRates::SetXML(CString sXML)
{
	LONG lIndex;
	clsXML oXML("Rates");
	oXML.SetSupportOptional(TRUE);
	oXML.SetXML(sXML);
	oXML.InitializeReader();
	mp_oCollection->m_Clear();
	for (lIndex = 1;lIndex <= oXML.ReadCollectionCount();lIndex++)
	{
		ResourceRate* oResourceRate = new ResourceRate();
		oResourceRate->SetXML(oXML.ReadCollectionObject(lIndex));
		mp_oCollection->SetAddMode(TRUE);
		CString sKey = _T("");
		oResourceRate->mp_oCollection = mp_oCollection;
		mp_oCollection->m_Add(oResourceRate, sKey, MP_ADD_1, MP_ADD_2, FALSE, MP_ADD_3);
	}
}