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

//{F2ED31AC-CB50-4A0E-9EA2-A4F6AE44E29E}
static const IID IID_IResourceRates = { 0xF2ED31AC, 0xCB50, 0x4A0E, { 0x9E, 0xA2, 0xA4, 0xF6, 0xAE, 0x44, 0xE2, 0x9E} };

//{B4C01870-961A-4E98-998B-145BB8FE3912}
IMPLEMENT_OLECREATE_FLAGS(ResourceRates, "MSP2010.ResourceRates", afxRegApartmentThreading, 0xB4C01870, 0x961A, 0x4E98, 0x99, 0x8B, 0x14, 0x5B, 0xB8, 0xFE, 0x39, 0x12)

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