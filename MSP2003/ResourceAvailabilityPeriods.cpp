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
#include "ResourceAvailabilityPeriods.h"

IMPLEMENT_DYNCREATE(ResourceAvailabilityPeriods, CCmdTarget)

//{6D17C8FA-96E3-48AD-AC08-D5BE0771008F}
static const IID IID_IResourceAvailabilityPeriods = { 0x6D17C8FA, 0x96E3, 0x48AD, { 0xAC, 0x08, 0xD5, 0xBE, 0x07, 0x71, 0x00, 0x8F} };

//{50172AEB-845C-472E-A206-5B958BF3E0C4}
IMPLEMENT_OLECREATE_FLAGS(ResourceAvailabilityPeriods, "MSP2003.ResourceAvailabilityPeriods", afxRegApartmentThreading, 0x50172AEB, 0x845C, 0x472E, 0xA2, 0x06, 0x5B, 0x95, 0x8B, 0xF3, 0xE0, 0xC4)

BEGIN_DISPATCH_MAP(ResourceAvailabilityPeriods, CCmdTarget)
DISP_PROPERTY_EX_ID(ResourceAvailabilityPeriods, "Count", 1, odl_GetCount, SetNotSupported, VT_I4)
DISP_PROPERTY_PARAM_ID(ResourceAvailabilityPeriods, "Item", 2, odl_Item, SetNotSupported, VT_DISPATCH, VTS_BSTR)
DISP_FUNCTION_ID(ResourceAvailabilityPeriods, "Add", 3, odl_Add, VT_DISPATCH, VTS_NONE)
DISP_FUNCTION_ID(ResourceAvailabilityPeriods, "Clear", 4, odl_Clear, VT_EMPTY, VTS_NONE)
DISP_FUNCTION_ID(ResourceAvailabilityPeriods, "Remove", 5, odl_Remove, VT_EMPTY, VTS_BSTR)
DISP_FUNCTION_ID(ResourceAvailabilityPeriods, "IsNull", 6, IsNull, VT_BOOL, VTS_NONE)
DISP_FUNCTION_ID(ResourceAvailabilityPeriods, "Initialize", 7, Initialize, VT_EMPTY, VTS_NONE)
DISP_FUNCTION_ID(ResourceAvailabilityPeriods, "SetXML", 8, odl_SetXML, VT_EMPTY, VTS_BSTR)
DISP_FUNCTION_ID(ResourceAvailabilityPeriods, "GetXML", 9, odl_GetXML, VT_BSTR, VTS_NONE)
DISP_DEFVALUE(ResourceAvailabilityPeriods, "Item")
END_DISPATCH_MAP()

BEGIN_INTERFACE_MAP(ResourceAvailabilityPeriods, CCmdTarget)
	INTERFACE_PART(ResourceAvailabilityPeriods, IID_IResourceAvailabilityPeriods, Dispatch)
END_INTERFACE_MAP()

BEGIN_MESSAGE_MAP(ResourceAvailabilityPeriods, CCmdTarget)
END_MESSAGE_MAP()


ResourceAvailabilityPeriods::ResourceAvailabilityPeriods()
{
	EnableAutomation();
	AfxOleLockApp();
	InitVars();
}

void ResourceAvailabilityPeriods::Initialize(void)
{
	InitVars();
}

void ResourceAvailabilityPeriods::InitVars(void)
{
	mp_oCollection = new clsCollectionBase("ResourceAvailabilityPeriod");
}

ResourceAvailabilityPeriods::~ResourceAvailabilityPeriods()
{
	delete mp_oCollection;
	AfxOleUnlockApp();
}

void ResourceAvailabilityPeriods::OnFinalRelease()
{
	CCmdTarget::OnFinalRelease();
}

LONG ResourceAvailabilityPeriods::odl_GetCount(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetCount();
}

LONG ResourceAvailabilityPeriods::GetCount(void)
{
	return mp_oCollection->m_lCount();
}

IDispatch* ResourceAvailabilityPeriods::odl_Item(LPCTSTR Index)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return Item(Index)->GetIDispatch(TRUE);
}

IDispatch* ResourceAvailabilityPeriods::odl_Add(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return Add()->GetIDispatch(TRUE);
}

void ResourceAvailabilityPeriods::odl_Clear(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	Clear();
}

void ResourceAvailabilityPeriods::odl_Remove(LPCTSTR Index)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	Remove(Index);
}

BOOL ResourceAvailabilityPeriods::IsNull(void)
{
	BOOL bReturn = TRUE;
	if (GetCount() > 0)
	{
		bReturn = FALSE;
	}
	return bReturn;
}

ResourceAvailabilityPeriod* ResourceAvailabilityPeriods::Item(CString Index)
{
	ResourceAvailabilityPeriod *oResourceAvailabilityPeriod;
	oResourceAvailabilityPeriod = (ResourceAvailabilityPeriod*)mp_oCollection->m_oItem(Index, MP_ITEM_1, MP_ITEM_2, MP_ITEM_3, MP_ITEM_4);
	return oResourceAvailabilityPeriod;
}

ResourceAvailabilityPeriod* ResourceAvailabilityPeriods::Add()
{
	mp_oCollection->SetAddMode(TRUE);
	ResourceAvailabilityPeriod* oResourceAvailabilityPeriod = new ResourceAvailabilityPeriod();
	oResourceAvailabilityPeriod->mp_oCollection = mp_oCollection;
	mp_oCollection->m_Add(oResourceAvailabilityPeriod, _T(""), MP_ADD_1, MP_ADD_2, FALSE, MP_ADD_3);
	return oResourceAvailabilityPeriod;
}

void ResourceAvailabilityPeriods::Clear(void)
{
	mp_oCollection->m_Clear();
}

void ResourceAvailabilityPeriods::Remove(CString Index)
{
	mp_oCollection->m_Remove(Index, MP_REMOVE_1, MP_REMOVE_2, MP_REMOVE_3, MP_REMOVE_4);
}

BSTR ResourceAvailabilityPeriods::odl_GetXML(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetXML().AllocSysString();
}

void ResourceAvailabilityPeriods::odl_SetXML(LPCTSTR sXML)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetXML(sXML);
}

CString ResourceAvailabilityPeriods::GetXML(void)
{
	if (IsNull() == TRUE)
	{
		return "<AvailabilityPeriods/>";
	}
	LONG lIndex;
	ResourceAvailabilityPeriod* oResourceAvailabilityPeriod;
	clsXML oXML("AvailabilityPeriods");
	oXML.InitializeWriter();
	for (lIndex = 1; lIndex <= GetCount(); lIndex++)
	{
		oResourceAvailabilityPeriod = (ResourceAvailabilityPeriod*) mp_oCollection->m_oReturnArrayElement(lIndex);
		oXML.WriteObject(oResourceAvailabilityPeriod->GetXML());
	}
	return oXML.GetXML();
}

void ResourceAvailabilityPeriods::SetXML(CString sXML)
{
	LONG lIndex;
	clsXML oXML("AvailabilityPeriods");
	oXML.SetSupportOptional(TRUE);
	oXML.SetXML(sXML);
	oXML.InitializeReader();
	mp_oCollection->m_Clear();
	for (lIndex = 1;lIndex <= oXML.ReadCollectionCount();lIndex++)
	{
		ResourceAvailabilityPeriod* oResourceAvailabilityPeriod = new ResourceAvailabilityPeriod();
		oResourceAvailabilityPeriod->SetXML(oXML.ReadCollectionObject(lIndex));
		mp_oCollection->SetAddMode(TRUE);
		CString sKey = _T("");
		oResourceAvailabilityPeriod->mp_oCollection = mp_oCollection;
		mp_oCollection->m_Add(oResourceAvailabilityPeriod, sKey, MP_ADD_1, MP_ADD_2, FALSE, MP_ADD_3);
	}
}