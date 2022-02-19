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
#include "ResourceExtendedAttribute_C.h"

IMPLEMENT_DYNCREATE(ResourceExtendedAttribute_C, CCmdTarget)

//{0DA61E3A-0EA5-4EDE-9ADA-615E8F5301A3}
static const IID IID_IResourceExtendedAttribute_C = { 0x0DA61E3A, 0x0EA5, 0x4EDE, { 0x9A, 0xDA, 0x61, 0x5E, 0x8F, 0x53, 0x01, 0xA3} };

//{1C03CFC6-5BDF-44EF-86FF-A3AF7430D644}
IMPLEMENT_OLECREATE_FLAGS(ResourceExtendedAttribute_C, "MSP2010.ResourceExtendedAttribute_C", afxRegApartmentThreading, 0x1C03CFC6, 0x5BDF, 0x44EF, 0x86, 0xFF, 0xA3, 0xAF, 0x74, 0x30, 0xD6, 0x44)

BEGIN_DISPATCH_MAP(ResourceExtendedAttribute_C, CCmdTarget)
DISP_PROPERTY_EX_ID(ResourceExtendedAttribute_C, "Count", 1, odl_GetCount, SetNotSupported, VT_I4)
DISP_PROPERTY_PARAM_ID(ResourceExtendedAttribute_C, "Item", 2, odl_Item, SetNotSupported, VT_DISPATCH, VTS_BSTR)
DISP_FUNCTION_ID(ResourceExtendedAttribute_C, "Add", 3, odl_Add, VT_DISPATCH, VTS_NONE)
DISP_FUNCTION_ID(ResourceExtendedAttribute_C, "Clear", 4, odl_Clear, VT_EMPTY, VTS_NONE)
DISP_FUNCTION_ID(ResourceExtendedAttribute_C, "Remove", 5, odl_Remove, VT_EMPTY, VTS_BSTR)
DISP_FUNCTION_ID(ResourceExtendedAttribute_C, "IsNull", 6, IsNull, VT_BOOL, VTS_NONE)
DISP_FUNCTION_ID(ResourceExtendedAttribute_C, "Initialize", 7, Initialize, VT_EMPTY, VTS_NONE)
DISP_DEFVALUE(ResourceExtendedAttribute_C, "Item")
END_DISPATCH_MAP()

BEGIN_INTERFACE_MAP(ResourceExtendedAttribute_C, CCmdTarget)
	INTERFACE_PART(ResourceExtendedAttribute_C, IID_IResourceExtendedAttribute_C, Dispatch)
END_INTERFACE_MAP()

BEGIN_MESSAGE_MAP(ResourceExtendedAttribute_C, CCmdTarget)
END_MESSAGE_MAP()


ResourceExtendedAttribute_C::ResourceExtendedAttribute_C()
{
	EnableAutomation();
	AfxOleLockApp();
	InitVars();
}

void ResourceExtendedAttribute_C::Initialize(void)
{
	InitVars();
}

void ResourceExtendedAttribute_C::InitVars(void)
{
	mp_oCollection = new clsCollectionBase("ResourceExtendedAttribute");
}

ResourceExtendedAttribute_C::~ResourceExtendedAttribute_C()
{
	delete mp_oCollection;
	AfxOleUnlockApp();
}

void ResourceExtendedAttribute_C::OnFinalRelease()
{
	CCmdTarget::OnFinalRelease();
}

LONG ResourceExtendedAttribute_C::odl_GetCount(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetCount();
}

LONG ResourceExtendedAttribute_C::GetCount(void)
{
	return mp_oCollection->m_lCount();
}

IDispatch* ResourceExtendedAttribute_C::odl_Item(LPCTSTR Index)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return Item(Index)->GetIDispatch(TRUE);
}

IDispatch* ResourceExtendedAttribute_C::odl_Add(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return Add()->GetIDispatch(TRUE);
}

void ResourceExtendedAttribute_C::odl_Clear(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	Clear();
}

void ResourceExtendedAttribute_C::odl_Remove(LPCTSTR Index)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	Remove(Index);
}

BOOL ResourceExtendedAttribute_C::IsNull(void)
{
	BOOL bReturn = TRUE;
	if (GetCount() > 0)
	{
		bReturn = FALSE;
	}
	return bReturn;
}

ResourceExtendedAttribute* ResourceExtendedAttribute_C::Item(CString Index)
{
	ResourceExtendedAttribute *oResourceExtendedAttribute;
	oResourceExtendedAttribute = (ResourceExtendedAttribute*)mp_oCollection->m_oItem(Index, MP_ITEM_1, MP_ITEM_2, MP_ITEM_3, MP_ITEM_4);
	return oResourceExtendedAttribute;
}

ResourceExtendedAttribute* ResourceExtendedAttribute_C::Add()
{
	mp_oCollection->SetAddMode(TRUE);
	ResourceExtendedAttribute* oResourceExtendedAttribute = new ResourceExtendedAttribute();
	oResourceExtendedAttribute->mp_oCollection = mp_oCollection;
	mp_oCollection->m_Add(oResourceExtendedAttribute, _T(""), MP_ADD_1, MP_ADD_2, FALSE, MP_ADD_3);
	return oResourceExtendedAttribute;
}

void ResourceExtendedAttribute_C::Clear(void)
{
	mp_oCollection->m_Clear();
}

void ResourceExtendedAttribute_C::Remove(CString Index)
{
	mp_oCollection->m_Remove(Index, MP_REMOVE_1, MP_REMOVE_2, MP_REMOVE_3, MP_REMOVE_4);
}

void ResourceExtendedAttribute_C::ReadObjectProtected(clsXML &oXML)
{
	LONG lIndex;
	for (lIndex = 1; lIndex <= oXML.ReadCollectionCount(); lIndex++)
	{
		if (oXML.GetCollectionObjectName(lIndex) == "ExtendedAttribute")
		{
			ResourceExtendedAttribute* oResourceExtendedAttribute = new ResourceExtendedAttribute();
			oResourceExtendedAttribute->SetXML(oXML.ReadCollectionObject(lIndex));
			mp_oCollection->SetAddMode(TRUE);
			CString sKey = _T("");
			oResourceExtendedAttribute->mp_oCollection = mp_oCollection;
			mp_oCollection->m_Add(oResourceExtendedAttribute, sKey, MP_ADD_1, MP_ADD_2, FALSE, MP_ADD_3);
		}
	}
}

void ResourceExtendedAttribute_C::WriteObjectProtected(clsXML &oXML)
{
	LONG lIndex;
	ResourceExtendedAttribute* oResourceExtendedAttribute;
	for (lIndex = 1; lIndex <= GetCount(); lIndex++)
	{
		oResourceExtendedAttribute = (ResourceExtendedAttribute*) mp_oCollection->m_oReturnArrayElement(lIndex);
		oXML.WriteObject(oResourceExtendedAttribute->GetXML());
	}
}