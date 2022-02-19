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

//{3DA5ED23-EB19-48BD-BB13-D9498662F02D}
static const IID IID_IResourceExtendedAttribute_C = { 0x3DA5ED23, 0xEB19, 0x48BD, { 0xBB, 0x13, 0xD9, 0x49, 0x86, 0x62, 0xF0, 0x2D} };

//{E889C7E3-9430-4720-A07D-126AB90F0776}
IMPLEMENT_OLECREATE_FLAGS(ResourceExtendedAttribute_C, "MSP2007.ResourceExtendedAttribute_C", afxRegApartmentThreading, 0xE889C7E3, 0x9430, 0x4720, 0xA0, 0x7D, 0x12, 0x6A, 0xB9, 0x0F, 0x07, 0x76)

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