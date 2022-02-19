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
#include "ResourceOutlineCode_C.h"

IMPLEMENT_DYNCREATE(ResourceOutlineCode_C, CCmdTarget)

//{83EF7CB3-BA3D-4159-AD5C-05471E57CCD6}
static const IID IID_IResourceOutlineCode_C = { 0x83EF7CB3, 0xBA3D, 0x4159, { 0xAD, 0x5C, 0x05, 0x47, 0x1E, 0x57, 0xCC, 0xD6} };

//{3328F9F5-58CB-4AC1-9DE6-A9B0EB4C83FD}
IMPLEMENT_OLECREATE_FLAGS(ResourceOutlineCode_C, "MSP2007.ResourceOutlineCode_C", afxRegApartmentThreading, 0x3328F9F5, 0x58CB, 0x4AC1, 0x9D, 0xE6, 0xA9, 0xB0, 0xEB, 0x4C, 0x83, 0xFD)

BEGIN_DISPATCH_MAP(ResourceOutlineCode_C, CCmdTarget)
DISP_PROPERTY_EX_ID(ResourceOutlineCode_C, "Count", 1, odl_GetCount, SetNotSupported, VT_I4)
DISP_PROPERTY_PARAM_ID(ResourceOutlineCode_C, "Item", 2, odl_Item, SetNotSupported, VT_DISPATCH, VTS_BSTR)
DISP_FUNCTION_ID(ResourceOutlineCode_C, "Add", 3, odl_Add, VT_DISPATCH, VTS_NONE)
DISP_FUNCTION_ID(ResourceOutlineCode_C, "Clear", 4, odl_Clear, VT_EMPTY, VTS_NONE)
DISP_FUNCTION_ID(ResourceOutlineCode_C, "Remove", 5, odl_Remove, VT_EMPTY, VTS_BSTR)
DISP_FUNCTION_ID(ResourceOutlineCode_C, "IsNull", 6, IsNull, VT_BOOL, VTS_NONE)
DISP_FUNCTION_ID(ResourceOutlineCode_C, "Initialize", 7, Initialize, VT_EMPTY, VTS_NONE)
DISP_DEFVALUE(ResourceOutlineCode_C, "Item")
END_DISPATCH_MAP()

BEGIN_INTERFACE_MAP(ResourceOutlineCode_C, CCmdTarget)
	INTERFACE_PART(ResourceOutlineCode_C, IID_IResourceOutlineCode_C, Dispatch)
END_INTERFACE_MAP()

BEGIN_MESSAGE_MAP(ResourceOutlineCode_C, CCmdTarget)
END_MESSAGE_MAP()


ResourceOutlineCode_C::ResourceOutlineCode_C()
{
	EnableAutomation();
	AfxOleLockApp();
	InitVars();
}

void ResourceOutlineCode_C::Initialize(void)
{
	InitVars();
}

void ResourceOutlineCode_C::InitVars(void)
{
	mp_oCollection = new clsCollectionBase("ResourceOutlineCode");
}

ResourceOutlineCode_C::~ResourceOutlineCode_C()
{
	delete mp_oCollection;
	AfxOleUnlockApp();
}

void ResourceOutlineCode_C::OnFinalRelease()
{
	CCmdTarget::OnFinalRelease();
}

LONG ResourceOutlineCode_C::odl_GetCount(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetCount();
}

LONG ResourceOutlineCode_C::GetCount(void)
{
	return mp_oCollection->m_lCount();
}

IDispatch* ResourceOutlineCode_C::odl_Item(LPCTSTR Index)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return Item(Index)->GetIDispatch(TRUE);
}

IDispatch* ResourceOutlineCode_C::odl_Add(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return Add()->GetIDispatch(TRUE);
}

void ResourceOutlineCode_C::odl_Clear(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	Clear();
}

void ResourceOutlineCode_C::odl_Remove(LPCTSTR Index)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	Remove(Index);
}

BOOL ResourceOutlineCode_C::IsNull(void)
{
	BOOL bReturn = TRUE;
	if (GetCount() > 0)
	{
		bReturn = FALSE;
	}
	return bReturn;
}

ResourceOutlineCode* ResourceOutlineCode_C::Item(CString Index)
{
	ResourceOutlineCode *oResourceOutlineCode;
	oResourceOutlineCode = (ResourceOutlineCode*)mp_oCollection->m_oItem(Index, MP_ITEM_1, MP_ITEM_2, MP_ITEM_3, MP_ITEM_4);
	return oResourceOutlineCode;
}

ResourceOutlineCode* ResourceOutlineCode_C::Add()
{
	mp_oCollection->SetAddMode(TRUE);
	ResourceOutlineCode* oResourceOutlineCode = new ResourceOutlineCode();
	oResourceOutlineCode->mp_oCollection = mp_oCollection;
	mp_oCollection->m_Add(oResourceOutlineCode, _T(""), MP_ADD_1, MP_ADD_2, FALSE, MP_ADD_3);
	return oResourceOutlineCode;
}

void ResourceOutlineCode_C::Clear(void)
{
	mp_oCollection->m_Clear();
}

void ResourceOutlineCode_C::Remove(CString Index)
{
	mp_oCollection->m_Remove(Index, MP_REMOVE_1, MP_REMOVE_2, MP_REMOVE_3, MP_REMOVE_4);
}

void ResourceOutlineCode_C::ReadObjectProtected(clsXML &oXML)
{
	LONG lIndex;
	for (lIndex = 1; lIndex <= oXML.ReadCollectionCount(); lIndex++)
	{
		if (oXML.GetCollectionObjectName(lIndex) == "OutlineCode")
		{
			ResourceOutlineCode* oResourceOutlineCode = new ResourceOutlineCode();
			oResourceOutlineCode->SetXML(oXML.ReadCollectionObject(lIndex));
			mp_oCollection->SetAddMode(TRUE);
			CString sKey = _T("");
			oResourceOutlineCode->mp_oCollection = mp_oCollection;
			mp_oCollection->m_Add(oResourceOutlineCode, sKey, MP_ADD_1, MP_ADD_2, FALSE, MP_ADD_3);
		}
	}
}

void ResourceOutlineCode_C::WriteObjectProtected(clsXML &oXML)
{
	LONG lIndex;
	ResourceOutlineCode* oResourceOutlineCode;
	for (lIndex = 1; lIndex <= GetCount(); lIndex++)
	{
		oResourceOutlineCode = (ResourceOutlineCode*) mp_oCollection->m_oReturnArrayElement(lIndex);
		oXML.WriteObject(oResourceOutlineCode->GetXML());
	}
}