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

//{7E699717-38ED-4600-8CC8-6F005D24A550}
static const IID IID_IResourceOutlineCode_C = { 0x7E699717, 0x38ED, 0x4600, { 0x8C, 0xC8, 0x6F, 0x00, 0x5D, 0x24, 0xA5, 0x50} };

//{34C55A84-356E-4370-9027-FCB2DD6CA17B}
IMPLEMENT_OLECREATE_FLAGS(ResourceOutlineCode_C, "MSP2003.ResourceOutlineCode_C", afxRegApartmentThreading, 0x34C55A84, 0x356E, 0x4370, 0x90, 0x27, 0xFC, 0xB2, 0xDD, 0x6C, 0xA1, 0x7B)

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