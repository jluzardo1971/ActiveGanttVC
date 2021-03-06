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
#include "ResourceBaseline_C.h"

IMPLEMENT_DYNCREATE(ResourceBaseline_C, CCmdTarget)

//{D969F2DC-1758-4DB9-A26B-971F39A5037E}
static const IID IID_IResourceBaseline_C = { 0xD969F2DC, 0x1758, 0x4DB9, { 0xA2, 0x6B, 0x97, 0x1F, 0x39, 0xA5, 0x03, 0x7E} };

//{77CC12D1-7703-4349-BE7A-1EE532B15B01}
IMPLEMENT_OLECREATE_FLAGS(ResourceBaseline_C, "MSP2007.ResourceBaseline_C", afxRegApartmentThreading, 0x77CC12D1, 0x7703, 0x4349, 0xBE, 0x7A, 0x1E, 0xE5, 0x32, 0xB1, 0x5B, 0x01)

BEGIN_DISPATCH_MAP(ResourceBaseline_C, CCmdTarget)
DISP_PROPERTY_EX_ID(ResourceBaseline_C, "Count", 1, odl_GetCount, SetNotSupported, VT_I4)
DISP_PROPERTY_PARAM_ID(ResourceBaseline_C, "Item", 2, odl_Item, SetNotSupported, VT_DISPATCH, VTS_BSTR)
DISP_FUNCTION_ID(ResourceBaseline_C, "Add", 3, odl_Add, VT_DISPATCH, VTS_NONE)
DISP_FUNCTION_ID(ResourceBaseline_C, "Clear", 4, odl_Clear, VT_EMPTY, VTS_NONE)
DISP_FUNCTION_ID(ResourceBaseline_C, "Remove", 5, odl_Remove, VT_EMPTY, VTS_BSTR)
DISP_FUNCTION_ID(ResourceBaseline_C, "IsNull", 6, IsNull, VT_BOOL, VTS_NONE)
DISP_FUNCTION_ID(ResourceBaseline_C, "Initialize", 7, Initialize, VT_EMPTY, VTS_NONE)
DISP_DEFVALUE(ResourceBaseline_C, "Item")
END_DISPATCH_MAP()

BEGIN_INTERFACE_MAP(ResourceBaseline_C, CCmdTarget)
	INTERFACE_PART(ResourceBaseline_C, IID_IResourceBaseline_C, Dispatch)
END_INTERFACE_MAP()

BEGIN_MESSAGE_MAP(ResourceBaseline_C, CCmdTarget)
END_MESSAGE_MAP()


ResourceBaseline_C::ResourceBaseline_C()
{
	EnableAutomation();
	AfxOleLockApp();
	InitVars();
}

void ResourceBaseline_C::Initialize(void)
{
	InitVars();
}

void ResourceBaseline_C::InitVars(void)
{
	mp_oCollection = new clsCollectionBase("ResourceBaseline");
}

ResourceBaseline_C::~ResourceBaseline_C()
{
	delete mp_oCollection;
	AfxOleUnlockApp();
}

void ResourceBaseline_C::OnFinalRelease()
{
	CCmdTarget::OnFinalRelease();
}

LONG ResourceBaseline_C::odl_GetCount(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetCount();
}

LONG ResourceBaseline_C::GetCount(void)
{
	return mp_oCollection->m_lCount();
}

IDispatch* ResourceBaseline_C::odl_Item(LPCTSTR Index)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return Item(Index)->GetIDispatch(TRUE);
}

IDispatch* ResourceBaseline_C::odl_Add(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return Add()->GetIDispatch(TRUE);
}

void ResourceBaseline_C::odl_Clear(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	Clear();
}

void ResourceBaseline_C::odl_Remove(LPCTSTR Index)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	Remove(Index);
}

BOOL ResourceBaseline_C::IsNull(void)
{
	BOOL bReturn = TRUE;
	if (GetCount() > 0)
	{
		bReturn = FALSE;
	}
	return bReturn;
}

ResourceBaseline* ResourceBaseline_C::Item(CString Index)
{
	ResourceBaseline *oResourceBaseline;
	oResourceBaseline = (ResourceBaseline*)mp_oCollection->m_oItem(Index, MP_ITEM_1, MP_ITEM_2, MP_ITEM_3, MP_ITEM_4);
	return oResourceBaseline;
}

ResourceBaseline* ResourceBaseline_C::Add()
{
	mp_oCollection->SetAddMode(TRUE);
	ResourceBaseline* oResourceBaseline = new ResourceBaseline();
	oResourceBaseline->mp_oCollection = mp_oCollection;
	mp_oCollection->m_Add(oResourceBaseline, _T(""), MP_ADD_1, MP_ADD_2, FALSE, MP_ADD_3);
	return oResourceBaseline;
}

void ResourceBaseline_C::Clear(void)
{
	mp_oCollection->m_Clear();
}

void ResourceBaseline_C::Remove(CString Index)
{
	mp_oCollection->m_Remove(Index, MP_REMOVE_1, MP_REMOVE_2, MP_REMOVE_3, MP_REMOVE_4);
}

void ResourceBaseline_C::ReadObjectProtected(clsXML &oXML)
{
	LONG lIndex;
	for (lIndex = 1; lIndex <= oXML.ReadCollectionCount(); lIndex++)
	{
		if (oXML.GetCollectionObjectName(lIndex) == "Baseline")
		{
			ResourceBaseline* oResourceBaseline = new ResourceBaseline();
			oResourceBaseline->SetXML(oXML.ReadCollectionObject(lIndex));
			mp_oCollection->SetAddMode(TRUE);
			CString sKey = _T("");
			oResourceBaseline->mp_oCollection = mp_oCollection;
			mp_oCollection->m_Add(oResourceBaseline, sKey, MP_ADD_1, MP_ADD_2, FALSE, MP_ADD_3);
		}
	}
}

void ResourceBaseline_C::WriteObjectProtected(clsXML &oXML)
{
	LONG lIndex;
	ResourceBaseline* oResourceBaseline;
	for (lIndex = 1; lIndex <= GetCount(); lIndex++)
	{
		oResourceBaseline = (ResourceBaseline*) mp_oCollection->m_oReturnArrayElement(lIndex);
		oXML.WriteObject(oResourceBaseline->GetXML());
	}
}