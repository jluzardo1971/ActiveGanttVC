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
#include "AssignmentExtendedAttribute_C.h"

IMPLEMENT_DYNCREATE(AssignmentExtendedAttribute_C, CCmdTarget)

//{C7DAF394-D4A7-4177-8110-07ECCB2B57DD}
static const IID IID_IAssignmentExtendedAttribute_C = { 0xC7DAF394, 0xD4A7, 0x4177, { 0x81, 0x10, 0x07, 0xEC, 0xCB, 0x2B, 0x57, 0xDD} };

//{6F6478A3-C44A-451A-9176-7BAC8E3D8553}
IMPLEMENT_OLECREATE_FLAGS(AssignmentExtendedAttribute_C, "MSP2010.AssignmentExtendedAttribute_C", afxRegApartmentThreading, 0x6F6478A3, 0xC44A, 0x451A, 0x91, 0x76, 0x7B, 0xAC, 0x8E, 0x3D, 0x85, 0x53)

BEGIN_DISPATCH_MAP(AssignmentExtendedAttribute_C, CCmdTarget)
DISP_PROPERTY_EX_ID(AssignmentExtendedAttribute_C, "Count", 1, odl_GetCount, SetNotSupported, VT_I4)
DISP_PROPERTY_PARAM_ID(AssignmentExtendedAttribute_C, "Item", 2, odl_Item, SetNotSupported, VT_DISPATCH, VTS_BSTR)
DISP_FUNCTION_ID(AssignmentExtendedAttribute_C, "Add", 3, odl_Add, VT_DISPATCH, VTS_NONE)
DISP_FUNCTION_ID(AssignmentExtendedAttribute_C, "Clear", 4, odl_Clear, VT_EMPTY, VTS_NONE)
DISP_FUNCTION_ID(AssignmentExtendedAttribute_C, "Remove", 5, odl_Remove, VT_EMPTY, VTS_BSTR)
DISP_FUNCTION_ID(AssignmentExtendedAttribute_C, "IsNull", 6, IsNull, VT_BOOL, VTS_NONE)
DISP_FUNCTION_ID(AssignmentExtendedAttribute_C, "Initialize", 7, Initialize, VT_EMPTY, VTS_NONE)
DISP_DEFVALUE(AssignmentExtendedAttribute_C, "Item")
END_DISPATCH_MAP()

BEGIN_INTERFACE_MAP(AssignmentExtendedAttribute_C, CCmdTarget)
	INTERFACE_PART(AssignmentExtendedAttribute_C, IID_IAssignmentExtendedAttribute_C, Dispatch)
END_INTERFACE_MAP()

BEGIN_MESSAGE_MAP(AssignmentExtendedAttribute_C, CCmdTarget)
END_MESSAGE_MAP()


AssignmentExtendedAttribute_C::AssignmentExtendedAttribute_C()
{
	EnableAutomation();
	AfxOleLockApp();
	InitVars();
}

void AssignmentExtendedAttribute_C::Initialize(void)
{
	InitVars();
}

void AssignmentExtendedAttribute_C::InitVars(void)
{
	mp_oCollection = new clsCollectionBase("AssignmentExtendedAttribute");
}

AssignmentExtendedAttribute_C::~AssignmentExtendedAttribute_C()
{
	delete mp_oCollection;
	AfxOleUnlockApp();
}

void AssignmentExtendedAttribute_C::OnFinalRelease()
{
	CCmdTarget::OnFinalRelease();
}

LONG AssignmentExtendedAttribute_C::odl_GetCount(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetCount();
}

LONG AssignmentExtendedAttribute_C::GetCount(void)
{
	return mp_oCollection->m_lCount();
}

IDispatch* AssignmentExtendedAttribute_C::odl_Item(LPCTSTR Index)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return Item(Index)->GetIDispatch(TRUE);
}

IDispatch* AssignmentExtendedAttribute_C::odl_Add(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return Add()->GetIDispatch(TRUE);
}

void AssignmentExtendedAttribute_C::odl_Clear(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	Clear();
}

void AssignmentExtendedAttribute_C::odl_Remove(LPCTSTR Index)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	Remove(Index);
}

BOOL AssignmentExtendedAttribute_C::IsNull(void)
{
	BOOL bReturn = TRUE;
	if (GetCount() > 0)
	{
		bReturn = FALSE;
	}
	return bReturn;
}

AssignmentExtendedAttribute* AssignmentExtendedAttribute_C::Item(CString Index)
{
	AssignmentExtendedAttribute *oAssignmentExtendedAttribute;
	oAssignmentExtendedAttribute = (AssignmentExtendedAttribute*)mp_oCollection->m_oItem(Index, MP_ITEM_1, MP_ITEM_2, MP_ITEM_3, MP_ITEM_4);
	return oAssignmentExtendedAttribute;
}

AssignmentExtendedAttribute* AssignmentExtendedAttribute_C::Add()
{
	mp_oCollection->SetAddMode(TRUE);
	AssignmentExtendedAttribute* oAssignmentExtendedAttribute = new AssignmentExtendedAttribute();
	oAssignmentExtendedAttribute->mp_oCollection = mp_oCollection;
	mp_oCollection->m_Add(oAssignmentExtendedAttribute, _T(""), MP_ADD_1, MP_ADD_2, FALSE, MP_ADD_3);
	return oAssignmentExtendedAttribute;
}

void AssignmentExtendedAttribute_C::Clear(void)
{
	mp_oCollection->m_Clear();
}

void AssignmentExtendedAttribute_C::Remove(CString Index)
{
	mp_oCollection->m_Remove(Index, MP_REMOVE_1, MP_REMOVE_2, MP_REMOVE_3, MP_REMOVE_4);
}

void AssignmentExtendedAttribute_C::ReadObjectProtected(clsXML &oXML)
{
	LONG lIndex;
	for (lIndex = 1; lIndex <= oXML.ReadCollectionCount(); lIndex++)
	{
		if (oXML.GetCollectionObjectName(lIndex) == "ExtendedAttribute")
		{
			AssignmentExtendedAttribute* oAssignmentExtendedAttribute = new AssignmentExtendedAttribute();
			oAssignmentExtendedAttribute->SetXML(oXML.ReadCollectionObject(lIndex));
			mp_oCollection->SetAddMode(TRUE);
			CString sKey = _T("");
			oAssignmentExtendedAttribute->mp_oCollection = mp_oCollection;
			mp_oCollection->m_Add(oAssignmentExtendedAttribute, sKey, MP_ADD_1, MP_ADD_2, FALSE, MP_ADD_3);
		}
	}
}

void AssignmentExtendedAttribute_C::WriteObjectProtected(clsXML &oXML)
{
	LONG lIndex;
	AssignmentExtendedAttribute* oAssignmentExtendedAttribute;
	for (lIndex = 1; lIndex <= GetCount(); lIndex++)
	{
		oAssignmentExtendedAttribute = (AssignmentExtendedAttribute*) mp_oCollection->m_oReturnArrayElement(lIndex);
		oXML.WriteObject(oAssignmentExtendedAttribute->GetXML());
	}
}