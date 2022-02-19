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
#include "AssignmentBaseline_C.h"

IMPLEMENT_DYNCREATE(AssignmentBaseline_C, CCmdTarget)

//{E0D9D24C-5837-47E9-9FFA-E62E2C7AAB0E}
static const IID IID_IAssignmentBaseline_C = { 0xE0D9D24C, 0x5837, 0x47E9, { 0x9F, 0xFA, 0xE6, 0x2E, 0x2C, 0x7A, 0xAB, 0x0E} };

//{5A2B366A-9177-49FD-8A0B-891D0E01FEEB}
IMPLEMENT_OLECREATE_FLAGS(AssignmentBaseline_C, "MSP2007.AssignmentBaseline_C", afxRegApartmentThreading, 0x5A2B366A, 0x9177, 0x49FD, 0x8A, 0x0B, 0x89, 0x1D, 0x0E, 0x01, 0xFE, 0xEB)

BEGIN_DISPATCH_MAP(AssignmentBaseline_C, CCmdTarget)
DISP_PROPERTY_EX_ID(AssignmentBaseline_C, "Count", 1, odl_GetCount, SetNotSupported, VT_I4)
DISP_PROPERTY_PARAM_ID(AssignmentBaseline_C, "Item", 2, odl_Item, SetNotSupported, VT_DISPATCH, VTS_BSTR)
DISP_FUNCTION_ID(AssignmentBaseline_C, "Add", 3, odl_Add, VT_DISPATCH, VTS_NONE)
DISP_FUNCTION_ID(AssignmentBaseline_C, "Clear", 4, odl_Clear, VT_EMPTY, VTS_NONE)
DISP_FUNCTION_ID(AssignmentBaseline_C, "Remove", 5, odl_Remove, VT_EMPTY, VTS_BSTR)
DISP_FUNCTION_ID(AssignmentBaseline_C, "IsNull", 6, IsNull, VT_BOOL, VTS_NONE)
DISP_FUNCTION_ID(AssignmentBaseline_C, "Initialize", 7, Initialize, VT_EMPTY, VTS_NONE)
DISP_DEFVALUE(AssignmentBaseline_C, "Item")
END_DISPATCH_MAP()

BEGIN_INTERFACE_MAP(AssignmentBaseline_C, CCmdTarget)
	INTERFACE_PART(AssignmentBaseline_C, IID_IAssignmentBaseline_C, Dispatch)
END_INTERFACE_MAP()

BEGIN_MESSAGE_MAP(AssignmentBaseline_C, CCmdTarget)
END_MESSAGE_MAP()


AssignmentBaseline_C::AssignmentBaseline_C()
{
	EnableAutomation();
	AfxOleLockApp();
	InitVars();
}

void AssignmentBaseline_C::Initialize(void)
{
	InitVars();
}

void AssignmentBaseline_C::InitVars(void)
{
	mp_oCollection = new clsCollectionBase("AssignmentBaseline");
}

AssignmentBaseline_C::~AssignmentBaseline_C()
{
	delete mp_oCollection;
	AfxOleUnlockApp();
}

void AssignmentBaseline_C::OnFinalRelease()
{
	CCmdTarget::OnFinalRelease();
}

LONG AssignmentBaseline_C::odl_GetCount(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetCount();
}

LONG AssignmentBaseline_C::GetCount(void)
{
	return mp_oCollection->m_lCount();
}

IDispatch* AssignmentBaseline_C::odl_Item(LPCTSTR Index)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return Item(Index)->GetIDispatch(TRUE);
}

IDispatch* AssignmentBaseline_C::odl_Add(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return Add()->GetIDispatch(TRUE);
}

void AssignmentBaseline_C::odl_Clear(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	Clear();
}

void AssignmentBaseline_C::odl_Remove(LPCTSTR Index)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	Remove(Index);
}

BOOL AssignmentBaseline_C::IsNull(void)
{
	BOOL bReturn = TRUE;
	if (GetCount() > 0)
	{
		bReturn = FALSE;
	}
	return bReturn;
}

AssignmentBaseline* AssignmentBaseline_C::Item(CString Index)
{
	AssignmentBaseline *oAssignmentBaseline;
	oAssignmentBaseline = (AssignmentBaseline*)mp_oCollection->m_oItem(Index, MP_ITEM_1, MP_ITEM_2, MP_ITEM_3, MP_ITEM_4);
	return oAssignmentBaseline;
}

AssignmentBaseline* AssignmentBaseline_C::Add()
{
	mp_oCollection->SetAddMode(TRUE);
	AssignmentBaseline* oAssignmentBaseline = new AssignmentBaseline();
	oAssignmentBaseline->mp_oCollection = mp_oCollection;
	mp_oCollection->m_Add(oAssignmentBaseline, _T(""), MP_ADD_1, MP_ADD_2, FALSE, MP_ADD_3);
	return oAssignmentBaseline;
}

void AssignmentBaseline_C::Clear(void)
{
	mp_oCollection->m_Clear();
}

void AssignmentBaseline_C::Remove(CString Index)
{
	mp_oCollection->m_Remove(Index, MP_REMOVE_1, MP_REMOVE_2, MP_REMOVE_3, MP_REMOVE_4);
}

void AssignmentBaseline_C::ReadObjectProtected(clsXML &oXML)
{
	LONG lIndex;
	for (lIndex = 1; lIndex <= oXML.ReadCollectionCount(); lIndex++)
	{
		if (oXML.GetCollectionObjectName(lIndex) == "Baseline")
		{
			AssignmentBaseline* oAssignmentBaseline = new AssignmentBaseline();
			oAssignmentBaseline->SetXML(oXML.ReadCollectionObject(lIndex));
			mp_oCollection->SetAddMode(TRUE);
			CString sKey = _T("");
			oAssignmentBaseline->mp_oCollection = mp_oCollection;
			mp_oCollection->m_Add(oAssignmentBaseline, sKey, MP_ADD_1, MP_ADD_2, FALSE, MP_ADD_3);
		}
	}
}

void AssignmentBaseline_C::WriteObjectProtected(clsXML &oXML)
{
	LONG lIndex;
	AssignmentBaseline* oAssignmentBaseline;
	for (lIndex = 1; lIndex <= GetCount(); lIndex++)
	{
		oAssignmentBaseline = (AssignmentBaseline*) mp_oCollection->m_oReturnArrayElement(lIndex);
		oXML.WriteObject(oAssignmentBaseline->GetXML());
	}
}