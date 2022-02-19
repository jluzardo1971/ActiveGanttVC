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
#include "Assignments.h"

IMPLEMENT_DYNCREATE(Assignments, CCmdTarget)

//{09E5A086-F13B-4F03-BC24-66A2C6EF0E81}
static const IID IID_IAssignments = { 0x09E5A086, 0xF13B, 0x4F03, { 0xBC, 0x24, 0x66, 0xA2, 0xC6, 0xEF, 0x0E, 0x81} };

//{D0E9FAEB-39F2-40B7-A969-68E762C63392}
IMPLEMENT_OLECREATE_FLAGS(Assignments, "MSP2010.Assignments", afxRegApartmentThreading, 0xD0E9FAEB, 0x39F2, 0x40B7, 0xA9, 0x69, 0x68, 0xE7, 0x62, 0xC6, 0x33, 0x92)

BEGIN_DISPATCH_MAP(Assignments, CCmdTarget)
DISP_PROPERTY_EX_ID(Assignments, "Count", 1, odl_GetCount, SetNotSupported, VT_I4)
DISP_PROPERTY_PARAM_ID(Assignments, "Item", 2, odl_Item, SetNotSupported, VT_DISPATCH, VTS_BSTR)
DISP_FUNCTION_ID(Assignments, "Add", 3, odl_Add, VT_DISPATCH, VTS_NONE)
DISP_FUNCTION_ID(Assignments, "Clear", 4, odl_Clear, VT_EMPTY, VTS_NONE)
DISP_FUNCTION_ID(Assignments, "Remove", 5, odl_Remove, VT_EMPTY, VTS_BSTR)
DISP_FUNCTION_ID(Assignments, "IsNull", 6, IsNull, VT_BOOL, VTS_NONE)
DISP_FUNCTION_ID(Assignments, "Initialize", 7, Initialize, VT_EMPTY, VTS_NONE)
DISP_FUNCTION_ID(Assignments, "SetXML", 8, odl_SetXML, VT_EMPTY, VTS_BSTR)
DISP_FUNCTION_ID(Assignments, "GetXML", 9, odl_GetXML, VT_BSTR, VTS_NONE)
DISP_DEFVALUE(Assignments, "Item")
END_DISPATCH_MAP()

BEGIN_INTERFACE_MAP(Assignments, CCmdTarget)
	INTERFACE_PART(Assignments, IID_IAssignments, Dispatch)
END_INTERFACE_MAP()

BEGIN_MESSAGE_MAP(Assignments, CCmdTarget)
END_MESSAGE_MAP()


Assignments::Assignments()
{
	EnableAutomation();
	AfxOleLockApp();
	InitVars();
}

void Assignments::Initialize(void)
{
	InitVars();
}

void Assignments::InitVars(void)
{
	mp_oCollection = new clsCollectionBase("Assignment");
}

Assignments::~Assignments()
{
	delete mp_oCollection;
	AfxOleUnlockApp();
}

void Assignments::OnFinalRelease()
{
	CCmdTarget::OnFinalRelease();
}

LONG Assignments::odl_GetCount(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetCount();
}

LONG Assignments::GetCount(void)
{
	return mp_oCollection->m_lCount();
}

IDispatch* Assignments::odl_Item(LPCTSTR Index)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return Item(Index)->GetIDispatch(TRUE);
}

IDispatch* Assignments::odl_Add(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return Add()->GetIDispatch(TRUE);
}

void Assignments::odl_Clear(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	Clear();
}

void Assignments::odl_Remove(LPCTSTR Index)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	Remove(Index);
}

BOOL Assignments::IsNull(void)
{
	BOOL bReturn = TRUE;
	if (GetCount() > 0)
	{
		bReturn = FALSE;
	}
	return bReturn;
}

Assignment* Assignments::Item(CString Index)
{
	Assignment *oAssignment;
	oAssignment = (Assignment*)mp_oCollection->m_oItem(Index, MP_ITEM_1, MP_ITEM_2, MP_ITEM_3, MP_ITEM_4);
	return oAssignment;
}

Assignment* Assignments::Add()
{
	mp_oCollection->SetAddMode(TRUE);
	Assignment* oAssignment = new Assignment();
	oAssignment->mp_oCollection = mp_oCollection;
	mp_oCollection->m_Add(oAssignment, _T(""), MP_ADD_1, MP_ADD_2, FALSE, MP_ADD_3);
	return oAssignment;
}

void Assignments::Clear(void)
{
	mp_oCollection->m_Clear();
}

void Assignments::Remove(CString Index)
{
	mp_oCollection->m_Remove(Index, MP_REMOVE_1, MP_REMOVE_2, MP_REMOVE_3, MP_REMOVE_4);
}

BSTR Assignments::odl_GetXML(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetXML().AllocSysString();
}

void Assignments::odl_SetXML(LPCTSTR sXML)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetXML(sXML);
}

CString Assignments::GetXML(void)
{
	if (IsNull() == TRUE)
	{
		return "<Assignments/>";
	}
	LONG lIndex;
	Assignment* oAssignment;
	clsXML oXML("Assignments");
	oXML.InitializeWriter();
	for (lIndex = 1; lIndex <= GetCount(); lIndex++)
	{
		oAssignment = (Assignment*) mp_oCollection->m_oReturnArrayElement(lIndex);
		oXML.WriteObject(oAssignment->GetXML());
	}
	return oXML.GetXML();
}

void Assignments::SetXML(CString sXML)
{
	LONG lIndex;
	clsXML oXML("Assignments");
	oXML.SetSupportOptional(TRUE);
	oXML.SetXML(sXML);
	oXML.InitializeReader();
	mp_oCollection->m_Clear();
	for (lIndex = 1;lIndex <= oXML.ReadCollectionCount();lIndex++)
	{
		Assignment* oAssignment = new Assignment();
		oAssignment->SetXML(oXML.ReadCollectionObject(lIndex));
		mp_oCollection->SetAddMode(TRUE);
		CString sKey = _T("");
		sKey.Format(_T("%d"), oAssignment->GetUID());
		sKey = _T("K") + sKey;
		oAssignment->mp_oCollection = mp_oCollection;
		oAssignment->SetKey(sKey);
		mp_oCollection->m_Add(oAssignment, sKey, MP_ADD_1, MP_ADD_2, FALSE, MP_ADD_3);
	}
}