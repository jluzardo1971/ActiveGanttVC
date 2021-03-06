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
#include "Tasks.h"

IMPLEMENT_DYNCREATE(Tasks, CCmdTarget)

//{30671723-EABF-4C04-98E0-8A3F7D354D10}
static const IID IID_ITasks = { 0x30671723, 0xEABF, 0x4C04, { 0x98, 0xE0, 0x8A, 0x3F, 0x7D, 0x35, 0x4D, 0x10} };

//{C74BA3FB-83E9-4EFB-B6B8-BD012F80EC5F}
IMPLEMENT_OLECREATE_FLAGS(Tasks, "MSP2003.Tasks", afxRegApartmentThreading, 0xC74BA3FB, 0x83E9, 0x4EFB, 0xB6, 0xB8, 0xBD, 0x01, 0x2F, 0x80, 0xEC, 0x5F)

BEGIN_DISPATCH_MAP(Tasks, CCmdTarget)
DISP_PROPERTY_EX_ID(Tasks, "Count", 1, odl_GetCount, SetNotSupported, VT_I4)
DISP_PROPERTY_PARAM_ID(Tasks, "Item", 2, odl_Item, SetNotSupported, VT_DISPATCH, VTS_BSTR)
DISP_FUNCTION_ID(Tasks, "Add", 3, odl_Add, VT_DISPATCH, VTS_NONE)
DISP_FUNCTION_ID(Tasks, "Clear", 4, odl_Clear, VT_EMPTY, VTS_NONE)
DISP_FUNCTION_ID(Tasks, "Remove", 5, odl_Remove, VT_EMPTY, VTS_BSTR)
DISP_FUNCTION_ID(Tasks, "IsNull", 6, IsNull, VT_BOOL, VTS_NONE)
DISP_FUNCTION_ID(Tasks, "Initialize", 7, Initialize, VT_EMPTY, VTS_NONE)
DISP_FUNCTION_ID(Tasks, "SetXML", 8, odl_SetXML, VT_EMPTY, VTS_BSTR)
DISP_FUNCTION_ID(Tasks, "GetXML", 9, odl_GetXML, VT_BSTR, VTS_NONE)
DISP_DEFVALUE(Tasks, "Item")
END_DISPATCH_MAP()

BEGIN_INTERFACE_MAP(Tasks, CCmdTarget)
	INTERFACE_PART(Tasks, IID_ITasks, Dispatch)
END_INTERFACE_MAP()

BEGIN_MESSAGE_MAP(Tasks, CCmdTarget)
END_MESSAGE_MAP()


Tasks::Tasks()
{
	EnableAutomation();
	AfxOleLockApp();
	InitVars();
}

void Tasks::Initialize(void)
{
	InitVars();
}

void Tasks::InitVars(void)
{
	mp_oCollection = new clsCollectionBase("Task");
}

Tasks::~Tasks()
{
	delete mp_oCollection;
	AfxOleUnlockApp();
}

void Tasks::OnFinalRelease()
{
	CCmdTarget::OnFinalRelease();
}

LONG Tasks::odl_GetCount(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetCount();
}

LONG Tasks::GetCount(void)
{
	return mp_oCollection->m_lCount();
}

IDispatch* Tasks::odl_Item(LPCTSTR Index)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return Item(Index)->GetIDispatch(TRUE);
}

IDispatch* Tasks::odl_Add(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return Add()->GetIDispatch(TRUE);
}

void Tasks::odl_Clear(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	Clear();
}

void Tasks::odl_Remove(LPCTSTR Index)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	Remove(Index);
}

BOOL Tasks::IsNull(void)
{
	BOOL bReturn = TRUE;
	if (GetCount() > 0)
	{
		bReturn = FALSE;
	}
	return bReturn;
}

Task* Tasks::Item(CString Index)
{
	Task *oTask;
	oTask = (Task*)mp_oCollection->m_oItem(Index, MP_ITEM_1, MP_ITEM_2, MP_ITEM_3, MP_ITEM_4);
	return oTask;
}

Task* Tasks::Add()
{
	mp_oCollection->SetAddMode(TRUE);
	Task* oTask = new Task();
	oTask->mp_oCollection = mp_oCollection;
	mp_oCollection->m_Add(oTask, _T(""), MP_ADD_1, MP_ADD_2, FALSE, MP_ADD_3);
	return oTask;
}

void Tasks::Clear(void)
{
	mp_oCollection->m_Clear();
}

void Tasks::Remove(CString Index)
{
	mp_oCollection->m_Remove(Index, MP_REMOVE_1, MP_REMOVE_2, MP_REMOVE_3, MP_REMOVE_4);
}

BSTR Tasks::odl_GetXML(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetXML().AllocSysString();
}

void Tasks::odl_SetXML(LPCTSTR sXML)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetXML(sXML);
}

CString Tasks::GetXML(void)
{
	if (IsNull() == TRUE)
	{
		return "<Tasks/>";
	}
	LONG lIndex;
	Task* oTask;
	clsXML oXML("Tasks");
	oXML.InitializeWriter();
	for (lIndex = 1; lIndex <= GetCount(); lIndex++)
	{
		oTask = (Task*) mp_oCollection->m_oReturnArrayElement(lIndex);
		oXML.WriteObject(oTask->GetXML());
	}
	return oXML.GetXML();
}

void Tasks::SetXML(CString sXML)
{
	LONG lIndex;
	clsXML oXML("Tasks");
	oXML.SetSupportOptional(TRUE);
	oXML.SetXML(sXML);
	oXML.InitializeReader();
	mp_oCollection->m_Clear();
	for (lIndex = 1;lIndex <= oXML.ReadCollectionCount();lIndex++)
	{
		Task* oTask = new Task();
		oTask->SetXML(oXML.ReadCollectionObject(lIndex));
		mp_oCollection->SetAddMode(TRUE);
		CString sKey = _T("");
		sKey.Format(_T("%d"), oTask->GetUID());
		sKey = _T("K") + sKey;
		oTask->mp_oCollection = mp_oCollection;
		oTask->SetKey(sKey);
		mp_oCollection->m_Add(oTask, sKey, MP_ADD_1, MP_ADD_2, FALSE, MP_ADD_3);
	}
}