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
#include "TaskPredecessorLink_C.h"

IMPLEMENT_DYNCREATE(TaskPredecessorLink_C, CCmdTarget)

//{6151B4AD-9270-4FAD-AB83-3F2B0D7B3364}
static const IID IID_ITaskPredecessorLink_C = { 0x6151B4AD, 0x9270, 0x4FAD, { 0xAB, 0x83, 0x3F, 0x2B, 0x0D, 0x7B, 0x33, 0x64} };

//{DF07AB43-CFF2-4AA4-84C6-31512D28F1EC}
IMPLEMENT_OLECREATE_FLAGS(TaskPredecessorLink_C, "MSP2010.TaskPredecessorLink_C", afxRegApartmentThreading, 0xDF07AB43, 0xCFF2, 0x4AA4, 0x84, 0xC6, 0x31, 0x51, 0x2D, 0x28, 0xF1, 0xEC)

BEGIN_DISPATCH_MAP(TaskPredecessorLink_C, CCmdTarget)
DISP_PROPERTY_EX_ID(TaskPredecessorLink_C, "Count", 1, odl_GetCount, SetNotSupported, VT_I4)
DISP_PROPERTY_PARAM_ID(TaskPredecessorLink_C, "Item", 2, odl_Item, SetNotSupported, VT_DISPATCH, VTS_BSTR)
DISP_FUNCTION_ID(TaskPredecessorLink_C, "Add", 3, odl_Add, VT_DISPATCH, VTS_NONE)
DISP_FUNCTION_ID(TaskPredecessorLink_C, "Clear", 4, odl_Clear, VT_EMPTY, VTS_NONE)
DISP_FUNCTION_ID(TaskPredecessorLink_C, "Remove", 5, odl_Remove, VT_EMPTY, VTS_BSTR)
DISP_FUNCTION_ID(TaskPredecessorLink_C, "IsNull", 6, IsNull, VT_BOOL, VTS_NONE)
DISP_FUNCTION_ID(TaskPredecessorLink_C, "Initialize", 7, Initialize, VT_EMPTY, VTS_NONE)
DISP_DEFVALUE(TaskPredecessorLink_C, "Item")
END_DISPATCH_MAP()

BEGIN_INTERFACE_MAP(TaskPredecessorLink_C, CCmdTarget)
	INTERFACE_PART(TaskPredecessorLink_C, IID_ITaskPredecessorLink_C, Dispatch)
END_INTERFACE_MAP()

BEGIN_MESSAGE_MAP(TaskPredecessorLink_C, CCmdTarget)
END_MESSAGE_MAP()


TaskPredecessorLink_C::TaskPredecessorLink_C()
{
	EnableAutomation();
	AfxOleLockApp();
	InitVars();
}

void TaskPredecessorLink_C::Initialize(void)
{
	InitVars();
}

void TaskPredecessorLink_C::InitVars(void)
{
	mp_oCollection = new clsCollectionBase("TaskPredecessorLink");
}

TaskPredecessorLink_C::~TaskPredecessorLink_C()
{
	delete mp_oCollection;
	AfxOleUnlockApp();
}

void TaskPredecessorLink_C::OnFinalRelease()
{
	CCmdTarget::OnFinalRelease();
}

LONG TaskPredecessorLink_C::odl_GetCount(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetCount();
}

LONG TaskPredecessorLink_C::GetCount(void)
{
	return mp_oCollection->m_lCount();
}

IDispatch* TaskPredecessorLink_C::odl_Item(LPCTSTR Index)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return Item(Index)->GetIDispatch(TRUE);
}

IDispatch* TaskPredecessorLink_C::odl_Add(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return Add()->GetIDispatch(TRUE);
}

void TaskPredecessorLink_C::odl_Clear(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	Clear();
}

void TaskPredecessorLink_C::odl_Remove(LPCTSTR Index)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	Remove(Index);
}

BOOL TaskPredecessorLink_C::IsNull(void)
{
	BOOL bReturn = TRUE;
	if (GetCount() > 0)
	{
		bReturn = FALSE;
	}
	return bReturn;
}

TaskPredecessorLink* TaskPredecessorLink_C::Item(CString Index)
{
	TaskPredecessorLink *oTaskPredecessorLink;
	oTaskPredecessorLink = (TaskPredecessorLink*)mp_oCollection->m_oItem(Index, MP_ITEM_1, MP_ITEM_2, MP_ITEM_3, MP_ITEM_4);
	return oTaskPredecessorLink;
}

TaskPredecessorLink* TaskPredecessorLink_C::Add()
{
	mp_oCollection->SetAddMode(TRUE);
	TaskPredecessorLink* oTaskPredecessorLink = new TaskPredecessorLink();
	oTaskPredecessorLink->mp_oCollection = mp_oCollection;
	mp_oCollection->m_Add(oTaskPredecessorLink, _T(""), MP_ADD_1, MP_ADD_2, FALSE, MP_ADD_3);
	return oTaskPredecessorLink;
}

void TaskPredecessorLink_C::Clear(void)
{
	mp_oCollection->m_Clear();
}

void TaskPredecessorLink_C::Remove(CString Index)
{
	mp_oCollection->m_Remove(Index, MP_REMOVE_1, MP_REMOVE_2, MP_REMOVE_3, MP_REMOVE_4);
}

void TaskPredecessorLink_C::ReadObjectProtected(clsXML &oXML)
{
	LONG lIndex;
	for (lIndex = 1; lIndex <= oXML.ReadCollectionCount(); lIndex++)
	{
		if (oXML.GetCollectionObjectName(lIndex) == "PredecessorLink")
		{
			TaskPredecessorLink* oTaskPredecessorLink = new TaskPredecessorLink();
			oTaskPredecessorLink->SetXML(oXML.ReadCollectionObject(lIndex));
			mp_oCollection->SetAddMode(TRUE);
			CString sKey = _T("");
			oTaskPredecessorLink->mp_oCollection = mp_oCollection;
			mp_oCollection->m_Add(oTaskPredecessorLink, sKey, MP_ADD_1, MP_ADD_2, FALSE, MP_ADD_3);
		}
	}
}

void TaskPredecessorLink_C::WriteObjectProtected(clsXML &oXML)
{
	LONG lIndex;
	TaskPredecessorLink* oTaskPredecessorLink;
	for (lIndex = 1; lIndex <= GetCount(); lIndex++)
	{
		oTaskPredecessorLink = (TaskPredecessorLink*) mp_oCollection->m_oReturnArrayElement(lIndex);
		oXML.WriteObject(oTaskPredecessorLink->GetXML());
	}
}