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
#include "TaskBaseline_C.h"

IMPLEMENT_DYNCREATE(TaskBaseline_C, CCmdTarget)

//{F525A3F4-72A2-4BE7-97F0-D29BC80163A8}
static const IID IID_ITaskBaseline_C = { 0xF525A3F4, 0x72A2, 0x4BE7, { 0x97, 0xF0, 0xD2, 0x9B, 0xC8, 0x01, 0x63, 0xA8} };

//{E8AF8B12-71BB-46D1-89AB-7E39D8DA2542}
IMPLEMENT_OLECREATE_FLAGS(TaskBaseline_C, "MSP2003.TaskBaseline_C", afxRegApartmentThreading, 0xE8AF8B12, 0x71BB, 0x46D1, 0x89, 0xAB, 0x7E, 0x39, 0xD8, 0xDA, 0x25, 0x42)

BEGIN_DISPATCH_MAP(TaskBaseline_C, CCmdTarget)
DISP_PROPERTY_EX_ID(TaskBaseline_C, "Count", 1, odl_GetCount, SetNotSupported, VT_I4)
DISP_PROPERTY_PARAM_ID(TaskBaseline_C, "Item", 2, odl_Item, SetNotSupported, VT_DISPATCH, VTS_BSTR)
DISP_FUNCTION_ID(TaskBaseline_C, "Add", 3, odl_Add, VT_DISPATCH, VTS_NONE)
DISP_FUNCTION_ID(TaskBaseline_C, "Clear", 4, odl_Clear, VT_EMPTY, VTS_NONE)
DISP_FUNCTION_ID(TaskBaseline_C, "Remove", 5, odl_Remove, VT_EMPTY, VTS_BSTR)
DISP_FUNCTION_ID(TaskBaseline_C, "IsNull", 6, IsNull, VT_BOOL, VTS_NONE)
DISP_FUNCTION_ID(TaskBaseline_C, "Initialize", 7, Initialize, VT_EMPTY, VTS_NONE)
DISP_DEFVALUE(TaskBaseline_C, "Item")
END_DISPATCH_MAP()

BEGIN_INTERFACE_MAP(TaskBaseline_C, CCmdTarget)
	INTERFACE_PART(TaskBaseline_C, IID_ITaskBaseline_C, Dispatch)
END_INTERFACE_MAP()

BEGIN_MESSAGE_MAP(TaskBaseline_C, CCmdTarget)
END_MESSAGE_MAP()


TaskBaseline_C::TaskBaseline_C()
{
	EnableAutomation();
	AfxOleLockApp();
	InitVars();
}

void TaskBaseline_C::Initialize(void)
{
	InitVars();
}

void TaskBaseline_C::InitVars(void)
{
	mp_oCollection = new clsCollectionBase("TaskBaseline");
}

TaskBaseline_C::~TaskBaseline_C()
{
	delete mp_oCollection;
	AfxOleUnlockApp();
}

void TaskBaseline_C::OnFinalRelease()
{
	CCmdTarget::OnFinalRelease();
}

LONG TaskBaseline_C::odl_GetCount(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetCount();
}

LONG TaskBaseline_C::GetCount(void)
{
	return mp_oCollection->m_lCount();
}

IDispatch* TaskBaseline_C::odl_Item(LPCTSTR Index)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return Item(Index)->GetIDispatch(TRUE);
}

IDispatch* TaskBaseline_C::odl_Add(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return Add()->GetIDispatch(TRUE);
}

void TaskBaseline_C::odl_Clear(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	Clear();
}

void TaskBaseline_C::odl_Remove(LPCTSTR Index)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	Remove(Index);
}

BOOL TaskBaseline_C::IsNull(void)
{
	BOOL bReturn = TRUE;
	if (GetCount() > 0)
	{
		bReturn = FALSE;
	}
	return bReturn;
}

TaskBaseline* TaskBaseline_C::Item(CString Index)
{
	TaskBaseline *oTaskBaseline;
	oTaskBaseline = (TaskBaseline*)mp_oCollection->m_oItem(Index, MP_ITEM_1, MP_ITEM_2, MP_ITEM_3, MP_ITEM_4);
	return oTaskBaseline;
}

TaskBaseline* TaskBaseline_C::Add()
{
	mp_oCollection->SetAddMode(TRUE);
	TaskBaseline* oTaskBaseline = new TaskBaseline();
	oTaskBaseline->mp_oCollection = mp_oCollection;
	mp_oCollection->m_Add(oTaskBaseline, _T(""), MP_ADD_1, MP_ADD_2, FALSE, MP_ADD_3);
	return oTaskBaseline;
}

void TaskBaseline_C::Clear(void)
{
	mp_oCollection->m_Clear();
}

void TaskBaseline_C::Remove(CString Index)
{
	mp_oCollection->m_Remove(Index, MP_REMOVE_1, MP_REMOVE_2, MP_REMOVE_3, MP_REMOVE_4);
}

void TaskBaseline_C::ReadObjectProtected(clsXML &oXML)
{
	LONG lIndex;
	for (lIndex = 1; lIndex <= oXML.ReadCollectionCount(); lIndex++)
	{
		if (oXML.GetCollectionObjectName(lIndex) == "Baseline")
		{
			TaskBaseline* oTaskBaseline = new TaskBaseline();
			oTaskBaseline->SetXML(oXML.ReadCollectionObject(lIndex));
			mp_oCollection->SetAddMode(TRUE);
			CString sKey = _T("");
			oTaskBaseline->mp_oCollection = mp_oCollection;
			mp_oCollection->m_Add(oTaskBaseline, sKey, MP_ADD_1, MP_ADD_2, FALSE, MP_ADD_3);
		}
	}
}

void TaskBaseline_C::WriteObjectProtected(clsXML &oXML)
{
	LONG lIndex;
	TaskBaseline* oTaskBaseline;
	for (lIndex = 1; lIndex <= GetCount(); lIndex++)
	{
		oTaskBaseline = (TaskBaseline*) mp_oCollection->m_oReturnArrayElement(lIndex);
		oXML.WriteObject(oTaskBaseline->GetXML());
	}
}