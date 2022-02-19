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
#include "TaskOutlineCode_C.h"

IMPLEMENT_DYNCREATE(TaskOutlineCode_C, CCmdTarget)

//{8C4FFF59-152F-4B98-9643-DDC31BDA2F08}
static const IID IID_ITaskOutlineCode_C = { 0x8C4FFF59, 0x152F, 0x4B98, { 0x96, 0x43, 0xDD, 0xC3, 0x1B, 0xDA, 0x2F, 0x08} };

//{14B025F4-3C04-4F33-8B17-78CB6AAED582}
IMPLEMENT_OLECREATE_FLAGS(TaskOutlineCode_C, "MSP2007.TaskOutlineCode_C", afxRegApartmentThreading, 0x14B025F4, 0x3C04, 0x4F33, 0x8B, 0x17, 0x78, 0xCB, 0x6A, 0xAE, 0xD5, 0x82)

BEGIN_DISPATCH_MAP(TaskOutlineCode_C, CCmdTarget)
DISP_PROPERTY_EX_ID(TaskOutlineCode_C, "Count", 1, odl_GetCount, SetNotSupported, VT_I4)
DISP_PROPERTY_PARAM_ID(TaskOutlineCode_C, "Item", 2, odl_Item, SetNotSupported, VT_DISPATCH, VTS_BSTR)
DISP_FUNCTION_ID(TaskOutlineCode_C, "Add", 3, odl_Add, VT_DISPATCH, VTS_NONE)
DISP_FUNCTION_ID(TaskOutlineCode_C, "Clear", 4, odl_Clear, VT_EMPTY, VTS_NONE)
DISP_FUNCTION_ID(TaskOutlineCode_C, "Remove", 5, odl_Remove, VT_EMPTY, VTS_BSTR)
DISP_FUNCTION_ID(TaskOutlineCode_C, "IsNull", 6, IsNull, VT_BOOL, VTS_NONE)
DISP_FUNCTION_ID(TaskOutlineCode_C, "Initialize", 7, Initialize, VT_EMPTY, VTS_NONE)
DISP_DEFVALUE(TaskOutlineCode_C, "Item")
END_DISPATCH_MAP()

BEGIN_INTERFACE_MAP(TaskOutlineCode_C, CCmdTarget)
	INTERFACE_PART(TaskOutlineCode_C, IID_ITaskOutlineCode_C, Dispatch)
END_INTERFACE_MAP()

BEGIN_MESSAGE_MAP(TaskOutlineCode_C, CCmdTarget)
END_MESSAGE_MAP()


TaskOutlineCode_C::TaskOutlineCode_C()
{
	EnableAutomation();
	AfxOleLockApp();
	InitVars();
}

void TaskOutlineCode_C::Initialize(void)
{
	InitVars();
}

void TaskOutlineCode_C::InitVars(void)
{
	mp_oCollection = new clsCollectionBase("TaskOutlineCode");
}

TaskOutlineCode_C::~TaskOutlineCode_C()
{
	delete mp_oCollection;
	AfxOleUnlockApp();
}

void TaskOutlineCode_C::OnFinalRelease()
{
	CCmdTarget::OnFinalRelease();
}

LONG TaskOutlineCode_C::odl_GetCount(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetCount();
}

LONG TaskOutlineCode_C::GetCount(void)
{
	return mp_oCollection->m_lCount();
}

IDispatch* TaskOutlineCode_C::odl_Item(LPCTSTR Index)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return Item(Index)->GetIDispatch(TRUE);
}

IDispatch* TaskOutlineCode_C::odl_Add(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return Add()->GetIDispatch(TRUE);
}

void TaskOutlineCode_C::odl_Clear(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	Clear();
}

void TaskOutlineCode_C::odl_Remove(LPCTSTR Index)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	Remove(Index);
}

BOOL TaskOutlineCode_C::IsNull(void)
{
	BOOL bReturn = TRUE;
	if (GetCount() > 0)
	{
		bReturn = FALSE;
	}
	return bReturn;
}

TaskOutlineCode* TaskOutlineCode_C::Item(CString Index)
{
	TaskOutlineCode *oTaskOutlineCode;
	oTaskOutlineCode = (TaskOutlineCode*)mp_oCollection->m_oItem(Index, MP_ITEM_1, MP_ITEM_2, MP_ITEM_3, MP_ITEM_4);
	return oTaskOutlineCode;
}

TaskOutlineCode* TaskOutlineCode_C::Add()
{
	mp_oCollection->SetAddMode(TRUE);
	TaskOutlineCode* oTaskOutlineCode = new TaskOutlineCode();
	oTaskOutlineCode->mp_oCollection = mp_oCollection;
	mp_oCollection->m_Add(oTaskOutlineCode, _T(""), MP_ADD_1, MP_ADD_2, FALSE, MP_ADD_3);
	return oTaskOutlineCode;
}

void TaskOutlineCode_C::Clear(void)
{
	mp_oCollection->m_Clear();
}

void TaskOutlineCode_C::Remove(CString Index)
{
	mp_oCollection->m_Remove(Index, MP_REMOVE_1, MP_REMOVE_2, MP_REMOVE_3, MP_REMOVE_4);
}

void TaskOutlineCode_C::ReadObjectProtected(clsXML &oXML)
{
	LONG lIndex;
	for (lIndex = 1; lIndex <= oXML.ReadCollectionCount(); lIndex++)
	{
		if (oXML.GetCollectionObjectName(lIndex) == "OutlineCode")
		{
			TaskOutlineCode* oTaskOutlineCode = new TaskOutlineCode();
			oTaskOutlineCode->SetXML(oXML.ReadCollectionObject(lIndex));
			mp_oCollection->SetAddMode(TRUE);
			CString sKey = _T("");
			oTaskOutlineCode->mp_oCollection = mp_oCollection;
			mp_oCollection->m_Add(oTaskOutlineCode, sKey, MP_ADD_1, MP_ADD_2, FALSE, MP_ADD_3);
		}
	}
}

void TaskOutlineCode_C::WriteObjectProtected(clsXML &oXML)
{
	LONG lIndex;
	TaskOutlineCode* oTaskOutlineCode;
	for (lIndex = 1; lIndex <= GetCount(); lIndex++)
	{
		oTaskOutlineCode = (TaskOutlineCode*) mp_oCollection->m_oReturnArrayElement(lIndex);
		oXML.WriteObject(oTaskOutlineCode->GetXML());
	}
}