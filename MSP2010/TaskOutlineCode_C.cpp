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

//{C3DCA0E3-4AF0-4590-84F5-6F402E5F75E3}
static const IID IID_ITaskOutlineCode_C = { 0xC3DCA0E3, 0x4AF0, 0x4590, { 0x84, 0xF5, 0x6F, 0x40, 0x2E, 0x5F, 0x75, 0xE3} };

//{DB86DCA6-C0EC-4095-BD81-C228197DDCCC}
IMPLEMENT_OLECREATE_FLAGS(TaskOutlineCode_C, "MSP2010.TaskOutlineCode_C", afxRegApartmentThreading, 0xDB86DCA6, 0xC0EC, 0x4095, 0xBD, 0x81, 0xC2, 0x28, 0x19, 0x7D, 0xDC, 0xCC)

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