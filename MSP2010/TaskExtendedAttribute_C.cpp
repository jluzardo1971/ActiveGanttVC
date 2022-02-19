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
#include "TaskExtendedAttribute_C.h"

IMPLEMENT_DYNCREATE(TaskExtendedAttribute_C, CCmdTarget)

//{DAD8F1E1-B7B0-42D2-A346-CB67C398ADCC}
static const IID IID_ITaskExtendedAttribute_C = { 0xDAD8F1E1, 0xB7B0, 0x42D2, { 0xA3, 0x46, 0xCB, 0x67, 0xC3, 0x98, 0xAD, 0xCC} };

//{9CB31D81-B93A-42A7-89F5-078FE5F34DF9}
IMPLEMENT_OLECREATE_FLAGS(TaskExtendedAttribute_C, "MSP2010.TaskExtendedAttribute_C", afxRegApartmentThreading, 0x9CB31D81, 0xB93A, 0x42A7, 0x89, 0xF5, 0x07, 0x8F, 0xE5, 0xF3, 0x4D, 0xF9)

BEGIN_DISPATCH_MAP(TaskExtendedAttribute_C, CCmdTarget)
DISP_PROPERTY_EX_ID(TaskExtendedAttribute_C, "Count", 1, odl_GetCount, SetNotSupported, VT_I4)
DISP_PROPERTY_PARAM_ID(TaskExtendedAttribute_C, "Item", 2, odl_Item, SetNotSupported, VT_DISPATCH, VTS_BSTR)
DISP_FUNCTION_ID(TaskExtendedAttribute_C, "Add", 3, odl_Add, VT_DISPATCH, VTS_NONE)
DISP_FUNCTION_ID(TaskExtendedAttribute_C, "Clear", 4, odl_Clear, VT_EMPTY, VTS_NONE)
DISP_FUNCTION_ID(TaskExtendedAttribute_C, "Remove", 5, odl_Remove, VT_EMPTY, VTS_BSTR)
DISP_FUNCTION_ID(TaskExtendedAttribute_C, "IsNull", 6, IsNull, VT_BOOL, VTS_NONE)
DISP_FUNCTION_ID(TaskExtendedAttribute_C, "Initialize", 7, Initialize, VT_EMPTY, VTS_NONE)
DISP_DEFVALUE(TaskExtendedAttribute_C, "Item")
END_DISPATCH_MAP()

BEGIN_INTERFACE_MAP(TaskExtendedAttribute_C, CCmdTarget)
	INTERFACE_PART(TaskExtendedAttribute_C, IID_ITaskExtendedAttribute_C, Dispatch)
END_INTERFACE_MAP()

BEGIN_MESSAGE_MAP(TaskExtendedAttribute_C, CCmdTarget)
END_MESSAGE_MAP()


TaskExtendedAttribute_C::TaskExtendedAttribute_C()
{
	EnableAutomation();
	AfxOleLockApp();
	InitVars();
}

void TaskExtendedAttribute_C::Initialize(void)
{
	InitVars();
}

void TaskExtendedAttribute_C::InitVars(void)
{
	mp_oCollection = new clsCollectionBase("TaskExtendedAttribute");
}

TaskExtendedAttribute_C::~TaskExtendedAttribute_C()
{
	delete mp_oCollection;
	AfxOleUnlockApp();
}

void TaskExtendedAttribute_C::OnFinalRelease()
{
	CCmdTarget::OnFinalRelease();
}

LONG TaskExtendedAttribute_C::odl_GetCount(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetCount();
}

LONG TaskExtendedAttribute_C::GetCount(void)
{
	return mp_oCollection->m_lCount();
}

IDispatch* TaskExtendedAttribute_C::odl_Item(LPCTSTR Index)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return Item(Index)->GetIDispatch(TRUE);
}

IDispatch* TaskExtendedAttribute_C::odl_Add(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return Add()->GetIDispatch(TRUE);
}

void TaskExtendedAttribute_C::odl_Clear(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	Clear();
}

void TaskExtendedAttribute_C::odl_Remove(LPCTSTR Index)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	Remove(Index);
}

BOOL TaskExtendedAttribute_C::IsNull(void)
{
	BOOL bReturn = TRUE;
	if (GetCount() > 0)
	{
		bReturn = FALSE;
	}
	return bReturn;
}

TaskExtendedAttribute* TaskExtendedAttribute_C::Item(CString Index)
{
	TaskExtendedAttribute *oTaskExtendedAttribute;
	oTaskExtendedAttribute = (TaskExtendedAttribute*)mp_oCollection->m_oItem(Index, MP_ITEM_1, MP_ITEM_2, MP_ITEM_3, MP_ITEM_4);
	return oTaskExtendedAttribute;
}

TaskExtendedAttribute* TaskExtendedAttribute_C::Add()
{
	mp_oCollection->SetAddMode(TRUE);
	TaskExtendedAttribute* oTaskExtendedAttribute = new TaskExtendedAttribute();
	oTaskExtendedAttribute->mp_oCollection = mp_oCollection;
	mp_oCollection->m_Add(oTaskExtendedAttribute, _T(""), MP_ADD_1, MP_ADD_2, FALSE, MP_ADD_3);
	return oTaskExtendedAttribute;
}

void TaskExtendedAttribute_C::Clear(void)
{
	mp_oCollection->m_Clear();
}

void TaskExtendedAttribute_C::Remove(CString Index)
{
	mp_oCollection->m_Remove(Index, MP_REMOVE_1, MP_REMOVE_2, MP_REMOVE_3, MP_REMOVE_4);
}

void TaskExtendedAttribute_C::ReadObjectProtected(clsXML &oXML)
{
	LONG lIndex;
	for (lIndex = 1; lIndex <= oXML.ReadCollectionCount(); lIndex++)
	{
		if (oXML.GetCollectionObjectName(lIndex) == "ExtendedAttribute")
		{
			TaskExtendedAttribute* oTaskExtendedAttribute = new TaskExtendedAttribute();
			oTaskExtendedAttribute->SetXML(oXML.ReadCollectionObject(lIndex));
			mp_oCollection->SetAddMode(TRUE);
			CString sKey = _T("");
			oTaskExtendedAttribute->mp_oCollection = mp_oCollection;
			mp_oCollection->m_Add(oTaskExtendedAttribute, sKey, MP_ADD_1, MP_ADD_2, FALSE, MP_ADD_3);
		}
	}
}

void TaskExtendedAttribute_C::WriteObjectProtected(clsXML &oXML)
{
	LONG lIndex;
	TaskExtendedAttribute* oTaskExtendedAttribute;
	for (lIndex = 1; lIndex <= GetCount(); lIndex++)
	{
		oTaskExtendedAttribute = (TaskExtendedAttribute*) mp_oCollection->m_oReturnArrayElement(lIndex);
		oXML.WriteObject(oTaskExtendedAttribute->GetXML());
	}
}