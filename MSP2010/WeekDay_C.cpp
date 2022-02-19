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
#include "WeekDay_C.h"

IMPLEMENT_DYNCREATE(WeekDay_C, CCmdTarget)

//{333E76C6-C412-4357-B667-7EBFE3A948F3}
static const IID IID_IWeekDay_C = { 0x333E76C6, 0xC412, 0x4357, { 0xB6, 0x67, 0x7E, 0xBF, 0xE3, 0xA9, 0x48, 0xF3} };

//{65DB9CEA-1CCB-4658-B796-6AAC44223F72}
IMPLEMENT_OLECREATE_FLAGS(WeekDay_C, "MSP2010.WeekDay_C", afxRegApartmentThreading, 0x65DB9CEA, 0x1CCB, 0x4658, 0xB7, 0x96, 0x6A, 0xAC, 0x44, 0x22, 0x3F, 0x72)

BEGIN_DISPATCH_MAP(WeekDay_C, CCmdTarget)
DISP_PROPERTY_EX_ID(WeekDay_C, "Count", 1, odl_GetCount, SetNotSupported, VT_I4)
DISP_PROPERTY_PARAM_ID(WeekDay_C, "Item", 2, odl_Item, SetNotSupported, VT_DISPATCH, VTS_BSTR)
DISP_FUNCTION_ID(WeekDay_C, "Add", 3, odl_Add, VT_DISPATCH, VTS_NONE)
DISP_FUNCTION_ID(WeekDay_C, "Clear", 4, odl_Clear, VT_EMPTY, VTS_NONE)
DISP_FUNCTION_ID(WeekDay_C, "Remove", 5, odl_Remove, VT_EMPTY, VTS_BSTR)
DISP_FUNCTION_ID(WeekDay_C, "IsNull", 6, IsNull, VT_BOOL, VTS_NONE)
DISP_FUNCTION_ID(WeekDay_C, "Initialize", 7, Initialize, VT_EMPTY, VTS_NONE)
DISP_DEFVALUE(WeekDay_C, "Item")
END_DISPATCH_MAP()

BEGIN_INTERFACE_MAP(WeekDay_C, CCmdTarget)
	INTERFACE_PART(WeekDay_C, IID_IWeekDay_C, Dispatch)
END_INTERFACE_MAP()

BEGIN_MESSAGE_MAP(WeekDay_C, CCmdTarget)
END_MESSAGE_MAP()


WeekDay_C::WeekDay_C()
{
	EnableAutomation();
	AfxOleLockApp();
	InitVars();
}

void WeekDay_C::Initialize(void)
{
	InitVars();
}

void WeekDay_C::InitVars(void)
{
	mp_oCollection = new clsCollectionBase("WeekDay");
}

WeekDay_C::~WeekDay_C()
{
	delete mp_oCollection;
	AfxOleUnlockApp();
}

void WeekDay_C::OnFinalRelease()
{
	CCmdTarget::OnFinalRelease();
}

LONG WeekDay_C::odl_GetCount(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetCount();
}

LONG WeekDay_C::GetCount(void)
{
	return mp_oCollection->m_lCount();
}

IDispatch* WeekDay_C::odl_Item(LPCTSTR Index)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return Item(Index)->GetIDispatch(TRUE);
}

IDispatch* WeekDay_C::odl_Add(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return Add()->GetIDispatch(TRUE);
}

void WeekDay_C::odl_Clear(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	Clear();
}

void WeekDay_C::odl_Remove(LPCTSTR Index)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	Remove(Index);
}

BOOL WeekDay_C::IsNull(void)
{
	BOOL bReturn = TRUE;
	if (GetCount() > 0)
	{
		bReturn = FALSE;
	}
	return bReturn;
}

WeekDay* WeekDay_C::Item(CString Index)
{
	WeekDay *oWeekDay;
	oWeekDay = (WeekDay*)mp_oCollection->m_oItem(Index, MP_ITEM_1, MP_ITEM_2, MP_ITEM_3, MP_ITEM_4);
	return oWeekDay;
}

WeekDay* WeekDay_C::Add()
{
	mp_oCollection->SetAddMode(TRUE);
	WeekDay* oWeekDay = new WeekDay();
	oWeekDay->mp_oCollection = mp_oCollection;
	mp_oCollection->m_Add(oWeekDay, _T(""), MP_ADD_1, MP_ADD_2, FALSE, MP_ADD_3);
	return oWeekDay;
}

void WeekDay_C::Clear(void)
{
	mp_oCollection->m_Clear();
}

void WeekDay_C::Remove(CString Index)
{
	mp_oCollection->m_Remove(Index, MP_REMOVE_1, MP_REMOVE_2, MP_REMOVE_3, MP_REMOVE_4);
}

void WeekDay_C::ReadObjectProtected(clsXML &oXML)
{
	LONG lIndex;
	for (lIndex = 1; lIndex <= oXML.ReadCollectionCount(); lIndex++)
	{
		if (oXML.GetCollectionObjectName(lIndex) == "WeekDay")
		{
			WeekDay* oWeekDay = new WeekDay();
			oWeekDay->SetXML(oXML.ReadCollectionObject(lIndex));
			mp_oCollection->SetAddMode(TRUE);
			CString sKey = _T("");
			oWeekDay->mp_oCollection = mp_oCollection;
			mp_oCollection->m_Add(oWeekDay, sKey, MP_ADD_1, MP_ADD_2, FALSE, MP_ADD_3);
		}
	}
}

void WeekDay_C::WriteObjectProtected(clsXML &oXML)
{
	LONG lIndex;
	WeekDay* oWeekDay;
	for (lIndex = 1; lIndex <= GetCount(); lIndex++)
	{
		oWeekDay = (WeekDay*) mp_oCollection->m_oReturnArrayElement(lIndex);
		oXML.WriteObject(oWeekDay->GetXML());
	}
}