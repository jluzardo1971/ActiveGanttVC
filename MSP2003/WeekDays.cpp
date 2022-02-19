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
#include "WeekDays.h"

IMPLEMENT_DYNCREATE(WeekDays, CCmdTarget)

//{616765DA-548B-487B-8234-CDE87540BF9E}
static const IID IID_IWeekDays = { 0x616765DA, 0x548B, 0x487B, { 0x82, 0x34, 0xCD, 0xE8, 0x75, 0x40, 0xBF, 0x9E} };

//{1E6127AB-FEF9-423C-9A41-7BD408EACEE1}
IMPLEMENT_OLECREATE_FLAGS(WeekDays, "MSP2003.WeekDays", afxRegApartmentThreading, 0x1E6127AB, 0xFEF9, 0x423C, 0x9A, 0x41, 0x7B, 0xD4, 0x08, 0xEA, 0xCE, 0xE1)

BEGIN_DISPATCH_MAP(WeekDays, CCmdTarget)
DISP_PROPERTY_EX_ID(WeekDays, "Count", 1, odl_GetCount, SetNotSupported, VT_I4)
DISP_PROPERTY_PARAM_ID(WeekDays, "Item", 2, odl_Item, SetNotSupported, VT_DISPATCH, VTS_BSTR)
DISP_FUNCTION_ID(WeekDays, "Add", 3, odl_Add, VT_DISPATCH, VTS_NONE)
DISP_FUNCTION_ID(WeekDays, "Clear", 4, odl_Clear, VT_EMPTY, VTS_NONE)
DISP_FUNCTION_ID(WeekDays, "Remove", 5, odl_Remove, VT_EMPTY, VTS_BSTR)
DISP_FUNCTION_ID(WeekDays, "IsNull", 6, IsNull, VT_BOOL, VTS_NONE)
DISP_FUNCTION_ID(WeekDays, "Initialize", 7, Initialize, VT_EMPTY, VTS_NONE)
DISP_FUNCTION_ID(WeekDays, "SetXML", 8, odl_SetXML, VT_EMPTY, VTS_BSTR)
DISP_FUNCTION_ID(WeekDays, "GetXML", 9, odl_GetXML, VT_BSTR, VTS_NONE)
DISP_DEFVALUE(WeekDays, "Item")
END_DISPATCH_MAP()

BEGIN_INTERFACE_MAP(WeekDays, CCmdTarget)
	INTERFACE_PART(WeekDays, IID_IWeekDays, Dispatch)
END_INTERFACE_MAP()

BEGIN_MESSAGE_MAP(WeekDays, CCmdTarget)
END_MESSAGE_MAP()


WeekDays::WeekDays()
{
	EnableAutomation();
	AfxOleLockApp();
	InitVars();
}

void WeekDays::Initialize(void)
{
	InitVars();
}

void WeekDays::InitVars(void)
{
	mp_oCollection = new clsCollectionBase("WeekDay");
}

WeekDays::~WeekDays()
{
	delete mp_oCollection;
	AfxOleUnlockApp();
}

void WeekDays::OnFinalRelease()
{
	CCmdTarget::OnFinalRelease();
}

LONG WeekDays::odl_GetCount(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetCount();
}

LONG WeekDays::GetCount(void)
{
	return mp_oCollection->m_lCount();
}

IDispatch* WeekDays::odl_Item(LPCTSTR Index)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return Item(Index)->GetIDispatch(TRUE);
}

IDispatch* WeekDays::odl_Add(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return Add()->GetIDispatch(TRUE);
}

void WeekDays::odl_Clear(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	Clear();
}

void WeekDays::odl_Remove(LPCTSTR Index)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	Remove(Index);
}

BOOL WeekDays::IsNull(void)
{
	BOOL bReturn = TRUE;
	if (GetCount() > 0)
	{
		bReturn = FALSE;
	}
	return bReturn;
}

WeekDay* WeekDays::Item(CString Index)
{
	WeekDay *oWeekDay;
	oWeekDay = (WeekDay*)mp_oCollection->m_oItem(Index, MP_ITEM_1, MP_ITEM_2, MP_ITEM_3, MP_ITEM_4);
	return oWeekDay;
}

WeekDay* WeekDays::Add()
{
	mp_oCollection->SetAddMode(TRUE);
	WeekDay* oWeekDay = new WeekDay();
	oWeekDay->mp_oCollection = mp_oCollection;
	mp_oCollection->m_Add(oWeekDay, _T(""), MP_ADD_1, MP_ADD_2, FALSE, MP_ADD_3);
	return oWeekDay;
}

void WeekDays::Clear(void)
{
	mp_oCollection->m_Clear();
}

void WeekDays::Remove(CString Index)
{
	mp_oCollection->m_Remove(Index, MP_REMOVE_1, MP_REMOVE_2, MP_REMOVE_3, MP_REMOVE_4);
}

BSTR WeekDays::odl_GetXML(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetXML().AllocSysString();
}

void WeekDays::odl_SetXML(LPCTSTR sXML)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetXML(sXML);
}

CString WeekDays::GetXML(void)
{
	if (IsNull() == TRUE)
	{
		return "<WeekDays/>";
	}
	LONG lIndex;
	WeekDay* oWeekDay;
	clsXML oXML("WeekDays");
	oXML.InitializeWriter();
	for (lIndex = 1; lIndex <= GetCount(); lIndex++)
	{
		oWeekDay = (WeekDay*) mp_oCollection->m_oReturnArrayElement(lIndex);
		oXML.WriteObject(oWeekDay->GetXML());
	}
	return oXML.GetXML();
}

void WeekDays::SetXML(CString sXML)
{
	LONG lIndex;
	clsXML oXML("WeekDays");
	oXML.SetSupportOptional(TRUE);
	oXML.SetXML(sXML);
	oXML.InitializeReader();
	mp_oCollection->m_Clear();
	for (lIndex = 1;lIndex <= oXML.ReadCollectionCount();lIndex++)
	{
		WeekDay* oWeekDay = new WeekDay();
		oWeekDay->SetXML(oXML.ReadCollectionObject(lIndex));
		mp_oCollection->SetAddMode(TRUE);
		CString sKey = _T("");
		oWeekDay->mp_oCollection = mp_oCollection;
		mp_oCollection->m_Add(oWeekDay, sKey, MP_ADD_1, MP_ADD_2, FALSE, MP_ADD_3);
	}
}