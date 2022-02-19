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
#include "CalendarWeekDays.h"

IMPLEMENT_DYNCREATE(CalendarWeekDays, CCmdTarget)

//{DBF3BD5B-347E-4735-A2AC-0D791B34E53A}
static const IID IID_ICalendarWeekDays = { 0xDBF3BD5B, 0x347E, 0x4735, { 0xA2, 0xAC, 0x0D, 0x79, 0x1B, 0x34, 0xE5, 0x3A} };

//{10EDAC97-46FA-4E8C-96AA-75F11D2EF4AC}
IMPLEMENT_OLECREATE_FLAGS(CalendarWeekDays, "MSP2010.CalendarWeekDays", afxRegApartmentThreading, 0x10EDAC97, 0x46FA, 0x4E8C, 0x96, 0xAA, 0x75, 0xF1, 0x1D, 0x2E, 0xF4, 0xAC)

BEGIN_DISPATCH_MAP(CalendarWeekDays, CCmdTarget)
DISP_PROPERTY_EX_ID(CalendarWeekDays, "Count", 1, odl_GetCount, SetNotSupported, VT_I4)
DISP_PROPERTY_PARAM_ID(CalendarWeekDays, "Item", 2, odl_Item, SetNotSupported, VT_DISPATCH, VTS_BSTR)
DISP_FUNCTION_ID(CalendarWeekDays, "Add", 3, odl_Add, VT_DISPATCH, VTS_NONE)
DISP_FUNCTION_ID(CalendarWeekDays, "Clear", 4, odl_Clear, VT_EMPTY, VTS_NONE)
DISP_FUNCTION_ID(CalendarWeekDays, "Remove", 5, odl_Remove, VT_EMPTY, VTS_BSTR)
DISP_FUNCTION_ID(CalendarWeekDays, "IsNull", 6, IsNull, VT_BOOL, VTS_NONE)
DISP_FUNCTION_ID(CalendarWeekDays, "Initialize", 7, Initialize, VT_EMPTY, VTS_NONE)
DISP_FUNCTION_ID(CalendarWeekDays, "SetXML", 8, odl_SetXML, VT_EMPTY, VTS_BSTR)
DISP_FUNCTION_ID(CalendarWeekDays, "GetXML", 9, odl_GetXML, VT_BSTR, VTS_NONE)
DISP_DEFVALUE(CalendarWeekDays, "Item")
END_DISPATCH_MAP()

BEGIN_INTERFACE_MAP(CalendarWeekDays, CCmdTarget)
	INTERFACE_PART(CalendarWeekDays, IID_ICalendarWeekDays, Dispatch)
END_INTERFACE_MAP()

BEGIN_MESSAGE_MAP(CalendarWeekDays, CCmdTarget)
END_MESSAGE_MAP()


CalendarWeekDays::CalendarWeekDays()
{
	EnableAutomation();
	AfxOleLockApp();
	InitVars();
}

void CalendarWeekDays::Initialize(void)
{
	InitVars();
}

void CalendarWeekDays::InitVars(void)
{
	mp_oCollection = new clsCollectionBase("CalendarWeekDay");
}

CalendarWeekDays::~CalendarWeekDays()
{
	delete mp_oCollection;
	AfxOleUnlockApp();
}

void CalendarWeekDays::OnFinalRelease()
{
	CCmdTarget::OnFinalRelease();
}

LONG CalendarWeekDays::odl_GetCount(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetCount();
}

LONG CalendarWeekDays::GetCount(void)
{
	return mp_oCollection->m_lCount();
}

IDispatch* CalendarWeekDays::odl_Item(LPCTSTR Index)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return Item(Index)->GetIDispatch(TRUE);
}

IDispatch* CalendarWeekDays::odl_Add(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return Add()->GetIDispatch(TRUE);
}

void CalendarWeekDays::odl_Clear(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	Clear();
}

void CalendarWeekDays::odl_Remove(LPCTSTR Index)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	Remove(Index);
}

BOOL CalendarWeekDays::IsNull(void)
{
	BOOL bReturn = TRUE;
	if (GetCount() > 0)
	{
		bReturn = FALSE;
	}
	return bReturn;
}

CalendarWeekDay* CalendarWeekDays::Item(CString Index)
{
	CalendarWeekDay *oCalendarWeekDay;
	oCalendarWeekDay = (CalendarWeekDay*)mp_oCollection->m_oItem(Index, MP_ITEM_1, MP_ITEM_2, MP_ITEM_3, MP_ITEM_4);
	return oCalendarWeekDay;
}

CalendarWeekDay* CalendarWeekDays::Add()
{
	mp_oCollection->SetAddMode(TRUE);
	CalendarWeekDay* oCalendarWeekDay = new CalendarWeekDay();
	oCalendarWeekDay->mp_oCollection = mp_oCollection;
	mp_oCollection->m_Add(oCalendarWeekDay, _T(""), MP_ADD_1, MP_ADD_2, FALSE, MP_ADD_3);
	return oCalendarWeekDay;
}

void CalendarWeekDays::Clear(void)
{
	mp_oCollection->m_Clear();
}

void CalendarWeekDays::Remove(CString Index)
{
	mp_oCollection->m_Remove(Index, MP_REMOVE_1, MP_REMOVE_2, MP_REMOVE_3, MP_REMOVE_4);
}

BSTR CalendarWeekDays::odl_GetXML(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetXML().AllocSysString();
}

void CalendarWeekDays::odl_SetXML(LPCTSTR sXML)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetXML(sXML);
}

CString CalendarWeekDays::GetXML(void)
{
	if (IsNull() == TRUE)
	{
		return "<WeekDays/>";
	}
	LONG lIndex;
	CalendarWeekDay* oCalendarWeekDay;
	clsXML oXML("WeekDays");
	oXML.InitializeWriter();
	for (lIndex = 1; lIndex <= GetCount(); lIndex++)
	{
		oCalendarWeekDay = (CalendarWeekDay*) mp_oCollection->m_oReturnArrayElement(lIndex);
		oXML.WriteObject(oCalendarWeekDay->GetXML());
	}
	return oXML.GetXML();
}

void CalendarWeekDays::SetXML(CString sXML)
{
	LONG lIndex;
	clsXML oXML("WeekDays");
	oXML.SetSupportOptional(TRUE);
	oXML.SetXML(sXML);
	oXML.InitializeReader();
	mp_oCollection->m_Clear();
	for (lIndex = 1;lIndex <= oXML.ReadCollectionCount();lIndex++)
	{
		CalendarWeekDay* oCalendarWeekDay = new CalendarWeekDay();
		oCalendarWeekDay->SetXML(oXML.ReadCollectionObject(lIndex));
		mp_oCollection->SetAddMode(TRUE);
		CString sKey = _T("");
		oCalendarWeekDay->mp_oCollection = mp_oCollection;
		mp_oCollection->m_Add(oCalendarWeekDay, sKey, MP_ADD_1, MP_ADD_2, FALSE, MP_ADD_3);
	}
}