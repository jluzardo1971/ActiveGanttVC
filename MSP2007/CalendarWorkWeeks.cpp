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
#include "CalendarWorkWeeks.h"

IMPLEMENT_DYNCREATE(CalendarWorkWeeks, CCmdTarget)

//{DBB34BDB-9DF2-4712-8E71-D43206ECC527}
static const IID IID_ICalendarWorkWeeks = { 0xDBB34BDB, 0x9DF2, 0x4712, { 0x8E, 0x71, 0xD4, 0x32, 0x06, 0xEC, 0xC5, 0x27} };

//{E437B795-35AD-4DE1-8D70-6AF127998FED}
IMPLEMENT_OLECREATE_FLAGS(CalendarWorkWeeks, "MSP2007.CalendarWorkWeeks", afxRegApartmentThreading, 0xE437B795, 0x35AD, 0x4DE1, 0x8D, 0x70, 0x6A, 0xF1, 0x27, 0x99, 0x8F, 0xED)

BEGIN_DISPATCH_MAP(CalendarWorkWeeks, CCmdTarget)
DISP_PROPERTY_EX_ID(CalendarWorkWeeks, "Count", 1, odl_GetCount, SetNotSupported, VT_I4)
DISP_PROPERTY_PARAM_ID(CalendarWorkWeeks, "Item", 2, odl_Item, SetNotSupported, VT_DISPATCH, VTS_BSTR)
DISP_FUNCTION_ID(CalendarWorkWeeks, "Add", 3, odl_Add, VT_DISPATCH, VTS_NONE)
DISP_FUNCTION_ID(CalendarWorkWeeks, "Clear", 4, odl_Clear, VT_EMPTY, VTS_NONE)
DISP_FUNCTION_ID(CalendarWorkWeeks, "Remove", 5, odl_Remove, VT_EMPTY, VTS_BSTR)
DISP_FUNCTION_ID(CalendarWorkWeeks, "IsNull", 6, IsNull, VT_BOOL, VTS_NONE)
DISP_FUNCTION_ID(CalendarWorkWeeks, "Initialize", 7, Initialize, VT_EMPTY, VTS_NONE)
DISP_FUNCTION_ID(CalendarWorkWeeks, "SetXML", 8, odl_SetXML, VT_EMPTY, VTS_BSTR)
DISP_FUNCTION_ID(CalendarWorkWeeks, "GetXML", 9, odl_GetXML, VT_BSTR, VTS_NONE)
DISP_DEFVALUE(CalendarWorkWeeks, "Item")
END_DISPATCH_MAP()

BEGIN_INTERFACE_MAP(CalendarWorkWeeks, CCmdTarget)
	INTERFACE_PART(CalendarWorkWeeks, IID_ICalendarWorkWeeks, Dispatch)
END_INTERFACE_MAP()

BEGIN_MESSAGE_MAP(CalendarWorkWeeks, CCmdTarget)
END_MESSAGE_MAP()


CalendarWorkWeeks::CalendarWorkWeeks()
{
	EnableAutomation();
	AfxOleLockApp();
	InitVars();
}

void CalendarWorkWeeks::Initialize(void)
{
	InitVars();
}

void CalendarWorkWeeks::InitVars(void)
{
	mp_oCollection = new clsCollectionBase("CalendarWorkWeek");
}

CalendarWorkWeeks::~CalendarWorkWeeks()
{
	delete mp_oCollection;
	AfxOleUnlockApp();
}

void CalendarWorkWeeks::OnFinalRelease()
{
	CCmdTarget::OnFinalRelease();
}

LONG CalendarWorkWeeks::odl_GetCount(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetCount();
}

LONG CalendarWorkWeeks::GetCount(void)
{
	return mp_oCollection->m_lCount();
}

IDispatch* CalendarWorkWeeks::odl_Item(LPCTSTR Index)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return Item(Index)->GetIDispatch(TRUE);
}

IDispatch* CalendarWorkWeeks::odl_Add(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return Add()->GetIDispatch(TRUE);
}

void CalendarWorkWeeks::odl_Clear(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	Clear();
}

void CalendarWorkWeeks::odl_Remove(LPCTSTR Index)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	Remove(Index);
}

BOOL CalendarWorkWeeks::IsNull(void)
{
	BOOL bReturn = TRUE;
	if (GetCount() > 0)
	{
		bReturn = FALSE;
	}
	return bReturn;
}

CalendarWorkWeek* CalendarWorkWeeks::Item(CString Index)
{
	CalendarWorkWeek *oCalendarWorkWeek;
	oCalendarWorkWeek = (CalendarWorkWeek*)mp_oCollection->m_oItem(Index, MP_ITEM_1, MP_ITEM_2, MP_ITEM_3, MP_ITEM_4);
	return oCalendarWorkWeek;
}

CalendarWorkWeek* CalendarWorkWeeks::Add()
{
	mp_oCollection->SetAddMode(TRUE);
	CalendarWorkWeek* oCalendarWorkWeek = new CalendarWorkWeek();
	oCalendarWorkWeek->mp_oCollection = mp_oCollection;
	mp_oCollection->m_Add(oCalendarWorkWeek, _T(""), MP_ADD_1, MP_ADD_2, FALSE, MP_ADD_3);
	return oCalendarWorkWeek;
}

void CalendarWorkWeeks::Clear(void)
{
	mp_oCollection->m_Clear();
}

void CalendarWorkWeeks::Remove(CString Index)
{
	mp_oCollection->m_Remove(Index, MP_REMOVE_1, MP_REMOVE_2, MP_REMOVE_3, MP_REMOVE_4);
}

BSTR CalendarWorkWeeks::odl_GetXML(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetXML().AllocSysString();
}

void CalendarWorkWeeks::odl_SetXML(LPCTSTR sXML)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetXML(sXML);
}

CString CalendarWorkWeeks::GetXML(void)
{
	if (IsNull() == TRUE)
	{
		return "<WorkWeeks/>";
	}
	LONG lIndex;
	CalendarWorkWeek* oCalendarWorkWeek;
	clsXML oXML("WorkWeeks");
	oXML.InitializeWriter();
	for (lIndex = 1; lIndex <= GetCount(); lIndex++)
	{
		oCalendarWorkWeek = (CalendarWorkWeek*) mp_oCollection->m_oReturnArrayElement(lIndex);
		oXML.WriteObject(oCalendarWorkWeek->GetXML());
	}
	return oXML.GetXML();
}

void CalendarWorkWeeks::SetXML(CString sXML)
{
	LONG lIndex;
	clsXML oXML("WorkWeeks");
	oXML.SetSupportOptional(TRUE);
	oXML.SetXML(sXML);
	oXML.InitializeReader();
	mp_oCollection->m_Clear();
	for (lIndex = 1;lIndex <= oXML.ReadCollectionCount();lIndex++)
	{
		CalendarWorkWeek* oCalendarWorkWeek = new CalendarWorkWeek();
		oCalendarWorkWeek->SetXML(oXML.ReadCollectionObject(lIndex));
		mp_oCollection->SetAddMode(TRUE);
		CString sKey = _T("");
		oCalendarWorkWeek->mp_oCollection = mp_oCollection;
		mp_oCollection->m_Add(oCalendarWorkWeek, sKey, MP_ADD_1, MP_ADD_2, FALSE, MP_ADD_3);
	}
}