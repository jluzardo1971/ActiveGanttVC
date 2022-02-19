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
#include "Calendars.h"

IMPLEMENT_DYNCREATE(Calendars, CCmdTarget)

//{A4B5D6CB-C08F-4790-BB6A-BAB695DEC2DE}
static const IID IID_ICalendars = { 0xA4B5D6CB, 0xC08F, 0x4790, { 0xBB, 0x6A, 0xBA, 0xB6, 0x95, 0xDE, 0xC2, 0xDE} };

//{6DD462CA-316B-41FC-92E3-71770074687C}
IMPLEMENT_OLECREATE_FLAGS(Calendars, "MSP2003.Calendars", afxRegApartmentThreading, 0x6DD462CA, 0x316B, 0x41FC, 0x92, 0xE3, 0x71, 0x77, 0x00, 0x74, 0x68, 0x7C)

BEGIN_DISPATCH_MAP(Calendars, CCmdTarget)
DISP_PROPERTY_EX_ID(Calendars, "Count", 1, odl_GetCount, SetNotSupported, VT_I4)
DISP_PROPERTY_PARAM_ID(Calendars, "Item", 2, odl_Item, SetNotSupported, VT_DISPATCH, VTS_BSTR)
DISP_FUNCTION_ID(Calendars, "Add", 3, odl_Add, VT_DISPATCH, VTS_NONE)
DISP_FUNCTION_ID(Calendars, "Clear", 4, odl_Clear, VT_EMPTY, VTS_NONE)
DISP_FUNCTION_ID(Calendars, "Remove", 5, odl_Remove, VT_EMPTY, VTS_BSTR)
DISP_FUNCTION_ID(Calendars, "IsNull", 6, IsNull, VT_BOOL, VTS_NONE)
DISP_FUNCTION_ID(Calendars, "Initialize", 7, Initialize, VT_EMPTY, VTS_NONE)
DISP_FUNCTION_ID(Calendars, "SetXML", 8, odl_SetXML, VT_EMPTY, VTS_BSTR)
DISP_FUNCTION_ID(Calendars, "GetXML", 9, odl_GetXML, VT_BSTR, VTS_NONE)
DISP_DEFVALUE(Calendars, "Item")
END_DISPATCH_MAP()

BEGIN_INTERFACE_MAP(Calendars, CCmdTarget)
	INTERFACE_PART(Calendars, IID_ICalendars, Dispatch)
END_INTERFACE_MAP()

BEGIN_MESSAGE_MAP(Calendars, CCmdTarget)
END_MESSAGE_MAP()


Calendars::Calendars()
{
	EnableAutomation();
	AfxOleLockApp();
	InitVars();
}

void Calendars::Initialize(void)
{
	InitVars();
}

void Calendars::InitVars(void)
{
	mp_oCollection = new clsCollectionBase("Calendar");
}

Calendars::~Calendars()
{
	delete mp_oCollection;
	AfxOleUnlockApp();
}

void Calendars::OnFinalRelease()
{
	CCmdTarget::OnFinalRelease();
}

LONG Calendars::odl_GetCount(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetCount();
}

LONG Calendars::GetCount(void)
{
	return mp_oCollection->m_lCount();
}

IDispatch* Calendars::odl_Item(LPCTSTR Index)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return Item(Index)->GetIDispatch(TRUE);
}

IDispatch* Calendars::odl_Add(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return Add()->GetIDispatch(TRUE);
}

void Calendars::odl_Clear(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	Clear();
}

void Calendars::odl_Remove(LPCTSTR Index)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	Remove(Index);
}

BOOL Calendars::IsNull(void)
{
	BOOL bReturn = TRUE;
	if (GetCount() > 0)
	{
		bReturn = FALSE;
	}
	return bReturn;
}

Calendar* Calendars::Item(CString Index)
{
	Calendar *oCalendar;
	oCalendar = (Calendar*)mp_oCollection->m_oItem(Index, MP_ITEM_1, MP_ITEM_2, MP_ITEM_3, MP_ITEM_4);
	return oCalendar;
}

Calendar* Calendars::Add()
{
	mp_oCollection->SetAddMode(TRUE);
	Calendar* oCalendar = new Calendar();
	oCalendar->mp_oCollection = mp_oCollection;
	mp_oCollection->m_Add(oCalendar, _T(""), MP_ADD_1, MP_ADD_2, FALSE, MP_ADD_3);
	return oCalendar;
}

void Calendars::Clear(void)
{
	mp_oCollection->m_Clear();
}

void Calendars::Remove(CString Index)
{
	mp_oCollection->m_Remove(Index, MP_REMOVE_1, MP_REMOVE_2, MP_REMOVE_3, MP_REMOVE_4);
}

BSTR Calendars::odl_GetXML(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetXML().AllocSysString();
}

void Calendars::odl_SetXML(LPCTSTR sXML)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetXML(sXML);
}

CString Calendars::GetXML(void)
{
	if (IsNull() == TRUE)
	{
		return "<Calendars/>";
	}
	LONG lIndex;
	Calendar* oCalendar;
	clsXML oXML("Calendars");
	oXML.InitializeWriter();
	for (lIndex = 1; lIndex <= GetCount(); lIndex++)
	{
		oCalendar = (Calendar*) mp_oCollection->m_oReturnArrayElement(lIndex);
		oXML.WriteObject(oCalendar->GetXML());
	}
	return oXML.GetXML();
}

void Calendars::SetXML(CString sXML)
{
	LONG lIndex;
	clsXML oXML("Calendars");
	oXML.SetSupportOptional(TRUE);
	oXML.SetXML(sXML);
	oXML.InitializeReader();
	mp_oCollection->m_Clear();
	for (lIndex = 1;lIndex <= oXML.ReadCollectionCount();lIndex++)
	{
		Calendar* oCalendar = new Calendar();
		oCalendar->SetXML(oXML.ReadCollectionObject(lIndex));
		mp_oCollection->SetAddMode(TRUE);
		CString sKey = _T("");
		sKey.Format(_T("%d"), oCalendar->GetUID());
		sKey = _T("K") + sKey;
		oCalendar->mp_oCollection = mp_oCollection;
		oCalendar->SetKey(sKey);
		mp_oCollection->m_Add(oCalendar, sKey, MP_ADD_1, MP_ADD_2, FALSE, MP_ADD_3);
	}
}