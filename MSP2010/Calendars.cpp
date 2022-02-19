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

//{66CC77E6-1E88-4D97-BED9-414F70727D86}
static const IID IID_ICalendars = { 0x66CC77E6, 0x1E88, 0x4D97, { 0xBE, 0xD9, 0x41, 0x4F, 0x70, 0x72, 0x7D, 0x86} };

//{764C1066-DE21-42F2-B474-AE39E8654F0A}
IMPLEMENT_OLECREATE_FLAGS(Calendars, "MSP2010.Calendars", afxRegApartmentThreading, 0x764C1066, 0xDE21, 0x42F2, 0xB4, 0x74, 0xAE, 0x39, 0xE8, 0x65, 0x4F, 0x0A)

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