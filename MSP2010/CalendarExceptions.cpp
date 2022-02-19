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
#include "CalendarExceptions.h"

IMPLEMENT_DYNCREATE(CalendarExceptions, CCmdTarget)

//{404D9DA7-3D96-4450-9447-8A9087B75752}
static const IID IID_ICalendarExceptions = { 0x404D9DA7, 0x3D96, 0x4450, { 0x94, 0x47, 0x8A, 0x90, 0x87, 0xB7, 0x57, 0x52} };

//{38AA04C1-920E-4F86-A99C-1D86FE371299}
IMPLEMENT_OLECREATE_FLAGS(CalendarExceptions, "MSP2010.CalendarExceptions", afxRegApartmentThreading, 0x38AA04C1, 0x920E, 0x4F86, 0xA9, 0x9C, 0x1D, 0x86, 0xFE, 0x37, 0x12, 0x99)

BEGIN_DISPATCH_MAP(CalendarExceptions, CCmdTarget)
DISP_PROPERTY_EX_ID(CalendarExceptions, "Count", 1, odl_GetCount, SetNotSupported, VT_I4)
DISP_PROPERTY_PARAM_ID(CalendarExceptions, "Item", 2, odl_Item, SetNotSupported, VT_DISPATCH, VTS_BSTR)
DISP_FUNCTION_ID(CalendarExceptions, "Add", 3, odl_Add, VT_DISPATCH, VTS_NONE)
DISP_FUNCTION_ID(CalendarExceptions, "Clear", 4, odl_Clear, VT_EMPTY, VTS_NONE)
DISP_FUNCTION_ID(CalendarExceptions, "Remove", 5, odl_Remove, VT_EMPTY, VTS_BSTR)
DISP_FUNCTION_ID(CalendarExceptions, "IsNull", 6, IsNull, VT_BOOL, VTS_NONE)
DISP_FUNCTION_ID(CalendarExceptions, "Initialize", 7, Initialize, VT_EMPTY, VTS_NONE)
DISP_FUNCTION_ID(CalendarExceptions, "SetXML", 8, odl_SetXML, VT_EMPTY, VTS_BSTR)
DISP_FUNCTION_ID(CalendarExceptions, "GetXML", 9, odl_GetXML, VT_BSTR, VTS_NONE)
DISP_DEFVALUE(CalendarExceptions, "Item")
END_DISPATCH_MAP()

BEGIN_INTERFACE_MAP(CalendarExceptions, CCmdTarget)
	INTERFACE_PART(CalendarExceptions, IID_ICalendarExceptions, Dispatch)
END_INTERFACE_MAP()

BEGIN_MESSAGE_MAP(CalendarExceptions, CCmdTarget)
END_MESSAGE_MAP()


CalendarExceptions::CalendarExceptions()
{
	EnableAutomation();
	AfxOleLockApp();
	InitVars();
}

void CalendarExceptions::Initialize(void)
{
	InitVars();
}

void CalendarExceptions::InitVars(void)
{
	mp_oCollection = new clsCollectionBase("CalendarException");
}

CalendarExceptions::~CalendarExceptions()
{
	delete mp_oCollection;
	AfxOleUnlockApp();
}

void CalendarExceptions::OnFinalRelease()
{
	CCmdTarget::OnFinalRelease();
}

LONG CalendarExceptions::odl_GetCount(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetCount();
}

LONG CalendarExceptions::GetCount(void)
{
	return mp_oCollection->m_lCount();
}

IDispatch* CalendarExceptions::odl_Item(LPCTSTR Index)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return Item(Index)->GetIDispatch(TRUE);
}

IDispatch* CalendarExceptions::odl_Add(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return Add()->GetIDispatch(TRUE);
}

void CalendarExceptions::odl_Clear(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	Clear();
}

void CalendarExceptions::odl_Remove(LPCTSTR Index)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	Remove(Index);
}

BOOL CalendarExceptions::IsNull(void)
{
	BOOL bReturn = TRUE;
	if (GetCount() > 0)
	{
		bReturn = FALSE;
	}
	return bReturn;
}

CalendarException* CalendarExceptions::Item(CString Index)
{
	CalendarException *oCalendarException;
	oCalendarException = (CalendarException*)mp_oCollection->m_oItem(Index, MP_ITEM_1, MP_ITEM_2, MP_ITEM_3, MP_ITEM_4);
	return oCalendarException;
}

CalendarException* CalendarExceptions::Add()
{
	mp_oCollection->SetAddMode(TRUE);
	CalendarException* oCalendarException = new CalendarException();
	oCalendarException->mp_oCollection = mp_oCollection;
	mp_oCollection->m_Add(oCalendarException, _T(""), MP_ADD_1, MP_ADD_2, FALSE, MP_ADD_3);
	return oCalendarException;
}

void CalendarExceptions::Clear(void)
{
	mp_oCollection->m_Clear();
}

void CalendarExceptions::Remove(CString Index)
{
	mp_oCollection->m_Remove(Index, MP_REMOVE_1, MP_REMOVE_2, MP_REMOVE_3, MP_REMOVE_4);
}

BSTR CalendarExceptions::odl_GetXML(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetXML().AllocSysString();
}

void CalendarExceptions::odl_SetXML(LPCTSTR sXML)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetXML(sXML);
}

CString CalendarExceptions::GetXML(void)
{
	if (IsNull() == TRUE)
	{
		return "<Exceptions/>";
	}
	LONG lIndex;
	CalendarException* oCalendarException;
	clsXML oXML("Exceptions");
	oXML.InitializeWriter();
	for (lIndex = 1; lIndex <= GetCount(); lIndex++)
	{
		oCalendarException = (CalendarException*) mp_oCollection->m_oReturnArrayElement(lIndex);
		oXML.WriteObject(oCalendarException->GetXML());
	}
	return oXML.GetXML();
}

void CalendarExceptions::SetXML(CString sXML)
{
	LONG lIndex;
	clsXML oXML("Exceptions");
	oXML.SetSupportOptional(TRUE);
	oXML.SetXML(sXML);
	oXML.InitializeReader();
	mp_oCollection->m_Clear();
	for (lIndex = 1;lIndex <= oXML.ReadCollectionCount();lIndex++)
	{
		CalendarException* oCalendarException = new CalendarException();
		oCalendarException->SetXML(oXML.ReadCollectionObject(lIndex));
		mp_oCollection->SetAddMode(TRUE);
		CString sKey = _T("");
		oCalendarException->mp_oCollection = mp_oCollection;
		mp_oCollection->m_Add(oCalendarException, sKey, MP_ADD_1, MP_ADD_2, FALSE, MP_ADD_3);
	}
}