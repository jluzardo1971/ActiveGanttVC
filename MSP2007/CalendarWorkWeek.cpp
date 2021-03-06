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
#include "CalendarWorkWeek.h"

IMPLEMENT_DYNCREATE(CalendarWorkWeek, CCmdTarget)

//{72A1F0C4-1841-4B32-BC09-36987E17021A}
static const IID IID_ICalendarWorkWeek = { 0x72A1F0C4, 0x1841, 0x4B32, { 0xBC, 0x09, 0x36, 0x98, 0x7E, 0x17, 0x02, 0x1A} };

//{3880EC9D-744A-4AD6-9E42-15EE98E08289}
IMPLEMENT_OLECREATE_FLAGS(CalendarWorkWeek, "MSP2007.CalendarWorkWeek", afxRegApartmentThreading, 0x3880EC9D, 0x744A, 0x4AD6, 0x9E, 0x42, 0x15, 0xEE, 0x98, 0xE0, 0x82, 0x89)

BEGIN_DISPATCH_MAP(CalendarWorkWeek, CCmdTarget)
DISP_PROPERTY_EX_ID(CalendarWorkWeek, "oTimePeriod", 1, odl_GetTimePeriod, SetNotSupported, VT_DISPATCH)
DISP_PROPERTY_EX_ID(CalendarWorkWeek, "sName", 2, odl_GetName, odl_SetName, VT_BSTR)
DISP_PROPERTY_EX_ID(CalendarWorkWeek, "oWeekDay", 3, odl_GetWeekDay, SetNotSupported, VT_DISPATCH)
DISP_PROPERTY_EX_ID(CalendarWorkWeek, "Key", 4, odl_GetKey, odl_SetKey, VT_BSTR)
DISP_FUNCTION_ID(CalendarWorkWeek, "GetXML", 5, odl_GetXML, VT_BSTR, VTS_NONE)
DISP_FUNCTION_ID(CalendarWorkWeek, "SetXML", 6, odl_SetXML, VT_EMPTY, VTS_BSTR)
DISP_FUNCTION_ID(CalendarWorkWeek, "IsNull", 7, IsNull, VT_BOOL, VTS_NONE)
DISP_FUNCTION_ID(CalendarWorkWeek, "Initialize", 8, Initialize, VT_EMPTY, VTS_NONE)
END_DISPATCH_MAP()

BEGIN_INTERFACE_MAP(CalendarWorkWeek, CCmdTarget)
	INTERFACE_PART(CalendarWorkWeek, IID_ICalendarWorkWeek, Dispatch)
END_INTERFACE_MAP()

BEGIN_MESSAGE_MAP(CalendarWorkWeek, CCmdTarget)
END_MESSAGE_MAP()


void CalendarWorkWeek::Initialize(void)
{
	InitVars();
}

CalendarWorkWeek::CalendarWorkWeek()
{
	EnableAutomation();
	AfxOleLockApp();
	InitVars();
}

void CalendarWorkWeek::InitVars(void)
{
	mp_oTimePeriod = new TimePeriod();
	mp_sName = "";
	mp_oWeekDay_C = new WeekDay_C();
}

CalendarWorkWeek::~CalendarWorkWeek()
{
	delete mp_oTimePeriod;
	delete mp_oWeekDay_C;
	AfxOleUnlockApp();
}

void CalendarWorkWeek::OnFinalRelease()
{
	clsItemBase::OnFinalRelease();
}

IDispatch* CalendarWorkWeek::odl_GetTimePeriod(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetTimePeriod()->GetIDispatch(TRUE);
}

TimePeriod* CalendarWorkWeek::GetTimePeriod(void)
{
	return mp_oTimePeriod;
}

BSTR CalendarWorkWeek::odl_GetName(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetName().AllocSysString();
}

CString CalendarWorkWeek::GetName(void)
{
	return (CString) mp_sName;
}

void CalendarWorkWeek::odl_SetName(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetName(newVal);
}

void CalendarWorkWeek::SetName(CString newVal)
{
	if (newVal.GetLength() > 512)
	{
		newVal = newVal.Mid(0, 512);
	}
	mp_sName = newVal;
}

IDispatch* CalendarWorkWeek::odl_GetWeekDay(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetWeekDay()->GetIDispatch(TRUE);
}

WeekDay_C* CalendarWorkWeek::GetWeekDay(void)
{
	return mp_oWeekDay_C;
}

BSTR CalendarWorkWeek::odl_GetKey(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetKey().AllocSysString();
}

void CalendarWorkWeek::odl_SetKey(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	USES_CONVERSION;
	SetKey(F_TOUCSTR(T2A(newVal)));
}

CString CalendarWorkWeek::GetKey(void)
{
	return mp_sKey;
}

void CalendarWorkWeek::SetKey(CString newVal)
{
	mp_oCollection->mp_SetKey(&mp_sKey, newVal, MP_SET_KEY);
}

BSTR CalendarWorkWeek::odl_GetXML(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetXML().AllocSysString();
}

void CalendarWorkWeek::odl_SetXML(LPCTSTR sXML)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetXML(sXML);
}

BOOL CalendarWorkWeek::IsNull(void)
{
	BOOL bReturn = TRUE;
	if (mp_oTimePeriod->IsNull() == FALSE)
	{
		bReturn = FALSE;
	}
	if (mp_sName != "")
	{
		bReturn = FALSE;
	}
	if (mp_oWeekDay_C->IsNull() == FALSE)
	{
		bReturn = FALSE;
	}
	return bReturn;
}

CString CalendarWorkWeek::GetXML(void)
{
	if (IsNull() == TRUE)
	{
		return "<WorkWeek/>";
	}
	clsXML oXML("WorkWeek");
	oXML.InitializeWriter();
	oXML.SetSupportOptional(TRUE);
	oXML.SetBoolsAreNumeric(TRUE);
	if (mp_oTimePeriod->IsNull() == FALSE)
	{
		oXML.WriteObject(mp_oTimePeriod->GetXML());
	}
	if (mp_sName != "")
	{
		oXML.WriteProperty("Name", mp_sName);
	}
	if (mp_oWeekDay_C->IsNull() == FALSE)
	{
		mp_oWeekDay_C->WriteObjectProtected(oXML);
	}
	return oXML.GetXML();
}

void CalendarWorkWeek::SetXML(CString sXML)
{
	clsXML oXML("WorkWeek");
	oXML.SetSupportOptional(TRUE);
	oXML.SetXML(sXML);
	oXML.InitializeReader();
	mp_oTimePeriod->SetXML(oXML.ReadObject("TimePeriod"));
	oXML.ReadProperty("Name", mp_sName);
	if (mp_sName.GetLength() > 512)
	{
		mp_sName = mp_sName.Mid(0, 512);
	}
	mp_oWeekDay_C->ReadObjectProtected(oXML);
}