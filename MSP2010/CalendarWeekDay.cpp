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
#include "CalendarWeekDay.h"

IMPLEMENT_DYNCREATE(CalendarWeekDay, CCmdTarget)

//{10D2F1FC-9A15-4209-9F27-255E8FD994B9}
static const IID IID_ICalendarWeekDay = { 0x10D2F1FC, 0x9A15, 0x4209, { 0x9F, 0x27, 0x25, 0x5E, 0x8F, 0xD9, 0x94, 0xB9} };

//{E0901EBE-D78C-4804-A0D6-8E2C732B616A}
IMPLEMENT_OLECREATE_FLAGS(CalendarWeekDay, "MSP2010.CalendarWeekDay", afxRegApartmentThreading, 0xE0901EBE, 0xD78C, 0x4804, 0xA0, 0xD6, 0x8E, 0x2C, 0x73, 0x2B, 0x61, 0x6A)

BEGIN_DISPATCH_MAP(CalendarWeekDay, CCmdTarget)
DISP_PROPERTY_EX_ID(CalendarWeekDay, "yDayType", 1, odl_GetDayType, odl_SetDayType, VT_I4)
DISP_PROPERTY_EX_ID(CalendarWeekDay, "bDayWorking", 2, odl_GetDayWorking, odl_SetDayWorking, VT_BOOL)
DISP_PROPERTY_EX_ID(CalendarWeekDay, "oTimePeriod", 3, odl_GetTimePeriod, SetNotSupported, VT_DISPATCH)
DISP_PROPERTY_EX_ID(CalendarWeekDay, "oWorkingTimes", 4, odl_GetWorkingTimes, SetNotSupported, VT_DISPATCH)
DISP_PROPERTY_EX_ID(CalendarWeekDay, "Key", 5, odl_GetKey, odl_SetKey, VT_BSTR)
DISP_FUNCTION_ID(CalendarWeekDay, "GetXML", 6, odl_GetXML, VT_BSTR, VTS_NONE)
DISP_FUNCTION_ID(CalendarWeekDay, "SetXML", 7, odl_SetXML, VT_EMPTY, VTS_BSTR)
DISP_FUNCTION_ID(CalendarWeekDay, "IsNull", 8, IsNull, VT_BOOL, VTS_NONE)
DISP_FUNCTION_ID(CalendarWeekDay, "Initialize", 9, Initialize, VT_EMPTY, VTS_NONE)
END_DISPATCH_MAP()

BEGIN_INTERFACE_MAP(CalendarWeekDay, CCmdTarget)
	INTERFACE_PART(CalendarWeekDay, IID_ICalendarWeekDay, Dispatch)
END_INTERFACE_MAP()

BEGIN_MESSAGE_MAP(CalendarWeekDay, CCmdTarget)
END_MESSAGE_MAP()


void CalendarWeekDay::Initialize(void)
{
	InitVars();
}

CalendarWeekDay::CalendarWeekDay()
{
	EnableAutomation();
	AfxOleLockApp();
	InitVars();
}

void CalendarWeekDay::InitVars(void)
{
	mp_yDayType = DT_EXCEPTION;
	mp_bDayWorking = FALSE;
	mp_oTimePeriod = new TimePeriod();
	mp_oWorkingTimes = new WorkingTimes();
}

CalendarWeekDay::~CalendarWeekDay()
{
	delete mp_oTimePeriod;
	delete mp_oWorkingTimes;
	AfxOleUnlockApp();
}

void CalendarWeekDay::OnFinalRelease()
{
	clsItemBase::OnFinalRelease();
}

LONG CalendarWeekDay::odl_GetDayType(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetDayType();
}

E_DAYTYPE CalendarWeekDay::GetDayType(void)
{
	return (E_DAYTYPE) mp_yDayType;
}

void CalendarWeekDay::odl_SetDayType(LONG newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetDayType((E_DAYTYPE)newVal);
}

void CalendarWeekDay::SetDayType(E_DAYTYPE newVal)
{
	mp_yDayType = newVal;
}

BOOL CalendarWeekDay::odl_GetDayWorking(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetDayWorking();
}

BOOL CalendarWeekDay::GetDayWorking(void)
{
	return (BOOL) mp_bDayWorking;
}

void CalendarWeekDay::odl_SetDayWorking(BOOL newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetDayWorking((BOOL)newVal);
}

void CalendarWeekDay::SetDayWorking(BOOL newVal)
{
	mp_bDayWorking = newVal;
}

IDispatch* CalendarWeekDay::odl_GetTimePeriod(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetTimePeriod()->GetIDispatch(TRUE);
}

TimePeriod* CalendarWeekDay::GetTimePeriod(void)
{
	return mp_oTimePeriod;
}

IDispatch* CalendarWeekDay::odl_GetWorkingTimes(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetWorkingTimes()->GetIDispatch(TRUE);
}

WorkingTimes* CalendarWeekDay::GetWorkingTimes(void)
{
	return mp_oWorkingTimes;
}

BSTR CalendarWeekDay::odl_GetKey(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetKey().AllocSysString();
}

void CalendarWeekDay::odl_SetKey(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	USES_CONVERSION;
	SetKey(F_TOUCSTR(T2A(newVal)));
}

CString CalendarWeekDay::GetKey(void)
{
	return mp_sKey;
}

void CalendarWeekDay::SetKey(CString newVal)
{
	mp_oCollection->mp_SetKey(&mp_sKey, newVal, MP_SET_KEY);
}

BSTR CalendarWeekDay::odl_GetXML(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetXML().AllocSysString();
}

void CalendarWeekDay::odl_SetXML(LPCTSTR sXML)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetXML(sXML);
}

BOOL CalendarWeekDay::IsNull(void)
{
	BOOL bReturn = TRUE;
	if (mp_yDayType != DT_EXCEPTION)
	{
		bReturn = FALSE;
	}
	if (mp_bDayWorking != FALSE)
	{
		bReturn = FALSE;
	}
	if (mp_oTimePeriod->IsNull() == FALSE)
	{
		bReturn = FALSE;
	}
	if (mp_oWorkingTimes->IsNull() == FALSE)
	{
		bReturn = FALSE;
	}
	return bReturn;
}

CString CalendarWeekDay::GetXML(void)
{
	if (IsNull() == TRUE)
	{
		return "<WeekDay/>";
	}
	clsXML oXML("WeekDay");
	oXML.InitializeWriter();
	oXML.SetSupportOptional(TRUE);
	oXML.SetBoolsAreNumeric(TRUE);
	oXML.WriteProperty("DayType", mp_yDayType);
	oXML.WriteProperty("DayWorking", mp_bDayWorking);
	if (mp_oTimePeriod->IsNull() == FALSE)
	{
		oXML.WriteObject(mp_oTimePeriod->GetXML());
	}
	if (mp_oWorkingTimes->IsNull() == FALSE)
	{
		oXML.WriteObject(mp_oWorkingTimes->GetXML());
	}
	return oXML.GetXML();
}

void CalendarWeekDay::SetXML(CString sXML)
{
	clsXML oXML("WeekDay");
	oXML.SetSupportOptional(TRUE);
	oXML.SetXML(sXML);
	oXML.InitializeReader();
	oXML.ReadProperty("DayType", mp_yDayType);
	oXML.ReadProperty("DayWorking", mp_bDayWorking);
	mp_oTimePeriod->SetXML(oXML.ReadObject("TimePeriod"));
	mp_oWorkingTimes->SetXML(oXML.ReadObject("WorkingTimes"));
}