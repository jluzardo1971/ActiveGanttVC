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
#include "CalendarException.h"

IMPLEMENT_DYNCREATE(CalendarException, CCmdTarget)

//{7E7DEC31-6C72-4B95-B811-079440723DCD}
static const IID IID_ICalendarException = { 0x7E7DEC31, 0x6C72, 0x4B95, { 0xB8, 0x11, 0x07, 0x94, 0x40, 0x72, 0x3D, 0xCD} };

//{5BAD0845-03A8-410E-ADB0-1ECCC63D3D8C}
IMPLEMENT_OLECREATE_FLAGS(CalendarException, "MSP2010.CalendarException", afxRegApartmentThreading, 0x5BAD0845, 0x03A8, 0x410E, 0xAD, 0xB0, 0x1E, 0xCC, 0xC6, 0x3D, 0x3D, 0x8C)

BEGIN_DISPATCH_MAP(CalendarException, CCmdTarget)
DISP_PROPERTY_EX_ID(CalendarException, "bEnteredByOccurrences", 1, odl_GetEnteredByOccurrences, odl_SetEnteredByOccurrences, VT_BOOL)
DISP_PROPERTY_EX_ID(CalendarException, "oTimePeriod", 2, odl_GetTimePeriod, SetNotSupported, VT_DISPATCH)
DISP_PROPERTY_EX_ID(CalendarException, "lOccurrences", 3, odl_GetOccurrences, odl_SetOccurrences, VT_I4)
DISP_PROPERTY_EX_ID(CalendarException, "sName", 4, odl_GetName, odl_SetName, VT_BSTR)
DISP_PROPERTY_EX_ID(CalendarException, "yType", 5, odl_GetType, odl_SetType, VT_I4)
DISP_PROPERTY_EX_ID(CalendarException, "lPeriod", 6, odl_GetPeriod, odl_SetPeriod, VT_I4)
DISP_PROPERTY_EX_ID(CalendarException, "lDaysOfWeek", 7, odl_GetDaysOfWeek, odl_SetDaysOfWeek, VT_I4)
DISP_PROPERTY_EX_ID(CalendarException, "yMonthItem", 8, odl_GetMonthItem, odl_SetMonthItem, VT_I4)
DISP_PROPERTY_EX_ID(CalendarException, "yMonthPosition", 9, odl_GetMonthPosition, odl_SetMonthPosition, VT_I4)
DISP_PROPERTY_EX_ID(CalendarException, "yMonth", 10, odl_GetMonth, odl_SetMonth, VT_I4)
DISP_PROPERTY_EX_ID(CalendarException, "lMonthDay", 11, odl_GetMonthDay, odl_SetMonthDay, VT_I4)
DISP_PROPERTY_EX_ID(CalendarException, "bDayWorking", 12, odl_GetDayWorking, odl_SetDayWorking, VT_BOOL)
DISP_PROPERTY_EX_ID(CalendarException, "oWorkingTimes", 13, odl_GetWorkingTimes, SetNotSupported, VT_DISPATCH)
DISP_PROPERTY_EX_ID(CalendarException, "Key", 14, odl_GetKey, odl_SetKey, VT_BSTR)
DISP_FUNCTION_ID(CalendarException, "GetXML", 15, odl_GetXML, VT_BSTR, VTS_NONE)
DISP_FUNCTION_ID(CalendarException, "SetXML", 16, odl_SetXML, VT_EMPTY, VTS_BSTR)
DISP_FUNCTION_ID(CalendarException, "IsNull", 17, IsNull, VT_BOOL, VTS_NONE)
DISP_FUNCTION_ID(CalendarException, "Initialize", 18, Initialize, VT_EMPTY, VTS_NONE)
END_DISPATCH_MAP()

BEGIN_INTERFACE_MAP(CalendarException, CCmdTarget)
	INTERFACE_PART(CalendarException, IID_ICalendarException, Dispatch)
END_INTERFACE_MAP()

BEGIN_MESSAGE_MAP(CalendarException, CCmdTarget)
END_MESSAGE_MAP()


void CalendarException::Initialize(void)
{
	InitVars();
}

CalendarException::CalendarException()
{
	EnableAutomation();
	AfxOleLockApp();
	InitVars();
}

void CalendarException::InitVars(void)
{
	mp_bEnteredByOccurrences = FALSE;
	mp_oTimePeriod = new TimePeriod();
	mp_lOccurrences = 0;
	mp_sName = "";
	mp_yType = T_3_DAILY;
	mp_lPeriod = 0;
	mp_lDaysOfWeek = 0;
	mp_yMonthItem = MI_DAY;
	mp_yMonthPosition = MP_FIRST_POSITION;
	mp_yMonth = M_JANUARY;
	mp_lMonthDay = 0;
	mp_bDayWorking = FALSE;
	mp_oWorkingTimes = new WorkingTimes();
}

CalendarException::~CalendarException()
{
	delete mp_oTimePeriod;
	delete mp_oWorkingTimes;
	AfxOleUnlockApp();
}

void CalendarException::OnFinalRelease()
{
	clsItemBase::OnFinalRelease();
}

BOOL CalendarException::odl_GetEnteredByOccurrences(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetEnteredByOccurrences();
}

BOOL CalendarException::GetEnteredByOccurrences(void)
{
	return (BOOL) mp_bEnteredByOccurrences;
}

void CalendarException::odl_SetEnteredByOccurrences(BOOL newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetEnteredByOccurrences((BOOL)newVal);
}

void CalendarException::SetEnteredByOccurrences(BOOL newVal)
{
	mp_bEnteredByOccurrences = newVal;
}

IDispatch* CalendarException::odl_GetTimePeriod(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetTimePeriod()->GetIDispatch(TRUE);
}

TimePeriod* CalendarException::GetTimePeriod(void)
{
	return mp_oTimePeriod;
}

LONG CalendarException::odl_GetOccurrences(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetOccurrences();
}

LONG CalendarException::GetOccurrences(void)
{
	return (LONG) mp_lOccurrences;
}

void CalendarException::odl_SetOccurrences(LONG newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetOccurrences((LONG)newVal);
}

void CalendarException::SetOccurrences(LONG newVal)
{
	mp_lOccurrences = newVal;
}

BSTR CalendarException::odl_GetName(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetName().AllocSysString();
}

CString CalendarException::GetName(void)
{
	return (CString) mp_sName;
}

void CalendarException::odl_SetName(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetName(newVal);
}

void CalendarException::SetName(CString newVal)
{
	if (newVal.GetLength() > 512)
	{
		newVal = newVal.Mid(0, 512);
	}
	mp_sName = newVal;
}

LONG CalendarException::odl_GetType(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetType();
}

E_TYPE_3 CalendarException::GetType(void)
{
	return (E_TYPE_3) mp_yType;
}

void CalendarException::odl_SetType(LONG newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetType((E_TYPE_3)newVal);
}

void CalendarException::SetType(E_TYPE_3 newVal)
{
	mp_yType = newVal;
}

LONG CalendarException::odl_GetPeriod(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetPeriod();
}

LONG CalendarException::GetPeriod(void)
{
	return (LONG) mp_lPeriod;
}

void CalendarException::odl_SetPeriod(LONG newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetPeriod((LONG)newVal);
}

void CalendarException::SetPeriod(LONG newVal)
{
	mp_lPeriod = newVal;
}

LONG CalendarException::odl_GetDaysOfWeek(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetDaysOfWeek();
}

LONG CalendarException::GetDaysOfWeek(void)
{
	return (LONG) mp_lDaysOfWeek;
}

void CalendarException::odl_SetDaysOfWeek(LONG newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetDaysOfWeek((LONG)newVal);
}

void CalendarException::SetDaysOfWeek(LONG newVal)
{
	mp_lDaysOfWeek = newVal;
}

LONG CalendarException::odl_GetMonthItem(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetMonthItem();
}

E_MONTHITEM CalendarException::GetMonthItem(void)
{
	return (E_MONTHITEM) mp_yMonthItem;
}

void CalendarException::odl_SetMonthItem(LONG newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetMonthItem((E_MONTHITEM)newVal);
}

void CalendarException::SetMonthItem(E_MONTHITEM newVal)
{
	mp_yMonthItem = newVal;
}

LONG CalendarException::odl_GetMonthPosition(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetMonthPosition();
}

E_MONTHPOSITION CalendarException::GetMonthPosition(void)
{
	return (E_MONTHPOSITION) mp_yMonthPosition;
}

void CalendarException::odl_SetMonthPosition(LONG newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetMonthPosition((E_MONTHPOSITION)newVal);
}

void CalendarException::SetMonthPosition(E_MONTHPOSITION newVal)
{
	mp_yMonthPosition = newVal;
}

LONG CalendarException::odl_GetMonth(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetMonth();
}

E_MONTH CalendarException::GetMonth(void)
{
	return (E_MONTH) mp_yMonth;
}

void CalendarException::odl_SetMonth(LONG newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetMonth((E_MONTH)newVal);
}

void CalendarException::SetMonth(E_MONTH newVal)
{
	mp_yMonth = newVal;
}

LONG CalendarException::odl_GetMonthDay(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetMonthDay();
}

LONG CalendarException::GetMonthDay(void)
{
	return (LONG) mp_lMonthDay;
}

void CalendarException::odl_SetMonthDay(LONG newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetMonthDay((LONG)newVal);
}

void CalendarException::SetMonthDay(LONG newVal)
{
	mp_lMonthDay = newVal;
}

BOOL CalendarException::odl_GetDayWorking(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetDayWorking();
}

BOOL CalendarException::GetDayWorking(void)
{
	return (BOOL) mp_bDayWorking;
}

void CalendarException::odl_SetDayWorking(BOOL newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetDayWorking((BOOL)newVal);
}

void CalendarException::SetDayWorking(BOOL newVal)
{
	mp_bDayWorking = newVal;
}

IDispatch* CalendarException::odl_GetWorkingTimes(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetWorkingTimes()->GetIDispatch(TRUE);
}

WorkingTimes* CalendarException::GetWorkingTimes(void)
{
	return mp_oWorkingTimes;
}

BSTR CalendarException::odl_GetKey(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetKey().AllocSysString();
}

void CalendarException::odl_SetKey(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	USES_CONVERSION;
	SetKey(F_TOUCSTR(T2A(newVal)));
}

CString CalendarException::GetKey(void)
{
	return mp_sKey;
}

void CalendarException::SetKey(CString newVal)
{
	mp_oCollection->mp_SetKey(&mp_sKey, newVal, MP_SET_KEY);
}

BSTR CalendarException::odl_GetXML(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetXML().AllocSysString();
}

void CalendarException::odl_SetXML(LPCTSTR sXML)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetXML(sXML);
}

BOOL CalendarException::IsNull(void)
{
	BOOL bReturn = TRUE;
	if (mp_bEnteredByOccurrences != FALSE)
	{
		bReturn = FALSE;
	}
	if (mp_oTimePeriod->IsNull() == FALSE)
	{
		bReturn = FALSE;
	}
	if (mp_lOccurrences != 0)
	{
		bReturn = FALSE;
	}
	if (mp_sName != "")
	{
		bReturn = FALSE;
	}
	if (mp_yType != T_3_DAILY)
	{
		bReturn = FALSE;
	}
	if (mp_lPeriod != 0)
	{
		bReturn = FALSE;
	}
	if (mp_lDaysOfWeek != 0)
	{
		bReturn = FALSE;
	}
	if (mp_yMonthItem != MI_DAY)
	{
		bReturn = FALSE;
	}
	if (mp_yMonthPosition != MP_FIRST_POSITION)
	{
		bReturn = FALSE;
	}
	if (mp_yMonth != M_JANUARY)
	{
		bReturn = FALSE;
	}
	if (mp_lMonthDay != 0)
	{
		bReturn = FALSE;
	}
	if (mp_bDayWorking != FALSE)
	{
		bReturn = FALSE;
	}
	if (mp_oWorkingTimes->IsNull() == FALSE)
	{
		bReturn = FALSE;
	}
	return bReturn;
}

CString CalendarException::GetXML(void)
{
	if (IsNull() == TRUE)
	{
		return "<Exception/>";
	}
	clsXML oXML("Exception");
	oXML.InitializeWriter();
	oXML.SetSupportOptional(TRUE);
	oXML.SetBoolsAreNumeric(TRUE);
	oXML.WriteProperty("EnteredByOccurrences", mp_bEnteredByOccurrences);
	if (mp_oTimePeriod->IsNull() == FALSE)
	{
		oXML.WriteObject(mp_oTimePeriod->GetXML());
	}
	oXML.WriteProperty("Occurrences", mp_lOccurrences);
	if (mp_sName != "")
	{
		oXML.WriteProperty("Name", mp_sName);
	}
	oXML.WriteProperty("Type", mp_yType);
	oXML.WriteProperty("Period", mp_lPeriod);
	oXML.WriteProperty("DaysOfWeek", mp_lDaysOfWeek);
	oXML.WriteProperty("MonthItem", mp_yMonthItem);
	oXML.WriteProperty("MonthPosition", mp_yMonthPosition);
	oXML.WriteProperty("Month", mp_yMonth);
	oXML.WriteProperty("MonthDay", mp_lMonthDay);
	oXML.WriteProperty("DayWorking", mp_bDayWorking);
	if (mp_oWorkingTimes->IsNull() == FALSE)
	{
		oXML.WriteObject(mp_oWorkingTimes->GetXML());
	}
	return oXML.GetXML();
}

void CalendarException::SetXML(CString sXML)
{
	clsXML oXML("Exception");
	oXML.SetSupportOptional(TRUE);
	oXML.SetXML(sXML);
	oXML.InitializeReader();
	oXML.ReadProperty("EnteredByOccurrences", mp_bEnteredByOccurrences);
	mp_oTimePeriod->SetXML(oXML.ReadObject("TimePeriod"));
	oXML.ReadProperty("Occurrences", mp_lOccurrences);
	oXML.ReadProperty("Name", mp_sName);
	if (mp_sName.GetLength() > 512)
	{
		mp_sName = mp_sName.Mid(0, 512);
	}
	oXML.ReadProperty("Type", mp_yType);
	oXML.ReadProperty("Period", mp_lPeriod);
	oXML.ReadProperty("DaysOfWeek", mp_lDaysOfWeek);
	oXML.ReadProperty("MonthItem", mp_yMonthItem);
	oXML.ReadProperty("MonthPosition", mp_yMonthPosition);
	oXML.ReadProperty("Month", mp_yMonth);
	oXML.ReadProperty("MonthDay", mp_lMonthDay);
	oXML.ReadProperty("DayWorking", mp_bDayWorking);
	mp_oWorkingTimes->SetXML(oXML.ReadObject("WorkingTimes"));
}