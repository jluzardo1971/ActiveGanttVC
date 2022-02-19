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
#include "Calendar.h"

IMPLEMENT_DYNCREATE(Calendar, CCmdTarget)

//{99D4D0EE-085B-4714-A4EF-D2C9ED12866E}
static const IID IID_ICalendar = { 0x99D4D0EE, 0x085B, 0x4714, { 0xA4, 0xEF, 0xD2, 0xC9, 0xED, 0x12, 0x86, 0x6E} };

//{B3FE80EF-CD73-47BA-BEB2-238F8B6B2782}
IMPLEMENT_OLECREATE_FLAGS(Calendar, "MSP2010.Calendar", afxRegApartmentThreading, 0xB3FE80EF, 0xCD73, 0x47BA, 0xBE, 0xB2, 0x23, 0x8F, 0x8B, 0x6B, 0x27, 0x82)

BEGIN_DISPATCH_MAP(Calendar, CCmdTarget)
DISP_PROPERTY_EX_ID(Calendar, "lUID", 1, odl_GetUID, odl_SetUID, VT_I4)
DISP_PROPERTY_EX_ID(Calendar, "sName", 2, odl_GetName, odl_SetName, VT_BSTR)
DISP_PROPERTY_EX_ID(Calendar, "bIsBaseCalendar", 3, odl_GetIsBaseCalendar, odl_SetIsBaseCalendar, VT_BOOL)
DISP_PROPERTY_EX_ID(Calendar, "bIsBaselineCalendar", 4, odl_GetIsBaselineCalendar, odl_SetIsBaselineCalendar, VT_BOOL)
DISP_PROPERTY_EX_ID(Calendar, "lBaseCalendarUID", 5, odl_GetBaseCalendarUID, odl_SetBaseCalendarUID, VT_I4)
DISP_PROPERTY_EX_ID(Calendar, "oWeekDays", 6, odl_GetWeekDays, SetNotSupported, VT_DISPATCH)
DISP_PROPERTY_EX_ID(Calendar, "oExceptions", 7, odl_GetExceptions, SetNotSupported, VT_DISPATCH)
DISP_PROPERTY_EX_ID(Calendar, "oWorkWeeks", 8, odl_GetWorkWeeks, SetNotSupported, VT_DISPATCH)
DISP_PROPERTY_EX_ID(Calendar, "Key", 9, odl_GetKey, odl_SetKey, VT_BSTR)
DISP_FUNCTION_ID(Calendar, "GetXML", 10, odl_GetXML, VT_BSTR, VTS_NONE)
DISP_FUNCTION_ID(Calendar, "SetXML", 11, odl_SetXML, VT_EMPTY, VTS_BSTR)
DISP_FUNCTION_ID(Calendar, "IsNull", 12, IsNull, VT_BOOL, VTS_NONE)
DISP_FUNCTION_ID(Calendar, "Initialize", 13, Initialize, VT_EMPTY, VTS_NONE)
END_DISPATCH_MAP()

BEGIN_INTERFACE_MAP(Calendar, CCmdTarget)
	INTERFACE_PART(Calendar, IID_ICalendar, Dispatch)
END_INTERFACE_MAP()

BEGIN_MESSAGE_MAP(Calendar, CCmdTarget)
END_MESSAGE_MAP()


void Calendar::Initialize(void)
{
	InitVars();
}

Calendar::Calendar()
{
	EnableAutomation();
	AfxOleLockApp();
	InitVars();
}

void Calendar::InitVars(void)
{
	mp_lUID = 0;
	mp_sName = "";
	mp_bIsBaseCalendar = FALSE;
	mp_bIsBaselineCalendar = FALSE;
	mp_lBaseCalendarUID = 0;
	mp_oWeekDays = new CalendarWeekDays();
	mp_oExceptions = new CalendarExceptions();
	mp_oWorkWeeks = new CalendarWorkWeeks();
}

Calendar::~Calendar()
{
	delete mp_oWeekDays;
	delete mp_oExceptions;
	delete mp_oWorkWeeks;
	AfxOleUnlockApp();
}

void Calendar::OnFinalRelease()
{
	clsItemBase::OnFinalRelease();
}

LONG Calendar::odl_GetUID(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetUID();
}

LONG Calendar::GetUID(void)
{
	return (LONG) mp_lUID;
}

void Calendar::odl_SetUID(LONG newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetUID((LONG)newVal);
}

void Calendar::SetUID(LONG newVal)
{
	mp_lUID = newVal;
}

BSTR Calendar::odl_GetName(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetName().AllocSysString();
}

CString Calendar::GetName(void)
{
	return (CString) mp_sName;
}

void Calendar::odl_SetName(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetName(newVal);
}

void Calendar::SetName(CString newVal)
{
	if (newVal.GetLength() > 512)
	{
		newVal = newVal.Mid(0, 512);
	}
	mp_sName = newVal;
}

BOOL Calendar::odl_GetIsBaseCalendar(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetIsBaseCalendar();
}

BOOL Calendar::GetIsBaseCalendar(void)
{
	return (BOOL) mp_bIsBaseCalendar;
}

void Calendar::odl_SetIsBaseCalendar(BOOL newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetIsBaseCalendar((BOOL)newVal);
}

void Calendar::SetIsBaseCalendar(BOOL newVal)
{
	mp_bIsBaseCalendar = newVal;
}

BOOL Calendar::odl_GetIsBaselineCalendar(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetIsBaselineCalendar();
}

BOOL Calendar::GetIsBaselineCalendar(void)
{
	return (BOOL) mp_bIsBaselineCalendar;
}

void Calendar::odl_SetIsBaselineCalendar(BOOL newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetIsBaselineCalendar((BOOL)newVal);
}

void Calendar::SetIsBaselineCalendar(BOOL newVal)
{
	mp_bIsBaselineCalendar = newVal;
}

LONG Calendar::odl_GetBaseCalendarUID(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetBaseCalendarUID();
}

LONG Calendar::GetBaseCalendarUID(void)
{
	return (LONG) mp_lBaseCalendarUID;
}

void Calendar::odl_SetBaseCalendarUID(LONG newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetBaseCalendarUID((LONG)newVal);
}

void Calendar::SetBaseCalendarUID(LONG newVal)
{
	mp_lBaseCalendarUID = newVal;
}

IDispatch* Calendar::odl_GetWeekDays(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetWeekDays()->GetIDispatch(TRUE);
}

CalendarWeekDays* Calendar::GetWeekDays(void)
{
	return mp_oWeekDays;
}

IDispatch* Calendar::odl_GetExceptions(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetExceptions()->GetIDispatch(TRUE);
}

CalendarExceptions* Calendar::GetExceptions(void)
{
	return mp_oExceptions;
}

IDispatch* Calendar::odl_GetWorkWeeks(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetWorkWeeks()->GetIDispatch(TRUE);
}

CalendarWorkWeeks* Calendar::GetWorkWeeks(void)
{
	return mp_oWorkWeeks;
}

BSTR Calendar::odl_GetKey(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetKey().AllocSysString();
}

void Calendar::odl_SetKey(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	USES_CONVERSION;
	SetKey(F_TOUCSTR(T2A(newVal)));
}

CString Calendar::GetKey(void)
{
	return mp_sKey;
}

void Calendar::SetKey(CString newVal)
{
	mp_oCollection->mp_SetKey(&mp_sKey, newVal, MP_SET_KEY);
}

BSTR Calendar::odl_GetXML(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetXML().AllocSysString();
}

void Calendar::odl_SetXML(LPCTSTR sXML)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetXML(sXML);
}

BOOL Calendar::IsNull(void)
{
	BOOL bReturn = TRUE;
	if (mp_lUID != 0)
	{
		bReturn = FALSE;
	}
	if (mp_sName != "")
	{
		bReturn = FALSE;
	}
	if (mp_bIsBaseCalendar != FALSE)
	{
		bReturn = FALSE;
	}
	if (mp_bIsBaselineCalendar != FALSE)
	{
		bReturn = FALSE;
	}
	if (mp_lBaseCalendarUID != 0)
	{
		bReturn = FALSE;
	}
	if (mp_oWeekDays->IsNull() == FALSE)
	{
		bReturn = FALSE;
	}
	if (mp_oExceptions->IsNull() == FALSE)
	{
		bReturn = FALSE;
	}
	if (mp_oWorkWeeks->IsNull() == FALSE)
	{
		bReturn = FALSE;
	}
	return bReturn;
}

CString Calendar::GetXML(void)
{
	if (IsNull() == TRUE)
	{
		return "<Calendar/>";
	}
	clsXML oXML("Calendar");
	oXML.InitializeWriter();
	oXML.SetSupportOptional(TRUE);
	oXML.SetBoolsAreNumeric(TRUE);
	oXML.WriteProperty("UID", mp_lUID);
	if (mp_sName != "")
	{
		oXML.WriteProperty("Name", mp_sName);
	}
	oXML.WriteProperty("IsBaseCalendar", mp_bIsBaseCalendar);
	oXML.WriteProperty("IsBaselineCalendar", mp_bIsBaselineCalendar);
	oXML.WriteProperty("BaseCalendarUID", mp_lBaseCalendarUID);
	if (mp_oWeekDays->IsNull() == FALSE)
	{
		oXML.WriteObject(mp_oWeekDays->GetXML());
	}
	if (mp_oExceptions->IsNull() == FALSE)
	{
		oXML.WriteObject(mp_oExceptions->GetXML());
	}
	if (mp_oWorkWeeks->IsNull() == FALSE)
	{
		oXML.WriteObject(mp_oWorkWeeks->GetXML());
	}
	return oXML.GetXML();
}

void Calendar::SetXML(CString sXML)
{
	clsXML oXML("Calendar");
	oXML.SetSupportOptional(TRUE);
	oXML.SetXML(sXML);
	oXML.InitializeReader();
	oXML.ReadProperty("UID", mp_lUID);
	oXML.ReadProperty("Name", mp_sName);
	if (mp_sName.GetLength() > 512)
	{
		mp_sName = mp_sName.Mid(0, 512);
	}
	oXML.ReadProperty("IsBaseCalendar", mp_bIsBaseCalendar);
	oXML.ReadProperty("IsBaselineCalendar", mp_bIsBaselineCalendar);
	oXML.ReadProperty("BaseCalendarUID", mp_lBaseCalendarUID);
	mp_oWeekDays->SetXML(oXML.ReadObject("WeekDays"));
	mp_oExceptions->SetXML(oXML.ReadObject("Exceptions"));
	mp_oWorkWeeks->SetXML(oXML.ReadObject("WorkWeeks"));
}