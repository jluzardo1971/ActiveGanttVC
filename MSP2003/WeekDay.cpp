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
#include "WeekDay.h"

IMPLEMENT_DYNCREATE(WeekDay, CCmdTarget)

//{5ACB2190-665B-4A0B-8E8D-F1ED438F5A46}
static const IID IID_IWeekDay = { 0x5ACB2190, 0x665B, 0x4A0B, { 0x8E, 0x8D, 0xF1, 0xED, 0x43, 0x8F, 0x5A, 0x46} };

//{2AACA94C-4442-42B3-9F5F-17C643A0D15D}
IMPLEMENT_OLECREATE_FLAGS(WeekDay, "MSP2003.WeekDay", afxRegApartmentThreading, 0x2AACA94C, 0x4442, 0x42B3, 0x9F, 0x5F, 0x17, 0xC6, 0x43, 0xA0, 0xD1, 0x5D)

BEGIN_DISPATCH_MAP(WeekDay, CCmdTarget)
DISP_PROPERTY_EX_ID(WeekDay, "yDayType", 1, odl_GetDayType, odl_SetDayType, VT_I4)
DISP_PROPERTY_EX_ID(WeekDay, "bDayWorking", 2, odl_GetDayWorking, odl_SetDayWorking, VT_BOOL)
DISP_PROPERTY_EX_ID(WeekDay, "oTimePeriod", 3, odl_GetTimePeriod, SetNotSupported, VT_DISPATCH)
DISP_PROPERTY_EX_ID(WeekDay, "oWorkingTimes", 4, odl_GetWorkingTimes, SetNotSupported, VT_DISPATCH)
DISP_PROPERTY_EX_ID(WeekDay, "Key", 5, odl_GetKey, odl_SetKey, VT_BSTR)
DISP_FUNCTION_ID(WeekDay, "GetXML", 6, odl_GetXML, VT_BSTR, VTS_NONE)
DISP_FUNCTION_ID(WeekDay, "SetXML", 7, odl_SetXML, VT_EMPTY, VTS_BSTR)
DISP_FUNCTION_ID(WeekDay, "IsNull", 8, IsNull, VT_BOOL, VTS_NONE)
DISP_FUNCTION_ID(WeekDay, "Initialize", 9, Initialize, VT_EMPTY, VTS_NONE)
END_DISPATCH_MAP()

BEGIN_INTERFACE_MAP(WeekDay, CCmdTarget)
	INTERFACE_PART(WeekDay, IID_IWeekDay, Dispatch)
END_INTERFACE_MAP()

BEGIN_MESSAGE_MAP(WeekDay, CCmdTarget)
END_MESSAGE_MAP()


void WeekDay::Initialize(void)
{
	InitVars();
}

WeekDay::WeekDay()
{
	EnableAutomation();
	AfxOleLockApp();
	InitVars();
}

void WeekDay::InitVars(void)
{
	mp_yDayType = DT_EXCEPTION;
	mp_bDayWorking = FALSE;
	mp_oTimePeriod = new TimePeriod();
	mp_oWorkingTimes = new WorkingTimes();
}

WeekDay::~WeekDay()
{
	delete mp_oTimePeriod;
	delete mp_oWorkingTimes;
	AfxOleUnlockApp();
}

void WeekDay::OnFinalRelease()
{
	clsItemBase::OnFinalRelease();
}

LONG WeekDay::odl_GetDayType(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetDayType();
}

E_DAYTYPE WeekDay::GetDayType(void)
{
	return (E_DAYTYPE) mp_yDayType;
}

void WeekDay::odl_SetDayType(LONG newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetDayType((E_DAYTYPE)newVal);
}

void WeekDay::SetDayType(E_DAYTYPE newVal)
{
	mp_yDayType = newVal;
}

BOOL WeekDay::odl_GetDayWorking(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetDayWorking();
}

BOOL WeekDay::GetDayWorking(void)
{
	return (BOOL) mp_bDayWorking;
}

void WeekDay::odl_SetDayWorking(BOOL newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetDayWorking((BOOL)newVal);
}

void WeekDay::SetDayWorking(BOOL newVal)
{
	mp_bDayWorking = newVal;
}

IDispatch* WeekDay::odl_GetTimePeriod(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetTimePeriod()->GetIDispatch(TRUE);
}

TimePeriod* WeekDay::GetTimePeriod(void)
{
	return mp_oTimePeriod;
}

IDispatch* WeekDay::odl_GetWorkingTimes(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetWorkingTimes()->GetIDispatch(TRUE);
}

WorkingTimes* WeekDay::GetWorkingTimes(void)
{
	return mp_oWorkingTimes;
}

BSTR WeekDay::odl_GetKey(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetKey().AllocSysString();
}

void WeekDay::odl_SetKey(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	USES_CONVERSION;
	SetKey(F_TOUCSTR(T2A(newVal)));
}

CString WeekDay::GetKey(void)
{
	return mp_sKey;
}

void WeekDay::SetKey(CString newVal)
{
	mp_oCollection->mp_SetKey(&mp_sKey, newVal, MP_SET_KEY);
}

BSTR WeekDay::odl_GetXML(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetXML().AllocSysString();
}

void WeekDay::odl_SetXML(LPCTSTR sXML)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetXML(sXML);
}

BOOL WeekDay::IsNull(void)
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

CString WeekDay::GetXML(void)
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

void WeekDay::SetXML(CString sXML)
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