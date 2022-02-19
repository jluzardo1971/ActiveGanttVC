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

//{3F7BA988-7D7C-4AF3-B8E7-784E4CE827A9}
static const IID IID_IWeekDay = { 0x3F7BA988, 0x7D7C, 0x4AF3, { 0xB8, 0xE7, 0x78, 0x4E, 0x4C, 0xE8, 0x27, 0xA9} };

//{99E5D75F-19EA-4CC3-AE5E-EBBA54EFEC4F}
IMPLEMENT_OLECREATE_FLAGS(WeekDay, "MSP2007.WeekDay", afxRegApartmentThreading, 0x99E5D75F, 0x19EA, 0x4CC3, 0xAE, 0x5E, 0xEB, 0xBA, 0x54, 0xEF, 0xEC, 0x4F)

BEGIN_DISPATCH_MAP(WeekDay, CCmdTarget)
DISP_PROPERTY_EX_ID(WeekDay, "yDayType", 1, odl_GetDayType, odl_SetDayType, VT_I4)
DISP_PROPERTY_EX_ID(WeekDay, "bDayWorking", 2, odl_GetDayWorking, odl_SetDayWorking, VT_BOOL)
DISP_PROPERTY_EX_ID(WeekDay, "Key", 3, odl_GetKey, odl_SetKey, VT_BSTR)
DISP_FUNCTION_ID(WeekDay, "GetXML", 4, odl_GetXML, VT_BSTR, VTS_NONE)
DISP_FUNCTION_ID(WeekDay, "SetXML", 5, odl_SetXML, VT_EMPTY, VTS_BSTR)
DISP_FUNCTION_ID(WeekDay, "IsNull", 6, IsNull, VT_BOOL, VTS_NONE)
DISP_FUNCTION_ID(WeekDay, "Initialize", 7, Initialize, VT_EMPTY, VTS_NONE)
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
}

WeekDay::~WeekDay()
{
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
}