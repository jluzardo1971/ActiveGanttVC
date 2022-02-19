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
#include "TimePeriod.h"

IMPLEMENT_DYNCREATE(TimePeriod, CCmdTarget)

//{93D8145F-03F0-40F6-B195-5714E00B711A}
static const IID IID_ITimePeriod = { 0x93D8145F, 0x03F0, 0x40F6, { 0xB1, 0x95, 0x57, 0x14, 0xE0, 0x0B, 0x71, 0x1A} };

//{ED9B4A3F-5C52-4AA2-A4D5-9AC3C67677F5}
IMPLEMENT_OLECREATE_FLAGS(TimePeriod, "MSP2003.TimePeriod", afxRegApartmentThreading, 0xED9B4A3F, 0x5C52, 0x4AA2, 0xA4, 0xD5, 0x9A, 0xC3, 0xC6, 0x76, 0x77, 0xF5)

BEGIN_DISPATCH_MAP(TimePeriod, CCmdTarget)
DISP_PROPERTY_EX_ID(TimePeriod, "dtFromDate", 1, odl_GetFromDate, odl_SetFromDate, VT_DATE)
DISP_PROPERTY_EX_ID(TimePeriod, "dtToDate", 2, odl_GetToDate, odl_SetToDate, VT_DATE)
DISP_PROPERTY_EX_ID(TimePeriod, "Key", 3, odl_GetKey, odl_SetKey, VT_BSTR)
DISP_FUNCTION_ID(TimePeriod, "GetXML", 4, odl_GetXML, VT_BSTR, VTS_NONE)
DISP_FUNCTION_ID(TimePeriod, "SetXML", 5, odl_SetXML, VT_EMPTY, VTS_BSTR)
DISP_FUNCTION_ID(TimePeriod, "IsNull", 6, IsNull, VT_BOOL, VTS_NONE)
DISP_FUNCTION_ID(TimePeriod, "Initialize", 7, Initialize, VT_EMPTY, VTS_NONE)
END_DISPATCH_MAP()

BEGIN_INTERFACE_MAP(TimePeriod, CCmdTarget)
	INTERFACE_PART(TimePeriod, IID_ITimePeriod, Dispatch)
END_INTERFACE_MAP()

BEGIN_MESSAGE_MAP(TimePeriod, CCmdTarget)
END_MESSAGE_MAP()


void TimePeriod::Initialize(void)
{
	InitVars();
}

TimePeriod::TimePeriod()
{
	EnableAutomation();
	AfxOleLockApp();
	InitVars();
}

void TimePeriod::InitVars(void)
{
	mp_dtFromDate = (DATE) 0;
	mp_dtToDate = (DATE) 0;
}

TimePeriod::~TimePeriod()
{
	AfxOleUnlockApp();
}

void TimePeriod::OnFinalRelease()
{
	clsItemBase::OnFinalRelease();
}

DATE TimePeriod::odl_GetFromDate(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetFromDate();
}

COleDateTime TimePeriod::GetFromDate(void)
{
	return (COleDateTime) mp_dtFromDate;
}

void TimePeriod::odl_SetFromDate(DATE newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetFromDate((COleDateTime)newVal);
}

void TimePeriod::SetFromDate(COleDateTime newVal)
{
	mp_dtFromDate = newVal;
}

DATE TimePeriod::odl_GetToDate(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetToDate();
}

COleDateTime TimePeriod::GetToDate(void)
{
	return (COleDateTime) mp_dtToDate;
}

void TimePeriod::odl_SetToDate(DATE newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetToDate((COleDateTime)newVal);
}

void TimePeriod::SetToDate(COleDateTime newVal)
{
	mp_dtToDate = newVal;
}

BSTR TimePeriod::odl_GetKey(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetKey().AllocSysString();
}

void TimePeriod::odl_SetKey(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	USES_CONVERSION;
	SetKey(F_TOUCSTR(T2A(newVal)));
}

CString TimePeriod::GetKey(void)
{
	return mp_sKey;
}

void TimePeriod::SetKey(CString newVal)
{
	mp_oCollection->mp_SetKey(&mp_sKey, newVal, MP_SET_KEY);
}

BSTR TimePeriod::odl_GetXML(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetXML().AllocSysString();
}

void TimePeriod::odl_SetXML(LPCTSTR sXML)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetXML(sXML);
}

BOOL TimePeriod::IsNull(void)
{
	BOOL bReturn = TRUE;
	if (mp_dtFromDate.m_dt != 36494)
	{
		bReturn = FALSE;
	}
	if (mp_dtToDate.m_dt != 36494)
	{
		bReturn = FALSE;
	}
	return bReturn;
}

CString TimePeriod::GetXML(void)
{
	if (IsNull() == TRUE)
	{
		return "<TimePeriod/>";
	}
	clsXML oXML("TimePeriod");
	oXML.InitializeWriter();
	oXML.SetSupportOptional(TRUE);
	oXML.SetBoolsAreNumeric(TRUE);
	if (mp_dtFromDate.m_dt != 36494)
	{
		oXML.WriteProperty("FromDate", mp_dtFromDate);
	}
	if (mp_dtToDate.m_dt != 36494)
	{
		oXML.WriteProperty("ToDate", mp_dtToDate);
	}
	return oXML.GetXML();
}

void TimePeriod::SetXML(CString sXML)
{
	clsXML oXML("TimePeriod");
	oXML.SetSupportOptional(TRUE);
	oXML.SetXML(sXML);
	oXML.InitializeReader();
	oXML.ReadProperty("FromDate", mp_dtFromDate);
	oXML.ReadProperty("ToDate", mp_dtToDate);
}