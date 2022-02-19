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
#include "WorkingTime.h"

IMPLEMENT_DYNCREATE(WorkingTime, CCmdTarget)

//{F084F09B-94CC-4BEE-B9CF-D9D78D394F0A}
static const IID IID_IWorkingTime = { 0xF084F09B, 0x94CC, 0x4BEE, { 0xB9, 0xCF, 0xD9, 0xD7, 0x8D, 0x39, 0x4F, 0x0A} };

//{81C6218D-101C-4F09-B780-E857F433395B}
IMPLEMENT_OLECREATE_FLAGS(WorkingTime, "MSP2007.WorkingTime", afxRegApartmentThreading, 0x81C6218D, 0x101C, 0x4F09, 0xB7, 0x80, 0xE8, 0x57, 0xF4, 0x33, 0x39, 0x5B)

BEGIN_DISPATCH_MAP(WorkingTime, CCmdTarget)
DISP_PROPERTY_EX_ID(WorkingTime, "oFromTime", 1, odl_GetFromTime, SetNotSupported, VT_DISPATCH)
DISP_PROPERTY_EX_ID(WorkingTime, "oToTime", 2, odl_GetToTime, SetNotSupported, VT_DISPATCH)
DISP_PROPERTY_EX_ID(WorkingTime, "Key", 3, odl_GetKey, odl_SetKey, VT_BSTR)
DISP_FUNCTION_ID(WorkingTime, "GetXML", 4, odl_GetXML, VT_BSTR, VTS_NONE)
DISP_FUNCTION_ID(WorkingTime, "SetXML", 5, odl_SetXML, VT_EMPTY, VTS_BSTR)
DISP_FUNCTION_ID(WorkingTime, "IsNull", 6, IsNull, VT_BOOL, VTS_NONE)
DISP_FUNCTION_ID(WorkingTime, "Initialize", 7, Initialize, VT_EMPTY, VTS_NONE)
END_DISPATCH_MAP()

BEGIN_INTERFACE_MAP(WorkingTime, CCmdTarget)
	INTERFACE_PART(WorkingTime, IID_IWorkingTime, Dispatch)
END_INTERFACE_MAP()

BEGIN_MESSAGE_MAP(WorkingTime, CCmdTarget)
END_MESSAGE_MAP()


void WorkingTime::Initialize(void)
{
	InitVars();
}

WorkingTime::WorkingTime()
{
	EnableAutomation();
	AfxOleLockApp();
	InitVars();
}

void WorkingTime::InitVars(void)
{
	mp_oFromTime = new Time();
	mp_oToTime = new Time();
}

WorkingTime::~WorkingTime()
{
	delete mp_oFromTime;
	delete mp_oToTime;
	AfxOleUnlockApp();
}

void WorkingTime::OnFinalRelease()
{
	clsItemBase::OnFinalRelease();
}

IDispatch* WorkingTime::odl_GetFromTime(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetFromTime()->GetIDispatch(TRUE);
}

Time* WorkingTime::GetFromTime(void)
{
	return mp_oFromTime;
}

IDispatch* WorkingTime::odl_GetToTime(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetToTime()->GetIDispatch(TRUE);
}

Time* WorkingTime::GetToTime(void)
{
	return mp_oToTime;
}

BSTR WorkingTime::odl_GetKey(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetKey().AllocSysString();
}

void WorkingTime::odl_SetKey(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	USES_CONVERSION;
	SetKey(F_TOUCSTR(T2A(newVal)));
}

CString WorkingTime::GetKey(void)
{
	return mp_sKey;
}

void WorkingTime::SetKey(CString newVal)
{
	mp_oCollection->mp_SetKey(&mp_sKey, newVal, MP_SET_KEY);
}

BSTR WorkingTime::odl_GetXML(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetXML().AllocSysString();
}

void WorkingTime::odl_SetXML(LPCTSTR sXML)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetXML(sXML);
}

BOOL WorkingTime::IsNull(void)
{
	BOOL bReturn = TRUE;
	if (mp_oFromTime->IsNull() == FALSE)
	{
		bReturn = FALSE;
	}
	if (mp_oToTime->IsNull() == FALSE)
	{
		bReturn = FALSE;
	}
	return bReturn;
}

CString WorkingTime::GetXML(void)
{
	if (IsNull() == TRUE)
	{
		return "<WorkingTime/>";
	}
	clsXML oXML("WorkingTime");
	oXML.InitializeWriter();
	oXML.SetSupportOptional(TRUE);
	oXML.SetBoolsAreNumeric(TRUE);
	if (mp_oFromTime->IsNull() == FALSE)
	{
		oXML.WriteProperty("FromTime", mp_oFromTime);
	}
	if (mp_oToTime->IsNull() == FALSE)
	{
		oXML.WriteProperty("ToTime", mp_oToTime);
	}
	return oXML.GetXML();
}

void WorkingTime::SetXML(CString sXML)
{
	clsXML oXML("WorkingTime");
	oXML.SetSupportOptional(TRUE);
	oXML.SetXML(sXML);
	oXML.InitializeReader();
	oXML.ReadProperty("FromTime", mp_oFromTime);
	oXML.ReadProperty("ToTime", mp_oToTime);
}