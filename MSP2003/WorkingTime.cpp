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

//{1F25A838-1764-42B4-9621-FAAE717C120E}
static const IID IID_IWorkingTime = { 0x1F25A838, 0x1764, 0x42B4, { 0x96, 0x21, 0xFA, 0xAE, 0x71, 0x7C, 0x12, 0x0E} };

//{A2336F7B-99A1-4610-BF8D-EF0C820E0195}
IMPLEMENT_OLECREATE_FLAGS(WorkingTime, "MSP2003.WorkingTime", afxRegApartmentThreading, 0xA2336F7B, 0x99A1, 0x4610, 0xBF, 0x8D, 0xEF, 0x0C, 0x82, 0x0E, 0x01, 0x95)

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