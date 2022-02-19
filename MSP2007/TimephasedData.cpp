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
#include "TimephasedData.h"

IMPLEMENT_DYNCREATE(TimephasedData, CCmdTarget)

//{14E45576-519F-41CD-923F-9CD89466BB5A}
static const IID IID_ITimephasedData = { 0x14E45576, 0x519F, 0x41CD, { 0x92, 0x3F, 0x9C, 0xD8, 0x94, 0x66, 0xBB, 0x5A} };

//{D9FF5819-FAFD-47AE-AD90-516AE230447E}
IMPLEMENT_OLECREATE_FLAGS(TimephasedData, "MSP2007.TimephasedData", afxRegApartmentThreading, 0xD9FF5819, 0xFAFD, 0x47AE, 0xAD, 0x90, 0x51, 0x6A, 0xE2, 0x30, 0x44, 0x7E)

BEGIN_DISPATCH_MAP(TimephasedData, CCmdTarget)
DISP_PROPERTY_EX_ID(TimephasedData, "yType", 1, odl_GetType, odl_SetType, VT_I4)
DISP_PROPERTY_EX_ID(TimephasedData, "lUID", 2, odl_GetUID, odl_SetUID, VT_I4)
DISP_PROPERTY_EX_ID(TimephasedData, "dtStart", 3, odl_GetStart, odl_SetStart, VT_DATE)
DISP_PROPERTY_EX_ID(TimephasedData, "dtFinish", 4, odl_GetFinish, odl_SetFinish, VT_DATE)
DISP_PROPERTY_EX_ID(TimephasedData, "yUnit", 5, odl_GetUnit, odl_SetUnit, VT_I4)
DISP_PROPERTY_EX_ID(TimephasedData, "sValue", 6, odl_GetValue, odl_SetValue, VT_BSTR)
DISP_PROPERTY_EX_ID(TimephasedData, "Key", 7, odl_GetKey, odl_SetKey, VT_BSTR)
DISP_FUNCTION_ID(TimephasedData, "GetXML", 8, odl_GetXML, VT_BSTR, VTS_NONE)
DISP_FUNCTION_ID(TimephasedData, "SetXML", 9, odl_SetXML, VT_EMPTY, VTS_BSTR)
DISP_FUNCTION_ID(TimephasedData, "IsNull", 10, IsNull, VT_BOOL, VTS_NONE)
DISP_FUNCTION_ID(TimephasedData, "Initialize", 11, Initialize, VT_EMPTY, VTS_NONE)
END_DISPATCH_MAP()

BEGIN_INTERFACE_MAP(TimephasedData, CCmdTarget)
	INTERFACE_PART(TimephasedData, IID_ITimephasedData, Dispatch)
END_INTERFACE_MAP()

BEGIN_MESSAGE_MAP(TimephasedData, CCmdTarget)
END_MESSAGE_MAP()


void TimephasedData::Initialize(void)
{
	InitVars();
}

TimephasedData::TimephasedData()
{
	EnableAutomation();
	AfxOleLockApp();
	InitVars();
}

void TimephasedData::InitVars(void)
{
	mp_yType = T_7_ASSIGNMENT_REMAINING_WORK;
	mp_lUID = 0;
	mp_dtStart = (DATE) 0;
	mp_dtFinish = (DATE) 0;
	mp_yUnit = U_M;
	mp_sValue = "";
}

TimephasedData::~TimephasedData()
{
	AfxOleUnlockApp();
}

void TimephasedData::OnFinalRelease()
{
	clsItemBase::OnFinalRelease();
}

LONG TimephasedData::odl_GetType(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetType();
}

E_TYPE_7 TimephasedData::GetType(void)
{
	return (E_TYPE_7) mp_yType;
}

void TimephasedData::odl_SetType(LONG newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetType((E_TYPE_7)newVal);
}

void TimephasedData::SetType(E_TYPE_7 newVal)
{
	mp_yType = newVal;
}

LONG TimephasedData::odl_GetUID(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetUID();
}

LONG TimephasedData::GetUID(void)
{
	return (LONG) mp_lUID;
}

void TimephasedData::odl_SetUID(LONG newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetUID((LONG)newVal);
}

void TimephasedData::SetUID(LONG newVal)
{
	mp_lUID = newVal;
}

DATE TimephasedData::odl_GetStart(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetStart();
}

COleDateTime TimephasedData::GetStart(void)
{
	return (COleDateTime) mp_dtStart;
}

void TimephasedData::odl_SetStart(DATE newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetStart((COleDateTime)newVal);
}

void TimephasedData::SetStart(COleDateTime newVal)
{
	mp_dtStart = newVal;
}

DATE TimephasedData::odl_GetFinish(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetFinish();
}

COleDateTime TimephasedData::GetFinish(void)
{
	return (COleDateTime) mp_dtFinish;
}

void TimephasedData::odl_SetFinish(DATE newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetFinish((COleDateTime)newVal);
}

void TimephasedData::SetFinish(COleDateTime newVal)
{
	mp_dtFinish = newVal;
}

LONG TimephasedData::odl_GetUnit(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetUnit();
}

E_UNIT TimephasedData::GetUnit(void)
{
	return (E_UNIT) mp_yUnit;
}

void TimephasedData::odl_SetUnit(LONG newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetUnit((E_UNIT)newVal);
}

void TimephasedData::SetUnit(E_UNIT newVal)
{
	mp_yUnit = newVal;
}

BSTR TimephasedData::odl_GetValue(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetValue().AllocSysString();
}

CString TimephasedData::GetValue(void)
{
	return (CString) mp_sValue;
}

void TimephasedData::odl_SetValue(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetValue(newVal);
}

void TimephasedData::SetValue(CString newVal)
{
	mp_sValue = newVal;
}

BSTR TimephasedData::odl_GetKey(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetKey().AllocSysString();
}

void TimephasedData::odl_SetKey(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	USES_CONVERSION;
	SetKey(F_TOUCSTR(T2A(newVal)));
}

CString TimephasedData::GetKey(void)
{
	return mp_sKey;
}

void TimephasedData::SetKey(CString newVal)
{
	mp_oCollection->mp_SetKey(&mp_sKey, newVal, MP_SET_KEY);
}

BSTR TimephasedData::odl_GetXML(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetXML().AllocSysString();
}

void TimephasedData::odl_SetXML(LPCTSTR sXML)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetXML(sXML);
}

BOOL TimephasedData::IsNull(void)
{
	BOOL bReturn = TRUE;
	if (mp_yType != T_7_ASSIGNMENT_REMAINING_WORK)
	{
		bReturn = FALSE;
	}
	if (mp_lUID != 0)
	{
		bReturn = FALSE;
	}
	if (mp_dtStart.m_dt != 36494)
	{
		bReturn = FALSE;
	}
	if (mp_dtFinish.m_dt != 36494)
	{
		bReturn = FALSE;
	}
	if (mp_yUnit != U_M)
	{
		bReturn = FALSE;
	}
	if (mp_sValue != "")
	{
		bReturn = FALSE;
	}
	return bReturn;
}

CString TimephasedData::GetXML(void)
{
	if (IsNull() == TRUE)
	{
		return "<TimephasedData/>";
	}
	clsXML oXML("TimephasedData");
	oXML.InitializeWriter();
	oXML.SetSupportOptional(TRUE);
	oXML.SetBoolsAreNumeric(TRUE);
	oXML.WriteProperty("Type", mp_yType);
	oXML.WriteProperty("UID", mp_lUID);
	if (mp_dtStart.m_dt != 36494)
	{
		oXML.WriteProperty("Start", mp_dtStart);
	}
	if (mp_dtFinish.m_dt != 36494)
	{
		oXML.WriteProperty("Finish", mp_dtFinish);
	}
	oXML.WriteProperty("Unit", mp_yUnit);
	if (mp_sValue != "")
	{
		oXML.WriteProperty("Value", mp_sValue);
	}
	return oXML.GetXML();
}

void TimephasedData::SetXML(CString sXML)
{
	clsXML oXML("TimephasedData");
	oXML.SetSupportOptional(TRUE);
	oXML.SetXML(sXML);
	oXML.InitializeReader();
	oXML.ReadProperty("Type", mp_yType);
	oXML.ReadProperty("UID", mp_lUID);
	oXML.ReadProperty("Start", mp_dtStart);
	oXML.ReadProperty("Finish", mp_dtFinish);
	oXML.ReadProperty("Unit", mp_yUnit);
	oXML.ReadProperty("Value", mp_sValue);
}