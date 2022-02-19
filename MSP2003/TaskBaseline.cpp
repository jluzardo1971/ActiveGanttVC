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
#include "TaskBaseline.h"

IMPLEMENT_DYNCREATE(TaskBaseline, CCmdTarget)

//{FB35FE98-95D2-414C-B4DD-FD13CB2397BD}
static const IID IID_ITaskBaseline = { 0xFB35FE98, 0x95D2, 0x414C, { 0xB4, 0xDD, 0xFD, 0x13, 0xCB, 0x23, 0x97, 0xBD} };

//{DB8ACF1A-21E3-4B5A-9270-6F5DE399D412}
IMPLEMENT_OLECREATE_FLAGS(TaskBaseline, "MSP2003.TaskBaseline", afxRegApartmentThreading, 0xDB8ACF1A, 0x21E3, 0x4B5A, 0x92, 0x70, 0x6F, 0x5D, 0xE3, 0x99, 0xD4, 0x12)

BEGIN_DISPATCH_MAP(TaskBaseline, CCmdTarget)
DISP_PROPERTY_EX_ID(TaskBaseline, "oTimephasedData", 1, odl_GetTimephasedData, SetNotSupported, VT_DISPATCH)
DISP_PROPERTY_EX_ID(TaskBaseline, "lNumber", 2, odl_GetNumber, odl_SetNumber, VT_I4)
DISP_PROPERTY_EX_ID(TaskBaseline, "bInterim", 3, odl_GetInterim, odl_SetInterim, VT_BOOL)
DISP_PROPERTY_EX_ID(TaskBaseline, "dtStart", 4, odl_GetStart, odl_SetStart, VT_DATE)
DISP_PROPERTY_EX_ID(TaskBaseline, "dtFinish", 5, odl_GetFinish, odl_SetFinish, VT_DATE)
DISP_PROPERTY_EX_ID(TaskBaseline, "oDuration", 6, odl_GetDuration, SetNotSupported, VT_DISPATCH)
DISP_PROPERTY_EX_ID(TaskBaseline, "yDurationFormat", 7, odl_GetDurationFormat, odl_SetDurationFormat, VT_I4)
DISP_PROPERTY_EX_ID(TaskBaseline, "bEstimatedDuration", 8, odl_GetEstimatedDuration, odl_SetEstimatedDuration, VT_BOOL)
DISP_PROPERTY_EX_ID(TaskBaseline, "oWork", 9, odl_GetWork, SetNotSupported, VT_DISPATCH)
DISP_PROPERTY_EX_ID(TaskBaseline, "cCost", 10, odl_GetCost, odl_SetCost, VT_R8)
DISP_PROPERTY_EX_ID(TaskBaseline, "fBCWS", 11, odl_GetBCWS, odl_SetBCWS, VT_R4)
DISP_PROPERTY_EX_ID(TaskBaseline, "fBCWP", 12, odl_GetBCWP, odl_SetBCWP, VT_R4)
DISP_PROPERTY_EX_ID(TaskBaseline, "Key", 13, odl_GetKey, odl_SetKey, VT_BSTR)
DISP_FUNCTION_ID(TaskBaseline, "GetXML", 14, odl_GetXML, VT_BSTR, VTS_NONE)
DISP_FUNCTION_ID(TaskBaseline, "SetXML", 15, odl_SetXML, VT_EMPTY, VTS_BSTR)
DISP_FUNCTION_ID(TaskBaseline, "IsNull", 16, IsNull, VT_BOOL, VTS_NONE)
DISP_FUNCTION_ID(TaskBaseline, "Initialize", 17, Initialize, VT_EMPTY, VTS_NONE)
END_DISPATCH_MAP()

BEGIN_INTERFACE_MAP(TaskBaseline, CCmdTarget)
	INTERFACE_PART(TaskBaseline, IID_ITaskBaseline, Dispatch)
END_INTERFACE_MAP()

BEGIN_MESSAGE_MAP(TaskBaseline, CCmdTarget)
END_MESSAGE_MAP()


void TaskBaseline::Initialize(void)
{
	InitVars();
}

TaskBaseline::TaskBaseline()
{
	EnableAutomation();
	AfxOleLockApp();
	InitVars();
}

void TaskBaseline::InitVars(void)
{
	mp_oTimephasedData_C = new TimephasedData_C();
	mp_lNumber = 0;
	mp_bInterim = FALSE;
	mp_dtStart = (DATE) 0;
	mp_dtFinish = (DATE) 0;
	mp_oDuration = new Duration();
	mp_yDurationFormat = DF_M;
	mp_bEstimatedDuration = FALSE;
	mp_oWork = new Duration();
	mp_cCost = 0;
	mp_fBCWS = 0;
	mp_fBCWP = 0;
}

TaskBaseline::~TaskBaseline()
{
	delete mp_oTimephasedData_C;
	delete mp_oDuration;
	delete mp_oWork;
	AfxOleUnlockApp();
}

void TaskBaseline::OnFinalRelease()
{
	clsItemBase::OnFinalRelease();
}

IDispatch* TaskBaseline::odl_GetTimephasedData(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetTimephasedData()->GetIDispatch(TRUE);
}

TimephasedData_C* TaskBaseline::GetTimephasedData(void)
{
	return mp_oTimephasedData_C;
}

LONG TaskBaseline::odl_GetNumber(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetNumber();
}

LONG TaskBaseline::GetNumber(void)
{
	return (LONG) mp_lNumber;
}

void TaskBaseline::odl_SetNumber(LONG newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetNumber((LONG)newVal);
}

void TaskBaseline::SetNumber(LONG newVal)
{
	mp_lNumber = newVal;
}

BOOL TaskBaseline::odl_GetInterim(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetInterim();
}

BOOL TaskBaseline::GetInterim(void)
{
	return (BOOL) mp_bInterim;
}

void TaskBaseline::odl_SetInterim(BOOL newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetInterim((BOOL)newVal);
}

void TaskBaseline::SetInterim(BOOL newVal)
{
	mp_bInterim = newVal;
}

DATE TaskBaseline::odl_GetStart(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetStart();
}

COleDateTime TaskBaseline::GetStart(void)
{
	return (COleDateTime) mp_dtStart;
}

void TaskBaseline::odl_SetStart(DATE newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetStart((COleDateTime)newVal);
}

void TaskBaseline::SetStart(COleDateTime newVal)
{
	mp_dtStart = newVal;
}

DATE TaskBaseline::odl_GetFinish(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetFinish();
}

COleDateTime TaskBaseline::GetFinish(void)
{
	return (COleDateTime) mp_dtFinish;
}

void TaskBaseline::odl_SetFinish(DATE newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetFinish((COleDateTime)newVal);
}

void TaskBaseline::SetFinish(COleDateTime newVal)
{
	mp_dtFinish = newVal;
}

IDispatch* TaskBaseline::odl_GetDuration(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetDuration()->GetIDispatch(TRUE);
}

Duration* TaskBaseline::GetDuration(void)
{
	return mp_oDuration;
}

LONG TaskBaseline::odl_GetDurationFormat(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetDurationFormat();
}

E_DURATIONFORMAT TaskBaseline::GetDurationFormat(void)
{
	return (E_DURATIONFORMAT) mp_yDurationFormat;
}

void TaskBaseline::odl_SetDurationFormat(LONG newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetDurationFormat((E_DURATIONFORMAT)newVal);
}

void TaskBaseline::SetDurationFormat(E_DURATIONFORMAT newVal)
{
	mp_yDurationFormat = newVal;
}

BOOL TaskBaseline::odl_GetEstimatedDuration(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetEstimatedDuration();
}

BOOL TaskBaseline::GetEstimatedDuration(void)
{
	return (BOOL) mp_bEstimatedDuration;
}

void TaskBaseline::odl_SetEstimatedDuration(BOOL newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetEstimatedDuration((BOOL)newVal);
}

void TaskBaseline::SetEstimatedDuration(BOOL newVal)
{
	mp_bEstimatedDuration = newVal;
}

IDispatch* TaskBaseline::odl_GetWork(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetWork()->GetIDispatch(TRUE);
}

Duration* TaskBaseline::GetWork(void)
{
	return mp_oWork;
}

DOUBLE TaskBaseline::odl_GetCost(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetCost();
}

DOUBLE TaskBaseline::GetCost(void)
{
	return (DOUBLE) mp_cCost;
}

void TaskBaseline::odl_SetCost(DOUBLE newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetCost((DOUBLE)newVal);
}

void TaskBaseline::SetCost(DOUBLE newVal)
{
	mp_cCost = newVal;
}

FLOAT TaskBaseline::odl_GetBCWS(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetBCWS();
}

FLOAT TaskBaseline::GetBCWS(void)
{
	return (FLOAT) mp_fBCWS;
}

void TaskBaseline::odl_SetBCWS(FLOAT newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetBCWS((FLOAT)newVal);
}

void TaskBaseline::SetBCWS(FLOAT newVal)
{
	mp_fBCWS = newVal;
}

FLOAT TaskBaseline::odl_GetBCWP(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetBCWP();
}

FLOAT TaskBaseline::GetBCWP(void)
{
	return (FLOAT) mp_fBCWP;
}

void TaskBaseline::odl_SetBCWP(FLOAT newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetBCWP((FLOAT)newVal);
}

void TaskBaseline::SetBCWP(FLOAT newVal)
{
	mp_fBCWP = newVal;
}

BSTR TaskBaseline::odl_GetKey(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetKey().AllocSysString();
}

void TaskBaseline::odl_SetKey(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	USES_CONVERSION;
	SetKey(F_TOUCSTR(T2A(newVal)));
}

CString TaskBaseline::GetKey(void)
{
	return mp_sKey;
}

void TaskBaseline::SetKey(CString newVal)
{
	mp_oCollection->mp_SetKey(&mp_sKey, newVal, MP_SET_KEY);
}

BSTR TaskBaseline::odl_GetXML(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetXML().AllocSysString();
}

void TaskBaseline::odl_SetXML(LPCTSTR sXML)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetXML(sXML);
}

BOOL TaskBaseline::IsNull(void)
{
	BOOL bReturn = TRUE;
	if (mp_oTimephasedData_C->IsNull() == FALSE)
	{
		bReturn = FALSE;
	}
	if (mp_lNumber != 0)
	{
		bReturn = FALSE;
	}
	if (mp_bInterim != FALSE)
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
	if (mp_oDuration->IsNull() == FALSE)
	{
		bReturn = FALSE;
	}
	if (mp_yDurationFormat != DF_M)
	{
		bReturn = FALSE;
	}
	if (mp_bEstimatedDuration != FALSE)
	{
		bReturn = FALSE;
	}
	if (mp_oWork->IsNull() == FALSE)
	{
		bReturn = FALSE;
	}
	if (mp_cCost != 0)
	{
		bReturn = FALSE;
	}
	if (mp_fBCWS != 0)
	{
		bReturn = FALSE;
	}
	if (mp_fBCWP != 0)
	{
		bReturn = FALSE;
	}
	return bReturn;
}

CString TaskBaseline::GetXML(void)
{
	if (IsNull() == TRUE)
	{
		return "<Baseline/>";
	}
	clsXML oXML("Baseline");
	oXML.InitializeWriter();
	oXML.SetSupportOptional(TRUE);
	oXML.SetBoolsAreNumeric(TRUE);
	if (mp_oTimephasedData_C->IsNull() == FALSE)
	{
		mp_oTimephasedData_C->WriteObjectProtected(oXML);
	}
	oXML.WriteProperty("Number", mp_lNumber);
	oXML.WriteProperty("Interim", mp_bInterim);
	if (mp_dtStart.m_dt != 36494)
	{
		oXML.WriteProperty("Start", mp_dtStart);
	}
	if (mp_dtFinish.m_dt != 36494)
	{
		oXML.WriteProperty("Finish", mp_dtFinish);
	}
	oXML.WriteProperty("Duration", mp_oDuration);
	oXML.WriteProperty("DurationFormat", mp_yDurationFormat);
	oXML.WriteProperty("EstimatedDuration", mp_bEstimatedDuration);
	oXML.WriteProperty("Work", mp_oWork);
	oXML.WriteProperty("Cost", mp_cCost);
	oXML.WriteProperty("BCWS", mp_fBCWS);
	oXML.WriteProperty("BCWP", mp_fBCWP);
	return oXML.GetXML();
}

void TaskBaseline::SetXML(CString sXML)
{
	clsXML oXML("Baseline");
	oXML.SetSupportOptional(TRUE);
	oXML.SetXML(sXML);
	oXML.InitializeReader();
	mp_oTimephasedData_C->ReadObjectProtected(oXML);
	oXML.ReadProperty("Number", mp_lNumber);
	oXML.ReadProperty("Interim", mp_bInterim);
	oXML.ReadProperty("Start", mp_dtStart);
	oXML.ReadProperty("Finish", mp_dtFinish);
	oXML.ReadProperty("Duration", mp_oDuration);
	oXML.ReadProperty("DurationFormat", mp_yDurationFormat);
	oXML.ReadProperty("EstimatedDuration", mp_bEstimatedDuration);
	oXML.ReadProperty("Work", mp_oWork);
	oXML.ReadProperty("Cost", mp_cCost);
	oXML.ReadProperty("BCWS", mp_fBCWS);
	oXML.ReadProperty("BCWP", mp_fBCWP);
}