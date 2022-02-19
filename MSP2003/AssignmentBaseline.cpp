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
#include "AssignmentBaseline.h"

IMPLEMENT_DYNCREATE(AssignmentBaseline, CCmdTarget)

//{8FEB05D7-0B8C-4463-A0C8-80E363A58FBA}
static const IID IID_IAssignmentBaseline = { 0x8FEB05D7, 0x0B8C, 0x4463, { 0xA0, 0xC8, 0x80, 0xE3, 0x63, 0xA5, 0x8F, 0xBA} };

//{50D939D1-FC1D-4D59-B586-2C78147726B6}
IMPLEMENT_OLECREATE_FLAGS(AssignmentBaseline, "MSP2003.AssignmentBaseline", afxRegApartmentThreading, 0x50D939D1, 0xFC1D, 0x4D59, 0xB5, 0x86, 0x2C, 0x78, 0x14, 0x77, 0x26, 0xB6)

BEGIN_DISPATCH_MAP(AssignmentBaseline, CCmdTarget)
DISP_PROPERTY_EX_ID(AssignmentBaseline, "oTimephasedData", 1, odl_GetTimephasedData, SetNotSupported, VT_DISPATCH)
DISP_PROPERTY_EX_ID(AssignmentBaseline, "sNumber", 2, odl_GetNumber, odl_SetNumber, VT_BSTR)
DISP_PROPERTY_EX_ID(AssignmentBaseline, "sStart", 3, odl_GetStart, odl_SetStart, VT_BSTR)
DISP_PROPERTY_EX_ID(AssignmentBaseline, "sFinish", 4, odl_GetFinish, odl_SetFinish, VT_BSTR)
DISP_PROPERTY_EX_ID(AssignmentBaseline, "oWork", 5, odl_GetWork, SetNotSupported, VT_DISPATCH)
DISP_PROPERTY_EX_ID(AssignmentBaseline, "sCost", 6, odl_GetCost, odl_SetCost, VT_BSTR)
DISP_PROPERTY_EX_ID(AssignmentBaseline, "fBCWS", 7, odl_GetBCWS, odl_SetBCWS, VT_R4)
DISP_PROPERTY_EX_ID(AssignmentBaseline, "fBCWP", 8, odl_GetBCWP, odl_SetBCWP, VT_R4)
DISP_PROPERTY_EX_ID(AssignmentBaseline, "Key", 9, odl_GetKey, odl_SetKey, VT_BSTR)
DISP_FUNCTION_ID(AssignmentBaseline, "GetXML", 10, odl_GetXML, VT_BSTR, VTS_NONE)
DISP_FUNCTION_ID(AssignmentBaseline, "SetXML", 11, odl_SetXML, VT_EMPTY, VTS_BSTR)
DISP_FUNCTION_ID(AssignmentBaseline, "IsNull", 12, IsNull, VT_BOOL, VTS_NONE)
DISP_FUNCTION_ID(AssignmentBaseline, "Initialize", 13, Initialize, VT_EMPTY, VTS_NONE)
END_DISPATCH_MAP()

BEGIN_INTERFACE_MAP(AssignmentBaseline, CCmdTarget)
	INTERFACE_PART(AssignmentBaseline, IID_IAssignmentBaseline, Dispatch)
END_INTERFACE_MAP()

BEGIN_MESSAGE_MAP(AssignmentBaseline, CCmdTarget)
END_MESSAGE_MAP()


void AssignmentBaseline::Initialize(void)
{
	InitVars();
}

AssignmentBaseline::AssignmentBaseline()
{
	EnableAutomation();
	AfxOleLockApp();
	InitVars();
}

void AssignmentBaseline::InitVars(void)
{
	mp_oTimephasedData_C = new TimephasedData_C();
	mp_sNumber = "";
	mp_sStart = "";
	mp_sFinish = "";
	mp_oWork = new Duration();
	mp_sCost = "";
	mp_fBCWS = 0;
	mp_fBCWP = 0;
}

AssignmentBaseline::~AssignmentBaseline()
{
	delete mp_oTimephasedData_C;
	delete mp_oWork;
	AfxOleUnlockApp();
}

void AssignmentBaseline::OnFinalRelease()
{
	clsItemBase::OnFinalRelease();
}

IDispatch* AssignmentBaseline::odl_GetTimephasedData(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetTimephasedData()->GetIDispatch(TRUE);
}

TimephasedData_C* AssignmentBaseline::GetTimephasedData(void)
{
	return mp_oTimephasedData_C;
}

BSTR AssignmentBaseline::odl_GetNumber(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetNumber().AllocSysString();
}

CString AssignmentBaseline::GetNumber(void)
{
	return (CString) mp_sNumber;
}

void AssignmentBaseline::odl_SetNumber(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetNumber(newVal);
}

void AssignmentBaseline::SetNumber(CString newVal)
{
	mp_sNumber = newVal;
}

BSTR AssignmentBaseline::odl_GetStart(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetStart().AllocSysString();
}

CString AssignmentBaseline::GetStart(void)
{
	return (CString) mp_sStart;
}

void AssignmentBaseline::odl_SetStart(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetStart(newVal);
}

void AssignmentBaseline::SetStart(CString newVal)
{
	mp_sStart = newVal;
}

BSTR AssignmentBaseline::odl_GetFinish(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetFinish().AllocSysString();
}

CString AssignmentBaseline::GetFinish(void)
{
	return (CString) mp_sFinish;
}

void AssignmentBaseline::odl_SetFinish(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetFinish(newVal);
}

void AssignmentBaseline::SetFinish(CString newVal)
{
	mp_sFinish = newVal;
}

IDispatch* AssignmentBaseline::odl_GetWork(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetWork()->GetIDispatch(TRUE);
}

Duration* AssignmentBaseline::GetWork(void)
{
	return mp_oWork;
}

BSTR AssignmentBaseline::odl_GetCost(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetCost().AllocSysString();
}

CString AssignmentBaseline::GetCost(void)
{
	return (CString) mp_sCost;
}

void AssignmentBaseline::odl_SetCost(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetCost(newVal);
}

void AssignmentBaseline::SetCost(CString newVal)
{
	mp_sCost = newVal;
}

FLOAT AssignmentBaseline::odl_GetBCWS(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetBCWS();
}

FLOAT AssignmentBaseline::GetBCWS(void)
{
	return (FLOAT) mp_fBCWS;
}

void AssignmentBaseline::odl_SetBCWS(FLOAT newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetBCWS((FLOAT)newVal);
}

void AssignmentBaseline::SetBCWS(FLOAT newVal)
{
	mp_fBCWS = newVal;
}

FLOAT AssignmentBaseline::odl_GetBCWP(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetBCWP();
}

FLOAT AssignmentBaseline::GetBCWP(void)
{
	return (FLOAT) mp_fBCWP;
}

void AssignmentBaseline::odl_SetBCWP(FLOAT newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetBCWP((FLOAT)newVal);
}

void AssignmentBaseline::SetBCWP(FLOAT newVal)
{
	mp_fBCWP = newVal;
}

BSTR AssignmentBaseline::odl_GetKey(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetKey().AllocSysString();
}

void AssignmentBaseline::odl_SetKey(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	USES_CONVERSION;
	SetKey(F_TOUCSTR(T2A(newVal)));
}

CString AssignmentBaseline::GetKey(void)
{
	return mp_sKey;
}

void AssignmentBaseline::SetKey(CString newVal)
{
	mp_oCollection->mp_SetKey(&mp_sKey, newVal, MP_SET_KEY);
}

BSTR AssignmentBaseline::odl_GetXML(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetXML().AllocSysString();
}

void AssignmentBaseline::odl_SetXML(LPCTSTR sXML)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetXML(sXML);
}

BOOL AssignmentBaseline::IsNull(void)
{
	BOOL bReturn = TRUE;
	if (mp_oTimephasedData_C->IsNull() == FALSE)
	{
		bReturn = FALSE;
	}
	if (mp_sNumber != "")
	{
		bReturn = FALSE;
	}
	if (mp_sStart != "")
	{
		bReturn = FALSE;
	}
	if (mp_sFinish != "")
	{
		bReturn = FALSE;
	}
	if (mp_oWork->IsNull() == FALSE)
	{
		bReturn = FALSE;
	}
	if (mp_sCost != "")
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

CString AssignmentBaseline::GetXML(void)
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
	oXML.WriteProperty("Number", mp_sNumber);
	if (mp_sStart != "")
	{
		oXML.WriteProperty("Start", mp_sStart);
	}
	if (mp_sFinish != "")
	{
		oXML.WriteProperty("Finish", mp_sFinish);
	}
	oXML.WriteProperty("Work", mp_oWork);
	if (mp_sCost != "")
	{
		oXML.WriteProperty("Cost", mp_sCost);
	}
	oXML.WriteProperty("BCWS", mp_fBCWS);
	oXML.WriteProperty("BCWP", mp_fBCWP);
	return oXML.GetXML();
}

void AssignmentBaseline::SetXML(CString sXML)
{
	clsXML oXML("Baseline");
	oXML.SetSupportOptional(TRUE);
	oXML.SetXML(sXML);
	oXML.InitializeReader();
	mp_oTimephasedData_C->ReadObjectProtected(oXML);
	oXML.ReadProperty("Number", mp_sNumber);
	oXML.ReadProperty("Start", mp_sStart);
	oXML.ReadProperty("Finish", mp_sFinish);
	oXML.ReadProperty("Work", mp_oWork);
	oXML.ReadProperty("Cost", mp_sCost);
	oXML.ReadProperty("BCWS", mp_fBCWS);
	oXML.ReadProperty("BCWP", mp_fBCWP);
}