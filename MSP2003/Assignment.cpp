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
#include "Assignment.h"

IMPLEMENT_DYNCREATE(Assignment, CCmdTarget)

//{3E92AACA-346D-467B-B526-C58A5DDFAE1C}
static const IID IID_IAssignment = { 0x3E92AACA, 0x346D, 0x467B, { 0xB5, 0x26, 0xC5, 0x8A, 0x5D, 0xDF, 0xAE, 0x1C} };

//{05D9C669-40C3-40F6-9C87-0C50960249ED}
IMPLEMENT_OLECREATE_FLAGS(Assignment, "MSP2003.Assignment", afxRegApartmentThreading, 0x05D9C669, 0x40C3, 0x40F6, 0x9C, 0x87, 0x0C, 0x50, 0x96, 0x02, 0x49, 0xED)

BEGIN_DISPATCH_MAP(Assignment, CCmdTarget)
DISP_PROPERTY_EX_ID(Assignment, "lUID", 1, odl_GetUID, odl_SetUID, VT_I4)
DISP_PROPERTY_EX_ID(Assignment, "lTaskUID", 2, odl_GetTaskUID, odl_SetTaskUID, VT_I4)
DISP_PROPERTY_EX_ID(Assignment, "lResourceUID", 3, odl_GetResourceUID, odl_SetResourceUID, VT_I4)
DISP_PROPERTY_EX_ID(Assignment, "lPercentWorkComplete", 4, odl_GetPercentWorkComplete, odl_SetPercentWorkComplete, VT_I4)
DISP_PROPERTY_EX_ID(Assignment, "cActualCost", 5, odl_GetActualCost, odl_SetActualCost, VT_R8)
DISP_PROPERTY_EX_ID(Assignment, "dtActualFinish", 6, odl_GetActualFinish, odl_SetActualFinish, VT_DATE)
DISP_PROPERTY_EX_ID(Assignment, "cActualOvertimeCost", 7, odl_GetActualOvertimeCost, odl_SetActualOvertimeCost, VT_R8)
DISP_PROPERTY_EX_ID(Assignment, "oActualOvertimeWork", 8, odl_GetActualOvertimeWork, SetNotSupported, VT_DISPATCH)
DISP_PROPERTY_EX_ID(Assignment, "dtActualStart", 9, odl_GetActualStart, odl_SetActualStart, VT_DATE)
DISP_PROPERTY_EX_ID(Assignment, "oActualWork", 10, odl_GetActualWork, SetNotSupported, VT_DISPATCH)
DISP_PROPERTY_EX_ID(Assignment, "fACWP", 11, odl_GetACWP, odl_SetACWP, VT_R4)
DISP_PROPERTY_EX_ID(Assignment, "bConfirmed", 12, odl_GetConfirmed, odl_SetConfirmed, VT_BOOL)
DISP_PROPERTY_EX_ID(Assignment, "cCost", 13, odl_GetCost, odl_SetCost, VT_R8)
DISP_PROPERTY_EX_ID(Assignment, "yCostRateTable", 14, odl_GetCostRateTable, odl_SetCostRateTable, VT_I4)
DISP_PROPERTY_EX_ID(Assignment, "fCostVariance", 15, odl_GetCostVariance, odl_SetCostVariance, VT_R4)
DISP_PROPERTY_EX_ID(Assignment, "fCV", 16, odl_GetCV, odl_SetCV, VT_R4)
DISP_PROPERTY_EX_ID(Assignment, "lDelay", 17, odl_GetDelay, odl_SetDelay, VT_I4)
DISP_PROPERTY_EX_ID(Assignment, "dtFinish", 18, odl_GetFinish, odl_SetFinish, VT_DATE)
DISP_PROPERTY_EX_ID(Assignment, "lFinishVariance", 19, odl_GetFinishVariance, odl_SetFinishVariance, VT_I4)
DISP_PROPERTY_EX_ID(Assignment, "sHyperlink", 20, odl_GetHyperlink, odl_SetHyperlink, VT_BSTR)
DISP_PROPERTY_EX_ID(Assignment, "sHyperlinkAddress", 21, odl_GetHyperlinkAddress, odl_SetHyperlinkAddress, VT_BSTR)
DISP_PROPERTY_EX_ID(Assignment, "sHyperlinkSubAddress", 22, odl_GetHyperlinkSubAddress, odl_SetHyperlinkSubAddress, VT_BSTR)
DISP_PROPERTY_EX_ID(Assignment, "fWorkVariance", 23, odl_GetWorkVariance, odl_SetWorkVariance, VT_R4)
DISP_PROPERTY_EX_ID(Assignment, "bHasFixedRateUnits", 24, odl_GetHasFixedRateUnits, odl_SetHasFixedRateUnits, VT_BOOL)
DISP_PROPERTY_EX_ID(Assignment, "bFixedMaterial", 25, odl_GetFixedMaterial, odl_SetFixedMaterial, VT_BOOL)
DISP_PROPERTY_EX_ID(Assignment, "lLevelingDelay", 26, odl_GetLevelingDelay, odl_SetLevelingDelay, VT_I4)
DISP_PROPERTY_EX_ID(Assignment, "yLevelingDelayFormat", 27, odl_GetLevelingDelayFormat, odl_SetLevelingDelayFormat, VT_I4)
DISP_PROPERTY_EX_ID(Assignment, "bLinkedFields", 28, odl_GetLinkedFields, odl_SetLinkedFields, VT_BOOL)
DISP_PROPERTY_EX_ID(Assignment, "bMilestone", 29, odl_GetMilestone, odl_SetMilestone, VT_BOOL)
DISP_PROPERTY_EX_ID(Assignment, "sNotes", 30, odl_GetNotes, odl_SetNotes, VT_BSTR)
DISP_PROPERTY_EX_ID(Assignment, "bOverallocated", 31, odl_GetOverallocated, odl_SetOverallocated, VT_BOOL)
DISP_PROPERTY_EX_ID(Assignment, "cOvertimeCost", 32, odl_GetOvertimeCost, odl_SetOvertimeCost, VT_R8)
DISP_PROPERTY_EX_ID(Assignment, "oOvertimeWork", 33, odl_GetOvertimeWork, SetNotSupported, VT_DISPATCH)
DISP_PROPERTY_EX_ID(Assignment, "oRegularWork", 34, odl_GetRegularWork, SetNotSupported, VT_DISPATCH)
DISP_PROPERTY_EX_ID(Assignment, "cRemainingCost", 35, odl_GetRemainingCost, odl_SetRemainingCost, VT_R8)
DISP_PROPERTY_EX_ID(Assignment, "cRemainingOvertimeCost", 36, odl_GetRemainingOvertimeCost, odl_SetRemainingOvertimeCost, VT_R8)
DISP_PROPERTY_EX_ID(Assignment, "oRemainingOvertimeWork", 37, odl_GetRemainingOvertimeWork, SetNotSupported, VT_DISPATCH)
DISP_PROPERTY_EX_ID(Assignment, "oRemainingWork", 38, odl_GetRemainingWork, SetNotSupported, VT_DISPATCH)
DISP_PROPERTY_EX_ID(Assignment, "bResponsePending", 39, odl_GetResponsePending, odl_SetResponsePending, VT_BOOL)
DISP_PROPERTY_EX_ID(Assignment, "dtStart", 40, odl_GetStart, odl_SetStart, VT_DATE)
DISP_PROPERTY_EX_ID(Assignment, "dtStop", 41, odl_GetStop, odl_SetStop, VT_DATE)
DISP_PROPERTY_EX_ID(Assignment, "dtResume", 42, odl_GetResume, odl_SetResume, VT_DATE)
DISP_PROPERTY_EX_ID(Assignment, "lStartVariance", 43, odl_GetStartVariance, odl_SetStartVariance, VT_I4)
DISP_PROPERTY_EX_ID(Assignment, "fUnits", 44, odl_GetUnits, odl_SetUnits, VT_R4)
DISP_PROPERTY_EX_ID(Assignment, "bUpdateNeeded", 45, odl_GetUpdateNeeded, odl_SetUpdateNeeded, VT_BOOL)
DISP_PROPERTY_EX_ID(Assignment, "fVAC", 46, odl_GetVAC, odl_SetVAC, VT_R4)
DISP_PROPERTY_EX_ID(Assignment, "oWork", 47, odl_GetWork, SetNotSupported, VT_DISPATCH)
DISP_PROPERTY_EX_ID(Assignment, "yWorkContour", 48, odl_GetWorkContour, odl_SetWorkContour, VT_I4)
DISP_PROPERTY_EX_ID(Assignment, "fBCWS", 49, odl_GetBCWS, odl_SetBCWS, VT_R4)
DISP_PROPERTY_EX_ID(Assignment, "fBCWP", 50, odl_GetBCWP, odl_SetBCWP, VT_R4)
DISP_PROPERTY_EX_ID(Assignment, "yBookingType", 51, odl_GetBookingType, odl_SetBookingType, VT_I4)
DISP_PROPERTY_EX_ID(Assignment, "oActualWorkProtected", 52, odl_GetActualWorkProtected, SetNotSupported, VT_DISPATCH)
DISP_PROPERTY_EX_ID(Assignment, "oActualOvertimeWorkProtected", 53, odl_GetActualOvertimeWorkProtected, SetNotSupported, VT_DISPATCH)
DISP_PROPERTY_EX_ID(Assignment, "dtCreationDate", 54, odl_GetCreationDate, odl_SetCreationDate, VT_DATE)
DISP_PROPERTY_EX_ID(Assignment, "oExtendedAttribute", 55, odl_GetExtendedAttribute, SetNotSupported, VT_DISPATCH)
DISP_PROPERTY_EX_ID(Assignment, "oBaseline", 56, odl_GetBaseline, SetNotSupported, VT_DISPATCH)
DISP_PROPERTY_EX_ID(Assignment, "oTimephasedData", 57, odl_GetTimephasedData, SetNotSupported, VT_DISPATCH)
DISP_PROPERTY_EX_ID(Assignment, "Key", 58, odl_GetKey, odl_SetKey, VT_BSTR)
DISP_FUNCTION_ID(Assignment, "GetXML", 59, odl_GetXML, VT_BSTR, VTS_NONE)
DISP_FUNCTION_ID(Assignment, "SetXML", 60, odl_SetXML, VT_EMPTY, VTS_BSTR)
DISP_FUNCTION_ID(Assignment, "IsNull", 61, IsNull, VT_BOOL, VTS_NONE)
DISP_FUNCTION_ID(Assignment, "Initialize", 62, Initialize, VT_EMPTY, VTS_NONE)
END_DISPATCH_MAP()

BEGIN_INTERFACE_MAP(Assignment, CCmdTarget)
	INTERFACE_PART(Assignment, IID_IAssignment, Dispatch)
END_INTERFACE_MAP()

BEGIN_MESSAGE_MAP(Assignment, CCmdTarget)
END_MESSAGE_MAP()


void Assignment::Initialize(void)
{
	InitVars();
}

Assignment::Assignment()
{
	EnableAutomation();
	AfxOleLockApp();
	InitVars();
}

void Assignment::InitVars(void)
{
	mp_lUID = 0;
	mp_lTaskUID = 0;
	mp_lResourceUID = 0;
	mp_lPercentWorkComplete = 0;
	mp_cActualCost = 0;
	mp_dtActualFinish = (DATE) 0;
	mp_cActualOvertimeCost = 0;
	mp_oActualOvertimeWork = new Duration();
	mp_dtActualStart = (DATE) 0;
	mp_oActualWork = new Duration();
	mp_fACWP = 0;
	mp_bConfirmed = FALSE;
	mp_cCost = 0;
	mp_yCostRateTable = CRT_COST_RATE_TABLE_0;
	mp_fCostVariance = 0;
	mp_fCV = 0;
	mp_lDelay = 0;
	mp_dtFinish = (DATE) 0;
	mp_lFinishVariance = 0;
	mp_sHyperlink = "";
	mp_sHyperlinkAddress = "";
	mp_sHyperlinkSubAddress = "";
	mp_fWorkVariance = 0;
	mp_bHasFixedRateUnits = FALSE;
	mp_bFixedMaterial = FALSE;
	mp_lLevelingDelay = 0;
	mp_yLevelingDelayFormat = LDF_M;
	mp_bLinkedFields = FALSE;
	mp_bMilestone = FALSE;
	mp_sNotes = "";
	mp_bOverallocated = FALSE;
	mp_cOvertimeCost = 0;
	mp_oOvertimeWork = new Duration();
	mp_oRegularWork = new Duration();
	mp_cRemainingCost = 0;
	mp_cRemainingOvertimeCost = 0;
	mp_oRemainingOvertimeWork = new Duration();
	mp_oRemainingWork = new Duration();
	mp_bResponsePending = FALSE;
	mp_dtStart = (DATE) 0;
	mp_dtStop = (DATE) 0;
	mp_dtResume = (DATE) 0;
	mp_lStartVariance = 0;
	mp_fUnits = 0;
	mp_bUpdateNeeded = FALSE;
	mp_fVAC = 0;
	mp_oWork = new Duration();
	mp_yWorkContour = WC_FLAT;
	mp_fBCWS = 0;
	mp_fBCWP = 0;
	mp_yBookingType = BT_COMMITED;
	mp_oActualWorkProtected = new Duration();
	mp_oActualOvertimeWorkProtected = new Duration();
	mp_dtCreationDate = (DATE) 0;
	mp_oExtendedAttribute_C = new AssignmentExtendedAttribute_C();
	mp_oBaseline_C = new AssignmentBaseline_C();
	mp_oTimephasedData_C = new TimephasedData_C();
}

Assignment::~Assignment()
{
	delete mp_oActualOvertimeWork;
	delete mp_oActualWork;
	delete mp_oOvertimeWork;
	delete mp_oRegularWork;
	delete mp_oRemainingOvertimeWork;
	delete mp_oRemainingWork;
	delete mp_oWork;
	delete mp_oActualWorkProtected;
	delete mp_oActualOvertimeWorkProtected;
	delete mp_oExtendedAttribute_C;
	delete mp_oBaseline_C;
	delete mp_oTimephasedData_C;
	AfxOleUnlockApp();
}

void Assignment::OnFinalRelease()
{
	clsItemBase::OnFinalRelease();
}

LONG Assignment::odl_GetUID(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetUID();
}

LONG Assignment::GetUID(void)
{
	return (LONG) mp_lUID;
}

void Assignment::odl_SetUID(LONG newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetUID((LONG)newVal);
}

void Assignment::SetUID(LONG newVal)
{
	mp_lUID = newVal;
}

LONG Assignment::odl_GetTaskUID(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetTaskUID();
}

LONG Assignment::GetTaskUID(void)
{
	return (LONG) mp_lTaskUID;
}

void Assignment::odl_SetTaskUID(LONG newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetTaskUID((LONG)newVal);
}

void Assignment::SetTaskUID(LONG newVal)
{
	mp_lTaskUID = newVal;
}

LONG Assignment::odl_GetResourceUID(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetResourceUID();
}

LONG Assignment::GetResourceUID(void)
{
	return (LONG) mp_lResourceUID;
}

void Assignment::odl_SetResourceUID(LONG newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetResourceUID((LONG)newVal);
}

void Assignment::SetResourceUID(LONG newVal)
{
	mp_lResourceUID = newVal;
}

LONG Assignment::odl_GetPercentWorkComplete(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetPercentWorkComplete();
}

LONG Assignment::GetPercentWorkComplete(void)
{
	return (LONG) mp_lPercentWorkComplete;
}

void Assignment::odl_SetPercentWorkComplete(LONG newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetPercentWorkComplete((LONG)newVal);
}

void Assignment::SetPercentWorkComplete(LONG newVal)
{
	mp_lPercentWorkComplete = newVal;
}

DOUBLE Assignment::odl_GetActualCost(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetActualCost();
}

DOUBLE Assignment::GetActualCost(void)
{
	return (DOUBLE) mp_cActualCost;
}

void Assignment::odl_SetActualCost(DOUBLE newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetActualCost((DOUBLE)newVal);
}

void Assignment::SetActualCost(DOUBLE newVal)
{
	mp_cActualCost = newVal;
}

DATE Assignment::odl_GetActualFinish(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetActualFinish();
}

COleDateTime Assignment::GetActualFinish(void)
{
	return (COleDateTime) mp_dtActualFinish;
}

void Assignment::odl_SetActualFinish(DATE newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetActualFinish((COleDateTime)newVal);
}

void Assignment::SetActualFinish(COleDateTime newVal)
{
	mp_dtActualFinish = newVal;
}

DOUBLE Assignment::odl_GetActualOvertimeCost(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetActualOvertimeCost();
}

DOUBLE Assignment::GetActualOvertimeCost(void)
{
	return (DOUBLE) mp_cActualOvertimeCost;
}

void Assignment::odl_SetActualOvertimeCost(DOUBLE newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetActualOvertimeCost((DOUBLE)newVal);
}

void Assignment::SetActualOvertimeCost(DOUBLE newVal)
{
	mp_cActualOvertimeCost = newVal;
}

IDispatch* Assignment::odl_GetActualOvertimeWork(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetActualOvertimeWork()->GetIDispatch(TRUE);
}

Duration* Assignment::GetActualOvertimeWork(void)
{
	return mp_oActualOvertimeWork;
}

DATE Assignment::odl_GetActualStart(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetActualStart();
}

COleDateTime Assignment::GetActualStart(void)
{
	return (COleDateTime) mp_dtActualStart;
}

void Assignment::odl_SetActualStart(DATE newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetActualStart((COleDateTime)newVal);
}

void Assignment::SetActualStart(COleDateTime newVal)
{
	mp_dtActualStart = newVal;
}

IDispatch* Assignment::odl_GetActualWork(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetActualWork()->GetIDispatch(TRUE);
}

Duration* Assignment::GetActualWork(void)
{
	return mp_oActualWork;
}

FLOAT Assignment::odl_GetACWP(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetACWP();
}

FLOAT Assignment::GetACWP(void)
{
	return (FLOAT) mp_fACWP;
}

void Assignment::odl_SetACWP(FLOAT newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetACWP((FLOAT)newVal);
}

void Assignment::SetACWP(FLOAT newVal)
{
	mp_fACWP = newVal;
}

BOOL Assignment::odl_GetConfirmed(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetConfirmed();
}

BOOL Assignment::GetConfirmed(void)
{
	return (BOOL) mp_bConfirmed;
}

void Assignment::odl_SetConfirmed(BOOL newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetConfirmed((BOOL)newVal);
}

void Assignment::SetConfirmed(BOOL newVal)
{
	mp_bConfirmed = newVal;
}

DOUBLE Assignment::odl_GetCost(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetCost();
}

DOUBLE Assignment::GetCost(void)
{
	return (DOUBLE) mp_cCost;
}

void Assignment::odl_SetCost(DOUBLE newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetCost((DOUBLE)newVal);
}

void Assignment::SetCost(DOUBLE newVal)
{
	mp_cCost = newVal;
}

LONG Assignment::odl_GetCostRateTable(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetCostRateTable();
}

E_COSTRATETABLE Assignment::GetCostRateTable(void)
{
	return (E_COSTRATETABLE) mp_yCostRateTable;
}

void Assignment::odl_SetCostRateTable(LONG newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetCostRateTable((E_COSTRATETABLE)newVal);
}

void Assignment::SetCostRateTable(E_COSTRATETABLE newVal)
{
	mp_yCostRateTable = newVal;
}

FLOAT Assignment::odl_GetCostVariance(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetCostVariance();
}

FLOAT Assignment::GetCostVariance(void)
{
	return (FLOAT) mp_fCostVariance;
}

void Assignment::odl_SetCostVariance(FLOAT newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetCostVariance((FLOAT)newVal);
}

void Assignment::SetCostVariance(FLOAT newVal)
{
	mp_fCostVariance = newVal;
}

FLOAT Assignment::odl_GetCV(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetCV();
}

FLOAT Assignment::GetCV(void)
{
	return (FLOAT) mp_fCV;
}

void Assignment::odl_SetCV(FLOAT newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetCV((FLOAT)newVal);
}

void Assignment::SetCV(FLOAT newVal)
{
	mp_fCV = newVal;
}

LONG Assignment::odl_GetDelay(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetDelay();
}

LONG Assignment::GetDelay(void)
{
	return (LONG) mp_lDelay;
}

void Assignment::odl_SetDelay(LONG newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetDelay((LONG)newVal);
}

void Assignment::SetDelay(LONG newVal)
{
	mp_lDelay = newVal;
}

DATE Assignment::odl_GetFinish(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetFinish();
}

COleDateTime Assignment::GetFinish(void)
{
	return (COleDateTime) mp_dtFinish;
}

void Assignment::odl_SetFinish(DATE newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetFinish((COleDateTime)newVal);
}

void Assignment::SetFinish(COleDateTime newVal)
{
	mp_dtFinish = newVal;
}

LONG Assignment::odl_GetFinishVariance(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetFinishVariance();
}

LONG Assignment::GetFinishVariance(void)
{
	return (LONG) mp_lFinishVariance;
}

void Assignment::odl_SetFinishVariance(LONG newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetFinishVariance((LONG)newVal);
}

void Assignment::SetFinishVariance(LONG newVal)
{
	mp_lFinishVariance = newVal;
}

BSTR Assignment::odl_GetHyperlink(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetHyperlink().AllocSysString();
}

CString Assignment::GetHyperlink(void)
{
	return (CString) mp_sHyperlink;
}

void Assignment::odl_SetHyperlink(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetHyperlink(newVal);
}

void Assignment::SetHyperlink(CString newVal)
{
	if (newVal.GetLength() > 512)
	{
		newVal = newVal.Mid(0, 512);
	}
	mp_sHyperlink = newVal;
}

BSTR Assignment::odl_GetHyperlinkAddress(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetHyperlinkAddress().AllocSysString();
}

CString Assignment::GetHyperlinkAddress(void)
{
	return (CString) mp_sHyperlinkAddress;
}

void Assignment::odl_SetHyperlinkAddress(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetHyperlinkAddress(newVal);
}

void Assignment::SetHyperlinkAddress(CString newVal)
{
	if (newVal.GetLength() > 512)
	{
		newVal = newVal.Mid(0, 512);
	}
	mp_sHyperlinkAddress = newVal;
}

BSTR Assignment::odl_GetHyperlinkSubAddress(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetHyperlinkSubAddress().AllocSysString();
}

CString Assignment::GetHyperlinkSubAddress(void)
{
	return (CString) mp_sHyperlinkSubAddress;
}

void Assignment::odl_SetHyperlinkSubAddress(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetHyperlinkSubAddress(newVal);
}

void Assignment::SetHyperlinkSubAddress(CString newVal)
{
	if (newVal.GetLength() > 512)
	{
		newVal = newVal.Mid(0, 512);
	}
	mp_sHyperlinkSubAddress = newVal;
}

FLOAT Assignment::odl_GetWorkVariance(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetWorkVariance();
}

FLOAT Assignment::GetWorkVariance(void)
{
	return (FLOAT) mp_fWorkVariance;
}

void Assignment::odl_SetWorkVariance(FLOAT newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetWorkVariance((FLOAT)newVal);
}

void Assignment::SetWorkVariance(FLOAT newVal)
{
	mp_fWorkVariance = newVal;
}

BOOL Assignment::odl_GetHasFixedRateUnits(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetHasFixedRateUnits();
}

BOOL Assignment::GetHasFixedRateUnits(void)
{
	return (BOOL) mp_bHasFixedRateUnits;
}

void Assignment::odl_SetHasFixedRateUnits(BOOL newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetHasFixedRateUnits((BOOL)newVal);
}

void Assignment::SetHasFixedRateUnits(BOOL newVal)
{
	mp_bHasFixedRateUnits = newVal;
}

BOOL Assignment::odl_GetFixedMaterial(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetFixedMaterial();
}

BOOL Assignment::GetFixedMaterial(void)
{
	return (BOOL) mp_bFixedMaterial;
}

void Assignment::odl_SetFixedMaterial(BOOL newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetFixedMaterial((BOOL)newVal);
}

void Assignment::SetFixedMaterial(BOOL newVal)
{
	mp_bFixedMaterial = newVal;
}

LONG Assignment::odl_GetLevelingDelay(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetLevelingDelay();
}

LONG Assignment::GetLevelingDelay(void)
{
	return (LONG) mp_lLevelingDelay;
}

void Assignment::odl_SetLevelingDelay(LONG newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetLevelingDelay((LONG)newVal);
}

void Assignment::SetLevelingDelay(LONG newVal)
{
	mp_lLevelingDelay = newVal;
}

LONG Assignment::odl_GetLevelingDelayFormat(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetLevelingDelayFormat();
}

E_LEVELINGDELAYFORMAT Assignment::GetLevelingDelayFormat(void)
{
	return (E_LEVELINGDELAYFORMAT) mp_yLevelingDelayFormat;
}

void Assignment::odl_SetLevelingDelayFormat(LONG newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetLevelingDelayFormat((E_LEVELINGDELAYFORMAT)newVal);
}

void Assignment::SetLevelingDelayFormat(E_LEVELINGDELAYFORMAT newVal)
{
	mp_yLevelingDelayFormat = newVal;
}

BOOL Assignment::odl_GetLinkedFields(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetLinkedFields();
}

BOOL Assignment::GetLinkedFields(void)
{
	return (BOOL) mp_bLinkedFields;
}

void Assignment::odl_SetLinkedFields(BOOL newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetLinkedFields((BOOL)newVal);
}

void Assignment::SetLinkedFields(BOOL newVal)
{
	mp_bLinkedFields = newVal;
}

BOOL Assignment::odl_GetMilestone(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetMilestone();
}

BOOL Assignment::GetMilestone(void)
{
	return (BOOL) mp_bMilestone;
}

void Assignment::odl_SetMilestone(BOOL newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetMilestone((BOOL)newVal);
}

void Assignment::SetMilestone(BOOL newVal)
{
	mp_bMilestone = newVal;
}

BSTR Assignment::odl_GetNotes(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetNotes().AllocSysString();
}

CString Assignment::GetNotes(void)
{
	return (CString) mp_sNotes;
}

void Assignment::odl_SetNotes(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetNotes(newVal);
}

void Assignment::SetNotes(CString newVal)
{
	mp_sNotes = newVal;
}

BOOL Assignment::odl_GetOverallocated(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetOverallocated();
}

BOOL Assignment::GetOverallocated(void)
{
	return (BOOL) mp_bOverallocated;
}

void Assignment::odl_SetOverallocated(BOOL newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetOverallocated((BOOL)newVal);
}

void Assignment::SetOverallocated(BOOL newVal)
{
	mp_bOverallocated = newVal;
}

DOUBLE Assignment::odl_GetOvertimeCost(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetOvertimeCost();
}

DOUBLE Assignment::GetOvertimeCost(void)
{
	return (DOUBLE) mp_cOvertimeCost;
}

void Assignment::odl_SetOvertimeCost(DOUBLE newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetOvertimeCost((DOUBLE)newVal);
}

void Assignment::SetOvertimeCost(DOUBLE newVal)
{
	mp_cOvertimeCost = newVal;
}

IDispatch* Assignment::odl_GetOvertimeWork(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetOvertimeWork()->GetIDispatch(TRUE);
}

Duration* Assignment::GetOvertimeWork(void)
{
	return mp_oOvertimeWork;
}

IDispatch* Assignment::odl_GetRegularWork(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetRegularWork()->GetIDispatch(TRUE);
}

Duration* Assignment::GetRegularWork(void)
{
	return mp_oRegularWork;
}

DOUBLE Assignment::odl_GetRemainingCost(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetRemainingCost();
}

DOUBLE Assignment::GetRemainingCost(void)
{
	return (DOUBLE) mp_cRemainingCost;
}

void Assignment::odl_SetRemainingCost(DOUBLE newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetRemainingCost((DOUBLE)newVal);
}

void Assignment::SetRemainingCost(DOUBLE newVal)
{
	mp_cRemainingCost = newVal;
}

DOUBLE Assignment::odl_GetRemainingOvertimeCost(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetRemainingOvertimeCost();
}

DOUBLE Assignment::GetRemainingOvertimeCost(void)
{
	return (DOUBLE) mp_cRemainingOvertimeCost;
}

void Assignment::odl_SetRemainingOvertimeCost(DOUBLE newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetRemainingOvertimeCost((DOUBLE)newVal);
}

void Assignment::SetRemainingOvertimeCost(DOUBLE newVal)
{
	mp_cRemainingOvertimeCost = newVal;
}

IDispatch* Assignment::odl_GetRemainingOvertimeWork(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetRemainingOvertimeWork()->GetIDispatch(TRUE);
}

Duration* Assignment::GetRemainingOvertimeWork(void)
{
	return mp_oRemainingOvertimeWork;
}

IDispatch* Assignment::odl_GetRemainingWork(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetRemainingWork()->GetIDispatch(TRUE);
}

Duration* Assignment::GetRemainingWork(void)
{
	return mp_oRemainingWork;
}

BOOL Assignment::odl_GetResponsePending(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetResponsePending();
}

BOOL Assignment::GetResponsePending(void)
{
	return (BOOL) mp_bResponsePending;
}

void Assignment::odl_SetResponsePending(BOOL newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetResponsePending((BOOL)newVal);
}

void Assignment::SetResponsePending(BOOL newVal)
{
	mp_bResponsePending = newVal;
}

DATE Assignment::odl_GetStart(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetStart();
}

COleDateTime Assignment::GetStart(void)
{
	return (COleDateTime) mp_dtStart;
}

void Assignment::odl_SetStart(DATE newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetStart((COleDateTime)newVal);
}

void Assignment::SetStart(COleDateTime newVal)
{
	mp_dtStart = newVal;
}

DATE Assignment::odl_GetStop(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetStop();
}

COleDateTime Assignment::GetStop(void)
{
	return (COleDateTime) mp_dtStop;
}

void Assignment::odl_SetStop(DATE newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetStop((COleDateTime)newVal);
}

void Assignment::SetStop(COleDateTime newVal)
{
	mp_dtStop = newVal;
}

DATE Assignment::odl_GetResume(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetResume();
}

COleDateTime Assignment::GetResume(void)
{
	return (COleDateTime) mp_dtResume;
}

void Assignment::odl_SetResume(DATE newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetResume((COleDateTime)newVal);
}

void Assignment::SetResume(COleDateTime newVal)
{
	mp_dtResume = newVal;
}

LONG Assignment::odl_GetStartVariance(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetStartVariance();
}

LONG Assignment::GetStartVariance(void)
{
	return (LONG) mp_lStartVariance;
}

void Assignment::odl_SetStartVariance(LONG newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetStartVariance((LONG)newVal);
}

void Assignment::SetStartVariance(LONG newVal)
{
	mp_lStartVariance = newVal;
}

FLOAT Assignment::odl_GetUnits(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetUnits();
}

FLOAT Assignment::GetUnits(void)
{
	return (FLOAT) mp_fUnits;
}

void Assignment::odl_SetUnits(FLOAT newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetUnits((FLOAT)newVal);
}

void Assignment::SetUnits(FLOAT newVal)
{
	mp_fUnits = newVal;
}

BOOL Assignment::odl_GetUpdateNeeded(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetUpdateNeeded();
}

BOOL Assignment::GetUpdateNeeded(void)
{
	return (BOOL) mp_bUpdateNeeded;
}

void Assignment::odl_SetUpdateNeeded(BOOL newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetUpdateNeeded((BOOL)newVal);
}

void Assignment::SetUpdateNeeded(BOOL newVal)
{
	mp_bUpdateNeeded = newVal;
}

FLOAT Assignment::odl_GetVAC(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetVAC();
}

FLOAT Assignment::GetVAC(void)
{
	return (FLOAT) mp_fVAC;
}

void Assignment::odl_SetVAC(FLOAT newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetVAC((FLOAT)newVal);
}

void Assignment::SetVAC(FLOAT newVal)
{
	mp_fVAC = newVal;
}

IDispatch* Assignment::odl_GetWork(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetWork()->GetIDispatch(TRUE);
}

Duration* Assignment::GetWork(void)
{
	return mp_oWork;
}

LONG Assignment::odl_GetWorkContour(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetWorkContour();
}

E_WORKCONTOUR Assignment::GetWorkContour(void)
{
	return (E_WORKCONTOUR) mp_yWorkContour;
}

void Assignment::odl_SetWorkContour(LONG newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetWorkContour((E_WORKCONTOUR)newVal);
}

void Assignment::SetWorkContour(E_WORKCONTOUR newVal)
{
	mp_yWorkContour = newVal;
}

FLOAT Assignment::odl_GetBCWS(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetBCWS();
}

FLOAT Assignment::GetBCWS(void)
{
	return (FLOAT) mp_fBCWS;
}

void Assignment::odl_SetBCWS(FLOAT newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetBCWS((FLOAT)newVal);
}

void Assignment::SetBCWS(FLOAT newVal)
{
	mp_fBCWS = newVal;
}

FLOAT Assignment::odl_GetBCWP(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetBCWP();
}

FLOAT Assignment::GetBCWP(void)
{
	return (FLOAT) mp_fBCWP;
}

void Assignment::odl_SetBCWP(FLOAT newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetBCWP((FLOAT)newVal);
}

void Assignment::SetBCWP(FLOAT newVal)
{
	mp_fBCWP = newVal;
}

LONG Assignment::odl_GetBookingType(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetBookingType();
}

E_BOOKINGTYPE Assignment::GetBookingType(void)
{
	return (E_BOOKINGTYPE) mp_yBookingType;
}

void Assignment::odl_SetBookingType(LONG newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetBookingType((E_BOOKINGTYPE)newVal);
}

void Assignment::SetBookingType(E_BOOKINGTYPE newVal)
{
	mp_yBookingType = newVal;
}

IDispatch* Assignment::odl_GetActualWorkProtected(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetActualWorkProtected()->GetIDispatch(TRUE);
}

Duration* Assignment::GetActualWorkProtected(void)
{
	return mp_oActualWorkProtected;
}

IDispatch* Assignment::odl_GetActualOvertimeWorkProtected(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetActualOvertimeWorkProtected()->GetIDispatch(TRUE);
}

Duration* Assignment::GetActualOvertimeWorkProtected(void)
{
	return mp_oActualOvertimeWorkProtected;
}

DATE Assignment::odl_GetCreationDate(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetCreationDate();
}

COleDateTime Assignment::GetCreationDate(void)
{
	return (COleDateTime) mp_dtCreationDate;
}

void Assignment::odl_SetCreationDate(DATE newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetCreationDate((COleDateTime)newVal);
}

void Assignment::SetCreationDate(COleDateTime newVal)
{
	mp_dtCreationDate = newVal;
}

IDispatch* Assignment::odl_GetExtendedAttribute(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetExtendedAttribute()->GetIDispatch(TRUE);
}

AssignmentExtendedAttribute_C* Assignment::GetExtendedAttribute(void)
{
	return mp_oExtendedAttribute_C;
}

IDispatch* Assignment::odl_GetBaseline(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetBaseline()->GetIDispatch(TRUE);
}

AssignmentBaseline_C* Assignment::GetBaseline(void)
{
	return mp_oBaseline_C;
}

IDispatch* Assignment::odl_GetTimephasedData(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetTimephasedData()->GetIDispatch(TRUE);
}

TimephasedData_C* Assignment::GetTimephasedData(void)
{
	return mp_oTimephasedData_C;
}

BSTR Assignment::odl_GetKey(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetKey().AllocSysString();
}

void Assignment::odl_SetKey(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	USES_CONVERSION;
	SetKey(F_TOUCSTR(T2A(newVal)));
}

CString Assignment::GetKey(void)
{
	return mp_sKey;
}

void Assignment::SetKey(CString newVal)
{
	mp_oCollection->mp_SetKey(&mp_sKey, newVal, MP_SET_KEY);
}

BSTR Assignment::odl_GetXML(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetXML().AllocSysString();
}

void Assignment::odl_SetXML(LPCTSTR sXML)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetXML(sXML);
}

BOOL Assignment::IsNull(void)
{
	BOOL bReturn = TRUE;
	if (mp_lUID != 0)
	{
		bReturn = FALSE;
	}
	if (mp_lTaskUID != 0)
	{
		bReturn = FALSE;
	}
	if (mp_lResourceUID != 0)
	{
		bReturn = FALSE;
	}
	if (mp_lPercentWorkComplete != 0)
	{
		bReturn = FALSE;
	}
	if (mp_cActualCost != 0)
	{
		bReturn = FALSE;
	}
	if (mp_dtActualFinish.m_dt != 36494)
	{
		bReturn = FALSE;
	}
	if (mp_cActualOvertimeCost != 0)
	{
		bReturn = FALSE;
	}
	if (mp_oActualOvertimeWork->IsNull() == FALSE)
	{
		bReturn = FALSE;
	}
	if (mp_dtActualStart.m_dt != 36494)
	{
		bReturn = FALSE;
	}
	if (mp_oActualWork->IsNull() == FALSE)
	{
		bReturn = FALSE;
	}
	if (mp_fACWP != 0)
	{
		bReturn = FALSE;
	}
	if (mp_bConfirmed != FALSE)
	{
		bReturn = FALSE;
	}
	if (mp_cCost != 0)
	{
		bReturn = FALSE;
	}
	if (mp_yCostRateTable != CRT_COST_RATE_TABLE_0)
	{
		bReturn = FALSE;
	}
	if (mp_fCostVariance != 0)
	{
		bReturn = FALSE;
	}
	if (mp_fCV != 0)
	{
		bReturn = FALSE;
	}
	if (mp_lDelay != 0)
	{
		bReturn = FALSE;
	}
	if (mp_dtFinish.m_dt != 36494)
	{
		bReturn = FALSE;
	}
	if (mp_lFinishVariance != 0)
	{
		bReturn = FALSE;
	}
	if (mp_sHyperlink != "")
	{
		bReturn = FALSE;
	}
	if (mp_sHyperlinkAddress != "")
	{
		bReturn = FALSE;
	}
	if (mp_sHyperlinkSubAddress != "")
	{
		bReturn = FALSE;
	}
	if (mp_fWorkVariance != 0)
	{
		bReturn = FALSE;
	}
	if (mp_bHasFixedRateUnits != FALSE)
	{
		bReturn = FALSE;
	}
	if (mp_bFixedMaterial != FALSE)
	{
		bReturn = FALSE;
	}
	if (mp_lLevelingDelay != 0)
	{
		bReturn = FALSE;
	}
	if (mp_yLevelingDelayFormat != LDF_M)
	{
		bReturn = FALSE;
	}
	if (mp_bLinkedFields != FALSE)
	{
		bReturn = FALSE;
	}
	if (mp_bMilestone != FALSE)
	{
		bReturn = FALSE;
	}
	if (mp_sNotes != "")
	{
		bReturn = FALSE;
	}
	if (mp_bOverallocated != FALSE)
	{
		bReturn = FALSE;
	}
	if (mp_cOvertimeCost != 0)
	{
		bReturn = FALSE;
	}
	if (mp_oOvertimeWork->IsNull() == FALSE)
	{
		bReturn = FALSE;
	}
	if (mp_oRegularWork->IsNull() == FALSE)
	{
		bReturn = FALSE;
	}
	if (mp_cRemainingCost != 0)
	{
		bReturn = FALSE;
	}
	if (mp_cRemainingOvertimeCost != 0)
	{
		bReturn = FALSE;
	}
	if (mp_oRemainingOvertimeWork->IsNull() == FALSE)
	{
		bReturn = FALSE;
	}
	if (mp_oRemainingWork->IsNull() == FALSE)
	{
		bReturn = FALSE;
	}
	if (mp_bResponsePending != FALSE)
	{
		bReturn = FALSE;
	}
	if (mp_dtStart.m_dt != 36494)
	{
		bReturn = FALSE;
	}
	if (mp_dtStop.m_dt != 36494)
	{
		bReturn = FALSE;
	}
	if (mp_dtResume.m_dt != 36494)
	{
		bReturn = FALSE;
	}
	if (mp_lStartVariance != 0)
	{
		bReturn = FALSE;
	}
	if (mp_fUnits != 0)
	{
		bReturn = FALSE;
	}
	if (mp_bUpdateNeeded != FALSE)
	{
		bReturn = FALSE;
	}
	if (mp_fVAC != 0)
	{
		bReturn = FALSE;
	}
	if (mp_oWork->IsNull() == FALSE)
	{
		bReturn = FALSE;
	}
	if (mp_yWorkContour != WC_FLAT)
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
	if (mp_yBookingType != BT_COMMITED)
	{
		bReturn = FALSE;
	}
	if (mp_oActualWorkProtected->IsNull() == FALSE)
	{
		bReturn = FALSE;
	}
	if (mp_oActualOvertimeWorkProtected->IsNull() == FALSE)
	{
		bReturn = FALSE;
	}
	if (mp_dtCreationDate.m_dt != 36494)
	{
		bReturn = FALSE;
	}
	if (mp_oExtendedAttribute_C->IsNull() == FALSE)
	{
		bReturn = FALSE;
	}
	if (mp_oBaseline_C->IsNull() == FALSE)
	{
		bReturn = FALSE;
	}
	if (mp_oTimephasedData_C->IsNull() == FALSE)
	{
		bReturn = FALSE;
	}
	return bReturn;
}

CString Assignment::GetXML(void)
{
	if (IsNull() == TRUE)
	{
		return "<Assignment/>";
	}
	clsXML oXML("Assignment");
	oXML.InitializeWriter();
	oXML.SetSupportOptional(TRUE);
	oXML.SetBoolsAreNumeric(TRUE);
	oXML.WriteProperty("UID", mp_lUID);
	oXML.WriteProperty("TaskUID", mp_lTaskUID);
	oXML.WriteProperty("ResourceUID", mp_lResourceUID);
	oXML.WriteProperty("PercentWorkComplete", mp_lPercentWorkComplete);
	oXML.WriteProperty("ActualCost", mp_cActualCost);
	if (mp_dtActualFinish.m_dt != 36494)
	{
		oXML.WriteProperty("ActualFinish", mp_dtActualFinish);
	}
	oXML.WriteProperty("ActualOvertimeCost", mp_cActualOvertimeCost);
	oXML.WriteProperty("ActualOvertimeWork", mp_oActualOvertimeWork);
	if (mp_dtActualStart.m_dt != 36494)
	{
		oXML.WriteProperty("ActualStart", mp_dtActualStart);
	}
	oXML.WriteProperty("ActualWork", mp_oActualWork);
	oXML.WriteProperty("ACWP", mp_fACWP);
	oXML.WriteProperty("Confirmed", mp_bConfirmed);
	oXML.WriteProperty("Cost", mp_cCost);
	oXML.WriteProperty("CostRateTable", mp_yCostRateTable);
	oXML.WriteProperty("CostVariance", mp_fCostVariance);
	oXML.WriteProperty("CV", mp_fCV);
	oXML.WriteProperty("Delay", mp_lDelay);
	if (mp_dtFinish.m_dt != 36494)
	{
		oXML.WriteProperty("Finish", mp_dtFinish);
	}
	oXML.WriteProperty("FinishVariance", mp_lFinishVariance);
	if (mp_sHyperlink != "")
	{
		oXML.WriteProperty("Hyperlink", mp_sHyperlink);
	}
	if (mp_sHyperlinkAddress != "")
	{
		oXML.WriteProperty("HyperlinkAddress", mp_sHyperlinkAddress);
	}
	if (mp_sHyperlinkSubAddress != "")
	{
		oXML.WriteProperty("HyperlinkSubAddress", mp_sHyperlinkSubAddress);
	}
	oXML.WriteProperty("WorkVariance", mp_fWorkVariance);
	oXML.WriteProperty("HasFixedRateUnits", mp_bHasFixedRateUnits);
	oXML.WriteProperty("FixedMaterial", mp_bFixedMaterial);
	oXML.WriteProperty("LevelingDelay", mp_lLevelingDelay);
	oXML.WriteProperty("LevelingDelayFormat", mp_yLevelingDelayFormat);
	oXML.WriteProperty("LinkedFields", mp_bLinkedFields);
	oXML.WriteProperty("Milestone", mp_bMilestone);
	if (mp_sNotes != "")
	{
		oXML.WriteProperty("Notes", mp_sNotes);
	}
	oXML.WriteProperty("Overallocated", mp_bOverallocated);
	oXML.WriteProperty("OvertimeCost", mp_cOvertimeCost);
	oXML.WriteProperty("OvertimeWork", mp_oOvertimeWork);
	oXML.WriteProperty("RegularWork", mp_oRegularWork);
	oXML.WriteProperty("RemainingCost", mp_cRemainingCost);
	oXML.WriteProperty("RemainingOvertimeCost", mp_cRemainingOvertimeCost);
	oXML.WriteProperty("RemainingOvertimeWork", mp_oRemainingOvertimeWork);
	oXML.WriteProperty("RemainingWork", mp_oRemainingWork);
	oXML.WriteProperty("ResponsePending", mp_bResponsePending);
	if (mp_dtStart.m_dt != 36494)
	{
		oXML.WriteProperty("Start", mp_dtStart);
	}
	if (mp_dtStop.m_dt != 36494)
	{
		oXML.WriteProperty("Stop", mp_dtStop);
	}
	if (mp_dtResume.m_dt != 36494)
	{
		oXML.WriteProperty("Resume", mp_dtResume);
	}
	oXML.WriteProperty("StartVariance", mp_lStartVariance);
	oXML.WriteProperty("Units", mp_fUnits);
	oXML.WriteProperty("UpdateNeeded", mp_bUpdateNeeded);
	oXML.WriteProperty("VAC", mp_fVAC);
	oXML.WriteProperty("Work", mp_oWork);
	oXML.WriteProperty("WorkContour", mp_yWorkContour);
	oXML.WriteProperty("BCWS", mp_fBCWS);
	oXML.WriteProperty("BCWP", mp_fBCWP);
	oXML.WriteProperty("BookingType", mp_yBookingType);
	oXML.WriteProperty("ActualWorkProtected", mp_oActualWorkProtected);
	oXML.WriteProperty("ActualOvertimeWorkProtected", mp_oActualOvertimeWorkProtected);
	if (mp_dtCreationDate.m_dt != 36494)
	{
		oXML.WriteProperty("CreationDate", mp_dtCreationDate);
	}
	if (mp_oExtendedAttribute_C->IsNull() == FALSE)
	{
		mp_oExtendedAttribute_C->WriteObjectProtected(oXML);
	}
	if (mp_oBaseline_C->IsNull() == FALSE)
	{
		mp_oBaseline_C->WriteObjectProtected(oXML);
	}
	if (mp_oTimephasedData_C->IsNull() == FALSE)
	{
		mp_oTimephasedData_C->WriteObjectProtected(oXML);
	}
	return oXML.GetXML();
}

void Assignment::SetXML(CString sXML)
{
	clsXML oXML("Assignment");
	oXML.SetSupportOptional(TRUE);
	oXML.SetXML(sXML);
	oXML.InitializeReader();
	oXML.ReadProperty("UID", mp_lUID);
	oXML.ReadProperty("TaskUID", mp_lTaskUID);
	oXML.ReadProperty("ResourceUID", mp_lResourceUID);
	oXML.ReadProperty("PercentWorkComplete", mp_lPercentWorkComplete);
	oXML.ReadProperty("ActualCost", mp_cActualCost);
	oXML.ReadProperty("ActualFinish", mp_dtActualFinish);
	oXML.ReadProperty("ActualOvertimeCost", mp_cActualOvertimeCost);
	oXML.ReadProperty("ActualOvertimeWork", mp_oActualOvertimeWork);
	oXML.ReadProperty("ActualStart", mp_dtActualStart);
	oXML.ReadProperty("ActualWork", mp_oActualWork);
	oXML.ReadProperty("ACWP", mp_fACWP);
	oXML.ReadProperty("Confirmed", mp_bConfirmed);
	oXML.ReadProperty("Cost", mp_cCost);
	oXML.ReadProperty("CostRateTable", mp_yCostRateTable);
	oXML.ReadProperty("CostVariance", mp_fCostVariance);
	oXML.ReadProperty("CV", mp_fCV);
	oXML.ReadProperty("Delay", mp_lDelay);
	oXML.ReadProperty("Finish", mp_dtFinish);
	oXML.ReadProperty("FinishVariance", mp_lFinishVariance);
	oXML.ReadProperty("Hyperlink", mp_sHyperlink);
	if (mp_sHyperlink.GetLength() > 512)
	{
		mp_sHyperlink = mp_sHyperlink.Mid(0, 512);
	}
	oXML.ReadProperty("HyperlinkAddress", mp_sHyperlinkAddress);
	if (mp_sHyperlinkAddress.GetLength() > 512)
	{
		mp_sHyperlinkAddress = mp_sHyperlinkAddress.Mid(0, 512);
	}
	oXML.ReadProperty("HyperlinkSubAddress", mp_sHyperlinkSubAddress);
	if (mp_sHyperlinkSubAddress.GetLength() > 512)
	{
		mp_sHyperlinkSubAddress = mp_sHyperlinkSubAddress.Mid(0, 512);
	}
	oXML.ReadProperty("WorkVariance", mp_fWorkVariance);
	oXML.ReadProperty("HasFixedRateUnits", mp_bHasFixedRateUnits);
	oXML.ReadProperty("FixedMaterial", mp_bFixedMaterial);
	oXML.ReadProperty("LevelingDelay", mp_lLevelingDelay);
	oXML.ReadProperty("LevelingDelayFormat", mp_yLevelingDelayFormat);
	oXML.ReadProperty("LinkedFields", mp_bLinkedFields);
	oXML.ReadProperty("Milestone", mp_bMilestone);
	oXML.ReadProperty("Notes", mp_sNotes);
	oXML.ReadProperty("Overallocated", mp_bOverallocated);
	oXML.ReadProperty("OvertimeCost", mp_cOvertimeCost);
	oXML.ReadProperty("OvertimeWork", mp_oOvertimeWork);
	oXML.ReadProperty("RegularWork", mp_oRegularWork);
	oXML.ReadProperty("RemainingCost", mp_cRemainingCost);
	oXML.ReadProperty("RemainingOvertimeCost", mp_cRemainingOvertimeCost);
	oXML.ReadProperty("RemainingOvertimeWork", mp_oRemainingOvertimeWork);
	oXML.ReadProperty("RemainingWork", mp_oRemainingWork);
	oXML.ReadProperty("ResponsePending", mp_bResponsePending);
	oXML.ReadProperty("Start", mp_dtStart);
	oXML.ReadProperty("Stop", mp_dtStop);
	oXML.ReadProperty("Resume", mp_dtResume);
	oXML.ReadProperty("StartVariance", mp_lStartVariance);
	oXML.ReadProperty("Units", mp_fUnits);
	oXML.ReadProperty("UpdateNeeded", mp_bUpdateNeeded);
	oXML.ReadProperty("VAC", mp_fVAC);
	oXML.ReadProperty("Work", mp_oWork);
	oXML.ReadProperty("WorkContour", mp_yWorkContour);
	oXML.ReadProperty("BCWS", mp_fBCWS);
	oXML.ReadProperty("BCWP", mp_fBCWP);
	oXML.ReadProperty("BookingType", mp_yBookingType);
	oXML.ReadProperty("ActualWorkProtected", mp_oActualWorkProtected);
	oXML.ReadProperty("ActualOvertimeWorkProtected", mp_oActualOvertimeWorkProtected);
	oXML.ReadProperty("CreationDate", mp_dtCreationDate);
	mp_oExtendedAttribute_C->ReadObjectProtected(oXML);
	mp_oBaseline_C->ReadObjectProtected(oXML);
	mp_oTimephasedData_C->ReadObjectProtected(oXML);
}