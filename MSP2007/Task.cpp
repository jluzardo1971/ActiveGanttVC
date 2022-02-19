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
#include "Task.h"

IMPLEMENT_DYNCREATE(Task, CCmdTarget)

//{38FD657E-4EAA-41F7-BF64-4AE746BB4988}
static const IID IID_ITask = { 0x38FD657E, 0x4EAA, 0x41F7, { 0xBF, 0x64, 0x4A, 0xE7, 0x46, 0xBB, 0x49, 0x88} };

//{1327766C-1AF7-4161-8207-78B36B29060F}
IMPLEMENT_OLECREATE_FLAGS(Task, "MSP2007.Task", afxRegApartmentThreading, 0x1327766C, 0x1AF7, 0x4161, 0x82, 0x07, 0x78, 0xB3, 0x6B, 0x29, 0x06, 0x0F)

BEGIN_DISPATCH_MAP(Task, CCmdTarget)
DISP_PROPERTY_EX_ID(Task, "lUID", 1, odl_GetUID, odl_SetUID, VT_I4)
DISP_PROPERTY_EX_ID(Task, "lID", 2, odl_GetID, odl_SetID, VT_I4)
DISP_PROPERTY_EX_ID(Task, "sName", 3, odl_GetName, odl_SetName, VT_BSTR)
DISP_PROPERTY_EX_ID(Task, "yType", 4, odl_GetType, odl_SetType, VT_I4)
DISP_PROPERTY_EX_ID(Task, "bIsNull", 5, odl_GetIsNull, odl_SetIsNull, VT_BOOL)
DISP_PROPERTY_EX_ID(Task, "dtCreateDate", 6, odl_GetCreateDate, odl_SetCreateDate, VT_DATE)
DISP_PROPERTY_EX_ID(Task, "sContact", 7, odl_GetContact, odl_SetContact, VT_BSTR)
DISP_PROPERTY_EX_ID(Task, "sWBS", 8, odl_GetWBS, odl_SetWBS, VT_BSTR)
DISP_PROPERTY_EX_ID(Task, "sWBSLevel", 9, odl_GetWBSLevel, odl_SetWBSLevel, VT_BSTR)
DISP_PROPERTY_EX_ID(Task, "sOutlineNumber", 10, odl_GetOutlineNumber, odl_SetOutlineNumber, VT_BSTR)
DISP_PROPERTY_EX_ID(Task, "lOutlineLevel", 11, odl_GetOutlineLevel, odl_SetOutlineLevel, VT_I4)
DISP_PROPERTY_EX_ID(Task, "lPriority", 12, odl_GetPriority, odl_SetPriority, VT_I4)
DISP_PROPERTY_EX_ID(Task, "dtStart", 13, odl_GetStart, odl_SetStart, VT_DATE)
DISP_PROPERTY_EX_ID(Task, "dtFinish", 14, odl_GetFinish, odl_SetFinish, VT_DATE)
DISP_PROPERTY_EX_ID(Task, "oDuration", 15, odl_GetDuration, SetNotSupported, VT_DISPATCH)
DISP_PROPERTY_EX_ID(Task, "yDurationFormat", 16, odl_GetDurationFormat, odl_SetDurationFormat, VT_I4)
DISP_PROPERTY_EX_ID(Task, "oWork", 17, odl_GetWork, SetNotSupported, VT_DISPATCH)
DISP_PROPERTY_EX_ID(Task, "dtStop", 18, odl_GetStop, odl_SetStop, VT_DATE)
DISP_PROPERTY_EX_ID(Task, "dtResume", 19, odl_GetResume, odl_SetResume, VT_DATE)
DISP_PROPERTY_EX_ID(Task, "bResumeValid", 20, odl_GetResumeValid, odl_SetResumeValid, VT_BOOL)
DISP_PROPERTY_EX_ID(Task, "bEffortDriven", 21, odl_GetEffortDriven, odl_SetEffortDriven, VT_BOOL)
DISP_PROPERTY_EX_ID(Task, "bRecurring", 22, odl_GetRecurring, odl_SetRecurring, VT_BOOL)
DISP_PROPERTY_EX_ID(Task, "bOverAllocated", 23, odl_GetOverAllocated, odl_SetOverAllocated, VT_BOOL)
DISP_PROPERTY_EX_ID(Task, "bEstimated", 24, odl_GetEstimated, odl_SetEstimated, VT_BOOL)
DISP_PROPERTY_EX_ID(Task, "bMilestone", 25, odl_GetMilestone, odl_SetMilestone, VT_BOOL)
DISP_PROPERTY_EX_ID(Task, "bSummary", 26, odl_GetSummary, odl_SetSummary, VT_BOOL)
DISP_PROPERTY_EX_ID(Task, "bCritical", 27, odl_GetCritical, odl_SetCritical, VT_BOOL)
DISP_PROPERTY_EX_ID(Task, "bIsSubproject", 28, odl_GetIsSubproject, odl_SetIsSubproject, VT_BOOL)
DISP_PROPERTY_EX_ID(Task, "bIsSubprojectReadOnly", 29, odl_GetIsSubprojectReadOnly, odl_SetIsSubprojectReadOnly, VT_BOOL)
DISP_PROPERTY_EX_ID(Task, "sSubprojectName", 30, odl_GetSubprojectName, odl_SetSubprojectName, VT_BSTR)
DISP_PROPERTY_EX_ID(Task, "bExternalTask", 31, odl_GetExternalTask, odl_SetExternalTask, VT_BOOL)
DISP_PROPERTY_EX_ID(Task, "sExternalTaskProject", 32, odl_GetExternalTaskProject, odl_SetExternalTaskProject, VT_BSTR)
DISP_PROPERTY_EX_ID(Task, "dtEarlyStart", 33, odl_GetEarlyStart, odl_SetEarlyStart, VT_DATE)
DISP_PROPERTY_EX_ID(Task, "dtEarlyFinish", 34, odl_GetEarlyFinish, odl_SetEarlyFinish, VT_DATE)
DISP_PROPERTY_EX_ID(Task, "dtLateStart", 35, odl_GetLateStart, odl_SetLateStart, VT_DATE)
DISP_PROPERTY_EX_ID(Task, "dtLateFinish", 36, odl_GetLateFinish, odl_SetLateFinish, VT_DATE)
DISP_PROPERTY_EX_ID(Task, "lStartVariance", 37, odl_GetStartVariance, odl_SetStartVariance, VT_I4)
DISP_PROPERTY_EX_ID(Task, "lFinishVariance", 38, odl_GetFinishVariance, odl_SetFinishVariance, VT_I4)
DISP_PROPERTY_EX_ID(Task, "fWorkVariance", 39, odl_GetWorkVariance, odl_SetWorkVariance, VT_R4)
DISP_PROPERTY_EX_ID(Task, "lFreeSlack", 40, odl_GetFreeSlack, odl_SetFreeSlack, VT_I4)
DISP_PROPERTY_EX_ID(Task, "lTotalSlack", 41, odl_GetTotalSlack, odl_SetTotalSlack, VT_I4)
DISP_PROPERTY_EX_ID(Task, "fFixedCost", 42, odl_GetFixedCost, odl_SetFixedCost, VT_R4)
DISP_PROPERTY_EX_ID(Task, "yFixedCostAccrual", 43, odl_GetFixedCostAccrual, odl_SetFixedCostAccrual, VT_I4)
DISP_PROPERTY_EX_ID(Task, "lPercentComplete", 44, odl_GetPercentComplete, odl_SetPercentComplete, VT_I4)
DISP_PROPERTY_EX_ID(Task, "lPercentWorkComplete", 45, odl_GetPercentWorkComplete, odl_SetPercentWorkComplete, VT_I4)
DISP_PROPERTY_EX_ID(Task, "cCost", 46, odl_GetCost, odl_SetCost, VT_R8)
DISP_PROPERTY_EX_ID(Task, "cOvertimeCost", 47, odl_GetOvertimeCost, odl_SetOvertimeCost, VT_R8)
DISP_PROPERTY_EX_ID(Task, "oOvertimeWork", 48, odl_GetOvertimeWork, SetNotSupported, VT_DISPATCH)
DISP_PROPERTY_EX_ID(Task, "dtActualStart", 49, odl_GetActualStart, odl_SetActualStart, VT_DATE)
DISP_PROPERTY_EX_ID(Task, "dtActualFinish", 50, odl_GetActualFinish, odl_SetActualFinish, VT_DATE)
DISP_PROPERTY_EX_ID(Task, "oActualDuration", 51, odl_GetActualDuration, SetNotSupported, VT_DISPATCH)
DISP_PROPERTY_EX_ID(Task, "cActualCost", 52, odl_GetActualCost, odl_SetActualCost, VT_R8)
DISP_PROPERTY_EX_ID(Task, "cActualOvertimeCost", 53, odl_GetActualOvertimeCost, odl_SetActualOvertimeCost, VT_R8)
DISP_PROPERTY_EX_ID(Task, "oActualWork", 54, odl_GetActualWork, SetNotSupported, VT_DISPATCH)
DISP_PROPERTY_EX_ID(Task, "oActualOvertimeWork", 55, odl_GetActualOvertimeWork, SetNotSupported, VT_DISPATCH)
DISP_PROPERTY_EX_ID(Task, "oRegularWork", 56, odl_GetRegularWork, SetNotSupported, VT_DISPATCH)
DISP_PROPERTY_EX_ID(Task, "oRemainingDuration", 57, odl_GetRemainingDuration, SetNotSupported, VT_DISPATCH)
DISP_PROPERTY_EX_ID(Task, "cRemainingCost", 58, odl_GetRemainingCost, odl_SetRemainingCost, VT_R8)
DISP_PROPERTY_EX_ID(Task, "oRemainingWork", 59, odl_GetRemainingWork, SetNotSupported, VT_DISPATCH)
DISP_PROPERTY_EX_ID(Task, "cRemainingOvertimeCost", 60, odl_GetRemainingOvertimeCost, odl_SetRemainingOvertimeCost, VT_R8)
DISP_PROPERTY_EX_ID(Task, "oRemainingOvertimeWork", 61, odl_GetRemainingOvertimeWork, SetNotSupported, VT_DISPATCH)
DISP_PROPERTY_EX_ID(Task, "fACWP", 62, odl_GetACWP, odl_SetACWP, VT_R4)
DISP_PROPERTY_EX_ID(Task, "fCV", 63, odl_GetCV, odl_SetCV, VT_R4)
DISP_PROPERTY_EX_ID(Task, "yConstraintType", 64, odl_GetConstraintType, odl_SetConstraintType, VT_I4)
DISP_PROPERTY_EX_ID(Task, "lCalendarUID", 65, odl_GetCalendarUID, odl_SetCalendarUID, VT_I4)
DISP_PROPERTY_EX_ID(Task, "dtConstraintDate", 66, odl_GetConstraintDate, odl_SetConstraintDate, VT_DATE)
DISP_PROPERTY_EX_ID(Task, "dtDeadline", 67, odl_GetDeadline, odl_SetDeadline, VT_DATE)
DISP_PROPERTY_EX_ID(Task, "bLevelAssignments", 68, odl_GetLevelAssignments, odl_SetLevelAssignments, VT_BOOL)
DISP_PROPERTY_EX_ID(Task, "bLevelingCanSplit", 69, odl_GetLevelingCanSplit, odl_SetLevelingCanSplit, VT_BOOL)
DISP_PROPERTY_EX_ID(Task, "lLevelingDelay", 70, odl_GetLevelingDelay, odl_SetLevelingDelay, VT_I4)
DISP_PROPERTY_EX_ID(Task, "yLevelingDelayFormat", 71, odl_GetLevelingDelayFormat, odl_SetLevelingDelayFormat, VT_I4)
DISP_PROPERTY_EX_ID(Task, "dtPreLeveledStart", 72, odl_GetPreLeveledStart, odl_SetPreLeveledStart, VT_DATE)
DISP_PROPERTY_EX_ID(Task, "dtPreLeveledFinish", 73, odl_GetPreLeveledFinish, odl_SetPreLeveledFinish, VT_DATE)
DISP_PROPERTY_EX_ID(Task, "sHyperlink", 74, odl_GetHyperlink, odl_SetHyperlink, VT_BSTR)
DISP_PROPERTY_EX_ID(Task, "sHyperlinkAddress", 75, odl_GetHyperlinkAddress, odl_SetHyperlinkAddress, VT_BSTR)
DISP_PROPERTY_EX_ID(Task, "sHyperlinkSubAddress", 76, odl_GetHyperlinkSubAddress, odl_SetHyperlinkSubAddress, VT_BSTR)
DISP_PROPERTY_EX_ID(Task, "bIgnoreResourceCalendar", 77, odl_GetIgnoreResourceCalendar, odl_SetIgnoreResourceCalendar, VT_BOOL)
DISP_PROPERTY_EX_ID(Task, "sNotes", 78, odl_GetNotes, odl_SetNotes, VT_BSTR)
DISP_PROPERTY_EX_ID(Task, "bHideBar", 79, odl_GetHideBar, odl_SetHideBar, VT_BOOL)
DISP_PROPERTY_EX_ID(Task, "bRollup", 80, odl_GetRollup, odl_SetRollup, VT_BOOL)
DISP_PROPERTY_EX_ID(Task, "fBCWS", 81, odl_GetBCWS, odl_SetBCWS, VT_R4)
DISP_PROPERTY_EX_ID(Task, "fBCWP", 82, odl_GetBCWP, odl_SetBCWP, VT_R4)
DISP_PROPERTY_EX_ID(Task, "lPhysicalPercentComplete", 83, odl_GetPhysicalPercentComplete, odl_SetPhysicalPercentComplete, VT_I4)
DISP_PROPERTY_EX_ID(Task, "yEarnedValueMethod", 84, odl_GetEarnedValueMethod, odl_SetEarnedValueMethod, VT_I4)
DISP_PROPERTY_EX_ID(Task, "oPredecessorLink", 85, odl_GetPredecessorLink, SetNotSupported, VT_DISPATCH)
DISP_PROPERTY_EX_ID(Task, "oActualWorkProtected", 86, odl_GetActualWorkProtected, SetNotSupported, VT_DISPATCH)
DISP_PROPERTY_EX_ID(Task, "oActualOvertimeWorkProtected", 87, odl_GetActualOvertimeWorkProtected, SetNotSupported, VT_DISPATCH)
DISP_PROPERTY_EX_ID(Task, "oExtendedAttribute", 88, odl_GetExtendedAttribute, SetNotSupported, VT_DISPATCH)
DISP_PROPERTY_EX_ID(Task, "oBaseline", 89, odl_GetBaseline, SetNotSupported, VT_DISPATCH)
DISP_PROPERTY_EX_ID(Task, "oOutlineCode", 90, odl_GetOutlineCode, SetNotSupported, VT_DISPATCH)
DISP_PROPERTY_EX_ID(Task, "bIsPublished", 91, odl_GetIsPublished, odl_SetIsPublished, VT_BOOL)
DISP_PROPERTY_EX_ID(Task, "sStatusManager", 92, odl_GetStatusManager, odl_SetStatusManager, VT_BSTR)
DISP_PROPERTY_EX_ID(Task, "dtCommitmentStart", 93, odl_GetCommitmentStart, odl_SetCommitmentStart, VT_DATE)
DISP_PROPERTY_EX_ID(Task, "dtCommitmentFinish", 94, odl_GetCommitmentFinish, odl_SetCommitmentFinish, VT_DATE)
DISP_PROPERTY_EX_ID(Task, "yCommitmentType", 95, odl_GetCommitmentType, odl_SetCommitmentType, VT_I4)
DISP_PROPERTY_EX_ID(Task, "oTimephasedData", 96, odl_GetTimephasedData, SetNotSupported, VT_DISPATCH)
DISP_PROPERTY_EX_ID(Task, "Key", 97, odl_GetKey, odl_SetKey, VT_BSTR)
DISP_FUNCTION_ID(Task, "GetXML", 98, odl_GetXML, VT_BSTR, VTS_NONE)
DISP_FUNCTION_ID(Task, "SetXML", 99, odl_SetXML, VT_EMPTY, VTS_BSTR)
DISP_FUNCTION_ID(Task, "IsNull", 100, IsNull, VT_BOOL, VTS_NONE)
DISP_FUNCTION_ID(Task, "Initialize", 101, Initialize, VT_EMPTY, VTS_NONE)
END_DISPATCH_MAP()

BEGIN_INTERFACE_MAP(Task, CCmdTarget)
	INTERFACE_PART(Task, IID_ITask, Dispatch)
END_INTERFACE_MAP()

BEGIN_MESSAGE_MAP(Task, CCmdTarget)
END_MESSAGE_MAP()


void Task::Initialize(void)
{
	InitVars();
}

Task::Task()
{
	EnableAutomation();
	AfxOleLockApp();
	InitVars();
}

void Task::InitVars(void)
{
	mp_lUID = 0;
	mp_lID = 0;
	mp_sName = "";
	mp_yType = T_4_FIXED_UNITS;
	mp_bIsNull = FALSE;
	mp_dtCreateDate = (DATE) 0;
	mp_sContact = "";
	mp_sWBS = "";
	mp_sWBSLevel = "";
	mp_sOutlineNumber = "";
	mp_lOutlineLevel = 0;
	mp_lPriority = 0;
	mp_dtStart = (DATE) 0;
	mp_dtFinish = (DATE) 0;
	mp_oDuration = new Duration();
	mp_yDurationFormat = DF_M;
	mp_oWork = new Duration();
	mp_dtStop = (DATE) 0;
	mp_dtResume = (DATE) 0;
	mp_bResumeValid = FALSE;
	mp_bEffortDriven = FALSE;
	mp_bRecurring = FALSE;
	mp_bOverAllocated = FALSE;
	mp_bEstimated = FALSE;
	mp_bMilestone = FALSE;
	mp_bSummary = FALSE;
	mp_bCritical = FALSE;
	mp_bIsSubproject = FALSE;
	mp_bIsSubprojectReadOnly = FALSE;
	mp_sSubprojectName = "";
	mp_bExternalTask = FALSE;
	mp_sExternalTaskProject = "";
	mp_dtEarlyStart = (DATE) 0;
	mp_dtEarlyFinish = (DATE) 0;
	mp_dtLateStart = (DATE) 0;
	mp_dtLateFinish = (DATE) 0;
	mp_lStartVariance = 0;
	mp_lFinishVariance = 0;
	mp_fWorkVariance = 0;
	mp_lFreeSlack = 0;
	mp_lTotalSlack = 0;
	mp_fFixedCost = 0;
	mp_yFixedCostAccrual = FCA_START;
	mp_lPercentComplete = 0;
	mp_lPercentWorkComplete = 0;
	mp_cCost = 0;
	mp_cOvertimeCost = 0;
	mp_oOvertimeWork = new Duration();
	mp_dtActualStart = (DATE) 0;
	mp_dtActualFinish = (DATE) 0;
	mp_oActualDuration = new Duration();
	mp_cActualCost = 0;
	mp_cActualOvertimeCost = 0;
	mp_oActualWork = new Duration();
	mp_oActualOvertimeWork = new Duration();
	mp_oRegularWork = new Duration();
	mp_oRemainingDuration = new Duration();
	mp_cRemainingCost = 0;
	mp_oRemainingWork = new Duration();
	mp_cRemainingOvertimeCost = 0;
	mp_oRemainingOvertimeWork = new Duration();
	mp_fACWP = 0;
	mp_fCV = 0;
	mp_yConstraintType = CT_AS_SOON_AS_POSSIBLE;
	mp_lCalendarUID = 0;
	mp_dtConstraintDate = (DATE) 0;
	mp_dtDeadline = (DATE) 0;
	mp_bLevelAssignments = FALSE;
	mp_bLevelingCanSplit = FALSE;
	mp_lLevelingDelay = 0;
	mp_yLevelingDelayFormat = LDF_M;
	mp_dtPreLeveledStart = (DATE) 0;
	mp_dtPreLeveledFinish = (DATE) 0;
	mp_sHyperlink = "";
	mp_sHyperlinkAddress = "";
	mp_sHyperlinkSubAddress = "";
	mp_bIgnoreResourceCalendar = FALSE;
	mp_sNotes = "";
	mp_bHideBar = FALSE;
	mp_bRollup = FALSE;
	mp_fBCWS = 0;
	mp_fBCWP = 0;
	mp_lPhysicalPercentComplete = 0;
	mp_yEarnedValueMethod = EVM_PERCENT_COMPLETE;
	mp_oPredecessorLink_C = new TaskPredecessorLink_C();
	mp_oActualWorkProtected = new Duration();
	mp_oActualOvertimeWorkProtected = new Duration();
	mp_oExtendedAttribute_C = new TaskExtendedAttribute_C();
	mp_oBaseline_C = new TaskBaseline_C();
	mp_oOutlineCode_C = new TaskOutlineCode_C();
	mp_bIsPublished = FALSE;
	mp_sStatusManager = "";
	mp_dtCommitmentStart = (DATE) 0;
	mp_dtCommitmentFinish = (DATE) 0;
	mp_yCommitmentType = CT_THE_TASK_HAS_NO_DELIVERABLE_OR_DEPENDENCY_ON_A_DELIVERABLE;
	mp_oTimephasedData_C = new TimephasedData_C();
}

Task::~Task()
{
	delete mp_oDuration;
	delete mp_oWork;
	delete mp_oOvertimeWork;
	delete mp_oActualDuration;
	delete mp_oActualWork;
	delete mp_oActualOvertimeWork;
	delete mp_oRegularWork;
	delete mp_oRemainingDuration;
	delete mp_oRemainingWork;
	delete mp_oRemainingOvertimeWork;
	delete mp_oPredecessorLink_C;
	delete mp_oActualWorkProtected;
	delete mp_oActualOvertimeWorkProtected;
	delete mp_oExtendedAttribute_C;
	delete mp_oBaseline_C;
	delete mp_oOutlineCode_C;
	delete mp_oTimephasedData_C;
	AfxOleUnlockApp();
}

void Task::OnFinalRelease()
{
	clsItemBase::OnFinalRelease();
}

LONG Task::odl_GetUID(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetUID();
}

LONG Task::GetUID(void)
{
	return (LONG) mp_lUID;
}

void Task::odl_SetUID(LONG newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetUID((LONG)newVal);
}

void Task::SetUID(LONG newVal)
{
	mp_lUID = newVal;
}

LONG Task::odl_GetID(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetID();
}

LONG Task::GetID(void)
{
	return (LONG) mp_lID;
}

void Task::odl_SetID(LONG newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetID((LONG)newVal);
}

void Task::SetID(LONG newVal)
{
	mp_lID = newVal;
}

BSTR Task::odl_GetName(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetName().AllocSysString();
}

CString Task::GetName(void)
{
	return (CString) mp_sName;
}

void Task::odl_SetName(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetName(newVal);
}

void Task::SetName(CString newVal)
{
	if (newVal.GetLength() > 512)
	{
		newVal = newVal.Mid(0, 512);
	}
	mp_sName = newVal;
}

LONG Task::odl_GetType(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetType();
}

E_TYPE_4 Task::GetType(void)
{
	return (E_TYPE_4) mp_yType;
}

void Task::odl_SetType(LONG newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetType((E_TYPE_4)newVal);
}

void Task::SetType(E_TYPE_4 newVal)
{
	mp_yType = newVal;
}

BOOL Task::odl_GetIsNull(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetIsNull();
}

BOOL Task::GetIsNull(void)
{
	return (BOOL) mp_bIsNull;
}

void Task::odl_SetIsNull(BOOL newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetIsNull((BOOL)newVal);
}

void Task::SetIsNull(BOOL newVal)
{
	mp_bIsNull = newVal;
}

DATE Task::odl_GetCreateDate(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetCreateDate();
}

COleDateTime Task::GetCreateDate(void)
{
	return (COleDateTime) mp_dtCreateDate;
}

void Task::odl_SetCreateDate(DATE newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetCreateDate((COleDateTime)newVal);
}

void Task::SetCreateDate(COleDateTime newVal)
{
	mp_dtCreateDate = newVal;
}

BSTR Task::odl_GetContact(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetContact().AllocSysString();
}

CString Task::GetContact(void)
{
	return (CString) mp_sContact;
}

void Task::odl_SetContact(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetContact(newVal);
}

void Task::SetContact(CString newVal)
{
	if (newVal.GetLength() > 512)
	{
		newVal = newVal.Mid(0, 512);
	}
	mp_sContact = newVal;
}

BSTR Task::odl_GetWBS(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetWBS().AllocSysString();
}

CString Task::GetWBS(void)
{
	return (CString) mp_sWBS;
}

void Task::odl_SetWBS(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetWBS(newVal);
}

void Task::SetWBS(CString newVal)
{
	mp_sWBS = newVal;
}

BSTR Task::odl_GetWBSLevel(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetWBSLevel().AllocSysString();
}

CString Task::GetWBSLevel(void)
{
	return (CString) mp_sWBSLevel;
}

void Task::odl_SetWBSLevel(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetWBSLevel(newVal);
}

void Task::SetWBSLevel(CString newVal)
{
	mp_sWBSLevel = newVal;
}

BSTR Task::odl_GetOutlineNumber(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetOutlineNumber().AllocSysString();
}

CString Task::GetOutlineNumber(void)
{
	return (CString) mp_sOutlineNumber;
}

void Task::odl_SetOutlineNumber(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetOutlineNumber(newVal);
}

void Task::SetOutlineNumber(CString newVal)
{
	if (newVal.GetLength() > 512)
	{
		newVal = newVal.Mid(0, 512);
	}
	mp_sOutlineNumber = newVal;
}

LONG Task::odl_GetOutlineLevel(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetOutlineLevel();
}

LONG Task::GetOutlineLevel(void)
{
	return (LONG) mp_lOutlineLevel;
}

void Task::odl_SetOutlineLevel(LONG newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetOutlineLevel((LONG)newVal);
}

void Task::SetOutlineLevel(LONG newVal)
{
	mp_lOutlineLevel = newVal;
}

LONG Task::odl_GetPriority(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetPriority();
}

LONG Task::GetPriority(void)
{
	return (LONG) mp_lPriority;
}

void Task::odl_SetPriority(LONG newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetPriority((LONG)newVal);
}

void Task::SetPriority(LONG newVal)
{
	mp_lPriority = newVal;
}

DATE Task::odl_GetStart(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetStart();
}

COleDateTime Task::GetStart(void)
{
	return (COleDateTime) mp_dtStart;
}

void Task::odl_SetStart(DATE newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetStart((COleDateTime)newVal);
}

void Task::SetStart(COleDateTime newVal)
{
	mp_dtStart = newVal;
}

DATE Task::odl_GetFinish(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetFinish();
}

COleDateTime Task::GetFinish(void)
{
	return (COleDateTime) mp_dtFinish;
}

void Task::odl_SetFinish(DATE newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetFinish((COleDateTime)newVal);
}

void Task::SetFinish(COleDateTime newVal)
{
	mp_dtFinish = newVal;
}

IDispatch* Task::odl_GetDuration(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetDuration()->GetIDispatch(TRUE);
}

Duration* Task::GetDuration(void)
{
	return mp_oDuration;
}

LONG Task::odl_GetDurationFormat(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetDurationFormat();
}

E_DURATIONFORMAT Task::GetDurationFormat(void)
{
	return (E_DURATIONFORMAT) mp_yDurationFormat;
}

void Task::odl_SetDurationFormat(LONG newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetDurationFormat((E_DURATIONFORMAT)newVal);
}

void Task::SetDurationFormat(E_DURATIONFORMAT newVal)
{
	mp_yDurationFormat = newVal;
}

IDispatch* Task::odl_GetWork(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetWork()->GetIDispatch(TRUE);
}

Duration* Task::GetWork(void)
{
	return mp_oWork;
}

DATE Task::odl_GetStop(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetStop();
}

COleDateTime Task::GetStop(void)
{
	return (COleDateTime) mp_dtStop;
}

void Task::odl_SetStop(DATE newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetStop((COleDateTime)newVal);
}

void Task::SetStop(COleDateTime newVal)
{
	mp_dtStop = newVal;
}

DATE Task::odl_GetResume(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetResume();
}

COleDateTime Task::GetResume(void)
{
	return (COleDateTime) mp_dtResume;
}

void Task::odl_SetResume(DATE newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetResume((COleDateTime)newVal);
}

void Task::SetResume(COleDateTime newVal)
{
	mp_dtResume = newVal;
}

BOOL Task::odl_GetResumeValid(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetResumeValid();
}

BOOL Task::GetResumeValid(void)
{
	return (BOOL) mp_bResumeValid;
}

void Task::odl_SetResumeValid(BOOL newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetResumeValid((BOOL)newVal);
}

void Task::SetResumeValid(BOOL newVal)
{
	mp_bResumeValid = newVal;
}

BOOL Task::odl_GetEffortDriven(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetEffortDriven();
}

BOOL Task::GetEffortDriven(void)
{
	return (BOOL) mp_bEffortDriven;
}

void Task::odl_SetEffortDriven(BOOL newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetEffortDriven((BOOL)newVal);
}

void Task::SetEffortDriven(BOOL newVal)
{
	mp_bEffortDriven = newVal;
}

BOOL Task::odl_GetRecurring(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetRecurring();
}

BOOL Task::GetRecurring(void)
{
	return (BOOL) mp_bRecurring;
}

void Task::odl_SetRecurring(BOOL newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetRecurring((BOOL)newVal);
}

void Task::SetRecurring(BOOL newVal)
{
	mp_bRecurring = newVal;
}

BOOL Task::odl_GetOverAllocated(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetOverAllocated();
}

BOOL Task::GetOverAllocated(void)
{
	return (BOOL) mp_bOverAllocated;
}

void Task::odl_SetOverAllocated(BOOL newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetOverAllocated((BOOL)newVal);
}

void Task::SetOverAllocated(BOOL newVal)
{
	mp_bOverAllocated = newVal;
}

BOOL Task::odl_GetEstimated(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetEstimated();
}

BOOL Task::GetEstimated(void)
{
	return (BOOL) mp_bEstimated;
}

void Task::odl_SetEstimated(BOOL newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetEstimated((BOOL)newVal);
}

void Task::SetEstimated(BOOL newVal)
{
	mp_bEstimated = newVal;
}

BOOL Task::odl_GetMilestone(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetMilestone();
}

BOOL Task::GetMilestone(void)
{
	return (BOOL) mp_bMilestone;
}

void Task::odl_SetMilestone(BOOL newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetMilestone((BOOL)newVal);
}

void Task::SetMilestone(BOOL newVal)
{
	mp_bMilestone = newVal;
}

BOOL Task::odl_GetSummary(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetSummary();
}

BOOL Task::GetSummary(void)
{
	return (BOOL) mp_bSummary;
}

void Task::odl_SetSummary(BOOL newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetSummary((BOOL)newVal);
}

void Task::SetSummary(BOOL newVal)
{
	mp_bSummary = newVal;
}

BOOL Task::odl_GetCritical(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetCritical();
}

BOOL Task::GetCritical(void)
{
	return (BOOL) mp_bCritical;
}

void Task::odl_SetCritical(BOOL newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetCritical((BOOL)newVal);
}

void Task::SetCritical(BOOL newVal)
{
	mp_bCritical = newVal;
}

BOOL Task::odl_GetIsSubproject(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetIsSubproject();
}

BOOL Task::GetIsSubproject(void)
{
	return (BOOL) mp_bIsSubproject;
}

void Task::odl_SetIsSubproject(BOOL newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetIsSubproject((BOOL)newVal);
}

void Task::SetIsSubproject(BOOL newVal)
{
	mp_bIsSubproject = newVal;
}

BOOL Task::odl_GetIsSubprojectReadOnly(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetIsSubprojectReadOnly();
}

BOOL Task::GetIsSubprojectReadOnly(void)
{
	return (BOOL) mp_bIsSubprojectReadOnly;
}

void Task::odl_SetIsSubprojectReadOnly(BOOL newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetIsSubprojectReadOnly((BOOL)newVal);
}

void Task::SetIsSubprojectReadOnly(BOOL newVal)
{
	mp_bIsSubprojectReadOnly = newVal;
}

BSTR Task::odl_GetSubprojectName(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetSubprojectName().AllocSysString();
}

CString Task::GetSubprojectName(void)
{
	return (CString) mp_sSubprojectName;
}

void Task::odl_SetSubprojectName(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetSubprojectName(newVal);
}

void Task::SetSubprojectName(CString newVal)
{
	if (newVal.GetLength() > 512)
	{
		newVal = newVal.Mid(0, 512);
	}
	mp_sSubprojectName = newVal;
}

BOOL Task::odl_GetExternalTask(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetExternalTask();
}

BOOL Task::GetExternalTask(void)
{
	return (BOOL) mp_bExternalTask;
}

void Task::odl_SetExternalTask(BOOL newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetExternalTask((BOOL)newVal);
}

void Task::SetExternalTask(BOOL newVal)
{
	mp_bExternalTask = newVal;
}

BSTR Task::odl_GetExternalTaskProject(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetExternalTaskProject().AllocSysString();
}

CString Task::GetExternalTaskProject(void)
{
	return (CString) mp_sExternalTaskProject;
}

void Task::odl_SetExternalTaskProject(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetExternalTaskProject(newVal);
}

void Task::SetExternalTaskProject(CString newVal)
{
	if (newVal.GetLength() > 512)
	{
		newVal = newVal.Mid(0, 512);
	}
	mp_sExternalTaskProject = newVal;
}

DATE Task::odl_GetEarlyStart(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetEarlyStart();
}

COleDateTime Task::GetEarlyStart(void)
{
	return (COleDateTime) mp_dtEarlyStart;
}

void Task::odl_SetEarlyStart(DATE newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetEarlyStart((COleDateTime)newVal);
}

void Task::SetEarlyStart(COleDateTime newVal)
{
	mp_dtEarlyStart = newVal;
}

DATE Task::odl_GetEarlyFinish(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetEarlyFinish();
}

COleDateTime Task::GetEarlyFinish(void)
{
	return (COleDateTime) mp_dtEarlyFinish;
}

void Task::odl_SetEarlyFinish(DATE newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetEarlyFinish((COleDateTime)newVal);
}

void Task::SetEarlyFinish(COleDateTime newVal)
{
	mp_dtEarlyFinish = newVal;
}

DATE Task::odl_GetLateStart(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetLateStart();
}

COleDateTime Task::GetLateStart(void)
{
	return (COleDateTime) mp_dtLateStart;
}

void Task::odl_SetLateStart(DATE newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetLateStart((COleDateTime)newVal);
}

void Task::SetLateStart(COleDateTime newVal)
{
	mp_dtLateStart = newVal;
}

DATE Task::odl_GetLateFinish(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetLateFinish();
}

COleDateTime Task::GetLateFinish(void)
{
	return (COleDateTime) mp_dtLateFinish;
}

void Task::odl_SetLateFinish(DATE newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetLateFinish((COleDateTime)newVal);
}

void Task::SetLateFinish(COleDateTime newVal)
{
	mp_dtLateFinish = newVal;
}

LONG Task::odl_GetStartVariance(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetStartVariance();
}

LONG Task::GetStartVariance(void)
{
	return (LONG) mp_lStartVariance;
}

void Task::odl_SetStartVariance(LONG newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetStartVariance((LONG)newVal);
}

void Task::SetStartVariance(LONG newVal)
{
	mp_lStartVariance = newVal;
}

LONG Task::odl_GetFinishVariance(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetFinishVariance();
}

LONG Task::GetFinishVariance(void)
{
	return (LONG) mp_lFinishVariance;
}

void Task::odl_SetFinishVariance(LONG newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetFinishVariance((LONG)newVal);
}

void Task::SetFinishVariance(LONG newVal)
{
	mp_lFinishVariance = newVal;
}

FLOAT Task::odl_GetWorkVariance(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetWorkVariance();
}

FLOAT Task::GetWorkVariance(void)
{
	return (FLOAT) mp_fWorkVariance;
}

void Task::odl_SetWorkVariance(FLOAT newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetWorkVariance((FLOAT)newVal);
}

void Task::SetWorkVariance(FLOAT newVal)
{
	mp_fWorkVariance = newVal;
}

LONG Task::odl_GetFreeSlack(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetFreeSlack();
}

LONG Task::GetFreeSlack(void)
{
	return (LONG) mp_lFreeSlack;
}

void Task::odl_SetFreeSlack(LONG newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetFreeSlack((LONG)newVal);
}

void Task::SetFreeSlack(LONG newVal)
{
	mp_lFreeSlack = newVal;
}

LONG Task::odl_GetTotalSlack(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetTotalSlack();
}

LONG Task::GetTotalSlack(void)
{
	return (LONG) mp_lTotalSlack;
}

void Task::odl_SetTotalSlack(LONG newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetTotalSlack((LONG)newVal);
}

void Task::SetTotalSlack(LONG newVal)
{
	mp_lTotalSlack = newVal;
}

FLOAT Task::odl_GetFixedCost(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetFixedCost();
}

FLOAT Task::GetFixedCost(void)
{
	return (FLOAT) mp_fFixedCost;
}

void Task::odl_SetFixedCost(FLOAT newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetFixedCost((FLOAT)newVal);
}

void Task::SetFixedCost(FLOAT newVal)
{
	mp_fFixedCost = newVal;
}

LONG Task::odl_GetFixedCostAccrual(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetFixedCostAccrual();
}

E_FIXEDCOSTACCRUAL Task::GetFixedCostAccrual(void)
{
	return (E_FIXEDCOSTACCRUAL) mp_yFixedCostAccrual;
}

void Task::odl_SetFixedCostAccrual(LONG newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetFixedCostAccrual((E_FIXEDCOSTACCRUAL)newVal);
}

void Task::SetFixedCostAccrual(E_FIXEDCOSTACCRUAL newVal)
{
	mp_yFixedCostAccrual = newVal;
}

LONG Task::odl_GetPercentComplete(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetPercentComplete();
}

LONG Task::GetPercentComplete(void)
{
	return (LONG) mp_lPercentComplete;
}

void Task::odl_SetPercentComplete(LONG newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetPercentComplete((LONG)newVal);
}

void Task::SetPercentComplete(LONG newVal)
{
	mp_lPercentComplete = newVal;
}

LONG Task::odl_GetPercentWorkComplete(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetPercentWorkComplete();
}

LONG Task::GetPercentWorkComplete(void)
{
	return (LONG) mp_lPercentWorkComplete;
}

void Task::odl_SetPercentWorkComplete(LONG newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetPercentWorkComplete((LONG)newVal);
}

void Task::SetPercentWorkComplete(LONG newVal)
{
	mp_lPercentWorkComplete = newVal;
}

DOUBLE Task::odl_GetCost(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetCost();
}

DOUBLE Task::GetCost(void)
{
	return (DOUBLE) mp_cCost;
}

void Task::odl_SetCost(DOUBLE newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetCost((DOUBLE)newVal);
}

void Task::SetCost(DOUBLE newVal)
{
	mp_cCost = newVal;
}

DOUBLE Task::odl_GetOvertimeCost(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetOvertimeCost();
}

DOUBLE Task::GetOvertimeCost(void)
{
	return (DOUBLE) mp_cOvertimeCost;
}

void Task::odl_SetOvertimeCost(DOUBLE newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetOvertimeCost((DOUBLE)newVal);
}

void Task::SetOvertimeCost(DOUBLE newVal)
{
	mp_cOvertimeCost = newVal;
}

IDispatch* Task::odl_GetOvertimeWork(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetOvertimeWork()->GetIDispatch(TRUE);
}

Duration* Task::GetOvertimeWork(void)
{
	return mp_oOvertimeWork;
}

DATE Task::odl_GetActualStart(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetActualStart();
}

COleDateTime Task::GetActualStart(void)
{
	return (COleDateTime) mp_dtActualStart;
}

void Task::odl_SetActualStart(DATE newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetActualStart((COleDateTime)newVal);
}

void Task::SetActualStart(COleDateTime newVal)
{
	mp_dtActualStart = newVal;
}

DATE Task::odl_GetActualFinish(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetActualFinish();
}

COleDateTime Task::GetActualFinish(void)
{
	return (COleDateTime) mp_dtActualFinish;
}

void Task::odl_SetActualFinish(DATE newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetActualFinish((COleDateTime)newVal);
}

void Task::SetActualFinish(COleDateTime newVal)
{
	mp_dtActualFinish = newVal;
}

IDispatch* Task::odl_GetActualDuration(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetActualDuration()->GetIDispatch(TRUE);
}

Duration* Task::GetActualDuration(void)
{
	return mp_oActualDuration;
}

DOUBLE Task::odl_GetActualCost(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetActualCost();
}

DOUBLE Task::GetActualCost(void)
{
	return (DOUBLE) mp_cActualCost;
}

void Task::odl_SetActualCost(DOUBLE newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetActualCost((DOUBLE)newVal);
}

void Task::SetActualCost(DOUBLE newVal)
{
	mp_cActualCost = newVal;
}

DOUBLE Task::odl_GetActualOvertimeCost(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetActualOvertimeCost();
}

DOUBLE Task::GetActualOvertimeCost(void)
{
	return (DOUBLE) mp_cActualOvertimeCost;
}

void Task::odl_SetActualOvertimeCost(DOUBLE newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetActualOvertimeCost((DOUBLE)newVal);
}

void Task::SetActualOvertimeCost(DOUBLE newVal)
{
	mp_cActualOvertimeCost = newVal;
}

IDispatch* Task::odl_GetActualWork(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetActualWork()->GetIDispatch(TRUE);
}

Duration* Task::GetActualWork(void)
{
	return mp_oActualWork;
}

IDispatch* Task::odl_GetActualOvertimeWork(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetActualOvertimeWork()->GetIDispatch(TRUE);
}

Duration* Task::GetActualOvertimeWork(void)
{
	return mp_oActualOvertimeWork;
}

IDispatch* Task::odl_GetRegularWork(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetRegularWork()->GetIDispatch(TRUE);
}

Duration* Task::GetRegularWork(void)
{
	return mp_oRegularWork;
}

IDispatch* Task::odl_GetRemainingDuration(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetRemainingDuration()->GetIDispatch(TRUE);
}

Duration* Task::GetRemainingDuration(void)
{
	return mp_oRemainingDuration;
}

DOUBLE Task::odl_GetRemainingCost(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetRemainingCost();
}

DOUBLE Task::GetRemainingCost(void)
{
	return (DOUBLE) mp_cRemainingCost;
}

void Task::odl_SetRemainingCost(DOUBLE newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetRemainingCost((DOUBLE)newVal);
}

void Task::SetRemainingCost(DOUBLE newVal)
{
	mp_cRemainingCost = newVal;
}

IDispatch* Task::odl_GetRemainingWork(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetRemainingWork()->GetIDispatch(TRUE);
}

Duration* Task::GetRemainingWork(void)
{
	return mp_oRemainingWork;
}

DOUBLE Task::odl_GetRemainingOvertimeCost(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetRemainingOvertimeCost();
}

DOUBLE Task::GetRemainingOvertimeCost(void)
{
	return (DOUBLE) mp_cRemainingOvertimeCost;
}

void Task::odl_SetRemainingOvertimeCost(DOUBLE newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetRemainingOvertimeCost((DOUBLE)newVal);
}

void Task::SetRemainingOvertimeCost(DOUBLE newVal)
{
	mp_cRemainingOvertimeCost = newVal;
}

IDispatch* Task::odl_GetRemainingOvertimeWork(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetRemainingOvertimeWork()->GetIDispatch(TRUE);
}

Duration* Task::GetRemainingOvertimeWork(void)
{
	return mp_oRemainingOvertimeWork;
}

FLOAT Task::odl_GetACWP(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetACWP();
}

FLOAT Task::GetACWP(void)
{
	return (FLOAT) mp_fACWP;
}

void Task::odl_SetACWP(FLOAT newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetACWP((FLOAT)newVal);
}

void Task::SetACWP(FLOAT newVal)
{
	mp_fACWP = newVal;
}

FLOAT Task::odl_GetCV(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetCV();
}

FLOAT Task::GetCV(void)
{
	return (FLOAT) mp_fCV;
}

void Task::odl_SetCV(FLOAT newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetCV((FLOAT)newVal);
}

void Task::SetCV(FLOAT newVal)
{
	mp_fCV = newVal;
}

LONG Task::odl_GetConstraintType(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetConstraintType();
}

E_CONSTRAINTTYPE Task::GetConstraintType(void)
{
	return (E_CONSTRAINTTYPE) mp_yConstraintType;
}

void Task::odl_SetConstraintType(LONG newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetConstraintType((E_CONSTRAINTTYPE)newVal);
}

void Task::SetConstraintType(E_CONSTRAINTTYPE newVal)
{
	mp_yConstraintType = newVal;
}

LONG Task::odl_GetCalendarUID(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetCalendarUID();
}

LONG Task::GetCalendarUID(void)
{
	return (LONG) mp_lCalendarUID;
}

void Task::odl_SetCalendarUID(LONG newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetCalendarUID((LONG)newVal);
}

void Task::SetCalendarUID(LONG newVal)
{
	mp_lCalendarUID = newVal;
}

DATE Task::odl_GetConstraintDate(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetConstraintDate();
}

COleDateTime Task::GetConstraintDate(void)
{
	return (COleDateTime) mp_dtConstraintDate;
}

void Task::odl_SetConstraintDate(DATE newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetConstraintDate((COleDateTime)newVal);
}

void Task::SetConstraintDate(COleDateTime newVal)
{
	mp_dtConstraintDate = newVal;
}

DATE Task::odl_GetDeadline(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetDeadline();
}

COleDateTime Task::GetDeadline(void)
{
	return (COleDateTime) mp_dtDeadline;
}

void Task::odl_SetDeadline(DATE newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetDeadline((COleDateTime)newVal);
}

void Task::SetDeadline(COleDateTime newVal)
{
	mp_dtDeadline = newVal;
}

BOOL Task::odl_GetLevelAssignments(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetLevelAssignments();
}

BOOL Task::GetLevelAssignments(void)
{
	return (BOOL) mp_bLevelAssignments;
}

void Task::odl_SetLevelAssignments(BOOL newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetLevelAssignments((BOOL)newVal);
}

void Task::SetLevelAssignments(BOOL newVal)
{
	mp_bLevelAssignments = newVal;
}

BOOL Task::odl_GetLevelingCanSplit(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetLevelingCanSplit();
}

BOOL Task::GetLevelingCanSplit(void)
{
	return (BOOL) mp_bLevelingCanSplit;
}

void Task::odl_SetLevelingCanSplit(BOOL newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetLevelingCanSplit((BOOL)newVal);
}

void Task::SetLevelingCanSplit(BOOL newVal)
{
	mp_bLevelingCanSplit = newVal;
}

LONG Task::odl_GetLevelingDelay(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetLevelingDelay();
}

LONG Task::GetLevelingDelay(void)
{
	return (LONG) mp_lLevelingDelay;
}

void Task::odl_SetLevelingDelay(LONG newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetLevelingDelay((LONG)newVal);
}

void Task::SetLevelingDelay(LONG newVal)
{
	mp_lLevelingDelay = newVal;
}

LONG Task::odl_GetLevelingDelayFormat(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetLevelingDelayFormat();
}

E_LEVELINGDELAYFORMAT Task::GetLevelingDelayFormat(void)
{
	return (E_LEVELINGDELAYFORMAT) mp_yLevelingDelayFormat;
}

void Task::odl_SetLevelingDelayFormat(LONG newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetLevelingDelayFormat((E_LEVELINGDELAYFORMAT)newVal);
}

void Task::SetLevelingDelayFormat(E_LEVELINGDELAYFORMAT newVal)
{
	mp_yLevelingDelayFormat = newVal;
}

DATE Task::odl_GetPreLeveledStart(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetPreLeveledStart();
}

COleDateTime Task::GetPreLeveledStart(void)
{
	return (COleDateTime) mp_dtPreLeveledStart;
}

void Task::odl_SetPreLeveledStart(DATE newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetPreLeveledStart((COleDateTime)newVal);
}

void Task::SetPreLeveledStart(COleDateTime newVal)
{
	mp_dtPreLeveledStart = newVal;
}

DATE Task::odl_GetPreLeveledFinish(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetPreLeveledFinish();
}

COleDateTime Task::GetPreLeveledFinish(void)
{
	return (COleDateTime) mp_dtPreLeveledFinish;
}

void Task::odl_SetPreLeveledFinish(DATE newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetPreLeveledFinish((COleDateTime)newVal);
}

void Task::SetPreLeveledFinish(COleDateTime newVal)
{
	mp_dtPreLeveledFinish = newVal;
}

BSTR Task::odl_GetHyperlink(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetHyperlink().AllocSysString();
}

CString Task::GetHyperlink(void)
{
	return (CString) mp_sHyperlink;
}

void Task::odl_SetHyperlink(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetHyperlink(newVal);
}

void Task::SetHyperlink(CString newVal)
{
	if (newVal.GetLength() > 512)
	{
		newVal = newVal.Mid(0, 512);
	}
	mp_sHyperlink = newVal;
}

BSTR Task::odl_GetHyperlinkAddress(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetHyperlinkAddress().AllocSysString();
}

CString Task::GetHyperlinkAddress(void)
{
	return (CString) mp_sHyperlinkAddress;
}

void Task::odl_SetHyperlinkAddress(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetHyperlinkAddress(newVal);
}

void Task::SetHyperlinkAddress(CString newVal)
{
	if (newVal.GetLength() > 512)
	{
		newVal = newVal.Mid(0, 512);
	}
	mp_sHyperlinkAddress = newVal;
}

BSTR Task::odl_GetHyperlinkSubAddress(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetHyperlinkSubAddress().AllocSysString();
}

CString Task::GetHyperlinkSubAddress(void)
{
	return (CString) mp_sHyperlinkSubAddress;
}

void Task::odl_SetHyperlinkSubAddress(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetHyperlinkSubAddress(newVal);
}

void Task::SetHyperlinkSubAddress(CString newVal)
{
	if (newVal.GetLength() > 512)
	{
		newVal = newVal.Mid(0, 512);
	}
	mp_sHyperlinkSubAddress = newVal;
}

BOOL Task::odl_GetIgnoreResourceCalendar(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetIgnoreResourceCalendar();
}

BOOL Task::GetIgnoreResourceCalendar(void)
{
	return (BOOL) mp_bIgnoreResourceCalendar;
}

void Task::odl_SetIgnoreResourceCalendar(BOOL newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetIgnoreResourceCalendar((BOOL)newVal);
}

void Task::SetIgnoreResourceCalendar(BOOL newVal)
{
	mp_bIgnoreResourceCalendar = newVal;
}

BSTR Task::odl_GetNotes(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetNotes().AllocSysString();
}

CString Task::GetNotes(void)
{
	return (CString) mp_sNotes;
}

void Task::odl_SetNotes(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetNotes(newVal);
}

void Task::SetNotes(CString newVal)
{
	mp_sNotes = newVal;
}

BOOL Task::odl_GetHideBar(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetHideBar();
}

BOOL Task::GetHideBar(void)
{
	return (BOOL) mp_bHideBar;
}

void Task::odl_SetHideBar(BOOL newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetHideBar((BOOL)newVal);
}

void Task::SetHideBar(BOOL newVal)
{
	mp_bHideBar = newVal;
}

BOOL Task::odl_GetRollup(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetRollup();
}

BOOL Task::GetRollup(void)
{
	return (BOOL) mp_bRollup;
}

void Task::odl_SetRollup(BOOL newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetRollup((BOOL)newVal);
}

void Task::SetRollup(BOOL newVal)
{
	mp_bRollup = newVal;
}

FLOAT Task::odl_GetBCWS(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetBCWS();
}

FLOAT Task::GetBCWS(void)
{
	return (FLOAT) mp_fBCWS;
}

void Task::odl_SetBCWS(FLOAT newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetBCWS((FLOAT)newVal);
}

void Task::SetBCWS(FLOAT newVal)
{
	mp_fBCWS = newVal;
}

FLOAT Task::odl_GetBCWP(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetBCWP();
}

FLOAT Task::GetBCWP(void)
{
	return (FLOAT) mp_fBCWP;
}

void Task::odl_SetBCWP(FLOAT newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetBCWP((FLOAT)newVal);
}

void Task::SetBCWP(FLOAT newVal)
{
	mp_fBCWP = newVal;
}

LONG Task::odl_GetPhysicalPercentComplete(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetPhysicalPercentComplete();
}

LONG Task::GetPhysicalPercentComplete(void)
{
	return (LONG) mp_lPhysicalPercentComplete;
}

void Task::odl_SetPhysicalPercentComplete(LONG newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetPhysicalPercentComplete((LONG)newVal);
}

void Task::SetPhysicalPercentComplete(LONG newVal)
{
	mp_lPhysicalPercentComplete = newVal;
}

LONG Task::odl_GetEarnedValueMethod(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetEarnedValueMethod();
}

E_EARNEDVALUEMETHOD Task::GetEarnedValueMethod(void)
{
	return (E_EARNEDVALUEMETHOD) mp_yEarnedValueMethod;
}

void Task::odl_SetEarnedValueMethod(LONG newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetEarnedValueMethod((E_EARNEDVALUEMETHOD)newVal);
}

void Task::SetEarnedValueMethod(E_EARNEDVALUEMETHOD newVal)
{
	mp_yEarnedValueMethod = newVal;
}

IDispatch* Task::odl_GetPredecessorLink(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetPredecessorLink()->GetIDispatch(TRUE);
}

TaskPredecessorLink_C* Task::GetPredecessorLink(void)
{
	return mp_oPredecessorLink_C;
}

IDispatch* Task::odl_GetActualWorkProtected(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetActualWorkProtected()->GetIDispatch(TRUE);
}

Duration* Task::GetActualWorkProtected(void)
{
	return mp_oActualWorkProtected;
}

IDispatch* Task::odl_GetActualOvertimeWorkProtected(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetActualOvertimeWorkProtected()->GetIDispatch(TRUE);
}

Duration* Task::GetActualOvertimeWorkProtected(void)
{
	return mp_oActualOvertimeWorkProtected;
}

IDispatch* Task::odl_GetExtendedAttribute(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetExtendedAttribute()->GetIDispatch(TRUE);
}

TaskExtendedAttribute_C* Task::GetExtendedAttribute(void)
{
	return mp_oExtendedAttribute_C;
}

IDispatch* Task::odl_GetBaseline(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetBaseline()->GetIDispatch(TRUE);
}

TaskBaseline_C* Task::GetBaseline(void)
{
	return mp_oBaseline_C;
}

IDispatch* Task::odl_GetOutlineCode(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetOutlineCode()->GetIDispatch(TRUE);
}

TaskOutlineCode_C* Task::GetOutlineCode(void)
{
	return mp_oOutlineCode_C;
}

BOOL Task::odl_GetIsPublished(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetIsPublished();
}

BOOL Task::GetIsPublished(void)
{
	return (BOOL) mp_bIsPublished;
}

void Task::odl_SetIsPublished(BOOL newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetIsPublished((BOOL)newVal);
}

void Task::SetIsPublished(BOOL newVal)
{
	mp_bIsPublished = newVal;
}

BSTR Task::odl_GetStatusManager(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetStatusManager().AllocSysString();
}

CString Task::GetStatusManager(void)
{
	return (CString) mp_sStatusManager;
}

void Task::odl_SetStatusManager(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetStatusManager(newVal);
}

void Task::SetStatusManager(CString newVal)
{
	mp_sStatusManager = newVal;
}

DATE Task::odl_GetCommitmentStart(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetCommitmentStart();
}

COleDateTime Task::GetCommitmentStart(void)
{
	return (COleDateTime) mp_dtCommitmentStart;
}

void Task::odl_SetCommitmentStart(DATE newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetCommitmentStart((COleDateTime)newVal);
}

void Task::SetCommitmentStart(COleDateTime newVal)
{
	mp_dtCommitmentStart = newVal;
}

DATE Task::odl_GetCommitmentFinish(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetCommitmentFinish();
}

COleDateTime Task::GetCommitmentFinish(void)
{
	return (COleDateTime) mp_dtCommitmentFinish;
}

void Task::odl_SetCommitmentFinish(DATE newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetCommitmentFinish((COleDateTime)newVal);
}

void Task::SetCommitmentFinish(COleDateTime newVal)
{
	mp_dtCommitmentFinish = newVal;
}

LONG Task::odl_GetCommitmentType(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetCommitmentType();
}

E_COMMITMENTTYPE Task::GetCommitmentType(void)
{
	return (E_COMMITMENTTYPE) mp_yCommitmentType;
}

void Task::odl_SetCommitmentType(LONG newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetCommitmentType((E_COMMITMENTTYPE)newVal);
}

void Task::SetCommitmentType(E_COMMITMENTTYPE newVal)
{
	mp_yCommitmentType = newVal;
}

IDispatch* Task::odl_GetTimephasedData(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetTimephasedData()->GetIDispatch(TRUE);
}

TimephasedData_C* Task::GetTimephasedData(void)
{
	return mp_oTimephasedData_C;
}

BSTR Task::odl_GetKey(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetKey().AllocSysString();
}

void Task::odl_SetKey(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	USES_CONVERSION;
	SetKey(F_TOUCSTR(T2A(newVal)));
}

CString Task::GetKey(void)
{
	return mp_sKey;
}

void Task::SetKey(CString newVal)
{
	mp_oCollection->mp_SetKey(&mp_sKey, newVal, MP_SET_KEY);
}

BSTR Task::odl_GetXML(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetXML().AllocSysString();
}

void Task::odl_SetXML(LPCTSTR sXML)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetXML(sXML);
}

BOOL Task::IsNull(void)
{
	BOOL bReturn = TRUE;
	if (mp_lUID != 0)
	{
		bReturn = FALSE;
	}
	if (mp_lID != 0)
	{
		bReturn = FALSE;
	}
	if (mp_sName != "")
	{
		bReturn = FALSE;
	}
	if (mp_yType != T_4_FIXED_UNITS)
	{
		bReturn = FALSE;
	}
	if (mp_bIsNull != FALSE)
	{
		bReturn = FALSE;
	}
	if (mp_dtCreateDate.m_dt != 36494)
	{
		bReturn = FALSE;
	}
	if (mp_sContact != "")
	{
		bReturn = FALSE;
	}
	if (mp_sWBS != "")
	{
		bReturn = FALSE;
	}
	if (mp_sWBSLevel != "")
	{
		bReturn = FALSE;
	}
	if (mp_sOutlineNumber != "")
	{
		bReturn = FALSE;
	}
	if (mp_lOutlineLevel != 0)
	{
		bReturn = FALSE;
	}
	if (mp_lPriority != 0)
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
	if (mp_oWork->IsNull() == FALSE)
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
	if (mp_bResumeValid != FALSE)
	{
		bReturn = FALSE;
	}
	if (mp_bEffortDriven != FALSE)
	{
		bReturn = FALSE;
	}
	if (mp_bRecurring != FALSE)
	{
		bReturn = FALSE;
	}
	if (mp_bOverAllocated != FALSE)
	{
		bReturn = FALSE;
	}
	if (mp_bEstimated != FALSE)
	{
		bReturn = FALSE;
	}
	if (mp_bMilestone != FALSE)
	{
		bReturn = FALSE;
	}
	if (mp_bSummary != FALSE)
	{
		bReturn = FALSE;
	}
	if (mp_bCritical != FALSE)
	{
		bReturn = FALSE;
	}
	if (mp_bIsSubproject != FALSE)
	{
		bReturn = FALSE;
	}
	if (mp_bIsSubprojectReadOnly != FALSE)
	{
		bReturn = FALSE;
	}
	if (mp_sSubprojectName != "")
	{
		bReturn = FALSE;
	}
	if (mp_bExternalTask != FALSE)
	{
		bReturn = FALSE;
	}
	if (mp_sExternalTaskProject != "")
	{
		bReturn = FALSE;
	}
	if (mp_dtEarlyStart.m_dt != 36494)
	{
		bReturn = FALSE;
	}
	if (mp_dtEarlyFinish.m_dt != 36494)
	{
		bReturn = FALSE;
	}
	if (mp_dtLateStart.m_dt != 36494)
	{
		bReturn = FALSE;
	}
	if (mp_dtLateFinish.m_dt != 36494)
	{
		bReturn = FALSE;
	}
	if (mp_lStartVariance != 0)
	{
		bReturn = FALSE;
	}
	if (mp_lFinishVariance != 0)
	{
		bReturn = FALSE;
	}
	if (mp_fWorkVariance != 0)
	{
		bReturn = FALSE;
	}
	if (mp_lFreeSlack != 0)
	{
		bReturn = FALSE;
	}
	if (mp_lTotalSlack != 0)
	{
		bReturn = FALSE;
	}
	if (mp_fFixedCost != 0)
	{
		bReturn = FALSE;
	}
	if (mp_yFixedCostAccrual != FCA_START)
	{
		bReturn = FALSE;
	}
	if (mp_lPercentComplete != 0)
	{
		bReturn = FALSE;
	}
	if (mp_lPercentWorkComplete != 0)
	{
		bReturn = FALSE;
	}
	if (mp_cCost != 0)
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
	if (mp_dtActualStart.m_dt != 36494)
	{
		bReturn = FALSE;
	}
	if (mp_dtActualFinish.m_dt != 36494)
	{
		bReturn = FALSE;
	}
	if (mp_oActualDuration->IsNull() == FALSE)
	{
		bReturn = FALSE;
	}
	if (mp_cActualCost != 0)
	{
		bReturn = FALSE;
	}
	if (mp_cActualOvertimeCost != 0)
	{
		bReturn = FALSE;
	}
	if (mp_oActualWork->IsNull() == FALSE)
	{
		bReturn = FALSE;
	}
	if (mp_oActualOvertimeWork->IsNull() == FALSE)
	{
		bReturn = FALSE;
	}
	if (mp_oRegularWork->IsNull() == FALSE)
	{
		bReturn = FALSE;
	}
	if (mp_oRemainingDuration->IsNull() == FALSE)
	{
		bReturn = FALSE;
	}
	if (mp_cRemainingCost != 0)
	{
		bReturn = FALSE;
	}
	if (mp_oRemainingWork->IsNull() == FALSE)
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
	if (mp_fACWP != 0)
	{
		bReturn = FALSE;
	}
	if (mp_fCV != 0)
	{
		bReturn = FALSE;
	}
	if (mp_yConstraintType != CT_AS_SOON_AS_POSSIBLE)
	{
		bReturn = FALSE;
	}
	if (mp_lCalendarUID != 0)
	{
		bReturn = FALSE;
	}
	if (mp_dtConstraintDate.m_dt != 36494)
	{
		bReturn = FALSE;
	}
	if (mp_dtDeadline.m_dt != 36494)
	{
		bReturn = FALSE;
	}
	if (mp_bLevelAssignments != FALSE)
	{
		bReturn = FALSE;
	}
	if (mp_bLevelingCanSplit != FALSE)
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
	if (mp_dtPreLeveledStart.m_dt != 36494)
	{
		bReturn = FALSE;
	}
	if (mp_dtPreLeveledFinish.m_dt != 36494)
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
	if (mp_bIgnoreResourceCalendar != FALSE)
	{
		bReturn = FALSE;
	}
	if (mp_sNotes != "")
	{
		bReturn = FALSE;
	}
	if (mp_bHideBar != FALSE)
	{
		bReturn = FALSE;
	}
	if (mp_bRollup != FALSE)
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
	if (mp_lPhysicalPercentComplete != 0)
	{
		bReturn = FALSE;
	}
	if (mp_yEarnedValueMethod != EVM_PERCENT_COMPLETE)
	{
		bReturn = FALSE;
	}
	if (mp_oPredecessorLink_C->IsNull() == FALSE)
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
	if (mp_oExtendedAttribute_C->IsNull() == FALSE)
	{
		bReturn = FALSE;
	}
	if (mp_oBaseline_C->IsNull() == FALSE)
	{
		bReturn = FALSE;
	}
	if (mp_oOutlineCode_C->IsNull() == FALSE)
	{
		bReturn = FALSE;
	}
	if (mp_bIsPublished != FALSE)
	{
		bReturn = FALSE;
	}
	if (mp_sStatusManager != "")
	{
		bReturn = FALSE;
	}
	if (mp_dtCommitmentStart.m_dt != 36494)
	{
		bReturn = FALSE;
	}
	if (mp_dtCommitmentFinish.m_dt != 36494)
	{
		bReturn = FALSE;
	}
	if (mp_yCommitmentType != CT_THE_TASK_HAS_NO_DELIVERABLE_OR_DEPENDENCY_ON_A_DELIVERABLE)
	{
		bReturn = FALSE;
	}
	if (mp_oTimephasedData_C->IsNull() == FALSE)
	{
		bReturn = FALSE;
	}
	return bReturn;
}

CString Task::GetXML(void)
{
	if (IsNull() == TRUE)
	{
		return "<Task/>";
	}
	clsXML oXML("Task");
	oXML.InitializeWriter();
	oXML.SetSupportOptional(TRUE);
	oXML.SetBoolsAreNumeric(TRUE);
	oXML.WriteProperty("UID", mp_lUID);
	oXML.WriteProperty("ID", mp_lID);
	if (mp_sName != "")
	{
		oXML.WriteProperty("Name", mp_sName);
	}
	oXML.WriteProperty("Type", mp_yType);
	oXML.WriteProperty("IsNull", mp_bIsNull);
	if (mp_dtCreateDate.m_dt != 36494)
	{
		oXML.WriteProperty("CreateDate", mp_dtCreateDate);
	}
	if (mp_sContact != "")
	{
		oXML.WriteProperty("Contact", mp_sContact);
	}
	if (mp_sWBS != "")
	{
		oXML.WriteProperty("WBS", mp_sWBS);
	}
	if (mp_sWBSLevel != "")
	{
		oXML.WriteProperty("WBSLevel", mp_sWBSLevel);
	}
	if (mp_sOutlineNumber != "")
	{
		oXML.WriteProperty("OutlineNumber", mp_sOutlineNumber);
	}
	oXML.WriteProperty("OutlineLevel", mp_lOutlineLevel);
	oXML.WriteProperty("Priority", mp_lPriority);
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
	oXML.WriteProperty("Work", mp_oWork);
	if (mp_dtStop.m_dt != 36494)
	{
		oXML.WriteProperty("Stop", mp_dtStop);
	}
	if (mp_dtResume.m_dt != 36494)
	{
		oXML.WriteProperty("Resume", mp_dtResume);
	}
	oXML.WriteProperty("ResumeValid", mp_bResumeValid);
	oXML.WriteProperty("EffortDriven", mp_bEffortDriven);
	oXML.WriteProperty("Recurring", mp_bRecurring);
	oXML.WriteProperty("OverAllocated", mp_bOverAllocated);
	oXML.WriteProperty("Estimated", mp_bEstimated);
	oXML.WriteProperty("Milestone", mp_bMilestone);
	oXML.WriteProperty("Summary", mp_bSummary);
	oXML.WriteProperty("Critical", mp_bCritical);
	oXML.WriteProperty("IsSubproject", mp_bIsSubproject);
	oXML.WriteProperty("IsSubprojectReadOnly", mp_bIsSubprojectReadOnly);
	if (mp_sSubprojectName != "")
	{
		oXML.WriteProperty("SubprojectName", mp_sSubprojectName);
	}
	oXML.WriteProperty("ExternalTask", mp_bExternalTask);
	if (mp_sExternalTaskProject != "")
	{
		oXML.WriteProperty("ExternalTaskProject", mp_sExternalTaskProject);
	}
	if (mp_dtEarlyStart.m_dt != 36494)
	{
		oXML.WriteProperty("EarlyStart", mp_dtEarlyStart);
	}
	if (mp_dtEarlyFinish.m_dt != 36494)
	{
		oXML.WriteProperty("EarlyFinish", mp_dtEarlyFinish);
	}
	if (mp_dtLateStart.m_dt != 36494)
	{
		oXML.WriteProperty("LateStart", mp_dtLateStart);
	}
	if (mp_dtLateFinish.m_dt != 36494)
	{
		oXML.WriteProperty("LateFinish", mp_dtLateFinish);
	}
	oXML.WriteProperty("StartVariance", mp_lStartVariance);
	oXML.WriteProperty("FinishVariance", mp_lFinishVariance);
	oXML.WriteProperty("WorkVariance", mp_fWorkVariance);
	oXML.WriteProperty("FreeSlack", mp_lFreeSlack);
	oXML.WriteProperty("TotalSlack", mp_lTotalSlack);
	oXML.WriteProperty("FixedCost", mp_fFixedCost);
	oXML.WriteProperty("FixedCostAccrual", mp_yFixedCostAccrual);
	oXML.WriteProperty("PercentComplete", mp_lPercentComplete);
	oXML.WriteProperty("PercentWorkComplete", mp_lPercentWorkComplete);
	oXML.WriteProperty("Cost", mp_cCost);
	oXML.WriteProperty("OvertimeCost", mp_cOvertimeCost);
	oXML.WriteProperty("OvertimeWork", mp_oOvertimeWork);
	if (mp_dtActualStart.m_dt != 36494)
	{
		oXML.WriteProperty("ActualStart", mp_dtActualStart);
	}
	if (mp_dtActualFinish.m_dt != 36494)
	{
		oXML.WriteProperty("ActualFinish", mp_dtActualFinish);
	}
	oXML.WriteProperty("ActualDuration", mp_oActualDuration);
	oXML.WriteProperty("ActualCost", mp_cActualCost);
	oXML.WriteProperty("ActualOvertimeCost", mp_cActualOvertimeCost);
	oXML.WriteProperty("ActualWork", mp_oActualWork);
	oXML.WriteProperty("ActualOvertimeWork", mp_oActualOvertimeWork);
	oXML.WriteProperty("RegularWork", mp_oRegularWork);
	oXML.WriteProperty("RemainingDuration", mp_oRemainingDuration);
	oXML.WriteProperty("RemainingCost", mp_cRemainingCost);
	oXML.WriteProperty("RemainingWork", mp_oRemainingWork);
	oXML.WriteProperty("RemainingOvertimeCost", mp_cRemainingOvertimeCost);
	oXML.WriteProperty("RemainingOvertimeWork", mp_oRemainingOvertimeWork);
	oXML.WriteProperty("ACWP", mp_fACWP);
	oXML.WriteProperty("CV", mp_fCV);
	oXML.WriteProperty("ConstraintType", mp_yConstraintType);
	oXML.WriteProperty("CalendarUID", mp_lCalendarUID);
	if (mp_dtConstraintDate.m_dt != 36494)
	{
		oXML.WriteProperty("ConstraintDate", mp_dtConstraintDate);
	}
	if (mp_dtDeadline.m_dt != 36494)
	{
		oXML.WriteProperty("Deadline", mp_dtDeadline);
	}
	oXML.WriteProperty("LevelAssignments", mp_bLevelAssignments);
	oXML.WriteProperty("LevelingCanSplit", mp_bLevelingCanSplit);
	oXML.WriteProperty("LevelingDelay", mp_lLevelingDelay);
	oXML.WriteProperty("LevelingDelayFormat", mp_yLevelingDelayFormat);
	if (mp_dtPreLeveledStart.m_dt != 36494)
	{
		oXML.WriteProperty("PreLeveledStart", mp_dtPreLeveledStart);
	}
	if (mp_dtPreLeveledFinish.m_dt != 36494)
	{
		oXML.WriteProperty("PreLeveledFinish", mp_dtPreLeveledFinish);
	}
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
	oXML.WriteProperty("IgnoreResourceCalendar", mp_bIgnoreResourceCalendar);
	if (mp_sNotes != "")
	{
		oXML.WriteProperty("Notes", mp_sNotes);
	}
	oXML.WriteProperty("HideBar", mp_bHideBar);
	oXML.WriteProperty("Rollup", mp_bRollup);
	oXML.WriteProperty("BCWS", mp_fBCWS);
	oXML.WriteProperty("BCWP", mp_fBCWP);
	oXML.WriteProperty("PhysicalPercentComplete", mp_lPhysicalPercentComplete);
	oXML.WriteProperty("EarnedValueMethod", mp_yEarnedValueMethod);
	if (mp_oPredecessorLink_C->IsNull() == FALSE)
	{
		mp_oPredecessorLink_C->WriteObjectProtected(oXML);
	}
	if (mp_oActualWorkProtected->IsNull() == FALSE)
	{
		oXML.WriteProperty("ActualWorkProtected", mp_oActualWorkProtected);
	}
	if (mp_oActualOvertimeWorkProtected->IsNull() == FALSE)
	{
		oXML.WriteProperty("ActualOvertimeWorkProtected", mp_oActualOvertimeWorkProtected);
	}
	if (mp_oExtendedAttribute_C->IsNull() == FALSE)
	{
		mp_oExtendedAttribute_C->WriteObjectProtected(oXML);
	}
	if (mp_oBaseline_C->IsNull() == FALSE)
	{
		mp_oBaseline_C->WriteObjectProtected(oXML);
	}
	if (mp_oOutlineCode_C->IsNull() == FALSE)
	{
		mp_oOutlineCode_C->WriteObjectProtected(oXML);
	}
	oXML.WriteProperty("IsPublished", mp_bIsPublished);
	if (mp_sStatusManager != "")
	{
		oXML.WriteProperty("StatusManager", mp_sStatusManager);
	}
	if (mp_dtCommitmentStart.m_dt != 36494)
	{
		oXML.WriteProperty("CommitmentStart", mp_dtCommitmentStart);
	}
	if (mp_dtCommitmentFinish.m_dt != 36494)
	{
		oXML.WriteProperty("CommitmentFinish", mp_dtCommitmentFinish);
	}
	oXML.WriteProperty("CommitmentType", mp_yCommitmentType);
	if (mp_oTimephasedData_C->IsNull() == FALSE)
	{
		mp_oTimephasedData_C->WriteObjectProtected(oXML);
	}
	return oXML.GetXML();
}

void Task::SetXML(CString sXML)
{
	clsXML oXML("Task");
	oXML.SetSupportOptional(TRUE);
	oXML.SetXML(sXML);
	oXML.InitializeReader();
	oXML.ReadProperty("UID", mp_lUID);
	oXML.ReadProperty("ID", mp_lID);
	oXML.ReadProperty("Name", mp_sName);
	if (mp_sName.GetLength() > 512)
	{
		mp_sName = mp_sName.Mid(0, 512);
	}
	oXML.ReadProperty("Type", mp_yType);
	oXML.ReadProperty("IsNull", mp_bIsNull);
	oXML.ReadProperty("CreateDate", mp_dtCreateDate);
	oXML.ReadProperty("Contact", mp_sContact);
	if (mp_sContact.GetLength() > 512)
	{
		mp_sContact = mp_sContact.Mid(0, 512);
	}
	oXML.ReadProperty("WBS", mp_sWBS);
	oXML.ReadProperty("WBSLevel", mp_sWBSLevel);
	oXML.ReadProperty("OutlineNumber", mp_sOutlineNumber);
	if (mp_sOutlineNumber.GetLength() > 512)
	{
		mp_sOutlineNumber = mp_sOutlineNumber.Mid(0, 512);
	}
	oXML.ReadProperty("OutlineLevel", mp_lOutlineLevel);
	oXML.ReadProperty("Priority", mp_lPriority);
	oXML.ReadProperty("Start", mp_dtStart);
	oXML.ReadProperty("Finish", mp_dtFinish);
	oXML.ReadProperty("Duration", mp_oDuration);
	oXML.ReadProperty("DurationFormat", mp_yDurationFormat);
	oXML.ReadProperty("Work", mp_oWork);
	oXML.ReadProperty("Stop", mp_dtStop);
	oXML.ReadProperty("Resume", mp_dtResume);
	oXML.ReadProperty("ResumeValid", mp_bResumeValid);
	oXML.ReadProperty("EffortDriven", mp_bEffortDriven);
	oXML.ReadProperty("Recurring", mp_bRecurring);
	oXML.ReadProperty("OverAllocated", mp_bOverAllocated);
	oXML.ReadProperty("Estimated", mp_bEstimated);
	oXML.ReadProperty("Milestone", mp_bMilestone);
	oXML.ReadProperty("Summary", mp_bSummary);
	oXML.ReadProperty("Critical", mp_bCritical);
	oXML.ReadProperty("IsSubproject", mp_bIsSubproject);
	oXML.ReadProperty("IsSubprojectReadOnly", mp_bIsSubprojectReadOnly);
	oXML.ReadProperty("SubprojectName", mp_sSubprojectName);
	if (mp_sSubprojectName.GetLength() > 512)
	{
		mp_sSubprojectName = mp_sSubprojectName.Mid(0, 512);
	}
	oXML.ReadProperty("ExternalTask", mp_bExternalTask);
	oXML.ReadProperty("ExternalTaskProject", mp_sExternalTaskProject);
	if (mp_sExternalTaskProject.GetLength() > 512)
	{
		mp_sExternalTaskProject = mp_sExternalTaskProject.Mid(0, 512);
	}
	oXML.ReadProperty("EarlyStart", mp_dtEarlyStart);
	oXML.ReadProperty("EarlyFinish", mp_dtEarlyFinish);
	oXML.ReadProperty("LateStart", mp_dtLateStart);
	oXML.ReadProperty("LateFinish", mp_dtLateFinish);
	oXML.ReadProperty("StartVariance", mp_lStartVariance);
	oXML.ReadProperty("FinishVariance", mp_lFinishVariance);
	oXML.ReadProperty("WorkVariance", mp_fWorkVariance);
	oXML.ReadProperty("FreeSlack", mp_lFreeSlack);
	oXML.ReadProperty("TotalSlack", mp_lTotalSlack);
	oXML.ReadProperty("FixedCost", mp_fFixedCost);
	oXML.ReadProperty("FixedCostAccrual", mp_yFixedCostAccrual);
	oXML.ReadProperty("PercentComplete", mp_lPercentComplete);
	oXML.ReadProperty("PercentWorkComplete", mp_lPercentWorkComplete);
	oXML.ReadProperty("Cost", mp_cCost);
	oXML.ReadProperty("OvertimeCost", mp_cOvertimeCost);
	oXML.ReadProperty("OvertimeWork", mp_oOvertimeWork);
	oXML.ReadProperty("ActualStart", mp_dtActualStart);
	oXML.ReadProperty("ActualFinish", mp_dtActualFinish);
	oXML.ReadProperty("ActualDuration", mp_oActualDuration);
	oXML.ReadProperty("ActualCost", mp_cActualCost);
	oXML.ReadProperty("ActualOvertimeCost", mp_cActualOvertimeCost);
	oXML.ReadProperty("ActualWork", mp_oActualWork);
	oXML.ReadProperty("ActualOvertimeWork", mp_oActualOvertimeWork);
	oXML.ReadProperty("RegularWork", mp_oRegularWork);
	oXML.ReadProperty("RemainingDuration", mp_oRemainingDuration);
	oXML.ReadProperty("RemainingCost", mp_cRemainingCost);
	oXML.ReadProperty("RemainingWork", mp_oRemainingWork);
	oXML.ReadProperty("RemainingOvertimeCost", mp_cRemainingOvertimeCost);
	oXML.ReadProperty("RemainingOvertimeWork", mp_oRemainingOvertimeWork);
	oXML.ReadProperty("ACWP", mp_fACWP);
	oXML.ReadProperty("CV", mp_fCV);
	oXML.ReadProperty("ConstraintType", mp_yConstraintType);
	oXML.ReadProperty("CalendarUID", mp_lCalendarUID);
	oXML.ReadProperty("ConstraintDate", mp_dtConstraintDate);
	oXML.ReadProperty("Deadline", mp_dtDeadline);
	oXML.ReadProperty("LevelAssignments", mp_bLevelAssignments);
	oXML.ReadProperty("LevelingCanSplit", mp_bLevelingCanSplit);
	oXML.ReadProperty("LevelingDelay", mp_lLevelingDelay);
	oXML.ReadProperty("LevelingDelayFormat", mp_yLevelingDelayFormat);
	oXML.ReadProperty("PreLeveledStart", mp_dtPreLeveledStart);
	oXML.ReadProperty("PreLeveledFinish", mp_dtPreLeveledFinish);
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
	oXML.ReadProperty("IgnoreResourceCalendar", mp_bIgnoreResourceCalendar);
	oXML.ReadProperty("Notes", mp_sNotes);
	oXML.ReadProperty("HideBar", mp_bHideBar);
	oXML.ReadProperty("Rollup", mp_bRollup);
	oXML.ReadProperty("BCWS", mp_fBCWS);
	oXML.ReadProperty("BCWP", mp_fBCWP);
	oXML.ReadProperty("PhysicalPercentComplete", mp_lPhysicalPercentComplete);
	oXML.ReadProperty("EarnedValueMethod", mp_yEarnedValueMethod);
	mp_oPredecessorLink_C->ReadObjectProtected(oXML);
	oXML.ReadProperty("ActualWorkProtected", mp_oActualWorkProtected);
	oXML.ReadProperty("ActualOvertimeWorkProtected", mp_oActualOvertimeWorkProtected);
	mp_oExtendedAttribute_C->ReadObjectProtected(oXML);
	mp_oBaseline_C->ReadObjectProtected(oXML);
	mp_oOutlineCode_C->ReadObjectProtected(oXML);
	oXML.ReadProperty("IsPublished", mp_bIsPublished);
	oXML.ReadProperty("StatusManager", mp_sStatusManager);
	oXML.ReadProperty("CommitmentStart", mp_dtCommitmentStart);
	oXML.ReadProperty("CommitmentFinish", mp_dtCommitmentFinish);
	oXML.ReadProperty("CommitmentType", mp_yCommitmentType);
	mp_oTimephasedData_C->ReadObjectProtected(oXML);
}