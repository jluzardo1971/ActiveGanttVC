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
#include "MP14.h"

IMPLEMENT_DYNCREATE(MP14, CCmdTarget)

//{919D066D-95AD-4664-8948-ADEB96CE7517}
static const IID IID_IMP14 = { 0x919D066D, 0x95AD, 0x4664, { 0x89, 0x48, 0xAD, 0xEB, 0x96, 0xCE, 0x75, 0x17} };

//{F5A0AC24-D255-4C91-83F2-F57178AC4C29}
IMPLEMENT_OLECREATE_FLAGS(MP14, "MSP2010.MP14", afxRegApartmentThreading, 0xF5A0AC24, 0xD255, 0x4C91, 0x83, 0xF2, 0xF5, 0x71, 0x78, 0xAC, 0x4C, 0x29)

BEGIN_DISPATCH_MAP(MP14, CCmdTarget)
DISP_PROPERTY_EX_ID(MP14, "lSaveVersion", 1, odl_GetSaveVersion, odl_SetSaveVersion, VT_I4)
DISP_PROPERTY_EX_ID(MP14, "sUID", 2, odl_GetUID, odl_SetUID, VT_BSTR)
DISP_PROPERTY_EX_ID(MP14, "sName", 3, odl_GetName, odl_SetName, VT_BSTR)
DISP_PROPERTY_EX_ID(MP14, "sTitle", 4, odl_GetTitle, odl_SetTitle, VT_BSTR)
DISP_PROPERTY_EX_ID(MP14, "sSubject", 5, odl_GetSubject, odl_SetSubject, VT_BSTR)
DISP_PROPERTY_EX_ID(MP14, "sCategory", 6, odl_GetCategory, odl_SetCategory, VT_BSTR)
DISP_PROPERTY_EX_ID(MP14, "sCompany", 7, odl_GetCompany, odl_SetCompany, VT_BSTR)
DISP_PROPERTY_EX_ID(MP14, "sManager", 8, odl_GetManager, odl_SetManager, VT_BSTR)
DISP_PROPERTY_EX_ID(MP14, "sAuthor", 9, odl_GetAuthor, odl_SetAuthor, VT_BSTR)
DISP_PROPERTY_EX_ID(MP14, "dtCreationDate", 10, odl_GetCreationDate, odl_SetCreationDate, VT_DATE)
DISP_PROPERTY_EX_ID(MP14, "lRevision", 11, odl_GetRevision, odl_SetRevision, VT_I4)
DISP_PROPERTY_EX_ID(MP14, "dtLastSaved", 12, odl_GetLastSaved, odl_SetLastSaved, VT_DATE)
DISP_PROPERTY_EX_ID(MP14, "bScheduleFromStart", 13, odl_GetScheduleFromStart, odl_SetScheduleFromStart, VT_BOOL)
DISP_PROPERTY_EX_ID(MP14, "dtStartDate", 14, odl_GetStartDate, odl_SetStartDate, VT_DATE)
DISP_PROPERTY_EX_ID(MP14, "dtFinishDate", 15, odl_GetFinishDate, odl_SetFinishDate, VT_DATE)
DISP_PROPERTY_EX_ID(MP14, "yFYStartDate", 16, odl_GetFYStartDate, odl_SetFYStartDate, VT_I4)
DISP_PROPERTY_EX_ID(MP14, "lCriticalSlackLimit", 17, odl_GetCriticalSlackLimit, odl_SetCriticalSlackLimit, VT_I4)
DISP_PROPERTY_EX_ID(MP14, "lCurrencyDigits", 18, odl_GetCurrencyDigits, odl_SetCurrencyDigits, VT_I4)
DISP_PROPERTY_EX_ID(MP14, "sCurrencySymbol", 19, odl_GetCurrencySymbol, odl_SetCurrencySymbol, VT_BSTR)
DISP_PROPERTY_EX_ID(MP14, "sCurrencyCode", 20, odl_GetCurrencyCode, odl_SetCurrencyCode, VT_BSTR)
DISP_PROPERTY_EX_ID(MP14, "yCurrencySymbolPosition", 21, odl_GetCurrencySymbolPosition, odl_SetCurrencySymbolPosition, VT_I4)
DISP_PROPERTY_EX_ID(MP14, "lCalendarUID", 22, odl_GetCalendarUID, odl_SetCalendarUID, VT_I4)
DISP_PROPERTY_EX_ID(MP14, "oDefaultStartTime", 23, odl_GetDefaultStartTime, SetNotSupported, VT_DISPATCH)
DISP_PROPERTY_EX_ID(MP14, "oDefaultFinishTime", 24, odl_GetDefaultFinishTime, SetNotSupported, VT_DISPATCH)
DISP_PROPERTY_EX_ID(MP14, "lMinutesPerDay", 25, odl_GetMinutesPerDay, odl_SetMinutesPerDay, VT_I4)
DISP_PROPERTY_EX_ID(MP14, "lMinutesPerWeek", 26, odl_GetMinutesPerWeek, odl_SetMinutesPerWeek, VT_I4)
DISP_PROPERTY_EX_ID(MP14, "lDaysPerMonth", 27, odl_GetDaysPerMonth, odl_SetDaysPerMonth, VT_I4)
DISP_PROPERTY_EX_ID(MP14, "yDefaultTaskType", 28, odl_GetDefaultTaskType, odl_SetDefaultTaskType, VT_I4)
DISP_PROPERTY_EX_ID(MP14, "yDefaultFixedCostAccrual", 29, odl_GetDefaultFixedCostAccrual, odl_SetDefaultFixedCostAccrual, VT_I4)
DISP_PROPERTY_EX_ID(MP14, "fDefaultStandardRate", 30, odl_GetDefaultStandardRate, odl_SetDefaultStandardRate, VT_R4)
DISP_PROPERTY_EX_ID(MP14, "fDefaultOvertimeRate", 31, odl_GetDefaultOvertimeRate, odl_SetDefaultOvertimeRate, VT_R4)
DISP_PROPERTY_EX_ID(MP14, "yDurationFormat", 32, odl_GetDurationFormat, odl_SetDurationFormat, VT_I4)
DISP_PROPERTY_EX_ID(MP14, "yWorkFormat", 33, odl_GetWorkFormat, odl_SetWorkFormat, VT_I4)
DISP_PROPERTY_EX_ID(MP14, "bEditableActualCosts", 34, odl_GetEditableActualCosts, odl_SetEditableActualCosts, VT_BOOL)
DISP_PROPERTY_EX_ID(MP14, "bHonorConstraints", 35, odl_GetHonorConstraints, odl_SetHonorConstraints, VT_BOOL)
DISP_PROPERTY_EX_ID(MP14, "yEarnedValueMethod", 36, odl_GetEarnedValueMethod, odl_SetEarnedValueMethod, VT_I4)
DISP_PROPERTY_EX_ID(MP14, "bInsertedProjectsLikeSummary", 37, odl_GetInsertedProjectsLikeSummary, odl_SetInsertedProjectsLikeSummary, VT_BOOL)
DISP_PROPERTY_EX_ID(MP14, "bMultipleCriticalPaths", 38, odl_GetMultipleCriticalPaths, odl_SetMultipleCriticalPaths, VT_BOOL)
DISP_PROPERTY_EX_ID(MP14, "bNewTasksEffortDriven", 39, odl_GetNewTasksEffortDriven, odl_SetNewTasksEffortDriven, VT_BOOL)
DISP_PROPERTY_EX_ID(MP14, "bNewTasksEstimated", 40, odl_GetNewTasksEstimated, odl_SetNewTasksEstimated, VT_BOOL)
DISP_PROPERTY_EX_ID(MP14, "bSplitsInProgressTasks", 41, odl_GetSplitsInProgressTasks, odl_SetSplitsInProgressTasks, VT_BOOL)
DISP_PROPERTY_EX_ID(MP14, "bSpreadActualCost", 42, odl_GetSpreadActualCost, odl_SetSpreadActualCost, VT_BOOL)
DISP_PROPERTY_EX_ID(MP14, "bSpreadPercentComplete", 43, odl_GetSpreadPercentComplete, odl_SetSpreadPercentComplete, VT_BOOL)
DISP_PROPERTY_EX_ID(MP14, "bTaskUpdatesResource", 44, odl_GetTaskUpdatesResource, odl_SetTaskUpdatesResource, VT_BOOL)
DISP_PROPERTY_EX_ID(MP14, "bFiscalYearStart", 45, odl_GetFiscalYearStart, odl_SetFiscalYearStart, VT_BOOL)
DISP_PROPERTY_EX_ID(MP14, "yWeekStartDay", 46, odl_GetWeekStartDay, odl_SetWeekStartDay, VT_I4)
DISP_PROPERTY_EX_ID(MP14, "bMoveCompletedEndsBack", 47, odl_GetMoveCompletedEndsBack, odl_SetMoveCompletedEndsBack, VT_BOOL)
DISP_PROPERTY_EX_ID(MP14, "bMoveRemainingStartsBack", 48, odl_GetMoveRemainingStartsBack, odl_SetMoveRemainingStartsBack, VT_BOOL)
DISP_PROPERTY_EX_ID(MP14, "bMoveRemainingStartsForward", 49, odl_GetMoveRemainingStartsForward, odl_SetMoveRemainingStartsForward, VT_BOOL)
DISP_PROPERTY_EX_ID(MP14, "bMoveCompletedEndsForward", 50, odl_GetMoveCompletedEndsForward, odl_SetMoveCompletedEndsForward, VT_BOOL)
DISP_PROPERTY_EX_ID(MP14, "yBaselineForEarnedValue", 51, odl_GetBaselineForEarnedValue, odl_SetBaselineForEarnedValue, VT_I4)
DISP_PROPERTY_EX_ID(MP14, "bAutoAddNewResourcesAndTasks", 52, odl_GetAutoAddNewResourcesAndTasks, odl_SetAutoAddNewResourcesAndTasks, VT_BOOL)
DISP_PROPERTY_EX_ID(MP14, "dtStatusDate", 53, odl_GetStatusDate, odl_SetStatusDate, VT_DATE)
DISP_PROPERTY_EX_ID(MP14, "dtCurrentDate", 54, odl_GetCurrentDate, odl_SetCurrentDate, VT_DATE)
DISP_PROPERTY_EX_ID(MP14, "bMicrosoftProjectServerURL", 55, odl_GetMicrosoftProjectServerURL, odl_SetMicrosoftProjectServerURL, VT_BOOL)
DISP_PROPERTY_EX_ID(MP14, "bAutolink", 56, odl_GetAutolink, odl_SetAutolink, VT_BOOL)
DISP_PROPERTY_EX_ID(MP14, "yNewTaskStartDate", 57, odl_GetNewTaskStartDate, odl_SetNewTaskStartDate, VT_I4)
DISP_PROPERTY_EX_ID(MP14, "yDefaultTaskEVMethod", 58, odl_GetDefaultTaskEVMethod, odl_SetDefaultTaskEVMethod, VT_I4)
DISP_PROPERTY_EX_ID(MP14, "bProjectExternallyEdited", 59, odl_GetProjectExternallyEdited, odl_SetProjectExternallyEdited, VT_BOOL)
DISP_PROPERTY_EX_ID(MP14, "dtExtendedCreationDate", 60, odl_GetExtendedCreationDate, odl_SetExtendedCreationDate, VT_DATE)
DISP_PROPERTY_EX_ID(MP14, "bActualsInSync", 61, odl_GetActualsInSync, odl_SetActualsInSync, VT_BOOL)
DISP_PROPERTY_EX_ID(MP14, "bRemoveFileProperties", 62, odl_GetRemoveFileProperties, odl_SetRemoveFileProperties, VT_BOOL)
DISP_PROPERTY_EX_ID(MP14, "bAdminProject", 63, odl_GetAdminProject, odl_SetAdminProject, VT_BOOL)
DISP_PROPERTY_EX_ID(MP14, "sBaslineCalendar", 64, odl_GetBaslineCalendar, odl_SetBaslineCalendar, VT_BSTR)
DISP_PROPERTY_EX_ID(MP14, "bNewTasksAreManual", 65, odl_GetNewTasksAreManual, odl_SetNewTasksAreManual, VT_BOOL)
DISP_PROPERTY_EX_ID(MP14, "bUpdateManuallyScheduledTasksWhenEditingLinks", 66, odl_GetUpdateManuallyScheduledTasksWhenEditingLinks, odl_SetUpdateManuallyScheduledTasksWhenEditingLinks, VT_BOOL)
DISP_PROPERTY_EX_ID(MP14, "bKeepTaskOnNearestWorkingTimeWhenMadeAutoScheduled", 67, odl_GetKeepTaskOnNearestWorkingTimeWhenMadeAutoScheduled, odl_SetKeepTaskOnNearestWorkingTimeWhenMadeAutoScheduled, VT_BOOL)
DISP_PROPERTY_EX_ID(MP14, "oOutlineCodes", 68, odl_GetOutlineCodes, SetNotSupported, VT_DISPATCH)
DISP_PROPERTY_EX_ID(MP14, "oWBSMasks", 69, odl_GetWBSMasks, SetNotSupported, VT_DISPATCH)
DISP_PROPERTY_EX_ID(MP14, "oExtendedAttributes", 70, odl_GetExtendedAttributes, SetNotSupported, VT_DISPATCH)
DISP_PROPERTY_EX_ID(MP14, "oCalendars", 71, odl_GetCalendars, SetNotSupported, VT_DISPATCH)
DISP_PROPERTY_EX_ID(MP14, "oTasks", 72, odl_GetTasks, SetNotSupported, VT_DISPATCH)
DISP_PROPERTY_EX_ID(MP14, "oResources", 73, odl_GetResources, SetNotSupported, VT_DISPATCH)
DISP_PROPERTY_EX_ID(MP14, "oAssignments", 74, odl_GetAssignments, SetNotSupported, VT_DISPATCH)
DISP_FUNCTION_ID(MP14, "WriteXML", 75, odl_WriteXML, VT_EMPTY, VTS_BSTR)
DISP_FUNCTION_ID(MP14, "ReadXML", 76, odl_ReadXML, VT_EMPTY, VTS_BSTR)
DISP_FUNCTION_ID(MP14, "GetXML", 77, odl_GetXML, VT_BSTR, VTS_NONE)
DISP_FUNCTION_ID(MP14, "SetXML", 78, odl_SetXML, VT_EMPTY, VTS_BSTR)
DISP_FUNCTION_ID(MP14, "IsNull", 79, IsNull, VT_BOOL, VTS_NONE)
DISP_FUNCTION_ID(MP14, "Initialize", 80, Initialize, VT_EMPTY, VTS_NONE)
END_DISPATCH_MAP()

BEGIN_INTERFACE_MAP(MP14, CCmdTarget)
	INTERFACE_PART(MP14, IID_IMP14, Dispatch)
END_INTERFACE_MAP()

BEGIN_MESSAGE_MAP(MP14, CCmdTarget)
END_MESSAGE_MAP()


void MP14::Initialize(void)
{
	InitVars();
}

MP14::MP14()
{
	EnableAutomation();
	AfxOleLockApp();
	InitVars();
}

void MP14::InitVars(void)
{
	mp_lSaveVersion = 0;
	mp_sUID = "";
	mp_sName = "";
	mp_sTitle = "";
	mp_sSubject = "";
	mp_sCategory = "";
	mp_sCompany = "";
	mp_sManager = "";
	mp_sAuthor = "";
	mp_dtCreationDate = (DATE) 0;
	mp_lRevision = 0;
	mp_dtLastSaved = (DATE) 0;
	mp_bScheduleFromStart = TRUE;
	mp_dtStartDate = (DATE) 0;
	mp_dtFinishDate = (DATE) 0;
	mp_yFYStartDate = FYSD_JANUARY;
	mp_lCriticalSlackLimit = 0;
	mp_lCurrencyDigits = 0;
	mp_sCurrencySymbol = "";
	mp_sCurrencyCode = "";
	mp_yCurrencySymbolPosition = CSP_BEFORE;
	mp_lCalendarUID = 0;
	mp_oDefaultStartTime = new Time();
	mp_oDefaultFinishTime = new Time();
	mp_lMinutesPerDay = 0;
	mp_lMinutesPerWeek = 0;
	mp_lDaysPerMonth = 0;
	mp_yDefaultTaskType = DTT_FIXED_DURATION;
	mp_yDefaultFixedCostAccrual = DFCA_START;
	mp_fDefaultStandardRate = 0;
	mp_fDefaultOvertimeRate = 0;
	mp_yDurationFormat = DF_M;
	mp_yWorkFormat = WF_M;
	mp_bEditableActualCosts = FALSE;
	mp_bHonorConstraints = TRUE;
	mp_yEarnedValueMethod = EVM_PERCENT_COMPLETE;
	mp_bInsertedProjectsLikeSummary = TRUE;
	mp_bMultipleCriticalPaths = FALSE;
	mp_bNewTasksEffortDriven = TRUE;
	mp_bNewTasksEstimated = TRUE;
	mp_bSplitsInProgressTasks = TRUE;
	mp_bSpreadActualCost = TRUE;
	mp_bSpreadPercentComplete = FALSE;
	mp_bTaskUpdatesResource = FALSE;
	mp_bFiscalYearStart = FALSE;
	mp_yWeekStartDay = WSD_SUNDAY;
	mp_bMoveCompletedEndsBack = FALSE;
	mp_bMoveRemainingStartsBack = FALSE;
	mp_bMoveRemainingStartsForward = FALSE;
	mp_bMoveCompletedEndsForward = FALSE;
	mp_yBaselineForEarnedValue = BFEV_BASELINE;
	mp_bAutoAddNewResourcesAndTasks = TRUE;
	mp_dtStatusDate = (DATE) 0;
	mp_dtCurrentDate = (DATE) 0;
	mp_bMicrosoftProjectServerURL = FALSE;
	mp_bAutolink = FALSE;
	mp_yNewTaskStartDate = NTSD_PROJECT_START_DATE;
	mp_yDefaultTaskEVMethod = DTEVM_PERCENT_COMPLETE;
	mp_bProjectExternallyEdited = FALSE;
	mp_dtExtendedCreationDate = (DATE) 0;
	mp_bActualsInSync = FALSE;
	mp_bRemoveFileProperties = FALSE;
	mp_bAdminProject = FALSE;
	mp_sBaslineCalendar = "";
	mp_bNewTasksAreManual = FALSE;
	mp_bUpdateManuallyScheduledTasksWhenEditingLinks = FALSE;
	mp_bKeepTaskOnNearestWorkingTimeWhenMadeAutoScheduled = FALSE;
	mp_oOutlineCodes = new OutlineCodes();
	mp_oWBSMasks = new WBSMasks();
	mp_oExtendedAttributes = new ExtendedAttributes();
	mp_oCalendars = new Calendars();
	mp_oTasks = new Tasks();
	mp_oResources = new Resources();
	mp_oAssignments = new Assignments();
}

MP14::~MP14()
{
	delete mp_oDefaultStartTime;
	delete mp_oDefaultFinishTime;
	delete mp_oOutlineCodes;
	delete mp_oWBSMasks;
	delete mp_oExtendedAttributes;
	delete mp_oCalendars;
	delete mp_oTasks;
	delete mp_oResources;
	delete mp_oAssignments;
	AfxOleUnlockApp();
}

void MP14::OnFinalRelease()
{
	CCmdTarget::OnFinalRelease();
}

LONG MP14::odl_GetSaveVersion(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetSaveVersion();
}

LONG MP14::GetSaveVersion(void)
{
	return (LONG) mp_lSaveVersion;
}

void MP14::odl_SetSaveVersion(LONG newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetSaveVersion((LONG)newVal);
}

void MP14::SetSaveVersion(LONG newVal)
{
	mp_lSaveVersion = newVal;
}

BSTR MP14::odl_GetUID(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetUID().AllocSysString();
}

CString MP14::GetUID(void)
{
	return (CString) mp_sUID;
}

void MP14::odl_SetUID(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetUID(newVal);
}

void MP14::SetUID(CString newVal)
{
	if (newVal.GetLength() > 16)
	{
		newVal = newVal.Mid(0, 16);
	}
	mp_sUID = newVal;
}

BSTR MP14::odl_GetName(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetName().AllocSysString();
}

CString MP14::GetName(void)
{
	return (CString) mp_sName;
}

void MP14::odl_SetName(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetName(newVal);
}

void MP14::SetName(CString newVal)
{
	if (newVal.GetLength() > 255)
	{
		newVal = newVal.Mid(0, 255);
	}
	mp_sName = newVal;
}

BSTR MP14::odl_GetTitle(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetTitle().AllocSysString();
}

CString MP14::GetTitle(void)
{
	return (CString) mp_sTitle;
}

void MP14::odl_SetTitle(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetTitle(newVal);
}

void MP14::SetTitle(CString newVal)
{
	if (newVal.GetLength() > 512)
	{
		newVal = newVal.Mid(0, 512);
	}
	mp_sTitle = newVal;
}

BSTR MP14::odl_GetSubject(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetSubject().AllocSysString();
}

CString MP14::GetSubject(void)
{
	return (CString) mp_sSubject;
}

void MP14::odl_SetSubject(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetSubject(newVal);
}

void MP14::SetSubject(CString newVal)
{
	if (newVal.GetLength() > 512)
	{
		newVal = newVal.Mid(0, 512);
	}
	mp_sSubject = newVal;
}

BSTR MP14::odl_GetCategory(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetCategory().AllocSysString();
}

CString MP14::GetCategory(void)
{
	return (CString) mp_sCategory;
}

void MP14::odl_SetCategory(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetCategory(newVal);
}

void MP14::SetCategory(CString newVal)
{
	if (newVal.GetLength() > 512)
	{
		newVal = newVal.Mid(0, 512);
	}
	mp_sCategory = newVal;
}

BSTR MP14::odl_GetCompany(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetCompany().AllocSysString();
}

CString MP14::GetCompany(void)
{
	return (CString) mp_sCompany;
}

void MP14::odl_SetCompany(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetCompany(newVal);
}

void MP14::SetCompany(CString newVal)
{
	if (newVal.GetLength() > 512)
	{
		newVal = newVal.Mid(0, 512);
	}
	mp_sCompany = newVal;
}

BSTR MP14::odl_GetManager(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetManager().AllocSysString();
}

CString MP14::GetManager(void)
{
	return (CString) mp_sManager;
}

void MP14::odl_SetManager(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetManager(newVal);
}

void MP14::SetManager(CString newVal)
{
	if (newVal.GetLength() > 512)
	{
		newVal = newVal.Mid(0, 512);
	}
	mp_sManager = newVal;
}

BSTR MP14::odl_GetAuthor(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetAuthor().AllocSysString();
}

CString MP14::GetAuthor(void)
{
	return (CString) mp_sAuthor;
}

void MP14::odl_SetAuthor(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetAuthor(newVal);
}

void MP14::SetAuthor(CString newVal)
{
	if (newVal.GetLength() > 512)
	{
		newVal = newVal.Mid(0, 512);
	}
	mp_sAuthor = newVal;
}

DATE MP14::odl_GetCreationDate(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetCreationDate();
}

COleDateTime MP14::GetCreationDate(void)
{
	return (COleDateTime) mp_dtCreationDate;
}

void MP14::odl_SetCreationDate(DATE newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetCreationDate((COleDateTime)newVal);
}

void MP14::SetCreationDate(COleDateTime newVal)
{
	mp_dtCreationDate = newVal;
}

LONG MP14::odl_GetRevision(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetRevision();
}

LONG MP14::GetRevision(void)
{
	return (LONG) mp_lRevision;
}

void MP14::odl_SetRevision(LONG newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetRevision((LONG)newVal);
}

void MP14::SetRevision(LONG newVal)
{
	mp_lRevision = newVal;
}

DATE MP14::odl_GetLastSaved(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetLastSaved();
}

COleDateTime MP14::GetLastSaved(void)
{
	return (COleDateTime) mp_dtLastSaved;
}

void MP14::odl_SetLastSaved(DATE newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetLastSaved((COleDateTime)newVal);
}

void MP14::SetLastSaved(COleDateTime newVal)
{
	mp_dtLastSaved = newVal;
}

BOOL MP14::odl_GetScheduleFromStart(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetScheduleFromStart();
}

BOOL MP14::GetScheduleFromStart(void)
{
	return (BOOL) mp_bScheduleFromStart;
}

void MP14::odl_SetScheduleFromStart(BOOL newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetScheduleFromStart((BOOL)newVal);
}

void MP14::SetScheduleFromStart(BOOL newVal)
{
	mp_bScheduleFromStart = newVal;
}

DATE MP14::odl_GetStartDate(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetStartDate();
}

COleDateTime MP14::GetStartDate(void)
{
	return (COleDateTime) mp_dtStartDate;
}

void MP14::odl_SetStartDate(DATE newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetStartDate((COleDateTime)newVal);
}

void MP14::SetStartDate(COleDateTime newVal)
{
	mp_dtStartDate = newVal;
}

DATE MP14::odl_GetFinishDate(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetFinishDate();
}

COleDateTime MP14::GetFinishDate(void)
{
	return (COleDateTime) mp_dtFinishDate;
}

void MP14::odl_SetFinishDate(DATE newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetFinishDate((COleDateTime)newVal);
}

void MP14::SetFinishDate(COleDateTime newVal)
{
	mp_dtFinishDate = newVal;
}

LONG MP14::odl_GetFYStartDate(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetFYStartDate();
}

E_FYSTARTDATE MP14::GetFYStartDate(void)
{
	return (E_FYSTARTDATE) mp_yFYStartDate;
}

void MP14::odl_SetFYStartDate(LONG newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetFYStartDate((E_FYSTARTDATE)newVal);
}

void MP14::SetFYStartDate(E_FYSTARTDATE newVal)
{
	mp_yFYStartDate = newVal;
}

LONG MP14::odl_GetCriticalSlackLimit(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetCriticalSlackLimit();
}

LONG MP14::GetCriticalSlackLimit(void)
{
	return (LONG) mp_lCriticalSlackLimit;
}

void MP14::odl_SetCriticalSlackLimit(LONG newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetCriticalSlackLimit((LONG)newVal);
}

void MP14::SetCriticalSlackLimit(LONG newVal)
{
	mp_lCriticalSlackLimit = newVal;
}

LONG MP14::odl_GetCurrencyDigits(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetCurrencyDigits();
}

LONG MP14::GetCurrencyDigits(void)
{
	return (LONG) mp_lCurrencyDigits;
}

void MP14::odl_SetCurrencyDigits(LONG newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetCurrencyDigits((LONG)newVal);
}

void MP14::SetCurrencyDigits(LONG newVal)
{
	mp_lCurrencyDigits = newVal;
}

BSTR MP14::odl_GetCurrencySymbol(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetCurrencySymbol().AllocSysString();
}

CString MP14::GetCurrencySymbol(void)
{
	return (CString) mp_sCurrencySymbol;
}

void MP14::odl_SetCurrencySymbol(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetCurrencySymbol(newVal);
}

void MP14::SetCurrencySymbol(CString newVal)
{
	if (newVal.GetLength() > 20)
	{
		newVal = newVal.Mid(0, 20);
	}
	mp_sCurrencySymbol = newVal;
}

BSTR MP14::odl_GetCurrencyCode(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetCurrencyCode().AllocSysString();
}

CString MP14::GetCurrencyCode(void)
{
	return (CString) mp_sCurrencyCode;
}

void MP14::odl_SetCurrencyCode(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetCurrencyCode(newVal);
}

void MP14::SetCurrencyCode(CString newVal)
{
	if (newVal.GetLength() > 3)
	{
		newVal = newVal.Mid(0, 3);
	}
	mp_sCurrencyCode = newVal;
}

LONG MP14::odl_GetCurrencySymbolPosition(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetCurrencySymbolPosition();
}

E_CURRENCYSYMBOLPOSITION MP14::GetCurrencySymbolPosition(void)
{
	return (E_CURRENCYSYMBOLPOSITION) mp_yCurrencySymbolPosition;
}

void MP14::odl_SetCurrencySymbolPosition(LONG newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetCurrencySymbolPosition((E_CURRENCYSYMBOLPOSITION)newVal);
}

void MP14::SetCurrencySymbolPosition(E_CURRENCYSYMBOLPOSITION newVal)
{
	mp_yCurrencySymbolPosition = newVal;
}

LONG MP14::odl_GetCalendarUID(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetCalendarUID();
}

LONG MP14::GetCalendarUID(void)
{
	return (LONG) mp_lCalendarUID;
}

void MP14::odl_SetCalendarUID(LONG newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetCalendarUID((LONG)newVal);
}

void MP14::SetCalendarUID(LONG newVal)
{
	mp_lCalendarUID = newVal;
}

IDispatch* MP14::odl_GetDefaultStartTime(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetDefaultStartTime()->GetIDispatch(TRUE);
}

Time* MP14::GetDefaultStartTime(void)
{
	return mp_oDefaultStartTime;
}

IDispatch* MP14::odl_GetDefaultFinishTime(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetDefaultFinishTime()->GetIDispatch(TRUE);
}

Time* MP14::GetDefaultFinishTime(void)
{
	return mp_oDefaultFinishTime;
}

LONG MP14::odl_GetMinutesPerDay(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetMinutesPerDay();
}

LONG MP14::GetMinutesPerDay(void)
{
	return (LONG) mp_lMinutesPerDay;
}

void MP14::odl_SetMinutesPerDay(LONG newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetMinutesPerDay((LONG)newVal);
}

void MP14::SetMinutesPerDay(LONG newVal)
{
	mp_lMinutesPerDay = newVal;
}

LONG MP14::odl_GetMinutesPerWeek(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetMinutesPerWeek();
}

LONG MP14::GetMinutesPerWeek(void)
{
	return (LONG) mp_lMinutesPerWeek;
}

void MP14::odl_SetMinutesPerWeek(LONG newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetMinutesPerWeek((LONG)newVal);
}

void MP14::SetMinutesPerWeek(LONG newVal)
{
	mp_lMinutesPerWeek = newVal;
}

LONG MP14::odl_GetDaysPerMonth(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetDaysPerMonth();
}

LONG MP14::GetDaysPerMonth(void)
{
	return (LONG) mp_lDaysPerMonth;
}

void MP14::odl_SetDaysPerMonth(LONG newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetDaysPerMonth((LONG)newVal);
}

void MP14::SetDaysPerMonth(LONG newVal)
{
	mp_lDaysPerMonth = newVal;
}

LONG MP14::odl_GetDefaultTaskType(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetDefaultTaskType();
}

E_DEFAULTTASKTYPE MP14::GetDefaultTaskType(void)
{
	return (E_DEFAULTTASKTYPE) mp_yDefaultTaskType;
}

void MP14::odl_SetDefaultTaskType(LONG newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetDefaultTaskType((E_DEFAULTTASKTYPE)newVal);
}

void MP14::SetDefaultTaskType(E_DEFAULTTASKTYPE newVal)
{
	mp_yDefaultTaskType = newVal;
}

LONG MP14::odl_GetDefaultFixedCostAccrual(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetDefaultFixedCostAccrual();
}

E_DEFAULTFIXEDCOSTACCRUAL MP14::GetDefaultFixedCostAccrual(void)
{
	return (E_DEFAULTFIXEDCOSTACCRUAL) mp_yDefaultFixedCostAccrual;
}

void MP14::odl_SetDefaultFixedCostAccrual(LONG newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetDefaultFixedCostAccrual((E_DEFAULTFIXEDCOSTACCRUAL)newVal);
}

void MP14::SetDefaultFixedCostAccrual(E_DEFAULTFIXEDCOSTACCRUAL newVal)
{
	mp_yDefaultFixedCostAccrual = newVal;
}

FLOAT MP14::odl_GetDefaultStandardRate(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetDefaultStandardRate();
}

FLOAT MP14::GetDefaultStandardRate(void)
{
	return (FLOAT) mp_fDefaultStandardRate;
}

void MP14::odl_SetDefaultStandardRate(FLOAT newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetDefaultStandardRate((FLOAT)newVal);
}

void MP14::SetDefaultStandardRate(FLOAT newVal)
{
	mp_fDefaultStandardRate = newVal;
}

FLOAT MP14::odl_GetDefaultOvertimeRate(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetDefaultOvertimeRate();
}

FLOAT MP14::GetDefaultOvertimeRate(void)
{
	return (FLOAT) mp_fDefaultOvertimeRate;
}

void MP14::odl_SetDefaultOvertimeRate(FLOAT newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetDefaultOvertimeRate((FLOAT)newVal);
}

void MP14::SetDefaultOvertimeRate(FLOAT newVal)
{
	mp_fDefaultOvertimeRate = newVal;
}

LONG MP14::odl_GetDurationFormat(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetDurationFormat();
}

E_DURATIONFORMAT MP14::GetDurationFormat(void)
{
	return (E_DURATIONFORMAT) mp_yDurationFormat;
}

void MP14::odl_SetDurationFormat(LONG newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetDurationFormat((E_DURATIONFORMAT)newVal);
}

void MP14::SetDurationFormat(E_DURATIONFORMAT newVal)
{
	mp_yDurationFormat = newVal;
}

LONG MP14::odl_GetWorkFormat(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetWorkFormat();
}

E_WORKFORMAT MP14::GetWorkFormat(void)
{
	return (E_WORKFORMAT) mp_yWorkFormat;
}

void MP14::odl_SetWorkFormat(LONG newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetWorkFormat((E_WORKFORMAT)newVal);
}

void MP14::SetWorkFormat(E_WORKFORMAT newVal)
{
	mp_yWorkFormat = newVal;
}

BOOL MP14::odl_GetEditableActualCosts(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetEditableActualCosts();
}

BOOL MP14::GetEditableActualCosts(void)
{
	return (BOOL) mp_bEditableActualCosts;
}

void MP14::odl_SetEditableActualCosts(BOOL newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetEditableActualCosts((BOOL)newVal);
}

void MP14::SetEditableActualCosts(BOOL newVal)
{
	mp_bEditableActualCosts = newVal;
}

BOOL MP14::odl_GetHonorConstraints(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetHonorConstraints();
}

BOOL MP14::GetHonorConstraints(void)
{
	return (BOOL) mp_bHonorConstraints;
}

void MP14::odl_SetHonorConstraints(BOOL newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetHonorConstraints((BOOL)newVal);
}

void MP14::SetHonorConstraints(BOOL newVal)
{
	mp_bHonorConstraints = newVal;
}

LONG MP14::odl_GetEarnedValueMethod(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetEarnedValueMethod();
}

E_EARNEDVALUEMETHOD MP14::GetEarnedValueMethod(void)
{
	return (E_EARNEDVALUEMETHOD) mp_yEarnedValueMethod;
}

void MP14::odl_SetEarnedValueMethod(LONG newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetEarnedValueMethod((E_EARNEDVALUEMETHOD)newVal);
}

void MP14::SetEarnedValueMethod(E_EARNEDVALUEMETHOD newVal)
{
	mp_yEarnedValueMethod = newVal;
}

BOOL MP14::odl_GetInsertedProjectsLikeSummary(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetInsertedProjectsLikeSummary();
}

BOOL MP14::GetInsertedProjectsLikeSummary(void)
{
	return (BOOL) mp_bInsertedProjectsLikeSummary;
}

void MP14::odl_SetInsertedProjectsLikeSummary(BOOL newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetInsertedProjectsLikeSummary((BOOL)newVal);
}

void MP14::SetInsertedProjectsLikeSummary(BOOL newVal)
{
	mp_bInsertedProjectsLikeSummary = newVal;
}

BOOL MP14::odl_GetMultipleCriticalPaths(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetMultipleCriticalPaths();
}

BOOL MP14::GetMultipleCriticalPaths(void)
{
	return (BOOL) mp_bMultipleCriticalPaths;
}

void MP14::odl_SetMultipleCriticalPaths(BOOL newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetMultipleCriticalPaths((BOOL)newVal);
}

void MP14::SetMultipleCriticalPaths(BOOL newVal)
{
	mp_bMultipleCriticalPaths = newVal;
}

BOOL MP14::odl_GetNewTasksEffortDriven(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetNewTasksEffortDriven();
}

BOOL MP14::GetNewTasksEffortDriven(void)
{
	return (BOOL) mp_bNewTasksEffortDriven;
}

void MP14::odl_SetNewTasksEffortDriven(BOOL newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetNewTasksEffortDriven((BOOL)newVal);
}

void MP14::SetNewTasksEffortDriven(BOOL newVal)
{
	mp_bNewTasksEffortDriven = newVal;
}

BOOL MP14::odl_GetNewTasksEstimated(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetNewTasksEstimated();
}

BOOL MP14::GetNewTasksEstimated(void)
{
	return (BOOL) mp_bNewTasksEstimated;
}

void MP14::odl_SetNewTasksEstimated(BOOL newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetNewTasksEstimated((BOOL)newVal);
}

void MP14::SetNewTasksEstimated(BOOL newVal)
{
	mp_bNewTasksEstimated = newVal;
}

BOOL MP14::odl_GetSplitsInProgressTasks(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetSplitsInProgressTasks();
}

BOOL MP14::GetSplitsInProgressTasks(void)
{
	return (BOOL) mp_bSplitsInProgressTasks;
}

void MP14::odl_SetSplitsInProgressTasks(BOOL newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetSplitsInProgressTasks((BOOL)newVal);
}

void MP14::SetSplitsInProgressTasks(BOOL newVal)
{
	mp_bSplitsInProgressTasks = newVal;
}

BOOL MP14::odl_GetSpreadActualCost(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetSpreadActualCost();
}

BOOL MP14::GetSpreadActualCost(void)
{
	return (BOOL) mp_bSpreadActualCost;
}

void MP14::odl_SetSpreadActualCost(BOOL newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetSpreadActualCost((BOOL)newVal);
}

void MP14::SetSpreadActualCost(BOOL newVal)
{
	mp_bSpreadActualCost = newVal;
}

BOOL MP14::odl_GetSpreadPercentComplete(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetSpreadPercentComplete();
}

BOOL MP14::GetSpreadPercentComplete(void)
{
	return (BOOL) mp_bSpreadPercentComplete;
}

void MP14::odl_SetSpreadPercentComplete(BOOL newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetSpreadPercentComplete((BOOL)newVal);
}

void MP14::SetSpreadPercentComplete(BOOL newVal)
{
	mp_bSpreadPercentComplete = newVal;
}

BOOL MP14::odl_GetTaskUpdatesResource(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetTaskUpdatesResource();
}

BOOL MP14::GetTaskUpdatesResource(void)
{
	return (BOOL) mp_bTaskUpdatesResource;
}

void MP14::odl_SetTaskUpdatesResource(BOOL newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetTaskUpdatesResource((BOOL)newVal);
}

void MP14::SetTaskUpdatesResource(BOOL newVal)
{
	mp_bTaskUpdatesResource = newVal;
}

BOOL MP14::odl_GetFiscalYearStart(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetFiscalYearStart();
}

BOOL MP14::GetFiscalYearStart(void)
{
	return (BOOL) mp_bFiscalYearStart;
}

void MP14::odl_SetFiscalYearStart(BOOL newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetFiscalYearStart((BOOL)newVal);
}

void MP14::SetFiscalYearStart(BOOL newVal)
{
	mp_bFiscalYearStart = newVal;
}

LONG MP14::odl_GetWeekStartDay(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetWeekStartDay();
}

E_WEEKSTARTDAY MP14::GetWeekStartDay(void)
{
	return (E_WEEKSTARTDAY) mp_yWeekStartDay;
}

void MP14::odl_SetWeekStartDay(LONG newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetWeekStartDay((E_WEEKSTARTDAY)newVal);
}

void MP14::SetWeekStartDay(E_WEEKSTARTDAY newVal)
{
	mp_yWeekStartDay = newVal;
}

BOOL MP14::odl_GetMoveCompletedEndsBack(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetMoveCompletedEndsBack();
}

BOOL MP14::GetMoveCompletedEndsBack(void)
{
	return (BOOL) mp_bMoveCompletedEndsBack;
}

void MP14::odl_SetMoveCompletedEndsBack(BOOL newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetMoveCompletedEndsBack((BOOL)newVal);
}

void MP14::SetMoveCompletedEndsBack(BOOL newVal)
{
	mp_bMoveCompletedEndsBack = newVal;
}

BOOL MP14::odl_GetMoveRemainingStartsBack(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetMoveRemainingStartsBack();
}

BOOL MP14::GetMoveRemainingStartsBack(void)
{
	return (BOOL) mp_bMoveRemainingStartsBack;
}

void MP14::odl_SetMoveRemainingStartsBack(BOOL newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetMoveRemainingStartsBack((BOOL)newVal);
}

void MP14::SetMoveRemainingStartsBack(BOOL newVal)
{
	mp_bMoveRemainingStartsBack = newVal;
}

BOOL MP14::odl_GetMoveRemainingStartsForward(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetMoveRemainingStartsForward();
}

BOOL MP14::GetMoveRemainingStartsForward(void)
{
	return (BOOL) mp_bMoveRemainingStartsForward;
}

void MP14::odl_SetMoveRemainingStartsForward(BOOL newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetMoveRemainingStartsForward((BOOL)newVal);
}

void MP14::SetMoveRemainingStartsForward(BOOL newVal)
{
	mp_bMoveRemainingStartsForward = newVal;
}

BOOL MP14::odl_GetMoveCompletedEndsForward(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetMoveCompletedEndsForward();
}

BOOL MP14::GetMoveCompletedEndsForward(void)
{
	return (BOOL) mp_bMoveCompletedEndsForward;
}

void MP14::odl_SetMoveCompletedEndsForward(BOOL newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetMoveCompletedEndsForward((BOOL)newVal);
}

void MP14::SetMoveCompletedEndsForward(BOOL newVal)
{
	mp_bMoveCompletedEndsForward = newVal;
}

LONG MP14::odl_GetBaselineForEarnedValue(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetBaselineForEarnedValue();
}

E_BASELINEFOREARNEDVALUE MP14::GetBaselineForEarnedValue(void)
{
	return (E_BASELINEFOREARNEDVALUE) mp_yBaselineForEarnedValue;
}

void MP14::odl_SetBaselineForEarnedValue(LONG newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetBaselineForEarnedValue((E_BASELINEFOREARNEDVALUE)newVal);
}

void MP14::SetBaselineForEarnedValue(E_BASELINEFOREARNEDVALUE newVal)
{
	mp_yBaselineForEarnedValue = newVal;
}

BOOL MP14::odl_GetAutoAddNewResourcesAndTasks(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetAutoAddNewResourcesAndTasks();
}

BOOL MP14::GetAutoAddNewResourcesAndTasks(void)
{
	return (BOOL) mp_bAutoAddNewResourcesAndTasks;
}

void MP14::odl_SetAutoAddNewResourcesAndTasks(BOOL newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetAutoAddNewResourcesAndTasks((BOOL)newVal);
}

void MP14::SetAutoAddNewResourcesAndTasks(BOOL newVal)
{
	mp_bAutoAddNewResourcesAndTasks = newVal;
}

DATE MP14::odl_GetStatusDate(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetStatusDate();
}

COleDateTime MP14::GetStatusDate(void)
{
	return (COleDateTime) mp_dtStatusDate;
}

void MP14::odl_SetStatusDate(DATE newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetStatusDate((COleDateTime)newVal);
}

void MP14::SetStatusDate(COleDateTime newVal)
{
	mp_dtStatusDate = newVal;
}

DATE MP14::odl_GetCurrentDate(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetCurrentDate();
}

COleDateTime MP14::GetCurrentDate(void)
{
	return (COleDateTime) mp_dtCurrentDate;
}

void MP14::odl_SetCurrentDate(DATE newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetCurrentDate((COleDateTime)newVal);
}

void MP14::SetCurrentDate(COleDateTime newVal)
{
	mp_dtCurrentDate = newVal;
}

BOOL MP14::odl_GetMicrosoftProjectServerURL(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetMicrosoftProjectServerURL();
}

BOOL MP14::GetMicrosoftProjectServerURL(void)
{
	return (BOOL) mp_bMicrosoftProjectServerURL;
}

void MP14::odl_SetMicrosoftProjectServerURL(BOOL newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetMicrosoftProjectServerURL((BOOL)newVal);
}

void MP14::SetMicrosoftProjectServerURL(BOOL newVal)
{
	mp_bMicrosoftProjectServerURL = newVal;
}

BOOL MP14::odl_GetAutolink(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetAutolink();
}

BOOL MP14::GetAutolink(void)
{
	return (BOOL) mp_bAutolink;
}

void MP14::odl_SetAutolink(BOOL newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetAutolink((BOOL)newVal);
}

void MP14::SetAutolink(BOOL newVal)
{
	mp_bAutolink = newVal;
}

LONG MP14::odl_GetNewTaskStartDate(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetNewTaskStartDate();
}

E_NEWTASKSTARTDATE MP14::GetNewTaskStartDate(void)
{
	return (E_NEWTASKSTARTDATE) mp_yNewTaskStartDate;
}

void MP14::odl_SetNewTaskStartDate(LONG newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetNewTaskStartDate((E_NEWTASKSTARTDATE)newVal);
}

void MP14::SetNewTaskStartDate(E_NEWTASKSTARTDATE newVal)
{
	mp_yNewTaskStartDate = newVal;
}

LONG MP14::odl_GetDefaultTaskEVMethod(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetDefaultTaskEVMethod();
}

E_DEFAULTTASKEVMETHOD MP14::GetDefaultTaskEVMethod(void)
{
	return (E_DEFAULTTASKEVMETHOD) mp_yDefaultTaskEVMethod;
}

void MP14::odl_SetDefaultTaskEVMethod(LONG newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetDefaultTaskEVMethod((E_DEFAULTTASKEVMETHOD)newVal);
}

void MP14::SetDefaultTaskEVMethod(E_DEFAULTTASKEVMETHOD newVal)
{
	mp_yDefaultTaskEVMethod = newVal;
}

BOOL MP14::odl_GetProjectExternallyEdited(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetProjectExternallyEdited();
}

BOOL MP14::GetProjectExternallyEdited(void)
{
	return (BOOL) mp_bProjectExternallyEdited;
}

void MP14::odl_SetProjectExternallyEdited(BOOL newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetProjectExternallyEdited((BOOL)newVal);
}

void MP14::SetProjectExternallyEdited(BOOL newVal)
{
	mp_bProjectExternallyEdited = newVal;
}

DATE MP14::odl_GetExtendedCreationDate(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetExtendedCreationDate();
}

COleDateTime MP14::GetExtendedCreationDate(void)
{
	return (COleDateTime) mp_dtExtendedCreationDate;
}

void MP14::odl_SetExtendedCreationDate(DATE newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetExtendedCreationDate((COleDateTime)newVal);
}

void MP14::SetExtendedCreationDate(COleDateTime newVal)
{
	mp_dtExtendedCreationDate = newVal;
}

BOOL MP14::odl_GetActualsInSync(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetActualsInSync();
}

BOOL MP14::GetActualsInSync(void)
{
	return (BOOL) mp_bActualsInSync;
}

void MP14::odl_SetActualsInSync(BOOL newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetActualsInSync((BOOL)newVal);
}

void MP14::SetActualsInSync(BOOL newVal)
{
	mp_bActualsInSync = newVal;
}

BOOL MP14::odl_GetRemoveFileProperties(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetRemoveFileProperties();
}

BOOL MP14::GetRemoveFileProperties(void)
{
	return (BOOL) mp_bRemoveFileProperties;
}

void MP14::odl_SetRemoveFileProperties(BOOL newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetRemoveFileProperties((BOOL)newVal);
}

void MP14::SetRemoveFileProperties(BOOL newVal)
{
	mp_bRemoveFileProperties = newVal;
}

BOOL MP14::odl_GetAdminProject(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetAdminProject();
}

BOOL MP14::GetAdminProject(void)
{
	return (BOOL) mp_bAdminProject;
}

void MP14::odl_SetAdminProject(BOOL newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetAdminProject((BOOL)newVal);
}

void MP14::SetAdminProject(BOOL newVal)
{
	mp_bAdminProject = newVal;
}

BSTR MP14::odl_GetBaslineCalendar(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetBaslineCalendar().AllocSysString();
}

CString MP14::GetBaslineCalendar(void)
{
	return (CString) mp_sBaslineCalendar;
}

void MP14::odl_SetBaslineCalendar(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetBaslineCalendar(newVal);
}

void MP14::SetBaslineCalendar(CString newVal)
{
	mp_sBaslineCalendar = newVal;
}

BOOL MP14::odl_GetNewTasksAreManual(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetNewTasksAreManual();
}

BOOL MP14::GetNewTasksAreManual(void)
{
	return (BOOL) mp_bNewTasksAreManual;
}

void MP14::odl_SetNewTasksAreManual(BOOL newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetNewTasksAreManual((BOOL)newVal);
}

void MP14::SetNewTasksAreManual(BOOL newVal)
{
	mp_bNewTasksAreManual = newVal;
}

BOOL MP14::odl_GetUpdateManuallyScheduledTasksWhenEditingLinks(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetUpdateManuallyScheduledTasksWhenEditingLinks();
}

BOOL MP14::GetUpdateManuallyScheduledTasksWhenEditingLinks(void)
{
	return (BOOL) mp_bUpdateManuallyScheduledTasksWhenEditingLinks;
}

void MP14::odl_SetUpdateManuallyScheduledTasksWhenEditingLinks(BOOL newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetUpdateManuallyScheduledTasksWhenEditingLinks((BOOL)newVal);
}

void MP14::SetUpdateManuallyScheduledTasksWhenEditingLinks(BOOL newVal)
{
	mp_bUpdateManuallyScheduledTasksWhenEditingLinks = newVal;
}

BOOL MP14::odl_GetKeepTaskOnNearestWorkingTimeWhenMadeAutoScheduled(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetKeepTaskOnNearestWorkingTimeWhenMadeAutoScheduled();
}

BOOL MP14::GetKeepTaskOnNearestWorkingTimeWhenMadeAutoScheduled(void)
{
	return (BOOL) mp_bKeepTaskOnNearestWorkingTimeWhenMadeAutoScheduled;
}

void MP14::odl_SetKeepTaskOnNearestWorkingTimeWhenMadeAutoScheduled(BOOL newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetKeepTaskOnNearestWorkingTimeWhenMadeAutoScheduled((BOOL)newVal);
}

void MP14::SetKeepTaskOnNearestWorkingTimeWhenMadeAutoScheduled(BOOL newVal)
{
	mp_bKeepTaskOnNearestWorkingTimeWhenMadeAutoScheduled = newVal;
}

IDispatch* MP14::odl_GetOutlineCodes(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetOutlineCodes()->GetIDispatch(TRUE);
}

OutlineCodes* MP14::GetOutlineCodes(void)
{
	return mp_oOutlineCodes;
}

IDispatch* MP14::odl_GetWBSMasks(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetWBSMasks()->GetIDispatch(TRUE);
}

WBSMasks* MP14::GetWBSMasks(void)
{
	return mp_oWBSMasks;
}

IDispatch* MP14::odl_GetExtendedAttributes(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetExtendedAttributes()->GetIDispatch(TRUE);
}

ExtendedAttributes* MP14::GetExtendedAttributes(void)
{
	return mp_oExtendedAttributes;
}

IDispatch* MP14::odl_GetCalendars(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetCalendars()->GetIDispatch(TRUE);
}

Calendars* MP14::GetCalendars(void)
{
	return mp_oCalendars;
}

IDispatch* MP14::odl_GetTasks(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetTasks()->GetIDispatch(TRUE);
}

Tasks* MP14::GetTasks(void)
{
	return mp_oTasks;
}

IDispatch* MP14::odl_GetResources(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetResources()->GetIDispatch(TRUE);
}

Resources* MP14::GetResources(void)
{
	return mp_oResources;
}

IDispatch* MP14::odl_GetAssignments(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetAssignments()->GetIDispatch(TRUE);
}

Assignments* MP14::GetAssignments(void)
{
	return mp_oAssignments;
}

void MP14::WriteXML(CString url)
{
	clsXML oXML("Project");
	mp_WriteXML(oXML);
	oXML.WriteXML(url);
}

void MP14::ReadXML(CString url)
{
	clsXML oXML("Project");
	oXML.ReadXML(url);
	oXML.RemoveNamespace();
	mp_ReadXML(oXML);
}

void MP14::SetXML(CString sXML)
{
	clsXML oXML("Project");
	oXML.SetXML(sXML);
	oXML.RemoveNamespace();
	mp_ReadXML(oXML);
}

CString MP14::GetXML()
{
	clsXML oXML("Project");
	mp_WriteXML(oXML);
	return oXML.GetXML();
}

void MP14::odl_WriteXML(LPCTSTR url)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	WriteXML(url);
}

void MP14::odl_ReadXML(LPCTSTR url)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	ReadXML(url);
}

void MP14::odl_SetXML(LPCTSTR sXML)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetXML(sXML);
}

BSTR MP14::odl_GetXML(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetXML().AllocSysString();
}

BOOL MP14::IsNull(void)
{
	BOOL bReturn = TRUE;
	if (mp_lSaveVersion != 0)
	{
		bReturn = FALSE;
	}
	if (mp_sUID != "")
	{
		bReturn = FALSE;
	}
	if (mp_sName != "")
	{
		bReturn = FALSE;
	}
	if (mp_sTitle != "")
	{
		bReturn = FALSE;
	}
	if (mp_sSubject != "")
	{
		bReturn = FALSE;
	}
	if (mp_sCategory != "")
	{
		bReturn = FALSE;
	}
	if (mp_sCompany != "")
	{
		bReturn = FALSE;
	}
	if (mp_sManager != "")
	{
		bReturn = FALSE;
	}
	if (mp_sAuthor != "")
	{
		bReturn = FALSE;
	}
	if (mp_dtCreationDate.m_dt != 36494)
	{
		bReturn = FALSE;
	}
	if (mp_lRevision != 0)
	{
		bReturn = FALSE;
	}
	if (mp_dtLastSaved.m_dt != 36494)
	{
		bReturn = FALSE;
	}
	if (mp_bScheduleFromStart != TRUE)
	{
		bReturn = FALSE;
	}
	if (mp_dtStartDate.m_dt != 36494)
	{
		bReturn = FALSE;
	}
	if (mp_dtFinishDate.m_dt != 36494)
	{
		bReturn = FALSE;
	}
	if (mp_yFYStartDate != FYSD_JANUARY)
	{
		bReturn = FALSE;
	}
	if (mp_lCriticalSlackLimit != 0)
	{
		bReturn = FALSE;
	}
	if (mp_lCurrencyDigits != 0)
	{
		bReturn = FALSE;
	}
	if (mp_sCurrencySymbol != "")
	{
		bReturn = FALSE;
	}
	if (mp_sCurrencyCode != "")
	{
		bReturn = FALSE;
	}
	if (mp_yCurrencySymbolPosition != CSP_BEFORE)
	{
		bReturn = FALSE;
	}
	if (mp_lCalendarUID != 0)
	{
		bReturn = FALSE;
	}
	if (mp_oDefaultStartTime->IsNull() == FALSE)
	{
		bReturn = FALSE;
	}
	if (mp_oDefaultFinishTime->IsNull() == FALSE)
	{
		bReturn = FALSE;
	}
	if (mp_lMinutesPerDay != 0)
	{
		bReturn = FALSE;
	}
	if (mp_lMinutesPerWeek != 0)
	{
		bReturn = FALSE;
	}
	if (mp_lDaysPerMonth != 0)
	{
		bReturn = FALSE;
	}
	if (mp_yDefaultTaskType != DTT_FIXED_DURATION)
	{
		bReturn = FALSE;
	}
	if (mp_yDefaultFixedCostAccrual != DFCA_START)
	{
		bReturn = FALSE;
	}
	if (mp_fDefaultStandardRate != 0)
	{
		bReturn = FALSE;
	}
	if (mp_fDefaultOvertimeRate != 0)
	{
		bReturn = FALSE;
	}
	if (mp_yDurationFormat != DF_M)
	{
		bReturn = FALSE;
	}
	if (mp_yWorkFormat != WF_M)
	{
		bReturn = FALSE;
	}
	if (mp_bEditableActualCosts != FALSE)
	{
		bReturn = FALSE;
	}
	if (mp_bHonorConstraints != TRUE)
	{
		bReturn = FALSE;
	}
	if (mp_yEarnedValueMethod != EVM_PERCENT_COMPLETE)
	{
		bReturn = FALSE;
	}
	if (mp_bInsertedProjectsLikeSummary != TRUE)
	{
		bReturn = FALSE;
	}
	if (mp_bMultipleCriticalPaths != FALSE)
	{
		bReturn = FALSE;
	}
	if (mp_bNewTasksEffortDriven != TRUE)
	{
		bReturn = FALSE;
	}
	if (mp_bNewTasksEstimated != TRUE)
	{
		bReturn = FALSE;
	}
	if (mp_bSplitsInProgressTasks != TRUE)
	{
		bReturn = FALSE;
	}
	if (mp_bSpreadActualCost != TRUE)
	{
		bReturn = FALSE;
	}
	if (mp_bSpreadPercentComplete != FALSE)
	{
		bReturn = FALSE;
	}
	if (mp_bTaskUpdatesResource != FALSE)
	{
		bReturn = FALSE;
	}
	if (mp_bFiscalYearStart != FALSE)
	{
		bReturn = FALSE;
	}
	if (mp_yWeekStartDay != WSD_SUNDAY)
	{
		bReturn = FALSE;
	}
	if (mp_bMoveCompletedEndsBack != FALSE)
	{
		bReturn = FALSE;
	}
	if (mp_bMoveRemainingStartsBack != FALSE)
	{
		bReturn = FALSE;
	}
	if (mp_bMoveRemainingStartsForward != FALSE)
	{
		bReturn = FALSE;
	}
	if (mp_bMoveCompletedEndsForward != FALSE)
	{
		bReturn = FALSE;
	}
	if (mp_yBaselineForEarnedValue != BFEV_BASELINE)
	{
		bReturn = FALSE;
	}
	if (mp_bAutoAddNewResourcesAndTasks != TRUE)
	{
		bReturn = FALSE;
	}
	if (mp_dtStatusDate.m_dt != 36494)
	{
		bReturn = FALSE;
	}
	if (mp_dtCurrentDate.m_dt != 36494)
	{
		bReturn = FALSE;
	}
	if (mp_bMicrosoftProjectServerURL != FALSE)
	{
		bReturn = FALSE;
	}
	if (mp_bAutolink != FALSE)
	{
		bReturn = FALSE;
	}
	if (mp_yNewTaskStartDate != NTSD_PROJECT_START_DATE)
	{
		bReturn = FALSE;
	}
	if (mp_yDefaultTaskEVMethod != DTEVM_PERCENT_COMPLETE)
	{
		bReturn = FALSE;
	}
	if (mp_bProjectExternallyEdited != FALSE)
	{
		bReturn = FALSE;
	}
	if (mp_dtExtendedCreationDate.m_dt != 36494)
	{
		bReturn = FALSE;
	}
	if (mp_bActualsInSync != FALSE)
	{
		bReturn = FALSE;
	}
	if (mp_bRemoveFileProperties != FALSE)
	{
		bReturn = FALSE;
	}
	if (mp_bAdminProject != FALSE)
	{
		bReturn = FALSE;
	}
	if (mp_sBaslineCalendar != "")
	{
		bReturn = FALSE;
	}
	if (mp_bNewTasksAreManual != FALSE)
	{
		bReturn = FALSE;
	}
	if (mp_bUpdateManuallyScheduledTasksWhenEditingLinks != FALSE)
	{
		bReturn = FALSE;
	}
	if (mp_bKeepTaskOnNearestWorkingTimeWhenMadeAutoScheduled != FALSE)
	{
		bReturn = FALSE;
	}
	if (mp_oOutlineCodes->IsNull() == FALSE)
	{
		bReturn = FALSE;
	}
	if (mp_oWBSMasks->IsNull() == FALSE)
	{
		bReturn = FALSE;
	}
	if (mp_oExtendedAttributes->IsNull() == FALSE)
	{
		bReturn = FALSE;
	}
	if (mp_oCalendars->IsNull() == FALSE)
	{
		bReturn = FALSE;
	}
	if (mp_oTasks->IsNull() == FALSE)
	{
		bReturn = FALSE;
	}
	if (mp_oResources->IsNull() == FALSE)
	{
		bReturn = FALSE;
	}
	if (mp_oAssignments->IsNull() == FALSE)
	{
		bReturn = FALSE;
	}
	return bReturn;
}

void MP14::mp_WriteXML(clsXML &oXML)
{
	oXML.InitializeWriter();
	oXML.SetSupportOptional(TRUE);
	oXML.SetBoolsAreNumeric(TRUE);
	oXML.WriteProperty("SaveVersion", mp_lSaveVersion);
	if (mp_sUID != "")
	{
		oXML.WriteProperty("UID", mp_sUID);
	}
	if (mp_sName != "")
	{
		oXML.WriteProperty("Name", mp_sName);
	}
	if (mp_sTitle != "")
	{
		oXML.WriteProperty("Title", mp_sTitle);
	}
	if (mp_sSubject != "")
	{
		oXML.WriteProperty("Subject", mp_sSubject);
	}
	if (mp_sCategory != "")
	{
		oXML.WriteProperty("Category", mp_sCategory);
	}
	if (mp_sCompany != "")
	{
		oXML.WriteProperty("Company", mp_sCompany);
	}
	if (mp_sManager != "")
	{
		oXML.WriteProperty("Manager", mp_sManager);
	}
	if (mp_sAuthor != "")
	{
		oXML.WriteProperty("Author", mp_sAuthor);
	}
	if (mp_dtCreationDate.m_dt != 36494)
	{
		oXML.WriteProperty("CreationDate", mp_dtCreationDate);
	}
	oXML.WriteProperty("Revision", mp_lRevision);
	if (mp_dtLastSaved.m_dt != 36494)
	{
		oXML.WriteProperty("LastSaved", mp_dtLastSaved);
	}
	oXML.WriteProperty("ScheduleFromStart", mp_bScheduleFromStart);
	if (mp_dtStartDate.m_dt != 36494)
	{
		oXML.WriteProperty("StartDate", mp_dtStartDate);
	}
	if (mp_dtFinishDate.m_dt != 36494)
	{
		oXML.WriteProperty("FinishDate", mp_dtFinishDate);
	}
	oXML.WriteProperty("FYStartDate", mp_yFYStartDate);
	oXML.WriteProperty("CriticalSlackLimit", mp_lCriticalSlackLimit);
	oXML.WriteProperty("CurrencyDigits", mp_lCurrencyDigits);
	if (mp_sCurrencySymbol != "")
	{
		oXML.WriteProperty("CurrencySymbol", mp_sCurrencySymbol);
	}
	oXML.WriteProperty("CurrencyCode", mp_sCurrencyCode);
	oXML.WriteProperty("CurrencySymbolPosition", mp_yCurrencySymbolPosition);
	oXML.WriteProperty("CalendarUID", mp_lCalendarUID);
	if (mp_oDefaultStartTime->IsNull() == FALSE)
	{
		oXML.WriteProperty("DefaultStartTime", mp_oDefaultStartTime);
	}
	if (mp_oDefaultFinishTime->IsNull() == FALSE)
	{
		oXML.WriteProperty("DefaultFinishTime", mp_oDefaultFinishTime);
	}
	oXML.WriteProperty("MinutesPerDay", mp_lMinutesPerDay);
	oXML.WriteProperty("MinutesPerWeek", mp_lMinutesPerWeek);
	oXML.WriteProperty("DaysPerMonth", mp_lDaysPerMonth);
	oXML.WriteProperty("DefaultTaskType", mp_yDefaultTaskType);
	oXML.WriteProperty("DefaultFixedCostAccrual", mp_yDefaultFixedCostAccrual);
	oXML.WriteProperty("DefaultStandardRate", mp_fDefaultStandardRate);
	oXML.WriteProperty("DefaultOvertimeRate", mp_fDefaultOvertimeRate);
	oXML.WriteProperty("DurationFormat", mp_yDurationFormat);
	oXML.WriteProperty("WorkFormat", mp_yWorkFormat);
	oXML.WriteProperty("EditableActualCosts", mp_bEditableActualCosts);
	oXML.WriteProperty("HonorConstraints", mp_bHonorConstraints);
	oXML.WriteProperty("EarnedValueMethod", mp_yEarnedValueMethod);
	oXML.WriteProperty("InsertedProjectsLikeSummary", mp_bInsertedProjectsLikeSummary);
	oXML.WriteProperty("MultipleCriticalPaths", mp_bMultipleCriticalPaths);
	oXML.WriteProperty("NewTasksEffortDriven", mp_bNewTasksEffortDriven);
	oXML.WriteProperty("NewTasksEstimated", mp_bNewTasksEstimated);
	oXML.WriteProperty("SplitsInProgressTasks", mp_bSplitsInProgressTasks);
	oXML.WriteProperty("SpreadActualCost", mp_bSpreadActualCost);
	oXML.WriteProperty("SpreadPercentComplete", mp_bSpreadPercentComplete);
	oXML.WriteProperty("TaskUpdatesResource", mp_bTaskUpdatesResource);
	oXML.WriteProperty("FiscalYearStart", mp_bFiscalYearStart);
	oXML.WriteProperty("WeekStartDay", mp_yWeekStartDay);
	oXML.WriteProperty("MoveCompletedEndsBack", mp_bMoveCompletedEndsBack);
	oXML.WriteProperty("MoveRemainingStartsBack", mp_bMoveRemainingStartsBack);
	oXML.WriteProperty("MoveRemainingStartsForward", mp_bMoveRemainingStartsForward);
	oXML.WriteProperty("MoveCompletedEndsForward", mp_bMoveCompletedEndsForward);
	oXML.WriteProperty("BaselineForEarnedValue", mp_yBaselineForEarnedValue);
	oXML.WriteProperty("AutoAddNewResourcesAndTasks", mp_bAutoAddNewResourcesAndTasks);
	if (mp_dtStatusDate.m_dt != 36494)
	{
		oXML.WriteProperty("StatusDate", mp_dtStatusDate);
	}
	if (mp_dtCurrentDate.m_dt != 36494)
	{
		oXML.WriteProperty("CurrentDate", mp_dtCurrentDate);
	}
	oXML.WriteProperty("MicrosoftProjectServerURL", mp_bMicrosoftProjectServerURL);
	oXML.WriteProperty("Autolink", mp_bAutolink);
	oXML.WriteProperty("NewTaskStartDate", mp_yNewTaskStartDate);
	oXML.WriteProperty("DefaultTaskEVMethod", mp_yDefaultTaskEVMethod);
	oXML.WriteProperty("ProjectExternallyEdited", mp_bProjectExternallyEdited);
	if (mp_dtExtendedCreationDate.m_dt != 36494)
	{
		oXML.WriteProperty("ExtendedCreationDate", mp_dtExtendedCreationDate);
	}
	oXML.WriteProperty("ActualsInSync", mp_bActualsInSync);
	oXML.WriteProperty("RemoveFileProperties", mp_bRemoveFileProperties);
	oXML.WriteProperty("AdminProject", mp_bAdminProject);
	if (mp_sBaslineCalendar != "")
	{
		oXML.WriteProperty("BaslineCalendar", mp_sBaslineCalendar);
	}
	oXML.WriteProperty("NewTasksAreManual", mp_bNewTasksAreManual);
	oXML.WriteProperty("UpdateManuallyScheduledTasksWhenEditingLinks", mp_bUpdateManuallyScheduledTasksWhenEditingLinks);
	oXML.WriteProperty("KeepTaskOnNearestWorkingTimeWhenMadeAutoScheduled", mp_bKeepTaskOnNearestWorkingTimeWhenMadeAutoScheduled);
	oXML.WriteObject(mp_oOutlineCodes->GetXML());
	oXML.WriteObject(mp_oWBSMasks->GetXML());
	oXML.WriteObject(mp_oExtendedAttributes->GetXML());
	oXML.WriteObject(mp_oCalendars->GetXML());
	oXML.WriteObject(mp_oTasks->GetXML());
	oXML.WriteObject(mp_oResources->GetXML());
	oXML.WriteObject(mp_oAssignments->GetXML());
	oXML.AddNamespace();
}

void MP14::mp_ReadXML(clsXML &oXML)
{
	oXML.SetSupportOptional(TRUE);
	oXML.InitializeReader();
	oXML.ReadProperty("SaveVersion", mp_lSaveVersion);
	oXML.ReadProperty("UID", mp_sUID);
	if (mp_sUID.GetLength() > 16)
	{
		mp_sUID = mp_sUID.Mid(0, 16);
	}
	oXML.ReadProperty("Name", mp_sName);
	if (mp_sName.GetLength() > 255)
	{
		mp_sName = mp_sName.Mid(0, 255);
	}
	oXML.ReadProperty("Title", mp_sTitle);
	if (mp_sTitle.GetLength() > 512)
	{
		mp_sTitle = mp_sTitle.Mid(0, 512);
	}
	oXML.ReadProperty("Subject", mp_sSubject);
	if (mp_sSubject.GetLength() > 512)
	{
		mp_sSubject = mp_sSubject.Mid(0, 512);
	}
	oXML.ReadProperty("Category", mp_sCategory);
	if (mp_sCategory.GetLength() > 512)
	{
		mp_sCategory = mp_sCategory.Mid(0, 512);
	}
	oXML.ReadProperty("Company", mp_sCompany);
	if (mp_sCompany.GetLength() > 512)
	{
		mp_sCompany = mp_sCompany.Mid(0, 512);
	}
	oXML.ReadProperty("Manager", mp_sManager);
	if (mp_sManager.GetLength() > 512)
	{
		mp_sManager = mp_sManager.Mid(0, 512);
	}
	oXML.ReadProperty("Author", mp_sAuthor);
	if (mp_sAuthor.GetLength() > 512)
	{
		mp_sAuthor = mp_sAuthor.Mid(0, 512);
	}
	oXML.ReadProperty("CreationDate", mp_dtCreationDate);
	oXML.ReadProperty("Revision", mp_lRevision);
	oXML.ReadProperty("LastSaved", mp_dtLastSaved);
	oXML.ReadProperty("ScheduleFromStart", mp_bScheduleFromStart);
	oXML.ReadProperty("StartDate", mp_dtStartDate);
	oXML.ReadProperty("FinishDate", mp_dtFinishDate);
	oXML.ReadProperty("FYStartDate", mp_yFYStartDate);
	oXML.ReadProperty("CriticalSlackLimit", mp_lCriticalSlackLimit);
	oXML.ReadProperty("CurrencyDigits", mp_lCurrencyDigits);
	oXML.ReadProperty("CurrencySymbol", mp_sCurrencySymbol);
	if (mp_sCurrencySymbol.GetLength() > 20)
	{
		mp_sCurrencySymbol = mp_sCurrencySymbol.Mid(0, 20);
	}
	oXML.ReadProperty("CurrencyCode", mp_sCurrencyCode);
	if (mp_sCurrencyCode.GetLength() > 3)
	{
		mp_sCurrencyCode = mp_sCurrencyCode.Mid(0, 3);
	}
	oXML.ReadProperty("CurrencySymbolPosition", mp_yCurrencySymbolPosition);
	oXML.ReadProperty("CalendarUID", mp_lCalendarUID);
	oXML.ReadProperty("DefaultStartTime", mp_oDefaultStartTime);
	oXML.ReadProperty("DefaultFinishTime", mp_oDefaultFinishTime);
	oXML.ReadProperty("MinutesPerDay", mp_lMinutesPerDay);
	oXML.ReadProperty("MinutesPerWeek", mp_lMinutesPerWeek);
	oXML.ReadProperty("DaysPerMonth", mp_lDaysPerMonth);
	oXML.ReadProperty("DefaultTaskType", mp_yDefaultTaskType);
	oXML.ReadProperty("DefaultFixedCostAccrual", mp_yDefaultFixedCostAccrual);
	oXML.ReadProperty("DefaultStandardRate", mp_fDefaultStandardRate);
	oXML.ReadProperty("DefaultOvertimeRate", mp_fDefaultOvertimeRate);
	oXML.ReadProperty("DurationFormat", mp_yDurationFormat);
	oXML.ReadProperty("WorkFormat", mp_yWorkFormat);
	oXML.ReadProperty("EditableActualCosts", mp_bEditableActualCosts);
	oXML.ReadProperty("HonorConstraints", mp_bHonorConstraints);
	oXML.ReadProperty("EarnedValueMethod", mp_yEarnedValueMethod);
	oXML.ReadProperty("InsertedProjectsLikeSummary", mp_bInsertedProjectsLikeSummary);
	oXML.ReadProperty("MultipleCriticalPaths", mp_bMultipleCriticalPaths);
	oXML.ReadProperty("NewTasksEffortDriven", mp_bNewTasksEffortDriven);
	oXML.ReadProperty("NewTasksEstimated", mp_bNewTasksEstimated);
	oXML.ReadProperty("SplitsInProgressTasks", mp_bSplitsInProgressTasks);
	oXML.ReadProperty("SpreadActualCost", mp_bSpreadActualCost);
	oXML.ReadProperty("SpreadPercentComplete", mp_bSpreadPercentComplete);
	oXML.ReadProperty("TaskUpdatesResource", mp_bTaskUpdatesResource);
	oXML.ReadProperty("FiscalYearStart", mp_bFiscalYearStart);
	oXML.ReadProperty("WeekStartDay", mp_yWeekStartDay);
	oXML.ReadProperty("MoveCompletedEndsBack", mp_bMoveCompletedEndsBack);
	oXML.ReadProperty("MoveRemainingStartsBack", mp_bMoveRemainingStartsBack);
	oXML.ReadProperty("MoveRemainingStartsForward", mp_bMoveRemainingStartsForward);
	oXML.ReadProperty("MoveCompletedEndsForward", mp_bMoveCompletedEndsForward);
	oXML.ReadProperty("BaselineForEarnedValue", mp_yBaselineForEarnedValue);
	oXML.ReadProperty("AutoAddNewResourcesAndTasks", mp_bAutoAddNewResourcesAndTasks);
	oXML.ReadProperty("StatusDate", mp_dtStatusDate);
	oXML.ReadProperty("CurrentDate", mp_dtCurrentDate);
	oXML.ReadProperty("MicrosoftProjectServerURL", mp_bMicrosoftProjectServerURL);
	oXML.ReadProperty("Autolink", mp_bAutolink);
	oXML.ReadProperty("NewTaskStartDate", mp_yNewTaskStartDate);
	oXML.ReadProperty("DefaultTaskEVMethod", mp_yDefaultTaskEVMethod);
	oXML.ReadProperty("ProjectExternallyEdited", mp_bProjectExternallyEdited);
	oXML.ReadProperty("ExtendedCreationDate", mp_dtExtendedCreationDate);
	oXML.ReadProperty("ActualsInSync", mp_bActualsInSync);
	oXML.ReadProperty("RemoveFileProperties", mp_bRemoveFileProperties);
	oXML.ReadProperty("AdminProject", mp_bAdminProject);
	oXML.ReadProperty("BaslineCalendar", mp_sBaslineCalendar);
	oXML.ReadProperty("NewTasksAreManual", mp_bNewTasksAreManual);
	oXML.ReadProperty("UpdateManuallyScheduledTasksWhenEditingLinks", mp_bUpdateManuallyScheduledTasksWhenEditingLinks);
	oXML.ReadProperty("KeepTaskOnNearestWorkingTimeWhenMadeAutoScheduled", mp_bKeepTaskOnNearestWorkingTimeWhenMadeAutoScheduled);
	mp_oOutlineCodes->SetXML(oXML.ReadObject("OutlineCodes"));
	mp_oWBSMasks->SetXML(oXML.ReadObject("WBSMasks"));
	mp_oExtendedAttributes->SetXML(oXML.ReadObject("ExtendedAttributes"));
	mp_oCalendars->SetXML(oXML.ReadObject("Calendars"));
	mp_oTasks->SetXML(oXML.ReadObject("Tasks"));
	mp_oResources->SetXML(oXML.ReadObject("Resources"));
	mp_oAssignments->SetXML(oXML.ReadObject("Assignments"));
}