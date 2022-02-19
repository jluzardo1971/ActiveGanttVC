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
#include "MP12.h"

IMPLEMENT_DYNCREATE(MP12, CCmdTarget)

//{556E86F5-A2B7-4291-9565-02128BF7414D}
static const IID IID_IMP12 = { 0x556E86F5, 0xA2B7, 0x4291, { 0x95, 0x65, 0x02, 0x12, 0x8B, 0xF7, 0x41, 0x4D} };

//{A1F4F09C-108A-4444-8674-6CB9CDC1817C}
IMPLEMENT_OLECREATE_FLAGS(MP12, "MSP2007.MP12", afxRegApartmentThreading, 0xA1F4F09C, 0x108A, 0x4444, 0x86, 0x74, 0x6C, 0xB9, 0xCD, 0xC1, 0x81, 0x7C)

BEGIN_DISPATCH_MAP(MP12, CCmdTarget)
DISP_PROPERTY_EX_ID(MP12, "lSaveVersion", 1, odl_GetSaveVersion, odl_SetSaveVersion, VT_I4)
DISP_PROPERTY_EX_ID(MP12, "sUID", 2, odl_GetUID, odl_SetUID, VT_BSTR)
DISP_PROPERTY_EX_ID(MP12, "sName", 3, odl_GetName, odl_SetName, VT_BSTR)
DISP_PROPERTY_EX_ID(MP12, "sTitle", 4, odl_GetTitle, odl_SetTitle, VT_BSTR)
DISP_PROPERTY_EX_ID(MP12, "sSubject", 5, odl_GetSubject, odl_SetSubject, VT_BSTR)
DISP_PROPERTY_EX_ID(MP12, "sCategory", 6, odl_GetCategory, odl_SetCategory, VT_BSTR)
DISP_PROPERTY_EX_ID(MP12, "sCompany", 7, odl_GetCompany, odl_SetCompany, VT_BSTR)
DISP_PROPERTY_EX_ID(MP12, "sManager", 8, odl_GetManager, odl_SetManager, VT_BSTR)
DISP_PROPERTY_EX_ID(MP12, "sAuthor", 9, odl_GetAuthor, odl_SetAuthor, VT_BSTR)
DISP_PROPERTY_EX_ID(MP12, "dtCreationDate", 10, odl_GetCreationDate, odl_SetCreationDate, VT_DATE)
DISP_PROPERTY_EX_ID(MP12, "lRevision", 11, odl_GetRevision, odl_SetRevision, VT_I4)
DISP_PROPERTY_EX_ID(MP12, "dtLastSaved", 12, odl_GetLastSaved, odl_SetLastSaved, VT_DATE)
DISP_PROPERTY_EX_ID(MP12, "bScheduleFromStart", 13, odl_GetScheduleFromStart, odl_SetScheduleFromStart, VT_BOOL)
DISP_PROPERTY_EX_ID(MP12, "dtStartDate", 14, odl_GetStartDate, odl_SetStartDate, VT_DATE)
DISP_PROPERTY_EX_ID(MP12, "dtFinishDate", 15, odl_GetFinishDate, odl_SetFinishDate, VT_DATE)
DISP_PROPERTY_EX_ID(MP12, "yFYStartDate", 16, odl_GetFYStartDate, odl_SetFYStartDate, VT_I4)
DISP_PROPERTY_EX_ID(MP12, "lCriticalSlackLimit", 17, odl_GetCriticalSlackLimit, odl_SetCriticalSlackLimit, VT_I4)
DISP_PROPERTY_EX_ID(MP12, "lCurrencyDigits", 18, odl_GetCurrencyDigits, odl_SetCurrencyDigits, VT_I4)
DISP_PROPERTY_EX_ID(MP12, "sCurrencySymbol", 19, odl_GetCurrencySymbol, odl_SetCurrencySymbol, VT_BSTR)
DISP_PROPERTY_EX_ID(MP12, "sCurrencyCode", 20, odl_GetCurrencyCode, odl_SetCurrencyCode, VT_BSTR)
DISP_PROPERTY_EX_ID(MP12, "yCurrencySymbolPosition", 21, odl_GetCurrencySymbolPosition, odl_SetCurrencySymbolPosition, VT_I4)
DISP_PROPERTY_EX_ID(MP12, "lCalendarUID", 22, odl_GetCalendarUID, odl_SetCalendarUID, VT_I4)
DISP_PROPERTY_EX_ID(MP12, "oDefaultStartTime", 23, odl_GetDefaultStartTime, SetNotSupported, VT_DISPATCH)
DISP_PROPERTY_EX_ID(MP12, "oDefaultFinishTime", 24, odl_GetDefaultFinishTime, SetNotSupported, VT_DISPATCH)
DISP_PROPERTY_EX_ID(MP12, "lMinutesPerDay", 25, odl_GetMinutesPerDay, odl_SetMinutesPerDay, VT_I4)
DISP_PROPERTY_EX_ID(MP12, "lMinutesPerWeek", 26, odl_GetMinutesPerWeek, odl_SetMinutesPerWeek, VT_I4)
DISP_PROPERTY_EX_ID(MP12, "lDaysPerMonth", 27, odl_GetDaysPerMonth, odl_SetDaysPerMonth, VT_I4)
DISP_PROPERTY_EX_ID(MP12, "yDefaultTaskType", 28, odl_GetDefaultTaskType, odl_SetDefaultTaskType, VT_I4)
DISP_PROPERTY_EX_ID(MP12, "yDefaultFixedCostAccrual", 29, odl_GetDefaultFixedCostAccrual, odl_SetDefaultFixedCostAccrual, VT_I4)
DISP_PROPERTY_EX_ID(MP12, "fDefaultStandardRate", 30, odl_GetDefaultStandardRate, odl_SetDefaultStandardRate, VT_R4)
DISP_PROPERTY_EX_ID(MP12, "fDefaultOvertimeRate", 31, odl_GetDefaultOvertimeRate, odl_SetDefaultOvertimeRate, VT_R4)
DISP_PROPERTY_EX_ID(MP12, "yDurationFormat", 32, odl_GetDurationFormat, odl_SetDurationFormat, VT_I4)
DISP_PROPERTY_EX_ID(MP12, "yWorkFormat", 33, odl_GetWorkFormat, odl_SetWorkFormat, VT_I4)
DISP_PROPERTY_EX_ID(MP12, "bEditableActualCosts", 34, odl_GetEditableActualCosts, odl_SetEditableActualCosts, VT_BOOL)
DISP_PROPERTY_EX_ID(MP12, "bHonorConstraints", 35, odl_GetHonorConstraints, odl_SetHonorConstraints, VT_BOOL)
DISP_PROPERTY_EX_ID(MP12, "yEarnedValueMethod", 36, odl_GetEarnedValueMethod, odl_SetEarnedValueMethod, VT_I4)
DISP_PROPERTY_EX_ID(MP12, "bInsertedProjectsLikeSummary", 37, odl_GetInsertedProjectsLikeSummary, odl_SetInsertedProjectsLikeSummary, VT_BOOL)
DISP_PROPERTY_EX_ID(MP12, "bMultipleCriticalPaths", 38, odl_GetMultipleCriticalPaths, odl_SetMultipleCriticalPaths, VT_BOOL)
DISP_PROPERTY_EX_ID(MP12, "bNewTasksEffortDriven", 39, odl_GetNewTasksEffortDriven, odl_SetNewTasksEffortDriven, VT_BOOL)
DISP_PROPERTY_EX_ID(MP12, "bNewTasksEstimated", 40, odl_GetNewTasksEstimated, odl_SetNewTasksEstimated, VT_BOOL)
DISP_PROPERTY_EX_ID(MP12, "bSplitsInProgressTasks", 41, odl_GetSplitsInProgressTasks, odl_SetSplitsInProgressTasks, VT_BOOL)
DISP_PROPERTY_EX_ID(MP12, "bSpreadActualCost", 42, odl_GetSpreadActualCost, odl_SetSpreadActualCost, VT_BOOL)
DISP_PROPERTY_EX_ID(MP12, "bSpreadPercentComplete", 43, odl_GetSpreadPercentComplete, odl_SetSpreadPercentComplete, VT_BOOL)
DISP_PROPERTY_EX_ID(MP12, "bTaskUpdatesResource", 44, odl_GetTaskUpdatesResource, odl_SetTaskUpdatesResource, VT_BOOL)
DISP_PROPERTY_EX_ID(MP12, "bFiscalYearStart", 45, odl_GetFiscalYearStart, odl_SetFiscalYearStart, VT_BOOL)
DISP_PROPERTY_EX_ID(MP12, "yWeekStartDay", 46, odl_GetWeekStartDay, odl_SetWeekStartDay, VT_I4)
DISP_PROPERTY_EX_ID(MP12, "bMoveCompletedEndsBack", 47, odl_GetMoveCompletedEndsBack, odl_SetMoveCompletedEndsBack, VT_BOOL)
DISP_PROPERTY_EX_ID(MP12, "bMoveRemainingStartsBack", 48, odl_GetMoveRemainingStartsBack, odl_SetMoveRemainingStartsBack, VT_BOOL)
DISP_PROPERTY_EX_ID(MP12, "bMoveRemainingStartsForward", 49, odl_GetMoveRemainingStartsForward, odl_SetMoveRemainingStartsForward, VT_BOOL)
DISP_PROPERTY_EX_ID(MP12, "bMoveCompletedEndsForward", 50, odl_GetMoveCompletedEndsForward, odl_SetMoveCompletedEndsForward, VT_BOOL)
DISP_PROPERTY_EX_ID(MP12, "yBaselineForEarnedValue", 51, odl_GetBaselineForEarnedValue, odl_SetBaselineForEarnedValue, VT_I4)
DISP_PROPERTY_EX_ID(MP12, "bAutoAddNewResourcesAndTasks", 52, odl_GetAutoAddNewResourcesAndTasks, odl_SetAutoAddNewResourcesAndTasks, VT_BOOL)
DISP_PROPERTY_EX_ID(MP12, "dtStatusDate", 53, odl_GetStatusDate, odl_SetStatusDate, VT_DATE)
DISP_PROPERTY_EX_ID(MP12, "dtCurrentDate", 54, odl_GetCurrentDate, odl_SetCurrentDate, VT_DATE)
DISP_PROPERTY_EX_ID(MP12, "bMicrosoftProjectServerURL", 55, odl_GetMicrosoftProjectServerURL, odl_SetMicrosoftProjectServerURL, VT_BOOL)
DISP_PROPERTY_EX_ID(MP12, "bAutolink", 56, odl_GetAutolink, odl_SetAutolink, VT_BOOL)
DISP_PROPERTY_EX_ID(MP12, "yNewTaskStartDate", 57, odl_GetNewTaskStartDate, odl_SetNewTaskStartDate, VT_I4)
DISP_PROPERTY_EX_ID(MP12, "yDefaultTaskEVMethod", 58, odl_GetDefaultTaskEVMethod, odl_SetDefaultTaskEVMethod, VT_I4)
DISP_PROPERTY_EX_ID(MP12, "bProjectExternallyEdited", 59, odl_GetProjectExternallyEdited, odl_SetProjectExternallyEdited, VT_BOOL)
DISP_PROPERTY_EX_ID(MP12, "dtExtendedCreationDate", 60, odl_GetExtendedCreationDate, odl_SetExtendedCreationDate, VT_DATE)
DISP_PROPERTY_EX_ID(MP12, "bActualsInSync", 61, odl_GetActualsInSync, odl_SetActualsInSync, VT_BOOL)
DISP_PROPERTY_EX_ID(MP12, "bRemoveFileProperties", 62, odl_GetRemoveFileProperties, odl_SetRemoveFileProperties, VT_BOOL)
DISP_PROPERTY_EX_ID(MP12, "bAdminProject", 63, odl_GetAdminProject, odl_SetAdminProject, VT_BOOL)
DISP_PROPERTY_EX_ID(MP12, "oOutlineCodes", 64, odl_GetOutlineCodes, SetNotSupported, VT_DISPATCH)
DISP_PROPERTY_EX_ID(MP12, "oWBSMasks", 65, odl_GetWBSMasks, SetNotSupported, VT_DISPATCH)
DISP_PROPERTY_EX_ID(MP12, "oExtendedAttributes", 66, odl_GetExtendedAttributes, SetNotSupported, VT_DISPATCH)
DISP_PROPERTY_EX_ID(MP12, "oCalendars", 67, odl_GetCalendars, SetNotSupported, VT_DISPATCH)
DISP_PROPERTY_EX_ID(MP12, "oTasks", 68, odl_GetTasks, SetNotSupported, VT_DISPATCH)
DISP_PROPERTY_EX_ID(MP12, "oResources", 69, odl_GetResources, SetNotSupported, VT_DISPATCH)
DISP_PROPERTY_EX_ID(MP12, "oAssignments", 70, odl_GetAssignments, SetNotSupported, VT_DISPATCH)
DISP_FUNCTION_ID(MP12, "WriteXML", 71, odl_WriteXML, VT_EMPTY, VTS_BSTR)
DISP_FUNCTION_ID(MP12, "ReadXML", 72, odl_ReadXML, VT_EMPTY, VTS_BSTR)
DISP_FUNCTION_ID(MP12, "GetXML", 73, odl_GetXML, VT_BSTR, VTS_NONE)
DISP_FUNCTION_ID(MP12, "SetXML", 74, odl_SetXML, VT_EMPTY, VTS_BSTR)
DISP_FUNCTION_ID(MP12, "IsNull", 75, IsNull, VT_BOOL, VTS_NONE)
DISP_FUNCTION_ID(MP12, "Initialize", 76, Initialize, VT_EMPTY, VTS_NONE)
END_DISPATCH_MAP()

BEGIN_INTERFACE_MAP(MP12, CCmdTarget)
	INTERFACE_PART(MP12, IID_IMP12, Dispatch)
END_INTERFACE_MAP()

BEGIN_MESSAGE_MAP(MP12, CCmdTarget)
END_MESSAGE_MAP()


void MP12::Initialize(void)
{
	InitVars();
}

MP12::MP12()
{
	EnableAutomation();
	AfxOleLockApp();
	InitVars();
}

void MP12::InitVars(void)
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
	mp_oOutlineCodes = new OutlineCodes();
	mp_oWBSMasks = new WBSMasks();
	mp_oExtendedAttributes = new ExtendedAttributes();
	mp_oCalendars = new Calendars();
	mp_oTasks = new Tasks();
	mp_oResources = new Resources();
	mp_oAssignments = new Assignments();
}

MP12::~MP12()
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

void MP12::OnFinalRelease()
{
	CCmdTarget::OnFinalRelease();
}

LONG MP12::odl_GetSaveVersion(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetSaveVersion();
}

LONG MP12::GetSaveVersion(void)
{
	return (LONG) mp_lSaveVersion;
}

void MP12::odl_SetSaveVersion(LONG newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetSaveVersion((LONG)newVal);
}

void MP12::SetSaveVersion(LONG newVal)
{
	mp_lSaveVersion = newVal;
}

BSTR MP12::odl_GetUID(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetUID().AllocSysString();
}

CString MP12::GetUID(void)
{
	return (CString) mp_sUID;
}

void MP12::odl_SetUID(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetUID(newVal);
}

void MP12::SetUID(CString newVal)
{
	if (newVal.GetLength() > 16)
	{
		newVal = newVal.Mid(0, 16);
	}
	mp_sUID = newVal;
}

BSTR MP12::odl_GetName(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetName().AllocSysString();
}

CString MP12::GetName(void)
{
	return (CString) mp_sName;
}

void MP12::odl_SetName(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetName(newVal);
}

void MP12::SetName(CString newVal)
{
	if (newVal.GetLength() > 255)
	{
		newVal = newVal.Mid(0, 255);
	}
	mp_sName = newVal;
}

BSTR MP12::odl_GetTitle(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetTitle().AllocSysString();
}

CString MP12::GetTitle(void)
{
	return (CString) mp_sTitle;
}

void MP12::odl_SetTitle(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetTitle(newVal);
}

void MP12::SetTitle(CString newVal)
{
	if (newVal.GetLength() > 512)
	{
		newVal = newVal.Mid(0, 512);
	}
	mp_sTitle = newVal;
}

BSTR MP12::odl_GetSubject(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetSubject().AllocSysString();
}

CString MP12::GetSubject(void)
{
	return (CString) mp_sSubject;
}

void MP12::odl_SetSubject(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetSubject(newVal);
}

void MP12::SetSubject(CString newVal)
{
	if (newVal.GetLength() > 512)
	{
		newVal = newVal.Mid(0, 512);
	}
	mp_sSubject = newVal;
}

BSTR MP12::odl_GetCategory(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetCategory().AllocSysString();
}

CString MP12::GetCategory(void)
{
	return (CString) mp_sCategory;
}

void MP12::odl_SetCategory(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetCategory(newVal);
}

void MP12::SetCategory(CString newVal)
{
	if (newVal.GetLength() > 512)
	{
		newVal = newVal.Mid(0, 512);
	}
	mp_sCategory = newVal;
}

BSTR MP12::odl_GetCompany(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetCompany().AllocSysString();
}

CString MP12::GetCompany(void)
{
	return (CString) mp_sCompany;
}

void MP12::odl_SetCompany(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetCompany(newVal);
}

void MP12::SetCompany(CString newVal)
{
	if (newVal.GetLength() > 512)
	{
		newVal = newVal.Mid(0, 512);
	}
	mp_sCompany = newVal;
}

BSTR MP12::odl_GetManager(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetManager().AllocSysString();
}

CString MP12::GetManager(void)
{
	return (CString) mp_sManager;
}

void MP12::odl_SetManager(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetManager(newVal);
}

void MP12::SetManager(CString newVal)
{
	if (newVal.GetLength() > 512)
	{
		newVal = newVal.Mid(0, 512);
	}
	mp_sManager = newVal;
}

BSTR MP12::odl_GetAuthor(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetAuthor().AllocSysString();
}

CString MP12::GetAuthor(void)
{
	return (CString) mp_sAuthor;
}

void MP12::odl_SetAuthor(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetAuthor(newVal);
}

void MP12::SetAuthor(CString newVal)
{
	if (newVal.GetLength() > 512)
	{
		newVal = newVal.Mid(0, 512);
	}
	mp_sAuthor = newVal;
}

DATE MP12::odl_GetCreationDate(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetCreationDate();
}

COleDateTime MP12::GetCreationDate(void)
{
	return (COleDateTime) mp_dtCreationDate;
}

void MP12::odl_SetCreationDate(DATE newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetCreationDate((COleDateTime)newVal);
}

void MP12::SetCreationDate(COleDateTime newVal)
{
	mp_dtCreationDate = newVal;
}

LONG MP12::odl_GetRevision(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetRevision();
}

LONG MP12::GetRevision(void)
{
	return (LONG) mp_lRevision;
}

void MP12::odl_SetRevision(LONG newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetRevision((LONG)newVal);
}

void MP12::SetRevision(LONG newVal)
{
	mp_lRevision = newVal;
}

DATE MP12::odl_GetLastSaved(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetLastSaved();
}

COleDateTime MP12::GetLastSaved(void)
{
	return (COleDateTime) mp_dtLastSaved;
}

void MP12::odl_SetLastSaved(DATE newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetLastSaved((COleDateTime)newVal);
}

void MP12::SetLastSaved(COleDateTime newVal)
{
	mp_dtLastSaved = newVal;
}

BOOL MP12::odl_GetScheduleFromStart(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetScheduleFromStart();
}

BOOL MP12::GetScheduleFromStart(void)
{
	return (BOOL) mp_bScheduleFromStart;
}

void MP12::odl_SetScheduleFromStart(BOOL newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetScheduleFromStart((BOOL)newVal);
}

void MP12::SetScheduleFromStart(BOOL newVal)
{
	mp_bScheduleFromStart = newVal;
}

DATE MP12::odl_GetStartDate(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetStartDate();
}

COleDateTime MP12::GetStartDate(void)
{
	return (COleDateTime) mp_dtStartDate;
}

void MP12::odl_SetStartDate(DATE newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetStartDate((COleDateTime)newVal);
}

void MP12::SetStartDate(COleDateTime newVal)
{
	mp_dtStartDate = newVal;
}

DATE MP12::odl_GetFinishDate(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetFinishDate();
}

COleDateTime MP12::GetFinishDate(void)
{
	return (COleDateTime) mp_dtFinishDate;
}

void MP12::odl_SetFinishDate(DATE newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetFinishDate((COleDateTime)newVal);
}

void MP12::SetFinishDate(COleDateTime newVal)
{
	mp_dtFinishDate = newVal;
}

LONG MP12::odl_GetFYStartDate(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetFYStartDate();
}

E_FYSTARTDATE MP12::GetFYStartDate(void)
{
	return (E_FYSTARTDATE) mp_yFYStartDate;
}

void MP12::odl_SetFYStartDate(LONG newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetFYStartDate((E_FYSTARTDATE)newVal);
}

void MP12::SetFYStartDate(E_FYSTARTDATE newVal)
{
	mp_yFYStartDate = newVal;
}

LONG MP12::odl_GetCriticalSlackLimit(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetCriticalSlackLimit();
}

LONG MP12::GetCriticalSlackLimit(void)
{
	return (LONG) mp_lCriticalSlackLimit;
}

void MP12::odl_SetCriticalSlackLimit(LONG newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetCriticalSlackLimit((LONG)newVal);
}

void MP12::SetCriticalSlackLimit(LONG newVal)
{
	mp_lCriticalSlackLimit = newVal;
}

LONG MP12::odl_GetCurrencyDigits(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetCurrencyDigits();
}

LONG MP12::GetCurrencyDigits(void)
{
	return (LONG) mp_lCurrencyDigits;
}

void MP12::odl_SetCurrencyDigits(LONG newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetCurrencyDigits((LONG)newVal);
}

void MP12::SetCurrencyDigits(LONG newVal)
{
	mp_lCurrencyDigits = newVal;
}

BSTR MP12::odl_GetCurrencySymbol(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetCurrencySymbol().AllocSysString();
}

CString MP12::GetCurrencySymbol(void)
{
	return (CString) mp_sCurrencySymbol;
}

void MP12::odl_SetCurrencySymbol(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetCurrencySymbol(newVal);
}

void MP12::SetCurrencySymbol(CString newVal)
{
	if (newVal.GetLength() > 20)
	{
		newVal = newVal.Mid(0, 20);
	}
	mp_sCurrencySymbol = newVal;
}

BSTR MP12::odl_GetCurrencyCode(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetCurrencyCode().AllocSysString();
}

CString MP12::GetCurrencyCode(void)
{
	return (CString) mp_sCurrencyCode;
}

void MP12::odl_SetCurrencyCode(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetCurrencyCode(newVal);
}

void MP12::SetCurrencyCode(CString newVal)
{
	if (newVal.GetLength() > 3)
	{
		newVal = newVal.Mid(0, 3);
	}
	mp_sCurrencyCode = newVal;
}

LONG MP12::odl_GetCurrencySymbolPosition(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetCurrencySymbolPosition();
}

E_CURRENCYSYMBOLPOSITION MP12::GetCurrencySymbolPosition(void)
{
	return (E_CURRENCYSYMBOLPOSITION) mp_yCurrencySymbolPosition;
}

void MP12::odl_SetCurrencySymbolPosition(LONG newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetCurrencySymbolPosition((E_CURRENCYSYMBOLPOSITION)newVal);
}

void MP12::SetCurrencySymbolPosition(E_CURRENCYSYMBOLPOSITION newVal)
{
	mp_yCurrencySymbolPosition = newVal;
}

LONG MP12::odl_GetCalendarUID(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetCalendarUID();
}

LONG MP12::GetCalendarUID(void)
{
	return (LONG) mp_lCalendarUID;
}

void MP12::odl_SetCalendarUID(LONG newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetCalendarUID((LONG)newVal);
}

void MP12::SetCalendarUID(LONG newVal)
{
	mp_lCalendarUID = newVal;
}

IDispatch* MP12::odl_GetDefaultStartTime(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetDefaultStartTime()->GetIDispatch(TRUE);
}

Time* MP12::GetDefaultStartTime(void)
{
	return mp_oDefaultStartTime;
}

IDispatch* MP12::odl_GetDefaultFinishTime(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetDefaultFinishTime()->GetIDispatch(TRUE);
}

Time* MP12::GetDefaultFinishTime(void)
{
	return mp_oDefaultFinishTime;
}

LONG MP12::odl_GetMinutesPerDay(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetMinutesPerDay();
}

LONG MP12::GetMinutesPerDay(void)
{
	return (LONG) mp_lMinutesPerDay;
}

void MP12::odl_SetMinutesPerDay(LONG newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetMinutesPerDay((LONG)newVal);
}

void MP12::SetMinutesPerDay(LONG newVal)
{
	mp_lMinutesPerDay = newVal;
}

LONG MP12::odl_GetMinutesPerWeek(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetMinutesPerWeek();
}

LONG MP12::GetMinutesPerWeek(void)
{
	return (LONG) mp_lMinutesPerWeek;
}

void MP12::odl_SetMinutesPerWeek(LONG newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetMinutesPerWeek((LONG)newVal);
}

void MP12::SetMinutesPerWeek(LONG newVal)
{
	mp_lMinutesPerWeek = newVal;
}

LONG MP12::odl_GetDaysPerMonth(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetDaysPerMonth();
}

LONG MP12::GetDaysPerMonth(void)
{
	return (LONG) mp_lDaysPerMonth;
}

void MP12::odl_SetDaysPerMonth(LONG newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetDaysPerMonth((LONG)newVal);
}

void MP12::SetDaysPerMonth(LONG newVal)
{
	mp_lDaysPerMonth = newVal;
}

LONG MP12::odl_GetDefaultTaskType(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetDefaultTaskType();
}

E_DEFAULTTASKTYPE MP12::GetDefaultTaskType(void)
{
	return (E_DEFAULTTASKTYPE) mp_yDefaultTaskType;
}

void MP12::odl_SetDefaultTaskType(LONG newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetDefaultTaskType((E_DEFAULTTASKTYPE)newVal);
}

void MP12::SetDefaultTaskType(E_DEFAULTTASKTYPE newVal)
{
	mp_yDefaultTaskType = newVal;
}

LONG MP12::odl_GetDefaultFixedCostAccrual(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetDefaultFixedCostAccrual();
}

E_DEFAULTFIXEDCOSTACCRUAL MP12::GetDefaultFixedCostAccrual(void)
{
	return (E_DEFAULTFIXEDCOSTACCRUAL) mp_yDefaultFixedCostAccrual;
}

void MP12::odl_SetDefaultFixedCostAccrual(LONG newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetDefaultFixedCostAccrual((E_DEFAULTFIXEDCOSTACCRUAL)newVal);
}

void MP12::SetDefaultFixedCostAccrual(E_DEFAULTFIXEDCOSTACCRUAL newVal)
{
	mp_yDefaultFixedCostAccrual = newVal;
}

FLOAT MP12::odl_GetDefaultStandardRate(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetDefaultStandardRate();
}

FLOAT MP12::GetDefaultStandardRate(void)
{
	return (FLOAT) mp_fDefaultStandardRate;
}

void MP12::odl_SetDefaultStandardRate(FLOAT newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetDefaultStandardRate((FLOAT)newVal);
}

void MP12::SetDefaultStandardRate(FLOAT newVal)
{
	mp_fDefaultStandardRate = newVal;
}

FLOAT MP12::odl_GetDefaultOvertimeRate(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetDefaultOvertimeRate();
}

FLOAT MP12::GetDefaultOvertimeRate(void)
{
	return (FLOAT) mp_fDefaultOvertimeRate;
}

void MP12::odl_SetDefaultOvertimeRate(FLOAT newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetDefaultOvertimeRate((FLOAT)newVal);
}

void MP12::SetDefaultOvertimeRate(FLOAT newVal)
{
	mp_fDefaultOvertimeRate = newVal;
}

LONG MP12::odl_GetDurationFormat(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetDurationFormat();
}

E_DURATIONFORMAT MP12::GetDurationFormat(void)
{
	return (E_DURATIONFORMAT) mp_yDurationFormat;
}

void MP12::odl_SetDurationFormat(LONG newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetDurationFormat((E_DURATIONFORMAT)newVal);
}

void MP12::SetDurationFormat(E_DURATIONFORMAT newVal)
{
	mp_yDurationFormat = newVal;
}

LONG MP12::odl_GetWorkFormat(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetWorkFormat();
}

E_WORKFORMAT MP12::GetWorkFormat(void)
{
	return (E_WORKFORMAT) mp_yWorkFormat;
}

void MP12::odl_SetWorkFormat(LONG newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetWorkFormat((E_WORKFORMAT)newVal);
}

void MP12::SetWorkFormat(E_WORKFORMAT newVal)
{
	mp_yWorkFormat = newVal;
}

BOOL MP12::odl_GetEditableActualCosts(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetEditableActualCosts();
}

BOOL MP12::GetEditableActualCosts(void)
{
	return (BOOL) mp_bEditableActualCosts;
}

void MP12::odl_SetEditableActualCosts(BOOL newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetEditableActualCosts((BOOL)newVal);
}

void MP12::SetEditableActualCosts(BOOL newVal)
{
	mp_bEditableActualCosts = newVal;
}

BOOL MP12::odl_GetHonorConstraints(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetHonorConstraints();
}

BOOL MP12::GetHonorConstraints(void)
{
	return (BOOL) mp_bHonorConstraints;
}

void MP12::odl_SetHonorConstraints(BOOL newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetHonorConstraints((BOOL)newVal);
}

void MP12::SetHonorConstraints(BOOL newVal)
{
	mp_bHonorConstraints = newVal;
}

LONG MP12::odl_GetEarnedValueMethod(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetEarnedValueMethod();
}

E_EARNEDVALUEMETHOD MP12::GetEarnedValueMethod(void)
{
	return (E_EARNEDVALUEMETHOD) mp_yEarnedValueMethod;
}

void MP12::odl_SetEarnedValueMethod(LONG newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetEarnedValueMethod((E_EARNEDVALUEMETHOD)newVal);
}

void MP12::SetEarnedValueMethod(E_EARNEDVALUEMETHOD newVal)
{
	mp_yEarnedValueMethod = newVal;
}

BOOL MP12::odl_GetInsertedProjectsLikeSummary(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetInsertedProjectsLikeSummary();
}

BOOL MP12::GetInsertedProjectsLikeSummary(void)
{
	return (BOOL) mp_bInsertedProjectsLikeSummary;
}

void MP12::odl_SetInsertedProjectsLikeSummary(BOOL newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetInsertedProjectsLikeSummary((BOOL)newVal);
}

void MP12::SetInsertedProjectsLikeSummary(BOOL newVal)
{
	mp_bInsertedProjectsLikeSummary = newVal;
}

BOOL MP12::odl_GetMultipleCriticalPaths(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetMultipleCriticalPaths();
}

BOOL MP12::GetMultipleCriticalPaths(void)
{
	return (BOOL) mp_bMultipleCriticalPaths;
}

void MP12::odl_SetMultipleCriticalPaths(BOOL newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetMultipleCriticalPaths((BOOL)newVal);
}

void MP12::SetMultipleCriticalPaths(BOOL newVal)
{
	mp_bMultipleCriticalPaths = newVal;
}

BOOL MP12::odl_GetNewTasksEffortDriven(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetNewTasksEffortDriven();
}

BOOL MP12::GetNewTasksEffortDriven(void)
{
	return (BOOL) mp_bNewTasksEffortDriven;
}

void MP12::odl_SetNewTasksEffortDriven(BOOL newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetNewTasksEffortDriven((BOOL)newVal);
}

void MP12::SetNewTasksEffortDriven(BOOL newVal)
{
	mp_bNewTasksEffortDriven = newVal;
}

BOOL MP12::odl_GetNewTasksEstimated(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetNewTasksEstimated();
}

BOOL MP12::GetNewTasksEstimated(void)
{
	return (BOOL) mp_bNewTasksEstimated;
}

void MP12::odl_SetNewTasksEstimated(BOOL newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetNewTasksEstimated((BOOL)newVal);
}

void MP12::SetNewTasksEstimated(BOOL newVal)
{
	mp_bNewTasksEstimated = newVal;
}

BOOL MP12::odl_GetSplitsInProgressTasks(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetSplitsInProgressTasks();
}

BOOL MP12::GetSplitsInProgressTasks(void)
{
	return (BOOL) mp_bSplitsInProgressTasks;
}

void MP12::odl_SetSplitsInProgressTasks(BOOL newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetSplitsInProgressTasks((BOOL)newVal);
}

void MP12::SetSplitsInProgressTasks(BOOL newVal)
{
	mp_bSplitsInProgressTasks = newVal;
}

BOOL MP12::odl_GetSpreadActualCost(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetSpreadActualCost();
}

BOOL MP12::GetSpreadActualCost(void)
{
	return (BOOL) mp_bSpreadActualCost;
}

void MP12::odl_SetSpreadActualCost(BOOL newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetSpreadActualCost((BOOL)newVal);
}

void MP12::SetSpreadActualCost(BOOL newVal)
{
	mp_bSpreadActualCost = newVal;
}

BOOL MP12::odl_GetSpreadPercentComplete(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetSpreadPercentComplete();
}

BOOL MP12::GetSpreadPercentComplete(void)
{
	return (BOOL) mp_bSpreadPercentComplete;
}

void MP12::odl_SetSpreadPercentComplete(BOOL newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetSpreadPercentComplete((BOOL)newVal);
}

void MP12::SetSpreadPercentComplete(BOOL newVal)
{
	mp_bSpreadPercentComplete = newVal;
}

BOOL MP12::odl_GetTaskUpdatesResource(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetTaskUpdatesResource();
}

BOOL MP12::GetTaskUpdatesResource(void)
{
	return (BOOL) mp_bTaskUpdatesResource;
}

void MP12::odl_SetTaskUpdatesResource(BOOL newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetTaskUpdatesResource((BOOL)newVal);
}

void MP12::SetTaskUpdatesResource(BOOL newVal)
{
	mp_bTaskUpdatesResource = newVal;
}

BOOL MP12::odl_GetFiscalYearStart(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetFiscalYearStart();
}

BOOL MP12::GetFiscalYearStart(void)
{
	return (BOOL) mp_bFiscalYearStart;
}

void MP12::odl_SetFiscalYearStart(BOOL newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetFiscalYearStart((BOOL)newVal);
}

void MP12::SetFiscalYearStart(BOOL newVal)
{
	mp_bFiscalYearStart = newVal;
}

LONG MP12::odl_GetWeekStartDay(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetWeekStartDay();
}

E_WEEKSTARTDAY MP12::GetWeekStartDay(void)
{
	return (E_WEEKSTARTDAY) mp_yWeekStartDay;
}

void MP12::odl_SetWeekStartDay(LONG newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetWeekStartDay((E_WEEKSTARTDAY)newVal);
}

void MP12::SetWeekStartDay(E_WEEKSTARTDAY newVal)
{
	mp_yWeekStartDay = newVal;
}

BOOL MP12::odl_GetMoveCompletedEndsBack(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetMoveCompletedEndsBack();
}

BOOL MP12::GetMoveCompletedEndsBack(void)
{
	return (BOOL) mp_bMoveCompletedEndsBack;
}

void MP12::odl_SetMoveCompletedEndsBack(BOOL newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetMoveCompletedEndsBack((BOOL)newVal);
}

void MP12::SetMoveCompletedEndsBack(BOOL newVal)
{
	mp_bMoveCompletedEndsBack = newVal;
}

BOOL MP12::odl_GetMoveRemainingStartsBack(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetMoveRemainingStartsBack();
}

BOOL MP12::GetMoveRemainingStartsBack(void)
{
	return (BOOL) mp_bMoveRemainingStartsBack;
}

void MP12::odl_SetMoveRemainingStartsBack(BOOL newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetMoveRemainingStartsBack((BOOL)newVal);
}

void MP12::SetMoveRemainingStartsBack(BOOL newVal)
{
	mp_bMoveRemainingStartsBack = newVal;
}

BOOL MP12::odl_GetMoveRemainingStartsForward(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetMoveRemainingStartsForward();
}

BOOL MP12::GetMoveRemainingStartsForward(void)
{
	return (BOOL) mp_bMoveRemainingStartsForward;
}

void MP12::odl_SetMoveRemainingStartsForward(BOOL newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetMoveRemainingStartsForward((BOOL)newVal);
}

void MP12::SetMoveRemainingStartsForward(BOOL newVal)
{
	mp_bMoveRemainingStartsForward = newVal;
}

BOOL MP12::odl_GetMoveCompletedEndsForward(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetMoveCompletedEndsForward();
}

BOOL MP12::GetMoveCompletedEndsForward(void)
{
	return (BOOL) mp_bMoveCompletedEndsForward;
}

void MP12::odl_SetMoveCompletedEndsForward(BOOL newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetMoveCompletedEndsForward((BOOL)newVal);
}

void MP12::SetMoveCompletedEndsForward(BOOL newVal)
{
	mp_bMoveCompletedEndsForward = newVal;
}

LONG MP12::odl_GetBaselineForEarnedValue(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetBaselineForEarnedValue();
}

E_BASELINEFOREARNEDVALUE MP12::GetBaselineForEarnedValue(void)
{
	return (E_BASELINEFOREARNEDVALUE) mp_yBaselineForEarnedValue;
}

void MP12::odl_SetBaselineForEarnedValue(LONG newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetBaselineForEarnedValue((E_BASELINEFOREARNEDVALUE)newVal);
}

void MP12::SetBaselineForEarnedValue(E_BASELINEFOREARNEDVALUE newVal)
{
	mp_yBaselineForEarnedValue = newVal;
}

BOOL MP12::odl_GetAutoAddNewResourcesAndTasks(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetAutoAddNewResourcesAndTasks();
}

BOOL MP12::GetAutoAddNewResourcesAndTasks(void)
{
	return (BOOL) mp_bAutoAddNewResourcesAndTasks;
}

void MP12::odl_SetAutoAddNewResourcesAndTasks(BOOL newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetAutoAddNewResourcesAndTasks((BOOL)newVal);
}

void MP12::SetAutoAddNewResourcesAndTasks(BOOL newVal)
{
	mp_bAutoAddNewResourcesAndTasks = newVal;
}

DATE MP12::odl_GetStatusDate(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetStatusDate();
}

COleDateTime MP12::GetStatusDate(void)
{
	return (COleDateTime) mp_dtStatusDate;
}

void MP12::odl_SetStatusDate(DATE newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetStatusDate((COleDateTime)newVal);
}

void MP12::SetStatusDate(COleDateTime newVal)
{
	mp_dtStatusDate = newVal;
}

DATE MP12::odl_GetCurrentDate(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetCurrentDate();
}

COleDateTime MP12::GetCurrentDate(void)
{
	return (COleDateTime) mp_dtCurrentDate;
}

void MP12::odl_SetCurrentDate(DATE newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetCurrentDate((COleDateTime)newVal);
}

void MP12::SetCurrentDate(COleDateTime newVal)
{
	mp_dtCurrentDate = newVal;
}

BOOL MP12::odl_GetMicrosoftProjectServerURL(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetMicrosoftProjectServerURL();
}

BOOL MP12::GetMicrosoftProjectServerURL(void)
{
	return (BOOL) mp_bMicrosoftProjectServerURL;
}

void MP12::odl_SetMicrosoftProjectServerURL(BOOL newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetMicrosoftProjectServerURL((BOOL)newVal);
}

void MP12::SetMicrosoftProjectServerURL(BOOL newVal)
{
	mp_bMicrosoftProjectServerURL = newVal;
}

BOOL MP12::odl_GetAutolink(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetAutolink();
}

BOOL MP12::GetAutolink(void)
{
	return (BOOL) mp_bAutolink;
}

void MP12::odl_SetAutolink(BOOL newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetAutolink((BOOL)newVal);
}

void MP12::SetAutolink(BOOL newVal)
{
	mp_bAutolink = newVal;
}

LONG MP12::odl_GetNewTaskStartDate(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetNewTaskStartDate();
}

E_NEWTASKSTARTDATE MP12::GetNewTaskStartDate(void)
{
	return (E_NEWTASKSTARTDATE) mp_yNewTaskStartDate;
}

void MP12::odl_SetNewTaskStartDate(LONG newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetNewTaskStartDate((E_NEWTASKSTARTDATE)newVal);
}

void MP12::SetNewTaskStartDate(E_NEWTASKSTARTDATE newVal)
{
	mp_yNewTaskStartDate = newVal;
}

LONG MP12::odl_GetDefaultTaskEVMethod(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetDefaultTaskEVMethod();
}

E_DEFAULTTASKEVMETHOD MP12::GetDefaultTaskEVMethod(void)
{
	return (E_DEFAULTTASKEVMETHOD) mp_yDefaultTaskEVMethod;
}

void MP12::odl_SetDefaultTaskEVMethod(LONG newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetDefaultTaskEVMethod((E_DEFAULTTASKEVMETHOD)newVal);
}

void MP12::SetDefaultTaskEVMethod(E_DEFAULTTASKEVMETHOD newVal)
{
	mp_yDefaultTaskEVMethod = newVal;
}

BOOL MP12::odl_GetProjectExternallyEdited(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetProjectExternallyEdited();
}

BOOL MP12::GetProjectExternallyEdited(void)
{
	return (BOOL) mp_bProjectExternallyEdited;
}

void MP12::odl_SetProjectExternallyEdited(BOOL newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetProjectExternallyEdited((BOOL)newVal);
}

void MP12::SetProjectExternallyEdited(BOOL newVal)
{
	mp_bProjectExternallyEdited = newVal;
}

DATE MP12::odl_GetExtendedCreationDate(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetExtendedCreationDate();
}

COleDateTime MP12::GetExtendedCreationDate(void)
{
	return (COleDateTime) mp_dtExtendedCreationDate;
}

void MP12::odl_SetExtendedCreationDate(DATE newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetExtendedCreationDate((COleDateTime)newVal);
}

void MP12::SetExtendedCreationDate(COleDateTime newVal)
{
	mp_dtExtendedCreationDate = newVal;
}

BOOL MP12::odl_GetActualsInSync(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetActualsInSync();
}

BOOL MP12::GetActualsInSync(void)
{
	return (BOOL) mp_bActualsInSync;
}

void MP12::odl_SetActualsInSync(BOOL newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetActualsInSync((BOOL)newVal);
}

void MP12::SetActualsInSync(BOOL newVal)
{
	mp_bActualsInSync = newVal;
}

BOOL MP12::odl_GetRemoveFileProperties(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetRemoveFileProperties();
}

BOOL MP12::GetRemoveFileProperties(void)
{
	return (BOOL) mp_bRemoveFileProperties;
}

void MP12::odl_SetRemoveFileProperties(BOOL newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetRemoveFileProperties((BOOL)newVal);
}

void MP12::SetRemoveFileProperties(BOOL newVal)
{
	mp_bRemoveFileProperties = newVal;
}

BOOL MP12::odl_GetAdminProject(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetAdminProject();
}

BOOL MP12::GetAdminProject(void)
{
	return (BOOL) mp_bAdminProject;
}

void MP12::odl_SetAdminProject(BOOL newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetAdminProject((BOOL)newVal);
}

void MP12::SetAdminProject(BOOL newVal)
{
	mp_bAdminProject = newVal;
}

IDispatch* MP12::odl_GetOutlineCodes(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetOutlineCodes()->GetIDispatch(TRUE);
}

OutlineCodes* MP12::GetOutlineCodes(void)
{
	return mp_oOutlineCodes;
}

IDispatch* MP12::odl_GetWBSMasks(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetWBSMasks()->GetIDispatch(TRUE);
}

WBSMasks* MP12::GetWBSMasks(void)
{
	return mp_oWBSMasks;
}

IDispatch* MP12::odl_GetExtendedAttributes(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetExtendedAttributes()->GetIDispatch(TRUE);
}

ExtendedAttributes* MP12::GetExtendedAttributes(void)
{
	return mp_oExtendedAttributes;
}

IDispatch* MP12::odl_GetCalendars(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetCalendars()->GetIDispatch(TRUE);
}

Calendars* MP12::GetCalendars(void)
{
	return mp_oCalendars;
}

IDispatch* MP12::odl_GetTasks(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetTasks()->GetIDispatch(TRUE);
}

Tasks* MP12::GetTasks(void)
{
	return mp_oTasks;
}

IDispatch* MP12::odl_GetResources(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetResources()->GetIDispatch(TRUE);
}

Resources* MP12::GetResources(void)
{
	return mp_oResources;
}

IDispatch* MP12::odl_GetAssignments(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetAssignments()->GetIDispatch(TRUE);
}

Assignments* MP12::GetAssignments(void)
{
	return mp_oAssignments;
}

void MP12::WriteXML(CString url)
{
	clsXML oXML("Project");
	mp_WriteXML(oXML);
	oXML.WriteXML(url);
}

void MP12::ReadXML(CString url)
{
	clsXML oXML("Project");
	oXML.ReadXML(url);
	oXML.RemoveNamespace();
	mp_ReadXML(oXML);
}

void MP12::SetXML(CString sXML)
{
	clsXML oXML("Project");
	oXML.SetXML(sXML);
	oXML.RemoveNamespace();
	mp_ReadXML(oXML);
}

CString MP12::GetXML()
{
	clsXML oXML("Project");
	mp_WriteXML(oXML);
	return oXML.GetXML();
}

void MP12::odl_WriteXML(LPCTSTR url)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	WriteXML(url);
}

void MP12::odl_ReadXML(LPCTSTR url)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	ReadXML(url);
}

void MP12::odl_SetXML(LPCTSTR sXML)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetXML(sXML);
}

BSTR MP12::odl_GetXML(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetXML().AllocSysString();
}

BOOL MP12::IsNull(void)
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

void MP12::mp_WriteXML(clsXML &oXML)
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
	oXML.WriteObject(mp_oOutlineCodes->GetXML());
	oXML.WriteObject(mp_oWBSMasks->GetXML());
	oXML.WriteObject(mp_oExtendedAttributes->GetXML());
	oXML.WriteObject(mp_oCalendars->GetXML());
	oXML.WriteObject(mp_oTasks->GetXML());
	oXML.WriteObject(mp_oResources->GetXML());
	oXML.WriteObject(mp_oAssignments->GetXML());
	oXML.AddNamespace();
}

void MP12::mp_ReadXML(clsXML &oXML)
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
	mp_oOutlineCodes->SetXML(oXML.ReadObject("OutlineCodes"));
	mp_oWBSMasks->SetXML(oXML.ReadObject("WBSMasks"));
	mp_oExtendedAttributes->SetXML(oXML.ReadObject("ExtendedAttributes"));
	mp_oCalendars->SetXML(oXML.ReadObject("Calendars"));
	mp_oTasks->SetXML(oXML.ReadObject("Tasks"));
	mp_oResources->SetXML(oXML.ReadObject("Resources"));
	mp_oAssignments->SetXML(oXML.ReadObject("Assignments"));
}