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
#pragma once

class clsXML;


class MP12 : public CCmdTarget
{
	DECLARE_DYNCREATE(MP12)

public:

	MP12();
	virtual ~MP12();
	virtual void OnFinalRelease();

	LONG mp_lSaveVersion;
	CString mp_sUID;
	CString mp_sName;
	CString mp_sTitle;
	CString mp_sSubject;
	CString mp_sCategory;
	CString mp_sCompany;
	CString mp_sManager;
	CString mp_sAuthor;
	COleDateTime mp_dtCreationDate;
	LONG mp_lRevision;
	COleDateTime mp_dtLastSaved;
	BOOL mp_bScheduleFromStart;
	COleDateTime mp_dtStartDate;
	COleDateTime mp_dtFinishDate;
	LONG mp_yFYStartDate;
	LONG mp_lCriticalSlackLimit;
	LONG mp_lCurrencyDigits;
	CString mp_sCurrencySymbol;
	CString mp_sCurrencyCode;
	LONG mp_yCurrencySymbolPosition;
	LONG mp_lCalendarUID;
	Time* mp_oDefaultStartTime;
	Time* mp_oDefaultFinishTime;
	LONG mp_lMinutesPerDay;
	LONG mp_lMinutesPerWeek;
	LONG mp_lDaysPerMonth;
	LONG mp_yDefaultTaskType;
	LONG mp_yDefaultFixedCostAccrual;
	FLOAT mp_fDefaultStandardRate;
	FLOAT mp_fDefaultOvertimeRate;
	LONG mp_yDurationFormat;
	LONG mp_yWorkFormat;
	BOOL mp_bEditableActualCosts;
	BOOL mp_bHonorConstraints;
	LONG mp_yEarnedValueMethod;
	BOOL mp_bInsertedProjectsLikeSummary;
	BOOL mp_bMultipleCriticalPaths;
	BOOL mp_bNewTasksEffortDriven;
	BOOL mp_bNewTasksEstimated;
	BOOL mp_bSplitsInProgressTasks;
	BOOL mp_bSpreadActualCost;
	BOOL mp_bSpreadPercentComplete;
	BOOL mp_bTaskUpdatesResource;
	BOOL mp_bFiscalYearStart;
	LONG mp_yWeekStartDay;
	BOOL mp_bMoveCompletedEndsBack;
	BOOL mp_bMoveRemainingStartsBack;
	BOOL mp_bMoveRemainingStartsForward;
	BOOL mp_bMoveCompletedEndsForward;
	LONG mp_yBaselineForEarnedValue;
	BOOL mp_bAutoAddNewResourcesAndTasks;
	COleDateTime mp_dtStatusDate;
	COleDateTime mp_dtCurrentDate;
	BOOL mp_bMicrosoftProjectServerURL;
	BOOL mp_bAutolink;
	LONG mp_yNewTaskStartDate;
	LONG mp_yDefaultTaskEVMethod;
	BOOL mp_bProjectExternallyEdited;
	COleDateTime mp_dtExtendedCreationDate;
	BOOL mp_bActualsInSync;
	BOOL mp_bRemoveFileProperties;
	BOOL mp_bAdminProject;
	OutlineCodes* mp_oOutlineCodes;
	WBSMasks* mp_oWBSMasks;
	ExtendedAttributes* mp_oExtendedAttributes;
	Calendars* mp_oCalendars;
	Tasks* mp_oTasks;
	Resources* mp_oResources;
	Assignments* mp_oAssignments;

	LONG odl_GetSaveVersion(void);
	LONG GetSaveVersion(void);
	void odl_SetSaveVersion(LONG newVal);
	void SetSaveVersion(LONG newVal);

	BSTR odl_GetUID(void);
	CString GetUID(void);
	void odl_SetUID(LPCTSTR newVal);
	void SetUID(CString newVal);

	BSTR odl_GetName(void);
	CString GetName(void);
	void odl_SetName(LPCTSTR newVal);
	void SetName(CString newVal);

	BSTR odl_GetTitle(void);
	CString GetTitle(void);
	void odl_SetTitle(LPCTSTR newVal);
	void SetTitle(CString newVal);

	BSTR odl_GetSubject(void);
	CString GetSubject(void);
	void odl_SetSubject(LPCTSTR newVal);
	void SetSubject(CString newVal);

	BSTR odl_GetCategory(void);
	CString GetCategory(void);
	void odl_SetCategory(LPCTSTR newVal);
	void SetCategory(CString newVal);

	BSTR odl_GetCompany(void);
	CString GetCompany(void);
	void odl_SetCompany(LPCTSTR newVal);
	void SetCompany(CString newVal);

	BSTR odl_GetManager(void);
	CString GetManager(void);
	void odl_SetManager(LPCTSTR newVal);
	void SetManager(CString newVal);

	BSTR odl_GetAuthor(void);
	CString GetAuthor(void);
	void odl_SetAuthor(LPCTSTR newVal);
	void SetAuthor(CString newVal);

	DATE odl_GetCreationDate(void);
	COleDateTime GetCreationDate(void);
	void odl_SetCreationDate(DATE newVal);
	void SetCreationDate(COleDateTime newVal);

	LONG odl_GetRevision(void);
	LONG GetRevision(void);
	void odl_SetRevision(LONG newVal);
	void SetRevision(LONG newVal);

	DATE odl_GetLastSaved(void);
	COleDateTime GetLastSaved(void);
	void odl_SetLastSaved(DATE newVal);
	void SetLastSaved(COleDateTime newVal);

	BOOL odl_GetScheduleFromStart(void);
	BOOL GetScheduleFromStart(void);
	void odl_SetScheduleFromStart(BOOL newVal);
	void SetScheduleFromStart(BOOL newVal);

	DATE odl_GetStartDate(void);
	COleDateTime GetStartDate(void);
	void odl_SetStartDate(DATE newVal);
	void SetStartDate(COleDateTime newVal);

	DATE odl_GetFinishDate(void);
	COleDateTime GetFinishDate(void);
	void odl_SetFinishDate(DATE newVal);
	void SetFinishDate(COleDateTime newVal);

	LONG odl_GetFYStartDate(void);
	E_FYSTARTDATE GetFYStartDate(void);
	void odl_SetFYStartDate(LONG newVal);
	void SetFYStartDate(E_FYSTARTDATE newVal);

	LONG odl_GetCriticalSlackLimit(void);
	LONG GetCriticalSlackLimit(void);
	void odl_SetCriticalSlackLimit(LONG newVal);
	void SetCriticalSlackLimit(LONG newVal);

	LONG odl_GetCurrencyDigits(void);
	LONG GetCurrencyDigits(void);
	void odl_SetCurrencyDigits(LONG newVal);
	void SetCurrencyDigits(LONG newVal);

	BSTR odl_GetCurrencySymbol(void);
	CString GetCurrencySymbol(void);
	void odl_SetCurrencySymbol(LPCTSTR newVal);
	void SetCurrencySymbol(CString newVal);

	BSTR odl_GetCurrencyCode(void);
	CString GetCurrencyCode(void);
	void odl_SetCurrencyCode(LPCTSTR newVal);
	void SetCurrencyCode(CString newVal);

	LONG odl_GetCurrencySymbolPosition(void);
	E_CURRENCYSYMBOLPOSITION GetCurrencySymbolPosition(void);
	void odl_SetCurrencySymbolPosition(LONG newVal);
	void SetCurrencySymbolPosition(E_CURRENCYSYMBOLPOSITION newVal);

	LONG odl_GetCalendarUID(void);
	LONG GetCalendarUID(void);
	void odl_SetCalendarUID(LONG newVal);
	void SetCalendarUID(LONG newVal);

	IDispatch* odl_GetDefaultStartTime(void);
	Time* GetDefaultStartTime(void);

	IDispatch* odl_GetDefaultFinishTime(void);
	Time* GetDefaultFinishTime(void);

	LONG odl_GetMinutesPerDay(void);
	LONG GetMinutesPerDay(void);
	void odl_SetMinutesPerDay(LONG newVal);
	void SetMinutesPerDay(LONG newVal);

	LONG odl_GetMinutesPerWeek(void);
	LONG GetMinutesPerWeek(void);
	void odl_SetMinutesPerWeek(LONG newVal);
	void SetMinutesPerWeek(LONG newVal);

	LONG odl_GetDaysPerMonth(void);
	LONG GetDaysPerMonth(void);
	void odl_SetDaysPerMonth(LONG newVal);
	void SetDaysPerMonth(LONG newVal);

	LONG odl_GetDefaultTaskType(void);
	E_DEFAULTTASKTYPE GetDefaultTaskType(void);
	void odl_SetDefaultTaskType(LONG newVal);
	void SetDefaultTaskType(E_DEFAULTTASKTYPE newVal);

	LONG odl_GetDefaultFixedCostAccrual(void);
	E_DEFAULTFIXEDCOSTACCRUAL GetDefaultFixedCostAccrual(void);
	void odl_SetDefaultFixedCostAccrual(LONG newVal);
	void SetDefaultFixedCostAccrual(E_DEFAULTFIXEDCOSTACCRUAL newVal);

	FLOAT odl_GetDefaultStandardRate(void);
	FLOAT GetDefaultStandardRate(void);
	void odl_SetDefaultStandardRate(FLOAT newVal);
	void SetDefaultStandardRate(FLOAT newVal);

	FLOAT odl_GetDefaultOvertimeRate(void);
	FLOAT GetDefaultOvertimeRate(void);
	void odl_SetDefaultOvertimeRate(FLOAT newVal);
	void SetDefaultOvertimeRate(FLOAT newVal);

	LONG odl_GetDurationFormat(void);
	E_DURATIONFORMAT GetDurationFormat(void);
	void odl_SetDurationFormat(LONG newVal);
	void SetDurationFormat(E_DURATIONFORMAT newVal);

	LONG odl_GetWorkFormat(void);
	E_WORKFORMAT GetWorkFormat(void);
	void odl_SetWorkFormat(LONG newVal);
	void SetWorkFormat(E_WORKFORMAT newVal);

	BOOL odl_GetEditableActualCosts(void);
	BOOL GetEditableActualCosts(void);
	void odl_SetEditableActualCosts(BOOL newVal);
	void SetEditableActualCosts(BOOL newVal);

	BOOL odl_GetHonorConstraints(void);
	BOOL GetHonorConstraints(void);
	void odl_SetHonorConstraints(BOOL newVal);
	void SetHonorConstraints(BOOL newVal);

	LONG odl_GetEarnedValueMethod(void);
	E_EARNEDVALUEMETHOD GetEarnedValueMethod(void);
	void odl_SetEarnedValueMethod(LONG newVal);
	void SetEarnedValueMethod(E_EARNEDVALUEMETHOD newVal);

	BOOL odl_GetInsertedProjectsLikeSummary(void);
	BOOL GetInsertedProjectsLikeSummary(void);
	void odl_SetInsertedProjectsLikeSummary(BOOL newVal);
	void SetInsertedProjectsLikeSummary(BOOL newVal);

	BOOL odl_GetMultipleCriticalPaths(void);
	BOOL GetMultipleCriticalPaths(void);
	void odl_SetMultipleCriticalPaths(BOOL newVal);
	void SetMultipleCriticalPaths(BOOL newVal);

	BOOL odl_GetNewTasksEffortDriven(void);
	BOOL GetNewTasksEffortDriven(void);
	void odl_SetNewTasksEffortDriven(BOOL newVal);
	void SetNewTasksEffortDriven(BOOL newVal);

	BOOL odl_GetNewTasksEstimated(void);
	BOOL GetNewTasksEstimated(void);
	void odl_SetNewTasksEstimated(BOOL newVal);
	void SetNewTasksEstimated(BOOL newVal);

	BOOL odl_GetSplitsInProgressTasks(void);
	BOOL GetSplitsInProgressTasks(void);
	void odl_SetSplitsInProgressTasks(BOOL newVal);
	void SetSplitsInProgressTasks(BOOL newVal);

	BOOL odl_GetSpreadActualCost(void);
	BOOL GetSpreadActualCost(void);
	void odl_SetSpreadActualCost(BOOL newVal);
	void SetSpreadActualCost(BOOL newVal);

	BOOL odl_GetSpreadPercentComplete(void);
	BOOL GetSpreadPercentComplete(void);
	void odl_SetSpreadPercentComplete(BOOL newVal);
	void SetSpreadPercentComplete(BOOL newVal);

	BOOL odl_GetTaskUpdatesResource(void);
	BOOL GetTaskUpdatesResource(void);
	void odl_SetTaskUpdatesResource(BOOL newVal);
	void SetTaskUpdatesResource(BOOL newVal);

	BOOL odl_GetFiscalYearStart(void);
	BOOL GetFiscalYearStart(void);
	void odl_SetFiscalYearStart(BOOL newVal);
	void SetFiscalYearStart(BOOL newVal);

	LONG odl_GetWeekStartDay(void);
	E_WEEKSTARTDAY GetWeekStartDay(void);
	void odl_SetWeekStartDay(LONG newVal);
	void SetWeekStartDay(E_WEEKSTARTDAY newVal);

	BOOL odl_GetMoveCompletedEndsBack(void);
	BOOL GetMoveCompletedEndsBack(void);
	void odl_SetMoveCompletedEndsBack(BOOL newVal);
	void SetMoveCompletedEndsBack(BOOL newVal);

	BOOL odl_GetMoveRemainingStartsBack(void);
	BOOL GetMoveRemainingStartsBack(void);
	void odl_SetMoveRemainingStartsBack(BOOL newVal);
	void SetMoveRemainingStartsBack(BOOL newVal);

	BOOL odl_GetMoveRemainingStartsForward(void);
	BOOL GetMoveRemainingStartsForward(void);
	void odl_SetMoveRemainingStartsForward(BOOL newVal);
	void SetMoveRemainingStartsForward(BOOL newVal);

	BOOL odl_GetMoveCompletedEndsForward(void);
	BOOL GetMoveCompletedEndsForward(void);
	void odl_SetMoveCompletedEndsForward(BOOL newVal);
	void SetMoveCompletedEndsForward(BOOL newVal);

	LONG odl_GetBaselineForEarnedValue(void);
	E_BASELINEFOREARNEDVALUE GetBaselineForEarnedValue(void);
	void odl_SetBaselineForEarnedValue(LONG newVal);
	void SetBaselineForEarnedValue(E_BASELINEFOREARNEDVALUE newVal);

	BOOL odl_GetAutoAddNewResourcesAndTasks(void);
	BOOL GetAutoAddNewResourcesAndTasks(void);
	void odl_SetAutoAddNewResourcesAndTasks(BOOL newVal);
	void SetAutoAddNewResourcesAndTasks(BOOL newVal);

	DATE odl_GetStatusDate(void);
	COleDateTime GetStatusDate(void);
	void odl_SetStatusDate(DATE newVal);
	void SetStatusDate(COleDateTime newVal);

	DATE odl_GetCurrentDate(void);
	COleDateTime GetCurrentDate(void);
	void odl_SetCurrentDate(DATE newVal);
	void SetCurrentDate(COleDateTime newVal);

	BOOL odl_GetMicrosoftProjectServerURL(void);
	BOOL GetMicrosoftProjectServerURL(void);
	void odl_SetMicrosoftProjectServerURL(BOOL newVal);
	void SetMicrosoftProjectServerURL(BOOL newVal);

	BOOL odl_GetAutolink(void);
	BOOL GetAutolink(void);
	void odl_SetAutolink(BOOL newVal);
	void SetAutolink(BOOL newVal);

	LONG odl_GetNewTaskStartDate(void);
	E_NEWTASKSTARTDATE GetNewTaskStartDate(void);
	void odl_SetNewTaskStartDate(LONG newVal);
	void SetNewTaskStartDate(E_NEWTASKSTARTDATE newVal);

	LONG odl_GetDefaultTaskEVMethod(void);
	E_DEFAULTTASKEVMETHOD GetDefaultTaskEVMethod(void);
	void odl_SetDefaultTaskEVMethod(LONG newVal);
	void SetDefaultTaskEVMethod(E_DEFAULTTASKEVMETHOD newVal);

	BOOL odl_GetProjectExternallyEdited(void);
	BOOL GetProjectExternallyEdited(void);
	void odl_SetProjectExternallyEdited(BOOL newVal);
	void SetProjectExternallyEdited(BOOL newVal);

	DATE odl_GetExtendedCreationDate(void);
	COleDateTime GetExtendedCreationDate(void);
	void odl_SetExtendedCreationDate(DATE newVal);
	void SetExtendedCreationDate(COleDateTime newVal);

	BOOL odl_GetActualsInSync(void);
	BOOL GetActualsInSync(void);
	void odl_SetActualsInSync(BOOL newVal);
	void SetActualsInSync(BOOL newVal);

	BOOL odl_GetRemoveFileProperties(void);
	BOOL GetRemoveFileProperties(void);
	void odl_SetRemoveFileProperties(BOOL newVal);
	void SetRemoveFileProperties(BOOL newVal);

	BOOL odl_GetAdminProject(void);
	BOOL GetAdminProject(void);
	void odl_SetAdminProject(BOOL newVal);
	void SetAdminProject(BOOL newVal);

	IDispatch* odl_GetOutlineCodes(void);
	OutlineCodes* GetOutlineCodes(void);

	IDispatch* odl_GetWBSMasks(void);
	WBSMasks* GetWBSMasks(void);

	IDispatch* odl_GetExtendedAttributes(void);
	ExtendedAttributes* GetExtendedAttributes(void);

	IDispatch* odl_GetCalendars(void);
	Calendars* GetCalendars(void);

	IDispatch* odl_GetTasks(void);
	Tasks* GetTasks(void);

	IDispatch* odl_GetResources(void);
	Resources* GetResources(void);

	IDispatch* odl_GetAssignments(void);
	Assignments* GetAssignments(void);


	void WriteXML(CString url);
	void ReadXML(CString url);
	void SetXML(CString sXML);
	CString GetXML();

	void odl_WriteXML(LPCTSTR url);
	void odl_ReadXML(LPCTSTR url);
	void odl_SetXML(LPCTSTR sXML);
	BSTR odl_GetXML(void);

	void mp_WriteXML(clsXML &oXML);
	void mp_ReadXML(clsXML &oXML);

	BOOL IsNull(void);

	void Initialize(void);

	void InitVars(void);

protected:
	DECLARE_DISPATCH_MAP()
	DECLARE_MESSAGE_MAP()
	DECLARE_OLECREATE(MP12)
	DECLARE_INTERFACE_MAP()

};