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



class Task : public clsItemBase
{
	DECLARE_DYNCREATE(Task)

public:

	Task();
	virtual ~Task();
	virtual void OnFinalRelease();

	clsCollectionBase* mp_oCollection;
	LONG mp_lUID;
	LONG mp_lID;
	CString mp_sName;
	LONG mp_yType;
	BOOL mp_bIsNull;
	COleDateTime mp_dtCreateDate;
	CString mp_sContact;
	CString mp_sWBS;
	CString mp_sWBSLevel;
	CString mp_sOutlineNumber;
	LONG mp_lOutlineLevel;
	LONG mp_lPriority;
	COleDateTime mp_dtStart;
	COleDateTime mp_dtFinish;
	Duration* mp_oDuration;
	LONG mp_yDurationFormat;
	Duration* mp_oWork;
	COleDateTime mp_dtStop;
	COleDateTime mp_dtResume;
	BOOL mp_bResumeValid;
	BOOL mp_bEffortDriven;
	BOOL mp_bRecurring;
	BOOL mp_bOverAllocated;
	BOOL mp_bEstimated;
	BOOL mp_bMilestone;
	BOOL mp_bSummary;
	BOOL mp_bDisplayAsSummary;
	BOOL mp_bCritical;
	BOOL mp_bIsSubproject;
	BOOL mp_bIsSubprojectReadOnly;
	CString mp_sSubprojectName;
	BOOL mp_bExternalTask;
	CString mp_sExternalTaskProject;
	COleDateTime mp_dtEarlyStart;
	COleDateTime mp_dtEarlyFinish;
	COleDateTime mp_dtLateStart;
	COleDateTime mp_dtLateFinish;
	LONG mp_lStartVariance;
	LONG mp_lFinishVariance;
	FLOAT mp_fWorkVariance;
	LONG mp_lFreeSlack;
	LONG mp_lStartSlack;
	LONG mp_lFinishSlack;
	LONG mp_lTotalSlack;
	FLOAT mp_fFixedCost;
	LONG mp_yFixedCostAccrual;
	LONG mp_lPercentComplete;
	LONG mp_lPercentWorkComplete;
	DOUBLE mp_cCost;
	DOUBLE mp_cOvertimeCost;
	Duration* mp_oOvertimeWork;
	COleDateTime mp_dtActualStart;
	COleDateTime mp_dtActualFinish;
	Duration* mp_oActualDuration;
	DOUBLE mp_cActualCost;
	DOUBLE mp_cActualOvertimeCost;
	Duration* mp_oActualWork;
	Duration* mp_oActualOvertimeWork;
	Duration* mp_oRegularWork;
	Duration* mp_oRemainingDuration;
	DOUBLE mp_cRemainingCost;
	Duration* mp_oRemainingWork;
	DOUBLE mp_cRemainingOvertimeCost;
	Duration* mp_oRemainingOvertimeWork;
	FLOAT mp_fACWP;
	FLOAT mp_fCV;
	LONG mp_yConstraintType;
	LONG mp_lCalendarUID;
	COleDateTime mp_dtConstraintDate;
	COleDateTime mp_dtDeadline;
	BOOL mp_bLevelAssignments;
	BOOL mp_bLevelingCanSplit;
	LONG mp_lLevelingDelay;
	LONG mp_yLevelingDelayFormat;
	COleDateTime mp_dtPreLeveledStart;
	COleDateTime mp_dtPreLeveledFinish;
	CString mp_sHyperlink;
	CString mp_sHyperlinkAddress;
	CString mp_sHyperlinkSubAddress;
	BOOL mp_bIgnoreResourceCalendar;
	CString mp_sNotes;
	BOOL mp_bHideBar;
	BOOL mp_bRollup;
	FLOAT mp_fBCWS;
	FLOAT mp_fBCWP;
	LONG mp_lPhysicalPercentComplete;
	LONG mp_yEarnedValueMethod;
	TaskPredecessorLink_C* mp_oPredecessorLink_C;
	Duration* mp_oActualWorkProtected;
	Duration* mp_oActualOvertimeWorkProtected;
	TaskExtendedAttribute_C* mp_oExtendedAttribute_C;
	TaskBaseline_C* mp_oBaseline_C;
	TaskOutlineCode_C* mp_oOutlineCode_C;
	BOOL mp_bIsPublished;
	CString mp_sStatusManager;
	COleDateTime mp_dtCommitmentStart;
	COleDateTime mp_dtCommitmentFinish;
	LONG mp_yCommitmentType;
	BOOL mp_bActive;
	BOOL mp_bPinned;
	CString mp_sPinnedStart;
	CString mp_sPinnedFinish;
	CString mp_sPinnedDuration;
	TimephasedData_C* mp_oTimephasedData_C;

	LONG odl_GetUID(void);
	LONG GetUID(void);
	void odl_SetUID(LONG newVal);
	void SetUID(LONG newVal);

	LONG odl_GetID(void);
	LONG GetID(void);
	void odl_SetID(LONG newVal);
	void SetID(LONG newVal);

	BSTR odl_GetName(void);
	CString GetName(void);
	void odl_SetName(LPCTSTR newVal);
	void SetName(CString newVal);

	LONG odl_GetType(void);
	E_TYPE_4 GetType(void);
	void odl_SetType(LONG newVal);
	void SetType(E_TYPE_4 newVal);

	BOOL odl_GetIsNull(void);
	BOOL GetIsNull(void);
	void odl_SetIsNull(BOOL newVal);
	void SetIsNull(BOOL newVal);

	DATE odl_GetCreateDate(void);
	COleDateTime GetCreateDate(void);
	void odl_SetCreateDate(DATE newVal);
	void SetCreateDate(COleDateTime newVal);

	BSTR odl_GetContact(void);
	CString GetContact(void);
	void odl_SetContact(LPCTSTR newVal);
	void SetContact(CString newVal);

	BSTR odl_GetWBS(void);
	CString GetWBS(void);
	void odl_SetWBS(LPCTSTR newVal);
	void SetWBS(CString newVal);

	BSTR odl_GetWBSLevel(void);
	CString GetWBSLevel(void);
	void odl_SetWBSLevel(LPCTSTR newVal);
	void SetWBSLevel(CString newVal);

	BSTR odl_GetOutlineNumber(void);
	CString GetOutlineNumber(void);
	void odl_SetOutlineNumber(LPCTSTR newVal);
	void SetOutlineNumber(CString newVal);

	LONG odl_GetOutlineLevel(void);
	LONG GetOutlineLevel(void);
	void odl_SetOutlineLevel(LONG newVal);
	void SetOutlineLevel(LONG newVal);

	LONG odl_GetPriority(void);
	LONG GetPriority(void);
	void odl_SetPriority(LONG newVal);
	void SetPriority(LONG newVal);

	DATE odl_GetStart(void);
	COleDateTime GetStart(void);
	void odl_SetStart(DATE newVal);
	void SetStart(COleDateTime newVal);

	DATE odl_GetFinish(void);
	COleDateTime GetFinish(void);
	void odl_SetFinish(DATE newVal);
	void SetFinish(COleDateTime newVal);

	IDispatch* odl_GetDuration(void);
	Duration* GetDuration(void);

	LONG odl_GetDurationFormat(void);
	E_DURATIONFORMAT GetDurationFormat(void);
	void odl_SetDurationFormat(LONG newVal);
	void SetDurationFormat(E_DURATIONFORMAT newVal);

	IDispatch* odl_GetWork(void);
	Duration* GetWork(void);

	DATE odl_GetStop(void);
	COleDateTime GetStop(void);
	void odl_SetStop(DATE newVal);
	void SetStop(COleDateTime newVal);

	DATE odl_GetResume(void);
	COleDateTime GetResume(void);
	void odl_SetResume(DATE newVal);
	void SetResume(COleDateTime newVal);

	BOOL odl_GetResumeValid(void);
	BOOL GetResumeValid(void);
	void odl_SetResumeValid(BOOL newVal);
	void SetResumeValid(BOOL newVal);

	BOOL odl_GetEffortDriven(void);
	BOOL GetEffortDriven(void);
	void odl_SetEffortDriven(BOOL newVal);
	void SetEffortDriven(BOOL newVal);

	BOOL odl_GetRecurring(void);
	BOOL GetRecurring(void);
	void odl_SetRecurring(BOOL newVal);
	void SetRecurring(BOOL newVal);

	BOOL odl_GetOverAllocated(void);
	BOOL GetOverAllocated(void);
	void odl_SetOverAllocated(BOOL newVal);
	void SetOverAllocated(BOOL newVal);

	BOOL odl_GetEstimated(void);
	BOOL GetEstimated(void);
	void odl_SetEstimated(BOOL newVal);
	void SetEstimated(BOOL newVal);

	BOOL odl_GetMilestone(void);
	BOOL GetMilestone(void);
	void odl_SetMilestone(BOOL newVal);
	void SetMilestone(BOOL newVal);

	BOOL odl_GetSummary(void);
	BOOL GetSummary(void);
	void odl_SetSummary(BOOL newVal);
	void SetSummary(BOOL newVal);

	BOOL odl_GetDisplayAsSummary(void);
	BOOL GetDisplayAsSummary(void);
	void odl_SetDisplayAsSummary(BOOL newVal);
	void SetDisplayAsSummary(BOOL newVal);

	BOOL odl_GetCritical(void);
	BOOL GetCritical(void);
	void odl_SetCritical(BOOL newVal);
	void SetCritical(BOOL newVal);

	BOOL odl_GetIsSubproject(void);
	BOOL GetIsSubproject(void);
	void odl_SetIsSubproject(BOOL newVal);
	void SetIsSubproject(BOOL newVal);

	BOOL odl_GetIsSubprojectReadOnly(void);
	BOOL GetIsSubprojectReadOnly(void);
	void odl_SetIsSubprojectReadOnly(BOOL newVal);
	void SetIsSubprojectReadOnly(BOOL newVal);

	BSTR odl_GetSubprojectName(void);
	CString GetSubprojectName(void);
	void odl_SetSubprojectName(LPCTSTR newVal);
	void SetSubprojectName(CString newVal);

	BOOL odl_GetExternalTask(void);
	BOOL GetExternalTask(void);
	void odl_SetExternalTask(BOOL newVal);
	void SetExternalTask(BOOL newVal);

	BSTR odl_GetExternalTaskProject(void);
	CString GetExternalTaskProject(void);
	void odl_SetExternalTaskProject(LPCTSTR newVal);
	void SetExternalTaskProject(CString newVal);

	DATE odl_GetEarlyStart(void);
	COleDateTime GetEarlyStart(void);
	void odl_SetEarlyStart(DATE newVal);
	void SetEarlyStart(COleDateTime newVal);

	DATE odl_GetEarlyFinish(void);
	COleDateTime GetEarlyFinish(void);
	void odl_SetEarlyFinish(DATE newVal);
	void SetEarlyFinish(COleDateTime newVal);

	DATE odl_GetLateStart(void);
	COleDateTime GetLateStart(void);
	void odl_SetLateStart(DATE newVal);
	void SetLateStart(COleDateTime newVal);

	DATE odl_GetLateFinish(void);
	COleDateTime GetLateFinish(void);
	void odl_SetLateFinish(DATE newVal);
	void SetLateFinish(COleDateTime newVal);

	LONG odl_GetStartVariance(void);
	LONG GetStartVariance(void);
	void odl_SetStartVariance(LONG newVal);
	void SetStartVariance(LONG newVal);

	LONG odl_GetFinishVariance(void);
	LONG GetFinishVariance(void);
	void odl_SetFinishVariance(LONG newVal);
	void SetFinishVariance(LONG newVal);

	FLOAT odl_GetWorkVariance(void);
	FLOAT GetWorkVariance(void);
	void odl_SetWorkVariance(FLOAT newVal);
	void SetWorkVariance(FLOAT newVal);

	LONG odl_GetFreeSlack(void);
	LONG GetFreeSlack(void);
	void odl_SetFreeSlack(LONG newVal);
	void SetFreeSlack(LONG newVal);

	LONG odl_GetStartSlack(void);
	LONG GetStartSlack(void);
	void odl_SetStartSlack(LONG newVal);
	void SetStartSlack(LONG newVal);

	LONG odl_GetFinishSlack(void);
	LONG GetFinishSlack(void);
	void odl_SetFinishSlack(LONG newVal);
	void SetFinishSlack(LONG newVal);

	LONG odl_GetTotalSlack(void);
	LONG GetTotalSlack(void);
	void odl_SetTotalSlack(LONG newVal);
	void SetTotalSlack(LONG newVal);

	FLOAT odl_GetFixedCost(void);
	FLOAT GetFixedCost(void);
	void odl_SetFixedCost(FLOAT newVal);
	void SetFixedCost(FLOAT newVal);

	LONG odl_GetFixedCostAccrual(void);
	E_FIXEDCOSTACCRUAL GetFixedCostAccrual(void);
	void odl_SetFixedCostAccrual(LONG newVal);
	void SetFixedCostAccrual(E_FIXEDCOSTACCRUAL newVal);

	LONG odl_GetPercentComplete(void);
	LONG GetPercentComplete(void);
	void odl_SetPercentComplete(LONG newVal);
	void SetPercentComplete(LONG newVal);

	LONG odl_GetPercentWorkComplete(void);
	LONG GetPercentWorkComplete(void);
	void odl_SetPercentWorkComplete(LONG newVal);
	void SetPercentWorkComplete(LONG newVal);

	DOUBLE odl_GetCost(void);
	DOUBLE GetCost(void);
	void odl_SetCost(DOUBLE newVal);
	void SetCost(DOUBLE newVal);

	DOUBLE odl_GetOvertimeCost(void);
	DOUBLE GetOvertimeCost(void);
	void odl_SetOvertimeCost(DOUBLE newVal);
	void SetOvertimeCost(DOUBLE newVal);

	IDispatch* odl_GetOvertimeWork(void);
	Duration* GetOvertimeWork(void);

	DATE odl_GetActualStart(void);
	COleDateTime GetActualStart(void);
	void odl_SetActualStart(DATE newVal);
	void SetActualStart(COleDateTime newVal);

	DATE odl_GetActualFinish(void);
	COleDateTime GetActualFinish(void);
	void odl_SetActualFinish(DATE newVal);
	void SetActualFinish(COleDateTime newVal);

	IDispatch* odl_GetActualDuration(void);
	Duration* GetActualDuration(void);

	DOUBLE odl_GetActualCost(void);
	DOUBLE GetActualCost(void);
	void odl_SetActualCost(DOUBLE newVal);
	void SetActualCost(DOUBLE newVal);

	DOUBLE odl_GetActualOvertimeCost(void);
	DOUBLE GetActualOvertimeCost(void);
	void odl_SetActualOvertimeCost(DOUBLE newVal);
	void SetActualOvertimeCost(DOUBLE newVal);

	IDispatch* odl_GetActualWork(void);
	Duration* GetActualWork(void);

	IDispatch* odl_GetActualOvertimeWork(void);
	Duration* GetActualOvertimeWork(void);

	IDispatch* odl_GetRegularWork(void);
	Duration* GetRegularWork(void);

	IDispatch* odl_GetRemainingDuration(void);
	Duration* GetRemainingDuration(void);

	DOUBLE odl_GetRemainingCost(void);
	DOUBLE GetRemainingCost(void);
	void odl_SetRemainingCost(DOUBLE newVal);
	void SetRemainingCost(DOUBLE newVal);

	IDispatch* odl_GetRemainingWork(void);
	Duration* GetRemainingWork(void);

	DOUBLE odl_GetRemainingOvertimeCost(void);
	DOUBLE GetRemainingOvertimeCost(void);
	void odl_SetRemainingOvertimeCost(DOUBLE newVal);
	void SetRemainingOvertimeCost(DOUBLE newVal);

	IDispatch* odl_GetRemainingOvertimeWork(void);
	Duration* GetRemainingOvertimeWork(void);

	FLOAT odl_GetACWP(void);
	FLOAT GetACWP(void);
	void odl_SetACWP(FLOAT newVal);
	void SetACWP(FLOAT newVal);

	FLOAT odl_GetCV(void);
	FLOAT GetCV(void);
	void odl_SetCV(FLOAT newVal);
	void SetCV(FLOAT newVal);

	LONG odl_GetConstraintType(void);
	E_CONSTRAINTTYPE GetConstraintType(void);
	void odl_SetConstraintType(LONG newVal);
	void SetConstraintType(E_CONSTRAINTTYPE newVal);

	LONG odl_GetCalendarUID(void);
	LONG GetCalendarUID(void);
	void odl_SetCalendarUID(LONG newVal);
	void SetCalendarUID(LONG newVal);

	DATE odl_GetConstraintDate(void);
	COleDateTime GetConstraintDate(void);
	void odl_SetConstraintDate(DATE newVal);
	void SetConstraintDate(COleDateTime newVal);

	DATE odl_GetDeadline(void);
	COleDateTime GetDeadline(void);
	void odl_SetDeadline(DATE newVal);
	void SetDeadline(COleDateTime newVal);

	BOOL odl_GetLevelAssignments(void);
	BOOL GetLevelAssignments(void);
	void odl_SetLevelAssignments(BOOL newVal);
	void SetLevelAssignments(BOOL newVal);

	BOOL odl_GetLevelingCanSplit(void);
	BOOL GetLevelingCanSplit(void);
	void odl_SetLevelingCanSplit(BOOL newVal);
	void SetLevelingCanSplit(BOOL newVal);

	LONG odl_GetLevelingDelay(void);
	LONG GetLevelingDelay(void);
	void odl_SetLevelingDelay(LONG newVal);
	void SetLevelingDelay(LONG newVal);

	LONG odl_GetLevelingDelayFormat(void);
	E_LEVELINGDELAYFORMAT GetLevelingDelayFormat(void);
	void odl_SetLevelingDelayFormat(LONG newVal);
	void SetLevelingDelayFormat(E_LEVELINGDELAYFORMAT newVal);

	DATE odl_GetPreLeveledStart(void);
	COleDateTime GetPreLeveledStart(void);
	void odl_SetPreLeveledStart(DATE newVal);
	void SetPreLeveledStart(COleDateTime newVal);

	DATE odl_GetPreLeveledFinish(void);
	COleDateTime GetPreLeveledFinish(void);
	void odl_SetPreLeveledFinish(DATE newVal);
	void SetPreLeveledFinish(COleDateTime newVal);

	BSTR odl_GetHyperlink(void);
	CString GetHyperlink(void);
	void odl_SetHyperlink(LPCTSTR newVal);
	void SetHyperlink(CString newVal);

	BSTR odl_GetHyperlinkAddress(void);
	CString GetHyperlinkAddress(void);
	void odl_SetHyperlinkAddress(LPCTSTR newVal);
	void SetHyperlinkAddress(CString newVal);

	BSTR odl_GetHyperlinkSubAddress(void);
	CString GetHyperlinkSubAddress(void);
	void odl_SetHyperlinkSubAddress(LPCTSTR newVal);
	void SetHyperlinkSubAddress(CString newVal);

	BOOL odl_GetIgnoreResourceCalendar(void);
	BOOL GetIgnoreResourceCalendar(void);
	void odl_SetIgnoreResourceCalendar(BOOL newVal);
	void SetIgnoreResourceCalendar(BOOL newVal);

	BSTR odl_GetNotes(void);
	CString GetNotes(void);
	void odl_SetNotes(LPCTSTR newVal);
	void SetNotes(CString newVal);

	BOOL odl_GetHideBar(void);
	BOOL GetHideBar(void);
	void odl_SetHideBar(BOOL newVal);
	void SetHideBar(BOOL newVal);

	BOOL odl_GetRollup(void);
	BOOL GetRollup(void);
	void odl_SetRollup(BOOL newVal);
	void SetRollup(BOOL newVal);

	FLOAT odl_GetBCWS(void);
	FLOAT GetBCWS(void);
	void odl_SetBCWS(FLOAT newVal);
	void SetBCWS(FLOAT newVal);

	FLOAT odl_GetBCWP(void);
	FLOAT GetBCWP(void);
	void odl_SetBCWP(FLOAT newVal);
	void SetBCWP(FLOAT newVal);

	LONG odl_GetPhysicalPercentComplete(void);
	LONG GetPhysicalPercentComplete(void);
	void odl_SetPhysicalPercentComplete(LONG newVal);
	void SetPhysicalPercentComplete(LONG newVal);

	LONG odl_GetEarnedValueMethod(void);
	E_EARNEDVALUEMETHOD GetEarnedValueMethod(void);
	void odl_SetEarnedValueMethod(LONG newVal);
	void SetEarnedValueMethod(E_EARNEDVALUEMETHOD newVal);

	IDispatch* odl_GetPredecessorLink(void);
	TaskPredecessorLink_C* GetPredecessorLink(void);

	IDispatch* odl_GetActualWorkProtected(void);
	Duration* GetActualWorkProtected(void);

	IDispatch* odl_GetActualOvertimeWorkProtected(void);
	Duration* GetActualOvertimeWorkProtected(void);

	IDispatch* odl_GetExtendedAttribute(void);
	TaskExtendedAttribute_C* GetExtendedAttribute(void);

	IDispatch* odl_GetBaseline(void);
	TaskBaseline_C* GetBaseline(void);

	IDispatch* odl_GetOutlineCode(void);
	TaskOutlineCode_C* GetOutlineCode(void);

	BOOL odl_GetIsPublished(void);
	BOOL GetIsPublished(void);
	void odl_SetIsPublished(BOOL newVal);
	void SetIsPublished(BOOL newVal);

	BSTR odl_GetStatusManager(void);
	CString GetStatusManager(void);
	void odl_SetStatusManager(LPCTSTR newVal);
	void SetStatusManager(CString newVal);

	DATE odl_GetCommitmentStart(void);
	COleDateTime GetCommitmentStart(void);
	void odl_SetCommitmentStart(DATE newVal);
	void SetCommitmentStart(COleDateTime newVal);

	DATE odl_GetCommitmentFinish(void);
	COleDateTime GetCommitmentFinish(void);
	void odl_SetCommitmentFinish(DATE newVal);
	void SetCommitmentFinish(COleDateTime newVal);

	LONG odl_GetCommitmentType(void);
	E_COMMITMENTTYPE GetCommitmentType(void);
	void odl_SetCommitmentType(LONG newVal);
	void SetCommitmentType(E_COMMITMENTTYPE newVal);

	BOOL odl_GetActive(void);
	BOOL GetActive(void);
	void odl_SetActive(BOOL newVal);
	void SetActive(BOOL newVal);

	BOOL odl_GetPinned(void);
	BOOL GetPinned(void);
	void odl_SetPinned(BOOL newVal);
	void SetPinned(BOOL newVal);

	BSTR odl_GetPinnedStart(void);
	CString GetPinnedStart(void);
	void odl_SetPinnedStart(LPCTSTR newVal);
	void SetPinnedStart(CString newVal);

	BSTR odl_GetPinnedFinish(void);
	CString GetPinnedFinish(void);
	void odl_SetPinnedFinish(LPCTSTR newVal);
	void SetPinnedFinish(CString newVal);

	BSTR odl_GetPinnedDuration(void);
	CString GetPinnedDuration(void);
	void odl_SetPinnedDuration(LPCTSTR newVal);
	void SetPinnedDuration(CString newVal);

	IDispatch* odl_GetTimephasedData(void);
	TimephasedData_C* GetTimephasedData(void);

	BSTR odl_GetKey(void);
	CString GetKey(void);
	void odl_SetKey(LPCTSTR newVal);
	void SetKey(CString newVal);


	BSTR odl_GetXML(void);
	CString GetXML(void);

	void odl_SetXML(LPCTSTR sXML);
	void SetXML(CString sXML);

	BOOL IsNull(void);

	void Initialize(void);

	void InitVars(void);

protected:
	DECLARE_DISPATCH_MAP()
	DECLARE_MESSAGE_MAP()
	DECLARE_OLECREATE(Task)
	DECLARE_INTERFACE_MAP()

};