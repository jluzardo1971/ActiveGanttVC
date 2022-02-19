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



class Assignment : public clsItemBase
{
	DECLARE_DYNCREATE(Assignment)

public:

	Assignment();
	virtual ~Assignment();
	virtual void OnFinalRelease();

	clsCollectionBase* mp_oCollection;
	LONG mp_lUID;
	LONG mp_lTaskUID;
	LONG mp_lResourceUID;
	LONG mp_lPercentWorkComplete;
	DOUBLE mp_cActualCost;
	COleDateTime mp_dtActualFinish;
	DOUBLE mp_cActualOvertimeCost;
	Duration* mp_oActualOvertimeWork;
	COleDateTime mp_dtActualStart;
	Duration* mp_oActualWork;
	FLOAT mp_fACWP;
	BOOL mp_bConfirmed;
	DOUBLE mp_cCost;
	LONG mp_yCostRateTable;
	FLOAT mp_fCostVariance;
	FLOAT mp_fCV;
	LONG mp_lDelay;
	COleDateTime mp_dtFinish;
	LONG mp_lFinishVariance;
	CString mp_sHyperlink;
	CString mp_sHyperlinkAddress;
	CString mp_sHyperlinkSubAddress;
	FLOAT mp_fWorkVariance;
	BOOL mp_bHasFixedRateUnits;
	BOOL mp_bFixedMaterial;
	LONG mp_lLevelingDelay;
	LONG mp_yLevelingDelayFormat;
	BOOL mp_bLinkedFields;
	BOOL mp_bMilestone;
	CString mp_sNotes;
	BOOL mp_bOverallocated;
	DOUBLE mp_cOvertimeCost;
	Duration* mp_oOvertimeWork;
	Duration* mp_oRegularWork;
	DOUBLE mp_cRemainingCost;
	DOUBLE mp_cRemainingOvertimeCost;
	Duration* mp_oRemainingOvertimeWork;
	Duration* mp_oRemainingWork;
	BOOL mp_bResponsePending;
	COleDateTime mp_dtStart;
	COleDateTime mp_dtStop;
	COleDateTime mp_dtResume;
	LONG mp_lStartVariance;
	FLOAT mp_fUnits;
	BOOL mp_bUpdateNeeded;
	FLOAT mp_fVAC;
	Duration* mp_oWork;
	LONG mp_yWorkContour;
	FLOAT mp_fBCWS;
	FLOAT mp_fBCWP;
	LONG mp_yBookingType;
	Duration* mp_oActualWorkProtected;
	Duration* mp_oActualOvertimeWorkProtected;
	COleDateTime mp_dtCreationDate;
	AssignmentExtendedAttribute_C* mp_oExtendedAttribute_C;
	AssignmentBaseline_C* mp_oBaseline_C;
	TimephasedData_C* mp_oTimephasedData_C;

	LONG odl_GetUID(void);
	LONG GetUID(void);
	void odl_SetUID(LONG newVal);
	void SetUID(LONG newVal);

	LONG odl_GetTaskUID(void);
	LONG GetTaskUID(void);
	void odl_SetTaskUID(LONG newVal);
	void SetTaskUID(LONG newVal);

	LONG odl_GetResourceUID(void);
	LONG GetResourceUID(void);
	void odl_SetResourceUID(LONG newVal);
	void SetResourceUID(LONG newVal);

	LONG odl_GetPercentWorkComplete(void);
	LONG GetPercentWorkComplete(void);
	void odl_SetPercentWorkComplete(LONG newVal);
	void SetPercentWorkComplete(LONG newVal);

	DOUBLE odl_GetActualCost(void);
	DOUBLE GetActualCost(void);
	void odl_SetActualCost(DOUBLE newVal);
	void SetActualCost(DOUBLE newVal);

	DATE odl_GetActualFinish(void);
	COleDateTime GetActualFinish(void);
	void odl_SetActualFinish(DATE newVal);
	void SetActualFinish(COleDateTime newVal);

	DOUBLE odl_GetActualOvertimeCost(void);
	DOUBLE GetActualOvertimeCost(void);
	void odl_SetActualOvertimeCost(DOUBLE newVal);
	void SetActualOvertimeCost(DOUBLE newVal);

	IDispatch* odl_GetActualOvertimeWork(void);
	Duration* GetActualOvertimeWork(void);

	DATE odl_GetActualStart(void);
	COleDateTime GetActualStart(void);
	void odl_SetActualStart(DATE newVal);
	void SetActualStart(COleDateTime newVal);

	IDispatch* odl_GetActualWork(void);
	Duration* GetActualWork(void);

	FLOAT odl_GetACWP(void);
	FLOAT GetACWP(void);
	void odl_SetACWP(FLOAT newVal);
	void SetACWP(FLOAT newVal);

	BOOL odl_GetConfirmed(void);
	BOOL GetConfirmed(void);
	void odl_SetConfirmed(BOOL newVal);
	void SetConfirmed(BOOL newVal);

	DOUBLE odl_GetCost(void);
	DOUBLE GetCost(void);
	void odl_SetCost(DOUBLE newVal);
	void SetCost(DOUBLE newVal);

	LONG odl_GetCostRateTable(void);
	E_COSTRATETABLE GetCostRateTable(void);
	void odl_SetCostRateTable(LONG newVal);
	void SetCostRateTable(E_COSTRATETABLE newVal);

	FLOAT odl_GetCostVariance(void);
	FLOAT GetCostVariance(void);
	void odl_SetCostVariance(FLOAT newVal);
	void SetCostVariance(FLOAT newVal);

	FLOAT odl_GetCV(void);
	FLOAT GetCV(void);
	void odl_SetCV(FLOAT newVal);
	void SetCV(FLOAT newVal);

	LONG odl_GetDelay(void);
	LONG GetDelay(void);
	void odl_SetDelay(LONG newVal);
	void SetDelay(LONG newVal);

	DATE odl_GetFinish(void);
	COleDateTime GetFinish(void);
	void odl_SetFinish(DATE newVal);
	void SetFinish(COleDateTime newVal);

	LONG odl_GetFinishVariance(void);
	LONG GetFinishVariance(void);
	void odl_SetFinishVariance(LONG newVal);
	void SetFinishVariance(LONG newVal);

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

	FLOAT odl_GetWorkVariance(void);
	FLOAT GetWorkVariance(void);
	void odl_SetWorkVariance(FLOAT newVal);
	void SetWorkVariance(FLOAT newVal);

	BOOL odl_GetHasFixedRateUnits(void);
	BOOL GetHasFixedRateUnits(void);
	void odl_SetHasFixedRateUnits(BOOL newVal);
	void SetHasFixedRateUnits(BOOL newVal);

	BOOL odl_GetFixedMaterial(void);
	BOOL GetFixedMaterial(void);
	void odl_SetFixedMaterial(BOOL newVal);
	void SetFixedMaterial(BOOL newVal);

	LONG odl_GetLevelingDelay(void);
	LONG GetLevelingDelay(void);
	void odl_SetLevelingDelay(LONG newVal);
	void SetLevelingDelay(LONG newVal);

	LONG odl_GetLevelingDelayFormat(void);
	E_LEVELINGDELAYFORMAT GetLevelingDelayFormat(void);
	void odl_SetLevelingDelayFormat(LONG newVal);
	void SetLevelingDelayFormat(E_LEVELINGDELAYFORMAT newVal);

	BOOL odl_GetLinkedFields(void);
	BOOL GetLinkedFields(void);
	void odl_SetLinkedFields(BOOL newVal);
	void SetLinkedFields(BOOL newVal);

	BOOL odl_GetMilestone(void);
	BOOL GetMilestone(void);
	void odl_SetMilestone(BOOL newVal);
	void SetMilestone(BOOL newVal);

	BSTR odl_GetNotes(void);
	CString GetNotes(void);
	void odl_SetNotes(LPCTSTR newVal);
	void SetNotes(CString newVal);

	BOOL odl_GetOverallocated(void);
	BOOL GetOverallocated(void);
	void odl_SetOverallocated(BOOL newVal);
	void SetOverallocated(BOOL newVal);

	DOUBLE odl_GetOvertimeCost(void);
	DOUBLE GetOvertimeCost(void);
	void odl_SetOvertimeCost(DOUBLE newVal);
	void SetOvertimeCost(DOUBLE newVal);

	IDispatch* odl_GetOvertimeWork(void);
	Duration* GetOvertimeWork(void);

	IDispatch* odl_GetRegularWork(void);
	Duration* GetRegularWork(void);

	DOUBLE odl_GetRemainingCost(void);
	DOUBLE GetRemainingCost(void);
	void odl_SetRemainingCost(DOUBLE newVal);
	void SetRemainingCost(DOUBLE newVal);

	DOUBLE odl_GetRemainingOvertimeCost(void);
	DOUBLE GetRemainingOvertimeCost(void);
	void odl_SetRemainingOvertimeCost(DOUBLE newVal);
	void SetRemainingOvertimeCost(DOUBLE newVal);

	IDispatch* odl_GetRemainingOvertimeWork(void);
	Duration* GetRemainingOvertimeWork(void);

	IDispatch* odl_GetRemainingWork(void);
	Duration* GetRemainingWork(void);

	BOOL odl_GetResponsePending(void);
	BOOL GetResponsePending(void);
	void odl_SetResponsePending(BOOL newVal);
	void SetResponsePending(BOOL newVal);

	DATE odl_GetStart(void);
	COleDateTime GetStart(void);
	void odl_SetStart(DATE newVal);
	void SetStart(COleDateTime newVal);

	DATE odl_GetStop(void);
	COleDateTime GetStop(void);
	void odl_SetStop(DATE newVal);
	void SetStop(COleDateTime newVal);

	DATE odl_GetResume(void);
	COleDateTime GetResume(void);
	void odl_SetResume(DATE newVal);
	void SetResume(COleDateTime newVal);

	LONG odl_GetStartVariance(void);
	LONG GetStartVariance(void);
	void odl_SetStartVariance(LONG newVal);
	void SetStartVariance(LONG newVal);

	FLOAT odl_GetUnits(void);
	FLOAT GetUnits(void);
	void odl_SetUnits(FLOAT newVal);
	void SetUnits(FLOAT newVal);

	BOOL odl_GetUpdateNeeded(void);
	BOOL GetUpdateNeeded(void);
	void odl_SetUpdateNeeded(BOOL newVal);
	void SetUpdateNeeded(BOOL newVal);

	FLOAT odl_GetVAC(void);
	FLOAT GetVAC(void);
	void odl_SetVAC(FLOAT newVal);
	void SetVAC(FLOAT newVal);

	IDispatch* odl_GetWork(void);
	Duration* GetWork(void);

	LONG odl_GetWorkContour(void);
	E_WORKCONTOUR GetWorkContour(void);
	void odl_SetWorkContour(LONG newVal);
	void SetWorkContour(E_WORKCONTOUR newVal);

	FLOAT odl_GetBCWS(void);
	FLOAT GetBCWS(void);
	void odl_SetBCWS(FLOAT newVal);
	void SetBCWS(FLOAT newVal);

	FLOAT odl_GetBCWP(void);
	FLOAT GetBCWP(void);
	void odl_SetBCWP(FLOAT newVal);
	void SetBCWP(FLOAT newVal);

	LONG odl_GetBookingType(void);
	E_BOOKINGTYPE GetBookingType(void);
	void odl_SetBookingType(LONG newVal);
	void SetBookingType(E_BOOKINGTYPE newVal);

	IDispatch* odl_GetActualWorkProtected(void);
	Duration* GetActualWorkProtected(void);

	IDispatch* odl_GetActualOvertimeWorkProtected(void);
	Duration* GetActualOvertimeWorkProtected(void);

	DATE odl_GetCreationDate(void);
	COleDateTime GetCreationDate(void);
	void odl_SetCreationDate(DATE newVal);
	void SetCreationDate(COleDateTime newVal);

	IDispatch* odl_GetExtendedAttribute(void);
	AssignmentExtendedAttribute_C* GetExtendedAttribute(void);

	IDispatch* odl_GetBaseline(void);
	AssignmentBaseline_C* GetBaseline(void);

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
	DECLARE_OLECREATE(Assignment)
	DECLARE_INTERFACE_MAP()

};