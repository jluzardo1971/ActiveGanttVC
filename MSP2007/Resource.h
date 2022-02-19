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



class Resource : public clsItemBase
{
	DECLARE_DYNCREATE(Resource)

public:

	Resource();
	virtual ~Resource();
	virtual void OnFinalRelease();

	clsCollectionBase* mp_oCollection;
	LONG mp_lUID;
	LONG mp_lID;
	CString mp_sName;
	LONG mp_yType;
	BOOL mp_bIsNull;
	CString mp_sInitials;
	CString mp_sPhonetics;
	CString mp_sNTAccount;
	CString mp_sMaterialLabel;
	CString mp_sCode;
	CString mp_sGroup;
	LONG mp_yWorkGroup;
	CString mp_sEmailAddress;
	CString mp_sHyperlink;
	CString mp_sHyperlinkAddress;
	CString mp_sHyperlinkSubAddress;
	FLOAT mp_fMaxUnits;
	FLOAT mp_fPeakUnits;
	BOOL mp_bOverAllocated;
	COleDateTime mp_dtAvailableFrom;
	COleDateTime mp_dtAvailableTo;
	COleDateTime mp_dtStart;
	COleDateTime mp_dtFinish;
	BOOL mp_bCanLevel;
	LONG mp_yAccrueAt;
	Duration* mp_oWork;
	Duration* mp_oRegularWork;
	Duration* mp_oOvertimeWork;
	Duration* mp_oActualWork;
	Duration* mp_oRemainingWork;
	Duration* mp_oActualOvertimeWork;
	Duration* mp_oRemainingOvertimeWork;
	LONG mp_lPercentWorkComplete;
	DOUBLE mp_cStandardRate;
	LONG mp_yStandardRateFormat;
	DOUBLE mp_cCost;
	DOUBLE mp_cOvertimeRate;
	LONG mp_yOvertimeRateFormat;
	DOUBLE mp_cOvertimeCost;
	DOUBLE mp_cCostPerUse;
	DOUBLE mp_cActualCost;
	DOUBLE mp_cActualOvertimeCost;
	DOUBLE mp_cRemainingCost;
	DOUBLE mp_cRemainingOvertimeCost;
	FLOAT mp_fWorkVariance;
	FLOAT mp_fCostVariance;
	FLOAT mp_fSV;
	FLOAT mp_fCV;
	FLOAT mp_fACWP;
	LONG mp_lCalendarUID;
	CString mp_sNotes;
	FLOAT mp_fBCWS;
	FLOAT mp_fBCWP;
	BOOL mp_bIsGeneric;
	BOOL mp_bIsInactive;
	BOOL mp_bIsEnterprise;
	LONG mp_yBookingType;
	Duration* mp_oActualWorkProtected;
	Duration* mp_oActualOvertimeWorkProtected;
	CString mp_sActiveDirectoryGUID;
	COleDateTime mp_dtCreationDate;
	ResourceExtendedAttribute_C* mp_oExtendedAttribute_C;
	ResourceBaseline_C* mp_oBaseline_C;
	ResourceOutlineCode_C* mp_oOutlineCode_C;
	BOOL mp_bIsCostResource;
	CString mp_sAssnOwner;
	CString mp_sAssnOwnerGuid;
	BOOL mp_bIsBudget;
	ResourceAvailabilityPeriods* mp_oAvailabilityPeriods;
	ResourceRates* mp_oRates;
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
	E_TYPE_6 GetType(void);
	void odl_SetType(LONG newVal);
	void SetType(E_TYPE_6 newVal);

	BOOL odl_GetIsNull(void);
	BOOL GetIsNull(void);
	void odl_SetIsNull(BOOL newVal);
	void SetIsNull(BOOL newVal);

	BSTR odl_GetInitials(void);
	CString GetInitials(void);
	void odl_SetInitials(LPCTSTR newVal);
	void SetInitials(CString newVal);

	BSTR odl_GetPhonetics(void);
	CString GetPhonetics(void);
	void odl_SetPhonetics(LPCTSTR newVal);
	void SetPhonetics(CString newVal);

	BSTR odl_GetNTAccount(void);
	CString GetNTAccount(void);
	void odl_SetNTAccount(LPCTSTR newVal);
	void SetNTAccount(CString newVal);

	BSTR odl_GetMaterialLabel(void);
	CString GetMaterialLabel(void);
	void odl_SetMaterialLabel(LPCTSTR newVal);
	void SetMaterialLabel(CString newVal);

	BSTR odl_GetCode(void);
	CString GetCode(void);
	void odl_SetCode(LPCTSTR newVal);
	void SetCode(CString newVal);

	BSTR odl_GetGroup(void);
	CString GetGroup(void);
	void odl_SetGroup(LPCTSTR newVal);
	void SetGroup(CString newVal);

	LONG odl_GetWorkGroup(void);
	E_WORKGROUP GetWorkGroup(void);
	void odl_SetWorkGroup(LONG newVal);
	void SetWorkGroup(E_WORKGROUP newVal);

	BSTR odl_GetEmailAddress(void);
	CString GetEmailAddress(void);
	void odl_SetEmailAddress(LPCTSTR newVal);
	void SetEmailAddress(CString newVal);

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

	FLOAT odl_GetMaxUnits(void);
	FLOAT GetMaxUnits(void);
	void odl_SetMaxUnits(FLOAT newVal);
	void SetMaxUnits(FLOAT newVal);

	FLOAT odl_GetPeakUnits(void);
	FLOAT GetPeakUnits(void);
	void odl_SetPeakUnits(FLOAT newVal);
	void SetPeakUnits(FLOAT newVal);

	BOOL odl_GetOverAllocated(void);
	BOOL GetOverAllocated(void);
	void odl_SetOverAllocated(BOOL newVal);
	void SetOverAllocated(BOOL newVal);

	DATE odl_GetAvailableFrom(void);
	COleDateTime GetAvailableFrom(void);
	void odl_SetAvailableFrom(DATE newVal);
	void SetAvailableFrom(COleDateTime newVal);

	DATE odl_GetAvailableTo(void);
	COleDateTime GetAvailableTo(void);
	void odl_SetAvailableTo(DATE newVal);
	void SetAvailableTo(COleDateTime newVal);

	DATE odl_GetStart(void);
	COleDateTime GetStart(void);
	void odl_SetStart(DATE newVal);
	void SetStart(COleDateTime newVal);

	DATE odl_GetFinish(void);
	COleDateTime GetFinish(void);
	void odl_SetFinish(DATE newVal);
	void SetFinish(COleDateTime newVal);

	BOOL odl_GetCanLevel(void);
	BOOL GetCanLevel(void);
	void odl_SetCanLevel(BOOL newVal);
	void SetCanLevel(BOOL newVal);

	LONG odl_GetAccrueAt(void);
	E_ACCRUEAT GetAccrueAt(void);
	void odl_SetAccrueAt(LONG newVal);
	void SetAccrueAt(E_ACCRUEAT newVal);

	IDispatch* odl_GetWork(void);
	Duration* GetWork(void);

	IDispatch* odl_GetRegularWork(void);
	Duration* GetRegularWork(void);

	IDispatch* odl_GetOvertimeWork(void);
	Duration* GetOvertimeWork(void);

	IDispatch* odl_GetActualWork(void);
	Duration* GetActualWork(void);

	IDispatch* odl_GetRemainingWork(void);
	Duration* GetRemainingWork(void);

	IDispatch* odl_GetActualOvertimeWork(void);
	Duration* GetActualOvertimeWork(void);

	IDispatch* odl_GetRemainingOvertimeWork(void);
	Duration* GetRemainingOvertimeWork(void);

	LONG odl_GetPercentWorkComplete(void);
	LONG GetPercentWorkComplete(void);
	void odl_SetPercentWorkComplete(LONG newVal);
	void SetPercentWorkComplete(LONG newVal);

	DOUBLE odl_GetStandardRate(void);
	DOUBLE GetStandardRate(void);
	void odl_SetStandardRate(DOUBLE newVal);
	void SetStandardRate(DOUBLE newVal);

	LONG odl_GetStandardRateFormat(void);
	E_STANDARDRATEFORMAT GetStandardRateFormat(void);
	void odl_SetStandardRateFormat(LONG newVal);
	void SetStandardRateFormat(E_STANDARDRATEFORMAT newVal);

	DOUBLE odl_GetCost(void);
	DOUBLE GetCost(void);
	void odl_SetCost(DOUBLE newVal);
	void SetCost(DOUBLE newVal);

	DOUBLE odl_GetOvertimeRate(void);
	DOUBLE GetOvertimeRate(void);
	void odl_SetOvertimeRate(DOUBLE newVal);
	void SetOvertimeRate(DOUBLE newVal);

	LONG odl_GetOvertimeRateFormat(void);
	E_OVERTIMERATEFORMAT GetOvertimeRateFormat(void);
	void odl_SetOvertimeRateFormat(LONG newVal);
	void SetOvertimeRateFormat(E_OVERTIMERATEFORMAT newVal);

	DOUBLE odl_GetOvertimeCost(void);
	DOUBLE GetOvertimeCost(void);
	void odl_SetOvertimeCost(DOUBLE newVal);
	void SetOvertimeCost(DOUBLE newVal);

	DOUBLE odl_GetCostPerUse(void);
	DOUBLE GetCostPerUse(void);
	void odl_SetCostPerUse(DOUBLE newVal);
	void SetCostPerUse(DOUBLE newVal);

	DOUBLE odl_GetActualCost(void);
	DOUBLE GetActualCost(void);
	void odl_SetActualCost(DOUBLE newVal);
	void SetActualCost(DOUBLE newVal);

	DOUBLE odl_GetActualOvertimeCost(void);
	DOUBLE GetActualOvertimeCost(void);
	void odl_SetActualOvertimeCost(DOUBLE newVal);
	void SetActualOvertimeCost(DOUBLE newVal);

	DOUBLE odl_GetRemainingCost(void);
	DOUBLE GetRemainingCost(void);
	void odl_SetRemainingCost(DOUBLE newVal);
	void SetRemainingCost(DOUBLE newVal);

	DOUBLE odl_GetRemainingOvertimeCost(void);
	DOUBLE GetRemainingOvertimeCost(void);
	void odl_SetRemainingOvertimeCost(DOUBLE newVal);
	void SetRemainingOvertimeCost(DOUBLE newVal);

	FLOAT odl_GetWorkVariance(void);
	FLOAT GetWorkVariance(void);
	void odl_SetWorkVariance(FLOAT newVal);
	void SetWorkVariance(FLOAT newVal);

	FLOAT odl_GetCostVariance(void);
	FLOAT GetCostVariance(void);
	void odl_SetCostVariance(FLOAT newVal);
	void SetCostVariance(FLOAT newVal);

	FLOAT odl_GetSV(void);
	FLOAT GetSV(void);
	void odl_SetSV(FLOAT newVal);
	void SetSV(FLOAT newVal);

	FLOAT odl_GetCV(void);
	FLOAT GetCV(void);
	void odl_SetCV(FLOAT newVal);
	void SetCV(FLOAT newVal);

	FLOAT odl_GetACWP(void);
	FLOAT GetACWP(void);
	void odl_SetACWP(FLOAT newVal);
	void SetACWP(FLOAT newVal);

	LONG odl_GetCalendarUID(void);
	LONG GetCalendarUID(void);
	void odl_SetCalendarUID(LONG newVal);
	void SetCalendarUID(LONG newVal);

	BSTR odl_GetNotes(void);
	CString GetNotes(void);
	void odl_SetNotes(LPCTSTR newVal);
	void SetNotes(CString newVal);

	FLOAT odl_GetBCWS(void);
	FLOAT GetBCWS(void);
	void odl_SetBCWS(FLOAT newVal);
	void SetBCWS(FLOAT newVal);

	FLOAT odl_GetBCWP(void);
	FLOAT GetBCWP(void);
	void odl_SetBCWP(FLOAT newVal);
	void SetBCWP(FLOAT newVal);

	BOOL odl_GetIsGeneric(void);
	BOOL GetIsGeneric(void);
	void odl_SetIsGeneric(BOOL newVal);
	void SetIsGeneric(BOOL newVal);

	BOOL odl_GetIsInactive(void);
	BOOL GetIsInactive(void);
	void odl_SetIsInactive(BOOL newVal);
	void SetIsInactive(BOOL newVal);

	BOOL odl_GetIsEnterprise(void);
	BOOL GetIsEnterprise(void);
	void odl_SetIsEnterprise(BOOL newVal);
	void SetIsEnterprise(BOOL newVal);

	LONG odl_GetBookingType(void);
	E_BOOKINGTYPE GetBookingType(void);
	void odl_SetBookingType(LONG newVal);
	void SetBookingType(E_BOOKINGTYPE newVal);

	IDispatch* odl_GetActualWorkProtected(void);
	Duration* GetActualWorkProtected(void);

	IDispatch* odl_GetActualOvertimeWorkProtected(void);
	Duration* GetActualOvertimeWorkProtected(void);

	BSTR odl_GetActiveDirectoryGUID(void);
	CString GetActiveDirectoryGUID(void);
	void odl_SetActiveDirectoryGUID(LPCTSTR newVal);
	void SetActiveDirectoryGUID(CString newVal);

	DATE odl_GetCreationDate(void);
	COleDateTime GetCreationDate(void);
	void odl_SetCreationDate(DATE newVal);
	void SetCreationDate(COleDateTime newVal);

	IDispatch* odl_GetExtendedAttribute(void);
	ResourceExtendedAttribute_C* GetExtendedAttribute(void);

	IDispatch* odl_GetBaseline(void);
	ResourceBaseline_C* GetBaseline(void);

	IDispatch* odl_GetOutlineCode(void);
	ResourceOutlineCode_C* GetOutlineCode(void);

	BOOL odl_GetIsCostResource(void);
	BOOL GetIsCostResource(void);
	void odl_SetIsCostResource(BOOL newVal);
	void SetIsCostResource(BOOL newVal);

	BSTR odl_GetAssnOwner(void);
	CString GetAssnOwner(void);
	void odl_SetAssnOwner(LPCTSTR newVal);
	void SetAssnOwner(CString newVal);

	BSTR odl_GetAssnOwnerGuid(void);
	CString GetAssnOwnerGuid(void);
	void odl_SetAssnOwnerGuid(LPCTSTR newVal);
	void SetAssnOwnerGuid(CString newVal);

	BOOL odl_GetIsBudget(void);
	BOOL GetIsBudget(void);
	void odl_SetIsBudget(BOOL newVal);
	void SetIsBudget(BOOL newVal);

	IDispatch* odl_GetAvailabilityPeriods(void);
	ResourceAvailabilityPeriods* GetAvailabilityPeriods(void);

	IDispatch* odl_GetRates(void);
	ResourceRates* GetRates(void);

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
	DECLARE_OLECREATE(Resource)
	DECLARE_INTERFACE_MAP()

};
