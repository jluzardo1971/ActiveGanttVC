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
	FLOAT mp_fPeakUnits;
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
	BOOL mp_bSummary;
	FLOAT mp_fSV;
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
	CString mp_sAssnOwner;
	CString mp_sAssnOwnerGuid;
	DOUBLE mp_cBudgetCost;
	Duration* mp_oBudgetWork;
	AssignmentExtendedAttribute_C* mp_oExtendedAttribute_C;
	AssignmentBaseline_C* mp_oBaseline_C;
	CString mp_sf404000;
	CString mp_sf404001;
	CString mp_sf404002;
	CString mp_sf404003;
	CString mp_sf404004;
	CString mp_sf404005;
	CString mp_sf404006;
	CString mp_sf404007;
	CString mp_sf404008;
	CString mp_sf404009;
	CString mp_sf40400a;
	CString mp_sf40400b;
	CString mp_sf40400c;
	CString mp_sf40400d;
	CString mp_sf40400e;
	CString mp_sf40400f;
	CString mp_sf404010;
	CString mp_sf404011;
	CString mp_sf404012;
	CString mp_sf404013;
	CString mp_sf404014;
	CString mp_sf404015;
	CString mp_sf404016;
	CString mp_sf404017;
	CString mp_sf404018;
	CString mp_sf404019;
	CString mp_sf40401a;
	CString mp_sf40401b;
	CString mp_sf40401c;
	CString mp_sf40401d;
	CString mp_sf40401e;
	CString mp_sf40401f;
	CString mp_sf404020;
	CString mp_sf404021;
	CString mp_sf404022;
	CString mp_sf404023;
	CString mp_sf404024;
	CString mp_sf404025;
	CString mp_sf404026;
	CString mp_sf404027;
	CString mp_sf404028;
	CString mp_sf404029;
	CString mp_sf40402a;
	CString mp_sf40402b;
	CString mp_sf40402c;
	CString mp_sf40402d;
	CString mp_sf40402e;
	CString mp_sf40402f;
	CString mp_sf404030;
	CString mp_sf404031;
	CString mp_sf404032;
	CString mp_sf404033;
	CString mp_sf404034;
	CString mp_sf404035;
	CString mp_sf404036;
	CString mp_sf404037;
	CString mp_sf404038;
	CString mp_sf404039;
	CString mp_sf40403a;
	CString mp_sf40403b;
	CString mp_sf40403c;
	CString mp_sf40403d;
	CString mp_sf40403e;
	CString mp_sf40403f;
	CString mp_sf404040;
	CString mp_sf404041;
	CString mp_sf404042;
	CString mp_sf404043;
	CString mp_sf404044;
	CString mp_sf404045;
	CString mp_sf404046;
	CString mp_sf404047;
	CString mp_sf404048;
	CString mp_sf404049;
	CString mp_sf40404a;
	CString mp_sf40404b;
	CString mp_sf40404c;
	CString mp_sf40404d;
	CString mp_sf40404e;
	CString mp_sf40404f;
	CString mp_sf404050;
	CString mp_sf404051;
	CString mp_sf404052;
	CString mp_sf404053;
	CString mp_sf404054;
	CString mp_sf404055;
	CString mp_sf404056;
	CString mp_sf404057;
	CString mp_sf404058;
	CString mp_sf404059;
	CString mp_sf40405a;
	CString mp_sf40405b;
	CString mp_sf40405c;
	CString mp_sf40405d;
	CString mp_sf40405e;
	CString mp_sf40405f;
	CString mp_sf404060;
	CString mp_sf404061;
	CString mp_sf404062;
	CString mp_sf404063;
	CString mp_sf404064;
	CString mp_sf404065;
	CString mp_sf404066;
	CString mp_sf404067;
	CString mp_sf404068;
	CString mp_sf404069;
	CString mp_sf40406a;
	CString mp_sf40406b;
	CString mp_sf40406c;
	CString mp_sf40406d;
	CString mp_sf40406e;
	CString mp_sf40406f;
	CString mp_sf404070;
	CString mp_sf404071;
	CString mp_sf404072;
	CString mp_sf404073;
	CString mp_sf404074;
	CString mp_sf404075;
	CString mp_sf404076;
	CString mp_sf404077;
	CString mp_sf404078;
	CString mp_sf404079;
	CString mp_sf40407a;
	CString mp_sf40407b;
	CString mp_sf40407c;
	CString mp_sf40407d;
	CString mp_sf40407e;
	CString mp_sf40407f;
	CString mp_sf404080;
	CString mp_sf404081;
	CString mp_sf404082;
	CString mp_sf404083;
	CString mp_sf404084;
	CString mp_sf404085;
	CString mp_sf404086;
	CString mp_sf404087;
	CString mp_sf404088;
	CString mp_sf404089;
	CString mp_sf40408a;
	CString mp_sf40408b;
	CString mp_sf40408c;
	CString mp_sf40408d;
	CString mp_sf40408e;
	CString mp_sf40408f;
	CString mp_sf404090;
	CString mp_sf404091;
	CString mp_sf404092;
	CString mp_sf404093;
	CString mp_sf404094;
	CString mp_sf404095;
	CString mp_sf404096;
	CString mp_sf404097;
	CString mp_sf404098;
	CString mp_sf404099;
	CString mp_sf40409a;
	CString mp_sf40409b;
	CString mp_sf40409c;
	CString mp_sf40409d;
	CString mp_sf40409e;
	CString mp_sf40409f;
	CString mp_sf4040a0;
	CString mp_sf4040a1;
	CString mp_sf4040a2;
	CString mp_sf4040a3;
	CString mp_sf4040a4;
	CString mp_sf4040a5;
	CString mp_sf4040a6;
	CString mp_sf4040a7;
	CString mp_sf4040a8;
	CString mp_sf4040a9;
	CString mp_sf4040aa;
	CString mp_sf4040ab;
	CString mp_sf4040ac;
	CString mp_sf4040ad;
	CString mp_sf4040ae;
	CString mp_sf4040af;
	CString mp_sf4040b0;
	CString mp_sf4040b1;
	CString mp_sf4040b2;
	CString mp_sf4040b3;
	CString mp_sf4040b4;
	CString mp_sf4040b5;
	CString mp_sf4040b6;
	CString mp_sf4040b7;
	CString mp_sf4040b8;
	CString mp_sf4040b9;
	CString mp_sf4040ba;
	CString mp_sf4040bb;
	CString mp_sf4040bc;
	CString mp_sf4040bd;
	CString mp_sf4040be;
	CString mp_sf4040bf;
	CString mp_sf4040c0;
	CString mp_sf4040c1;
	CString mp_sf4040c2;
	CString mp_sf4040c3;
	CString mp_sf4040c4;
	CString mp_sf4040c5;
	CString mp_sf4040c6;
	CString mp_sf4040c7;
	CString mp_sf4040c8;
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

	FLOAT odl_GetPeakUnits(void);
	FLOAT GetPeakUnits(void);
	void odl_SetPeakUnits(FLOAT newVal);
	void SetPeakUnits(FLOAT newVal);

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

	BOOL odl_GetSummary(void);
	BOOL GetSummary(void);
	void odl_SetSummary(BOOL newVal);
	void SetSummary(BOOL newVal);

	FLOAT odl_GetSV(void);
	FLOAT GetSV(void);
	void odl_SetSV(FLOAT newVal);
	void SetSV(FLOAT newVal);

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

	BSTR odl_GetAssnOwner(void);
	CString GetAssnOwner(void);
	void odl_SetAssnOwner(LPCTSTR newVal);
	void SetAssnOwner(CString newVal);

	BSTR odl_GetAssnOwnerGuid(void);
	CString GetAssnOwnerGuid(void);
	void odl_SetAssnOwnerGuid(LPCTSTR newVal);
	void SetAssnOwnerGuid(CString newVal);

	DOUBLE odl_GetBudgetCost(void);
	DOUBLE GetBudgetCost(void);
	void odl_SetBudgetCost(DOUBLE newVal);
	void SetBudgetCost(DOUBLE newVal);

	IDispatch* odl_GetBudgetWork(void);
	Duration* GetBudgetWork(void);

	IDispatch* odl_GetExtendedAttribute(void);
	AssignmentExtendedAttribute_C* GetExtendedAttribute(void);

	IDispatch* odl_GetBaseline(void);
	AssignmentBaseline_C* GetBaseline(void);

	BSTR odl_Getf404000(void);
	CString Getf404000(void);
	void odl_Setf404000(LPCTSTR newVal);
	void Setf404000(CString newVal);

	BSTR odl_Getf404001(void);
	CString Getf404001(void);
	void odl_Setf404001(LPCTSTR newVal);
	void Setf404001(CString newVal);

	BSTR odl_Getf404002(void);
	CString Getf404002(void);
	void odl_Setf404002(LPCTSTR newVal);
	void Setf404002(CString newVal);

	BSTR odl_Getf404003(void);
	CString Getf404003(void);
	void odl_Setf404003(LPCTSTR newVal);
	void Setf404003(CString newVal);

	BSTR odl_Getf404004(void);
	CString Getf404004(void);
	void odl_Setf404004(LPCTSTR newVal);
	void Setf404004(CString newVal);

	BSTR odl_Getf404005(void);
	CString Getf404005(void);
	void odl_Setf404005(LPCTSTR newVal);
	void Setf404005(CString newVal);

	BSTR odl_Getf404006(void);
	CString Getf404006(void);
	void odl_Setf404006(LPCTSTR newVal);
	void Setf404006(CString newVal);

	BSTR odl_Getf404007(void);
	CString Getf404007(void);
	void odl_Setf404007(LPCTSTR newVal);
	void Setf404007(CString newVal);

	BSTR odl_Getf404008(void);
	CString Getf404008(void);
	void odl_Setf404008(LPCTSTR newVal);
	void Setf404008(CString newVal);

	BSTR odl_Getf404009(void);
	CString Getf404009(void);
	void odl_Setf404009(LPCTSTR newVal);
	void Setf404009(CString newVal);

	BSTR odl_Getf40400a(void);
	CString Getf40400a(void);
	void odl_Setf40400a(LPCTSTR newVal);
	void Setf40400a(CString newVal);

	BSTR odl_Getf40400b(void);
	CString Getf40400b(void);
	void odl_Setf40400b(LPCTSTR newVal);
	void Setf40400b(CString newVal);

	BSTR odl_Getf40400c(void);
	CString Getf40400c(void);
	void odl_Setf40400c(LPCTSTR newVal);
	void Setf40400c(CString newVal);

	BSTR odl_Getf40400d(void);
	CString Getf40400d(void);
	void odl_Setf40400d(LPCTSTR newVal);
	void Setf40400d(CString newVal);

	BSTR odl_Getf40400e(void);
	CString Getf40400e(void);
	void odl_Setf40400e(LPCTSTR newVal);
	void Setf40400e(CString newVal);

	BSTR odl_Getf40400f(void);
	CString Getf40400f(void);
	void odl_Setf40400f(LPCTSTR newVal);
	void Setf40400f(CString newVal);

	BSTR odl_Getf404010(void);
	CString Getf404010(void);
	void odl_Setf404010(LPCTSTR newVal);
	void Setf404010(CString newVal);

	BSTR odl_Getf404011(void);
	CString Getf404011(void);
	void odl_Setf404011(LPCTSTR newVal);
	void Setf404011(CString newVal);

	BSTR odl_Getf404012(void);
	CString Getf404012(void);
	void odl_Setf404012(LPCTSTR newVal);
	void Setf404012(CString newVal);

	BSTR odl_Getf404013(void);
	CString Getf404013(void);
	void odl_Setf404013(LPCTSTR newVal);
	void Setf404013(CString newVal);

	BSTR odl_Getf404014(void);
	CString Getf404014(void);
	void odl_Setf404014(LPCTSTR newVal);
	void Setf404014(CString newVal);

	BSTR odl_Getf404015(void);
	CString Getf404015(void);
	void odl_Setf404015(LPCTSTR newVal);
	void Setf404015(CString newVal);

	BSTR odl_Getf404016(void);
	CString Getf404016(void);
	void odl_Setf404016(LPCTSTR newVal);
	void Setf404016(CString newVal);

	BSTR odl_Getf404017(void);
	CString Getf404017(void);
	void odl_Setf404017(LPCTSTR newVal);
	void Setf404017(CString newVal);

	BSTR odl_Getf404018(void);
	CString Getf404018(void);
	void odl_Setf404018(LPCTSTR newVal);
	void Setf404018(CString newVal);

	BSTR odl_Getf404019(void);
	CString Getf404019(void);
	void odl_Setf404019(LPCTSTR newVal);
	void Setf404019(CString newVal);

	BSTR odl_Getf40401a(void);
	CString Getf40401a(void);
	void odl_Setf40401a(LPCTSTR newVal);
	void Setf40401a(CString newVal);

	BSTR odl_Getf40401b(void);
	CString Getf40401b(void);
	void odl_Setf40401b(LPCTSTR newVal);
	void Setf40401b(CString newVal);

	BSTR odl_Getf40401c(void);
	CString Getf40401c(void);
	void odl_Setf40401c(LPCTSTR newVal);
	void Setf40401c(CString newVal);

	BSTR odl_Getf40401d(void);
	CString Getf40401d(void);
	void odl_Setf40401d(LPCTSTR newVal);
	void Setf40401d(CString newVal);

	BSTR odl_Getf40401e(void);
	CString Getf40401e(void);
	void odl_Setf40401e(LPCTSTR newVal);
	void Setf40401e(CString newVal);

	BSTR odl_Getf40401f(void);
	CString Getf40401f(void);
	void odl_Setf40401f(LPCTSTR newVal);
	void Setf40401f(CString newVal);

	BSTR odl_Getf404020(void);
	CString Getf404020(void);
	void odl_Setf404020(LPCTSTR newVal);
	void Setf404020(CString newVal);

	BSTR odl_Getf404021(void);
	CString Getf404021(void);
	void odl_Setf404021(LPCTSTR newVal);
	void Setf404021(CString newVal);

	BSTR odl_Getf404022(void);
	CString Getf404022(void);
	void odl_Setf404022(LPCTSTR newVal);
	void Setf404022(CString newVal);

	BSTR odl_Getf404023(void);
	CString Getf404023(void);
	void odl_Setf404023(LPCTSTR newVal);
	void Setf404023(CString newVal);

	BSTR odl_Getf404024(void);
	CString Getf404024(void);
	void odl_Setf404024(LPCTSTR newVal);
	void Setf404024(CString newVal);

	BSTR odl_Getf404025(void);
	CString Getf404025(void);
	void odl_Setf404025(LPCTSTR newVal);
	void Setf404025(CString newVal);

	BSTR odl_Getf404026(void);
	CString Getf404026(void);
	void odl_Setf404026(LPCTSTR newVal);
	void Setf404026(CString newVal);

	BSTR odl_Getf404027(void);
	CString Getf404027(void);
	void odl_Setf404027(LPCTSTR newVal);
	void Setf404027(CString newVal);

	BSTR odl_Getf404028(void);
	CString Getf404028(void);
	void odl_Setf404028(LPCTSTR newVal);
	void Setf404028(CString newVal);

	BSTR odl_Getf404029(void);
	CString Getf404029(void);
	void odl_Setf404029(LPCTSTR newVal);
	void Setf404029(CString newVal);

	BSTR odl_Getf40402a(void);
	CString Getf40402a(void);
	void odl_Setf40402a(LPCTSTR newVal);
	void Setf40402a(CString newVal);

	BSTR odl_Getf40402b(void);
	CString Getf40402b(void);
	void odl_Setf40402b(LPCTSTR newVal);
	void Setf40402b(CString newVal);

	BSTR odl_Getf40402c(void);
	CString Getf40402c(void);
	void odl_Setf40402c(LPCTSTR newVal);
	void Setf40402c(CString newVal);

	BSTR odl_Getf40402d(void);
	CString Getf40402d(void);
	void odl_Setf40402d(LPCTSTR newVal);
	void Setf40402d(CString newVal);

	BSTR odl_Getf40402e(void);
	CString Getf40402e(void);
	void odl_Setf40402e(LPCTSTR newVal);
	void Setf40402e(CString newVal);

	BSTR odl_Getf40402f(void);
	CString Getf40402f(void);
	void odl_Setf40402f(LPCTSTR newVal);
	void Setf40402f(CString newVal);

	BSTR odl_Getf404030(void);
	CString Getf404030(void);
	void odl_Setf404030(LPCTSTR newVal);
	void Setf404030(CString newVal);

	BSTR odl_Getf404031(void);
	CString Getf404031(void);
	void odl_Setf404031(LPCTSTR newVal);
	void Setf404031(CString newVal);

	BSTR odl_Getf404032(void);
	CString Getf404032(void);
	void odl_Setf404032(LPCTSTR newVal);
	void Setf404032(CString newVal);

	BSTR odl_Getf404033(void);
	CString Getf404033(void);
	void odl_Setf404033(LPCTSTR newVal);
	void Setf404033(CString newVal);

	BSTR odl_Getf404034(void);
	CString Getf404034(void);
	void odl_Setf404034(LPCTSTR newVal);
	void Setf404034(CString newVal);

	BSTR odl_Getf404035(void);
	CString Getf404035(void);
	void odl_Setf404035(LPCTSTR newVal);
	void Setf404035(CString newVal);

	BSTR odl_Getf404036(void);
	CString Getf404036(void);
	void odl_Setf404036(LPCTSTR newVal);
	void Setf404036(CString newVal);

	BSTR odl_Getf404037(void);
	CString Getf404037(void);
	void odl_Setf404037(LPCTSTR newVal);
	void Setf404037(CString newVal);

	BSTR odl_Getf404038(void);
	CString Getf404038(void);
	void odl_Setf404038(LPCTSTR newVal);
	void Setf404038(CString newVal);

	BSTR odl_Getf404039(void);
	CString Getf404039(void);
	void odl_Setf404039(LPCTSTR newVal);
	void Setf404039(CString newVal);

	BSTR odl_Getf40403a(void);
	CString Getf40403a(void);
	void odl_Setf40403a(LPCTSTR newVal);
	void Setf40403a(CString newVal);

	BSTR odl_Getf40403b(void);
	CString Getf40403b(void);
	void odl_Setf40403b(LPCTSTR newVal);
	void Setf40403b(CString newVal);

	BSTR odl_Getf40403c(void);
	CString Getf40403c(void);
	void odl_Setf40403c(LPCTSTR newVal);
	void Setf40403c(CString newVal);

	BSTR odl_Getf40403d(void);
	CString Getf40403d(void);
	void odl_Setf40403d(LPCTSTR newVal);
	void Setf40403d(CString newVal);

	BSTR odl_Getf40403e(void);
	CString Getf40403e(void);
	void odl_Setf40403e(LPCTSTR newVal);
	void Setf40403e(CString newVal);

	BSTR odl_Getf40403f(void);
	CString Getf40403f(void);
	void odl_Setf40403f(LPCTSTR newVal);
	void Setf40403f(CString newVal);

	BSTR odl_Getf404040(void);
	CString Getf404040(void);
	void odl_Setf404040(LPCTSTR newVal);
	void Setf404040(CString newVal);

	BSTR odl_Getf404041(void);
	CString Getf404041(void);
	void odl_Setf404041(LPCTSTR newVal);
	void Setf404041(CString newVal);

	BSTR odl_Getf404042(void);
	CString Getf404042(void);
	void odl_Setf404042(LPCTSTR newVal);
	void Setf404042(CString newVal);

	BSTR odl_Getf404043(void);
	CString Getf404043(void);
	void odl_Setf404043(LPCTSTR newVal);
	void Setf404043(CString newVal);

	BSTR odl_Getf404044(void);
	CString Getf404044(void);
	void odl_Setf404044(LPCTSTR newVal);
	void Setf404044(CString newVal);

	BSTR odl_Getf404045(void);
	CString Getf404045(void);
	void odl_Setf404045(LPCTSTR newVal);
	void Setf404045(CString newVal);

	BSTR odl_Getf404046(void);
	CString Getf404046(void);
	void odl_Setf404046(LPCTSTR newVal);
	void Setf404046(CString newVal);

	BSTR odl_Getf404047(void);
	CString Getf404047(void);
	void odl_Setf404047(LPCTSTR newVal);
	void Setf404047(CString newVal);

	BSTR odl_Getf404048(void);
	CString Getf404048(void);
	void odl_Setf404048(LPCTSTR newVal);
	void Setf404048(CString newVal);

	BSTR odl_Getf404049(void);
	CString Getf404049(void);
	void odl_Setf404049(LPCTSTR newVal);
	void Setf404049(CString newVal);

	BSTR odl_Getf40404a(void);
	CString Getf40404a(void);
	void odl_Setf40404a(LPCTSTR newVal);
	void Setf40404a(CString newVal);

	BSTR odl_Getf40404b(void);
	CString Getf40404b(void);
	void odl_Setf40404b(LPCTSTR newVal);
	void Setf40404b(CString newVal);

	BSTR odl_Getf40404c(void);
	CString Getf40404c(void);
	void odl_Setf40404c(LPCTSTR newVal);
	void Setf40404c(CString newVal);

	BSTR odl_Getf40404d(void);
	CString Getf40404d(void);
	void odl_Setf40404d(LPCTSTR newVal);
	void Setf40404d(CString newVal);

	BSTR odl_Getf40404e(void);
	CString Getf40404e(void);
	void odl_Setf40404e(LPCTSTR newVal);
	void Setf40404e(CString newVal);

	BSTR odl_Getf40404f(void);
	CString Getf40404f(void);
	void odl_Setf40404f(LPCTSTR newVal);
	void Setf40404f(CString newVal);

	BSTR odl_Getf404050(void);
	CString Getf404050(void);
	void odl_Setf404050(LPCTSTR newVal);
	void Setf404050(CString newVal);

	BSTR odl_Getf404051(void);
	CString Getf404051(void);
	void odl_Setf404051(LPCTSTR newVal);
	void Setf404051(CString newVal);

	BSTR odl_Getf404052(void);
	CString Getf404052(void);
	void odl_Setf404052(LPCTSTR newVal);
	void Setf404052(CString newVal);

	BSTR odl_Getf404053(void);
	CString Getf404053(void);
	void odl_Setf404053(LPCTSTR newVal);
	void Setf404053(CString newVal);

	BSTR odl_Getf404054(void);
	CString Getf404054(void);
	void odl_Setf404054(LPCTSTR newVal);
	void Setf404054(CString newVal);

	BSTR odl_Getf404055(void);
	CString Getf404055(void);
	void odl_Setf404055(LPCTSTR newVal);
	void Setf404055(CString newVal);

	BSTR odl_Getf404056(void);
	CString Getf404056(void);
	void odl_Setf404056(LPCTSTR newVal);
	void Setf404056(CString newVal);

	BSTR odl_Getf404057(void);
	CString Getf404057(void);
	void odl_Setf404057(LPCTSTR newVal);
	void Setf404057(CString newVal);

	BSTR odl_Getf404058(void);
	CString Getf404058(void);
	void odl_Setf404058(LPCTSTR newVal);
	void Setf404058(CString newVal);

	BSTR odl_Getf404059(void);
	CString Getf404059(void);
	void odl_Setf404059(LPCTSTR newVal);
	void Setf404059(CString newVal);

	BSTR odl_Getf40405a(void);
	CString Getf40405a(void);
	void odl_Setf40405a(LPCTSTR newVal);
	void Setf40405a(CString newVal);

	BSTR odl_Getf40405b(void);
	CString Getf40405b(void);
	void odl_Setf40405b(LPCTSTR newVal);
	void Setf40405b(CString newVal);

	BSTR odl_Getf40405c(void);
	CString Getf40405c(void);
	void odl_Setf40405c(LPCTSTR newVal);
	void Setf40405c(CString newVal);

	BSTR odl_Getf40405d(void);
	CString Getf40405d(void);
	void odl_Setf40405d(LPCTSTR newVal);
	void Setf40405d(CString newVal);

	BSTR odl_Getf40405e(void);
	CString Getf40405e(void);
	void odl_Setf40405e(LPCTSTR newVal);
	void Setf40405e(CString newVal);

	BSTR odl_Getf40405f(void);
	CString Getf40405f(void);
	void odl_Setf40405f(LPCTSTR newVal);
	void Setf40405f(CString newVal);

	BSTR odl_Getf404060(void);
	CString Getf404060(void);
	void odl_Setf404060(LPCTSTR newVal);
	void Setf404060(CString newVal);

	BSTR odl_Getf404061(void);
	CString Getf404061(void);
	void odl_Setf404061(LPCTSTR newVal);
	void Setf404061(CString newVal);

	BSTR odl_Getf404062(void);
	CString Getf404062(void);
	void odl_Setf404062(LPCTSTR newVal);
	void Setf404062(CString newVal);

	BSTR odl_Getf404063(void);
	CString Getf404063(void);
	void odl_Setf404063(LPCTSTR newVal);
	void Setf404063(CString newVal);

	BSTR odl_Getf404064(void);
	CString Getf404064(void);
	void odl_Setf404064(LPCTSTR newVal);
	void Setf404064(CString newVal);

	BSTR odl_Getf404065(void);
	CString Getf404065(void);
	void odl_Setf404065(LPCTSTR newVal);
	void Setf404065(CString newVal);

	BSTR odl_Getf404066(void);
	CString Getf404066(void);
	void odl_Setf404066(LPCTSTR newVal);
	void Setf404066(CString newVal);

	BSTR odl_Getf404067(void);
	CString Getf404067(void);
	void odl_Setf404067(LPCTSTR newVal);
	void Setf404067(CString newVal);

	BSTR odl_Getf404068(void);
	CString Getf404068(void);
	void odl_Setf404068(LPCTSTR newVal);
	void Setf404068(CString newVal);

	BSTR odl_Getf404069(void);
	CString Getf404069(void);
	void odl_Setf404069(LPCTSTR newVal);
	void Setf404069(CString newVal);

	BSTR odl_Getf40406a(void);
	CString Getf40406a(void);
	void odl_Setf40406a(LPCTSTR newVal);
	void Setf40406a(CString newVal);

	BSTR odl_Getf40406b(void);
	CString Getf40406b(void);
	void odl_Setf40406b(LPCTSTR newVal);
	void Setf40406b(CString newVal);

	BSTR odl_Getf40406c(void);
	CString Getf40406c(void);
	void odl_Setf40406c(LPCTSTR newVal);
	void Setf40406c(CString newVal);

	BSTR odl_Getf40406d(void);
	CString Getf40406d(void);
	void odl_Setf40406d(LPCTSTR newVal);
	void Setf40406d(CString newVal);

	BSTR odl_Getf40406e(void);
	CString Getf40406e(void);
	void odl_Setf40406e(LPCTSTR newVal);
	void Setf40406e(CString newVal);

	BSTR odl_Getf40406f(void);
	CString Getf40406f(void);
	void odl_Setf40406f(LPCTSTR newVal);
	void Setf40406f(CString newVal);

	BSTR odl_Getf404070(void);
	CString Getf404070(void);
	void odl_Setf404070(LPCTSTR newVal);
	void Setf404070(CString newVal);

	BSTR odl_Getf404071(void);
	CString Getf404071(void);
	void odl_Setf404071(LPCTSTR newVal);
	void Setf404071(CString newVal);

	BSTR odl_Getf404072(void);
	CString Getf404072(void);
	void odl_Setf404072(LPCTSTR newVal);
	void Setf404072(CString newVal);

	BSTR odl_Getf404073(void);
	CString Getf404073(void);
	void odl_Setf404073(LPCTSTR newVal);
	void Setf404073(CString newVal);

	BSTR odl_Getf404074(void);
	CString Getf404074(void);
	void odl_Setf404074(LPCTSTR newVal);
	void Setf404074(CString newVal);

	BSTR odl_Getf404075(void);
	CString Getf404075(void);
	void odl_Setf404075(LPCTSTR newVal);
	void Setf404075(CString newVal);

	BSTR odl_Getf404076(void);
	CString Getf404076(void);
	void odl_Setf404076(LPCTSTR newVal);
	void Setf404076(CString newVal);

	BSTR odl_Getf404077(void);
	CString Getf404077(void);
	void odl_Setf404077(LPCTSTR newVal);
	void Setf404077(CString newVal);

	BSTR odl_Getf404078(void);
	CString Getf404078(void);
	void odl_Setf404078(LPCTSTR newVal);
	void Setf404078(CString newVal);

	BSTR odl_Getf404079(void);
	CString Getf404079(void);
	void odl_Setf404079(LPCTSTR newVal);
	void Setf404079(CString newVal);

	BSTR odl_Getf40407a(void);
	CString Getf40407a(void);
	void odl_Setf40407a(LPCTSTR newVal);
	void Setf40407a(CString newVal);

	BSTR odl_Getf40407b(void);
	CString Getf40407b(void);
	void odl_Setf40407b(LPCTSTR newVal);
	void Setf40407b(CString newVal);

	BSTR odl_Getf40407c(void);
	CString Getf40407c(void);
	void odl_Setf40407c(LPCTSTR newVal);
	void Setf40407c(CString newVal);

	BSTR odl_Getf40407d(void);
	CString Getf40407d(void);
	void odl_Setf40407d(LPCTSTR newVal);
	void Setf40407d(CString newVal);

	BSTR odl_Getf40407e(void);
	CString Getf40407e(void);
	void odl_Setf40407e(LPCTSTR newVal);
	void Setf40407e(CString newVal);

	BSTR odl_Getf40407f(void);
	CString Getf40407f(void);
	void odl_Setf40407f(LPCTSTR newVal);
	void Setf40407f(CString newVal);

	BSTR odl_Getf404080(void);
	CString Getf404080(void);
	void odl_Setf404080(LPCTSTR newVal);
	void Setf404080(CString newVal);

	BSTR odl_Getf404081(void);
	CString Getf404081(void);
	void odl_Setf404081(LPCTSTR newVal);
	void Setf404081(CString newVal);

	BSTR odl_Getf404082(void);
	CString Getf404082(void);
	void odl_Setf404082(LPCTSTR newVal);
	void Setf404082(CString newVal);

	BSTR odl_Getf404083(void);
	CString Getf404083(void);
	void odl_Setf404083(LPCTSTR newVal);
	void Setf404083(CString newVal);

	BSTR odl_Getf404084(void);
	CString Getf404084(void);
	void odl_Setf404084(LPCTSTR newVal);
	void Setf404084(CString newVal);

	BSTR odl_Getf404085(void);
	CString Getf404085(void);
	void odl_Setf404085(LPCTSTR newVal);
	void Setf404085(CString newVal);

	BSTR odl_Getf404086(void);
	CString Getf404086(void);
	void odl_Setf404086(LPCTSTR newVal);
	void Setf404086(CString newVal);

	BSTR odl_Getf404087(void);
	CString Getf404087(void);
	void odl_Setf404087(LPCTSTR newVal);
	void Setf404087(CString newVal);

	BSTR odl_Getf404088(void);
	CString Getf404088(void);
	void odl_Setf404088(LPCTSTR newVal);
	void Setf404088(CString newVal);

	BSTR odl_Getf404089(void);
	CString Getf404089(void);
	void odl_Setf404089(LPCTSTR newVal);
	void Setf404089(CString newVal);

	BSTR odl_Getf40408a(void);
	CString Getf40408a(void);
	void odl_Setf40408a(LPCTSTR newVal);
	void Setf40408a(CString newVal);

	BSTR odl_Getf40408b(void);
	CString Getf40408b(void);
	void odl_Setf40408b(LPCTSTR newVal);
	void Setf40408b(CString newVal);

	BSTR odl_Getf40408c(void);
	CString Getf40408c(void);
	void odl_Setf40408c(LPCTSTR newVal);
	void Setf40408c(CString newVal);

	BSTR odl_Getf40408d(void);
	CString Getf40408d(void);
	void odl_Setf40408d(LPCTSTR newVal);
	void Setf40408d(CString newVal);

	BSTR odl_Getf40408e(void);
	CString Getf40408e(void);
	void odl_Setf40408e(LPCTSTR newVal);
	void Setf40408e(CString newVal);

	BSTR odl_Getf40408f(void);
	CString Getf40408f(void);
	void odl_Setf40408f(LPCTSTR newVal);
	void Setf40408f(CString newVal);

	BSTR odl_Getf404090(void);
	CString Getf404090(void);
	void odl_Setf404090(LPCTSTR newVal);
	void Setf404090(CString newVal);

	BSTR odl_Getf404091(void);
	CString Getf404091(void);
	void odl_Setf404091(LPCTSTR newVal);
	void Setf404091(CString newVal);

	BSTR odl_Getf404092(void);
	CString Getf404092(void);
	void odl_Setf404092(LPCTSTR newVal);
	void Setf404092(CString newVal);

	BSTR odl_Getf404093(void);
	CString Getf404093(void);
	void odl_Setf404093(LPCTSTR newVal);
	void Setf404093(CString newVal);

	BSTR odl_Getf404094(void);
	CString Getf404094(void);
	void odl_Setf404094(LPCTSTR newVal);
	void Setf404094(CString newVal);

	BSTR odl_Getf404095(void);
	CString Getf404095(void);
	void odl_Setf404095(LPCTSTR newVal);
	void Setf404095(CString newVal);

	BSTR odl_Getf404096(void);
	CString Getf404096(void);
	void odl_Setf404096(LPCTSTR newVal);
	void Setf404096(CString newVal);

	BSTR odl_Getf404097(void);
	CString Getf404097(void);
	void odl_Setf404097(LPCTSTR newVal);
	void Setf404097(CString newVal);

	BSTR odl_Getf404098(void);
	CString Getf404098(void);
	void odl_Setf404098(LPCTSTR newVal);
	void Setf404098(CString newVal);

	BSTR odl_Getf404099(void);
	CString Getf404099(void);
	void odl_Setf404099(LPCTSTR newVal);
	void Setf404099(CString newVal);

	BSTR odl_Getf40409a(void);
	CString Getf40409a(void);
	void odl_Setf40409a(LPCTSTR newVal);
	void Setf40409a(CString newVal);

	BSTR odl_Getf40409b(void);
	CString Getf40409b(void);
	void odl_Setf40409b(LPCTSTR newVal);
	void Setf40409b(CString newVal);

	BSTR odl_Getf40409c(void);
	CString Getf40409c(void);
	void odl_Setf40409c(LPCTSTR newVal);
	void Setf40409c(CString newVal);

	BSTR odl_Getf40409d(void);
	CString Getf40409d(void);
	void odl_Setf40409d(LPCTSTR newVal);
	void Setf40409d(CString newVal);

	BSTR odl_Getf40409e(void);
	CString Getf40409e(void);
	void odl_Setf40409e(LPCTSTR newVal);
	void Setf40409e(CString newVal);

	BSTR odl_Getf40409f(void);
	CString Getf40409f(void);
	void odl_Setf40409f(LPCTSTR newVal);
	void Setf40409f(CString newVal);

	BSTR odl_Getf4040a0(void);
	CString Getf4040a0(void);
	void odl_Setf4040a0(LPCTSTR newVal);
	void Setf4040a0(CString newVal);

	BSTR odl_Getf4040a1(void);
	CString Getf4040a1(void);
	void odl_Setf4040a1(LPCTSTR newVal);
	void Setf4040a1(CString newVal);

	BSTR odl_Getf4040a2(void);
	CString Getf4040a2(void);
	void odl_Setf4040a2(LPCTSTR newVal);
	void Setf4040a2(CString newVal);

	BSTR odl_Getf4040a3(void);
	CString Getf4040a3(void);
	void odl_Setf4040a3(LPCTSTR newVal);
	void Setf4040a3(CString newVal);

	BSTR odl_Getf4040a4(void);
	CString Getf4040a4(void);
	void odl_Setf4040a4(LPCTSTR newVal);
	void Setf4040a4(CString newVal);

	BSTR odl_Getf4040a5(void);
	CString Getf4040a5(void);
	void odl_Setf4040a5(LPCTSTR newVal);
	void Setf4040a5(CString newVal);

	BSTR odl_Getf4040a6(void);
	CString Getf4040a6(void);
	void odl_Setf4040a6(LPCTSTR newVal);
	void Setf4040a6(CString newVal);

	BSTR odl_Getf4040a7(void);
	CString Getf4040a7(void);
	void odl_Setf4040a7(LPCTSTR newVal);
	void Setf4040a7(CString newVal);

	BSTR odl_Getf4040a8(void);
	CString Getf4040a8(void);
	void odl_Setf4040a8(LPCTSTR newVal);
	void Setf4040a8(CString newVal);

	BSTR odl_Getf4040a9(void);
	CString Getf4040a9(void);
	void odl_Setf4040a9(LPCTSTR newVal);
	void Setf4040a9(CString newVal);

	BSTR odl_Getf4040aa(void);
	CString Getf4040aa(void);
	void odl_Setf4040aa(LPCTSTR newVal);
	void Setf4040aa(CString newVal);

	BSTR odl_Getf4040ab(void);
	CString Getf4040ab(void);
	void odl_Setf4040ab(LPCTSTR newVal);
	void Setf4040ab(CString newVal);

	BSTR odl_Getf4040ac(void);
	CString Getf4040ac(void);
	void odl_Setf4040ac(LPCTSTR newVal);
	void Setf4040ac(CString newVal);

	BSTR odl_Getf4040ad(void);
	CString Getf4040ad(void);
	void odl_Setf4040ad(LPCTSTR newVal);
	void Setf4040ad(CString newVal);

	BSTR odl_Getf4040ae(void);
	CString Getf4040ae(void);
	void odl_Setf4040ae(LPCTSTR newVal);
	void Setf4040ae(CString newVal);

	BSTR odl_Getf4040af(void);
	CString Getf4040af(void);
	void odl_Setf4040af(LPCTSTR newVal);
	void Setf4040af(CString newVal);

	BSTR odl_Getf4040b0(void);
	CString Getf4040b0(void);
	void odl_Setf4040b0(LPCTSTR newVal);
	void Setf4040b0(CString newVal);

	BSTR odl_Getf4040b1(void);
	CString Getf4040b1(void);
	void odl_Setf4040b1(LPCTSTR newVal);
	void Setf4040b1(CString newVal);

	BSTR odl_Getf4040b2(void);
	CString Getf4040b2(void);
	void odl_Setf4040b2(LPCTSTR newVal);
	void Setf4040b2(CString newVal);

	BSTR odl_Getf4040b3(void);
	CString Getf4040b3(void);
	void odl_Setf4040b3(LPCTSTR newVal);
	void Setf4040b3(CString newVal);

	BSTR odl_Getf4040b4(void);
	CString Getf4040b4(void);
	void odl_Setf4040b4(LPCTSTR newVal);
	void Setf4040b4(CString newVal);

	BSTR odl_Getf4040b5(void);
	CString Getf4040b5(void);
	void odl_Setf4040b5(LPCTSTR newVal);
	void Setf4040b5(CString newVal);

	BSTR odl_Getf4040b6(void);
	CString Getf4040b6(void);
	void odl_Setf4040b6(LPCTSTR newVal);
	void Setf4040b6(CString newVal);

	BSTR odl_Getf4040b7(void);
	CString Getf4040b7(void);
	void odl_Setf4040b7(LPCTSTR newVal);
	void Setf4040b7(CString newVal);

	BSTR odl_Getf4040b8(void);
	CString Getf4040b8(void);
	void odl_Setf4040b8(LPCTSTR newVal);
	void Setf4040b8(CString newVal);

	BSTR odl_Getf4040b9(void);
	CString Getf4040b9(void);
	void odl_Setf4040b9(LPCTSTR newVal);
	void Setf4040b9(CString newVal);

	BSTR odl_Getf4040ba(void);
	CString Getf4040ba(void);
	void odl_Setf4040ba(LPCTSTR newVal);
	void Setf4040ba(CString newVal);

	BSTR odl_Getf4040bb(void);
	CString Getf4040bb(void);
	void odl_Setf4040bb(LPCTSTR newVal);
	void Setf4040bb(CString newVal);

	BSTR odl_Getf4040bc(void);
	CString Getf4040bc(void);
	void odl_Setf4040bc(LPCTSTR newVal);
	void Setf4040bc(CString newVal);

	BSTR odl_Getf4040bd(void);
	CString Getf4040bd(void);
	void odl_Setf4040bd(LPCTSTR newVal);
	void Setf4040bd(CString newVal);

	BSTR odl_Getf4040be(void);
	CString Getf4040be(void);
	void odl_Setf4040be(LPCTSTR newVal);
	void Setf4040be(CString newVal);

	BSTR odl_Getf4040bf(void);
	CString Getf4040bf(void);
	void odl_Setf4040bf(LPCTSTR newVal);
	void Setf4040bf(CString newVal);

	BSTR odl_Getf4040c0(void);
	CString Getf4040c0(void);
	void odl_Setf4040c0(LPCTSTR newVal);
	void Setf4040c0(CString newVal);

	BSTR odl_Getf4040c1(void);
	CString Getf4040c1(void);
	void odl_Setf4040c1(LPCTSTR newVal);
	void Setf4040c1(CString newVal);

	BSTR odl_Getf4040c2(void);
	CString Getf4040c2(void);
	void odl_Setf4040c2(LPCTSTR newVal);
	void Setf4040c2(CString newVal);

	BSTR odl_Getf4040c3(void);
	CString Getf4040c3(void);
	void odl_Setf4040c3(LPCTSTR newVal);
	void Setf4040c3(CString newVal);

	BSTR odl_Getf4040c4(void);
	CString Getf4040c4(void);
	void odl_Setf4040c4(LPCTSTR newVal);
	void Setf4040c4(CString newVal);

	BSTR odl_Getf4040c5(void);
	CString Getf4040c5(void);
	void odl_Setf4040c5(LPCTSTR newVal);
	void Setf4040c5(CString newVal);

	BSTR odl_Getf4040c6(void);
	CString Getf4040c6(void);
	void odl_Setf4040c6(LPCTSTR newVal);
	void Setf4040c6(CString newVal);

	BSTR odl_Getf4040c7(void);
	CString Getf4040c7(void);
	void odl_Setf4040c7(LPCTSTR newVal);
	void Setf4040c7(CString newVal);

	BSTR odl_Getf4040c8(void);
	CString Getf4040c8(void);
	void odl_Setf4040c8(LPCTSTR newVal);
	void Setf4040c8(CString newVal);

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