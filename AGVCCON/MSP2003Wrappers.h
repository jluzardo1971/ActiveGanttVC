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

namespace MSP2003
{

class CDuration : public COleDispatchDriver
{
// Constructors
public:
	CDuration() {}	
	CDuration(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CDuration(const CDuration& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Properties
public:
	void SetYear(LONG propval);
	LONG GetYear();
	void SetMonth(LONG propval);
	LONG GetMonth();
	void SetDay(LONG propval);
	LONG GetDay();
	void SetHour(LONG propval);
	LONG GetHour();
	void SetMinute(LONG propval);
	LONG GetMinute();
	void SetSecond(LONG propval);
	LONG GetSecond();

// Methods
public:
	BOOL IsNull(void);
	CString ToString(void);
	CString FromString(LPCTSTR sString);
};

class CTime : public COleDispatchDriver
{
// Constructors
public:
	CTime() {}	
	CTime(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CTime(const CTime& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Properties
public:
	void SetHour(LONG propval);
	LONG GetHour();
	void SetMinute(LONG propval);
	LONG GetMinute();
	void SetSecond(LONG propval);
	LONG GetSecond();

// Methods
public:
	BOOL IsNull(void);
	CString ToString(void);
	CString FromString(LPCTSTR sString);
};

typedef enum E_FYSTARTDATE
{
	FYSD_JANUARY = 1,
	FYSD_FEBRUARY = 2,
	FYSD_MARCH = 3,
	FYSD_APRIL = 4,
	FYSD_MAY = 5,
	FYSD_JUNE = 6,
	FYSD_JULY = 7,
	FYSD_AUGUST = 8,
	FYSD_SEPTEMBER = 9,
	FYSD_OCTOBER = 10,
	FYSD_NOVEMBER = 11,
	FYSD_DECEMBER = 12,
}E_FYSTARTDATE;

typedef enum E_CURRENCYSYMBOLPOSITION
{
	CSP_BEFORE = 0,
	CSP_AFTER = 1,
	CSP_BEFORE_WITH_SPACE = 2,
	CSP_AFTER_WITH_SPACE = 3,
}E_CURRENCYSYMBOLPOSITION;

typedef enum E_DEFAULTTASKTYPE
{
	DTT_FIXED_UNITS = 0,
	DTT_FIXED_DURATION = 1,
	DTT_FIXED_WORK = 2,
}E_DEFAULTTASKTYPE;

typedef enum E_DEFAULTFIXEDCOSTACCRUAL
{
	DFCA_START = 1,
	DFCA_PRORATED = 2,
	DFCA_END = 3,
}E_DEFAULTFIXEDCOSTACCRUAL;

typedef enum E_DURATIONFORMAT
{
	DF_M = 3,
	DF_EM = 4,
	DF_H = 5,
	DF_EH = 6,
	DF_D = 7,
	DF_ED = 8,
	DF_W = 9,
	DF_EW = 10,
	DF_MO = 11,
	DF_EMO = 12,
	DF__PERCENT = 19,
	DF_E_PERCENT = 20,
	DF_NULL = 21,
	DF_M_QUESTIONMRK = 35,
	DF_EM_QUESTIONMRK = 36,
	DF_H_QUESTIONMRK = 37,
	DF_EH_QUESTIONMRK = 38,
	DF_D_QUESTIONMRK = 39,
	DF_ED_QUESTIONMRK = 40,
	DF_W_QUESTIONMRK = 41,
	DF_EW_QUESTIONMRK = 42,
	DF_MO_QUESTIONMRK = 43,
	DF_EMO_QUESTIONMRK = 44,
	DF__PERCENT_QUESTIONMRK = 51,
	DF_E_PERCENT_QUESTIONMRK = 52,
	DF_NULL_A = 53,
}E_DURATIONFORMAT;

typedef enum E_WORKFORMAT
{
	WF_M = 1,
	WF_H = 2,
	WF_D = 3,
	WF_W = 4,
	WF_MO = 5,
}E_WORKFORMAT;

typedef enum E_EARNEDVALUEMETHOD
{
	EVM_PERCENT_COMPLETE = 0,
	EVM_PHYSICAL_PERCENT_COMPLETE = 1,
}E_EARNEDVALUEMETHOD;

typedef enum E_WEEKSTARTDAY
{
	WSD_SUNDAY = 0,
	WSD_MONDAY = 1,
	WSD_TUESDAY = 2,
	WSD_WEDNESDAY = 3,
	WSD_THURSDAY = 4,
	WSD_FRIDAY = 5,
	WSD_SATURDAY = 6,
}E_WEEKSTARTDAY;

typedef enum E_BASELINEFOREARNEDVALUE
{
	BFEV_BASELINE = 0,
	BFEV_BASELINE_1 = 1,
	BFEV_BASELINE_2 = 2,
	BFEV_BASELINE_3 = 3,
	BFEV_BASELINE_4 = 4,
	BFEV_BASELINE_5 = 5,
	BFEV_BASELINE_6 = 6,
	BFEV_BASELINE_7 = 7,
	BFEV_BASELINE_8 = 8,
	BFEV_BASELINE_9 = 9,
	BFEV_BASELINE_10 = 10,
}E_BASELINEFOREARNEDVALUE;

typedef enum E_NEWTASKSTARTDATE
{
	NTSD_PROJECT_START_DATE = 0,
	NTSD_CURRENT_DATE = 1,
}E_NEWTASKSTARTDATE;

typedef enum E_DEFAULTTASKEVMETHOD
{
	DTEVM_PERCENT_COMPLETE = 0,
	DTEVM_PHYSICAL_PERCENT_COMPLETE = 1,
}E_DEFAULTTASKEVMETHOD;

typedef enum E_TYPE
{
	T_NUMBERS = 0,
	T_UPPER_CASE_LETTERS = 1,
	T_LOWER_CASE_LETTERS = 2,
	T_CHARACTERS = 3,
}E_TYPE;

typedef enum E_TYPE_1
{
	T_1_NUMBERS = 0,
	T_1_UPPERCASE_LETTERS = 1,
	T_1_LOWERCASE_LETTERS = 2,
	T_1_CHARACTERS = 3,
}E_TYPE_1;

typedef enum E_ROLLUPTYPE
{
	RT_MAXIMUM_OR_FOR_FLAG_FIELDS = 0,
	RT_MINIMUM_AND_FOR_FLAG_FIELDS = 1,
	RT_COUNT_ALL = 2,
	RT_SUM = 3,
	RT_AVERAGE = 4,
	RT_AVERAGE_FIRST_SUBLEVEL = 5,
	RT_COUNT_FIRST_SUBLEVEL = 6,
	RT_COUNT_NONSUMMARIES = 7,
}E_ROLLUPTYPE;

typedef enum E_CALCULATIONTYPE
{
	CT_NONE = 0,
	CT_ROLLUP = 1,
	CT_CALCULATION = 2,
}E_CALCULATIONTYPE;

typedef enum E_VALUELISTSORTORDER
{
	VSO_DESCENDING = 0,
	VSO_ASCENDING = 1,
}E_VALUELISTSORTORDER;

typedef enum E_DAYTYPE
{
	DT_EXCEPTION = 0,
	DT_MONDAY = 1,
	DT_TUESDAY = 2,
	DT_WEDNESDAY = 3,
	DT_THURSDAY = 4,
	DT_FRIDAY = 5,
	DT_SATURDAY = 6,
	DT_SUNDAY = 7,
}E_DAYTYPE;

typedef enum E_TYPE_2
{
	T_2_FIXED_UNITS = 0,
	T_2_FIXED_DURATION = 1,
	T_2_FIXED_WORK = 2,
}E_TYPE_2;

typedef enum E_FIXEDCOSTACCRUAL
{
	FCA_START = 1,
	FCA_PRORATED = 2,
	FCA_END = 3,
}E_FIXEDCOSTACCRUAL;

typedef enum E_CONSTRAINTTYPE
{
	CT_AS_SOON_AS_POSSIBLE = 0,
	CT_AS_LATE_AS_POSSIBLE = 1,
	CT_MUST_START_ON = 2,
	CT_MUST_FINISH_ON = 3,
	CT_START_NO_EARLIER_THAN = 4,
	CT_START_NO_LATER_THAN = 5,
	CT_FINISH_NO_EARLIER_THAN = 6,
	CT_FINISH_NO_LATER_THAN = 7,
}E_CONSTRAINTTYPE;

typedef enum E_LEVELINGDELAYFORMAT
{
	LDF_M = 3,
	LDF_EM = 4,
	LDF_H = 5,
	LDF_EH = 6,
	LDF_D = 7,
	LDF_ED = 8,
	LDF_W = 9,
	LDF_EW = 10,
	LDF_MO = 11,
	LDF_EMO = 12,
	LDF__PERCENT = 19,
	LDF_E_PERCENT = 20,
	LDF_NULL = 21,
	LDF_M_QUESTIONMRK = 35,
	LDF_EM_QUESTIONMRK = 36,
	LDF_H_QUESTIONMRK = 37,
	LDF_EH_QUESTIONMRK = 38,
	LDF_D_QUESTIONMRK = 39,
	LDF_ED_QUESTIONMRK = 40,
	LDF_W_QUESTIONMRK = 41,
	LDF_EW_QUESTIONMRK = 42,
	LDF_MO_QUESTIONMRK = 43,
	LDF_EMO_QUESTIONMRK = 44,
	LDF__PERCENT_QUESTIONMRK = 51,
	LDF_E_PERCENT_QUESTIONMRK = 52,
	LDF_NULL_A = 53,
}E_LEVELINGDELAYFORMAT;

typedef enum E_TYPE_3
{
	T_3_FF = 0,
	T_3_FS = 1,
	T_3_SF = 2,
	T_3_SS = 3,
}E_TYPE_3;

typedef enum E_LAGFORMAT
{
	LF_M = 3,
	LF_EM = 4,
	LF_H = 5,
	LF_EH = 6,
	LF_D = 7,
	LF_ED = 8,
	LF_W = 9,
	LF_EW = 10,
	LF_MO = 11,
	LF_EMO = 12,
	LF__PERCENT = 19,
	LF_E_PERCENT = 20,
	LF_M_QUESTIONMRK = 35,
	LF_EM_QUESTIONMRK = 36,
	LF_H_QUESTIONMRK = 37,
	LF_EH_QUESTIONMRK = 38,
	LF_D_QUESTIONMRK = 39,
	LF_ED_QUESTIONMRK = 40,
	LF_W_QUESTIONMRK = 41,
	LF_EW_QUESTIONMRK = 42,
	LF_MO_QUESTIONMRK = 43,
	LF_EMO_QUESTIONMRK = 44,
	LF__PERCENT_QUESTIONMRK = 51,
	LF_E_PERCENT_QUESTIONMRK = 52,
	LF_NULL_A = 53,
}E_LAGFORMAT;

typedef enum E_TYPE_4
{
	T_4_MATERIAL = 0,
	T_4_WORK = 1,
}E_TYPE_4;

typedef enum E_WORKGROUP
{
	WG_DEFAULT = 0,
	WG_NONE = 1,
	WG_EMAIL = 2,
	WG_WEB = 3,
}E_WORKGROUP;

typedef enum E_ACCRUEAT
{
	AA_START = 1,
	AA_END = 2,
	AA_PRORATED = 3,
}E_ACCRUEAT;

typedef enum E_STANDARDRATEFORMAT
{
	SRF_M = 1,
	SRF_H = 2,
	SRF_D = 3,
	SRF_W = 4,
	SRF_MO = 5,
	SRF_Y = 7,
	SRF_MATERIAL_RESOURCE_RATE_OR_BLANK_SYMBOL_SPECIFIED = 8,
}E_STANDARDRATEFORMAT;

typedef enum E_OVERTIMERATEFORMAT
{
	ORF_M = 1,
	ORF_H = 2,
	ORF_D = 3,
	ORF_W = 4,
	ORF_MO = 5,
	ORF_Y = 7,
}E_OVERTIMERATEFORMAT;

typedef enum E_BOOKINGTYPE
{
	BT_COMMITED = 1,
	BT_PROPOSED = 2,
}E_BOOKINGTYPE;

typedef enum E_RATETABLE
{
	RT_A = 0,
	RT_B = 1,
	RT_C = 2,
	RT_D = 3,
	RT_E = 4,
}E_RATETABLE;

typedef enum E_STANDARDRATEFORMAT_1
{
	SRF_1_M = 1,
	SRF_1_H = 2,
	SRF_1_D = 3,
	SRF_1_W = 4,
	SRF_1_MO = 5,
	SRF_1_Y = 7,
}E_STANDARDRATEFORMAT_1;

typedef enum E_COSTRATETABLE
{
	CRT_COST_RATE_TABLE_0 = 0,
	CRT_COST_RATE_TABLE_1 = 1,
	CRT_COST_RATE_TABLE_2 = 2,
	CRT_COST_RATE_TABLE_3 = 3,
	CRT_COST_RATE_TABLE_4 = 4,
}E_COSTRATETABLE;

typedef enum E_WORKCONTOUR
{
	WC_FLAT = 0,
	WC_BACK_LOADED = 1,
	WC_FRONT_LOADED = 2,
	WC_DOUBLE_PEAK = 3,
	WC_EARLY_PEAK = 4,
	WC_LATE_PEAK = 5,
	WC_BELL = 6,
	WC_TURTLE = 7,
	WC_CONTOURED = 8,
}E_WORKCONTOUR;

typedef enum E_TYPE_5
{
	T_5_ASSIGNMENT_REMAINING_WORK = 1,
	T_5_ASSIGNMENT_ACTUAL_WORK = 2,
	T_5_ASSIGNMENT_ACTUAL_OVERTIME_WORK = 3,
	T_5_ASSIGNMENT_BASELINE_WORK = 4,
	T_5_ASSIGNMENT_BASELINE_COST = 5,
	T_5_ASSIGNMENT_ACTUAL_COST = 6,
	T_5_RESOURCE_BASELINE_WORK = 7,
	T_5_RESOURCE_BASELINE_COST = 8,
	T_5_TASK_BASELINE_WORK = 9,
	T_5_TASK_BASELINE_COST = 10,
	T_5_TASK_PERCENT_COMPLETE = 11,
	T_5_ASSIGNMENT_BASELINE_1_WORK = 16,
	T_5_ASSIGNMENT_BASELINE_1_COST = 17,
	T_5_TASK_BASELINE_1_WORK = 18,
	T_5_TASK_BASELINE_1_COST = 19,
	T_5_RESOURCE_BASELINE_1_WORK = 20,
	T_5_RESOURCE_BASELINE_1_COST = 21,
	T_5_ASSIGNMENT_BASELINE_2_WORK = 22,
	T_5_ASSIGNMENT_BASELINE_2_COST = 23,
	T_5_TASK_BASELINE_2_WORK = 24,
	T_5_TASK_BASELINE_2_COST = 25,
	T_5_RESOURCE_BASELINE_2_WORK = 26,
	T_5_RESOURCE_BASELINE_2_COST = 27,
	T_5_ASSIGNMENT_BASELINE_3_WORK = 28,
	T_5_ASSIGNMENT_BASELINE_3_COST = 29,
	T_5_TASK_BASELINE_3_WORK = 30,
	T_5_TASK_BASELINE_3_COST = 31,
	T_5_RESOURCE_BASELINE_3_WORK = 32,
	T_5_RESOURCE_BASELINE_3_COST = 33,
	T_5_ASSIGNMENT_BASELINE_4_WORK = 34,
	T_5_ASSIGNMENT_BASELINE_4_COST = 35,
	T_5_TASK_BASELINE_4_WORK = 36,
	T_5_TASK_BASELINE_4_COST = 37,
	T_5_RESOURCE_BASELINE_4_WORK = 38,
	T_5_RESOURCE_BASELINE_4_COST = 39,
	T_5_ASSIGNMENT_BASELINE_5_WORK = 40,
	T_5_ASSIGNMENT_BASELINE_5_COST = 41,
	T_5_TASK_BASELINE_5_WORK = 42,
	T_5_TASK_BASELINE_5_COST = 43,
	T_5_RESOURCE_BASELINE_5_WORK = 44,
	T_5_RESOURCE_BASELINE_5_COST = 45,
	T_5_ASSIGNMENT_BASELINE_6_WORK = 46,
	T_5_ASSIGNMENT_BASELINE_6_COST = 47,
	T_5_TASK_BASELINE_6_WORK = 48,
	T_5_TASK_BASELINE_6_COST = 49,
	T_5_RESOURCE_BASELINE_6_WORK = 50,
	T_5_RESOURCE_BASELINE_6_COST = 51,
	T_5_ASSIGNMENT_BASELINE_7_WORK = 52,
	T_5_ASSIGNMENT_BASELINE_7_COST = 53,
	T_5_TASK_BASELINE_7_WORK = 54,
	T_5_TASK_BASELINE_7_COST = 55,
	T_5_RESOURCE_BASELINE_7_WORK = 56,
	T_5_RESOURCE_BASELINE_7_COST = 57,
	T_5_ASSIGNMENT_BASELINE_8_WORK = 58,
	T_5_ASSIGNMENT_BASELINE_8_COST = 59,
	T_5_TASK_BASELINE_8_WORK = 60,
	T_5_TASK_BASELINE_8_COST = 61,
	T_5_RESOURCE_BASELINE_8_WORK = 62,
	T_5_RESOURCE_BASELINE_8_COST = 63,
	T_5_ASSIGNMENT_BASELINE_9_WORK = 64,
	T_5_ASSIGNMENT_BASELINE_9_COST = 65,
	T_5_TASK_BASELINE_9_WORK = 66,
	T_5_TASK_BASELINE_9_COST = 67,
	T_5_RESOURCE_BASELINE_9_WORK = 68,
	T_5_RESOURCE_BASELINE_9_COST = 69,
	T_5_ASSIGNMENT_BASELINE_10_WORK = 70,
	T_5_ASSIGNMENT_BASELINE_10_COST = 71,
	T_5_TASK_BASELINE_10_WORK = 72,
	T_5_TASK_BASELINE_10_COST = 73,
	T_5_RESOURCE_BASELINE_10_WORK = 74,
	T_5_RESOURCE_BASELINE_10_COST = 75,
	T_5_PHYSICAL_PERCENT_COMPLETE = 76,
}E_TYPE_5;

typedef enum E_UNIT
{
	U_M = 0,
	U_H = 1,
	U_D = 2,
	U_W = 3,
	U_MO = 5,
	U_Y = 8,
}E_UNIT;

class CTimephasedData : public COleDispatchDriver
{
// Constructors
public:
	CTimephasedData() {}	
	CTimephasedData(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CTimephasedData(const CTimephasedData& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Properties
public:
	void SetyType(LONG propval);
	LONG GetyType();
	void SetlUID(LONG propval);
	LONG GetlUID();
	void SetdtStart(DATE propval);
	DATE GetdtStart();
	void SetdtFinish(DATE propval);
	DATE GetdtFinish();
	void SetyUnit(LONG propval);
	LONG GetyUnit();
	void SetsValue(LPCTSTR propval);
	CString GetsValue();
	void SetKey(LPCTSTR propval);
	CString GetKey();

// Methods
public:
	CString GetXML(void);
	void SetXML(LPCTSTR sXML);
	BOOL IsNull(void);
	void Initialize(void);
};

class CTimephasedData_C : public COleDispatchDriver
{
// Constructors
public:
	CTimephasedData_C() {}	
	CTimephasedData_C(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CTimephasedData_C(const CTimephasedData_C& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Properties
public:
	LONG GetCount();

// Methods
public:
	CTimephasedData Item(LPCTSTR Index);
	CTimephasedData Add(void);
	void Clear(void);
	void Remove(LPCTSTR Index);
	BOOL IsNull(void);
	void Initialize(void);
};

class CAssignmentBaseline : public COleDispatchDriver
{
// Constructors
public:
	CAssignmentBaseline() {}	
	CAssignmentBaseline(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CAssignmentBaseline(const CAssignmentBaseline& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Properties
public:
	CTimephasedData_C GetoTimephasedData_C();
	void SetsNumber(LPCTSTR propval);
	CString GetsNumber();
	void SetsStart(LPCTSTR propval);
	CString GetsStart();
	void SetsFinish(LPCTSTR propval);
	CString GetsFinish();
	CDuration GetoWork();
	void SetsCost(LPCTSTR propval);
	CString GetsCost();
	void SetfBCWS(FLOAT propval);
	FLOAT GetfBCWS();
	void SetfBCWP(FLOAT propval);
	FLOAT GetfBCWP();
	void SetKey(LPCTSTR propval);
	CString GetKey();

// Methods
public:
	CString GetXML(void);
	void SetXML(LPCTSTR sXML);
	BOOL IsNull(void);
	void Initialize(void);
};

class CAssignmentBaseline_C : public COleDispatchDriver
{
// Constructors
public:
	CAssignmentBaseline_C() {}	
	CAssignmentBaseline_C(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CAssignmentBaseline_C(const CAssignmentBaseline_C& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Properties
public:
	LONG GetCount();

// Methods
public:
	CAssignmentBaseline Item(LPCTSTR Index);
	CAssignmentBaseline Add(void);
	void Clear(void);
	void Remove(LPCTSTR Index);
	BOOL IsNull(void);
	void Initialize(void);
};

class CAssignmentExtendedAttribute : public COleDispatchDriver
{
// Constructors
public:
	CAssignmentExtendedAttribute() {}	
	CAssignmentExtendedAttribute(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CAssignmentExtendedAttribute(const CAssignmentExtendedAttribute& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Properties
public:
	void SetlUID(LONG propval);
	LONG GetlUID();
	void SetsFieldID(LPCTSTR propval);
	CString GetsFieldID();
	void SetsValue(LPCTSTR propval);
	CString GetsValue();
	void SetlValueID(LONG propval);
	LONG GetlValueID();
	void SetyDurationFormat(LONG propval);
	LONG GetyDurationFormat();
	void SetKey(LPCTSTR propval);
	CString GetKey();

// Methods
public:
	CString GetXML(void);
	void SetXML(LPCTSTR sXML);
	BOOL IsNull(void);
	void Initialize(void);
};

class CAssignmentExtendedAttribute_C : public COleDispatchDriver
{
// Constructors
public:
	CAssignmentExtendedAttribute_C() {}	
	CAssignmentExtendedAttribute_C(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CAssignmentExtendedAttribute_C(const CAssignmentExtendedAttribute_C& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Properties
public:
	LONG GetCount();

// Methods
public:
	CAssignmentExtendedAttribute Item(LPCTSTR Index);
	CAssignmentExtendedAttribute Add(void);
	void Clear(void);
	void Remove(LPCTSTR Index);
	BOOL IsNull(void);
	void Initialize(void);
};

class CAssignment : public COleDispatchDriver
{
// Constructors
public:
	CAssignment() {}	
	CAssignment(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CAssignment(const CAssignment& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Properties
public:
	void SetlUID(LONG propval);
	LONG GetlUID();
	void SetlTaskUID(LONG propval);
	LONG GetlTaskUID();
	void SetlResourceUID(LONG propval);
	LONG GetlResourceUID();
	void SetlPercentWorkComplete(LONG propval);
	LONG GetlPercentWorkComplete();
	void SetcActualCost(DOUBLE propval);
	DOUBLE GetcActualCost();
	void SetdtActualFinish(DATE propval);
	DATE GetdtActualFinish();
	void SetcActualOvertimeCost(DOUBLE propval);
	DOUBLE GetcActualOvertimeCost();
	CDuration GetoActualOvertimeWork();
	void SetdtActualStart(DATE propval);
	DATE GetdtActualStart();
	CDuration GetoActualWork();
	void SetfACWP(FLOAT propval);
	FLOAT GetfACWP();
	void SetbConfirmed(BOOL propval);
	BOOL GetbConfirmed();
	void SetcCost(DOUBLE propval);
	DOUBLE GetcCost();
	void SetyCostRateTable(LONG propval);
	LONG GetyCostRateTable();
	void SetfCostVariance(FLOAT propval);
	FLOAT GetfCostVariance();
	void SetfCV(FLOAT propval);
	FLOAT GetfCV();
	void SetlDelay(LONG propval);
	LONG GetlDelay();
	void SetdtFinish(DATE propval);
	DATE GetdtFinish();
	void SetlFinishVariance(LONG propval);
	LONG GetlFinishVariance();
	void SetsHyperlink(LPCTSTR propval);
	CString GetsHyperlink();
	void SetsHyperlinkAddress(LPCTSTR propval);
	CString GetsHyperlinkAddress();
	void SetsHyperlinkSubAddress(LPCTSTR propval);
	CString GetsHyperlinkSubAddress();
	void SetfWorkVariance(FLOAT propval);
	FLOAT GetfWorkVariance();
	void SetbHasFixedRateUnits(BOOL propval);
	BOOL GetbHasFixedRateUnits();
	void SetbFixedMaterial(BOOL propval);
	BOOL GetbFixedMaterial();
	void SetlLevelingDelay(LONG propval);
	LONG GetlLevelingDelay();
	void SetyLevelingDelayFormat(LONG propval);
	LONG GetyLevelingDelayFormat();
	void SetbLinkedFields(BOOL propval);
	BOOL GetbLinkedFields();
	void SetbMilestone(BOOL propval);
	BOOL GetbMilestone();
	void SetsNotes(LPCTSTR propval);
	CString GetsNotes();
	void SetbOverallocated(BOOL propval);
	BOOL GetbOverallocated();
	void SetcOvertimeCost(DOUBLE propval);
	DOUBLE GetcOvertimeCost();
	CDuration GetoOvertimeWork();
	CDuration GetoRegularWork();
	void SetcRemainingCost(DOUBLE propval);
	DOUBLE GetcRemainingCost();
	void SetcRemainingOvertimeCost(DOUBLE propval);
	DOUBLE GetcRemainingOvertimeCost();
	CDuration GetoRemainingOvertimeWork();
	CDuration GetoRemainingWork();
	void SetbResponsePending(BOOL propval);
	BOOL GetbResponsePending();
	void SetdtStart(DATE propval);
	DATE GetdtStart();
	void SetdtStop(DATE propval);
	DATE GetdtStop();
	void SetdtResume(DATE propval);
	DATE GetdtResume();
	void SetlStartVariance(LONG propval);
	LONG GetlStartVariance();
	void SetfUnits(FLOAT propval);
	FLOAT GetfUnits();
	void SetbUpdateNeeded(BOOL propval);
	BOOL GetbUpdateNeeded();
	void SetfVAC(FLOAT propval);
	FLOAT GetfVAC();
	CDuration GetoWork();
	void SetyWorkContour(LONG propval);
	LONG GetyWorkContour();
	void SetfBCWS(FLOAT propval);
	FLOAT GetfBCWS();
	void SetfBCWP(FLOAT propval);
	FLOAT GetfBCWP();
	void SetyBookingType(LONG propval);
	LONG GetyBookingType();
	CDuration GetoActualWorkProtected();
	CDuration GetoActualOvertimeWorkProtected();
	void SetdtCreationDate(DATE propval);
	DATE GetdtCreationDate();
	CAssignmentExtendedAttribute_C GetoExtendedAttribute_C();
	CAssignmentBaseline_C GetoBaseline_C();
	CTimephasedData_C GetoTimephasedData_C();
	void SetKey(LPCTSTR propval);
	CString GetKey();

// Methods
public:
	CString GetXML(void);
	void SetXML(LPCTSTR sXML);
	BOOL IsNull(void);
	void Initialize(void);
};

class CAssignments : public COleDispatchDriver
{
// Constructors
public:
	CAssignments() {}	
	CAssignments(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CAssignments(const CAssignments& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Properties
public:
	LONG GetCount();

// Methods
public:
	CAssignment Item(LPCTSTR Index);
	CAssignment Add(void);
	void Clear(void);
	void Remove(LPCTSTR Index);
	BOOL IsNull(void);
	void Initialize(void);
	void SetXML(LPCTSTR sXML);
	CString GetXML(void);
};

class CResourceRate : public COleDispatchDriver
{
// Constructors
public:
	CResourceRate() {}	
	CResourceRate(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CResourceRate(const CResourceRate& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Properties
public:
	void SetdtRatesFrom(DATE propval);
	DATE GetdtRatesFrom();
	void SetdtRatesTo(DATE propval);
	DATE GetdtRatesTo();
	void SetyRateTable(LONG propval);
	LONG GetyRateTable();
	void SetcStandardRate(DOUBLE propval);
	DOUBLE GetcStandardRate();
	void SetyStandardRateFormat(LONG propval);
	LONG GetyStandardRateFormat();
	void SetcOvertimeRate(DOUBLE propval);
	DOUBLE GetcOvertimeRate();
	void SetyOvertimeRateFormat(LONG propval);
	LONG GetyOvertimeRateFormat();
	void SetcCostPerUse(DOUBLE propval);
	DOUBLE GetcCostPerUse();
	void SetKey(LPCTSTR propval);
	CString GetKey();

// Methods
public:
	CString GetXML(void);
	void SetXML(LPCTSTR sXML);
	BOOL IsNull(void);
	void Initialize(void);
};

class CResourceRates : public COleDispatchDriver
{
// Constructors
public:
	CResourceRates() {}	
	CResourceRates(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CResourceRates(const CResourceRates& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Properties
public:
	LONG GetCount();

// Methods
public:
	CResourceRate Item(LPCTSTR Index);
	CResourceRate Add(void);
	void Clear(void);
	void Remove(LPCTSTR Index);
	BOOL IsNull(void);
	void Initialize(void);
	void SetXML(LPCTSTR sXML);
	CString GetXML(void);
};

class CResourceAvailabilityPeriod : public COleDispatchDriver
{
// Constructors
public:
	CResourceAvailabilityPeriod() {}	
	CResourceAvailabilityPeriod(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CResourceAvailabilityPeriod(const CResourceAvailabilityPeriod& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Properties
public:
	void SetdtAvailableFrom(DATE propval);
	DATE GetdtAvailableFrom();
	void SetdtAvailableTo(DATE propval);
	DATE GetdtAvailableTo();
	void SetfAvailableUnits(FLOAT propval);
	FLOAT GetfAvailableUnits();
	void SetKey(LPCTSTR propval);
	CString GetKey();

// Methods
public:
	CString GetXML(void);
	void SetXML(LPCTSTR sXML);
	BOOL IsNull(void);
	void Initialize(void);
};

class CResourceAvailabilityPeriods : public COleDispatchDriver
{
// Constructors
public:
	CResourceAvailabilityPeriods() {}	
	CResourceAvailabilityPeriods(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CResourceAvailabilityPeriods(const CResourceAvailabilityPeriods& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Properties
public:
	LONG GetCount();

// Methods
public:
	CResourceAvailabilityPeriod Item(LPCTSTR Index);
	CResourceAvailabilityPeriod Add(void);
	void Clear(void);
	void Remove(LPCTSTR Index);
	BOOL IsNull(void);
	void Initialize(void);
	void SetXML(LPCTSTR sXML);
	CString GetXML(void);
};

class CResourceOutlineCode : public COleDispatchDriver
{
// Constructors
public:
	CResourceOutlineCode() {}	
	CResourceOutlineCode(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CResourceOutlineCode(const CResourceOutlineCode& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Properties
public:
	void SetlUID(LONG propval);
	LONG GetlUID();
	void SetsFieldID(LPCTSTR propval);
	CString GetsFieldID();
	void SetlValueID(LONG propval);
	LONG GetlValueID();
	void SetKey(LPCTSTR propval);
	CString GetKey();

// Methods
public:
	CString GetXML(void);
	void SetXML(LPCTSTR sXML);
	BOOL IsNull(void);
	void Initialize(void);
};

class CResourceOutlineCode_C : public COleDispatchDriver
{
// Constructors
public:
	CResourceOutlineCode_C() {}	
	CResourceOutlineCode_C(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CResourceOutlineCode_C(const CResourceOutlineCode_C& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Properties
public:
	LONG GetCount();

// Methods
public:
	CResourceOutlineCode Item(LPCTSTR Index);
	CResourceOutlineCode Add(void);
	void Clear(void);
	void Remove(LPCTSTR Index);
	BOOL IsNull(void);
	void Initialize(void);
};

class CResourceBaseline : public COleDispatchDriver
{
// Constructors
public:
	CResourceBaseline() {}	
	CResourceBaseline(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CResourceBaseline(const CResourceBaseline& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Properties
public:
	void SetlNumber(LONG propval);
	LONG GetlNumber();
	CDuration GetoWork();
	void SetfCost(FLOAT propval);
	FLOAT GetfCost();
	void SetfBCWS(FLOAT propval);
	FLOAT GetfBCWS();
	void SetfBCWP(FLOAT propval);
	FLOAT GetfBCWP();
	void SetKey(LPCTSTR propval);
	CString GetKey();

// Methods
public:
	CString GetXML(void);
	void SetXML(LPCTSTR sXML);
	BOOL IsNull(void);
	void Initialize(void);
};

class CResourceBaseline_C : public COleDispatchDriver
{
// Constructors
public:
	CResourceBaseline_C() {}	
	CResourceBaseline_C(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CResourceBaseline_C(const CResourceBaseline_C& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Properties
public:
	LONG GetCount();

// Methods
public:
	CResourceBaseline Item(LPCTSTR Index);
	CResourceBaseline Add(void);
	void Clear(void);
	void Remove(LPCTSTR Index);
	BOOL IsNull(void);
	void Initialize(void);
};

class CResourceExtendedAttribute : public COleDispatchDriver
{
// Constructors
public:
	CResourceExtendedAttribute() {}	
	CResourceExtendedAttribute(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CResourceExtendedAttribute(const CResourceExtendedAttribute& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Properties
public:
	void SetlUID(LONG propval);
	LONG GetlUID();
	void SetsFieldID(LPCTSTR propval);
	CString GetsFieldID();
	void SetsValue(LPCTSTR propval);
	CString GetsValue();
	void SetlValueID(LONG propval);
	LONG GetlValueID();
	void SetyDurationFormat(LONG propval);
	LONG GetyDurationFormat();
	void SetKey(LPCTSTR propval);
	CString GetKey();

// Methods
public:
	CString GetXML(void);
	void SetXML(LPCTSTR sXML);
	BOOL IsNull(void);
	void Initialize(void);
};

class CResourceExtendedAttribute_C : public COleDispatchDriver
{
// Constructors
public:
	CResourceExtendedAttribute_C() {}	
	CResourceExtendedAttribute_C(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CResourceExtendedAttribute_C(const CResourceExtendedAttribute_C& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Properties
public:
	LONG GetCount();

// Methods
public:
	CResourceExtendedAttribute Item(LPCTSTR Index);
	CResourceExtendedAttribute Add(void);
	void Clear(void);
	void Remove(LPCTSTR Index);
	BOOL IsNull(void);
	void Initialize(void);
};

class CResource : public COleDispatchDriver
{
// Constructors
public:
	CResource() {}	
	CResource(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CResource(const CResource& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Properties
public:
	void SetlUID(LONG propval);
	LONG GetlUID();
	void SetlID(LONG propval);
	LONG GetlID();
	void SetsName(LPCTSTR propval);
	CString GetsName();
	void SetyType(LONG propval);
	LONG GetyType();
	void SetbIsNull(BOOL propval);
	BOOL GetbIsNull();
	void SetsInitials(LPCTSTR propval);
	CString GetsInitials();
	void SetsPhonetics(LPCTSTR propval);
	CString GetsPhonetics();
	void SetsNTAccount(LPCTSTR propval);
	CString GetsNTAccount();
	void SetsMaterialLabel(LPCTSTR propval);
	CString GetsMaterialLabel();
	void SetsCode(LPCTSTR propval);
	CString GetsCode();
	void SetsGroup(LPCTSTR propval);
	CString GetsGroup();
	void SetyWorkGroup(LONG propval);
	LONG GetyWorkGroup();
	void SetsEmailAddress(LPCTSTR propval);
	CString GetsEmailAddress();
	void SetsHyperlink(LPCTSTR propval);
	CString GetsHyperlink();
	void SetsHyperlinkAddress(LPCTSTR propval);
	CString GetsHyperlinkAddress();
	void SetsHyperlinkSubAddress(LPCTSTR propval);
	CString GetsHyperlinkSubAddress();
	void SetfMaxUnits(FLOAT propval);
	FLOAT GetfMaxUnits();
	void SetfPeakUnits(FLOAT propval);
	FLOAT GetfPeakUnits();
	void SetbOverAllocated(BOOL propval);
	BOOL GetbOverAllocated();
	void SetdtAvailableFrom(DATE propval);
	DATE GetdtAvailableFrom();
	void SetdtAvailableTo(DATE propval);
	DATE GetdtAvailableTo();
	void SetdtStart(DATE propval);
	DATE GetdtStart();
	void SetdtFinish(DATE propval);
	DATE GetdtFinish();
	void SetbCanLevel(BOOL propval);
	BOOL GetbCanLevel();
	void SetyAccrueAt(LONG propval);
	LONG GetyAccrueAt();
	CDuration GetoWork();
	CDuration GetoRegularWork();
	CDuration GetoOvertimeWork();
	CDuration GetoActualWork();
	CDuration GetoRemainingWork();
	CDuration GetoActualOvertimeWork();
	CDuration GetoRemainingOvertimeWork();
	void SetlPercentWorkComplete(LONG propval);
	LONG GetlPercentWorkComplete();
	void SetcStandardRate(DOUBLE propval);
	DOUBLE GetcStandardRate();
	void SetyStandardRateFormat(LONG propval);
	LONG GetyStandardRateFormat();
	void SetcCost(DOUBLE propval);
	DOUBLE GetcCost();
	void SetcOvertimeRate(DOUBLE propval);
	DOUBLE GetcOvertimeRate();
	void SetyOvertimeRateFormat(LONG propval);
	LONG GetyOvertimeRateFormat();
	void SetcOvertimeCost(DOUBLE propval);
	DOUBLE GetcOvertimeCost();
	void SetcCostPerUse(DOUBLE propval);
	DOUBLE GetcCostPerUse();
	void SetcActualCost(DOUBLE propval);
	DOUBLE GetcActualCost();
	void SetcActualOvertimeCost(DOUBLE propval);
	DOUBLE GetcActualOvertimeCost();
	void SetcRemainingCost(DOUBLE propval);
	DOUBLE GetcRemainingCost();
	void SetcRemainingOvertimeCost(DOUBLE propval);
	DOUBLE GetcRemainingOvertimeCost();
	void SetfWorkVariance(FLOAT propval);
	FLOAT GetfWorkVariance();
	void SetfCostVariance(FLOAT propval);
	FLOAT GetfCostVariance();
	void SetfSV(FLOAT propval);
	FLOAT GetfSV();
	void SetfCV(FLOAT propval);
	FLOAT GetfCV();
	void SetfACWP(FLOAT propval);
	FLOAT GetfACWP();
	void SetlCalendarUID(LONG propval);
	LONG GetlCalendarUID();
	void SetsNotes(LPCTSTR propval);
	CString GetsNotes();
	void SetfBCWS(FLOAT propval);
	FLOAT GetfBCWS();
	void SetfBCWP(FLOAT propval);
	FLOAT GetfBCWP();
	void SetbIsGeneric(BOOL propval);
	BOOL GetbIsGeneric();
	void SetbIsInactive(BOOL propval);
	BOOL GetbIsInactive();
	void SetbIsEnterprise(BOOL propval);
	BOOL GetbIsEnterprise();
	void SetyBookingType(LONG propval);
	LONG GetyBookingType();
	CDuration GetoActualWorkProtected();
	CDuration GetoActualOvertimeWorkProtected();
	void SetsActiveDirectoryGUID(LPCTSTR propval);
	CString GetsActiveDirectoryGUID();
	void SetdtCreationDate(DATE propval);
	DATE GetdtCreationDate();
	CResourceExtendedAttribute_C GetoExtendedAttribute_C();
	CResourceBaseline_C GetoBaseline_C();
	CResourceOutlineCode_C GetoOutlineCode_C();
	CResourceAvailabilityPeriods GetoAvailabilityPeriods();
	CResourceRates GetoRates();
	CTimephasedData_C GetoTimephasedData_C();
	void SetKey(LPCTSTR propval);
	CString GetKey();

// Methods
public:
	CString GetXML(void);
	void SetXML(LPCTSTR sXML);
	BOOL IsNull(void);
	void Initialize(void);
};

class CResources : public COleDispatchDriver
{
// Constructors
public:
	CResources() {}	
	CResources(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CResources(const CResources& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Properties
public:
	LONG GetCount();

// Methods
public:
	CResource Item(LPCTSTR Index);
	CResource Add(void);
	void Clear(void);
	void Remove(LPCTSTR Index);
	BOOL IsNull(void);
	void Initialize(void);
	void SetXML(LPCTSTR sXML);
	CString GetXML(void);
};

class CTaskOutlineCode : public COleDispatchDriver
{
// Constructors
public:
	CTaskOutlineCode() {}	
	CTaskOutlineCode(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CTaskOutlineCode(const CTaskOutlineCode& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Properties
public:
	void SetlUID(LONG propval);
	LONG GetlUID();
	void SetsFieldID(LPCTSTR propval);
	CString GetsFieldID();
	void SetlValueID(LONG propval);
	LONG GetlValueID();
	void SetKey(LPCTSTR propval);
	CString GetKey();

// Methods
public:
	CString GetXML(void);
	void SetXML(LPCTSTR sXML);
	BOOL IsNull(void);
	void Initialize(void);
};

class CTaskOutlineCode_C : public COleDispatchDriver
{
// Constructors
public:
	CTaskOutlineCode_C() {}	
	CTaskOutlineCode_C(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CTaskOutlineCode_C(const CTaskOutlineCode_C& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Properties
public:
	LONG GetCount();

// Methods
public:
	CTaskOutlineCode Item(LPCTSTR Index);
	CTaskOutlineCode Add(void);
	void Clear(void);
	void Remove(LPCTSTR Index);
	BOOL IsNull(void);
	void Initialize(void);
};

class CTaskBaseline : public COleDispatchDriver
{
// Constructors
public:
	CTaskBaseline() {}	
	CTaskBaseline(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CTaskBaseline(const CTaskBaseline& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Properties
public:
	CTimephasedData_C GetoTimephasedData_C();
	void SetlNumber(LONG propval);
	LONG GetlNumber();
	void SetbInterim(BOOL propval);
	BOOL GetbInterim();
	void SetdtStart(DATE propval);
	DATE GetdtStart();
	void SetdtFinish(DATE propval);
	DATE GetdtFinish();
	CDuration GetoDuration();
	void SetyDurationFormat(LONG propval);
	LONG GetyDurationFormat();
	void SetbEstimatedDuration(BOOL propval);
	BOOL GetbEstimatedDuration();
	CDuration GetoWork();
	void SetcCost(DOUBLE propval);
	DOUBLE GetcCost();
	void SetfBCWS(FLOAT propval);
	FLOAT GetfBCWS();
	void SetfBCWP(FLOAT propval);
	FLOAT GetfBCWP();
	void SetKey(LPCTSTR propval);
	CString GetKey();

// Methods
public:
	CString GetXML(void);
	void SetXML(LPCTSTR sXML);
	BOOL IsNull(void);
	void Initialize(void);
};

class CTaskBaseline_C : public COleDispatchDriver
{
// Constructors
public:
	CTaskBaseline_C() {}	
	CTaskBaseline_C(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CTaskBaseline_C(const CTaskBaseline_C& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Properties
public:
	LONG GetCount();

// Methods
public:
	CTaskBaseline Item(LPCTSTR Index);
	CTaskBaseline Add(void);
	void Clear(void);
	void Remove(LPCTSTR Index);
	BOOL IsNull(void);
	void Initialize(void);
};

class CTaskExtendedAttribute : public COleDispatchDriver
{
// Constructors
public:
	CTaskExtendedAttribute() {}	
	CTaskExtendedAttribute(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CTaskExtendedAttribute(const CTaskExtendedAttribute& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Properties
public:
	void SetlUID(LONG propval);
	LONG GetlUID();
	void SetsFieldID(LPCTSTR propval);
	CString GetsFieldID();
	void SetsValue(LPCTSTR propval);
	CString GetsValue();
	void SetlValueID(LONG propval);
	LONG GetlValueID();
	void SetyDurationFormat(LONG propval);
	LONG GetyDurationFormat();
	void SetKey(LPCTSTR propval);
	CString GetKey();

// Methods
public:
	CString GetXML(void);
	void SetXML(LPCTSTR sXML);
	BOOL IsNull(void);
	void Initialize(void);
};

class CTaskExtendedAttribute_C : public COleDispatchDriver
{
// Constructors
public:
	CTaskExtendedAttribute_C() {}	
	CTaskExtendedAttribute_C(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CTaskExtendedAttribute_C(const CTaskExtendedAttribute_C& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Properties
public:
	LONG GetCount();

// Methods
public:
	CTaskExtendedAttribute Item(LPCTSTR Index);
	CTaskExtendedAttribute Add(void);
	void Clear(void);
	void Remove(LPCTSTR Index);
	BOOL IsNull(void);
	void Initialize(void);
};

class CTaskPredecessorLink : public COleDispatchDriver
{
// Constructors
public:
	CTaskPredecessorLink() {}	
	CTaskPredecessorLink(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CTaskPredecessorLink(const CTaskPredecessorLink& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Properties
public:
	void SetlPredecessorUID(LONG propval);
	LONG GetlPredecessorUID();
	void SetyType(LONG propval);
	LONG GetyType();
	void SetbCrossProject(BOOL propval);
	BOOL GetbCrossProject();
	void SetsCrossProjectName(LPCTSTR propval);
	CString GetsCrossProjectName();
	void SetlLinkLag(LONG propval);
	LONG GetlLinkLag();
	void SetyLagFormat(LONG propval);
	LONG GetyLagFormat();
	void SetKey(LPCTSTR propval);
	CString GetKey();

// Methods
public:
	CString GetXML(void);
	void SetXML(LPCTSTR sXML);
	BOOL IsNull(void);
	void Initialize(void);
};

class CTaskPredecessorLink_C : public COleDispatchDriver
{
// Constructors
public:
	CTaskPredecessorLink_C() {}	
	CTaskPredecessorLink_C(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CTaskPredecessorLink_C(const CTaskPredecessorLink_C& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Properties
public:
	LONG GetCount();

// Methods
public:
	CTaskPredecessorLink Item(LPCTSTR Index);
	CTaskPredecessorLink Add(void);
	void Clear(void);
	void Remove(LPCTSTR Index);
	BOOL IsNull(void);
	void Initialize(void);
};

class CTask : public COleDispatchDriver
{
// Constructors
public:
	CTask() {}	
	CTask(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CTask(const CTask& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Properties
public:
	void SetlUID(LONG propval);
	LONG GetlUID();
	void SetlID(LONG propval);
	LONG GetlID();
	void SetsName(LPCTSTR propval);
	CString GetsName();
	void SetyType(LONG propval);
	LONG GetyType();
	void SetbIsNull(BOOL propval);
	BOOL GetbIsNull();
	void SetdtCreateDate(DATE propval);
	DATE GetdtCreateDate();
	void SetsContact(LPCTSTR propval);
	CString GetsContact();
	void SetsWBS(LPCTSTR propval);
	CString GetsWBS();
	void SetsWBSLevel(LPCTSTR propval);
	CString GetsWBSLevel();
	void SetsOutlineNumber(LPCTSTR propval);
	CString GetsOutlineNumber();
	void SetlOutlineLevel(LONG propval);
	LONG GetlOutlineLevel();
	void SetlPriority(LONG propval);
	LONG GetlPriority();
	void SetdtStart(DATE propval);
	DATE GetdtStart();
	void SetdtFinish(DATE propval);
	DATE GetdtFinish();
	CDuration GetoDuration();
	void SetyDurationFormat(LONG propval);
	LONG GetyDurationFormat();
	CDuration GetoWork();
	void SetdtStop(DATE propval);
	DATE GetdtStop();
	void SetdtResume(DATE propval);
	DATE GetdtResume();
	void SetbResumeValid(BOOL propval);
	BOOL GetbResumeValid();
	void SetbEffortDriven(BOOL propval);
	BOOL GetbEffortDriven();
	void SetbRecurring(BOOL propval);
	BOOL GetbRecurring();
	void SetbOverAllocated(BOOL propval);
	BOOL GetbOverAllocated();
	void SetbEstimated(BOOL propval);
	BOOL GetbEstimated();
	void SetbMilestone(BOOL propval);
	BOOL GetbMilestone();
	void SetbSummary(BOOL propval);
	BOOL GetbSummary();
	void SetbCritical(BOOL propval);
	BOOL GetbCritical();
	void SetbIsSubproject(BOOL propval);
	BOOL GetbIsSubproject();
	void SetbIsSubprojectReadOnly(BOOL propval);
	BOOL GetbIsSubprojectReadOnly();
	void SetsSubprojectName(LPCTSTR propval);
	CString GetsSubprojectName();
	void SetbExternalTask(BOOL propval);
	BOOL GetbExternalTask();
	void SetsExternalTaskProject(LPCTSTR propval);
	CString GetsExternalTaskProject();
	void SetdtEarlyStart(DATE propval);
	DATE GetdtEarlyStart();
	void SetdtEarlyFinish(DATE propval);
	DATE GetdtEarlyFinish();
	void SetdtLateStart(DATE propval);
	DATE GetdtLateStart();
	void SetdtLateFinish(DATE propval);
	DATE GetdtLateFinish();
	void SetlStartVariance(LONG propval);
	LONG GetlStartVariance();
	void SetlFinishVariance(LONG propval);
	LONG GetlFinishVariance();
	void SetfWorkVariance(FLOAT propval);
	FLOAT GetfWorkVariance();
	void SetlFreeSlack(LONG propval);
	LONG GetlFreeSlack();
	void SetlTotalSlack(LONG propval);
	LONG GetlTotalSlack();
	void SetfFixedCost(FLOAT propval);
	FLOAT GetfFixedCost();
	void SetyFixedCostAccrual(LONG propval);
	LONG GetyFixedCostAccrual();
	void SetlPercentComplete(LONG propval);
	LONG GetlPercentComplete();
	void SetlPercentWorkComplete(LONG propval);
	LONG GetlPercentWorkComplete();
	void SetcCost(DOUBLE propval);
	DOUBLE GetcCost();
	void SetcOvertimeCost(DOUBLE propval);
	DOUBLE GetcOvertimeCost();
	CDuration GetoOvertimeWork();
	void SetdtActualStart(DATE propval);
	DATE GetdtActualStart();
	void SetdtActualFinish(DATE propval);
	DATE GetdtActualFinish();
	CDuration GetoActualDuration();
	void SetcActualCost(DOUBLE propval);
	DOUBLE GetcActualCost();
	void SetcActualOvertimeCost(DOUBLE propval);
	DOUBLE GetcActualOvertimeCost();
	CDuration GetoActualWork();
	CDuration GetoActualOvertimeWork();
	CDuration GetoRegularWork();
	CDuration GetoRemainingDuration();
	void SetcRemainingCost(DOUBLE propval);
	DOUBLE GetcRemainingCost();
	CDuration GetoRemainingWork();
	void SetcRemainingOvertimeCost(DOUBLE propval);
	DOUBLE GetcRemainingOvertimeCost();
	CDuration GetoRemainingOvertimeWork();
	void SetfACWP(FLOAT propval);
	FLOAT GetfACWP();
	void SetfCV(FLOAT propval);
	FLOAT GetfCV();
	void SetyConstraintType(LONG propval);
	LONG GetyConstraintType();
	void SetlCalendarUID(LONG propval);
	LONG GetlCalendarUID();
	void SetdtConstraintDate(DATE propval);
	DATE GetdtConstraintDate();
	void SetdtDeadline(DATE propval);
	DATE GetdtDeadline();
	void SetbLevelAssignments(BOOL propval);
	BOOL GetbLevelAssignments();
	void SetbLevelingCanSplit(BOOL propval);
	BOOL GetbLevelingCanSplit();
	void SetlLevelingDelay(LONG propval);
	LONG GetlLevelingDelay();
	void SetyLevelingDelayFormat(LONG propval);
	LONG GetyLevelingDelayFormat();
	void SetdtPreLeveledStart(DATE propval);
	DATE GetdtPreLeveledStart();
	void SetdtPreLeveledFinish(DATE propval);
	DATE GetdtPreLeveledFinish();
	void SetsHyperlink(LPCTSTR propval);
	CString GetsHyperlink();
	void SetsHyperlinkAddress(LPCTSTR propval);
	CString GetsHyperlinkAddress();
	void SetsHyperlinkSubAddress(LPCTSTR propval);
	CString GetsHyperlinkSubAddress();
	void SetbIgnoreResourceCalendar(BOOL propval);
	BOOL GetbIgnoreResourceCalendar();
	void SetsNotes(LPCTSTR propval);
	CString GetsNotes();
	void SetbHideBar(BOOL propval);
	BOOL GetbHideBar();
	void SetbRollup(BOOL propval);
	BOOL GetbRollup();
	void SetfBCWS(FLOAT propval);
	FLOAT GetfBCWS();
	void SetfBCWP(FLOAT propval);
	FLOAT GetfBCWP();
	void SetlPhysicalPercentComplete(LONG propval);
	LONG GetlPhysicalPercentComplete();
	void SetyEarnedValueMethod(LONG propval);
	LONG GetyEarnedValueMethod();
	CTaskPredecessorLink_C GetoPredecessorLink_C();
	CDuration GetoActualWorkProtected();
	CDuration GetoActualOvertimeWorkProtected();
	CTaskExtendedAttribute_C GetoExtendedAttribute_C();
	CTaskBaseline_C GetoBaseline_C();
	CTaskOutlineCode_C GetoOutlineCode_C();
	CTimephasedData_C GetoTimephasedData_C();
	void SetKey(LPCTSTR propval);
	CString GetKey();

// Methods
public:
	CString GetXML(void);
	void SetXML(LPCTSTR sXML);
	BOOL IsNull(void);
	void Initialize(void);
};

class CTasks : public COleDispatchDriver
{
// Constructors
public:
	CTasks() {}	
	CTasks(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CTasks(const CTasks& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Properties
public:
	LONG GetCount();

// Methods
public:
	CTask Item(LPCTSTR Index);
	CTask Add(void);
	void Clear(void);
	void Remove(LPCTSTR Index);
	BOOL IsNull(void);
	void Initialize(void);
	void SetXML(LPCTSTR sXML);
	CString GetXML(void);
};

class CWorkingTime : public COleDispatchDriver
{
// Constructors
public:
	CWorkingTime() {}	
	CWorkingTime(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CWorkingTime(const CWorkingTime& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Properties
public:
	CTime GetoFromTime();
	CTime GetoToTime();
	void SetKey(LPCTSTR propval);
	CString GetKey();

// Methods
public:
	CString GetXML(void);
	void SetXML(LPCTSTR sXML);
	BOOL IsNull(void);
	void Initialize(void);
};

class CWorkingTimes : public COleDispatchDriver
{
// Constructors
public:
	CWorkingTimes() {}	
	CWorkingTimes(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CWorkingTimes(const CWorkingTimes& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Properties
public:
	LONG GetCount();

// Methods
public:
	CWorkingTime Item(LPCTSTR Index);
	CWorkingTime Add(void);
	void Clear(void);
	void Remove(LPCTSTR Index);
	BOOL IsNull(void);
	void Initialize(void);
	void SetXML(LPCTSTR sXML);
	CString GetXML(void);
};

class CTimePeriod : public COleDispatchDriver
{
// Constructors
public:
	CTimePeriod() {}	
	CTimePeriod(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CTimePeriod(const CTimePeriod& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Properties
public:
	void SetdtFromDate(DATE propval);
	DATE GetdtFromDate();
	void SetdtToDate(DATE propval);
	DATE GetdtToDate();
	void SetKey(LPCTSTR propval);
	CString GetKey();

// Methods
public:
	CString GetXML(void);
	void SetXML(LPCTSTR sXML);
	BOOL IsNull(void);
	void Initialize(void);
};

class CTimePeriods : public COleDispatchDriver
{
// Constructors
public:
	CTimePeriods() {}	
	CTimePeriods(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CTimePeriods(const CTimePeriods& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Properties
public:
	LONG GetCount();

// Methods
public:
	CTimePeriod Item(LPCTSTR Index);
	CTimePeriod Add(void);
	void Clear(void);
	void Remove(LPCTSTR Index);
	BOOL IsNull(void);
	void Initialize(void);
	void SetXML(LPCTSTR sXML);
	CString GetXML(void);
};

class CWeekDay : public COleDispatchDriver
{
// Constructors
public:
	CWeekDay() {}	
	CWeekDay(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CWeekDay(const CWeekDay& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Properties
public:
	void SetyDayType(LONG propval);
	LONG GetyDayType();
	void SetbDayWorking(BOOL propval);
	BOOL GetbDayWorking();
	CTimePeriod GetoTimePeriod();
	CWorkingTimes GetoWorkingTimes();
	void SetKey(LPCTSTR propval);
	CString GetKey();

// Methods
public:
	CString GetXML(void);
	void SetXML(LPCTSTR sXML);
	BOOL IsNull(void);
	void Initialize(void);
};

class CWeekDays : public COleDispatchDriver
{
// Constructors
public:
	CWeekDays() {}	
	CWeekDays(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CWeekDays(const CWeekDays& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Properties
public:
	LONG GetCount();

// Methods
public:
	CWeekDay Item(LPCTSTR Index);
	CWeekDay Add(void);
	void Clear(void);
	void Remove(LPCTSTR Index);
	BOOL IsNull(void);
	void Initialize(void);
	void SetXML(LPCTSTR sXML);
	CString GetXML(void);
};

class CCalendar : public COleDispatchDriver
{
// Constructors
public:
	CCalendar() {}	
	CCalendar(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CCalendar(const CCalendar& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Properties
public:
	void SetlUID(LONG propval);
	LONG GetlUID();
	void SetsName(LPCTSTR propval);
	CString GetsName();
	void SetbIsBaseCalendar(BOOL propval);
	BOOL GetbIsBaseCalendar();
	void SetlBaseCalendarUID(LONG propval);
	LONG GetlBaseCalendarUID();
	CWeekDays GetoWeekDays();
	void SetKey(LPCTSTR propval);
	CString GetKey();

// Methods
public:
	CString GetXML(void);
	void SetXML(LPCTSTR sXML);
	BOOL IsNull(void);
	void Initialize(void);
};

class CCalendars : public COleDispatchDriver
{
// Constructors
public:
	CCalendars() {}	
	CCalendars(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CCalendars(const CCalendars& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Properties
public:
	LONG GetCount();

// Methods
public:
	CCalendar Item(LPCTSTR Index);
	CCalendar Add(void);
	void Clear(void);
	void Remove(LPCTSTR Index);
	BOOL IsNull(void);
	void Initialize(void);
	void SetXML(LPCTSTR sXML);
	CString GetXML(void);
};

class CExtendedAttributeValue : public COleDispatchDriver
{
// Constructors
public:
	CExtendedAttributeValue() {}	
	CExtendedAttributeValue(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CExtendedAttributeValue(const CExtendedAttributeValue& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Properties
public:
	void SetlID(LONG propval);
	LONG GetlID();
	void SetsValue(LPCTSTR propval);
	CString GetsValue();
	void SetsDescription(LPCTSTR propval);
	CString GetsDescription();
	void SetKey(LPCTSTR propval);
	CString GetKey();

// Methods
public:
	CString GetXML(void);
	void SetXML(LPCTSTR sXML);
	BOOL IsNull(void);
	void Initialize(void);
};

class CExtendedAttributeValueList : public COleDispatchDriver
{
// Constructors
public:
	CExtendedAttributeValueList() {}	
	CExtendedAttributeValueList(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CExtendedAttributeValueList(const CExtendedAttributeValueList& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Properties
public:
	LONG GetCount();

// Methods
public:
	CExtendedAttributeValue Item(LPCTSTR Index);
	CExtendedAttributeValue Add(void);
	void Clear(void);
	void Remove(LPCTSTR Index);
	BOOL IsNull(void);
	void Initialize(void);
	void SetXML(LPCTSTR sXML);
	CString GetXML(void);
};

class CExtendedAttribute : public COleDispatchDriver
{
// Constructors
public:
	CExtendedAttribute() {}	
	CExtendedAttribute(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CExtendedAttribute(const CExtendedAttribute& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Properties
public:
	void SetsFieldID(LPCTSTR propval);
	CString GetsFieldID();
	void SetsFieldName(LPCTSTR propval);
	CString GetsFieldName();
	void SetsAlias(LPCTSTR propval);
	CString GetsAlias();
	void SetsPhoneticAlias(LPCTSTR propval);
	CString GetsPhoneticAlias();
	void SetyRollupType(LONG propval);
	LONG GetyRollupType();
	void SetyCalculationType(LONG propval);
	LONG GetyCalculationType();
	void SetsFormula(LPCTSTR propval);
	CString GetsFormula();
	void SetbRestrictValues(BOOL propval);
	BOOL GetbRestrictValues();
	void SetyValuelistSortOrder(LONG propval);
	LONG GetyValuelistSortOrder();
	void SetbAppendNewValues(BOOL propval);
	BOOL GetbAppendNewValues();
	void SetsDefault(LPCTSTR propval);
	CString GetsDefault();
	CExtendedAttributeValueList GetoValueList();
	void SetKey(LPCTSTR propval);
	CString GetKey();

// Methods
public:
	CString GetXML(void);
	void SetXML(LPCTSTR sXML);
	BOOL IsNull(void);
	void Initialize(void);
};

class CExtendedAttributes : public COleDispatchDriver
{
// Constructors
public:
	CExtendedAttributes() {}	
	CExtendedAttributes(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CExtendedAttributes(const CExtendedAttributes& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Properties
public:
	LONG GetCount();

// Methods
public:
	CExtendedAttribute Item(LPCTSTR Index);
	CExtendedAttribute Add(void);
	void Clear(void);
	void Remove(LPCTSTR Index);
	BOOL IsNull(void);
	void Initialize(void);
	void SetXML(LPCTSTR sXML);
	CString GetXML(void);
};

class CWBSMask : public COleDispatchDriver
{
// Constructors
public:
	CWBSMask() {}	
	CWBSMask(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CWBSMask(const CWBSMask& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Properties
public:
	void SetlLevel(LONG propval);
	LONG GetlLevel();
	void SetyType(LONG propval);
	LONG GetyType();
	void SetsLength(LPCTSTR propval);
	CString GetsLength();
	void SetsSeparator(LPCTSTR propval);
	CString GetsSeparator();
	void SetKey(LPCTSTR propval);
	CString GetKey();

// Methods
public:
	CString GetXML(void);
	void SetXML(LPCTSTR sXML);
	BOOL IsNull(void);
	void Initialize(void);
};

class CWBSMasks : public COleDispatchDriver
{
// Constructors
public:
	CWBSMasks() {}	
	CWBSMasks(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CWBSMasks(const CWBSMasks& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Properties
public:
	LONG GetCount();

// Methods
public:
	CWBSMask Item(LPCTSTR Index);
	CWBSMask Add(void);
	void Clear(void);
	void Remove(LPCTSTR Index);
	BOOL IsNull(void);
	void Initialize(void);
	void SetXML(LPCTSTR sXML);
	CString GetXML(void);
};

class COutlineCodeMask : public COleDispatchDriver
{
// Constructors
public:
	COutlineCodeMask() {}	
	COutlineCodeMask(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	COutlineCodeMask(const COutlineCodeMask& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Properties
public:
	void SetlLevel(LONG propval);
	LONG GetlLevel();
	void SetyType(LONG propval);
	LONG GetyType();
	void SetlLength(LONG propval);
	LONG GetlLength();
	void SetsSeparator(LPCTSTR propval);
	CString GetsSeparator();
	void SetKey(LPCTSTR propval);
	CString GetKey();

// Methods
public:
	CString GetXML(void);
	void SetXML(LPCTSTR sXML);
	BOOL IsNull(void);
	void Initialize(void);
};

class COutlineCodeMasks : public COleDispatchDriver
{
// Constructors
public:
	COutlineCodeMasks() {}	
	COutlineCodeMasks(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	COutlineCodeMasks(const COutlineCodeMasks& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Properties
public:
	LONG GetCount();

// Methods
public:
	COutlineCodeMask Item(LPCTSTR Index);
	COutlineCodeMask Add(void);
	void Clear(void);
	void Remove(LPCTSTR Index);
	BOOL IsNull(void);
	void Initialize(void);
	void SetXML(LPCTSTR sXML);
	CString GetXML(void);
};

class COutlineCodeValue : public COleDispatchDriver
{
// Constructors
public:
	COutlineCodeValue() {}	
	COutlineCodeValue(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	COutlineCodeValue(const COutlineCodeValue& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Properties
public:
	void SetlValueID(LONG propval);
	LONG GetlValueID();
	void SetlParentValueID(LONG propval);
	LONG GetlParentValueID();
	void SetsValue(LPCTSTR propval);
	CString GetsValue();
	void SetsDescription(LPCTSTR propval);
	CString GetsDescription();
	void SetKey(LPCTSTR propval);
	CString GetKey();

// Methods
public:
	CString GetXML(void);
	void SetXML(LPCTSTR sXML);
	BOOL IsNull(void);
	void Initialize(void);
};

class COutlineCodeValues : public COleDispatchDriver
{
// Constructors
public:
	COutlineCodeValues() {}	
	COutlineCodeValues(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	COutlineCodeValues(const COutlineCodeValues& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Properties
public:
	LONG GetCount();

// Methods
public:
	COutlineCodeValue Item(LPCTSTR Index);
	COutlineCodeValue Add(void);
	void Clear(void);
	void Remove(LPCTSTR Index);
	BOOL IsNull(void);
	void Initialize(void);
	void SetXML(LPCTSTR sXML);
	CString GetXML(void);
};

class COutlineCode : public COleDispatchDriver
{
// Constructors
public:
	COutlineCode() {}	
	COutlineCode(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	COutlineCode(const COutlineCode& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Properties
public:
	void SetsFieldID(LPCTSTR propval);
	CString GetsFieldID();
	void SetsFieldName(LPCTSTR propval);
	CString GetsFieldName();
	void SetsAlias(LPCTSTR propval);
	CString GetsAlias();
	void SetsPhoneticAlias(LPCTSTR propval);
	CString GetsPhoneticAlias();
	COutlineCodeValues GetoValues();
	void SetbEnterprise(BOOL propval);
	BOOL GetbEnterprise();
	void SetlEnterpriseOutlineCodeAlias(LONG propval);
	LONG GetlEnterpriseOutlineCodeAlias();
	void SetbResourceSubstitutionEnabled(BOOL propval);
	BOOL GetbResourceSubstitutionEnabled();
	void SetbLeafOnly(BOOL propval);
	BOOL GetbLeafOnly();
	void SetbAllLevelsRequired(BOOL propval);
	BOOL GetbAllLevelsRequired();
	void SetbOnlyTableValuesAllowed(BOOL propval);
	BOOL GetbOnlyTableValuesAllowed();
	COutlineCodeMasks GetoMasks();
	void SetKey(LPCTSTR propval);
	CString GetKey();

// Methods
public:
	CString GetXML(void);
	void SetXML(LPCTSTR sXML);
	BOOL IsNull(void);
	void Initialize(void);
};

class COutlineCodes : public COleDispatchDriver
{
// Constructors
public:
	COutlineCodes() {}	
	COutlineCodes(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	COutlineCodes(const COutlineCodes& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Properties
public:
	LONG GetCount();

// Methods
public:
	COutlineCode Item(LPCTSTR Index);
	COutlineCode Add(void);
	void Clear(void);
	void Remove(LPCTSTR Index);
	BOOL IsNull(void);
	void Initialize(void);
	void SetXML(LPCTSTR sXML);
	CString GetXML(void);
};

class CMP11 : public COleDispatchDriver
{
// Constructors
public:
	CMP11() {}	
	CMP11(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CMP11(const CMP11& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Properties
public:
	void SetsUID(LPCTSTR propval);
	CString GetsUID();
	void SetsName(LPCTSTR propval);
	CString GetsName();
	void SetsTitle(LPCTSTR propval);
	CString GetsTitle();
	void SetsSubject(LPCTSTR propval);
	CString GetsSubject();
	void SetsCategory(LPCTSTR propval);
	CString GetsCategory();
	void SetsCompany(LPCTSTR propval);
	CString GetsCompany();
	void SetsManager(LPCTSTR propval);
	CString GetsManager();
	void SetsAuthor(LPCTSTR propval);
	CString GetsAuthor();
	void SetdtCreationDate(DATE propval);
	DATE GetdtCreationDate();
	void SetlRevision(LONG propval);
	LONG GetlRevision();
	void SetdtLastSaved(DATE propval);
	DATE GetdtLastSaved();
	void SetbScheduleFromStart(BOOL propval);
	BOOL GetbScheduleFromStart();
	void SetdtStartDate(DATE propval);
	DATE GetdtStartDate();
	void SetdtFinishDate(DATE propval);
	DATE GetdtFinishDate();
	void SetyFYStartDate(LONG propval);
	LONG GetyFYStartDate();
	void SetlCriticalSlackLimit(LONG propval);
	LONG GetlCriticalSlackLimit();
	void SetlCurrencyDigits(LONG propval);
	LONG GetlCurrencyDigits();
	void SetsCurrencySymbol(LPCTSTR propval);
	CString GetsCurrencySymbol();
	void SetyCurrencySymbolPosition(LONG propval);
	LONG GetyCurrencySymbolPosition();
	void SetlCalendarUID(LONG propval);
	LONG GetlCalendarUID();
	CTime GetoDefaultStartTime();
	CTime GetoDefaultFinishTime();
	void SetlMinutesPerDay(LONG propval);
	LONG GetlMinutesPerDay();
	void SetlMinutesPerWeek(LONG propval);
	LONG GetlMinutesPerWeek();
	void SetlDaysPerMonth(LONG propval);
	LONG GetlDaysPerMonth();
	void SetyDefaultTaskType(LONG propval);
	LONG GetyDefaultTaskType();
	void SetyDefaultFixedCostAccrual(LONG propval);
	LONG GetyDefaultFixedCostAccrual();
	void SetfDefaultStandardRate(FLOAT propval);
	FLOAT GetfDefaultStandardRate();
	void SetfDefaultOvertimeRate(FLOAT propval);
	FLOAT GetfDefaultOvertimeRate();
	void SetyDurationFormat(LONG propval);
	LONG GetyDurationFormat();
	void SetyWorkFormat(LONG propval);
	LONG GetyWorkFormat();
	void SetbEditableActualCosts(BOOL propval);
	BOOL GetbEditableActualCosts();
	void SetbHonorConstraints(BOOL propval);
	BOOL GetbHonorConstraints();
	void SetyEarnedValueMethod(LONG propval);
	LONG GetyEarnedValueMethod();
	void SetbInsertedProjectsLikeSummary(BOOL propval);
	BOOL GetbInsertedProjectsLikeSummary();
	void SetbMultipleCriticalPaths(BOOL propval);
	BOOL GetbMultipleCriticalPaths();
	void SetbNewTasksEffortDriven(BOOL propval);
	BOOL GetbNewTasksEffortDriven();
	void SetbNewTasksEstimated(BOOL propval);
	BOOL GetbNewTasksEstimated();
	void SetbSplitsInProgressTasks(BOOL propval);
	BOOL GetbSplitsInProgressTasks();
	void SetbSpreadActualCost(BOOL propval);
	BOOL GetbSpreadActualCost();
	void SetbSpreadPercentComplete(BOOL propval);
	BOOL GetbSpreadPercentComplete();
	void SetbTaskUpdatesResource(BOOL propval);
	BOOL GetbTaskUpdatesResource();
	void SetbFiscalYearStart(BOOL propval);
	BOOL GetbFiscalYearStart();
	void SetyWeekStartDay(LONG propval);
	LONG GetyWeekStartDay();
	void SetbMoveCompletedEndsBack(BOOL propval);
	BOOL GetbMoveCompletedEndsBack();
	void SetbMoveRemainingStartsBack(BOOL propval);
	BOOL GetbMoveRemainingStartsBack();
	void SetbMoveRemainingStartsForward(BOOL propval);
	BOOL GetbMoveRemainingStartsForward();
	void SetbMoveCompletedEndsForward(BOOL propval);
	BOOL GetbMoveCompletedEndsForward();
	void SetyBaselineForEarnedValue(LONG propval);
	LONG GetyBaselineForEarnedValue();
	void SetbAutoAddNewResourcesAndTasks(BOOL propval);
	BOOL GetbAutoAddNewResourcesAndTasks();
	void SetdtStatusDate(DATE propval);
	DATE GetdtStatusDate();
	void SetdtCurrentDate(DATE propval);
	DATE GetdtCurrentDate();
	void SetbMicrosoftProjectServerURL(BOOL propval);
	BOOL GetbMicrosoftProjectServerURL();
	void SetbAutolink(BOOL propval);
	BOOL GetbAutolink();
	void SetyNewTaskStartDate(LONG propval);
	LONG GetyNewTaskStartDate();
	void SetyDefaultTaskEVMethod(LONG propval);
	LONG GetyDefaultTaskEVMethod();
	void SetbProjectExternallyEdited(BOOL propval);
	BOOL GetbProjectExternallyEdited();
	void SetdtExtendedCreationDate(DATE propval);
	DATE GetdtExtendedCreationDate();
	void SetbActualsInSync(BOOL propval);
	BOOL GetbActualsInSync();
	void SetbRemoveFileProperties(BOOL propval);
	BOOL GetbRemoveFileProperties();
	void SetbAdminProject(BOOL propval);
	BOOL GetbAdminProject();
	COutlineCodes GetoOutlineCodes();
	CWBSMasks GetoWBSMasks();
	CExtendedAttributes GetoExtendedAttributes();
	CCalendars GetoCalendars();
	CTasks GetoTasks();
	CResources GetoResources();
	CAssignments GetoAssignments();

// Methods
public:
	void WriteXML(LPCTSTR url);
	void ReadXML(LPCTSTR url);
	CString GetXML(void);
	void SetXML(LPCTSTR sXML);
	BOOL IsNull(void);
	void Initialize(void);
};

class CMSP2003Ctl : public CWnd
{
protected:
	DECLARE_DYNCREATE(CMSP2003Ctl)
public:
	CLSID const& GetClsid()
	{
		static CLSID const clsid = { 0xAF1509E7, 0x4A8C, 0x4306, { 0x81, 0x16, 0xE5, 0x11, 0xD4, 0xFF, 0x55, 0x82} };
		return clsid;
	}

	virtual BOOL Create(LPCTSTR lpszClassName, LPCTSTR lpszWindowName, DWORD dwStyle, const RECT& rect, CWnd* pParentWnd, UINT nID, CCreateContext* pContext = NULL)
	{
		return CreateControl(GetClsid(), lpszWindowName, dwStyle, rect, pParentWnd, nID);
	}

	BOOL Create(LPCTSTR lpszWindowName, DWORD dwStyle, const RECT& rect, CWnd* pParentWnd, UINT nID, CFile* pPersist = NULL, BOOL bStorage = FALSE, BSTR bstrLicKey = NULL)
	{
		return CreateControl(GetClsid(), lpszWindowName, dwStyle, rect, pParentWnd, nID, pPersist, bStorage, bstrLicKey);
	}

// Properties
public:
	CMP11 GetMP11();

// Methods
public:
	void AboutBox(void);
	void Clear(void);
};
}