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

namespace MSP2010
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
	T_DATE = 4,
	T_DURATION = 6,
	T_COST = 9,
	T_NUMBER = 15,
	T_FLAG = 17,
	T_TEXT = 21,
	T_FINISH_DATE = 27,
}E_TYPE;

typedef enum E_TYPE_1
{
	T_1_NUMBERS = 0,
	T_1_UPPER_CASE_LETTERS = 1,
	T_1_LOWER_CASE_LETTERS = 2,
	T_1_CHARACTERS = 3,
}E_TYPE_1;

typedef enum E_TYPE_2
{
	T_2_NUMBERS = 0,
	T_2_UPPERCASE_LETTERS = 1,
	T_2_LOWERCASE_LETTERS = 2,
	T_2_CHARACTERS = 3,
}E_TYPE_2;

typedef enum E_CFTYPE
{
	CFT_COST = 0,
	CFT_DATE = 1,
	CFT_DURATION = 2,
	CFT_FINISH = 3,
	CFT_FLAG = 4,
	CFT_NUMBER = 5,
	CFT_START = 6,
	CFT_TEXT = 7,
}E_CFTYPE;

typedef enum E_ELEMTYPE
{
	ET_TASK = 20,
	ET_RESOURCE = 21,
	ET_CALENDAR = 22,
	ET_ASSIGNMENT = 23,
}E_ELEMTYPE;

typedef enum E_ROLLUPTYPE
{
	RT_MAXIMUM_OR_FOR_FLAG_FIELDS = 0,
	RT_MINIMUM__AND__FOR_FLAG_FIELDS = 1,
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
	DT_SUNDAY = 1,
	DT_MONDAY = 2,
	DT_TUESDAY = 3,
	DT_WEDNESDAY = 4,
	DT_THURSDAY = 5,
	DT_FRIDAY = 6,
	DT_SATURDAY = 7,
}E_DAYTYPE;

typedef enum E_TYPE_3
{
	T_3_DAILY = 1,
	T_3_YEARLY_BY_DAY_OF_THE_MONTH = 2,
	T_3_YEARLY_BY_POSITION = 3,
	T_3_MONTHLY_BY_DAY_OF_THE_MONTH = 4,
	T_3_MONTHLY_BY_POSITION = 5,
	T_3_WEEKLY = 6,
	T_3_BY_DAY_COUNT = 7,
	T_3_BY_WEEKDAY_COUNT = 8,
	T_3_NO_EXCEPTION_TYPE = 9,
}E_TYPE_3;

typedef enum E_MONTHITEM
{
	MI_DAY = 0,
	MI_WEEKDAY = 1,
	MI_WEEKENDDAY = 2,
	MI_SUNDAY = 3,
	MI_MONDAY = 4,
	MI_TUESDAY = 5,
	MI_WEDNESDAY = 6,
	MI_THURSDAY = 7,
	MI_FRIDAY = 8,
	MI_SATURDAY = 9,
}E_MONTHITEM;

typedef enum E_MONTHPOSITION
{
	MP_FIRST_POSITION = 0,
	MP_SECOND_POSITION = 1,
	MP_THIRD_POSITION = 2,
	MP_FOURTH_POSITION = 3,
	MP_LAST_POSITION = 4,
}E_MONTHPOSITION;

typedef enum E_MONTH
{
	M_JANUARY = 0,
	M_FEBRUARY = 1,
	M_MARCH = 2,
	M_APRIL = 3,
	M_MAY = 4,
	M_JUNE = 5,
	M_JULY = 6,
	M_AUGUST = 7,
	M_SEPTEMBER = 8,
	M_OCTOBER = 9,
	M_NOVEMBER = 10,
	M_DECEMBER = 11,
}E_MONTH;

typedef enum E_TYPE_4
{
	T_4_FIXED_UNITS = 0,
	T_4_FIXED_DURATION = 1,
	T_4_FIXED_WORK = 2,
}E_TYPE_4;

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

typedef enum E_COMMITMENTTYPE
{
	CT_THE_TASK_HAS_NO_DELIVERABLE_OR_DEPENDENCY_ON_A_DELIVERABLE = 0,
	CT_THE_TASK_HAS_AN_ASSOCIATED_DELIVERABLE = 1,
	CT_THE_TASK_HAS_A_DEPENDENCY_ON_AN_ASSOCIATED_DELIVERABLE = 2,
}E_COMMITMENTTYPE;

typedef enum E_TYPE_5
{
	T_5_FF = 0,
	T_5_FS = 1,
	T_5_SF = 2,
	T_5_SS = 3,
}E_TYPE_5;

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
}E_LAGFORMAT;

typedef enum E_TYPE_6
{
	T_6_MATERIAL = 0,
	T_6_WORK = 1,
}E_TYPE_6;

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
	AA_INVALID = 4,
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

typedef enum E_TYPE_7
{
	T_7_ASSIGNMENT_REMAINING_WORK = 1,
	T_7_ASSIGNMENT_ACTUAL_WORK = 2,
	T_7_ASSIGNMENT_ACTUAL_OVERTIME_WORK = 3,
	T_7_ASSIGNMENT_BASELINE_WORK = 4,
	T_7_ASSIGNMENT_BASELINE_COST = 5,
	T_7_ASSIGNMENT_ACTUAL_COST = 6,
	T_7_RESOURCE_BASELINE_WORK = 7,
	T_7_RESOURCE_BASELINE_COST = 8,
	T_7_TASK_BASELINE_WORK = 9,
	T_7_TASK_BASELINE_COST = 10,
	T_7_TASK_PERCENT_COMPLETE = 11,
	T_7_ASSIGNMENT_BASELINE_1_WORK = 16,
	T_7_ASSIGNMENT_BASELINE_1_COST = 17,
	T_7_TASK_BASELINE_1_WORK = 18,
	T_7_TASK_BASELINE_1_COST = 19,
	T_7_RESOURCE_BASELINE_1_WORK = 20,
	T_7_RESOURCE_BASELINE_1_COST = 21,
	T_7_ASSIGNMENT_BASELINE_2_WORK = 22,
	T_7_ASSIGNMENT_BASELINE_2_COST = 23,
	T_7_TASK_BASELINE_2_WORK = 24,
	T_7_TASK_BASELINE_2_COST = 25,
	T_7_RESOURCE_BASELINE_2_WORK = 26,
	T_7_RESOURCE_BASELINE_2_COST = 27,
	T_7_ASSIGNMENT_BASELINE_3_WORK = 28,
	T_7_ASSIGNMENT_BASELINE_3_COST = 29,
	T_7_TASK_BASELINE_3_WORK = 30,
	T_7_TASK_BASELINE_3_COST = 31,
	T_7_RESOURCE_BASELINE_3_WORK = 32,
	T_7_RESOURCE_BASELINE_3_COST = 33,
	T_7_ASSIGNMENT_BASELINE_4_WORK = 34,
	T_7_ASSIGNMENT_BASELINE_4_COST = 35,
	T_7_TASK_BASELINE_4_WORK = 36,
	T_7_TASK_BASELINE_4_COST = 37,
	T_7_RESOURCE_BASELINE_4_WORK = 38,
	T_7_RESOURCE_BASELINE_4_COST = 39,
	T_7_ASSIGNMENT_BASELINE_5_WORK = 40,
	T_7_ASSIGNMENT_BASELINE_5_COST = 41,
	T_7_TASK_BASELINE_5_WORK = 42,
	T_7_TASK_BASELINE_5_COST = 43,
	T_7_RESOURCE_BASELINE_5_WORK = 44,
	T_7_RESOURCE_BASELINE_5_COST = 45,
	T_7_ASSIGNMENT_BASELINE_6_WORK = 46,
	T_7_ASSIGNMENT_BASELINE_6_COST = 47,
	T_7_TASK_BASELINE_6_WORK = 48,
	T_7_TASK_BASELINE_6_COST = 49,
	T_7_RESOURCE_BASELINE_6_WORK = 50,
	T_7_RESOURCE_BASELINE_6_COST = 51,
	T_7_ASSIGNMENT_BASELINE_7_WORK = 52,
	T_7_ASSIGNMENT_BASELINE_7_COST = 53,
	T_7_TASK_BASELINE_7_WORK = 54,
	T_7_TASK_BASELINE_7_COST = 55,
	T_7_RESOURCE_BASELINE_7_WORK = 56,
	T_7_RESOURCE_BASELINE_7_COST = 57,
	T_7_ASSIGNMENT_BASELINE_8_WORK = 58,
	T_7_ASSIGNMENT_BASELINE_8_COST = 59,
	T_7_TASK_BASELINE_8_WORK = 60,
	T_7_TASK_BASELINE_8_COST = 61,
	T_7_RESOURCE_BASELINE_8_WORK = 62,
	T_7_RESOURCE_BASELINE_8_COST = 63,
	T_7_ASSIGNMENT_BASELINE_9_WORK = 64,
	T_7_ASSIGNMENT_BASELINE_9_COST = 65,
	T_7_TASK_BASELINE_9_WORK = 66,
	T_7_TASK_BASELINE_9_COST = 67,
	T_7_RESOURCE_BASELINE_9_WORK = 68,
	T_7_RESOURCE_BASELINE_9_COST = 69,
	T_7_ASSIGNMENT_BASELINE_10_WORK = 70,
	T_7_ASSIGNMENT_BASELINE_10_COST = 71,
	T_7_TASK_BASELINE_10_WORK = 72,
	T_7_TASK_BASELINE_10_COST = 73,
	T_7_RESOURCE_BASELINE_10_WORK = 74,
	T_7_RESOURCE_BASELINE_10_COST = 75,
	T_7_PHYSICAL_PERCENT_COMPLETE = 76,
}E_TYPE_7;

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
	void SetsFieldID(LPCTSTR propval);
	CString GetsFieldID();
	void SetsValue(LPCTSTR propval);
	CString GetsValue();
	void SetlValueGUID(LONG propval);
	LONG GetlValueGUID();
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
	void SetfPeakUnits(FLOAT propval);
	FLOAT GetfPeakUnits();
	void SetlRateScale(LONG propval);
	LONG GetlRateScale();
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
	void SetbSummary(BOOL propval);
	BOOL GetbSummary();
	void SetfSV(FLOAT propval);
	FLOAT GetfSV();
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
	void SetsAssnOwner(LPCTSTR propval);
	CString GetsAssnOwner();
	void SetsAssnOwnerGuid(LPCTSTR propval);
	CString GetsAssnOwnerGuid();
	void SetcBudgetCost(DOUBLE propval);
	DOUBLE GetcBudgetCost();
	CDuration GetoBudgetWork();
	CAssignmentExtendedAttribute_C GetoExtendedAttribute_C();
	CAssignmentBaseline_C GetoBaseline_C();
	void Setsf404000(LPCTSTR propval);
	CString Getsf404000();
	void Setsf404001(LPCTSTR propval);
	CString Getsf404001();
	void Setsf404002(LPCTSTR propval);
	CString Getsf404002();
	void Setsf404003(LPCTSTR propval);
	CString Getsf404003();
	void Setsf404004(LPCTSTR propval);
	CString Getsf404004();
	void Setsf404005(LPCTSTR propval);
	CString Getsf404005();
	void Setsf404006(LPCTSTR propval);
	CString Getsf404006();
	void Setsf404007(LPCTSTR propval);
	CString Getsf404007();
	void Setsf404008(LPCTSTR propval);
	CString Getsf404008();
	void Setsf404009(LPCTSTR propval);
	CString Getsf404009();
	void Setsf40400a(LPCTSTR propval);
	CString Getsf40400a();
	void Setsf40400b(LPCTSTR propval);
	CString Getsf40400b();
	void Setsf40400c(LPCTSTR propval);
	CString Getsf40400c();
	void Setsf40400d(LPCTSTR propval);
	CString Getsf40400d();
	void Setsf40400e(LPCTSTR propval);
	CString Getsf40400e();
	void Setsf40400f(LPCTSTR propval);
	CString Getsf40400f();
	void Setsf404010(LPCTSTR propval);
	CString Getsf404010();
	void Setsf404011(LPCTSTR propval);
	CString Getsf404011();
	void Setsf404012(LPCTSTR propval);
	CString Getsf404012();
	void Setsf404013(LPCTSTR propval);
	CString Getsf404013();
	void Setsf404014(LPCTSTR propval);
	CString Getsf404014();
	void Setsf404015(LPCTSTR propval);
	CString Getsf404015();
	void Setsf404016(LPCTSTR propval);
	CString Getsf404016();
	void Setsf404017(LPCTSTR propval);
	CString Getsf404017();
	void Setsf404018(LPCTSTR propval);
	CString Getsf404018();
	void Setsf404019(LPCTSTR propval);
	CString Getsf404019();
	void Setsf40401a(LPCTSTR propval);
	CString Getsf40401a();
	void Setsf40401b(LPCTSTR propval);
	CString Getsf40401b();
	void Setsf40401c(LPCTSTR propval);
	CString Getsf40401c();
	void Setsf40401d(LPCTSTR propval);
	CString Getsf40401d();
	void Setsf40401e(LPCTSTR propval);
	CString Getsf40401e();
	void Setsf40401f(LPCTSTR propval);
	CString Getsf40401f();
	void Setsf404020(LPCTSTR propval);
	CString Getsf404020();
	void Setsf404021(LPCTSTR propval);
	CString Getsf404021();
	void Setsf404022(LPCTSTR propval);
	CString Getsf404022();
	void Setsf404023(LPCTSTR propval);
	CString Getsf404023();
	void Setsf404024(LPCTSTR propval);
	CString Getsf404024();
	void Setsf404025(LPCTSTR propval);
	CString Getsf404025();
	void Setsf404026(LPCTSTR propval);
	CString Getsf404026();
	void Setsf404027(LPCTSTR propval);
	CString Getsf404027();
	void Setsf404028(LPCTSTR propval);
	CString Getsf404028();
	void Setsf404029(LPCTSTR propval);
	CString Getsf404029();
	void Setsf40402a(LPCTSTR propval);
	CString Getsf40402a();
	void Setsf40402b(LPCTSTR propval);
	CString Getsf40402b();
	void Setsf40402c(LPCTSTR propval);
	CString Getsf40402c();
	void Setsf40402d(LPCTSTR propval);
	CString Getsf40402d();
	void Setsf40402e(LPCTSTR propval);
	CString Getsf40402e();
	void Setsf40402f(LPCTSTR propval);
	CString Getsf40402f();
	void Setsf404030(LPCTSTR propval);
	CString Getsf404030();
	void Setsf404031(LPCTSTR propval);
	CString Getsf404031();
	void Setsf404032(LPCTSTR propval);
	CString Getsf404032();
	void Setsf404033(LPCTSTR propval);
	CString Getsf404033();
	void Setsf404034(LPCTSTR propval);
	CString Getsf404034();
	void Setsf404035(LPCTSTR propval);
	CString Getsf404035();
	void Setsf404036(LPCTSTR propval);
	CString Getsf404036();
	void Setsf404037(LPCTSTR propval);
	CString Getsf404037();
	void Setsf404038(LPCTSTR propval);
	CString Getsf404038();
	void Setsf404039(LPCTSTR propval);
	CString Getsf404039();
	void Setsf40403a(LPCTSTR propval);
	CString Getsf40403a();
	void Setsf40403b(LPCTSTR propval);
	CString Getsf40403b();
	void Setsf40403c(LPCTSTR propval);
	CString Getsf40403c();
	void Setsf40403d(LPCTSTR propval);
	CString Getsf40403d();
	void Setsf40403e(LPCTSTR propval);
	CString Getsf40403e();
	void Setsf40403f(LPCTSTR propval);
	CString Getsf40403f();
	void Setsf404040(LPCTSTR propval);
	CString Getsf404040();
	void Setsf404041(LPCTSTR propval);
	CString Getsf404041();
	void Setsf404042(LPCTSTR propval);
	CString Getsf404042();
	void Setsf404043(LPCTSTR propval);
	CString Getsf404043();
	void Setsf404044(LPCTSTR propval);
	CString Getsf404044();
	void Setsf404045(LPCTSTR propval);
	CString Getsf404045();
	void Setsf404046(LPCTSTR propval);
	CString Getsf404046();
	void Setsf404047(LPCTSTR propval);
	CString Getsf404047();
	void Setsf404048(LPCTSTR propval);
	CString Getsf404048();
	void Setsf404049(LPCTSTR propval);
	CString Getsf404049();
	void Setsf40404a(LPCTSTR propval);
	CString Getsf40404a();
	void Setsf40404b(LPCTSTR propval);
	CString Getsf40404b();
	void Setsf40404c(LPCTSTR propval);
	CString Getsf40404c();
	void Setsf40404d(LPCTSTR propval);
	CString Getsf40404d();
	void Setsf40404e(LPCTSTR propval);
	CString Getsf40404e();
	void Setsf40404f(LPCTSTR propval);
	CString Getsf40404f();
	void Setsf404050(LPCTSTR propval);
	CString Getsf404050();
	void Setsf404051(LPCTSTR propval);
	CString Getsf404051();
	void Setsf404052(LPCTSTR propval);
	CString Getsf404052();
	void Setsf404053(LPCTSTR propval);
	CString Getsf404053();
	void Setsf404054(LPCTSTR propval);
	CString Getsf404054();
	void Setsf404055(LPCTSTR propval);
	CString Getsf404055();
	void Setsf404056(LPCTSTR propval);
	CString Getsf404056();
	void Setsf404057(LPCTSTR propval);
	CString Getsf404057();
	void Setsf404058(LPCTSTR propval);
	CString Getsf404058();
	void Setsf404059(LPCTSTR propval);
	CString Getsf404059();
	void Setsf40405a(LPCTSTR propval);
	CString Getsf40405a();
	void Setsf40405b(LPCTSTR propval);
	CString Getsf40405b();
	void Setsf40405c(LPCTSTR propval);
	CString Getsf40405c();
	void Setsf40405d(LPCTSTR propval);
	CString Getsf40405d();
	void Setsf40405e(LPCTSTR propval);
	CString Getsf40405e();
	void Setsf40405f(LPCTSTR propval);
	CString Getsf40405f();
	void Setsf404060(LPCTSTR propval);
	CString Getsf404060();
	void Setsf404061(LPCTSTR propval);
	CString Getsf404061();
	void Setsf404062(LPCTSTR propval);
	CString Getsf404062();
	void Setsf404063(LPCTSTR propval);
	CString Getsf404063();
	void Setsf404064(LPCTSTR propval);
	CString Getsf404064();
	void Setsf404065(LPCTSTR propval);
	CString Getsf404065();
	void Setsf404066(LPCTSTR propval);
	CString Getsf404066();
	void Setsf404067(LPCTSTR propval);
	CString Getsf404067();
	void Setsf404068(LPCTSTR propval);
	CString Getsf404068();
	void Setsf404069(LPCTSTR propval);
	CString Getsf404069();
	void Setsf40406a(LPCTSTR propval);
	CString Getsf40406a();
	void Setsf40406b(LPCTSTR propval);
	CString Getsf40406b();
	void Setsf40406c(LPCTSTR propval);
	CString Getsf40406c();
	void Setsf40406d(LPCTSTR propval);
	CString Getsf40406d();
	void Setsf40406e(LPCTSTR propval);
	CString Getsf40406e();
	void Setsf40406f(LPCTSTR propval);
	CString Getsf40406f();
	void Setsf404070(LPCTSTR propval);
	CString Getsf404070();
	void Setsf404071(LPCTSTR propval);
	CString Getsf404071();
	void Setsf404072(LPCTSTR propval);
	CString Getsf404072();
	void Setsf404073(LPCTSTR propval);
	CString Getsf404073();
	void Setsf404074(LPCTSTR propval);
	CString Getsf404074();
	void Setsf404075(LPCTSTR propval);
	CString Getsf404075();
	void Setsf404076(LPCTSTR propval);
	CString Getsf404076();
	void Setsf404077(LPCTSTR propval);
	CString Getsf404077();
	void Setsf404078(LPCTSTR propval);
	CString Getsf404078();
	void Setsf404079(LPCTSTR propval);
	CString Getsf404079();
	void Setsf40407a(LPCTSTR propval);
	CString Getsf40407a();
	void Setsf40407b(LPCTSTR propval);
	CString Getsf40407b();
	void Setsf40407c(LPCTSTR propval);
	CString Getsf40407c();
	void Setsf40407d(LPCTSTR propval);
	CString Getsf40407d();
	void Setsf40407e(LPCTSTR propval);
	CString Getsf40407e();
	void Setsf40407f(LPCTSTR propval);
	CString Getsf40407f();
	void Setsf404080(LPCTSTR propval);
	CString Getsf404080();
	void Setsf404081(LPCTSTR propval);
	CString Getsf404081();
	void Setsf404082(LPCTSTR propval);
	CString Getsf404082();
	void Setsf404083(LPCTSTR propval);
	CString Getsf404083();
	void Setsf404084(LPCTSTR propval);
	CString Getsf404084();
	void Setsf404085(LPCTSTR propval);
	CString Getsf404085();
	void Setsf404086(LPCTSTR propval);
	CString Getsf404086();
	void Setsf404087(LPCTSTR propval);
	CString Getsf404087();
	void Setsf404088(LPCTSTR propval);
	CString Getsf404088();
	void Setsf404089(LPCTSTR propval);
	CString Getsf404089();
	void Setsf40408a(LPCTSTR propval);
	CString Getsf40408a();
	void Setsf40408b(LPCTSTR propval);
	CString Getsf40408b();
	void Setsf40408c(LPCTSTR propval);
	CString Getsf40408c();
	void Setsf40408d(LPCTSTR propval);
	CString Getsf40408d();
	void Setsf40408e(LPCTSTR propval);
	CString Getsf40408e();
	void Setsf40408f(LPCTSTR propval);
	CString Getsf40408f();
	void Setsf404090(LPCTSTR propval);
	CString Getsf404090();
	void Setsf404091(LPCTSTR propval);
	CString Getsf404091();
	void Setsf404092(LPCTSTR propval);
	CString Getsf404092();
	void Setsf404093(LPCTSTR propval);
	CString Getsf404093();
	void Setsf404094(LPCTSTR propval);
	CString Getsf404094();
	void Setsf404095(LPCTSTR propval);
	CString Getsf404095();
	void Setsf404096(LPCTSTR propval);
	CString Getsf404096();
	void Setsf404097(LPCTSTR propval);
	CString Getsf404097();
	void Setsf404098(LPCTSTR propval);
	CString Getsf404098();
	void Setsf404099(LPCTSTR propval);
	CString Getsf404099();
	void Setsf40409a(LPCTSTR propval);
	CString Getsf40409a();
	void Setsf40409b(LPCTSTR propval);
	CString Getsf40409b();
	void Setsf40409c(LPCTSTR propval);
	CString Getsf40409c();
	void Setsf40409d(LPCTSTR propval);
	CString Getsf40409d();
	void Setsf40409e(LPCTSTR propval);
	CString Getsf40409e();
	void Setsf40409f(LPCTSTR propval);
	CString Getsf40409f();
	void Setsf4040a0(LPCTSTR propval);
	CString Getsf4040a0();
	void Setsf4040a1(LPCTSTR propval);
	CString Getsf4040a1();
	void Setsf4040a2(LPCTSTR propval);
	CString Getsf4040a2();
	void Setsf4040a3(LPCTSTR propval);
	CString Getsf4040a3();
	void Setsf4040a4(LPCTSTR propval);
	CString Getsf4040a4();
	void Setsf4040a5(LPCTSTR propval);
	CString Getsf4040a5();
	void Setsf4040a6(LPCTSTR propval);
	CString Getsf4040a6();
	void Setsf4040a7(LPCTSTR propval);
	CString Getsf4040a7();
	void Setsf4040a8(LPCTSTR propval);
	CString Getsf4040a8();
	void Setsf4040a9(LPCTSTR propval);
	CString Getsf4040a9();
	void Setsf4040aa(LPCTSTR propval);
	CString Getsf4040aa();
	void Setsf4040ab(LPCTSTR propval);
	CString Getsf4040ab();
	void Setsf4040ac(LPCTSTR propval);
	CString Getsf4040ac();
	void Setsf4040ad(LPCTSTR propval);
	CString Getsf4040ad();
	void Setsf4040ae(LPCTSTR propval);
	CString Getsf4040ae();
	void Setsf4040af(LPCTSTR propval);
	CString Getsf4040af();
	void Setsf4040b0(LPCTSTR propval);
	CString Getsf4040b0();
	void Setsf4040b1(LPCTSTR propval);
	CString Getsf4040b1();
	void Setsf4040b2(LPCTSTR propval);
	CString Getsf4040b2();
	void Setsf4040b3(LPCTSTR propval);
	CString Getsf4040b3();
	void Setsf4040b4(LPCTSTR propval);
	CString Getsf4040b4();
	void Setsf4040b5(LPCTSTR propval);
	CString Getsf4040b5();
	void Setsf4040b6(LPCTSTR propval);
	CString Getsf4040b6();
	void Setsf4040b7(LPCTSTR propval);
	CString Getsf4040b7();
	void Setsf4040b8(LPCTSTR propval);
	CString Getsf4040b8();
	void Setsf4040b9(LPCTSTR propval);
	CString Getsf4040b9();
	void Setsf4040ba(LPCTSTR propval);
	CString Getsf4040ba();
	void Setsf4040bb(LPCTSTR propval);
	CString Getsf4040bb();
	void Setsf4040bc(LPCTSTR propval);
	CString Getsf4040bc();
	void Setsf4040bd(LPCTSTR propval);
	CString Getsf4040bd();
	void Setsf4040be(LPCTSTR propval);
	CString Getsf4040be();
	void Setsf4040bf(LPCTSTR propval);
	CString Getsf4040bf();
	void Setsf4040c0(LPCTSTR propval);
	CString Getsf4040c0();
	void Setsf4040c1(LPCTSTR propval);
	CString Getsf4040c1();
	void Setsf4040c2(LPCTSTR propval);
	CString Getsf4040c2();
	void Setsf4040c3(LPCTSTR propval);
	CString Getsf4040c3();
	void Setsf4040c4(LPCTSTR propval);
	CString Getsf4040c4();
	void Setsf4040c5(LPCTSTR propval);
	CString Getsf4040c5();
	void Setsf4040c6(LPCTSTR propval);
	CString Getsf4040c6();
	void Setsf4040c7(LPCTSTR propval);
	CString Getsf4040c7();
	void Setsf4040c8(LPCTSTR propval);
	CString Getsf4040c8();
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
	void SetsFieldID(LPCTSTR propval);
	CString GetsFieldID();
	void SetlValueID(LONG propval);
	LONG GetlValueID();
	void SetlValueGUID(LONG propval);
	LONG GetlValueGUID();
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
	void SetsFieldID(LPCTSTR propval);
	CString GetsFieldID();
	void SetsValue(LPCTSTR propval);
	CString GetsValue();
	void SetlValueGUID(LONG propval);
	LONG GetlValueGUID();
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
	void SetbIsCostResource(BOOL propval);
	BOOL GetbIsCostResource();
	void SetsAssnOwner(LPCTSTR propval);
	CString GetsAssnOwner();
	void SetsAssnOwnerGuid(LPCTSTR propval);
	CString GetsAssnOwnerGuid();
	void SetbIsBudget(BOOL propval);
	BOOL GetbIsBudget();
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
	void SetsFieldID(LPCTSTR propval);
	CString GetsFieldID();
	void SetlValueID(LONG propval);
	LONG GetlValueID();
	void SetlValueGUID(LONG propval);
	LONG GetlValueGUID();
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
	void SetfFixedCost(FLOAT propval);
	FLOAT GetfFixedCost();
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
	void SetsFieldID(LPCTSTR propval);
	CString GetsFieldID();
	void SetsValue(LPCTSTR propval);
	CString GetsValue();
	void SetlValueGUID(LONG propval);
	LONG GetlValueGUID();
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
	void SetbDisplayAsSummary(BOOL propval);
	BOOL GetbDisplayAsSummary();
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
	void SetlStartSlack(LONG propval);
	LONG GetlStartSlack();
	void SetlFinishSlack(LONG propval);
	LONG GetlFinishSlack();
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
	void SetbIsPublished(BOOL propval);
	BOOL GetbIsPublished();
	void SetsStatusManager(LPCTSTR propval);
	CString GetsStatusManager();
	void SetdtCommitmentStart(DATE propval);
	DATE GetdtCommitmentStart();
	void SetdtCommitmentFinish(DATE propval);
	DATE GetdtCommitmentFinish();
	void SetyCommitmentType(LONG propval);
	LONG GetyCommitmentType();
	void SetbActive(BOOL propval);
	BOOL GetbActive();
	void SetbPinned(BOOL propval);
	BOOL GetbPinned();
	void SetsPinnedStart(LPCTSTR propval);
	CString GetsPinnedStart();
	void SetsPinnedFinish(LPCTSTR propval);
	CString GetsPinnedFinish();
	void SetsPinnedDuration(LPCTSTR propval);
	CString GetsPinnedDuration();
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
	void SetKey(LPCTSTR propval);
	CString GetKey();

// Methods
public:
	CString GetXML(void);
	void SetXML(LPCTSTR sXML);
	BOOL IsNull(void);
	void Initialize(void);
};

class CWeekDay_C : public COleDispatchDriver
{
// Constructors
public:
	CWeekDay_C() {}	
	CWeekDay_C(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CWeekDay_C(const CWeekDay_C& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

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
};

class CCalendarWorkWeek : public COleDispatchDriver
{
// Constructors
public:
	CCalendarWorkWeek() {}	
	CCalendarWorkWeek(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CCalendarWorkWeek(const CCalendarWorkWeek& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Properties
public:
	CTimePeriod GetoTimePeriod();
	void SetsName(LPCTSTR propval);
	CString GetsName();
	CWeekDay_C GetoWeekDay_C();
	void SetKey(LPCTSTR propval);
	CString GetKey();

// Methods
public:
	CString GetXML(void);
	void SetXML(LPCTSTR sXML);
	BOOL IsNull(void);
	void Initialize(void);
};

class CCalendarWorkWeeks : public COleDispatchDriver
{
// Constructors
public:
	CCalendarWorkWeeks() {}	
	CCalendarWorkWeeks(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CCalendarWorkWeeks(const CCalendarWorkWeeks& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Properties
public:
	LONG GetCount();

// Methods
public:
	CCalendarWorkWeek Item(LPCTSTR Index);
	CCalendarWorkWeek Add(void);
	void Clear(void);
	void Remove(LPCTSTR Index);
	BOOL IsNull(void);
	void Initialize(void);
	void SetXML(LPCTSTR sXML);
	CString GetXML(void);
};

class CCalendarException : public COleDispatchDriver
{
// Constructors
public:
	CCalendarException() {}	
	CCalendarException(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CCalendarException(const CCalendarException& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Properties
public:
	void SetbEnteredByOccurrences(BOOL propval);
	BOOL GetbEnteredByOccurrences();
	CTimePeriod GetoTimePeriod();
	void SetlOccurrences(LONG propval);
	LONG GetlOccurrences();
	void SetsName(LPCTSTR propval);
	CString GetsName();
	void SetyType(LONG propval);
	LONG GetyType();
	void SetlPeriod(LONG propval);
	LONG GetlPeriod();
	void SetlDaysOfWeek(LONG propval);
	LONG GetlDaysOfWeek();
	void SetyMonthItem(LONG propval);
	LONG GetyMonthItem();
	void SetyMonthPosition(LONG propval);
	LONG GetyMonthPosition();
	void SetyMonth(LONG propval);
	LONG GetyMonth();
	void SetlMonthDay(LONG propval);
	LONG GetlMonthDay();
	void SetbDayWorking(BOOL propval);
	BOOL GetbDayWorking();
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

class CCalendarExceptions : public COleDispatchDriver
{
// Constructors
public:
	CCalendarExceptions() {}	
	CCalendarExceptions(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CCalendarExceptions(const CCalendarExceptions& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Properties
public:
	LONG GetCount();

// Methods
public:
	CCalendarException Item(LPCTSTR Index);
	CCalendarException Add(void);
	void Clear(void);
	void Remove(LPCTSTR Index);
	BOOL IsNull(void);
	void Initialize(void);
	void SetXML(LPCTSTR sXML);
	CString GetXML(void);
};

class CCalendarWeekDay : public COleDispatchDriver
{
// Constructors
public:
	CCalendarWeekDay() {}	
	CCalendarWeekDay(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CCalendarWeekDay(const CCalendarWeekDay& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

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

class CCalendarWeekDays : public COleDispatchDriver
{
// Constructors
public:
	CCalendarWeekDays() {}	
	CCalendarWeekDays(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CCalendarWeekDays(const CCalendarWeekDays& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Properties
public:
	LONG GetCount();

// Methods
public:
	CCalendarWeekDay Item(LPCTSTR Index);
	CCalendarWeekDay Add(void);
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
	void SetbIsBaselineCalendar(BOOL propval);
	BOOL GetbIsBaselineCalendar();
	void SetlBaseCalendarUID(LONG propval);
	LONG GetlBaseCalendarUID();
	CCalendarWeekDays GetoWeekDays();
	CCalendarExceptions GetoExceptions();
	CCalendarWorkWeeks GetoWorkWeeks();
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
	void SetsPhonetic(LPCTSTR propval);
	CString GetsPhonetic();
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
	void SetyCFType(LONG propval);
	LONG GetyCFType();
	void SetsGuid(LPCTSTR propval);
	CString GetsGuid();
	void SetyElemType(LONG propval);
	LONG GetyElemType();
	void SetlMaxMultiValues(LONG propval);
	LONG GetlMaxMultiValues();
	void SetbUserDef(BOOL propval);
	BOOL GetbUserDef();
	void SetsAlias(LPCTSTR propval);
	CString GetsAlias();
	void SetsSecondaryPID(LPCTSTR propval);
	CString GetsSecondaryPID();
	void SetbAutoRollDown(BOOL propval);
	BOOL GetbAutoRollDown();
	void SetsDefaultGuid(LPCTSTR propval);
	CString GetsDefaultGuid();
	void SetsLtuid(LPCTSTR propval);
	CString GetsLtuid();
	void SetsSecondaryGuid(LPCTSTR propval);
	CString GetsSecondaryGuid();
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
	void SetsFieldGUID(LPCTSTR propval);
	CString GetsFieldGUID();
	void SetyType(LONG propval);
	LONG GetyType();
	void SetbIsCollapsed(BOOL propval);
	BOOL GetbIsCollapsed();
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
	void SetsGuid(LPCTSTR propval);
	CString GetsGuid();
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
	void SetbShowIndent(BOOL propval);
	BOOL GetbShowIndent();
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

class CMP14 : public COleDispatchDriver
{
// Constructors
public:
	CMP14() {}	
	CMP14(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CMP14(const CMP14& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Properties
public:
	void SetlSaveVersion(LONG propval);
	LONG GetlSaveVersion();
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
	void SetsCurrencyCode(LPCTSTR propval);
	CString GetsCurrencyCode();
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
	void SetsBaslineCalendar(LPCTSTR propval);
	CString GetsBaslineCalendar();
	void SetbNewTasksAreManual(BOOL propval);
	BOOL GetbNewTasksAreManual();
	void SetbUpdateManuallyScheduledTasksWhenEditingLinks(BOOL propval);
	BOOL GetbUpdateManuallyScheduledTasksWhenEditingLinks();
	void SetbKeepTaskOnNearestWorkingTimeWhenMadeAutoScheduled(BOOL propval);
	BOOL GetbKeepTaskOnNearestWorkingTimeWhenMadeAutoScheduled();
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

class CMSP2010Ctl : public CWnd
{
protected:
	DECLARE_DYNCREATE(CMSP2010Ctl)
public:
	CLSID const& GetClsid()
	{
		static CLSID const clsid = { 0xC56E2638, 0x476C, 0x44C6, { 0xB2, 0x1F, 0x91, 0xFD, 0x43, 0x48, 0xB8, 0xA5} };
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
	CMP14 GetMP14();

// Methods
public:
	void AboutBox(void);
	void Clear(void);
};
}