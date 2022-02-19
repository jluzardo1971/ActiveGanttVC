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



class CalendarException : public clsItemBase
{
	DECLARE_DYNCREATE(CalendarException)

public:

	CalendarException();
	virtual ~CalendarException();
	virtual void OnFinalRelease();

	clsCollectionBase* mp_oCollection;
	BOOL mp_bEnteredByOccurrences;
	TimePeriod* mp_oTimePeriod;
	LONG mp_lOccurrences;
	CString mp_sName;
	LONG mp_yType;
	LONG mp_lPeriod;
	LONG mp_lDaysOfWeek;
	LONG mp_yMonthItem;
	LONG mp_yMonthPosition;
	LONG mp_yMonth;
	LONG mp_lMonthDay;
	BOOL mp_bDayWorking;
	WorkingTimes* mp_oWorkingTimes;

	BOOL odl_GetEnteredByOccurrences(void);
	BOOL GetEnteredByOccurrences(void);
	void odl_SetEnteredByOccurrences(BOOL newVal);
	void SetEnteredByOccurrences(BOOL newVal);

	IDispatch* odl_GetTimePeriod(void);
	TimePeriod* GetTimePeriod(void);

	LONG odl_GetOccurrences(void);
	LONG GetOccurrences(void);
	void odl_SetOccurrences(LONG newVal);
	void SetOccurrences(LONG newVal);

	BSTR odl_GetName(void);
	CString GetName(void);
	void odl_SetName(LPCTSTR newVal);
	void SetName(CString newVal);

	LONG odl_GetType(void);
	E_TYPE_3 GetType(void);
	void odl_SetType(LONG newVal);
	void SetType(E_TYPE_3 newVal);

	LONG odl_GetPeriod(void);
	LONG GetPeriod(void);
	void odl_SetPeriod(LONG newVal);
	void SetPeriod(LONG newVal);

	LONG odl_GetDaysOfWeek(void);
	LONG GetDaysOfWeek(void);
	void odl_SetDaysOfWeek(LONG newVal);
	void SetDaysOfWeek(LONG newVal);

	LONG odl_GetMonthItem(void);
	E_MONTHITEM GetMonthItem(void);
	void odl_SetMonthItem(LONG newVal);
	void SetMonthItem(E_MONTHITEM newVal);

	LONG odl_GetMonthPosition(void);
	E_MONTHPOSITION GetMonthPosition(void);
	void odl_SetMonthPosition(LONG newVal);
	void SetMonthPosition(E_MONTHPOSITION newVal);

	LONG odl_GetMonth(void);
	E_MONTH GetMonth(void);
	void odl_SetMonth(LONG newVal);
	void SetMonth(E_MONTH newVal);

	LONG odl_GetMonthDay(void);
	LONG GetMonthDay(void);
	void odl_SetMonthDay(LONG newVal);
	void SetMonthDay(LONG newVal);

	BOOL odl_GetDayWorking(void);
	BOOL GetDayWorking(void);
	void odl_SetDayWorking(BOOL newVal);
	void SetDayWorking(BOOL newVal);

	IDispatch* odl_GetWorkingTimes(void);
	WorkingTimes* GetWorkingTimes(void);

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
	DECLARE_OLECREATE(CalendarException)
	DECLARE_INTERFACE_MAP()

};