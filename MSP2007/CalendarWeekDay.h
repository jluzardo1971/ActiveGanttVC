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



class CalendarWeekDay : public clsItemBase
{
	DECLARE_DYNCREATE(CalendarWeekDay)

public:

	CalendarWeekDay();
	virtual ~CalendarWeekDay();
	virtual void OnFinalRelease();

	clsCollectionBase* mp_oCollection;
	LONG mp_yDayType;
	BOOL mp_bDayWorking;
	TimePeriod* mp_oTimePeriod;
	WorkingTimes* mp_oWorkingTimes;

	LONG odl_GetDayType(void);
	E_DAYTYPE GetDayType(void);
	void odl_SetDayType(LONG newVal);
	void SetDayType(E_DAYTYPE newVal);

	BOOL odl_GetDayWorking(void);
	BOOL GetDayWorking(void);
	void odl_SetDayWorking(BOOL newVal);
	void SetDayWorking(BOOL newVal);

	IDispatch* odl_GetTimePeriod(void);
	TimePeriod* GetTimePeriod(void);

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
	DECLARE_OLECREATE(CalendarWeekDay)
	DECLARE_INTERFACE_MAP()

};