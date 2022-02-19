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



class CalendarWorkWeek : public clsItemBase
{
	DECLARE_DYNCREATE(CalendarWorkWeek)

public:

	CalendarWorkWeek();
	virtual ~CalendarWorkWeek();
	virtual void OnFinalRelease();

	clsCollectionBase* mp_oCollection;
	TimePeriod* mp_oTimePeriod;
	CString mp_sName;
	WeekDay_C* mp_oWeekDay_C;

	IDispatch* odl_GetTimePeriod(void);
	TimePeriod* GetTimePeriod(void);

	BSTR odl_GetName(void);
	CString GetName(void);
	void odl_SetName(LPCTSTR newVal);
	void SetName(CString newVal);

	IDispatch* odl_GetWeekDay(void);
	WeekDay_C* GetWeekDay(void);

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
	DECLARE_OLECREATE(CalendarWorkWeek)
	DECLARE_INTERFACE_MAP()

};