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



class Calendar : public clsItemBase
{
	DECLARE_DYNCREATE(Calendar)

public:

	Calendar();
	virtual ~Calendar();
	virtual void OnFinalRelease();

	clsCollectionBase* mp_oCollection;
	LONG mp_lUID;
	CString mp_sName;
	BOOL mp_bIsBaseCalendar;
	BOOL mp_bIsBaselineCalendar;
	LONG mp_lBaseCalendarUID;
	CalendarWeekDays* mp_oWeekDays;
	CalendarExceptions* mp_oExceptions;
	CalendarWorkWeeks* mp_oWorkWeeks;

	LONG odl_GetUID(void);
	LONG GetUID(void);
	void odl_SetUID(LONG newVal);
	void SetUID(LONG newVal);

	BSTR odl_GetName(void);
	CString GetName(void);
	void odl_SetName(LPCTSTR newVal);
	void SetName(CString newVal);

	BOOL odl_GetIsBaseCalendar(void);
	BOOL GetIsBaseCalendar(void);
	void odl_SetIsBaseCalendar(BOOL newVal);
	void SetIsBaseCalendar(BOOL newVal);

	BOOL odl_GetIsBaselineCalendar(void);
	BOOL GetIsBaselineCalendar(void);
	void odl_SetIsBaselineCalendar(BOOL newVal);
	void SetIsBaselineCalendar(BOOL newVal);

	LONG odl_GetBaseCalendarUID(void);
	LONG GetBaseCalendarUID(void);
	void odl_SetBaseCalendarUID(LONG newVal);
	void SetBaseCalendarUID(LONG newVal);

	IDispatch* odl_GetWeekDays(void);
	CalendarWeekDays* GetWeekDays(void);

	IDispatch* odl_GetExceptions(void);
	CalendarExceptions* GetExceptions(void);

	IDispatch* odl_GetWorkWeeks(void);
	CalendarWorkWeeks* GetWorkWeeks(void);

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
	DECLARE_OLECREATE(Calendar)
	DECLARE_INTERFACE_MAP()

};