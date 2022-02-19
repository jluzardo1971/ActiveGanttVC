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


class CalendarWorkWeeks : public CCmdTarget
{
	DECLARE_DYNCREATE(CalendarWorkWeeks)

public:

	CalendarWorkWeeks();
	virtual ~CalendarWorkWeeks();
	virtual void OnFinalRelease();

	clsCollectionBase* mp_oCollection;

	LONG odl_GetCount(void);
	LONG GetCount(void);

	IDispatch* odl_Item(LPCTSTR Index);
	CalendarWorkWeek* Item(CString Index);

	IDispatch* odl_Add(void);
	CalendarWorkWeek* Add(void);

	void odl_Clear(void);
	void Clear(void);

	void odl_Remove(LPCTSTR Index);
	void Remove(CString Index);

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
	DECLARE_OLECREATE(CalendarWorkWeeks)
	DECLARE_INTERFACE_MAP()

};