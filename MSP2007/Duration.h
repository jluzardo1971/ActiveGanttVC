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

class Duration : public CCmdTarget
{
	DECLARE_DYNCREATE(Duration)

public:
	Duration();
	virtual ~Duration();
	virtual void OnFinalRelease();

	LONG mp_lYear;
	LONG mp_lMonth;
	LONG mp_lDay;
	LONG mp_lHour;
	LONG mp_lMinute;
	LONG mp_lSecond;

	BOOL IsNull(void);
	CString ToString(void);
	void FromString(CString sString);
	LONG GetYear(void);
	void SetYear(LONG nNewValue);
	LONG GetMonth(void);
	void SetMonth(LONG nNewValue);
	LONG GetDay(void);
	void SetDay(LONG nNewValue);
	LONG GetHour(void);
	void SetHour(LONG nNewValue);
	LONG GetMinute(void);
	void SetMinute(LONG nNewValue);
	LONG GetSecond(void);
	void SetSecond(LONG nNewValue);

	BSTR odl_ToString(void);
	void odl_FromString(LPCTSTR sString);

protected:
	DECLARE_DISPATCH_MAP()
	DECLARE_MESSAGE_MAP()
	DECLARE_OLECREATE(Duration)
	DECLARE_INTERFACE_MAP()
};