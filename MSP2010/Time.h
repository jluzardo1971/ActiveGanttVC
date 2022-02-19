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

class Time : public CCmdTarget
{
	DECLARE_DYNCREATE(Time)

public:
	Time();
	virtual ~Time();
	virtual void OnFinalRelease();

	LONG mp_lHour;
	LONG mp_lMinute;
	LONG mp_lSecond;

	LONG GetHour();
	void SetHour(LONG nNewValue);
	LONG GetMinute();
	void SetMinute(LONG nNewValue);
	LONG GetSecond();
	void SetSecond(LONG nNewValue);

	CString ToString();
	void FromString(CString sString);
	COleDateTime ToDateTime();
	void FromDateTime(COleDateTime dtDate);
	BOOL IsNull(void);

	BSTR odl_ToString(void);
	void odl_FromString(LPCTSTR sString);


protected:
	DECLARE_DISPATCH_MAP()
	DECLARE_MESSAGE_MAP()
	DECLARE_OLECREATE(Time)
	DECLARE_INTERFACE_MAP()
};