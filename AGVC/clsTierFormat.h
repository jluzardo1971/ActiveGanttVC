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


class clsTierFormat : public CCmdTarget
{
	DECLARE_DYNCREATE(clsTierFormat)
	clsTierFormat();

public:
	CActiveGanttVCCtl* mp_oControl;

	clsTierFormat(CActiveGanttVCCtl* oControl);
	virtual ~clsTierFormat();
	virtual void OnFinalRelease();

// Member Variables
public:
	CString mp_sMicrosecondIntervalFormat;
	CString mp_sMillisecondIntervalFormat;
	CString mp_sSecondIntervalFormat;
	CString mp_sMinuteIntervalFormat;
	CString mp_sHourIntervalFormat;
	CString mp_sDayIntervalFormat;
	CString mp_sDayOfWeekIntervalFormat;
	CString mp_sDayOfYearIntervalFormat;
	CString mp_sWeekIntervalFormat;
	CString mp_sMonthIntervalFormat;
	CString mp_sQuarterIntervalFormat;
	CString mp_sYearIntervalFormat;

//Internal Properties (Extensions to ODL)
public:
	CString GetMicrosecondIntervalFormat(void);
	void SetMicrosecondIntervalFormat(CString Value);
	CString GetMillisecondIntervalFormat(void);
	void SetMillisecondIntervalFormat(CString Value);
	CString GetSecondIntervalFormat(void);
	void SetSecondIntervalFormat(CString Value);
	CString GetMinuteIntervalFormat(void);
	void SetMinuteIntervalFormat(CString Value);
	CString GetHourIntervalFormat(void);
	void SetHourIntervalFormat(CString Value);
	CString GetDayIntervalFormat(void);
	void SetDayIntervalFormat(CString Value);
	CString GetDayOfWeekIntervalFormat(void);
	void SetDayOfWeekIntervalFormat(CString Value);
	CString GetDayOfYearIntervalFormat(void);
	void SetDayOfYearIntervalFormat(CString Value);
	CString GetWeekIntervalFormat(void);
	void SetWeekIntervalFormat(CString Value);
	CString GetMonthIntervalFormat(void);
	void SetMonthIntervalFormat(CString Value);
	CString GetQuarterIntervalFormat(void);
	void SetQuarterIntervalFormat(CString Value);
	CString GetYearIntervalFormat(void);
	void SetYearIntervalFormat(CString Value);

//Internal Properties
public:

//Internal Methods
public:
	void Finalize(void);
	CString GetXML(void);
	void SetXML(CString sXML);
	void Clear();
	void Clone(clsTierFormat* oClone);

protected:
	DECLARE_DISPATCH_MAP()
	DECLARE_MESSAGE_MAP()
	DECLARE_OLECREATE(clsTierFormat)
	DECLARE_INTERFACE_MAP()

BSTR odl_GetMicrosecondIntervalFormat(void)
{
	
	return GetMicrosecondIntervalFormat().AllocSysString();
}

void odl_SetMicrosecondIntervalFormat(LPCTSTR Value)
{
	
	SetMicrosecondIntervalFormat(Value);
}

BSTR odl_GetMillisecondIntervalFormat(void)
{
	
	return GetMillisecondIntervalFormat().AllocSysString();
}

void odl_SetMillisecondIntervalFormat(LPCTSTR Value)
{
	
	SetMillisecondIntervalFormat(Value);
}

BSTR odl_GetSecondIntervalFormat(void)
{
	
	return GetSecondIntervalFormat().AllocSysString();
}

void odl_SetSecondIntervalFormat(LPCTSTR Value)
{
	
	SetSecondIntervalFormat(Value);
}

BSTR odl_GetMinuteIntervalFormat(void)
{
	
	return GetMinuteIntervalFormat().AllocSysString();
}

void odl_SetMinuteIntervalFormat(LPCTSTR Value)
{
	
	SetMinuteIntervalFormat(Value);
}

BSTR odl_GetHourIntervalFormat(void)
{
	
	return GetHourIntervalFormat().AllocSysString();
}

void odl_SetHourIntervalFormat(LPCTSTR Value)
{
	
	SetHourIntervalFormat(Value);
}

BSTR odl_GetDayIntervalFormat(void)
{
	
	return GetDayIntervalFormat().AllocSysString();
}

void odl_SetDayIntervalFormat(LPCTSTR Value)
{
	
	SetDayIntervalFormat(Value);
}

BSTR odl_GetDayOfWeekIntervalFormat(void)
{
	
	return GetDayOfWeekIntervalFormat().AllocSysString();
}

void odl_SetDayOfWeekIntervalFormat(LPCTSTR Value)
{
	
	SetDayOfWeekIntervalFormat(Value);
}

BSTR odl_GetDayOfYearIntervalFormat(void)
{
	
	return GetDayOfYearIntervalFormat().AllocSysString();
}

void odl_SetDayOfYearIntervalFormat(LPCTSTR Value)
{
	
	SetDayOfYearIntervalFormat(Value);
}

BSTR odl_GetWeekIntervalFormat(void)
{
	
	return GetWeekIntervalFormat().AllocSysString();
}

void odl_SetWeekIntervalFormat(LPCTSTR Value)
{
	
	SetWeekIntervalFormat(Value);
}

BSTR odl_GetMonthIntervalFormat(void)
{
	
	return GetMonthIntervalFormat().AllocSysString();
}

void odl_SetMonthIntervalFormat(LPCTSTR Value)
{
	
	SetMonthIntervalFormat(Value);
}

BSTR odl_GetQuarterIntervalFormat(void)
{
	
	return GetQuarterIntervalFormat().AllocSysString();
}

void odl_SetQuarterIntervalFormat(LPCTSTR Value)
{
	
	SetQuarterIntervalFormat(Value);
}

BSTR odl_GetYearIntervalFormat(void)
{
	
	return GetYearIntervalFormat().AllocSysString();
}

void odl_SetYearIntervalFormat(LPCTSTR Value)
{
	
	SetYearIntervalFormat(Value);
}

BSTR odl_GetXML(void)
{
	
	return GetXML().AllocSysString();
}

void odl_SetXML(LPCTSTR sXML)
{
	
	SetXML(sXML);
}

};