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
#include "stdafx.h"
#include "ActiveGanttVC.h"
#include "ActiveGanttVCCtl.h"
#include "clsXML.h"
#include "clsTierFormat.h"

IMPLEMENT_DYNCREATE(clsTierFormat, CCmdTarget)

//{F0E83CB9-4115-476B-B026-F88213427B23}
static const IID IID_IclsTierFormat = { 0xF0E83CB9, 0x4115, 0x476B, { 0xB0, 0x26, 0xF8, 0x82, 0x13, 0x42, 0x7B, 0x23} };

//{28B7B4F5-0ED9-4070-BBDF-56E7603AB391}
IMPLEMENT_OLECREATE_FLAGS(clsTierFormat, "AGVC.clsTierFormat", afxRegApartmentThreading, 0x28B7B4F5, 0x0ED9, 0x4070, 0xBB, 0xDF, 0x56, 0xE7, 0x60, 0x3A, 0xB3, 0x91)

BEGIN_DISPATCH_MAP(clsTierFormat, CCmdTarget)
	DISP_PROPERTY_EX_ID(clsTierFormat, "MinuteIntervalFormat", 1, odl_GetMinuteIntervalFormat, odl_SetMinuteIntervalFormat, VT_BSTR) 
	DISP_PROPERTY_EX_ID(clsTierFormat, "HourIntervalFormat", 2, odl_GetHourIntervalFormat, odl_SetHourIntervalFormat, VT_BSTR) 
	DISP_PROPERTY_EX_ID(clsTierFormat, "DayIntervalFormat", 3, odl_GetDayIntervalFormat, odl_SetDayIntervalFormat, VT_BSTR) 
	DISP_PROPERTY_EX_ID(clsTierFormat, "DayOfWeekIntervalFormat", 4, odl_GetDayOfWeekIntervalFormat, odl_SetDayOfWeekIntervalFormat, VT_BSTR) 
	DISP_PROPERTY_EX_ID(clsTierFormat, "DayOfYearIntervalFormat", 5, odl_GetDayOfYearIntervalFormat, odl_SetDayOfYearIntervalFormat, VT_BSTR) 
	DISP_PROPERTY_EX_ID(clsTierFormat, "WeekIntervalFormat", 6, odl_GetWeekIntervalFormat, odl_SetWeekIntervalFormat, VT_BSTR) 
	DISP_PROPERTY_EX_ID(clsTierFormat, "MonthIntervalFormat", 7, odl_GetMonthIntervalFormat, odl_SetMonthIntervalFormat, VT_BSTR) 
	DISP_PROPERTY_EX_ID(clsTierFormat, "QuarterIntervalFormat", 8, odl_GetQuarterIntervalFormat, odl_SetQuarterIntervalFormat, VT_BSTR) 
	DISP_PROPERTY_EX_ID(clsTierFormat, "YearIntervalFormat", 9, odl_GetYearIntervalFormat, odl_SetYearIntervalFormat, VT_BSTR) 
	DISP_FUNCTION_ID(clsTierFormat, "GetXML", 10, odl_GetXML, VT_BSTR, VTS_NONE) 
	DISP_FUNCTION_ID(clsTierFormat, "SetXML", 11, odl_SetXML, VT_EMPTY, VTS_BSTR)
	DISP_PROPERTY_EX_ID(clsTierFormat, "SecondIntervalFormat", 12, odl_GetSecondIntervalFormat, odl_SetSecondIntervalFormat, VT_BSTR)
	DISP_PROPERTY_EX_ID(clsTierFormat, "MillisecondIntervalFormat", 13, odl_GetMillisecondIntervalFormat, odl_SetSecondIntervalFormat, VT_BSTR)
	DISP_PROPERTY_EX_ID(clsTierFormat, "MicrosecondIntervalFormat", 14, odl_GetMicrosecondIntervalFormat, odl_SetSecondIntervalFormat, VT_BSTR) 
END_DISPATCH_MAP()

BEGIN_INTERFACE_MAP(clsTierFormat, CCmdTarget)
	INTERFACE_PART(clsTierFormat, IID_IclsTierFormat, Dispatch)
END_INTERFACE_MAP()

BEGIN_MESSAGE_MAP(clsTierFormat, CCmdTarget)
END_MESSAGE_MAP()

clsTierFormat::clsTierFormat()
{
	AfxMessageBox(_T("Invalid constructor has been invoked for clsTierFormat. Please inform support@sourcecodestore.com of this error"), MB_OK | MB_ICONSTOP | MB_APPLMODAL);
}

clsTierFormat::clsTierFormat(CActiveGanttVCCtl* oControl)
{
	EnableAutomation();
	AfxOleLockApp();
	mp_oControl = oControl;
	Clear();
}

clsTierFormat::~clsTierFormat()
{
	AfxOleUnlockApp();
}

void clsTierFormat::OnFinalRelease()
{
	CCmdTarget::OnFinalRelease();
}

CString clsTierFormat::GetMicrosecondIntervalFormat(void)
{
	return mp_sMicrosecondIntervalFormat;
}

void clsTierFormat::SetMicrosecondIntervalFormat(CString Value)
{
	mp_sMicrosecondIntervalFormat = Value;
}

CString clsTierFormat::GetMillisecondIntervalFormat(void)
{
	return mp_sMillisecondIntervalFormat;
}

void clsTierFormat::SetMillisecondIntervalFormat(CString Value)
{
	mp_sMillisecondIntervalFormat = Value;
}

CString clsTierFormat::GetSecondIntervalFormat(void)
{
	return mp_sSecondIntervalFormat;
}

void clsTierFormat::SetSecondIntervalFormat(CString Value)
{
	mp_sSecondIntervalFormat = Value;
}

CString clsTierFormat::GetMinuteIntervalFormat(void)
{
	return mp_sMinuteIntervalFormat;
}

void clsTierFormat::SetMinuteIntervalFormat(CString Value)
{
	mp_sMinuteIntervalFormat = Value;
}

CString clsTierFormat::GetHourIntervalFormat(void)
{
	return mp_sHourIntervalFormat;
}

void clsTierFormat::SetHourIntervalFormat(CString Value)
{
	mp_sHourIntervalFormat = Value;
}

CString clsTierFormat::GetDayIntervalFormat(void)
{
	return mp_sDayIntervalFormat;
}

void clsTierFormat::SetDayIntervalFormat(CString Value)
{
	mp_sDayIntervalFormat = Value;
}

CString clsTierFormat::GetDayOfWeekIntervalFormat(void)
{
	return mp_sDayOfWeekIntervalFormat;
}

void clsTierFormat::SetDayOfWeekIntervalFormat(CString Value)
{
	mp_sDayOfWeekIntervalFormat = Value;
}

CString clsTierFormat::GetDayOfYearIntervalFormat(void)
{
	return mp_sDayOfYearIntervalFormat;
}

void clsTierFormat::SetDayOfYearIntervalFormat(CString Value)
{
	mp_sDayOfYearIntervalFormat = Value;
}

CString clsTierFormat::GetWeekIntervalFormat(void)
{
	return mp_sWeekIntervalFormat;
}

void clsTierFormat::SetWeekIntervalFormat(CString Value)
{
	mp_sWeekIntervalFormat = Value;
}

CString clsTierFormat::GetMonthIntervalFormat(void)
{
	return mp_sMonthIntervalFormat;
}

void clsTierFormat::SetMonthIntervalFormat(CString Value)
{
	mp_sMonthIntervalFormat = Value;
}

CString clsTierFormat::GetQuarterIntervalFormat(void)
{
	return mp_sQuarterIntervalFormat;
}

void clsTierFormat::SetQuarterIntervalFormat(CString Value)
{
	mp_sQuarterIntervalFormat = Value;
}

CString clsTierFormat::GetYearIntervalFormat(void)
{
	return mp_sYearIntervalFormat;
}

void clsTierFormat::SetYearIntervalFormat(CString Value)
{
	mp_sYearIntervalFormat = Value;
}

void clsTierFormat::Finalize(void)
{
}

CString clsTierFormat::GetXML(void)
{
	clsXML oXML(mp_oControl, "TierFormat");
	oXML.InitializeWriter();
	oXML.WriteProperty("DayIntervalFormat", mp_sDayIntervalFormat);
	oXML.WriteProperty("DayOfWeekIntervalFormat", mp_sDayOfWeekIntervalFormat);
	oXML.WriteProperty("DayOfYearIntervalFormat", mp_sDayOfYearIntervalFormat);
	oXML.WriteProperty("HourIntervalFormat", mp_sHourIntervalFormat);
	oXML.WriteProperty("MinuteIntervalFormat", mp_sMinuteIntervalFormat);
	oXML.WriteProperty("SecondIntervalFormat", mp_sSecondIntervalFormat);
	oXML.WriteProperty("MillisecondIntervalFormat", mp_sMillisecondIntervalFormat);
	oXML.WriteProperty("MicrosecondIntervalFormat", mp_sMicrosecondIntervalFormat);
	oXML.WriteProperty("MonthIntervalFormat", mp_sMonthIntervalFormat);
	oXML.WriteProperty("QuarterIntervalFormat", mp_sQuarterIntervalFormat);
	oXML.WriteProperty("WeekIntervalFormat", mp_sWeekIntervalFormat);
	oXML.WriteProperty("YearIntervalFormat", mp_sYearIntervalFormat);
	return oXML.GetXML();
}

void clsTierFormat::SetXML(CString sXML)
{
	clsXML oXML(mp_oControl, "TierFormat");
	oXML.SetXML(sXML);
	oXML.InitializeReader();
    oXML.ReadProperty("DayIntervalFormat", mp_sDayIntervalFormat);
    oXML.ReadProperty("DayOfWeekIntervalFormat", mp_sDayOfWeekIntervalFormat);
    oXML.ReadProperty("DayOfYearIntervalFormat", mp_sDayOfYearIntervalFormat);
    oXML.ReadProperty("HourIntervalFormat", mp_sHourIntervalFormat);
    oXML.ReadProperty("MinuteIntervalFormat", mp_sMinuteIntervalFormat);
	oXML.ReadProperty("SecondIntervalFormat", mp_sSecondIntervalFormat);
	oXML.ReadProperty("MillisecondIntervalFormat", mp_sSecondIntervalFormat);
	oXML.ReadProperty("MicrosecondIntervalFormat", mp_sSecondIntervalFormat);
    oXML.ReadProperty("MonthIntervalFormat", mp_sMonthIntervalFormat);
    oXML.ReadProperty("QuarterIntervalFormat", mp_sQuarterIntervalFormat);
    oXML.ReadProperty("WeekIntervalFormat", mp_sWeekIntervalFormat);
    oXML.ReadProperty("YearIntervalFormat", mp_sYearIntervalFormat);
}

void clsTierFormat::Clear()
{
	mp_sMicrosecondIntervalFormat = "ffffff";
	mp_sMillisecondIntervalFormat = "fff";
	mp_sSecondIntervalFormat = "ss";
	mp_sMinuteIntervalFormat = "mm";
	mp_sHourIntervalFormat = "HH:mm";
	mp_sDayIntervalFormat = "d";
	mp_sDayOfWeekIntervalFormat = "dddd d";
	mp_sDayOfYearIntervalFormat = "y";
	mp_sWeekIntervalFormat = "ww";
	mp_sMonthIntervalFormat = "MMMM yyyy";
	mp_sQuarterIntervalFormat = "Q\\Q yyyy";
	mp_sYearIntervalFormat = "yyyy";
}

void clsTierFormat::Clone(clsTierFormat* oClone)
{
    oClone->SetMicrosecondIntervalFormat(mp_sMicrosecondIntervalFormat);
    oClone->SetMillisecondIntervalFormat(mp_sMillisecondIntervalFormat);
    oClone->SetSecondIntervalFormat(mp_sSecondIntervalFormat);
    oClone->SetMinuteIntervalFormat(mp_sMinuteIntervalFormat);
    oClone->SetHourIntervalFormat(mp_sHourIntervalFormat);
    oClone->SetDayIntervalFormat(mp_sDayIntervalFormat);
    oClone->SetDayOfWeekIntervalFormat(mp_sDayOfWeekIntervalFormat);
    oClone->SetDayOfYearIntervalFormat(mp_sDayOfYearIntervalFormat);
    oClone->SetWeekIntervalFormat(mp_sWeekIntervalFormat);
    oClone->SetMonthIntervalFormat(mp_sMonthIntervalFormat);
    oClone->SetQuarterIntervalFormat(mp_sQuarterIntervalFormat);
    oClone->SetYearIntervalFormat(mp_sYearIntervalFormat);
}