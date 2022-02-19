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
#include "clsTierAppearance.h"

IMPLEMENT_DYNCREATE(clsTierAppearance, CCmdTarget)

//{F3635208-6FBD-4998-A7F2-4C3BEB83D582}
static const IID IID_IclsTierAppearance = { 0xF3635208, 0x6FBD, 0x4998, { 0xA7, 0xF2, 0x4C, 0x3B, 0xEB, 0x83, 0xD5, 0x82} };

//{12A423DB-A1CD-45B5-9279-DE47CE00665F}
IMPLEMENT_OLECREATE_FLAGS(clsTierAppearance, "AGVC.clsTierAppearance", afxRegApartmentThreading, 0x12A423DB, 0xA1CD, 0x45B5, 0x92, 0x79, 0xDE, 0x47, 0xCE, 0x00, 0x66, 0x5F)

BEGIN_DISPATCH_MAP(clsTierAppearance, CCmdTarget)
	DISP_PROPERTY_EX_ID(clsTierAppearance, "MinuteColors", 1, odl_GetMinuteColors, SetNotSupported, VT_DISPATCH) 
	DISP_PROPERTY_EX_ID(clsTierAppearance, "HourColors", 2, odl_GetHourColors, SetNotSupported, VT_DISPATCH) 
	DISP_PROPERTY_EX_ID(clsTierAppearance, "DayColors", 3, odl_GetDayColors, SetNotSupported, VT_DISPATCH) 
	DISP_PROPERTY_EX_ID(clsTierAppearance, "DayOfWeekColors", 4, odl_GetDayOfWeekColors, SetNotSupported, VT_DISPATCH) 
	DISP_PROPERTY_EX_ID(clsTierAppearance, "DayOfYearColors", 5, odl_GetDayOfYearColors, SetNotSupported, VT_DISPATCH) 
	DISP_PROPERTY_EX_ID(clsTierAppearance, "WeekColors", 6, odl_GetWeekColors, SetNotSupported, VT_DISPATCH) 
	DISP_PROPERTY_EX_ID(clsTierAppearance, "MonthColors", 7, odl_GetMonthColors, SetNotSupported, VT_DISPATCH) 
	DISP_PROPERTY_EX_ID(clsTierAppearance, "QuarterColors", 8, odl_GetQuarterColors, SetNotSupported, VT_DISPATCH) 
	DISP_PROPERTY_EX_ID(clsTierAppearance, "YearColors", 9, odl_GetYearColors, SetNotSupported, VT_DISPATCH)  
	DISP_FUNCTION_ID(clsTierAppearance, "GetXML", 10, odl_GetXML, VT_BSTR, VTS_NONE) 
	DISP_FUNCTION_ID(clsTierAppearance, "SetXML", 11, odl_SetXML, VT_EMPTY, VTS_BSTR)
	DISP_PROPERTY_EX_ID(clsTierAppearance, "SecondColors", 12, odl_GetSecondColors, SetNotSupported, VT_DISPATCH)
	DISP_PROPERTY_EX_ID(clsTierAppearance, "MillisecondColors", 13, odl_GetMillisecondColors, SetNotSupported, VT_DISPATCH)
	DISP_PROPERTY_EX_ID(clsTierAppearance, "MicrosecondColors", 14, odl_GetMicrosecondColors, SetNotSupported, VT_DISPATCH)
END_DISPATCH_MAP()

BEGIN_INTERFACE_MAP(clsTierAppearance, CCmdTarget)
	INTERFACE_PART(clsTierAppearance, IID_IclsTierAppearance, Dispatch)
END_INTERFACE_MAP()

BEGIN_MESSAGE_MAP(clsTierAppearance, CCmdTarget)
END_MESSAGE_MAP()

clsTierAppearance::clsTierAppearance()
{
	AfxMessageBox(_T("Invalid constructor has been invoked for clsTierAppearance. Please inform support@sourcecodestore.com of this error"), MB_OK | MB_ICONSTOP | MB_APPLMODAL);
}

clsTierAppearance::clsTierAppearance(CActiveGanttVCCtl* oControl)
{
	EnableAutomation();
	AfxOleLockApp();
	mp_oControl = oControl;
    mp_oMicrosecondColors = new clsTierColors(mp_oControl, ST_MICROSECOND);
    mp_oMillisecondColors = new clsTierColors(mp_oControl, ST_MILLISECOND);
    mp_oSecondColors = new clsTierColors(mp_oControl, ST_SECOND);
    mp_oMinuteColors = new clsTierColors(mp_oControl, ST_MINUTE);
    mp_oHourColors = new clsTierColors(mp_oControl, ST_HOUR);
    mp_oDayColors = new clsTierColors(mp_oControl, ST_DAY);
    mp_oDayOfWeekColors = new clsTierColors(mp_oControl, ST_DAYOFWEEK);
    mp_oDayOfYearColors = new clsTierColors(mp_oControl, ST_DAYOFYEAR);
    mp_oWeekColors = new clsTierColors(mp_oControl, ST_WEEK);
    mp_oMonthColors = new clsTierColors(mp_oControl, ST_MONTH);
    mp_oQuarterColors = new clsTierColors(mp_oControl, ST_QUARTER);
    mp_oYearColors = new clsTierColors(mp_oControl, ST_YEAR);
	Clear();
}

clsTierAppearance::~clsTierAppearance()
{
	delete mp_oMicrosecondColors;
	mp_oMicrosecondColors = NULL;
	delete mp_oMillisecondColors;
	mp_oMillisecondColors = NULL;
	delete mp_oSecondColors;
	mp_oSecondColors = NULL;
	delete mp_oMinuteColors;
	mp_oMinuteColors = NULL;
	delete mp_oHourColors;
	mp_oHourColors = NULL;
	delete mp_oDayColors;
	mp_oDayColors = NULL;
	delete mp_oDayOfWeekColors;
	mp_oDayOfWeekColors = NULL;
	delete mp_oDayOfYearColors;
	mp_oDayOfYearColors = NULL;
	delete mp_oWeekColors;
	mp_oWeekColors = NULL;
	delete mp_oMonthColors;
	mp_oMonthColors = NULL;
	delete mp_oQuarterColors;
	mp_oQuarterColors = NULL;
	delete mp_oYearColors;
	mp_oYearColors = NULL;
	AfxOleUnlockApp();
}

void clsTierAppearance::OnFinalRelease()
{
	CCmdTarget::OnFinalRelease();
}



clsTierColors* clsTierAppearance::GetSecondColors(void)
{
	return mp_oSecondColors;
}

clsTierColors* clsTierAppearance::GetMicrosecondColors(void)
{
	return mp_oMicrosecondColors;
}

clsTierColors* clsTierAppearance::GetMillisecondColors(void)
{
	return mp_oMillisecondColors;
}

clsTierColors* clsTierAppearance::GetMinuteColors(void)
{
	return mp_oMinuteColors;
}

clsTierColors* clsTierAppearance::GetHourColors(void)
{
	return mp_oHourColors;
}

clsTierColors* clsTierAppearance::GetDayColors(void)
{
	return mp_oDayColors;
}

clsTierColors* clsTierAppearance::GetDayOfWeekColors(void)
{
	return mp_oDayOfWeekColors;
}

clsTierColors* clsTierAppearance::GetDayOfYearColors(void)
{
	return mp_oDayOfYearColors;
}

clsTierColors* clsTierAppearance::GetWeekColors(void)
{
	return mp_oWeekColors;
}

clsTierColors* clsTierAppearance::GetMonthColors(void)
{
	return mp_oMonthColors;
}

clsTierColors* clsTierAppearance::GetQuarterColors(void)
{
	return mp_oQuarterColors;
}

clsTierColors* clsTierAppearance::GetYearColors(void)
{
	return mp_oYearColors;
}

CString clsTierAppearance::GetXML(void)
{
	clsXML oXML(mp_oControl, "TierAppearance");
	oXML.InitializeWriter();
	oXML.WriteObject(mp_oDayColors->GetXML());
	oXML.WriteObject(mp_oDayOfWeekColors->GetXML());
	oXML.WriteObject(mp_oDayOfYearColors->GetXML());
	oXML.WriteObject(mp_oHourColors->GetXML());
	oXML.WriteObject(mp_oMinuteColors->GetXML());
	oXML.WriteObject(mp_oSecondColors->GetXML());
	oXML.WriteObject(mp_oMillisecondColors->GetXML());
	oXML.WriteObject(mp_oMicrosecondColors->GetXML());
	oXML.WriteObject(mp_oMonthColors->GetXML());
	oXML.WriteObject(mp_oQuarterColors->GetXML());
	oXML.WriteObject(mp_oWeekColors->GetXML());
	oXML.WriteObject(mp_oYearColors->GetXML());
	return oXML.GetXML();
}

void clsTierAppearance::SetXML(CString sXML)
{
	clsXML oXML(mp_oControl, "TierAppearance");
	oXML.SetXML(sXML);
	oXML.InitializeReader();
	mp_oDayColors->SetXML(oXML.ReadObject("DayColors"));
	mp_oDayOfWeekColors->SetXML(oXML.ReadObject("DayOfWeekColors"));
	mp_oDayOfYearColors->SetXML(oXML.ReadObject("DayOfYearColors"));
	mp_oHourColors->SetXML(oXML.ReadObject("HourColors"));
	mp_oMinuteColors->SetXML(oXML.ReadObject("MinuteColors"));
	mp_oSecondColors->SetXML(oXML.ReadObject("SecondColors"));
	mp_oMillisecondColors->SetXML(oXML.ReadObject("MillisecondColors"));
	mp_oMicrosecondColors->SetXML(oXML.ReadObject("MicrosecondColors"));
	mp_oMonthColors->SetXML(oXML.ReadObject("MonthColors"));
	mp_oQuarterColors->SetXML(oXML.ReadObject("QuarterColors"));
	mp_oWeekColors->SetXML(oXML.ReadObject("WeekColors"));
	mp_oYearColors->SetXML(oXML.ReadObject("YearColors"));
}

void clsTierAppearance::Clear()
{
        mp_oMicrosecondColors->Clear();
        mp_oMicrosecondColors->Add(Color::DimGray, Color::Black, Color::DimGray, Color::Black, Color::DimGray, Color::Black);
        mp_oMicrosecondColors->Add(Color::Gray, Color::Black, Color::Gray, Color::Black, Color::Gray, Color::Black);
        mp_oMicrosecondColors->Add(Color::DarkGray, Color::Black, Color::DarkGray, Color::Black, Color::DarkGray, Color::Black);
        mp_oMicrosecondColors->Add(Color::Silver, Color::Black, Color::Silver, Color::Black, Color::Silver, Color::Black);
        mp_oMicrosecondColors->Add(Color::DimGray, Color::Black, Color::DimGray, Color::Black, Color::DimGray, Color::Black);
        mp_oMicrosecondColors->Add(Color::Gray, Color::Black, Color::Gray, Color::Black, Color::Gray, Color::Black);
        mp_oMicrosecondColors->Add(Color::DarkGray, Color::Black, Color::DarkGray, Color::Black, Color::DarkGray, Color::Black);
        mp_oMicrosecondColors->Add(Color::Silver, Color::Black, Color::Silver, Color::Black, Color::Silver, Color::Black);
        mp_oMicrosecondColors->Add(Color::DimGray, Color::Black, Color::DimGray, Color::Black, Color::DimGray, Color::Black);
        mp_oMicrosecondColors->Add(Color::Gray, Color::Black, Color::Gray, Color::Black, Color::Gray, Color::Black);
        mp_oMillisecondColors->Clear();
        mp_oMillisecondColors->Add(Color::DimGray, Color::Black, Color::DimGray, Color::Black, Color::DimGray, Color::Black);
        mp_oMillisecondColors->Add(Color::Gray, Color::Black, Color::Gray, Color::Black, Color::Gray, Color::Black);
        mp_oMillisecondColors->Add(Color::DarkGray, Color::Black, Color::DarkGray, Color::Black, Color::DarkGray, Color::Black);
        mp_oMillisecondColors->Add(Color::Silver, Color::Black, Color::Silver, Color::Black, Color::Silver, Color::Black);
        mp_oMillisecondColors->Add(Color::DimGray, Color::Black, Color::DimGray, Color::Black, Color::DimGray, Color::Black);
        mp_oMillisecondColors->Add(Color::Gray, Color::Black, Color::Gray, Color::Black, Color::Gray, Color::Black);
        mp_oMillisecondColors->Add(Color::DarkGray, Color::Black, Color::DarkGray, Color::Black, Color::DarkGray, Color::Black);
        mp_oMillisecondColors->Add(Color::Silver, Color::Black, Color::Silver, Color::Black, Color::Silver, Color::Black);
        mp_oMillisecondColors->Add(Color::DimGray, Color::Black, Color::DimGray, Color::Black, Color::DimGray, Color::Black);
        mp_oMillisecondColors->Add(Color::Gray, Color::Black, Color::Gray, Color::Black, Color::Gray, Color::Black);
        mp_oSecondColors->Clear();
        mp_oSecondColors->Add(Color::DimGray, Color::Black, Color::DimGray, Color::Black, Color::DimGray, Color::Black);
        mp_oSecondColors->Add(Color::Gray, Color::Black, Color::Gray, Color::Black, Color::Gray, Color::Black);
        mp_oSecondColors->Add(Color::DarkGray, Color::Black, Color::DarkGray, Color::Black, Color::DarkGray, Color::Black);
        mp_oSecondColors->Add(Color::Silver, Color::Black, Color::Silver, Color::Black, Color::Silver, Color::Black);
        mp_oSecondColors->Add(Color::DimGray, Color::Black, Color::DimGray, Color::Black, Color::DimGray, Color::Black);
        mp_oSecondColors->Add(Color::Gray, Color::Black, Color::Gray, Color::Black, Color::Gray, Color::Black);
        mp_oSecondColors->Add(Color::DarkGray, Color::Black, Color::DarkGray, Color::Black, Color::DarkGray, Color::Black);
        mp_oSecondColors->Add(Color::Silver, Color::Black, Color::Silver, Color::Black, Color::Silver, Color::Black);
        mp_oSecondColors->Add(Color::DimGray, Color::Black, Color::DimGray, Color::Black, Color::DimGray, Color::Black);
        mp_oSecondColors->Add(Color::Gray, Color::Black, Color::Gray, Color::Black, Color::Gray, Color::Black);
        mp_oMinuteColors->Clear();
        mp_oMinuteColors->Add(Color::DimGray, Color::Black, Color::DimGray, Color::Black, Color::DimGray, Color::Black);
        mp_oMinuteColors->Add(Color::Gray, Color::Black, Color::Gray, Color::Black, Color::Gray, Color::Black);
        mp_oMinuteColors->Add(Color::DarkGray, Color::Black, Color::DarkGray, Color::Black, Color::DarkGray, Color::Black);
        mp_oMinuteColors->Add(Color::Silver, Color::Black, Color::Silver, Color::Black, Color::Silver, Color::Black);
        mp_oMinuteColors->Add(Color::DimGray, Color::Black, Color::DimGray, Color::Black, Color::DimGray, Color::Black);
        mp_oMinuteColors->Add(Color::Gray, Color::Black, Color::Gray, Color::Black, Color::Gray, Color::Black);
        mp_oMinuteColors->Add(Color::DarkGray, Color::Black, Color::DarkGray, Color::Black, Color::DarkGray, Color::Black);
        mp_oMinuteColors->Add(Color::Silver, Color::Black, Color::Silver, Color::Black, Color::Silver, Color::Black);
        mp_oMinuteColors->Add(Color::DimGray, Color::Black, Color::DimGray, Color::Black, Color::DimGray, Color::Black);
        mp_oMinuteColors->Add(Color::Gray, Color::Black, Color::Gray, Color::Black, Color::Gray, Color::Black);
        mp_oHourColors->Clear();
        mp_oHourColors->Add(Color::DimGray, Color::Black, Color::DimGray, Color::Black, Color::DimGray, Color::Black);
        mp_oHourColors->Add(Color::Gray, Color::Black, Color::Gray, Color::Black, Color::Gray, Color::Black);
        mp_oHourColors->Add(Color::DarkGray, Color::Black, Color::DarkGray, Color::Black, Color::DarkGray, Color::Black);
        mp_oHourColors->Add(Color::Silver, Color::Black, Color::Silver, Color::Black, Color::Silver, Color::Black);
        mp_oHourColors->Add(Color::DimGray, Color::Black, Color::DimGray, Color::Black, Color::DimGray, Color::Black);
        mp_oHourColors->Add(Color::Gray, Color::Black, Color::Gray, Color::Black, Color::Gray, Color::Black);
        mp_oHourColors->Add(Color::DarkGray, Color::Black, Color::DarkGray, Color::Black, Color::DarkGray, Color::Black);
        mp_oHourColors->Add(Color::Silver, Color::Black, Color::Silver, Color::Black, Color::Silver, Color::Black);
        mp_oHourColors->Add(Color::DimGray, Color::Black, Color::DimGray, Color::Black, Color::DimGray, Color::Black);
        mp_oHourColors->Add(Color::Gray, Color::Black, Color::Gray, Color::Black, Color::Gray, Color::Black);
        mp_oDayColors->Clear();
        mp_oDayColors->Add(Color::DimGray, Color::Black, Color::DimGray, Color::Black, Color::DimGray, Color::Black);
        mp_oDayColors->Add(Color::Gray, Color::Black, Color::Gray, Color::Black, Color::Gray, Color::Black);
        mp_oDayColors->Add(Color::DarkGray, Color::Black, Color::DarkGray, Color::Black, Color::DarkGray, Color::Black);
        mp_oDayColors->Add(Color::Silver, Color::Black, Color::Silver, Color::Black, Color::Silver, Color::Black);
        mp_oDayColors->Add(Color::DimGray, Color::Black, Color::DimGray, Color::Black, Color::DimGray, Color::Black);
        mp_oDayColors->Add(Color::Gray, Color::Black, Color::Gray, Color::Black, Color::Gray, Color::Black);
        mp_oDayColors->Add(Color::DarkGray, Color::Black, Color::DarkGray, Color::Black, Color::DarkGray, Color::Black);
        mp_oDayColors->Add(Color::Silver, Color::Black, Color::Silver, Color::Black, Color::Silver, Color::Black);
        mp_oDayColors->Add(Color::DimGray, Color::Black, Color::DimGray, Color::Black, Color::DimGray, Color::Black);
        mp_oDayColors->Add(Color::Gray, Color::Black, Color::Gray, Color::Black, Color::Gray, Color::Black);
        mp_oDayOfWeekColors->Clear();
        mp_oDayOfWeekColors->Add(Color::CornflowerBlue, Color::Black, Color::CornflowerBlue, Color::Black, Color::CornflowerBlue, Color::Black, "Sunday");
        mp_oDayOfWeekColors->Add(Color::MediumSlateBlue, Color::Black, Color::MediumSlateBlue, Color::Black, Color::MediumSlateBlue, Color::Black, "Monday");
        mp_oDayOfWeekColors->Add(Color::SlateBlue, Color::Black, Color::SlateBlue, Color::Black, Color::SlateBlue, Color::Black, "Tuesday");
        mp_oDayOfWeekColors->Add(Color::RoyalBlue, Color::Black, Color::RoyalBlue, Color::Black, Color::RoyalBlue, Color::Black, "Wednesday");
        mp_oDayOfWeekColors->Add(Color::SteelBlue, Color::Black, Color::SteelBlue, Color::Black, Color::SteelBlue, Color::Black, "Thursday");
        mp_oDayOfWeekColors->Add(Color::CadetBlue, Color::Black, Color::CadetBlue, Color::Black, Color::CadetBlue, Color::Black, "Friday");
        mp_oDayOfWeekColors->Add(Color::DodgerBlue, Color::Black, Color::DodgerBlue, Color::Black, Color::DodgerBlue, Color::Black, "Saturday");
        mp_oDayOfYearColors->Clear();
        mp_oDayOfYearColors->Add(Color::DimGray, Color::Black, Color::DimGray, Color::Black, Color::DimGray, Color::Black);
        mp_oDayOfYearColors->Add(Color::Gray, Color::Black, Color::Gray, Color::Black, Color::Gray, Color::Black);
        mp_oDayOfYearColors->Add(Color::DarkGray, Color::Black, Color::DarkGray, Color::Black, Color::DarkGray, Color::Black);
        mp_oDayOfYearColors->Add(Color::Silver, Color::Black, Color::Silver, Color::Black, Color::Silver, Color::Black);
        mp_oDayOfYearColors->Add(Color::DimGray, Color::Black, Color::DimGray, Color::Black, Color::DimGray, Color::Black);
        mp_oDayOfYearColors->Add(Color::Gray, Color::Black, Color::Gray, Color::Black, Color::Gray, Color::Black);
        mp_oDayOfYearColors->Add(Color::DarkGray, Color::Black, Color::DarkGray, Color::Black, Color::DarkGray, Color::Black);
        mp_oDayOfYearColors->Add(Color::Silver, Color::Black, Color::Silver, Color::Black, Color::Silver, Color::Black);
        mp_oDayOfYearColors->Add(Color::DimGray, Color::Black, Color::DimGray, Color::Black, Color::DimGray, Color::Black);
        mp_oDayOfYearColors->Add(Color::Gray, Color::Black, Color::Gray, Color::Black, Color::Gray, Color::Black);
        mp_oWeekColors->Clear();
        mp_oWeekColors->Add(Color::DimGray, Color::Black, Color::DimGray, Color::Black, Color::DimGray, Color::Black);
        mp_oWeekColors->Add(Color::Gray, Color::Black, Color::Gray, Color::Black, Color::Gray, Color::Black);
        mp_oWeekColors->Add(Color::DarkGray, Color::Black, Color::DarkGray, Color::Black, Color::DarkGray, Color::Black);
        mp_oWeekColors->Add(Color::Silver, Color::Black, Color::Silver, Color::Black, Color::Silver, Color::Black);
        mp_oWeekColors->Add(Color::DimGray, Color::Black, Color::DimGray, Color::Black, Color::DimGray, Color::Black);
        mp_oWeekColors->Add(Color::Gray, Color::Black, Color::Gray, Color::Black, Color::Gray, Color::Black);
        mp_oWeekColors->Add(Color::DarkGray, Color::Black, Color::DarkGray, Color::Black, Color::DarkGray, Color::Black);
        mp_oWeekColors->Add(Color::Silver, Color::Black, Color::Silver, Color::Black, Color::Silver, Color::Black);
        mp_oWeekColors->Add(Color::DimGray, Color::Black, Color::DimGray, Color::Black, Color::DimGray, Color::Black);
        mp_oWeekColors->Add(Color::Gray, Color::Black, Color::Gray, Color::Black, Color::Gray, Color::Black);
        mp_oMonthColors->Clear();
        mp_oMonthColors->Add(Color::DimGray, Color::Black, Color::DimGray, Color::Black, Color::DimGray, Color::Black, "January");
        mp_oMonthColors->Add(Color::Gray, Color::Black, Color::Gray, Color::Black, Color::Gray, Color::Black, "February");
        mp_oMonthColors->Add(Color::DarkGray, Color::Black, Color::DarkGray, Color::Black, Color::DarkGray, Color::Black, "March");
        mp_oMonthColors->Add(Color::Silver, Color::Black, Color::Silver, Color::Black, Color::Silver, Color::Black, "April");
        mp_oMonthColors->Add(Color::DimGray, Color::Black, Color::DimGray, Color::Black, Color::DimGray, Color::Black, "May");
        mp_oMonthColors->Add(Color::Gray, Color::Black, Color::Gray, Color::Black, Color::Gray, Color::Black, "June");
        mp_oMonthColors->Add(Color::DarkGray, Color::Black, Color::DarkGray, Color::Black, Color::DarkGray, Color::Black, "July");
        mp_oMonthColors->Add(Color::Silver, Color::Black, Color::Silver, Color::Black, Color::Silver, Color::Black, "August");
        mp_oMonthColors->Add(Color::DimGray, Color::Black, Color::DimGray, Color::Black, Color::DimGray, Color::Black, "September");
        mp_oMonthColors->Add(Color::Gray, Color::Black, Color::Gray, Color::Black, Color::Gray, Color::Black, "October");
        mp_oMonthColors->Add(Color::DarkGray, Color::Black, Color::DarkGray, Color::Black, Color::DarkGray, Color::Black, "November");
        mp_oMonthColors->Add(Color::Silver, Color::Black, Color::Silver, Color::Black, Color::Silver, Color::Black, "December");
        mp_oQuarterColors->Clear();
        mp_oQuarterColors->Add(Color::DimGray, Color::Black, Color::DimGray, Color::Black, Color::DimGray, Color::Black);
        mp_oQuarterColors->Add(Color::Gray, Color::Black, Color::Gray, Color::Black, Color::Gray, Color::Black);
        mp_oQuarterColors->Add(Color::DarkGray, Color::Black, Color::DarkGray, Color::Black, Color::DarkGray, Color::Black);
        mp_oQuarterColors->Add(Color::Silver, Color::Black, Color::Silver, Color::Black, Color::Silver, Color::Black);
        mp_oYearColors->Clear();
        mp_oYearColors->Add(Color::DimGray, Color::Black, Color::DimGray, Color::Black, Color::DimGray, Color::Black);
        mp_oYearColors->Add(Color::Gray, Color::Black, Color::Gray, Color::Black, Color::Gray, Color::Black);
        mp_oYearColors->Add(Color::DarkGray, Color::Black, Color::DarkGray, Color::Black, Color::DarkGray, Color::Black);
        mp_oYearColors->Add(Color::Silver, Color::Black, Color::Silver, Color::Black, Color::Silver, Color::Black);
        mp_oYearColors->Add(Color::DimGray, Color::Black, Color::DimGray, Color::Black, Color::DimGray, Color::Black);
        mp_oYearColors->Add(Color::Gray, Color::Black, Color::Gray, Color::Black, Color::Gray, Color::Black);
        mp_oYearColors->Add(Color::DarkGray, Color::Black, Color::DarkGray, Color::Black, Color::DarkGray, Color::Black);
        mp_oYearColors->Add(Color::Silver, Color::Black, Color::Silver, Color::Black, Color::Silver, Color::Black);
        mp_oYearColors->Add(Color::DimGray, Color::Black, Color::DimGray, Color::Black, Color::DimGray, Color::Black);
        mp_oYearColors->Add(Color::Gray, Color::Black, Color::Gray, Color::Black, Color::Gray, Color::Black);
}

void clsTierAppearance::Clone(clsTierAppearance* oClone)
{
    GetMicrosecondColors()->Clone(oClone->GetMicrosecondColors());
    GetMillisecondColors()->Clone(oClone->GetMillisecondColors());
    GetSecondColors()->Clone(oClone->GetSecondColors());
    GetMinuteColors()->Clone(oClone->GetMinuteColors());
    GetHourColors()->Clone(oClone->GetHourColors());
    GetDayColors()->Clone(oClone->GetDayColors());
    GetDayOfWeekColors()->Clone(oClone->GetDayOfWeekColors());
    GetDayOfYearColors()->Clone(oClone->GetDayOfYearColors());
    GetWeekColors()->Clone(oClone->GetWeekColors());
    GetMonthColors()->Clone(oClone->GetMonthColors());
    GetQuarterColors()->Clone(oClone->GetQuarterColors());
    GetYearColors()->Clone(oClone->GetYearColors());
}