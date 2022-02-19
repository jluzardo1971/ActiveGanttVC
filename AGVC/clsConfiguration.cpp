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
#include "clsConfiguration.h"
#include "ActiveGanttVCCtl.h"
#include "clsXML.h"




IMPLEMENT_DYNCREATE(clsConfiguration, CCmdTarget)

//{7A3C6473-E7D2-4898-8951-3C87AA1F69A5}
static const IID IID_IclsConfiguration = { 0x7A3C6473, 0xE7D2, 0x4898, { 0x89, 0x51, 0x3C, 0x87, 0xAA, 0x1F, 0x69, 0xA5} };


//{C105F698-DED6-470A-8A59-CDC14E858C72}
IMPLEMENT_OLECREATE_FLAGS(clsConfiguration, "AGVC.clsConfiguration", afxRegApartmentThreading, 0xC105F698, 0xDED6, 0x470A, 0x8A, 0x59, 0xCD, 0xC1, 0x4E, 0x85, 0x8C, 0x72)

BEGIN_MESSAGE_MAP(clsConfiguration, CCmdTarget)
END_MESSAGE_MAP()

BEGIN_DISPATCH_MAP(clsConfiguration, CCmdTarget)
	DISP_PROPERTY_EX_ID(clsConfiguration, "FirstDayOfWeek", 1, odl_GetFirstDayOfWeek, odl_SetFirstDayOfWeek, VT_I4)
	DISP_PROPERTY_EX_ID(clsConfiguration, "CalendarWeekRule", 2, odl_GetCalendarWeekRule, odl_SetCalendarWeekRule, VT_I4) 
	DISP_PROPERTY_EX_ID(clsConfiguration, "ISO8601Weeks", 3, odl_GetISO8601Weeks, odl_SetISO8601Weeks, VT_BOOL)
	DISP_FUNCTION_ID(clsConfiguration, "GetXML", 4, odl_GetXML, VT_BSTR, VTS_NONE) 
	DISP_FUNCTION_ID(clsConfiguration, "SetXML", 5, odl_SetXML, VT_EMPTY, VTS_BSTR)
	DISP_PROPERTY_EX_ID(clsConfiguration, "RowHighlightedBackColor", 6, odl_GetRowHighlightedBackColor, odl_SetRowHighlightedBackColor, VT_COLOR)
	DISP_PROPERTY_EX_ID(clsConfiguration, "RowEvenBackColor", 7, odl_GetRowEvenBackColor, odl_SetRowEvenBackColor, VT_COLOR)
	DISP_PROPERTY_EX_ID(clsConfiguration, "RowOddBackColor", 8, odl_GetRowOddBackColor, odl_SetRowOddBackColor, VT_COLOR)
	DISP_PROPERTY_EX_ID(clsConfiguration, "AlternatingRowStyles", 9, odl_GetAlternatingRowStyles, odl_SetAlternatingRowStyles, VT_BOOL)
	DISP_PROPERTY_EX_ID(clsConfiguration, "DefaultFont", 10, odl_GetDefaultFont, SetNotSupported, VT_DISPATCH)
	DISP_PROPERTY_EX_ID(clsConfiguration, "DefaultControlStyle", 11, odl_GetDefaultControlStyle, SetNotSupported, VT_DISPATCH)
	DISP_PROPERTY_EX_ID(clsConfiguration, "DefaultTaskStyle", 12, odl_GetDefaultTaskStyle, SetNotSupported, VT_DISPATCH)
	DISP_PROPERTY_EX_ID(clsConfiguration, "DefaultRowStyle", 13, odl_GetDefaultRowStyle, SetNotSupported, VT_DISPATCH)
	DISP_PROPERTY_EX_ID(clsConfiguration, "DefaultClientAreaStyle", 14, odl_GetDefaultClientAreaStyle, SetNotSupported, VT_DISPATCH)
	DISP_PROPERTY_EX_ID(clsConfiguration, "DefaultCellStyle", 15, odl_GetDefaultCellStyle, SetNotSupported, VT_DISPATCH)
	DISP_PROPERTY_EX_ID(clsConfiguration, "DefaultColumnStyle", 16, odl_GetDefaultColumnStyle, SetNotSupported, VT_DISPATCH)
	DISP_PROPERTY_EX_ID(clsConfiguration, "DefaultPercentageStyle", 17, odl_GetDefaultPercentageStyle, SetNotSupported, VT_DISPATCH)
	DISP_PROPERTY_EX_ID(clsConfiguration, "DefaultPredecessorStyle", 18, odl_GetDefaultPredecessorStyle, SetNotSupported, VT_DISPATCH)
	DISP_PROPERTY_EX_ID(clsConfiguration, "DefaultTimeLineStyle", 19, odl_GetDefaultTimeLineStyle, SetNotSupported, VT_DISPATCH)
	DISP_PROPERTY_EX_ID(clsConfiguration, "DefaultTimeBlockStyle", 20, odl_GetDefaultTimeBlockStyle, SetNotSupported, VT_DISPATCH)
	DISP_PROPERTY_EX_ID(clsConfiguration, "DefaultTickMarkAreaStyle", 21, odl_GetDefaultTickMarkAreaStyle, SetNotSupported, VT_DISPATCH)
	DISP_PROPERTY_EX_ID(clsConfiguration, "DefaultSplitterStyle", 22, odl_GetDefaultSplitterStyle, SetNotSupported, VT_DISPATCH)
	DISP_PROPERTY_EX_ID(clsConfiguration, "DefaultProgressLineStyle", 23, odl_GetDefaultProgressLineStyle, SetNotSupported, VT_DISPATCH)
	DISP_PROPERTY_EX_ID(clsConfiguration, "DefaultNodeStyle", 24, odl_GetDefaultNodeStyle, SetNotSupported, VT_DISPATCH)
	DISP_PROPERTY_EX_ID(clsConfiguration, "DefaultTierStyle", 25, odl_GetDefaultTierStyle, SetNotSupported, VT_DISPATCH)
	DISP_PROPERTY_EX_ID(clsConfiguration, "DefaultScrollBarStyle", 26, odl_GetDefaultScrollBarStyle, SetNotSupported, VT_DISPATCH)
	DISP_PROPERTY_EX_ID(clsConfiguration, "DefaultSBSeparatorStyle", 27, odl_GetDefaultSBSeparatorStyle, SetNotSupported, VT_DISPATCH)
	DISP_PROPERTY_EX_ID(clsConfiguration, "DefaultSBNormalStyle", 28, odl_GetDefaultSBNormalStyle, SetNotSupported, VT_DISPATCH)
	DISP_PROPERTY_EX_ID(clsConfiguration, "DefaultSBPressedStyle", 29, odl_GetDefaultSBPressedStyle, SetNotSupported, VT_DISPATCH)
	DISP_PROPERTY_EX_ID(clsConfiguration, "DefaultSBDisabledStyle", 30, odl_GetDefaultSBDisabledStyle, SetNotSupported, VT_DISPATCH)
END_DISPATCH_MAP()



BEGIN_INTERFACE_MAP(clsConfiguration, CCmdTarget)
	INTERFACE_PART(clsConfiguration, IID_IclsConfiguration, Dispatch)
END_INTERFACE_MAP()

clsConfiguration::clsConfiguration()
{
	AfxMessageBox(_T("Invalid constructor has been invoked for clsConfiguration. Please inform support@sourcecodestore.com of this error"), MB_OK | MB_ICONSTOP | MB_APPLMODAL);
}

clsConfiguration::clsConfiguration(CActiveGanttVCCtl* oControl)
{
	EnableAutomation();	
	AfxOleLockApp();
	mp_oControl = oControl;
    mp_yFirstDayOfWeek = WD_SUNDAY;
    mp_yCalendarWeekRule = CWR_FIRSTDAY;
    mp_bISO8601Weeks = FALSE;
    mp_clrRowHighlightedBackColor = Color::Red;
    mp_clrRowEvenBackColor = Color::White;
    mp_clrRowOddBackColor = Color::MakeARGB(255, 192, 192, 192);
    mp_bAlternatingRowStyles = FALSE;
	mp_oDefaultFont.SetPtFont(_T("Tahoma"), 8, FS_FONTSTYLEREGULAR);
}

clsConfiguration::~clsConfiguration()
{
	AfxOleUnlockApp();
}


void clsConfiguration::OnFinalRelease()
{
	CCmdTarget::OnFinalRelease();
}

clsFont* clsConfiguration::GetDefaultFont(void)
{
	return &mp_oDefaultFont;
}

Gdiplus::Color clsConfiguration::GetRowHighlightedBackColor(void)
{
	return mp_clrRowHighlightedBackColor;
}

void clsConfiguration::SetRowHighlightedBackColor(Gdiplus::Color Value)
{
	mp_clrRowHighlightedBackColor = Value;
}

Gdiplus::Color clsConfiguration::GetRowEvenBackColor(void)
{
	return mp_clrRowEvenBackColor;
}

void clsConfiguration::SetRowEvenBackColor(Gdiplus::Color Value)
{
	mp_clrRowEvenBackColor = Value;
}

Gdiplus::Color clsConfiguration::GetRowOddBackColor(void)
{
	return mp_clrRowOddBackColor;
}

void clsConfiguration::SetRowOddBackColor(Gdiplus::Color Value)
{
	mp_clrRowOddBackColor = Value;
}

BOOL clsConfiguration::GetAlternatingRowStyles(void)
{
	return mp_bAlternatingRowStyles;
}

void clsConfiguration::SetAlternatingRowStyles(BOOL value)
{
	mp_bAlternatingRowStyles = value;
}

LONG clsConfiguration::GetFirstDayOfWeek(void)
{
	return mp_yFirstDayOfWeek;
}

void clsConfiguration::SetFirstDayOfWeek(LONG value)
{
	mp_yFirstDayOfWeek = value;
}

LONG clsConfiguration::GetCalendarWeekRule(void)
{
	return mp_yCalendarWeekRule;
}

void clsConfiguration::SetCalendarWeekRule(LONG value)
{
	mp_yCalendarWeekRule = value;
}

BOOL clsConfiguration::GetISO8601Weeks(void)
{
	return mp_bISO8601Weeks;
}

void clsConfiguration::SetISO8601Weeks(BOOL value)
{
	mp_bISO8601Weeks = value;
}

clsStyle*  clsConfiguration::GetDefaultControlStyle(void)
{
	return mp_oControl->GetStyles()->mp_oDefaultControlStyle;
}

clsStyle*  clsConfiguration::GetDefaultTaskStyle(void)
{
	return mp_oControl->GetStyles()->mp_oDefaultTaskStyle;
}

clsStyle*  clsConfiguration::GetDefaultRowStyle(void)
{
	return mp_oControl->GetStyles()->mp_oDefaultRowStyle;
}

clsStyle*  clsConfiguration::GetDefaultClientAreaStyle(void)
{
	return mp_oControl->GetStyles()->mp_oDefaultClientAreaStyle;
}

clsStyle*  clsConfiguration::GetDefaultCellStyle(void)
{
	return mp_oControl->GetStyles()->mp_oDefaultCellStyle;
}

clsStyle*  clsConfiguration::GetDefaultColumnStyle(void)
{
	return mp_oControl->GetStyles()->mp_oDefaultColumnStyle;
}

clsStyle*  clsConfiguration::GetDefaultPercentageStyle(void)
{
	return mp_oControl->GetStyles()->mp_oDefaultPercentageStyle;
}

clsStyle*  clsConfiguration::GetDefaultPredecessorStyle(void)
{
	return mp_oControl->GetStyles()->mp_oDefaultPredecessorStyle;
}

clsStyle*  clsConfiguration::GetDefaultTimeLineStyle(void)
{
	return mp_oControl->GetStyles()->mp_oDefaultTimeLineStyle;
}

clsStyle*  clsConfiguration::GetDefaultTimeBlockStyle(void)
{
	return mp_oControl->GetStyles()->mp_oDefaultTimeBlockStyle;
}

clsStyle*  clsConfiguration::GetDefaultTickMarkAreaStyle(void)
{
	return mp_oControl->GetStyles()->mp_oDefaultTickMarkAreaStyle;
}

clsStyle*  clsConfiguration::GetDefaultSplitterStyle(void)
{
	return mp_oControl->GetStyles()->mp_oDefaultSplitterStyle;
}

clsStyle*  clsConfiguration::GetDefaultProgressLineStyle(void)
{
	return mp_oControl->GetStyles()->mp_oDefaultProgressLineStyle;
}

clsStyle*  clsConfiguration::GetDefaultNodeStyle(void)
{
	return mp_oControl->GetStyles()->mp_oDefaultNodeStyle;
}

clsStyle*  clsConfiguration::GetDefaultTierStyle(void)
{
	return mp_oControl->GetStyles()->mp_oDefaultTierStyle;
}

clsStyle*  clsConfiguration::GetDefaultScrollBarStyle(void)
{
	return mp_oControl->GetStyles()->mp_oDefaultScrollBarStyle;
}

clsStyle*  clsConfiguration::GetDefaultSBSeparatorStyle(void)
{
	return mp_oControl->GetStyles()->mp_oDefaultSBSeparatorStyle;
}

clsStyle*  clsConfiguration::GetDefaultSBNormalStyle(void)
{
	return mp_oControl->GetStyles()->mp_oDefaultSBNormalStyle;
}

clsStyle*  clsConfiguration::GetDefaultSBPressedStyle(void)
{
	return mp_oControl->GetStyles()->mp_oDefaultSBPressedStyle;
}

clsStyle*  clsConfiguration::GetDefaultSBDisabledStyle(void)
{
	return mp_oControl->GetStyles()->mp_oDefaultSBDisabledStyle;
}

CString clsConfiguration::Format(COleDateTime dtExpression, CString sFormat)
{
	CString sBuff;
	int lBuff;

	//"d", "f", "F", "g", "h", "H", "K", "m", "M", "s", "t", "y", "z"
	if (sFormat == _T("%d")) { sFormat = _T("d"); }
	if (sFormat == _T("%f")) { sFormat = _T("f"); }
	if (sFormat == _T("%F")) { sFormat = _T("F"); }
	if (sFormat == _T("%g")) { sFormat = _T("g"); }
	if (sFormat == _T("%h")) { sFormat = _T("h"); }
	if (sFormat == _T("%H")) { sFormat = _T("H"); }
	if (sFormat == _T("%K")) { sFormat = _T("K"); }
	if (sFormat == _T("%m")) { sFormat = _T("m"); }
	if (sFormat == _T("%M")) { sFormat = _T("M"); }
	if (sFormat == _T("%s")) { sFormat = _T("s"); }
	if (sFormat == _T("%t")) { sFormat = _T("t"); }
	if (sFormat == _T("%y")) { sFormat = _T("y"); }
	if (sFormat == _T("%z")) { sFormat = _T("z"); }

	////Escape Character
	sFormat.Replace(_T("\\\\"), _T("{~°|1}"));

	////YEARS
	sFormat.Replace(_T("\\y"), _T("{~°|}"));
	if (sFormat.Find(_T("yyyyy"), 0) != -1)
	{
		//The year as a five-digit number.
		//   1/1/0001 12:00:00 AM -> 00001
		//   6/15/2009 1:45:30 PM -> 02009
		//The year as a four-digit number.
		//   1/1/0001 12:00:00 AM -> 0001
		sBuff = CStr(dtExpression.GetYear());
		if (sBuff.GetLength() == 4)
		{
			sBuff = _T("0") + sBuff;
		}
		sFormat.Replace(_T("yyyyy"), sBuff);
	}
	if (sFormat.Find(_T("yyyy"), 0) != -1)
	{
		//The year as a four-digit number.
		//   1/1/0001 12:00:00 AM -> 0001
		sBuff = CStr(dtExpression.GetYear());
		sFormat.Replace(_T("yyyy"), sBuff);
	}
	if (sFormat.Find(_T("yyy"), 0) != -1)
	{
		//The year, with a minimum of three digits.
		//   1/1/0001 12:00:00 AM -> 001
		sBuff = CStr(dtExpression.GetYear());
		sBuff = sBuff.Right(3);
		sFormat.Replace(_T("yyy"), sBuff);
	}
	if (sFormat.Find(_T("yy"), 0) != -1)
	{
		//The year, from 00 to 99.
		//   1/1/0001 12:00:00 AM -> 01
		sBuff = CStr(dtExpression.GetYear());
		sBuff = sBuff.Right(2);
		sFormat.Replace(_T("yy"), sBuff);
	}
	if (sFormat.Find(_T("y"), 0) != -1)
	{
		//The year, from 0 to 99.
		//   1/1/0001 12:00:00 AM -> 1
		sBuff = CStr(dtExpression.GetYear());
		sBuff = sBuff.Right(1);
		sFormat.Replace(_T("yy"), sBuff);
	}
	sFormat.Replace(_T("{~°|}"), _T("y"));

	////QUARTERS
	sFormat.Replace(_T("\\Q"), _T("{~°|}"));
	if (sFormat.Find(_T("QQ"), 0) != -1)
	{
		//The quarter of the year with a leading zero.
		//   6/15/2009 1:45:30 PM  -> 4
		sBuff = CStr(mp_oControl->GetMathLib()->Quarter(dtExpression));
		if (sBuff.GetLength() == 1)
		{
			sBuff = _T("0") + sBuff;
		}
		sFormat.Replace(_T("QQ"), sBuff);
	}
	if (sFormat.Find(_T("Q"), 0) != -1)
	{
		//The quarter of the year.
		//   6/15/2009 1:45:30 PM  -> 4
		sBuff = CStr(mp_oControl->GetMathLib()->Quarter(dtExpression));
		sFormat.Replace(_T("Q"), sBuff);

	}
	sFormat.Replace(_T("{~°|}"), _T("Q"));

	////MONTHS

	sFormat.Replace(_T("\\M"), _T("{~°|}"));
	if (sFormat.Find(_T("MMMM"), 0) != -1)
	{
		//The full name of the month.
		//   6/15/2009 1:45:30 PM -> June (en-US)
		switch (dtExpression.GetMonth())
		{
		case 1:
			sFormat.Replace(_T("MMMM"), mp_EscString(GetLocaleInfoEx(LOCALE_SMONTHNAME1)));
			break;
		case 2:
			sFormat.Replace(_T("MMMM"), mp_EscString(GetLocaleInfoEx(LOCALE_SMONTHNAME2)));
			break;
		case 3:
			sFormat.Replace(_T("MMMM"), mp_EscString(GetLocaleInfoEx(LOCALE_SMONTHNAME3)));
			break;
		case 4:
			sFormat.Replace(_T("MMMM"), mp_EscString(GetLocaleInfoEx(LOCALE_SMONTHNAME4)));
			break;
		case 5:
			sFormat.Replace(_T("MMMM"), mp_EscString(GetLocaleInfoEx(LOCALE_SMONTHNAME5)));
			break;
		case 6:
			sFormat.Replace(_T("MMMM"), mp_EscString(GetLocaleInfoEx(LOCALE_SMONTHNAME6)));
			break;
		case 7:
			sFormat.Replace(_T("MMMM"), mp_EscString(GetLocaleInfoEx(LOCALE_SMONTHNAME7)));
			break;
		case 8:
			sFormat.Replace(_T("MMMM"), mp_EscString(GetLocaleInfoEx(LOCALE_SMONTHNAME8)));
			break;
		case 9:
			sFormat.Replace(_T("MMMM"), mp_EscString(GetLocaleInfoEx(LOCALE_SMONTHNAME9)));
			break;
		case 10:
			sFormat.Replace(_T("MMMM"), mp_EscString(GetLocaleInfoEx(LOCALE_SMONTHNAME10)));
			break;
		case 11:
			sFormat.Replace(_T("MMMM"), mp_EscString(GetLocaleInfoEx(LOCALE_SMONTHNAME11)));
			break;
		case 12:
			sFormat.Replace(_T("MMMM"), mp_EscString(GetLocaleInfoEx(LOCALE_SMONTHNAME12)));
			break;
		}
		sFormat.Replace(_T("\\M"), _T("{~°|}"));
	}
	if (sFormat.Find(_T("MMM"), 0) != -1)
	{
		//The abbreviated name of the month.
		//   6/15/2009 1:45:30 PM -> Jun (en-US)
		switch (dtExpression.GetMonth())
		{
		case 1:
			sFormat.Replace(_T("MMM"), mp_EscString(GetLocaleInfoEx(LOCALE_SABBREVMONTHNAME1)));
			break;
		case 2:
			sFormat.Replace(_T("MMM"), mp_EscString(GetLocaleInfoEx(LOCALE_SABBREVMONTHNAME2)));
			break;
		case 3:
			sFormat.Replace(_T("MMM"), mp_EscString(GetLocaleInfoEx(LOCALE_SABBREVMONTHNAME3)));
			break;
		case 4:
			sFormat.Replace(_T("MMM"), mp_EscString(GetLocaleInfoEx(LOCALE_SABBREVMONTHNAME4)));
			break;
		case 5:
			sFormat.Replace(_T("MMM"), mp_EscString(GetLocaleInfoEx(LOCALE_SABBREVMONTHNAME5)));
			break;
		case 6:
			sFormat.Replace(_T("MMM"), mp_EscString(GetLocaleInfoEx(LOCALE_SABBREVMONTHNAME6)));
			break;
		case 7:
			sFormat.Replace(_T("MMM"), mp_EscString(GetLocaleInfoEx(LOCALE_SABBREVMONTHNAME7)));
			break;
		case 8:
			sFormat.Replace(_T("MMM"), mp_EscString(GetLocaleInfoEx(LOCALE_SABBREVMONTHNAME8)));
			break;
		case 9:
			sFormat.Replace(_T("MMM"), mp_EscString(GetLocaleInfoEx(LOCALE_SABBREVMONTHNAME9)));
			break;
		case 10:
			sFormat.Replace(_T("MMM"), mp_EscString(GetLocaleInfoEx(LOCALE_SABBREVMONTHNAME10)));
			break;
		case 11:
			sFormat.Replace(_T("MMM"), mp_EscString(GetLocaleInfoEx(LOCALE_SABBREVMONTHNAME11)));
			break;
		case 12:
			sFormat.Replace(_T("MMM"), mp_EscString(GetLocaleInfoEx(LOCALE_SABBREVMONTHNAME12)));
			break;
		}
		sFormat.Replace(_T("\\M"), _T("{~°|}"));
	}
	if (sFormat.Find(_T("MM"), 0) != -1)
	{
		//The month, from 01 through 12.
		//   6/15/2009 1:45:30 PM -> 06
		sBuff = CStr(dtExpression.GetMonth());
		if (sBuff.GetLength() == 1)
		{
			sBuff = _T("0") + sBuff;
		}
		sFormat.Replace(_T("MM"), sBuff);
	}
	if (sFormat.Find(_T("M"), 0) != -1)
	{
		//The month, from 1 through 12.
		//   6/15/2009 1:45:30 PM -> 6
		sBuff = CStr(dtExpression.GetMonth());
		sFormat.Replace(_T("M"), sBuff);

	}
	sFormat.Replace(_T("{~°|}"), _T("M"));

	////DAYS

	sFormat.Replace(_T("\\d"), _T("{~°|}"));
	if (sFormat.Find(_T("dddd"), 0) != -1)
	{
		//The full name of the day of the week. 
		//   6/15/2009 1:45:30 PM -> Monday (en-US)
		switch (dtExpression.GetDayOfWeek() - 1)
		{
		case WD_SUNDAY:
			sFormat.Replace(_T("dddd"), mp_EscString(GetLocaleInfoEx(LOCALE_SDAYNAME7)));
			break;
		case WD_MONDAY:
			sFormat.Replace(_T("dddd"), mp_EscString(GetLocaleInfoEx(LOCALE_SDAYNAME1)));
			break;
		case WD_TUESDAY:
			sFormat.Replace(_T("dddd"), mp_EscString(GetLocaleInfoEx(LOCALE_SDAYNAME2)));
			break;
		case WD_WEDNESDAY:
			sFormat.Replace(_T("dddd"), mp_EscString(GetLocaleInfoEx(LOCALE_SDAYNAME3)));
			break;
		case WD_THURSDAY:
			sFormat.Replace(_T("dddd"), mp_EscString(GetLocaleInfoEx(LOCALE_SDAYNAME4)));
			break;
		case WD_FRIDAY:
			sFormat.Replace(_T("dddd"), mp_EscString(GetLocaleInfoEx(LOCALE_SDAYNAME5)));
			break;
		case WD_SATURDAY:
			sFormat.Replace(_T("dddd"), mp_EscString(GetLocaleInfoEx(LOCALE_SDAYNAME6)));
			break;
		}
		sFormat.Replace(_T("\\d"), _T("{~°|}"));
	}
	if (sFormat.Find(_T("ddd"), 0) != -1)
	{
		//The abbreviated name of the day of the week. 
		//   6/15/2009 1:45:30 PM -> Mon (en-US)
		switch (dtExpression.GetDayOfWeek() - 1)
		{
		case WD_SUNDAY:
			sFormat.Replace(_T("ddd"), mp_EscString(GetLocaleInfoEx(LOCALE_SABBREVDAYNAME7)));
			break;
		case WD_MONDAY:
			sFormat.Replace(_T("ddd"), mp_EscString(GetLocaleInfoEx(LOCALE_SABBREVDAYNAME1)));
			break;
		case WD_TUESDAY:
			sFormat.Replace(_T("ddd"), mp_EscString(GetLocaleInfoEx(LOCALE_SABBREVDAYNAME2)));
			break;
		case WD_WEDNESDAY:
			sFormat.Replace(_T("ddd"), mp_EscString(GetLocaleInfoEx(LOCALE_SABBREVDAYNAME3)));
			break;
		case WD_THURSDAY:
			sFormat.Replace(_T("ddd"), mp_EscString(GetLocaleInfoEx(LOCALE_SABBREVDAYNAME4)));
			break;
		case WD_FRIDAY:
			sFormat.Replace(_T("ddd"), mp_EscString(GetLocaleInfoEx(LOCALE_SABBREVDAYNAME5)));
			break;
		case WD_SATURDAY:
			sFormat.Replace(_T("ddd"), mp_EscString(GetLocaleInfoEx(LOCALE_SABBREVDAYNAME6)));
			break;
		}
		sFormat.Replace(_T("\\d"), _T("{~°|}"));
	}
	if (sFormat.Find(_T("dd"), 0) != -1)
	{
		//The day of the month, from 01 through 31. 
		//  6/1/2009 1:45:30 PM -> 01
		sBuff = CStr(dtExpression.GetDay());
		if (sBuff.GetLength() == 1)
		{
			sBuff = _T("0") + sBuff;
		}
		sFormat.Replace(_T("dd"), sBuff);
	}
	if (sFormat.Find(_T("d"), 0) != -1)
	{
		//The day of the month, from 01 through 31. 
		//  6/1/2009 1:45:30 PM -> 01
		sBuff = CStr(dtExpression.GetDay());
		sFormat.Replace(_T("d"), sBuff);
	}
	sFormat.Replace(_T("{~°|}"), _T("d"));



	////DAY IN YEAR
	sFormat.Replace(_T("\\Y"), _T("{~°|}"));
	if (sFormat.Find(_T("Y"), 0) != -1)
	{
		//The number of the day in the year.
		//   1/3/2009 1:45:30 PM  -> 3
		sBuff = CStr(dtExpression.GetDayOfYear());
		sFormat.Replace(_T("Y"), sBuff);
	}
	sFormat.Replace(_T("{~°|}"), _T("Y"));

	////DAY IN WEEK
	sFormat.Replace(_T("\\W"), _T("{~°|}"));
	if (sFormat.Find(_T("W"), 0) != -1)
	{
		//The number of the day in the week.
		//   1/3/2009 1:45:30 PM  -> ??
		sBuff = CStr(dtExpression.GetDayOfWeek() - 1);
		sFormat.Replace(_T("W"), sBuff);
	}
	sFormat.Replace(_T("{~°|}"), _T("W"));

	////HOURS
	sFormat.Replace(_T("\\h"), _T("{~°|}"));
	if (sFormat.Find(_T("hh"), 0) != -1)
	{
		//The hour, using a 12-hour clock from 01 to 12.
		//   6/15/2009 1:45:30 AM -> 01
		//   6/15/2009 1:45:30 PM -> 01
		lBuff = dtExpression.GetHour();
		lBuff = lBuff;
		if ((lBuff > 12))
		{
			lBuff = lBuff - 12;
		}
		if ((lBuff == 0))
		{
			lBuff = 12;
		}
		sBuff = CStr(lBuff);
		if (sBuff.GetLength() == 1)
		{
			sBuff = _T("0") + sBuff;
		}
		sFormat.Replace(_T("hh"), sBuff);

	}
	if (sFormat.Find(_T("h"), 0) != -1)
	{
		//The hour, using a 12-hour clock from 1 to 12.
		//   6/15/2009 1:45:30 AM -> 1
		//   6/15/2009 1:45:30 PM -> 1
		lBuff = dtExpression.GetHour();
		if ((lBuff > 12))
		{
			lBuff = lBuff - 12;
		}
		if ((lBuff == 0))
		{
			lBuff = 12;
		}
		sBuff = CStr(lBuff);
		sFormat.Replace(_T("h"), sBuff);

	}
	sFormat.Replace(_T("{~°|}"), _T("h"));

	////24 HOUR
	sFormat.Replace(_T("\\H"), _T("{~°|}"));
	if (sFormat.Find(_T("HH"), 0) != -1)
	{
		//The hour, using a 24-hour clock from 00 to 23.
		//   6/15/2009 1:45:30 AM -> 01
		//   6/15/2009 1:45:30 PM -> 13
		lBuff = dtExpression.GetHour();
		sBuff = CStr(lBuff);
		if (sBuff.GetLength() == 1)
		{
			sBuff = _T("0") + sBuff;
		}
		sFormat.Replace(_T("HH"), sBuff);

	}
	if (sFormat.Find(_T("H"), 0) != -1)
	{
		//The hour, using a 24-hour clock from 0 to 23.
		//   6/15/2009 1:45:30 AM -> 1
		//   6/15/2009 1:45:30 PM -> 13
		lBuff = dtExpression.GetHour();
		sBuff = CStr(lBuff);
		sFormat.Replace(_T("H"), sBuff);
	}
	sFormat.Replace(_T("{~°|}"), _T("H"));

	////MINUTES
	sFormat.Replace(_T("\\m"), _T("{~°|}"));
	if (sFormat.Find(_T("mm"), 0) != -1)
	{
		//The minute, from 00 through 59.
		//   6/15/2009 1:09:30 AM -> 09
		//   6/15/2009 1:09:30 PM -> 09
		sBuff = CStr(dtExpression.GetMinute());
		if (sBuff.GetLength() == 1)
		{
			sBuff = _T("0") + sBuff;
		}
		sFormat.Replace(_T("mm"), sBuff);
	}
	if (sFormat.Find(_T("m"), 0) != -1)
	{
		//The minute, from 0 through 59.
		//   6/15/2009 1:09:30 AM -> 9
		//   6/15/2009 1:09:30 PM -> 9
		sBuff = CStr(dtExpression.GetMinute());
		sFormat.Replace(_T("m"), sBuff);

	}
	sFormat.Replace(_T("{~°|}"), _T("m"));

	////SECONDS
	sFormat.Replace(_T("\\s"), _T("{~°|}"));
	if (sFormat.Find(_T("ss"), 0) != -1)
	{
		//The second, from 00 through 59.
		//   6/15/2009 1:45:09 PM -> 09
		sBuff = CStr(dtExpression.GetSecond());
		if (sBuff.GetLength() == 1)
		{
			sBuff = _T("0") + sBuff;
		}
		sFormat.Replace(_T("ss"), sBuff);
	}
	if (sFormat.Find(_T("s"), 0) != -1)
	{
		//The second, from 0 through 59.
		//   6/15/2009 1:45:09 PM -> 9
		sBuff = CStr(dtExpression.GetSecond());
		sFormat.Replace(_T("s"), sBuff);
	}
	sFormat.Replace(_T("{~°|}"), _T("s"));

	////AM/PM DESIGNATOR
	sFormat.Replace(_T("\\t"), _T("{~°|}"));
	if (sFormat.Find(_T("tt"), 0) != -1)
	{
		//The AM/PM designator.
		//   6/15/2009 1:45:30 PM -> PM (en-US)
		if (dtExpression.GetHour() < 12)
		{
			sBuff = GetLocaleInfoEx(LOCALE_S1159);
		}
		else
		{
			sBuff = GetLocaleInfoEx(LOCALE_S2359);
		}
		sFormat.Replace(_T("tt"), mp_EscString(sBuff));
		sFormat.Replace(_T("\\t"), _T("{~°|}"));
	}
	if (sFormat.Find(_T("t"), 0) != -1)
	{
		//The first character of the AM/PM designator.
		//   6/15/2009 1:45:30 PM -> P (en-US)
		if (dtExpression.GetHour() < 12)
		{
			sBuff = GetLocaleInfoEx(LOCALE_S1159);
		}
		else
		{
			sBuff = GetLocaleInfoEx(LOCALE_S2359);
		}
		sBuff = sBuff.Left(1);
		sFormat.Replace(_T("t"), mp_EscString(sBuff));
		sFormat.Replace(_T("\\t"), _T("{~°|}"));
	}
	sFormat.Replace(_T("{~°|}"), _T("t"));

	////TIME ZONE
	sFormat.Replace(_T("\\K"), _T("{~°|}"));
	if (sFormat.Find(_T("K"), 0) != -1)
	{
		//Time zone information.
		//TzSpecificLocalTimeToSystemTime
		// Get the local system time.
		SYSTEMTIME LocalTime = { 0 };
		GetSystemTime(&LocalTime);

		// Get the timezone info.
		TIME_ZONE_INFORMATION TimeZoneInfo;
		GetTimeZoneInformation(&TimeZoneInfo);

		// Convert local time to UTC.
		SYSTEMTIME GmtTime = { 0 };
		TzSpecificLocalTimeToSystemTime(&TimeZoneInfo, &LocalTime, &GmtTime);

		// GMT = LocalTime + TimeZoneInfo.Bias
		// TimeZoneInfo.Bias is the difference between local time
		// and GMT in minutes.

		// Local time expressed in terms of GMT bias.
		float TimeZoneDifference = -(float(TimeZoneInfo.Bias) / 60);
		CString csLocalTimeInGmt;
		csLocalTimeInGmt.Format(_T("%2.2f"), TimeZoneDifference);
		sFormat.Replace(_T("K"), csLocalTimeInGmt);
	}
	sFormat.Replace(_T("{~°|}"), _T("K"));

	////UTC OFFSET
	sFormat.Replace(_T("\\z"), _T("{~°|}"));
	if (sFormat.Find(_T("zzz"), 0) != -1)
	{
		//Hours and minutes offset from UTC.
		//    6/15/2009 1:45:30 PM -07:00 -> -07:00
	}
	if (sFormat.Find(_T("zz"), 0) != -1)
	{
		//Hours offset from UTC, with a leading zero for a single-digit value.
		//    6/15/2009 1:45:30 PM -07:00 -> -07
	}
	if (sFormat.Find(_T("z"), 0) != -1)
	{
		//Hours offset from UTC, with no leading zeros.
		//    6/15/2009 1:45:30 PM -07:00 -> -7
	}
	sFormat.Replace(_T("{~°|}"), _T("z"));

	////PERIOD OR ERA
	sFormat.Replace(_T("\\g"), _T("{~°|}"));
	if (sFormat.Find(_T("gg"), 0) != -1)
	{
		//The period or era.
		//   6/15/2009 1:45:30 PM -> A.D.
	}
	if (sFormat.Find(_T("g"), 0) != -1)
	{
		//The period or era.
		//   6/15/2009 1:45:30 PM -> A.D.
	}
	sFormat.Replace(_T("{~°|}"), _T("g"));

	////LITERAL STRING DELIMITER
	if (sFormat.Find(_T("'"), 0) != -1)
	{
		//Literal string delimiter. 
		//   6/15/2009 1:45:30 PM ("arr:" h:m t) -> arr: 1:45 P
		//   6/15/2009 1:45:30 PM ('arr:' h:m t) -> arr: 1:45 P

	}

	////SEPARATORS
	if (sFormat.Find(_T(":"), 0) != -1)
	{
		//The time separator.
		//    6/15/2009 1:45:30 PM -> : (en-US)
		//    6/15/2009 1:45:30 PM -> . (it-IT)
		//    6/15/2009 1:45:30 PM -> : (ja-JP)
		sFormat.Replace(_T(":"), GetLocaleInfoEx(LOCALE_STIME));
	}
	if (sFormat.Find(_T("/"), 0) != -1)
	{
		//The date separator.
		//    6/15/2009 1:45:30 PM -> / (en-US)
		//    6/15/2009 1:45:30 PM -> - (ar-DZ)
		//    
		sFormat.Replace(_T("/"), GetLocaleInfoEx(LOCALE_SDATE));
	}

	////Remove Escaping
	sFormat.Replace(_T("\\"), _T(""));

	////Escape Character
	sFormat.Replace(_T("{~°|1}"), _T("\\"));

	return sFormat;

}

CString clsConfiguration::GetXML(void)
{
	clsXML oXML(mp_oControl, "Configuration");
	oXML.InitializeWriter();
    oXML.WriteProperty("FirstDayOfWeek", mp_yFirstDayOfWeek);
    oXML.WriteProperty("CalendarWeekRule", mp_yCalendarWeekRule);
    oXML.WriteProperty("ISO8601Weeks", mp_bISO8601Weeks);
    oXML.WriteProperty("RowHighlightedBackColor", mp_clrRowHighlightedBackColor);
    oXML.WriteProperty("AlternatingRowStyles", mp_bAlternatingRowStyles);
    oXML.WriteProperty("RowOddBackColor", mp_clrRowOddBackColor);
    oXML.WriteProperty("RowEvenBackColor", mp_clrRowEvenBackColor);
    oXML.WriteProperty("DefaultFont", mp_oDefaultFont);
    oXML.WriteProperty("DefaultControlStyle", mp_oControl->GetStyles()->mp_oDefaultControlStyle);
    oXML.WriteProperty("DefaultTaskStyle", mp_oControl->GetStyles()->mp_oDefaultTaskStyle);
    oXML.WriteProperty("DefaultRowStyle", mp_oControl->GetStyles()->mp_oDefaultRowStyle);
    oXML.WriteProperty("DefaultClientAreaStyle", mp_oControl->GetStyles()->mp_oDefaultClientAreaStyle);
    oXML.WriteProperty("DefaultCellStyle", mp_oControl->GetStyles()->mp_oDefaultCellStyle);
    oXML.WriteProperty("DefaultColumnStyle", mp_oControl->GetStyles()->mp_oDefaultColumnStyle);
    oXML.WriteProperty("DefaultPercentageStyle", mp_oControl->GetStyles()->mp_oDefaultPercentageStyle);
    oXML.WriteProperty("DefaultPredecessorStyle", mp_oControl->GetStyles()->mp_oDefaultPredecessorStyle);
    oXML.WriteProperty("DefaultTimeLineStyle", mp_oControl->GetStyles()->mp_oDefaultTimeLineStyle);
    oXML.WriteProperty("DefaultTimeBlockStyle", mp_oControl->GetStyles()->mp_oDefaultTimeBlockStyle);
    oXML.WriteProperty("DefaultTickMarkAreaStyle", mp_oControl->GetStyles()->mp_oDefaultTickMarkAreaStyle);
    oXML.WriteProperty("DefaultSplitterStyle", mp_oControl->GetStyles()->mp_oDefaultSplitterStyle);
    oXML.WriteProperty("DefaultProgressLineStyle", mp_oControl->GetStyles()->mp_oDefaultProgressLineStyle);
    oXML.WriteProperty("DefaultNodeStyle", mp_oControl->GetStyles()->mp_oDefaultNodeStyle);
    oXML.WriteProperty("DefaultTierStyle", mp_oControl->GetStyles()->mp_oDefaultTierStyle);
    oXML.WriteProperty("DefaultScrollBarStyle", mp_oControl->GetStyles()->mp_oDefaultScrollBarStyle);
    oXML.WriteProperty("DefaultSBSeparatorStyle", mp_oControl->GetStyles()->mp_oDefaultSBSeparatorStyle);
    oXML.WriteProperty("DefaultSBNormalStyle", mp_oControl->GetStyles()->mp_oDefaultSBNormalStyle);
    oXML.WriteProperty("DefaultSBPressedStyle", mp_oControl->GetStyles()->mp_oDefaultSBPressedStyle);
    oXML.WriteProperty("DefaultSBDisabledStyle", mp_oControl->GetStyles()->mp_oDefaultSBDisabledStyle);
	return oXML.GetXML();
}

void clsConfiguration::SetXML(CString sXML)
{
	clsXML oXML(mp_oControl, "Configuration");
	oXML.SetXML(sXML);
	oXML.InitializeReader();
    oXML.ReadProperty("FirstDayOfWeek", mp_yFirstDayOfWeek);
    oXML.ReadProperty("CalendarWeekRule", mp_yCalendarWeekRule);
    oXML.ReadProperty("ISO8601Weeks", mp_bISO8601Weeks);
    oXML.ReadProperty("RowHighlightedBackColor", mp_clrRowHighlightedBackColor);
    oXML.ReadProperty("AlternatingRowStyles", mp_bAlternatingRowStyles);
    oXML.ReadProperty("RowOddBackColor", mp_clrRowOddBackColor);
    oXML.ReadProperty("RowEvenBackColor", mp_clrRowEvenBackColor);
    oXML.ReadProperty("DefaultFont", mp_oDefaultFont);
    oXML.ReadProperty("DefaultControlStyle", mp_oControl->GetStyles()->mp_oDefaultControlStyle);
    oXML.ReadProperty("DefaultTaskStyle", mp_oControl->GetStyles()->mp_oDefaultTaskStyle);
    oXML.ReadProperty("DefaultRowStyle", mp_oControl->GetStyles()->mp_oDefaultRowStyle);
    oXML.ReadProperty("DefaultClientAreaStyle", mp_oControl->GetStyles()->mp_oDefaultClientAreaStyle);
    oXML.ReadProperty("DefaultCellStyle", mp_oControl->GetStyles()->mp_oDefaultCellStyle);
    oXML.ReadProperty("DefaultColumnStyle", mp_oControl->GetStyles()->mp_oDefaultColumnStyle);
    oXML.ReadProperty("DefaultPercentageStyle", mp_oControl->GetStyles()->mp_oDefaultPercentageStyle);
    oXML.ReadProperty("DefaultPredecessorStyle", mp_oControl->GetStyles()->mp_oDefaultPredecessorStyle);
    oXML.ReadProperty("DefaultTimeLineStyle", mp_oControl->GetStyles()->mp_oDefaultTimeLineStyle);
    oXML.ReadProperty("DefaultTimeBlockStyle", mp_oControl->GetStyles()->mp_oDefaultTimeBlockStyle);
    oXML.ReadProperty("DefaultTickMarkAreaStyle", mp_oControl->GetStyles()->mp_oDefaultTickMarkAreaStyle);
    oXML.ReadProperty("DefaultSplitterStyle", mp_oControl->GetStyles()->mp_oDefaultSplitterStyle);
    oXML.ReadProperty("DefaultProgressLineStyle", mp_oControl->GetStyles()->mp_oDefaultProgressLineStyle);
    oXML.ReadProperty("DefaultNodeStyle", mp_oControl->GetStyles()->mp_oDefaultNodeStyle);
    oXML.ReadProperty("DefaultTierStyle", mp_oControl->GetStyles()->mp_oDefaultTierStyle);
    oXML.ReadProperty("DefaultScrollBarStyle", mp_oControl->GetStyles()->mp_oDefaultScrollBarStyle);
    oXML.ReadProperty("DefaultSBSeparatorStyle", mp_oControl->GetStyles()->mp_oDefaultSBSeparatorStyle);
    oXML.ReadProperty("DefaultSBNormalStyle", mp_oControl->GetStyles()->mp_oDefaultSBNormalStyle);
    oXML.ReadProperty("DefaultSBPressedStyle", mp_oControl->GetStyles()->mp_oDefaultSBPressedStyle);
    oXML.ReadProperty("DefaultSBDisabledStyle", mp_oControl->GetStyles()->mp_oDefaultSBDisabledStyle);
}