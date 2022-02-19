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
#include "clsXML.h"
#include "Duration.h"

IMPLEMENT_DYNCREATE(Duration, CCmdTarget)

static const IID IID_IDuration = {0xc46b0f9e,0xd1d6,0x4f95,{0xb8,0xf5,0x80,0x2f,0xf3,0x49,0x10,0x3b}};
IMPLEMENT_OLECREATE_FLAGS(Duration, "MSP2010.Duration", afxRegApartmentThreading,0xc6827a97,0x0c1b,0x4a34,0xae,0x67,0x41,0xe8,0x8e,0x89,0xfb,0xb0)

BEGIN_DISPATCH_MAP(Duration, CCmdTarget)
	DISP_FUNCTION_ID(Duration, "IsNull", 1, IsNull, VT_BOOL, VTS_NONE)
	DISP_FUNCTION_ID(Duration, "ToString", 2, odl_ToString, VT_BSTR, VTS_NONE)
	DISP_FUNCTION_ID(Duration, "FromString", 3, odl_FromString, VT_EMPTY, VTS_BSTR)
	DISP_PROPERTY_EX_ID(Duration, "Year", 4, GetYear, SetYear, VT_I4)
	DISP_PROPERTY_EX_ID(Duration, "Month", 5, GetMonth, SetMonth, VT_I4)
	DISP_PROPERTY_EX_ID(Duration, "Day", 6, GetDay, SetDay, VT_I4)
	DISP_PROPERTY_EX_ID(Duration, "Hour", 7, GetHour, SetHour, VT_I4)
	DISP_PROPERTY_EX_ID(Duration, "Minute", 8, GetMinute, SetMinute, VT_I4)
	DISP_PROPERTY_EX_ID(Duration, "Second", 9, GetSecond, SetSecond, VT_I4)
END_DISPATCH_MAP()

BEGIN_INTERFACE_MAP(Duration, CCmdTarget)
	INTERFACE_PART(Duration, IID_IDuration, Dispatch)
END_INTERFACE_MAP()

BEGIN_MESSAGE_MAP(Duration, CCmdTarget)
END_MESSAGE_MAP()

Duration::Duration()
{
	EnableAutomation();
	AfxOleLockApp();
	mp_lYear = 0;
	mp_lMonth = 0;
	mp_lDay = 0;
	mp_lHour = 0;
	mp_lMinute = 0;
	mp_lSecond = 0;
}

Duration::~Duration()
{
	AfxOleUnlockApp();
}

void Duration::OnFinalRelease()
{
	CCmdTarget::OnFinalRelease();
}

BOOL Duration::IsNull(void)
{
    bool bReturn = TRUE;
    if (mp_lYear != 0)
    {
        bReturn = FALSE;
    }
    if (mp_lMonth != 0)
    {
        bReturn = FALSE;
    }
    if (mp_lDay != 0)
    {
        bReturn = FALSE;
    }
    if (mp_lHour != 0)
    {
        bReturn = FALSE;
    }
    if (mp_lMinute != 0)
    {
        bReturn = FALSE;
    }
    if (mp_lSecond != 0)
    {
        bReturn = FALSE;
    }
    return bReturn;
}

CString Duration::ToString(void)
{
	CString sReturn = "P";
	CString sYear;
	sYear.Format(_T("%d"), mp_lYear);
	CString sMonth;
	sMonth.Format(_T("%d"), mp_lMonth);
	CString sDay;
	sDay.Format(_T("%d"), mp_lDay);
	CString sHour;
	sHour.Format(_T("%d"), mp_lHour);
	CString sMinute;
	sMinute.Format(_T("%d"), mp_lMinute);
	CString sSecond;
	sSecond.Format(_T("%d"), mp_lSecond);
    if (mp_lYear != 0)
    {
        sReturn = sReturn + sYear + "Y";
    }
    if (mp_lMonth != 0)
    {
        sReturn = sReturn + sMonth + "M";
    }
    if (mp_lDay != 0)
    {
        sReturn = sReturn + sDay + "D";
    }
    sReturn = sReturn + "T";
    sReturn = sReturn + sHour + "H";
    sReturn = sReturn + sMinute + "M";
    sReturn = sReturn + sSecond + "S";
    return sReturn;
}

void Duration::FromString(CString sString)
{
	CString sTime = _T("");
	CString sDate = _T("");
	CString sBuff = _T("");
	if (sString.Left(1) != _T("P") || sString.GetLength() == 0)
	{
		return;
	}
	if (sString.Find(_T("T")) > -1)
	{
		sTime = sString.Mid(sString.Find(_T("T")) + 1, sString.GetLength() - sString.Find(_T("T")) - 1);
		sDate = sString;
		sDate.Replace(_T("T") + sTime, _T(""));
	}
	else
	{
		sTime = "";
		sDate = sString;
	}
	sDate = sDate.Mid(1, sDate.GetLength() - 1);
	if (sTime.GetLength() > 0)
	{
		if (sTime.Find(_T("H")) > -1)
		{
			sBuff = sTime.Mid(0, sTime.Find(_T("H")));
			sTime.Replace(sBuff + _T("H"), _T(""));
			mp_lHour = _wtoi(sBuff);
		}
		if (sTime.Find(_T("M")) > -1)
		{
			sBuff = sTime.Mid(0, sTime.Find(_T("M")));
			sTime.Replace(sBuff + _T("M"), _T(""));
			mp_lMinute = _wtoi(sBuff);
		}
		if (sTime.Find(_T("S")) > -1)
		{
			sBuff = sTime.Mid(0, sTime.Find(_T("S")));
			sTime.Replace(sBuff + _T("S"), _T(""));
			mp_lSecond = _wtoi(sBuff);
		}
	}
	if (sDate.GetLength() > 0)
	{
		if (sDate.Find(_T("Y")) > -1)
		{
			sBuff = sTime.Mid(0, sTime.Find(_T("Y")));
			sTime.Replace(sBuff + _T("Y"), _T(""));
			mp_lYear = _wtoi(sBuff);
		}
		if (sDate.Find(_T("M")) > -1)
		{
			sBuff = sTime.Mid(0, sTime.Find(_T("M")));
			sTime.Replace(sBuff + _T("M"), _T(""));
			mp_lMonth = _wtoi(sBuff);
		}
		if (sDate.Find(_T("D")) > -1)
		{
			sBuff = sTime.Mid(0, sTime.Find(_T("D")));
			sTime.Replace(sBuff + _T("D"), _T(""));
			mp_lDay = _wtoi(sBuff);
		}
	}
}

LONG Duration::GetYear(void)
{
	return mp_lYear;
}

void Duration::SetYear(LONG nNewValue)
{
	mp_lYear = nNewValue;
}

LONG Duration::GetMonth(void)
{
	return mp_lMonth;
}

void Duration::SetMonth(LONG nNewValue)
{
	mp_lMonth = nNewValue;
}

LONG Duration::GetDay(void)
{
	return mp_lDay;
}

void Duration::SetDay(LONG nNewValue)
{
	mp_lDay = nNewValue;
}

LONG Duration::GetHour(void)
{
	return mp_lHour;
}

void Duration::SetHour(LONG nNewValue)
{
	mp_lHour = nNewValue;
}

LONG Duration::GetMinute(void)
{
	return mp_lMinute;
}

void Duration::SetMinute(LONG nNewValue)
{
	mp_lMinute = nNewValue;
}

LONG Duration::GetSecond(void)
{
	return mp_lSecond;
}

void Duration::SetSecond(LONG nNewValue)
{
	mp_lSecond = nNewValue;
}

BSTR Duration::odl_ToString(void)
{
	return ToString().AllocSysString();
}

void Duration::odl_FromString(LPCTSTR sString)
{
	FromString(sString);
}