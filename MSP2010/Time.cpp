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
#include "Time.h"

IMPLEMENT_DYNCREATE(Time, CCmdTarget)

static const IID IID_ITime = {0x0b13fc95,0xaaba,0x4a0f,{0x8a,0xe7,0xd4,0x06,0x3f,0x2d,0x0a,0x28}};
IMPLEMENT_OLECREATE_FLAGS(Time, "MSP2010.Time", afxRegApartmentThreading,0xb0d62f8e,0x3c3e,0x4bba,0x8a,0x86,0x27,0x50,0x9d,0x27,0x7c,0xad)

BEGIN_DISPATCH_MAP(Time, CCmdTarget)
	DISP_FUNCTION_ID(Time, "IsNull", 1, IsNull, VT_BOOL, VTS_NONE)
	DISP_FUNCTION_ID(Time, "ToString", 2, odl_ToString, VT_BSTR, VTS_NONE)
	DISP_FUNCTION_ID(Time, "FromString", 3, odl_FromString, VT_EMPTY, VTS_BSTR)
	DISP_PROPERTY_EX_ID(Time, "Hour", 4, GetHour, SetHour, VT_I4)
	DISP_PROPERTY_EX_ID(Time, "Minute", 5, GetMinute, SetMinute, VT_I4)
	DISP_PROPERTY_EX_ID(Time, "Second", 6, GetSecond, SetSecond, VT_I4)

END_DISPATCH_MAP()

BEGIN_INTERFACE_MAP(Time, CCmdTarget)
	INTERFACE_PART(Time, IID_ITime, Dispatch)
END_INTERFACE_MAP()

BEGIN_MESSAGE_MAP(Time, CCmdTarget)
END_MESSAGE_MAP()

Time::Time()
{
	EnableAutomation();
	AfxOleLockApp();
	mp_lHour = 0;
	mp_lMinute = 0;
	mp_lSecond = 0;
}

Time::~Time()
{
	AfxOleUnlockApp();
}

void Time::OnFinalRelease()
{
	CCmdTarget::OnFinalRelease();
}

BOOL Time::IsNull(void)
{
    BOOL bReturn = TRUE;
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

CString Time::ToString()
{ 
	CString sHour;
	sHour.Format(_T("%02d"), mp_lHour);
	CString sMinute;
	sMinute.Format(_T("%02d"), mp_lMinute);
	CString sSecond;
	sSecond.Format(_T("%02d"), mp_lSecond);
	return sHour + ":" + sMinute + ":" + sSecond;
}

void Time::FromString(CString sString)
{
	if (sString.GetLength() == 8)
	{
		mp_lHour = _wtoi(sString.Mid(0, 2));
		mp_lMinute = _wtoi(sString.Mid(3, 2));
		mp_lSecond = _wtoi(sString.Mid(6, 2));
	}
	else
	{
		mp_lHour = 0;
		mp_lMinute = 0;
		mp_lSecond = 0;
	}
}

COleDateTime Time::ToDateTime()
{
	COleDateTime dtReturn = (DATE)0;
	dtReturn.SetTime(mp_lHour, mp_lMinute, mp_lSecond);
	return dtReturn;
}

void Time::FromDateTime(COleDateTime dtDate)
{
	mp_lHour = dtDate.GetHour();
	mp_lMinute = dtDate.GetMinute();
	mp_lSecond = dtDate.GetSecond();
}

LONG Time::GetHour()
{
	return mp_lHour;
}

void Time::SetHour(LONG nNewValue)
{
	mp_lHour = nNewValue;
}

LONG Time::GetMinute()
{
	return mp_lMinute;
}

void Time::SetMinute(LONG nNewValue)
{
	mp_lMinute = nNewValue;
}

LONG Time::GetSecond()
{
	return mp_lSecond;
}

void Time::SetSecond(LONG nNewValue)
{
	mp_lSecond = nNewValue;
}

BSTR Time::odl_ToString(void)
{
	return ToString().AllocSysString();
}

void Time::odl_FromString(LPCTSTR sString)
{
	FromString(sString);
}