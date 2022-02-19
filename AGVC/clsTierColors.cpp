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
#include "clsTierColors.h"

IMPLEMENT_DYNCREATE(clsTierColors, CCmdTarget)

//{DC6F9823-0DA9-4B39-94DC-004351CFA082}
static const IID IID_IclsTierColors = { 0xDC6F9823, 0x0DA9, 0x4B39, { 0x94, 0xDC, 0x00, 0x43, 0x51, 0xCF, 0xA0, 0x82} };

//{40C39476-CC07-45B5-A768-91EC4E8C21D8}
IMPLEMENT_OLECREATE_FLAGS(clsTierColors, "AGVC.clsTierColors", afxRegApartmentThreading, 0x40C39476, 0xCC07, 0x45B5, 0xA7, 0x68, 0x91, 0xEC, 0x4E, 0x8C, 0x21, 0xD8)

BEGIN_DISPATCH_MAP(clsTierColors, CCmdTarget)
	DISP_PROPERTY_EX_ID(clsTierColors, "Count", 1, odl_GetCount, SetNotSupported, VT_I4) 
	DISP_PROPERTY_PARAM_ID(clsTierColors, "Item", 2, odl_Item, SetNotSupported, VT_DISPATCH, VTS_BSTR) 
	DISP_FUNCTION_ID(clsTierColors, "GetXML", 3, odl_GetXML, VT_BSTR, VTS_NONE) 
	DISP_FUNCTION_ID(clsTierColors, "SetXML", 4, odl_SetXML, VT_EMPTY, VTS_BSTR)
	DISP_DEFVALUE(clsTierColors, "Item")
END_DISPATCH_MAP()

BEGIN_INTERFACE_MAP(clsTierColors, CCmdTarget)
	INTERFACE_PART(clsTierColors, IID_IclsTierColors, Dispatch)
END_INTERFACE_MAP()

BEGIN_MESSAGE_MAP(clsTierColors, CCmdTarget)
END_MESSAGE_MAP()

clsTierColors::clsTierColors(CActiveGanttVCCtl* oControl, LONG yTierType)
{
	EnableAutomation();
	AfxOleLockApp();
	mp_oControl = oControl;
	mp_oCollection = new clsCollectionBase(mp_oControl, "TierColor");
	mp_yTierType = yTierType;
}

clsTierColors::clsTierColors()
{
	AfxMessageBox(_T("Invalid constructor has been invoked for clsTierColors. Please inform support@sourcecodestore.com of this error"), MB_OK | MB_ICONSTOP | MB_APPLMODAL);
}

clsTierColors::~clsTierColors()
{
	delete mp_oCollection;
	mp_oCollection = NULL;
	AfxOleUnlockApp();
}

void clsTierColors::OnFinalRelease()
{	
	CCmdTarget::OnFinalRelease();
}

LONG clsTierColors::GetCount(void)
{
	return mp_oCollection->m_lCount();
}

clsCollectionBase* clsTierColors::GetCollection(void)
{
	return mp_oCollection;
}


void clsTierColors::Finalize(void)
{
}

clsTierColor* clsTierColors::Item(CString Index)
{
	clsTierColor *oTierColor;
	oTierColor = (clsTierColor*)mp_oCollection->m_oItem(Index, TIERCOLORS_ITEM_1, TIERCOLORS_ITEM_2, TIERCOLORS_ITEM_3, TIERCOLORS_ITEM_4);
	return oTierColor;
}

clsTierColor* clsTierColors::Add(Gdiplus::Color BackColor, Gdiplus::Color ForeColor, Gdiplus::Color StartGradientColor, Gdiplus::Color EndGradientColor, Gdiplus::Color HatchBackColor, Gdiplus::Color HatchForeColor, CString Key)
{
	mp_oCollection->SetAddMode(TRUE);
	clsTierColor* oTierColor = new clsTierColor(mp_oControl, this);
	oTierColor->SetBackColor(BackColor);
	oTierColor->SetForeColor(ForeColor);
	oTierColor->SetStartGradientColor(StartGradientColor);
	oTierColor->SetEndGradientColor(EndGradientColor);
	oTierColor->SetHatchBackColor(HatchBackColor);
	oTierColor->SetHatchForeColor(HatchForeColor);
	oTierColor->SetKey(Key);
	mp_oCollection->m_Add(oTierColor, Key, TIERCOLORS_ADD_1, TIERCOLORS_ADD_2, FALSE, TIERCOLORS_ADD_3);
	return oTierColor;
}

clsTierColor* clsTierColors::Add(Gdiplus::Color BackColor, Gdiplus::Color ForeColor, Gdiplus::Color StartGradientColor, Gdiplus::Color EndGradientColor, Gdiplus::Color HatchBackColor, Gdiplus::Color HatchForeColor)
{
	return Add(BackColor, ForeColor, StartGradientColor, EndGradientColor, HatchBackColor, HatchForeColor, "");
}

CString clsTierColors::mp_CollectionName(void)
{
	if (mp_yTierType == ST_MICROSECOND)
	{
		return _T("MicrosecondColors");
	}
	else if (mp_yTierType == ST_MILLISECOND)
	{
		return _T("MillisecondColors");
	}
	else if (mp_yTierType == ST_SECOND)
	{
		return _T("SecondColors");
	}
	else if (mp_yTierType == ST_MINUTE)
	{
		return _T("MinuteColors");
	}
	else if (mp_yTierType == ST_HOUR)
	{
		return _T("HourColors");
	}
	else if (mp_yTierType == ST_DAY)
	{
		return _T("DayColors");
	}
	else if (mp_yTierType == ST_DAYOFWEEK)
	{
		return _T("DayOfWeekColors");
	}
	else if (mp_yTierType == ST_DAYOFYEAR)
	{
		return _T("DayOfYearColors");
	}
	else if (mp_yTierType == ST_WEEK)
	{
		return _T("WeekColors");
	}
	else if (mp_yTierType == ST_MONTH)
	{
		return _T("MonthColors");
	}
	else if (mp_yTierType == ST_QUARTER)
	{
		return _T("QuarterColors");
	}
	else if (mp_yTierType == ST_YEAR)
	{
		return _T("YearColors");
	}
	return _T("");
}

CString clsTierColors::GetXML(void)
{
	LONG lIndex;
	clsTierColor* oTierColor; 
	clsXML oXML(mp_oControl, mp_CollectionName());
	oXML.InitializeWriter();
	for (lIndex = 1; lIndex <= GetCount(); lIndex++)
	{
		oTierColor = (clsTierColor*) mp_oCollection->m_oReturnArrayElement(lIndex);
		oXML.WriteObject(oTierColor->GetXML());
	}
	return oXML.GetXML();
}

void clsTierColors::SetXML(CString sXML)
{
	LONG lIndex;
	clsXML oXML(mp_oControl, mp_CollectionName());
	oXML.SetXML(sXML);
	oXML.InitializeReader();
	mp_oCollection->m_Clear();
	for (lIndex = 1;lIndex <= oXML.ReadCollectionCount();lIndex++)
	{
		clsTierColor* oTierColor = new clsTierColor(mp_oControl, this);
		oTierColor->SetXML(oXML.ReadCollectionObject(lIndex));
		mp_oCollection->SetAddMode(TRUE);
		mp_oCollection->m_Add(oTierColor, oTierColor->mp_sKey, TIERCOLORS_ADD_1, TIERCOLORS_ADD_2, FALSE, TIERCOLORS_ADD_3);
	}
}

void clsTierColors::Clear()
{
	mp_oCollection->m_Clear();
}

void clsTierColors::Clone(clsTierColors* oClone)
{
    oClone->Clear();
    int i;
	for (i = 1; i <= GetCount(); i++)
	{
        clsTierColor* oTierColor = Item(CStr(i));
        clsTierColor* oTierColorClone;
        oTierColorClone = oClone->Add(oTierColor->GetBackColor(), oTierColor->GetForeColor(), oTierColor->GetStartGradientColor(), oTierColor->GetEndGradientColor(), oTierColor->GetHatchBackColor(), oTierColor->GetHatchForeColor(), oTierColor->GetKey());
        oTierColor->Clone(oTierColorClone);
	}
}