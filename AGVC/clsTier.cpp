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
#include "clsTier.h"



IMPLEMENT_DYNCREATE(clsTier, CCmdTarget)

	//{B3F99910-DCE3-4B7A-B550-D8994835D475}
	static const IID IID_IclsTier = { 0xB3F99910, 0xDCE3, 0x4B7A, { 0xB5, 0x50, 0xD8, 0x99, 0x48, 0x35, 0xD4, 0x75} };

//{D28E88BF-9715-4A22-B299-583359927F2B}
IMPLEMENT_OLECREATE_FLAGS(clsTier, "AGVC.clsTier", afxRegApartmentThreading, 0xD28E88BF, 0x9715, 0x4A22, 0xB2, 0x99, 0x58, 0x33, 0x59, 0x92, 0x7F, 0x2B)

	BEGIN_DISPATCH_MAP(clsTier, CCmdTarget)
		DISP_PROPERTY_EX_ID(clsTier, "Visible", 1, odl_GetVisible, odl_SetVisible, VT_BOOL) 
		DISP_PROPERTY_EX_ID(clsTier, "Tag", 2, odl_GetTag, odl_SetTag, VT_BSTR) 
		DISP_PROPERTY_EX_ID(clsTier, "Interval", 3, odl_GetInterval, odl_SetInterval, VT_I4) 
		DISP_PROPERTY_EX_ID(clsTier, "TierType", 4, odl_GetTierType, odl_SetTierType, VT_I4) 
		DISP_PROPERTY_EX_ID(clsTier, "Height", 5, odl_GetHeight, odl_SetHeight, VT_I4) 
		DISP_FUNCTION_ID(clsTier, "GetXML", 6, odl_GetXML, VT_BSTR, VTS_NONE) 
		DISP_FUNCTION_ID(clsTier, "SetXML", 7, odl_SetXML, VT_EMPTY, VTS_BSTR)
		DISP_PROPERTY_EX_ID(clsTier, "Factor", 8, odl_GetFactor, odl_SetFactor, VT_I4)
		DISP_PROPERTY_EX_ID(clsTier, "StyleIndex", 9, odl_GetStyleIndex, odl_SetStyleIndex, VT_BSTR) 
		DISP_PROPERTY_EX_ID(clsTier, "Style", 10, odl_GetStyle, SetNotSupported, VT_DISPATCH)
		DISP_PROPERTY_EX_ID(clsTier, "BackgroundMode", 11, odl_GetBackgroundMode, odl_SetBackgroundMode, VT_I4)

	END_DISPATCH_MAP()

	BEGIN_INTERFACE_MAP(clsTier, CCmdTarget)
		INTERFACE_PART(clsTier, IID_IclsTier, Dispatch)
	END_INTERFACE_MAP()

	BEGIN_MESSAGE_MAP(clsTier, CCmdTarget)
	END_MESSAGE_MAP()

	clsTier::clsTier(CActiveGanttVCCtl* oControl, clsTierArea* oTierArea,LONG yTierPosition)
	{
		EnableAutomation();
		AfxOleLockApp();
		mp_oControl = oControl;
		mp_oTierArea = oTierArea;
		mp_yTierPosition = yTierPosition;
		switch (mp_yTierPosition) 
		{
		case SP_UPPER:
			mp_sTierPosition = "UpperTier";
			break;
		case SP_MIDDLE:
			mp_sTierPosition = "MiddleTier";
			break;
		case SP_LOWER:
			mp_sTierPosition = "LowerTier";
			break;
		}
		Clear();
	}

	clsTier::clsTier()
	{
		AfxMessageBox(_T("Invalid constructor has been invoked for clsTier. Please inform support@sourcecodestore.com of this error"), MB_OK | MB_ICONSTOP | MB_APPLMODAL);
	}

	clsTier::~clsTier()
	{
		AfxOleUnlockApp();
	}

	void clsTier::OnFinalRelease()
	{
		CCmdTarget::OnFinalRelease();
	}

	BOOL clsTier::GetVisible(void)
	{
        if (mp_yTierType == ST_NOT_VISIBLE)
        {
            return FALSE;
        }
        else
        {
            return mp_bVisible;
        }
	}

	void clsTier::SetVisible(BOOL Value)
	{
		mp_bVisible = Value;
	}

	CString clsTier::GetTag(void)
	{
		return mp_sTag;
	}

	void clsTier::SetTag(CString Value)
	{
		mp_sTag = Value;
	}

	LONG clsTier::GetInterval(void)
	{
		return mp_yInterval;
	}

	void clsTier::SetInterval(LONG Value)
	{
		mp_yInterval = Value;
		mp_yTierType = ST_CUSTOM;
	}

	LONG clsTier::GetFactor(void)
	{
		return mp_lFactor;
	}

	void clsTier::SetFactor(LONG Value)
	{
		mp_lFactor = Value;
	}

	LONG clsTier::GetTierType(void)
	{
		return mp_yTierType;
	}

	void clsTier::SetTierType(LONG Value)
	{
		mp_yTierType = Value;
		switch (mp_yTierType) 
		{
        case ST_NOT_VISIBLE:
            mp_yInterval = IL_HOUR;
            mp_lFactor = 1;
            break;
		case ST_YEAR:
			mp_yInterval = IL_YEAR;
			mp_lFactor = 1;
			break;
		case ST_QUARTER:
			mp_yInterval = IL_QUARTER;
			mp_lFactor = 1;
			break;
		case ST_MONTH:
			mp_yInterval = IL_MONTH;
			mp_lFactor = 1;
			break;
		case ST_WEEK:
			mp_yInterval = IL_WEEK;
			mp_lFactor = 1;
			break;
		case ST_DAYOFWEEK:
			mp_yInterval = IL_DAY;
			mp_lFactor = 1;
			break;
		case ST_DAY:
			mp_yInterval = IL_DAY;
			mp_lFactor = 1;
			break;
		case ST_DAYOFYEAR:
			mp_yInterval = IL_DAY;
			mp_lFactor = 1;
			break;
		case ST_HOUR:
			mp_yInterval = IL_HOUR;
			mp_lFactor = 1;
			break;
		case ST_MINUTE:
			mp_yInterval = IL_MINUTE;
			mp_lFactor = 1;
			break;
		case ST_SECOND:
			mp_yInterval = IL_SECOND;
			mp_lFactor = 1;
			break;
		}
	}

	LONG clsTier::GetHeight(void)
	{
		return mp_lHeight;
	}

	void clsTier::SetHeight(LONG Value)
	{
		mp_lHeight = Value;
	}

	void clsTier::Position(void)
	{
		if (GetVisible() == FALSE) 
		{
			return;
		}
		COleDateTime dtStart;
		COleDateTime dtEnd;
		LONG lTop = 0;
		LONG lBottom = 0;
		LONG lTierHeight = 0;
		lTierHeight = GetHeight();
		lTop = mp_oTierArea->GetTimeLine()->TiersTickMarksPosition(mp_sTierPosition);
		lBottom = lTop + lTierHeight - 1;
		if ((mp_oControl->GetMathLib()->GetXCoordinateFromDate(mp_oControl->GetMathLib()->DateTimeAdd(mp_yInterval, mp_lFactor, mp_oTierArea->GetTimeLine()->GetStartDate())) - mp_oControl->GetMathLib()->GetXCoordinateFromDate(mp_oTierArea->GetTimeLine()->GetStartDate()) > 5)) 
		{
			dtEnd = mp_oControl->GetMathLib()->RoundDate(mp_yInterval, mp_lFactor, mp_oTierArea->GetTimeLine()->GetStartDate());
			if ((mp_oControl->GetMathLib()->GetXCoordinateFromDate(dtEnd) >= mp_oTierArea->GetTimeLine()->Getf_lStart())) 
			{
				dtStart = mp_oControl->GetMathLib()->DateTimeAdd(mp_yInterval, -mp_lFactor, dtEnd);
				dtStart = mp_oControl->GetMathLib()->RoundDate(mp_yInterval, mp_lFactor, dtStart);
				Draw(dtStart, dtEnd, lTop, lBottom);
			}
			while ((dtEnd < mp_oTierArea->GetTimeLine()->GetEndDate())) 
			{
				dtStart = dtEnd;
				dtEnd = mp_oControl->GetMathLib()->DateTimeAdd(mp_yInterval, mp_lFactor, dtEnd);
				dtStart = mp_oControl->GetMathLib()->RoundDate(mp_yInterval, mp_lFactor, dtStart);
				dtEnd = mp_oControl->GetMathLib()->RoundDate(mp_yInterval, mp_lFactor, dtEnd);
				Draw(dtStart, dtEnd, lTop, lBottom);
			}
		}
	}

	void clsTier::Draw(COleDateTime dtStart,COleDateTime dtEnd,LONG lTop,LONG lBottom)
	{
		LONG lStart = 0;
		LONG lEnd = 0;
		LONG lStartTrim = 0;
		LONG lEndTrim = 0;

		CString sText = "";
		LONG lTextWidth = 0;

		Color clrForeColor = Color::White;
		Color clrBackColor = Color::White;
		Color clrStartGradientColor = Color::White;
		Color clrEndGradientColor = Color::White;
		Color clrHatchBackColor = Color::White;
		Color clrHatchForeColor = Color::White;


		lStart = mp_oControl->GetMathLib()->GetXCoordinateFromDate(dtStart);
		lEnd = mp_oControl->GetMathLib()->GetXCoordinateFromDate(dtEnd) - 1;
		if ((lStart < mp_oTierArea->GetTimeLine()->Getf_lStart())) 
		{
			lStartTrim = mp_oTierArea->GetTimeLine()->Getf_lStart();
		}
		else 
		{
			lStartTrim = lStart;
		}
		if ((lEnd > mp_oTierArea->GetTimeLine()->Getf_lEnd())) 
		{
			lEndTrim = mp_oTierArea->GetTimeLine()->Getf_lEnd();
		}
		else 
		{
			lEndTrim = lEnd;
		}
		sText = "";
		if ((mp_yTierType == ST_CUSTOM)) 
		{
			mp_oControl->GetCustomTierDrawEventArgs()->Clear();
			mp_oControl->GetCustomTierDrawEventArgs()->SetTierPosition(mp_yTierPosition);
			mp_oControl->GetCustomTierDrawEventArgs()->SetStartDate(dtStart);
			mp_oControl->GetCustomTierDrawEventArgs()->SetEndDate(dtEnd);
			mp_oControl->GetCustomTierDrawEventArgs()->SetLeft(lStart);
			mp_oControl->GetCustomTierDrawEventArgs()->SetTop(lTop);
			mp_oControl->GetCustomTierDrawEventArgs()->SetRight(lEnd);
			mp_oControl->GetCustomTierDrawEventArgs()->SetBottom(lBottom);
			mp_oControl->GetCustomTierDrawEventArgs()->SetLeftTrim(lStartTrim);
			mp_oControl->GetCustomTierDrawEventArgs()->SetRightTrim(lEndTrim);
			mp_oControl->GetCustomTierDrawEventArgs()->SetText(sText);
			mp_oControl->GetCustomTierDrawEventArgs()->SetStyleIndex("");
			mp_oControl->GetCustomTierDrawEventArgs()->SetInterval(GetInterval());
			mp_oControl->GetCustomTierDrawEventArgs()->SetFactor(GetFactor());
			mp_oControl->FireCustomTierDraw();
			CString sTest1 = mp_oControl->GetCustomTierDrawEventArgs()->GetStyleIndex();
			if (!(mp_oControl->GetCustomTierDrawEventArgs()->GetStyleIndex().GetLength()== 0)) 
			{	
				mp_oControl->clsG->mp_DrawItem(lStart, lTop, lEnd, lBottom, mp_oControl->GetCustomTierDrawEventArgs()->GetStyleIndex(), mp_oControl->GetCustomTierDrawEventArgs()->GetText(), FALSE, NULL, lStartTrim, lEndTrim, NULL);
			}
			else
			{
				mp_oControl->clsG->mp_DrawItem(lStart, lTop, lEnd, lBottom, "", mp_oControl->GetCustomTierDrawEventArgs()->GetText(), FALSE, NULL, lStartTrim, lEndTrim, mp_oControl->GetConfiguration()->GetDefaultTierStyle());
			}
		}
		else 
		{

			clsTierAppearance* oTierAppearance = NULL;
			clsTierFormat* oTierFormat = NULL;
			if ((mp_oControl->GetTierAppearanceScope() == OS_CONTROL)) 
			{
				oTierAppearance = mp_oControl->GetTierAppearance();
			} 
			else if ((mp_oControl->GetTierAppearanceScope() == OS_VIEW)) 
			{
				oTierAppearance = mp_oTierArea->GetTierAppearance();
			}
			if ((mp_oControl->GetTierFormatScope() == OS_CONTROL)) 
			{
				oTierFormat = mp_oControl->GetTierFormat();
			}
			else if ((mp_oControl->GetTierFormatScope() == OS_VIEW)) 
			{
				oTierFormat = mp_oTierArea->GetTierFormat();
			}
			if ((mp_yInterval == IL_YEAR)) 
			{
				clrForeColor = oTierAppearance->GetYearColors()->Item(mp_oControl->GetMathLib()->GetTierIndex(dtStart.GetYear()))->GetForeColor();
				clrBackColor = oTierAppearance->GetYearColors()->Item(mp_oControl->GetMathLib()->GetTierIndex(dtStart.GetYear()))->GetBackColor();
				clrStartGradientColor = oTierAppearance->GetYearColors()->Item(mp_oControl->GetMathLib()->GetTierIndex(dtStart.GetYear()))->GetStartGradientColor();
				clrEndGradientColor = oTierAppearance->GetYearColors()->Item(mp_oControl->GetMathLib()->GetTierIndex(dtStart.GetYear()))->GetEndGradientColor();
				clrHatchBackColor = oTierAppearance->GetYearColors()->Item(mp_oControl->GetMathLib()->GetTierIndex(dtStart.GetYear()))->GetHatchBackColor();
				clrHatchForeColor = oTierAppearance->GetYearColors()->Item(mp_oControl->GetMathLib()->GetTierIndex(dtStart.GetYear()))->GetHatchForeColor();
				sText = mp_oControl->GetConfiguration()->Format(dtStart, oTierFormat->GetYearIntervalFormat());
			}
			else if ((mp_yInterval == IL_QUARTER)) 
			{
				clrForeColor = oTierAppearance->GetQuarterColors()->Item(CStr(mp_oControl->GetMathLib()->Quarter(dtStart)))->GetForeColor();
				clrBackColor = oTierAppearance->GetQuarterColors()->Item(CStr(mp_oControl->GetMathLib()->Quarter(dtStart)))->GetBackColor();
				clrStartGradientColor = oTierAppearance->GetQuarterColors()->Item(CStr(mp_oControl->GetMathLib()->Quarter(dtStart)))->GetStartGradientColor();
				clrEndGradientColor = oTierAppearance->GetQuarterColors()->Item(CStr(mp_oControl->GetMathLib()->Quarter(dtStart)))->GetEndGradientColor();
				clrHatchBackColor = oTierAppearance->GetQuarterColors()->Item(CStr(mp_oControl->GetMathLib()->Quarter(dtStart)))->GetHatchBackColor();
				clrHatchForeColor = oTierAppearance->GetQuarterColors()->Item(CStr(mp_oControl->GetMathLib()->Quarter(dtStart)))->GetHatchForeColor();
				sText = mp_oControl->GetConfiguration()->Format(dtStart, oTierFormat->GetQuarterIntervalFormat());
			}
			else if ((mp_yInterval == IL_MONTH)) 
			{
				clrForeColor = oTierAppearance->GetMonthColors()->Item(CStr(dtStart.GetMonth()))->GetForeColor();
				clrBackColor = oTierAppearance->GetMonthColors()->Item(CStr(dtStart.GetMonth()))->GetBackColor();
				clrStartGradientColor = oTierAppearance->GetMonthColors()->Item(CStr(dtStart.GetMonth()))->GetStartGradientColor();
				clrEndGradientColor = oTierAppearance->GetMonthColors()->Item(CStr(dtStart.GetMonth()))->GetEndGradientColor();
				clrHatchBackColor = oTierAppearance->GetMonthColors()->Item(CStr(dtStart.GetMonth()))->GetHatchBackColor();
				clrHatchForeColor = oTierAppearance->GetMonthColors()->Item(CStr(dtStart.GetMonth()))->GetHatchForeColor();
				sText = mp_oControl->GetConfiguration()->Format(dtStart, oTierFormat->GetMonthIntervalFormat());
			}
			else if ((mp_yInterval == IL_WEEK)) 
			{
				clrForeColor = oTierAppearance->GetWeekColors()->Item(CStr(mp_oControl->GetMathLib()->Week(dtStart)))->GetForeColor();
				clrBackColor = oTierAppearance->GetWeekColors()->Item(CStr(mp_oControl->GetMathLib()->Week(dtStart)))->GetBackColor();
				clrStartGradientColor = oTierAppearance->GetWeekColors()->Item(CStr(mp_oControl->GetMathLib()->Week(dtStart)))->GetStartGradientColor();
				clrEndGradientColor = oTierAppearance->GetWeekColors()->Item(CStr(mp_oControl->GetMathLib()->Week(dtStart)))->GetEndGradientColor();
				clrHatchBackColor = oTierAppearance->GetWeekColors()->Item(CStr(mp_oControl->GetMathLib()->Week(dtStart)))->GetHatchBackColor();
				clrHatchForeColor = oTierAppearance->GetWeekColors()->Item(CStr(mp_oControl->GetMathLib()->Week(dtStart)))->GetHatchForeColor();
				sText = mp_oControl->GetConfiguration()->Format(dtStart, oTierFormat->GetWeekIntervalFormat());
			}
			else if ((mp_yInterval == IL_DAY)) 
			{
				if ((mp_yTierType == ST_DAYOFWEEK)) 
				{
					clrBackColor = oTierAppearance->GetDayOfWeekColors()->Item(CStr(mp_oControl->GetMathLib()->NumericDayOfWeek(dtStart)))->GetBackColor();
					clrForeColor = oTierAppearance->GetDayOfWeekColors()->Item(CStr(mp_oControl->GetMathLib()->NumericDayOfWeek(dtStart)))->GetForeColor();
					clrStartGradientColor = oTierAppearance->GetDayOfWeekColors()->Item(CStr(mp_oControl->GetMathLib()->NumericDayOfWeek(dtStart)))->GetStartGradientColor();
					clrEndGradientColor = oTierAppearance->GetDayOfWeekColors()->Item(CStr(mp_oControl->GetMathLib()->NumericDayOfWeek(dtStart)))->GetEndGradientColor();
					clrHatchBackColor = oTierAppearance->GetDayOfWeekColors()->Item(CStr(mp_oControl->GetMathLib()->NumericDayOfWeek(dtStart)))->GetHatchBackColor();
					clrHatchForeColor = oTierAppearance->GetDayOfWeekColors()->Item(CStr(mp_oControl->GetMathLib()->NumericDayOfWeek(dtStart)))->GetHatchForeColor();
					sText = mp_oControl->GetConfiguration()->Format(dtStart, oTierFormat->GetDayOfWeekIntervalFormat());
				}
				else if ((mp_yTierType == ST_DAY)) 
				{
					clrBackColor = oTierAppearance->GetDayColors()->Item(mp_oControl->GetMathLib()->GetTierIndex(dtStart.GetDay()))->GetBackColor();
					clrForeColor = oTierAppearance->GetDayColors()->Item(mp_oControl->GetMathLib()->GetTierIndex(dtStart.GetDay()))->GetForeColor();
					clrStartGradientColor = oTierAppearance->GetDayColors()->Item(mp_oControl->GetMathLib()->GetTierIndex(dtStart.GetDay()))->GetStartGradientColor();
					clrEndGradientColor = oTierAppearance->GetDayColors()->Item(mp_oControl->GetMathLib()->GetTierIndex(dtStart.GetDay()))->GetEndGradientColor();
					clrHatchBackColor = oTierAppearance->GetDayColors()->Item(mp_oControl->GetMathLib()->GetTierIndex(dtStart.GetDay()))->GetHatchBackColor();
					clrHatchForeColor = oTierAppearance->GetDayColors()->Item(mp_oControl->GetMathLib()->GetTierIndex(dtStart.GetDay()))->GetHatchForeColor();
					sText = mp_oControl->GetConfiguration()->Format(dtStart, oTierFormat->GetDayIntervalFormat());
				}
				else if ((mp_yTierType == ST_DAYOFYEAR)) 
				{
					clrBackColor = oTierAppearance->GetDayOfYearColors()->Item(mp_oControl->GetMathLib()->GetTierIndex(dtStart.GetDayOfYear()))->GetBackColor();
					clrForeColor = oTierAppearance->GetDayOfYearColors()->Item(mp_oControl->GetMathLib()->GetTierIndex(dtStart.GetDayOfYear()))->GetForeColor();
					clrStartGradientColor = oTierAppearance->GetDayOfYearColors()->Item(mp_oControl->GetMathLib()->GetTierIndex(dtStart.GetDayOfYear()))->GetStartGradientColor();
					clrEndGradientColor = oTierAppearance->GetDayOfYearColors()->Item(mp_oControl->GetMathLib()->GetTierIndex(dtStart.GetDayOfYear()))->GetEndGradientColor();
					clrHatchBackColor = oTierAppearance->GetDayOfYearColors()->Item(mp_oControl->GetMathLib()->GetTierIndex(dtStart.GetDayOfYear()))->GetHatchBackColor();
					clrHatchForeColor = oTierAppearance->GetDayOfYearColors()->Item(mp_oControl->GetMathLib()->GetTierIndex(dtStart.GetDayOfYear()))->GetHatchForeColor();
					sText = mp_oControl->GetConfiguration()->Format(dtStart, oTierFormat->GetDayOfYearIntervalFormat());
				}
			}
			else if ((mp_yInterval == IL_HOUR)) 
			{
				clrBackColor = oTierAppearance->GetHourColors()->Item(mp_oControl->GetMathLib()->GetTierIndex(dtStart.GetHour()))->GetBackColor();
				clrForeColor = oTierAppearance->GetHourColors()->Item(mp_oControl->GetMathLib()->GetTierIndex(dtStart.GetHour()))->GetForeColor();
				clrStartGradientColor = oTierAppearance->GetHourColors()->Item(mp_oControl->GetMathLib()->GetTierIndex(dtStart.GetHour()))->GetStartGradientColor();
				clrEndGradientColor = oTierAppearance->GetHourColors()->Item(mp_oControl->GetMathLib()->GetTierIndex(dtStart.GetHour()))->GetEndGradientColor();
				clrHatchBackColor = oTierAppearance->GetHourColors()->Item(mp_oControl->GetMathLib()->GetTierIndex(dtStart.GetHour()))->GetHatchBackColor();
				clrHatchForeColor = oTierAppearance->GetHourColors()->Item(mp_oControl->GetMathLib()->GetTierIndex(dtStart.GetHour()))->GetHatchForeColor();
				sText = mp_oControl->GetConfiguration()->Format(dtStart, oTierFormat->GetHourIntervalFormat());
			}
			else if ((mp_yInterval == IL_MINUTE)) 
			{
				clrBackColor = oTierAppearance->GetMinuteColors()->Item(mp_oControl->GetMathLib()->GetTierIndex(dtStart.GetMinute()))->GetBackColor();
				clrForeColor = oTierAppearance->GetMinuteColors()->Item(mp_oControl->GetMathLib()->GetTierIndex(dtStart.GetMinute()))->GetForeColor();
				clrStartGradientColor = oTierAppearance->GetMinuteColors()->Item(mp_oControl->GetMathLib()->GetTierIndex(dtStart.GetMinute()))->GetStartGradientColor();
				clrEndGradientColor = oTierAppearance->GetMinuteColors()->Item(mp_oControl->GetMathLib()->GetTierIndex(dtStart.GetMinute()))->GetEndGradientColor();
				clrHatchBackColor = oTierAppearance->GetMinuteColors()->Item(mp_oControl->GetMathLib()->GetTierIndex(dtStart.GetMinute()))->GetHatchBackColor();
				clrHatchForeColor = oTierAppearance->GetMinuteColors()->Item(mp_oControl->GetMathLib()->GetTierIndex(dtStart.GetMinute()))->GetHatchForeColor();
				sText = mp_oControl->GetConfiguration()->Format(dtStart, oTierFormat->GetMinuteIntervalFormat());
			}
			else if ((mp_yInterval == IL_SECOND)) 
			{
				clrBackColor = oTierAppearance->GetSecondColors()->Item(mp_oControl->GetMathLib()->GetTierIndex(dtStart.GetSecond()))->GetBackColor();
				clrForeColor = oTierAppearance->GetSecondColors()->Item(mp_oControl->GetMathLib()->GetTierIndex(dtStart.GetSecond()))->GetForeColor();
				clrStartGradientColor = oTierAppearance->GetSecondColors()->Item(mp_oControl->GetMathLib()->GetTierIndex(dtStart.GetSecond()))->GetStartGradientColor();
				clrEndGradientColor = oTierAppearance->GetSecondColors()->Item(mp_oControl->GetMathLib()->GetTierIndex(dtStart.GetSecond()))->GetEndGradientColor();
				clrHatchBackColor = oTierAppearance->GetSecondColors()->Item(mp_oControl->GetMathLib()->GetTierIndex(dtStart.GetSecond()))->GetHatchBackColor();
				clrHatchForeColor = oTierAppearance->GetSecondColors()->Item(mp_oControl->GetMathLib()->GetTierIndex(dtStart.GetSecond()))->GetHatchForeColor();
				sText = mp_oControl->GetConfiguration()->Format(dtStart, oTierFormat->GetSecondIntervalFormat());
			}
			if (lEnd > mp_oTierArea->GetTimeLine()->Getf_lEnd()) 
			{
				lEnd = mp_oTierArea->GetTimeLine()->Getf_lEnd();
			}
			lTextWidth = mp_oControl->clsG->mp_lStrWidth(sText, &mp_oStyle->mp_oFont);
			if ((lEnd - lStart) > lTextWidth) 
			{
				mp_oControl->GetCustomTierDrawEventArgs()->Clear();
				mp_oControl->GetCustomTierDrawEventArgs()->SetTierPosition(mp_yTierPosition);
				mp_oControl->GetCustomTierDrawEventArgs()->SetText(sText);
				mp_oControl->GetCustomTierDrawEventArgs()->SetStartDate(dtStart);
				mp_oControl->FireTierTextDraw();
				if (mp_yBackgroundMode == ETBGM_TIERAPPEARANCE) 
				{
					mp_oControl->clsG->mp_DrawItemEx(lStart, lTop, lEnd, lBottom, sText, FALSE, NULL, lStartTrim, lEndTrim, mp_oStyle, clrBackColor, clrForeColor, clrStartGradientColor, clrEndGradientColor, clrHatchBackColor, clrHatchForeColor);
				} 
				else if (mp_yBackgroundMode == ETBGM_STYLE) 
				{
					mp_oControl->clsG->mp_DrawItem(lStart, lTop, lEnd, lBottom, "", sText, FALSE, NULL, lStartTrim, lEndTrim, mp_oStyle);
				}
			}
			else 
			{
				if (mp_yBackgroundMode == ETBGM_TIERAPPEARANCE) 
				{
					mp_oControl->clsG->mp_DrawItemEx(lStart, lTop, lEnd, lBottom, "", FALSE, NULL, lStartTrim, lEndTrim, mp_oStyle, clrBackColor, clrForeColor, clrStartGradientColor, clrEndGradientColor, clrHatchBackColor, clrHatchForeColor);
				}
				else if (mp_yBackgroundMode == ETBGM_STYLE) 
				{
					mp_oControl->clsG->mp_DrawItem(lStart, lTop, lEnd, lBottom, "", "", FALSE, NULL, lStartTrim, lEndTrim, mp_oStyle);
				}
			}
		}
	}

	CString clsTier::GetStyleIndex(void)
	{
		if (mp_sStyleIndex == "DS_TIER") 
		{
			return "";
		}
		else 
		{
			return mp_sStyleIndex;
		}
	}

	void clsTier::SetStyleIndex(CString Value)
	{
		Value = Value.Trim();
		if (Value.GetLength() == 0) {Value = "DS_TIER";} 
		mp_sStyleIndex = Value;
		mp_oStyle = mp_oControl->GetStyles()->FItem(Value);
	}

	clsStyle* clsTier::GetStyle(void)
	{
		return mp_oStyle;
	}

	LONG clsTier::GetBackgroundMode(void)
	{
		return mp_yBackgroundMode;
	}

	void clsTier::SetBackgroundMode(LONG Value)
	{
		mp_yBackgroundMode = Value;
	}

	CString clsTier::GetXML(void)
	{
		clsXML oXML(mp_oControl, mp_sTierPosition);
		oXML.InitializeWriter();
		oXML.WriteProperty("yTierPosition", mp_yTierPosition);
		oXML.WriteProperty("sTierPosition", mp_sTierPosition);
		oXML.WriteProperty("Height", mp_lHeight);
		oXML.WriteProperty("Interval", mp_yInterval);
		oXML.WriteProperty("Factor", mp_lFactor);
		oXML.WriteProperty("Tag", mp_sTag);
		oXML.WriteProperty("TierType", mp_yTierType);
		oXML.WriteProperty("Visible", mp_bVisible);
		oXML.WriteProperty("StyleIndex", mp_sStyleIndex);
		oXML.WriteProperty("BackgroundMode", mp_yBackgroundMode);
		return oXML.GetXML();
	}

	void clsTier::SetXML(CString sXML)
	{
		clsXML oXML(mp_oControl, mp_sTierPosition);
		oXML.SetXML(sXML);
		oXML.InitializeReader();
		oXML.ReadProperty("yTierPosition", mp_yTierPosition);
		oXML.ReadProperty("sTierPosition", mp_sTierPosition);
		oXML.ReadProperty("Height", mp_lHeight);
		oXML.ReadProperty("Interval", mp_yInterval);
		oXML.ReadProperty("Factor", mp_lFactor);
		oXML.ReadProperty("Tag", mp_sTag);
		oXML.ReadProperty("TierType", mp_yTierType);
		oXML.ReadProperty("Visible", mp_bVisible);
		oXML.ReadProperty("StyleIndex", mp_sStyleIndex);
		SetStyleIndex(mp_sStyleIndex);
		oXML.ReadProperty("BackgroundMode", mp_yBackgroundMode);
	}

void clsTier::Clear()
{
	mp_bVisible = TRUE;
	mp_lFactor = 1;
	mp_lHeight = 14;
	mp_sTag = "";
	mp_yInterval = IL_WEEK;
	mp_yTierType = ST_DAYOFWEEK;
	mp_sStyleIndex = "DS_TIER";
	mp_oStyle = mp_oControl->GetStyles()->FItem("DS_TIER");
	mp_yBackgroundMode = ETBGM_TIERAPPEARANCE;
}

void clsTier::Clone(clsTier* oClone)
{
    oClone->SetHeight(mp_lHeight);
    oClone->SetInterval(mp_yInterval);
    oClone->SetFactor(mp_lFactor);
    oClone->SetTag(mp_sTag);
    oClone->SetTierType(mp_yTierType);
    oClone->SetVisible(mp_bVisible);
    oClone->SetStyleIndex(mp_sStyleIndex);
    oClone->SetBackgroundMode(mp_yBackgroundMode);
}