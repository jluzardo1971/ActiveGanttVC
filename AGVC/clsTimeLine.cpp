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
#include "clsTimeLine.h"



IMPLEMENT_DYNCREATE(clsTimeLine, CCmdTarget)

//{8C665285-172E-4311-A086-D861C33C1110}
static const IID IID_IclsTimeLine = { 0x8C665285, 0x172E, 0x4311, { 0xA0, 0x86, 0xD8, 0x61, 0xC3, 0x3C, 0x11, 0x10} };

//{3A9BBC36-07A0-447C-A11A-6C374B25EE80}
IMPLEMENT_OLECREATE_FLAGS(clsTimeLine, "AGVC.clsTimeLine", afxRegApartmentThreading, 0x3A9BBC36, 0x07A0, 0x447C, 0xA1, 0x1A, 0x6C, 0x37, 0x4B, 0x25, 0xEE, 0x80)

BEGIN_DISPATCH_MAP(clsTimeLine, CCmdTarget)
	DISP_PROPERTY_EX_ID(clsTimeLine, "StyleIndex", 1, odl_GetStyleIndex, odl_SetStyleIndex, VT_BSTR) 
	DISP_PROPERTY_EX_ID(clsTimeLine, "Style", 2, odl_GetStyle, SetNotSupported, VT_DISPATCH) 
	DISP_PROPERTY_EX_ID(clsTimeLine, "ForeColor", 3, odl_GetForeColor, odl_SetForeColor, VT_COLOR) 
	DISP_PROPERTY_EX_ID(clsTimeLine, "EndDate", 4, odl_GetEndDate, SetNotSupported, VT_DATE) 
	DISP_PROPERTY_EX_ID(clsTimeLine, "StartDate", 5, odl_GetStartDate, SetNotSupported, VT_DATE) 
	DISP_PROPERTY_EX_ID(clsTimeLine, "Height", 6, odl_GetHeight, SetNotSupported, VT_I4) 
	DISP_PROPERTY_EX_ID(clsTimeLine, "Top", 7, odl_GetTop, SetNotSupported, VT_I4) 
	DISP_PROPERTY_EX_ID(clsTimeLine, "Bottom", 8, odl_GetBottom, SetNotSupported, VT_I4) 
	DISP_PROPERTY_EX_ID(clsTimeLine, "TimeLineScrollBar", 9, odl_GetTimeLineScrollBar, SetNotSupported, VT_DISPATCH) 
	DISP_PROPERTY_EX_ID(clsTimeLine, "TierArea", 10, odl_GetTierArea, SetNotSupported, VT_DISPATCH) 
	DISP_PROPERTY_EX_ID(clsTimeLine, "TickMarkArea", 11, odl_GetTickMarkArea, SetNotSupported, VT_DISPATCH) 
	DISP_PROPERTY_EX_ID(clsTimeLine, "ProgressLine", 12, odl_GetProgressLine, SetNotSupported, VT_DISPATCH) 
	DISP_FUNCTION_ID(clsTimeLine, "Move", 13, odl_Move, VT_EMPTY, VTS_I4 VTS_I4) 
	DISP_FUNCTION_ID(clsTimeLine, "Position", 14, odl_Position, VT_EMPTY, VTS_DATE) 
	DISP_FUNCTION_ID(clsTimeLine, "GetXML", 15, odl_GetXML, VT_BSTR, VTS_NONE) 
	DISP_FUNCTION_ID(clsTimeLine, "SetXML", 16, odl_SetXML, VT_EMPTY, VTS_BSTR) 
END_DISPATCH_MAP()

BEGIN_INTERFACE_MAP(clsTimeLine, CCmdTarget)
	INTERFACE_PART(clsTimeLine, IID_IclsTimeLine, Dispatch)
END_INTERFACE_MAP()

BEGIN_MESSAGE_MAP(clsTimeLine, CCmdTarget)
END_MESSAGE_MAP()

clsTimeLine::clsTimeLine()
{
	AfxMessageBox(_T("Invalid constructor has been invoked for clsTimeLine. Please inform support@sourcecodestore.com of this error"), MB_OK | MB_ICONSTOP | MB_APPLMODAL);
}

clsTimeLine::clsTimeLine(CActiveGanttVCCtl* oControl, clsView* oView)
{
	EnableAutomation();
	AfxOleLockApp();
	mp_oControl = oControl;
	mp_oView = oView;
	mp_oTimeLineScrollBar = new clsTimeLineScrollBar(mp_oControl);
	mp_oTierArea = new clsTierArea(mp_oControl, this);
	mp_oTickMarkArea = new clsTickMarkArea(mp_oControl, this);
	mp_oProgressLine = new clsProgressLine(mp_oControl);
	Clear();
}

clsTimeLine::~clsTimeLine()
{
	delete mp_oTimeLineScrollBar;
	mp_oTimeLineScrollBar = NULL;
	delete mp_oTierArea;
	mp_oTierArea = NULL;
	delete mp_oTickMarkArea;
	mp_oTickMarkArea = NULL;
	delete mp_oProgressLine;
	mp_oProgressLine = NULL;
	AfxOleUnlockApp();
}

void clsTimeLine::OnFinalRelease()
{
	CCmdTarget::OnFinalRelease();
}

CString clsTimeLine::GetStyleIndex(void)
{
	if (mp_sStyleIndex == "DS_TIMELINE") 
	{
		return "";
	}
	else 
	{
		return mp_sStyleIndex;
	}
}

void clsTimeLine::SetStyleIndex(CString Value)
{
	Value = Value.Trim();
	if (Value.GetLength() == 0) {Value = "DS_TIMELINE";} 
	mp_sStyleIndex = Value;
	mp_oStyle = mp_oControl->GetStyles()->FItem(Value);
}

clsStyle* clsTimeLine::GetStyle(void)
{
	return mp_oStyle;
}

Gdiplus::Color clsTimeLine::GetForeColor(void)
{
	return mp_clrForeColor;
}

void clsTimeLine::SetForeColor(Gdiplus::Color Value)
{
	mp_clrForeColor = Value;
}

COleDateTime clsTimeLine::GetEndDate(void)
{
	Calculate();
	mp_dtEndDate = mp_oControl->GetMathLib()->DateTimeAdd(mp_oView->GetInterval(), (mp_lEnd - mp_lStart) * mp_oView->GetFactor(), mp_dtStartDate);
	return mp_dtEndDate;
}

COleDateTime clsTimeLine::GetStartDate(void)
{
	return mp_dtStartDate;
}

LONG clsTimeLine::GetHeight(void)
{
	return GetBottom() - GetTop();
}

LONG clsTimeLine::GetTop(void)
{
	return mp_oControl->Getmt_BorderThickness();
}

LONG clsTimeLine::GetBottom(void)
{
	LONG lReturn = 0;
	LONG lUpperTierHeight = 0;
	LONG lMiddleTierHeight = 0;
	LONG lLowerTierHeight = 0;
	LONG lTickMarkAreaHeight = 0;
	lReturn = 0;
	lUpperTierHeight = 0;
	lLowerTierHeight = 0;
	lTickMarkAreaHeight = 0;
	if ((mp_oTierArea->GetUpperTier()->GetVisible() == TRUE)) 
	{
		lUpperTierHeight = mp_oTierArea->GetUpperTier()->GetHeight();
	}
	if ((mp_oTierArea->GetMiddleTier()->GetVisible() == TRUE)) 
	{
		lMiddleTierHeight = mp_oTierArea->GetMiddleTier()->GetHeight();
	}
	if ((mp_oTierArea->GetLowerTier()->GetVisible() == TRUE)) 
	{
		lLowerTierHeight = mp_oTierArea->GetLowerTier()->GetHeight();
	}
	if ((mp_oTickMarkArea->GetVisible() == TRUE)) 
	{
		lTickMarkAreaHeight = mp_oTickMarkArea->GetHeight();
	}
	lReturn = lUpperTierHeight + lMiddleTierHeight + lLowerTierHeight + lTickMarkAreaHeight;
	lReturn = lReturn + mp_oControl->Getmt_BorderThickness();
	return lReturn;
}

clsTimeLineScrollBar* clsTimeLine::GetTimeLineScrollBar(void)
{
	return mp_oTimeLineScrollBar;
}

clsTierArea* clsTimeLine::GetTierArea(void)
{
	return mp_oTierArea;
}

clsTickMarkArea* clsTimeLine::GetTickMarkArea(void)
{
	return mp_oTickMarkArea;
}

clsProgressLine* clsTimeLine::GetProgressLine(void)
{
	return mp_oProgressLine;
}

void clsTimeLine::Setf_StartDate(COleDateTime Value)
{
	mp_dtStartDate = Value;
	Calculate();
}

LONG clsTimeLine::Getf_lStart(void)
{
	return mp_lStart;
}

LONG clsTimeLine::Getf_lEnd(void)
{
	return mp_lEnd;;
}

void clsTimeLine::Move(LONG Interval, LONG Factor)
{
	Setf_StartDate(mp_oControl->GetMathLib()->DateTimeAdd(Interval, Factor, mp_dtStartDate));
}

void clsTimeLine::Position(COleDateTime TimeLineStartDate)
{
	Setf_StartDate(TimeLineStartDate);
}

void clsTimeLine::Calculate(void)
{
	if (mp_oControl->f_TimeLineScrollBar()->GetEnabled() == TRUE) 
	{
		mp_dtStartDate = mp_oControl->GetMathLib()->DateTimeAdd(mp_oControl->f_TimeLineScrollBar()->GetInterval(), mp_oControl->f_TimeLineScrollBar()->GetValue() * mp_oControl->f_TimeLineScrollBar()->GetFactor(), mp_oControl->f_TimeLineScrollBar()->GetStartDate());
	}
	mp_dtStartDate = mp_oControl->GetMathLib()->RoundDate(mp_oView->GetInterval(), mp_oView->GetFactor(), mp_dtStartDate);
	mp_lStart = mp_oControl->GetSplitter()->GetRight();
	mp_lEnd = mp_oControl->Getmt_RightMargin();
	if (mp_oStyle->GetAppearance() == SA_RAISED) 
	{
		if (mp_oStyle->GetButtonStyle() == BT_NORMALWINDOWS) 
		{
			mp_lStart = mp_oControl->GetSplitter()->GetRight() + 1;
			mp_lEnd = mp_oControl->Getmt_RightMargin() - 1;
		}
		else if (mp_oStyle->GetButtonStyle() == BT_LIGHTWEIGHT) 
		{
			mp_lStart = mp_oControl->GetSplitter()->GetRight() + 2;
			mp_lEnd = mp_oControl->Getmt_RightMargin() - 2;
		}
	}
	else 
	{
		mp_lStart = mp_oControl->GetSplitter()->GetRight();
		mp_lEnd = mp_oControl->Getmt_RightMargin();
	}
}

LONG clsTimeLine::TiersTickMarksPosition(CString v_yTier)
{
	LONG lReturn = 0;
	LONG lUpperTierHeight = 0;
	LONG lMiddleTierHeight = 0;
	LONG lLowerTierHeight = 0;
	LONG lTickMarkAreaHeight = 0;
	lReturn = 0;
	lUpperTierHeight = 0;
	lLowerTierHeight = 0;
	lTickMarkAreaHeight = 0;
	if ((mp_oTierArea->GetUpperTier()->GetVisible() == TRUE)) 
	{
		lUpperTierHeight = mp_oTierArea->GetUpperTier()->GetHeight();
	}
	if ((mp_oTierArea->GetMiddleTier()->GetVisible() == TRUE)) 
	{
		lMiddleTierHeight = mp_oTierArea->GetMiddleTier()->GetHeight();
	}
	if ((mp_oTierArea->GetLowerTier()->GetVisible() == TRUE)) 
	{
		lLowerTierHeight = mp_oTierArea->GetLowerTier()->GetHeight();
	}
	if ((mp_oTickMarkArea->GetVisible() == TRUE)) 
	{
		lTickMarkAreaHeight = mp_oTickMarkArea->GetHeight();
	}
	lReturn = lUpperTierHeight + lMiddleTierHeight + lLowerTierHeight + lTickMarkAreaHeight;
	if ((lReturn > 0)) 
	{
		lReturn = lReturn;
	}
	lReturn = lReturn + mp_oControl->Getmt_BorderThickness();
	if (v_yTier == "UpperTier")
	{
		lReturn = lReturn - lUpperTierHeight - lMiddleTierHeight - lLowerTierHeight - lTickMarkAreaHeight;
	}
	else if (v_yTier == "MiddleTier")
	{
		lReturn = lReturn - lMiddleTierHeight - lLowerTierHeight - lTickMarkAreaHeight;
	}
	else if (v_yTier == "LowerTier")
	{
		lReturn = lReturn - lLowerTierHeight - lTickMarkAreaHeight;
	}
	else if (v_yTier == "TickMarkArea")
	{
		lReturn = lReturn - lTickMarkAreaHeight;
	}
	return lReturn;
}

void clsTimeLine::Draw(void)
{
	LONG lBottom = 0;
	LONG lTop = 0;
	LONG lLeft = 0;
	LONG lRight = 0;
	if ((GetHeight() == 0)) 
	{
		return;
	}
	lBottom = GetBottom();
	lTop = GetTop();
	lLeft = mp_oControl->GetSplitter()->GetRight();
	lRight = mp_oControl->Getmt_RightMargin();
	mp_oControl->clsG->mp_ClipRegion(lLeft, lTop, lRight, lBottom, TRUE);
	mp_oControl->clsG->mp_DrawItem(lLeft, lTop, lRight, lBottom, "", "", FALSE, NULL, 0, 0, 
		mp_oStyle);
	if (mp_oStyle->GetAppearance() == SA_RAISED) 
	{
		if (mp_oStyle->GetButtonStyle() == BT_NORMALWINDOWS) 
		{
			mp_oControl->clsG->mp_ClipRegion(lLeft + 2, lTop + 2, lRight - 2, lBottom - 2, TRUE);
		}
		else if (mp_oStyle->GetButtonStyle() == BT_LIGHTWEIGHT) 
		{
			mp_oControl->clsG->mp_ClipRegion(lLeft + 1, lTop + 1, lRight - 1, lBottom - 1, TRUE);
		}
	}
	mp_oTierArea->GetUpperTier()->Position();
	mp_oTierArea->GetMiddleTier()->Position();
	mp_oTierArea->GetLowerTier()->Position();
	mp_oTickMarkArea->Draw();
}

CString clsTimeLine::GetXML(void)
{
	clsXML oXML(mp_oControl, "TimeLine");
	oXML.InitializeWriter();
	oXML.WriteProperty("ForeColor", mp_clrForeColor);
	oXML.WriteProperty("StyleIndex", mp_sStyleIndex);
	oXML.WriteProperty("StartDate", mp_dtStartDate);
	oXML.WriteObject(mp_oProgressLine->GetXML());
	oXML.WriteObject(mp_oTimeLineScrollBar->GetXML());
	oXML.WriteObject(mp_oTickMarkArea->GetXML());
	oXML.WriteObject(mp_oTierArea->GetXML());
	return oXML.GetXML();
}

void clsTimeLine::SetXML(CString sXML)
{
	clsXML oXML(mp_oControl, "TimeLine");
	oXML.SetXML(sXML);
	oXML.InitializeReader();
	oXML.ReadProperty("ForeColor", mp_clrForeColor);
	oXML.ReadProperty("StyleIndex", mp_sStyleIndex);
	oXML.ReadProperty("StartDate", mp_dtStartDate);
	SetStyleIndex(mp_sStyleIndex);
	Calculate();
	mp_oProgressLine->SetXML(oXML.ReadObject("ProgressLine"));
	mp_oTimeLineScrollBar->SetXML(oXML.ReadObject("TimeLineScrollBar"));
	mp_oTickMarkArea->SetXML(oXML.ReadObject("TickMarkArea"));
	mp_oTierArea->SetXML(oXML.ReadObject("TierArea"));
}

void clsTimeLine::Clear()
{
	mp_clrForeColor = Color::Black;
	mp_sStyleIndex = "DS_TIMELINE";
	mp_oStyle = mp_oControl->GetStyles()->FItem("DS_TIMELINE");
	COleDateTime dtTimeNow;
	SYSTEMTIME lt;
	GetLocalTime(&lt);
	dtTimeNow.SetDateTime(lt.wYear, lt.wMonth, lt.wDay, lt.wHour, lt.wMinute, lt.wSecond);
	Setf_StartDate(mp_oControl->GetMathLib()->DateTimeAdd(IL_HOUR, -3, dtTimeNow));
	GetProgressLine()->Clear();
    GetTimeLineScrollBar()->Clear();
    GetTickMarkArea()->Clear();
    GetTierArea()->Clear();
}

void clsTimeLine::Clone(clsTimeLine* oClone)
{
	oClone->SetForeColor(mp_clrForeColor);
	oClone->SetStyleIndex(mp_sStyleIndex);
	oClone->Setf_StartDate(mp_dtStartDate);
	GetProgressLine()->Clone(oClone->GetProgressLine());
	GetTimeLineScrollBar()->Clone(oClone->GetTimeLineScrollBar());
	GetTickMarkArea()->Clone(oClone->GetTickMarkArea());
	GetTierArea()->Clone(oClone->GetTierArea());
}