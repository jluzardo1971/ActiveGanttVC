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
#include "clsTickMarkArea.h"



IMPLEMENT_DYNCREATE(clsTickMarkArea, CCmdTarget)

//{B1FAA52F-E118-4B96-9E01-8CF988191E95}
static const IID IID_IclsTickMarkArea = { 0xB1FAA52F, 0xE118, 0x4B96, { 0x9E, 0x01, 0x8C, 0xF9, 0x88, 0x19, 0x1E, 0x95} };

//{9D312576-FE4E-4C95-9976-C62260882B80}
IMPLEMENT_OLECREATE_FLAGS(clsTickMarkArea, "AGVC.clsTickMarkArea", afxRegApartmentThreading, 0x9D312576, 0xFE4E, 0x4C95, 0x99, 0x76, 0xC6, 0x22, 0x60, 0x88, 0x2B, 0x80)

BEGIN_DISPATCH_MAP(clsTickMarkArea, CCmdTarget)
	DISP_PROPERTY_EX_ID(clsTickMarkArea, "Height", 1, odl_GetHeight, odl_SetHeight, VT_I4) 
	DISP_PROPERTY_EX_ID(clsTickMarkArea, "BigTickMarkHeight", 2, odl_GetBigTickMarkHeight, odl_SetBigTickMarkHeight, VT_I4) 
	DISP_PROPERTY_EX_ID(clsTickMarkArea, "MediumTickMarkHeight", 3, odl_GetMediumTickMarkHeight, odl_SetMediumTickMarkHeight, VT_I4) 
	DISP_PROPERTY_EX_ID(clsTickMarkArea, "SmallTickMarkHeight", 4, odl_GetSmallTickMarkHeight, odl_SetSmallTickMarkHeight, VT_I4) 
	DISP_PROPERTY_EX_ID(clsTickMarkArea, "Visible", 5, odl_GetVisible, odl_SetVisible, VT_BOOL) 
	DISP_PROPERTY_EX_ID(clsTickMarkArea, "StyleIndex", 6, odl_GetStyleIndex, odl_SetStyleIndex, VT_BSTR) 
	DISP_PROPERTY_EX_ID(clsTickMarkArea, "TextOffset", 7, odl_GetTextOffset, odl_SetTextOffset, VT_I4) 
	DISP_PROPERTY_EX_ID(clsTickMarkArea, "TickMarks", 8, odl_GetTickMarks, SetNotSupported, VT_DISPATCH) 
	DISP_FUNCTION_ID(clsTickMarkArea, "GetXML", 9, odl_GetXML, VT_BSTR, VTS_NONE) 
	DISP_FUNCTION_ID(clsTickMarkArea, "SetXML", 10, odl_SetXML, VT_EMPTY, VTS_BSTR)
	DISP_PROPERTY_EX_ID(clsTickMarkArea, "Style", 11, odl_GetStyle, SetNotSupported, VT_DISPATCH) 
END_DISPATCH_MAP()

BEGIN_INTERFACE_MAP(clsTickMarkArea, CCmdTarget)
	INTERFACE_PART(clsTickMarkArea, IID_IclsTickMarkArea, Dispatch)
END_INTERFACE_MAP()

BEGIN_MESSAGE_MAP(clsTickMarkArea, CCmdTarget)
END_MESSAGE_MAP()

clsTickMarkArea::clsTickMarkArea(CActiveGanttVCCtl* oControl, clsTimeLine* oTimeLine)
{
	EnableAutomation();
	AfxOleLockApp();
	mp_oControl = oControl;
	mp_oTimeLine = oTimeLine;
	mp_oTickMarks = new clsTickMarks(mp_oControl);
	Clear();
}

clsTickMarkArea::clsTickMarkArea()
{
	AfxMessageBox(_T("Invalid constructor has been invoked for clsTickMarkArea. Please inform support@sourcecodestore.com of this error"), MB_OK | MB_ICONSTOP | MB_APPLMODAL);
}

clsTickMarkArea::~clsTickMarkArea()
{
	delete mp_oTickMarks;
	mp_oTickMarks = NULL;
	AfxOleUnlockApp();
}

void clsTickMarkArea::OnFinalRelease()
{
	CCmdTarget::OnFinalRelease();
}

LONG clsTickMarkArea::GetHeight(void)
{
	return mp_lHeight;
}

void clsTickMarkArea::SetHeight(LONG Value)
{
	mp_lHeight = Value;
}

LONG clsTickMarkArea::GetBigTickMarkHeight(void)
{
	return mp_lBigTickMarkHeight;
}

void clsTickMarkArea::SetBigTickMarkHeight(LONG Value)
{
	mp_lBigTickMarkHeight = Value;
}

LONG clsTickMarkArea::GetMediumTickMarkHeight(void)
{
	return mp_lMediumTickMarkHeight;
}

void clsTickMarkArea::SetMediumTickMarkHeight(LONG Value)
{
	mp_lMediumTickMarkHeight = Value;
}

LONG clsTickMarkArea::GetSmallTickMarkHeight(void)
{
	return mp_lSmallTickMarkHeight;
}

void clsTickMarkArea::SetSmallTickMarkHeight(LONG Value)
{
	mp_lSmallTickMarkHeight = Value;
}

BOOL clsTickMarkArea::GetVisible(void)
{
	return mp_bVisible;
}

void clsTickMarkArea::SetVisible(BOOL Value)
{
	mp_bVisible = Value;
}

CString clsTickMarkArea::GetStyleIndex(void)
{
	if (mp_sStyleIndex == "DS_TICKMARKAREA") 
	{
		return "";
	}
	else 
	{
		return mp_sStyleIndex;
	}
}

void clsTickMarkArea::SetStyleIndex(CString Value)
{
	Value = Value.Trim();
	if (Value.GetLength() == 0) {Value = "DS_TICKMARKAREA";} 
	mp_sStyleIndex = Value;
	mp_oStyle = mp_oControl->GetStyles()->FItem(Value);
}

LONG clsTickMarkArea::GetTextOffset(void)
{
	return mp_lTextOffset;
}

void clsTickMarkArea::SetTextOffset(LONG Value)
{
	mp_lTextOffset = Value;
}

clsTickMarks* clsTickMarkArea::GetTickMarks(void)
{
	return mp_oTickMarks;
}

clsStyle* clsTickMarkArea::GetStyle(void)
{
	return mp_oStyle;
}

void clsTickMarkArea::Draw(void)
{
	COleDateTime dtBuff;
	clsTickMark* oTickMark = NULL;
	LONG lIndex = 0;
	if (GetVisible() == FALSE) 
	{
		return;
	}
    mp_oControl->GetCustomTickMarkAreaDrawEventArgs()->Clear();
    mp_oControl->GetCustomTickMarkAreaDrawEventArgs()->SetCustomDraw(FALSE);
    mp_oControl->GetCustomTickMarkAreaDrawEventArgs()->SetLeft(mp_oTimeLine->Getf_lStart());
    mp_oControl->GetCustomTickMarkAreaDrawEventArgs()->SetTop(mp_oTimeLine->GetBottom() - GetHeight());
    mp_oControl->GetCustomTickMarkAreaDrawEventArgs()->SetRight(mp_oTimeLine->Getf_lEnd());
    mp_oControl->GetCustomTickMarkAreaDrawEventArgs()->SetBottom(mp_oTimeLine->GetBottom());
    mp_oControl->FireCustomTickMarkAreaDraw();
	if (mp_oControl->GetCustomTickMarkAreaDrawEventArgs()->GetCustomDraw() == FALSE)
    {
        mp_oControl->clsG->mp_DrawItem(mp_oTimeLine->Getf_lStart(), mp_oTimeLine->GetBottom() - GetHeight(), mp_oTimeLine->Getf_lEnd(), mp_oTimeLine->GetBottom(), "", "", FALSE, NULL, mp_oTimeLine->Getf_lStart(), mp_oTimeLine->Getf_lEnd(), mp_oStyle);
    }
	mp_oControl->clsG->mp_ClipRegion(mp_oTimeLine->Getf_lStart(), mp_oTimeLine->GetBottom() - GetHeight(), mp_oTimeLine->Getf_lEnd(), mp_oTimeLine->GetBottom(), FALSE);
	for (lIndex = 1; lIndex <= mp_oTickMarks->GetCount(); lIndex++) 
	{
		LONG yInterval;
		LONG lFactor = 0;
		oTickMark = mp_oTickMarks->Item(CStr(lIndex));
		yInterval = oTickMark->GetInterval();
		lFactor = oTickMark->GetFactor();
		if (mp_oControl->GetMathLib()->GetXCoordinateFromDate(mp_oControl->GetMathLib()->DateTimeAdd(yInterval, lFactor, mp_oTimeLine->GetStartDate())) - mp_oControl->GetMathLib()->GetXCoordinateFromDate(mp_oTimeLine->GetStartDate()) < 5) 
		{
			break; 
		}
		dtBuff = mp_oControl->GetMathLib()->RoundDate(yInterval, lFactor, mp_oTimeLine->GetStartDate());

		if (mp_oControl->GetMathLib()->GetXCoordinateFromDate(dtBuff) >= mp_oTimeLine->Getf_lStart()) 
		{
			mp_oControl->GetCustomTickMarkAreaDrawEventArgs()->Clear();
			mp_oControl->GetCustomTickMarkAreaDrawEventArgs()->SetLeft(mp_oTimeLine->Getf_lStart());
			mp_oControl->GetCustomTickMarkAreaDrawEventArgs()->SetTop(mp_oTimeLine->GetBottom() - GetHeight());
			mp_oControl->GetCustomTickMarkAreaDrawEventArgs()->SetRight(mp_oTimeLine->Getf_lEnd());
			mp_oControl->GetCustomTickMarkAreaDrawEventArgs()->SetBottom(mp_oTimeLine->GetBottom());
			mp_oControl->GetCustomTickMarkAreaDrawEventArgs()->SetoTickMark(oTickMark);
			mp_oControl->GetCustomTickMarkAreaDrawEventArgs()->SetdtDate(dtBuff);
			mp_oControl->GetCustomTickMarkAreaDrawEventArgs()->SetX(mp_oControl->GetMathLib()->GetXCoordinateFromDate(dtBuff));
			mp_oControl->FireCustomTickMarkDraw();
			if (mp_oControl->GetCustomTickMarkAreaDrawEventArgs()->GetCustomDraw() == FALSE)
			{
				PaintTickMark(dtBuff, oTickMark);
				if (oTickMark->GetDisplayText() == TRUE) 
				{
					PaintText(mp_oControl->GetMathLib()->GetXCoordinateFromDate(dtBuff), oTickMark);
				}
			}
		}
		while (dtBuff < mp_oTimeLine->GetEndDate()) 
		{
			dtBuff = mp_oControl->GetMathLib()->DateTimeAdd(yInterval, lFactor, dtBuff);
			mp_oControl->GetCustomTickMarkAreaDrawEventArgs()->Clear();
			mp_oControl->GetCustomTickMarkAreaDrawEventArgs()->SetLeft(mp_oTimeLine->Getf_lStart());
			mp_oControl->GetCustomTickMarkAreaDrawEventArgs()->SetTop(mp_oTimeLine->GetBottom() - GetHeight());
			mp_oControl->GetCustomTickMarkAreaDrawEventArgs()->SetRight(mp_oTimeLine->Getf_lEnd());
			mp_oControl->GetCustomTickMarkAreaDrawEventArgs()->SetBottom(mp_oTimeLine->GetBottom());
			mp_oControl->GetCustomTickMarkAreaDrawEventArgs()->SetoTickMark(oTickMark);
			mp_oControl->GetCustomTickMarkAreaDrawEventArgs()->SetdtDate(dtBuff);
			mp_oControl->GetCustomTickMarkAreaDrawEventArgs()->SetX(mp_oControl->GetMathLib()->GetXCoordinateFromDate(dtBuff));
			mp_oControl->FireCustomTickMarkDraw();
			if (mp_oControl->GetCustomTickMarkAreaDrawEventArgs()->GetCustomDraw() == FALSE)
			{
				PaintTickMark(dtBuff, oTickMark);
				if (oTickMark->GetDisplayText() == TRUE) 
				{
					PaintText(mp_oControl->GetMathLib()->GetXCoordinateFromDate(dtBuff), oTickMark);
				}
			}
		}
	}
	mp_oControl->clsG->mp_ClearClipRegion();
}

void clsTickMarkArea::PaintTickMark(COleDateTime dtDate, clsTickMark* oTickMark)
{
	int fXCoordinate = 0;
	LONG lTickMarkHeight = 0;
	COleDateTime dtBuff = mp_oControl->GetMathLib()->RoundDate(oTickMark->GetInterval(), oTickMark->GetFactor(), dtDate);
	fXCoordinate = mp_oControl->GetMathLib()->GetXCoordinateFromDate(dtBuff);
	switch (oTickMark->GetTickMarkType()) 
	{
	case TLT_BIG:
		lTickMarkHeight = mp_lBigTickMarkHeight;
		break;
	case TLT_MEDIUM:
		lTickMarkHeight = mp_lMediumTickMarkHeight;
		break;
	case TLT_SMALL:
		lTickMarkHeight = mp_lSmallTickMarkHeight;
		break;
	}
	mp_oControl->clsG->mp_DrawLine(fXCoordinate, mp_oTimeLine->GetBottom() - lTickMarkHeight, fXCoordinate, mp_oTimeLine->GetBottom(), LT_NORMAL, mp_oStyle->GetBorderColor(), LDS_SOLID);
}

void clsTickMarkArea::PaintText(LONG fXCoordinate, clsTickMark* oTickMark)
{
	COleDateTime dtDateBuff;
	CString sDateBuff = "";
	LONG lLeft = 0;
	LONG lTop = 0;
	LONG lRight = 0;
	LONG lBottom = 0;
	LONG lStringWidth = 0;
	LONG lStringHeight = 0;
	dtDateBuff = mp_oControl->GetMathLib()->GetDateFromXCoordinate(fXCoordinate);
	dtDateBuff = mp_oControl->GetMathLib()->RoundDate(oTickMark->GetInterval(), oTickMark->GetFactor(), dtDateBuff);
	sDateBuff = mp_oControl->GetConfiguration()->Format(dtDateBuff, oTickMark->GetTextFormat());

	lStringWidth = mp_oControl->clsG->mp_lStrWidth(sDateBuff, mp_oStyle->GetFont());
	lStringHeight = mp_oControl->clsG->mp_lStrHeight(sDateBuff, mp_oStyle->GetFont());
	lLeft = fXCoordinate - (lStringWidth / 2) - 10;
	lTop = mp_oTimeLine->GetBottom() - mp_lTextOffset - lStringHeight;
	lRight = fXCoordinate + (lStringWidth / 2) + 10;
	lBottom = lTop + lStringHeight;
	mp_oControl->GetTickMarkTextDrawEventArgs()->Clear();
	mp_oControl->GetTickMarkTextDrawEventArgs()->mp_dtDate = dtDateBuff;
    mp_oControl->GetTickMarkTextDrawEventArgs()->SetText(sDateBuff);
    mp_oControl->GetTickMarkTextDrawEventArgs()->SetCustomDraw(FALSE);
    mp_oControl->FireTickMarkTextDraw();
	if (mp_oControl->GetTickMarkTextDrawEventArgs()->GetCustomDraw() == FALSE)
	{
		mp_oControl->clsG->mp_DrawAlignedText(lLeft, lTop, lRight, lBottom, sDateBuff, HAL_CENTER, VAL_CENTER, mp_oStyle->GetForeColor(), mp_oStyle->GetFont());
	}
	else
	{
		mp_oControl->clsG->mp_DrawAlignedText(lLeft, lTop, lRight, lBottom, mp_oControl->GetTickMarkTextDrawEventArgs()->GetText(), HAL_CENTER, VAL_CENTER, mp_oStyle->GetForeColor(), mp_oStyle->GetFont());
	}
}

CString clsTickMarkArea::GetXML(void)
{
	clsXML oXML(mp_oControl, "TickMarkArea");
	oXML.InitializeWriter();
	oXML.WriteProperty("BigTickMarkHeight", mp_lBigTickMarkHeight);
	oXML.WriteProperty("Height", mp_lHeight);
	oXML.WriteProperty("MediumTickMarkHeight", mp_lMediumTickMarkHeight);
	oXML.WriteProperty("SmallTickMarkHeight", mp_lSmallTickMarkHeight);
	oXML.WriteProperty("StyleIndex", mp_sStyleIndex);
	oXML.WriteProperty("TextOffset", mp_lTextOffset);
	oXML.WriteProperty("Visible", mp_bVisible);
	oXML.WriteObject(mp_oTickMarks->GetXML());
	return oXML.GetXML();
}

void clsTickMarkArea::SetXML(CString sXML)
{
	clsXML oXML(mp_oControl, "TickMarkArea");
	oXML.SetXML(sXML);
	oXML.InitializeReader();
	oXML.ReadProperty("BigTickMarkHeight", mp_lBigTickMarkHeight);
	oXML.ReadProperty("Height", mp_lHeight);
	oXML.ReadProperty("MediumTickMarkHeight", mp_lMediumTickMarkHeight);
	oXML.ReadProperty("SmallTickMarkHeight", mp_lSmallTickMarkHeight);
	oXML.ReadProperty("StyleIndex", mp_sStyleIndex);
	SetStyleIndex(mp_sStyleIndex);
	oXML.ReadProperty("TextOffset", mp_lTextOffset);
	oXML.ReadProperty("Visible", mp_bVisible);
	mp_oTickMarks->SetXML(oXML.ReadObject("TickMarks"));
}

void clsTickMarkArea::Clear()
{
	mp_sStyleIndex = "DS_TICKMARKAREA";
	mp_oStyle = mp_oControl->GetStyles()->FItem("DS_TICKMARKAREA");
	mp_lHeight = 23;
	mp_lBigTickMarkHeight = 12;
	mp_lMediumTickMarkHeight = 9;
	mp_lSmallTickMarkHeight = 7;
	mp_bVisible = TRUE;
	mp_lTextOffset = 11;
	GetTickMarks()->Clear();
}

void clsTickMarkArea::Clone(clsTickMarkArea* oClone)
{
    oClone->SetBigTickMarkHeight(mp_lBigTickMarkHeight);
    oClone->SetHeight(mp_lHeight);
    oClone->SetMediumTickMarkHeight(mp_lMediumTickMarkHeight);
    oClone->SetSmallTickMarkHeight(mp_lSmallTickMarkHeight);
    oClone->SetStyleIndex(mp_sStyleIndex);
    oClone->SetTextOffset(mp_lTextOffset);
    oClone->SetVisible(mp_bVisible);
    GetTickMarks()->Clone(oClone->GetTickMarks());
}