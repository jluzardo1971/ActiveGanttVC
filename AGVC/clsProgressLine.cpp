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
#include "clsProgressLine.h"



IMPLEMENT_DYNCREATE(clsProgressLine, CCmdTarget)

//{9930967C-34EF-40E6-9EF6-96F34EECCF2C}
static const IID IID_IclsProgressLine = { 0x9930967C, 0x34EF, 0x40E6, { 0x9E, 0xF6, 0x96, 0xF3, 0x4E, 0xEC, 0xCF, 0x2C} };

//{DC07329D-92E4-4C79-973C-07E9B299560D}
IMPLEMENT_OLECREATE_FLAGS(clsProgressLine, "AGVC.clsProgressLine", afxRegApartmentThreading, 0xDC07329D, 0x92E4, 0x4C79, 0x97, 0x3C, 0x07, 0xE9, 0xB2, 0x99, 0x56, 0x0D)

BEGIN_DISPATCH_MAP(clsProgressLine, CCmdTarget)
	DISP_PROPERTY_EX_ID(clsProgressLine, "Position", 1, odl_GetPosition, odl_SetPosition, VT_DATE) 
	DISP_PROPERTY_EX_ID(clsProgressLine, "ForeColor", 2, odl_GetForeColor, odl_SetForeColor, VT_COLOR) 
	DISP_PROPERTY_EX_ID(clsProgressLine, "Length", 3, odl_GetLength, odl_SetLength, VT_I4) 
	DISP_PROPERTY_EX_ID(clsProgressLine, "LineType", 4, odl_GetLineType, odl_SetLineType, VT_I4) 
	DISP_FUNCTION_ID(clsProgressLine, "GetXML", 5, odl_GetXML, VT_BSTR, VTS_NONE) 
	DISP_FUNCTION_ID(clsProgressLine, "SetXML", 6, odl_SetXML, VT_EMPTY, VTS_BSTR)
	DISP_PROPERTY_EX_ID(clsProgressLine, "Width", 7, odl_GetWidth, odl_SetWidth, VT_I4)
	DISP_PROPERTY_EX_ID(clsProgressLine, "LineStyle", 8, odl_GetLineStyle, odl_SetLineStyle, VT_I4)
	DISP_FUNCTION_ID(clsProgressLine, "SetColor", 9, odl_SetColor, VT_EMPTY, VTS_I4 VTS_COLOR)
	DISP_FUNCTION_ID(clsProgressLine, "GetColor", 10, odl_GetColor, VT_COLOR, VTS_I4)
	DISP_PROPERTY_EX_ID(clsProgressLine, "StyleIndex", 11, odl_GetStyleIndex, odl_SetStyleIndex, VT_BSTR) 
	DISP_PROPERTY_EX_ID(clsProgressLine, "Style", 12, odl_GetStyle, SetNotSupported, VT_DISPATCH)
END_DISPATCH_MAP()

BEGIN_INTERFACE_MAP(clsProgressLine, CCmdTarget)
	INTERFACE_PART(clsProgressLine, IID_IclsProgressLine, Dispatch)
END_INTERFACE_MAP()

BEGIN_MESSAGE_MAP(clsProgressLine, CCmdTarget)
END_MESSAGE_MAP()

clsProgressLine::clsProgressLine()
{
	AfxMessageBox(_T("Invalid constructor has been invoked for clsProgressLine. Please inform support@sourcecodestore.com of this error"), MB_OK | MB_ICONSTOP | MB_APPLMODAL);
}

clsProgressLine::clsProgressLine(CActiveGanttVCCtl* oControl)
{
	EnableAutomation();
	AfxOleLockApp();
	mp_oControl = oControl;
	Clear();
}

clsProgressLine::~clsProgressLine()
{
	AfxOleUnlockApp();
}

void clsProgressLine::OnFinalRelease()
{
	CCmdTarget::OnFinalRelease();
}

COleDateTime clsProgressLine::GetPosition(void)
{
	return mp_dtPosition;
}

void clsProgressLine::SetPosition(COleDateTime Value)
{
	mp_dtPosition = Value;
}

Gdiplus::Color clsProgressLine::GetForeColor(void)
{
	return mp_clrForeColor;
}

void clsProgressLine::SetForeColor(Gdiplus::Color Value)
{
	mp_clrForeColor = Value;
}

LONG clsProgressLine::GetLength(void)
{
	return mp_yLength;
}

void clsProgressLine::SetLength(LONG Value)
{
	mp_yLength = Value;
}

LONG clsProgressLine::GetLineType(void)
{
	return mp_yLineType;
}

void clsProgressLine::SetLineType(LONG Value)
{
	mp_yLineType = Value;
}

void clsProgressLine::SetColor(LONG Index, Gdiplus::Color oColor)
{
	Index = Index - 1;
    if (mp_yLineStyle != PLT_USERDEFINED)
    {
        mp_oControl->mp_ErrorReport(PROGRESSLINE_INVALIDOP, "Invalid Operation", "ActiveGanttVCCtl.clsProgressLine.SetColor");
        return;
    }

    if (Index < 0)
    {
        mp_oControl->mp_ErrorReport(PROGRESSLINE_INVALID_INDEX, "Invalid Index", "ActiveGanttVCCtl.clsProgressLine.SetColor");
        return;
    }
    if (Index > mp_aColors.GetCount() - 1)
    {
        mp_oControl->mp_ErrorReport(PROGRESSLINE_INVALID_INDEX, "Invalid Index", "ActiveGanttVCCtl.clsProgressLine.SetColor");
        return;
    }
    mp_aColors.SetAt(Index, oColor);
}

Gdiplus::Color clsProgressLine::GetColor(LONG Index)
{
	Index = Index - 1;
	        
    if (mp_yLineStyle != PLT_USERDEFINED)
    {
        mp_oControl->mp_ErrorReport(PROGRESSLINE_INVALIDOP, "Invalid Operation", "ActiveGanttVCCtl.clsProgressLine.GetColor");
        return NULL;
    }
    if (Index < 0)
    {
        mp_oControl->mp_ErrorReport(PROGRESSLINE_INVALID_INDEX, "Invalid Index", "ActiveGanttVCCtl.clsProgressLine.GetColor");
        return NULL;
    }
    if (Index > mp_aColors.GetCount() - 1)
    {
        mp_oControl->mp_ErrorReport(PROGRESSLINE_INVALID_INDEX, "Invalid Index", "ActiveGanttCSNCtl.clsProgressLine.GetColor");
        return NULL;
    }
	Gdiplus::Color oReturnColor;
	oReturnColor = mp_aColors.GetAt(Index);
	return oReturnColor;
}

LONG clsProgressLine::GetLineStyle(void)
{
	return mp_yLineStyle;
}

void clsProgressLine::SetLineStyle(LONG Value)
{
	mp_yLineStyle = Value;
}

LONG clsProgressLine::GetWidth(void)
{
    if (mp_yLineStyle == PLT_STANDARD)
    {
        return 1;
    }
    else
    {
        return mp_lWidth;
    }
}

void clsProgressLine::SetWidth(LONG Value)
{
    if (mp_yLineStyle == PLT_STANDARD)
    {
        mp_lWidth = 1;
    }
    else
    {
        if (Value < 0)
        {
            mp_oControl->mp_ErrorReport(PROGRESSLINE_INVALID_WIDTH, "Invalid Width Value", "ActiveGanttVCCtl.clsProgressLine.Width");
            return;
        }
        mp_lWidth = Value;
    }
    int i = 0;
    mp_aColors.RemoveAll();
    for (i = 0; i <= mp_lWidth - 1; i++)
    {
        mp_aColors.Add(Color::MakeARGB(255, 255, 255, 255));
    }
}

CString clsProgressLine::GetStyleIndex(void)
{
	if (mp_sStyleIndex == "DS_PROGRESSLINE") 
	{
		return "";
	}
	else 
	{
		return mp_sStyleIndex;
	}
}

void clsProgressLine::SetStyleIndex(CString Value)
{
	Value = Value.Trim();
	if (Value.GetLength() == 0) {Value = "DS_PROGRESSLINE";} 
	mp_sStyleIndex = Value;
	mp_oStyle = mp_oControl->GetStyles()->FItem(Value);
}

clsStyle* clsProgressLine::GetStyle(void)
{
	return mp_oStyle;
}

void clsProgressLine::Finalize(void)
{
}

clsTimeLine* clsProgressLine::mp_oTimeLine()
{
	return mp_oControl->GetCurrentViewObject()->GetTimeLine();
}

void clsProgressLine::Draw(void)
{
	LONG lX1 = 0;
	LONG lX2 = 0;
	E_PROGRESSLINELENGTH yTimeLineMarkerLength;
	COleDateTime dtDate;
    LONG lTop = 0;
    LONG lBottom = 0;
	if (mp_yLineType == TLMT_SYSTEMTIME)
	{
		SYSTEMTIME lt;
		GetLocalTime(&lt);
		dtDate.SetDateTime(lt.wYear, lt.wMonth, lt.wDay, lt.wHour, lt.wMinute, lt.wSecond);
	}
	else
	{
		dtDate = mp_dtPosition;
	}
	if (dtDate >= mp_oTimeLine()->GetStartDate() && dtDate <= mp_oTimeLine()->GetEndDate())
	{
		yTimeLineMarkerLength = (E_PROGRESSLINELENGTH)mp_yLength;
		lX1 = mp_oControl->GetMathLib()->GetXCoordinateFromDate(mp_dtPosition);
		    if (mp_oTimeLine()->GetTickMarkArea()->GetVisible() == FALSE && yTimeLineMarkerLength == TLMA_CA_TICKMARKAREA)
            {
                yTimeLineMarkerLength = TLMA_CLIENTAREA;
            }
            if (mp_oTimeLine()->GetTickMarkArea()->GetVisible() == FALSE && yTimeLineMarkerLength == TLMA_TICKMARKAREA)
            {
                return;
            }
            if ((mp_oTimeLine()->GetTickMarkArea()->GetVisible() == FALSE) && (yTimeLineMarkerLength == TLMA_TIMELINE) && (mp_oTimeLine()->GetTierArea()->GetLowerTier()->GetVisible() == FALSE) && (mp_oTimeLine()->GetTierArea()->GetMiddleTier()->GetVisible() == FALSE) && (mp_oTimeLine()->GetTierArea()->GetUpperTier()->GetVisible() == FALSE))
            {
                return;
            }

		switch (yTimeLineMarkerLength)
        {
            case TLMA_TICKMARKAREA:
                lTop = mp_oTimeLine()->TiersTickMarksPosition("TickMarkArea");
                lBottom = mp_oTimeLine()->GetBottom();
                break;
            case TLMA_CA_TIMELINE:
                lTop = mp_oTimeLine()->GetTop();
                lBottom = mp_oControl->GetCurrentViewObject()->GetClientArea()->GetBottom();
                break;
            case TLMA_CLIENTAREA:
                lTop = mp_oControl->GetCurrentViewObject()->GetClientArea()->GetTop();
                lBottom = mp_oControl->GetCurrentViewObject()->GetClientArea()->GetBottom();
                break;
            case TLMA_CA_TICKMARKAREA:
                lTop = mp_oTimeLine()->TiersTickMarksPosition("TickMarkArea");
                lBottom = mp_oControl->GetCurrentViewObject()->GetClientArea()->GetBottom();
                break;
            case TLMA_NONE:
                return;
        }

		mp_oControl->clsG->mp_ClipRegion(mp_oTimeLine()->Getf_lStart(), lTop, mp_oTimeLine()->Getf_lEnd(), lBottom, TRUE);
        switch (mp_yLineStyle)
        {
            case PLT_STANDARD:
                mp_oControl->clsG->mp_DrawLine(lX1, lTop, lX1, lBottom, LT_NORMAL, mp_clrForeColor, LDS_SOLID);
                break;
            case PLT_USERDEFINED:
				{
					lX1 = lX1 - (int) floor((double)mp_lWidth / 2);
					int i = 0;
					for (i = 0; i <= mp_aColors.GetCount() - 1; i++)
					{
						mp_oControl->clsG->mp_DrawLine(lX1 + i, lTop, lX1 + i, lBottom, LT_NORMAL, (Color)mp_aColors[i], LDS_SOLID);
					}
				}
                break;
            case PLT_STYLE:
                if (mp_lWidth > 1)
                {
                    lX1 = lX1 - (int) floor((double)mp_lWidth / 2);
                    lX2 = lX1 + mp_lWidth - 1;
                }
                else
                {
                    lX2 = lX1 + 1;
                }
                mp_oControl->clsG->mp_DrawItem(lX1, lTop, lX2, lBottom, "", "", FALSE, NULL, 0, 0, mp_oStyle);
                break;
        }
        mp_oControl->clsG->mp_ClearClipRegion();


	}
}

CString clsProgressLine::GetXML(void)
{
	clsXML oXML(mp_oControl, "ProgressLine");
	oXML.InitializeWriter();
	oXML.WriteProperty("ForeColor", mp_clrForeColor);
	oXML.WriteProperty("Length", mp_yLength);
	oXML.WriteProperty("LineType", mp_yLineType);
	oXML.WriteProperty("Position", mp_dtPosition);
	oXML.WriteProperty("LineStyle", mp_yLineStyle);
	oXML.WriteProperty("StyleIndex", mp_sStyleIndex);
	oXML.WriteProperty("Width", mp_lWidth);
	if (mp_yLineStyle == PLT_USERDEFINED)
	{
		LONG lColorCount = mp_aColors.GetCount();
		oXML.WriteProperty("ColorCount", lColorCount);
		int i = 0;
		for (i = 0; i <= mp_aColors.GetCount() - 1; i++)
		{
			oXML.WriteProperty("Color" + CStr(i), mp_aColors.GetAt(i));
		}
	}
	else
	{
		LONG lColorCount = 0;
		oXML.WriteProperty("ColorCount", lColorCount);
	}
	return oXML.GetXML();
}

void clsProgressLine::SetXML(CString sXML)
{
	clsXML oXML(mp_oControl, "ProgressLine");
	oXML.SetXML(sXML);
	oXML.InitializeReader();
	oXML.ReadProperty("ForeColor", mp_clrForeColor);
	oXML.ReadProperty("Length", mp_yLength);
	oXML.ReadProperty("LineType", mp_yLineType);
	oXML.ReadProperty("Position", mp_dtPosition);
	oXML.ReadProperty("LineStyle", mp_yLineStyle);
	oXML.ReadProperty("StyleIndex", mp_sStyleIndex);
	SetStyleIndex(mp_sStyleIndex);
	oXML.ReadProperty("Width", mp_lWidth);
	mp_aColors.RemoveAll();
	int lColorCount = 0;
	oXML.ReadProperty("ColorCount", lColorCount);
	if (lColorCount > 0)
	{
		int i = 0;
		for (i = 0; i <= lColorCount - 1; i++)
		{
			Gdiplus::Color oColor;
			oXML.ReadProperty("Color" + CStr(i), oColor);
			mp_aColors.Add(oColor);
		}
	}
	SetStyleIndex(mp_sStyleIndex);
}

void clsProgressLine::Clear()
{
	SetWidth(6);
	mp_clrForeColor = Color::Red;
	mp_dtPosition = (DATE)0;
	mp_yLength = TLMA_TICKMARKAREA;
	mp_yLineType = TLMT_SYSTEMTIME;
	mp_yLineStyle = SA_APPEARANCE;
	mp_sStyleIndex = "DS_PROGRESSLINE";
	mp_oStyle = mp_oControl->GetStyles()->FItem("DS_PROGRESSLINE");
}

void clsProgressLine::Clone(clsProgressLine* oClone)
{
    oClone->SetForeColor(mp_clrForeColor);
    oClone->SetLength(mp_yLength);
    oClone->SetLineType(mp_yLineType);
    oClone->SetLineStyle(mp_yLineStyle);
    oClone->SetPosition(mp_dtPosition);
    oClone->SetStyleIndex(mp_sStyleIndex);
    oClone->SetWidth(mp_lWidth);
    int i;
    for (i = 0; i <= mp_aColors.GetCount() - 1; i++)
	{
        oClone->mp_aColors.Add(mp_aColors.GetAt(i));
	}
}