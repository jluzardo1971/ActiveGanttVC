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
#include "clsSplitter.h"



IMPLEMENT_DYNCREATE(clsSplitter, CCmdTarget)

//{3D059D62-42B7-4801-B87D-346B478137F2}
static const IID IID_IclsSplitter = { 0x3D059D62, 0x42B7, 0x4801, { 0xB8, 0x7D, 0x34, 0x6B, 0x47, 0x81, 0x37, 0xF2} };

//{9D06745C-D1B8-4C80-AC32-5D59ED6D2D34}
IMPLEMENT_OLECREATE_FLAGS(clsSplitter, "AGVC.clsSplitter", afxRegApartmentThreading, 0x9D06745C, 0xD1B8, 0x4C80, 0xAC, 0x32, 0x5D, 0x59, 0xED, 0x6D, 0x2D, 0x34)

BEGIN_DISPATCH_MAP(clsSplitter, CCmdTarget)
	DISP_PROPERTY_EX_ID(clsSplitter, "Appearance", 1, odl_GetAppearance, odl_SetAppearance, VT_I4) 
	DISP_PROPERTY_EX_ID(clsSplitter, "Position", 2, odl_GetPosition, odl_SetPosition, VT_I4) 
	DISP_FUNCTION_ID(clsSplitter, "GetXML", 3, odl_GetXML, VT_BSTR, VTS_NONE) 
	DISP_FUNCTION_ID(clsSplitter, "SetXML", 4, odl_SetXML, VT_EMPTY, VTS_BSTR)
	DISP_FUNCTION_ID(clsSplitter, "SetColor", 5, odl_SetColor, VT_EMPTY, VTS_I4 VTS_COLOR)
	DISP_FUNCTION_ID(clsSplitter, "GetColor", 6, odl_GetColor, VT_COLOR, VTS_I4)
	DISP_PROPERTY_EX_ID(clsSplitter, "Type", 7, odl_GetType, odl_SetType, VT_I4)
	DISP_PROPERTY_EX_ID(clsSplitter, "Width", 8, odl_GetWidth, odl_SetWidth, VT_I4)
	DISP_PROPERTY_EX_ID(clsSplitter, "StyleIndex", 9, odl_GetStyleIndex, odl_SetStyleIndex, VT_BSTR) 
	DISP_PROPERTY_EX_ID(clsSplitter, "Style", 10, odl_GetStyle, SetNotSupported, VT_DISPATCH)
	DISP_PROPERTY_EX_ID(clsSplitter, "Offset", 11, odl_GetOffset, odl_SetOffset, VT_I4) 
END_DISPATCH_MAP()

BEGIN_INTERFACE_MAP(clsSplitter, CCmdTarget)
	INTERFACE_PART(clsSplitter, IID_IclsSplitter, Dispatch)
END_INTERFACE_MAP()

BEGIN_MESSAGE_MAP(clsSplitter, CCmdTarget)
END_MESSAGE_MAP()

clsSplitter::clsSplitter()
{
	AfxMessageBox(_T("Invalid constructor has been invoked for clsSplitter. Please inform support@sourcecodestore.com of this error"), MB_OK | MB_ICONSTOP | MB_APPLMODAL);
}

clsSplitter::clsSplitter(CActiveGanttVCCtl* oControl)
{
	EnableAutomation();
	AfxOleLockApp();
	mp_oControl = oControl;
	mp_lPosition = 125;
	mp_yAppearance = TLB_3D;
	mp_yType = SA_APPEARANCE;
	SetWidth(6);
	mp_sStyleIndex = "DS_SPLITTER";
	mp_oStyle = mp_oControl->GetStyles()->FItem("DS_SPLITTER");
	mp_lOffset = -1;
}

clsSplitter::~clsSplitter()
{
	AfxOleUnlockApp();
}

void clsSplitter::OnFinalRelease()
{
	CCmdTarget::OnFinalRelease();
}

LONG clsSplitter::GetAppearance(void)
{
	return mp_yAppearance;
}

void clsSplitter::SetAppearance(LONG Value)
{
	mp_yAppearance = Value;
}

LONG clsSplitter::GetPosition(void)
{
	return mp_lPosition;
}

void clsSplitter::SetPosition(LONG Value)
{
	if ((Value <= 0))
	{
		return;
	}
	mp_lPosition = Value;
	if (mp_lPosition > (mp_oControl->GetColumns()->GetWidth() + mp_lOffset))
	{
		mp_lPosition = mp_oControl->GetColumns()->GetWidth() + mp_lOffset;
		mp_oControl->GetHorizontalScrollBar()->SetValue(0);
	}
}

LONG clsSplitter::GetOffset(void)
{
	return mp_lOffset;
}

void clsSplitter::SetOffset(LONG Value)
{
	mp_lOffset = Value;
}

void clsSplitter::SetColor(LONG Index, Gdiplus::Color oColor)
{
	Index = Index - 1;
    if (mp_yType != SA_USERDEFINED)
    {
        mp_oControl->mp_ErrorReport(SPLITTER_INVALIDOP, "Invalid Operation", "ActiveGanttVCCtl.clsSplitter.SetColor");
        return;
    }

    if (Index < 0)
    {
        mp_oControl->mp_ErrorReport(SPLITTER_INVALID_INDEX, "Invalid Index", "ActiveGanttVCCtl.clsSplitter.SetColor");
        return;
    }
    if (Index > mp_aColors.GetCount() - 1)
    {
        mp_oControl->mp_ErrorReport(SPLITTER_INVALID_INDEX, "Invalid Index", "ActiveGanttVCCtl.clsSplitter.SetColor");
        return;
    }
    mp_aColors.SetAt(Index, oColor);
}

Gdiplus::Color clsSplitter::GetColor(LONG Index)
{
	Index = Index - 1;
	        
    if (mp_yType != SA_USERDEFINED)
    {
        mp_oControl->mp_ErrorReport(SPLITTER_INVALIDOP, "Invalid Operation", "ActiveGanttVCCtl.clsSplitter.GetColor");
        return NULL;
    }
    if (Index < 0)
    {
        mp_oControl->mp_ErrorReport(SPLITTER_INVALID_INDEX, "Invalid Index", "ActiveGanttVCCtl.clsSplitter.GetColor");
        return NULL;
    }
    if (Index > mp_aColors.GetCount() - 1)
    {
        mp_oControl->mp_ErrorReport(SPLITTER_INVALID_INDEX, "Invalid Index", "ActiveGanttCSNCtl.clsSplitter.GetColor");
        return NULL;
    }
	Gdiplus::Color oReturnColor;
	oReturnColor = mp_aColors.GetAt(Index);
	return oReturnColor;
}

LONG clsSplitter::GetType(void)
{
	return mp_yType;
}

void clsSplitter::SetType(LONG Value)
{
	mp_yType = Value;
}

LONG clsSplitter::GetWidth(void)
{
    if (mp_yType == SA_APPEARANCE)
    {
        return 6;
    }
    else
    {
        return mp_lWidth;
    }
}

void clsSplitter::SetWidth(LONG Value)
{
    if (mp_yType == SA_APPEARANCE)
    {
        mp_lWidth = 6;
    }
    else
    {
        if (Value < 0)
        {
            mp_oControl->mp_ErrorReport(SPLITTER_INVALID_WIDTH, "Invalid Width Value", "ActiveGanttVCCtl.clsSplitter.Width");
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

CString clsSplitter::GetStyleIndex(void)
{
	if (mp_sStyleIndex == "DS_SPLITTER") 
	{
		return "";
	}
	else 
	{
		return mp_sStyleIndex;
	}
}

void clsSplitter::SetStyleIndex(CString Value)
{
	Value = Value.Trim();
	if (Value.GetLength() == 0) {Value = "DS_SPLITTER";} 
	mp_sStyleIndex = Value;
	mp_oStyle = mp_oControl->GetStyles()->FItem(Value);
}

clsStyle* clsSplitter::GetStyle(void)
{
	return mp_oStyle;
}

LONG clsSplitter::GetLeft(void)
{
	if (mp_oControl->GetColumns()->GetCount() != 0)
	{
		return GetPosition() + mp_oControl->Getmt_BorderThickness() - 1;
	}
	else
	{
		if ((mp_oControl->Getf_UserMode() == TRUE))
		{
			return 0;
		}
		else
		{
			return 125 + mp_oControl->Getmt_BorderThickness() - 1;
		}
	}
}

LONG clsSplitter::GetRight(void)
{
	if (mp_oControl->GetColumns()->GetCount() != 0)
	{
		return GetPosition() + mp_oControl->Getmt_BorderThickness() + this->GetWidth();
	}
	else
	{
		if ((mp_oControl->Getf_UserMode() == TRUE))
		{
			return mp_oControl->Getmt_BorderThickness();
		}
		else
		{
			return 125 + mp_oControl->Getmt_BorderThickness() + this->GetWidth();
		}
	}
}

void clsSplitter::Finalize(void)
{
}

void clsSplitter::Draw(void)
{
	if (mp_oControl->GetColumns()->GetCount() == 0 && mp_oControl->Getf_UserMode() == TRUE)
	{
		return;
	}
	mp_oControl->clsG->mp_ClipRegion(GetLeft() + 1, 0, GetLeft() + 6, mp_oControl->Getmt_BottomMargin(), TRUE);
    if (mp_yType == SA_STYLE)
    {
        mp_oControl->clsG->mp_DrawItem(this->GetLeft() + 1, 0, this->GetLeft() + this->GetWidth() + 1, mp_oControl->clsG->Height(), "", "", FALSE, NULL, 0, 0, this->mp_oStyle);
    }
    else
    {
        int i = 0;
        if (mp_yType == SA_APPEARANCE)
        {
            mp_aColors.RemoveAll();
            for (i = 0; i <= 5; i++)
            {
                mp_aColors.Add(Color::MakeARGB(255, 255, 255, 255));
            }
            switch (mp_yAppearance)
            {
                case TLB_3D:
					mp_aColors.SetAt(0, Color::MakeARGB(255, 192, 192, 192));
					mp_aColors.SetAt(1, Color::White);
                    mp_aColors.SetAt(2, Color::MakeARGB(255, 192, 192, 192));
                    mp_aColors.SetAt(3, Color::MakeARGB(255, 192, 192, 192));
                    mp_aColors.SetAt(4, Color::MakeARGB(255, 64, 64, 64));
                    mp_aColors.SetAt(5, Color::MakeARGB(255, 66, 66, 66));
                    break;
                case TLB_NONE:
                    mp_aColors.SetAt(0, Color::MakeARGB(255, 192, 192, 192));
                    mp_aColors.SetAt(1, Color::MakeARGB(255, 192, 192, 192));
                    mp_aColors.SetAt(2, Color::MakeARGB(255, 192, 192, 192));
                    mp_aColors.SetAt(3, Color::MakeARGB(255, 192, 192, 192));
                    mp_aColors.SetAt(4, Color::MakeARGB(255, 192, 192, 192));
                    mp_aColors.SetAt(5, Color::MakeARGB(255, 192, 192, 192));
                    break;
                case TLB_SINGLE:
					mp_aColors.SetAt(0, Color::Black);
                    mp_aColors.SetAt(1, Color::MakeARGB(255, 192, 192, 192));
                    mp_aColors.SetAt(2, Color::MakeARGB(255, 192, 192, 192));
                    mp_aColors.SetAt(3, Color::MakeARGB(255, 192, 192, 192));
                    mp_aColors.SetAt(4, Color::MakeARGB(255, 192, 192, 192));
                    mp_aColors.SetAt(5, Color::Black);
                    break;
            }
        }
        for (i = 0; i <= mp_aColors.GetCount() - 1; i++)
        {
            mp_oControl->clsG->mp_DrawLine(this->GetLeft() + i + 1, 0, this->GetLeft() + i + 1, mp_oControl->clsG->Height(), LT_NORMAL, mp_aColors.GetAt(i), LDS_SOLID);
        }
    }
}

void clsSplitter::f_AdjustPosition(void)
{
	LONG lWidth;
	lWidth = mp_oControl->GetColumns()->GetWidth() + mp_lOffset;
	if (lWidth < (mp_oControl->clsG->Width() - 25))
	{
		if (mp_lPosition < lWidth)
		{
			mp_lPosition = lWidth;
		}
	}
	if (mp_lPosition > lWidth)
	{
		mp_lPosition = lWidth;
		mp_oControl->GetHorizontalScrollBar()->SetValue(0);
	}
}

CString clsSplitter::GetXML(void)
{
	clsXML oXML(mp_oControl, "Splitter");
	oXML.InitializeWriter();
	oXML.WriteProperty("Appearance", mp_yAppearance);
	oXML.WriteProperty("Offset", mp_lOffset);
	oXML.WriteProperty("Position", mp_lPosition);
	oXML.WriteProperty("Type", mp_yType);
	oXML.WriteProperty("StyleIndex", mp_sStyleIndex);
	oXML.WriteProperty("Width", mp_lWidth);
	if (mp_yType == SA_USERDEFINED)
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

void clsSplitter::SetXML(CString sXML)
{
	clsXML oXML(mp_oControl, "Splitter");
	oXML.SetXML(sXML);
	oXML.InitializeReader();
	oXML.ReadProperty("Appearance", mp_yAppearance);
	oXML.ReadProperty("Offset", mp_lOffset);
	oXML.ReadProperty("Position", mp_lPosition);
	SetPosition(mp_lPosition);
	oXML.ReadProperty("Type", mp_yType);
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