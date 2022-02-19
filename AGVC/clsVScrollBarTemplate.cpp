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
#include "clsVScrollBarTemplate.h"



IMPLEMENT_DYNCREATE(clsVScrollBarTemplate, CCmdTarget)

//{EC67DC20-E7E5-486D-8F25-8207EB4BA079}
static const IID IID_IclsVScrollBarTemplate = { 0xEC67DC20, 0xE7E5, 0x486D, { 0x8F, 0x25, 0x82, 0x07, 0xEB, 0x4B, 0xA0, 0x79} };

//{6C5157A6-9DF1-4405-946F-920FB3BD95EE}
IMPLEMENT_OLECREATE_FLAGS(clsVScrollBarTemplate, "AGVC.clsVScrollBarTemplate", afxRegApartmentThreading, 0x6C5157A6, 0x9DF1, 0x4405, 0x94, 0x6F, 0x92, 0x0F, 0xB3, 0xBD, 0x95, 0xEE)

BEGIN_DISPATCH_MAP(clsVScrollBarTemplate, CCmdTarget)
	DISP_PROPERTY_EX_ID(clsVScrollBarTemplate, "TimerInterval", 1, odl_GetTimerInterval, odl_SetTimerInterval, VT_I4)
	DISP_PROPERTY_EX_ID(clsVScrollBarTemplate, "StyleIndex", 2, odl_GetStyleIndex, odl_SetStyleIndex, VT_BSTR) 
	DISP_PROPERTY_EX_ID(clsVScrollBarTemplate, "Style", 3, odl_GetStyle, SetNotSupported, VT_DISPATCH)
	DISP_PROPERTY_EX_ID(clsVScrollBarTemplate, "ArrowButtons", 4, odl_GetArrowButtons, SetNotSupported, VT_DISPATCH) 
	DISP_PROPERTY_EX_ID(clsVScrollBarTemplate, "ThumbButton", 5, odl_GetThumbButton, SetNotSupported, VT_DISPATCH)
	DISP_FUNCTION_ID(clsVScrollBarTemplate, "GetXML", 6, odl_GetXML, VT_BSTR, VTS_NONE) 
	DISP_FUNCTION_ID(clsVScrollBarTemplate, "SetXML", 7, odl_SetXML, VT_EMPTY, VTS_BSTR)
END_DISPATCH_MAP()

BEGIN_INTERFACE_MAP(clsVScrollBarTemplate, CCmdTarget)
	INTERFACE_PART(clsVScrollBarTemplate, IID_IclsVScrollBarTemplate, Dispatch)
END_INTERFACE_MAP()

BEGIN_MESSAGE_MAP(clsVScrollBarTemplate, CCmdTarget)
END_MESSAGE_MAP()

clsVScrollBarTemplate::clsVScrollBarTemplate()
{
	AfxMessageBox(_T("Invalid constructor has been invoked for clsVScrollBarTemplate. Please inform support@sourcecodestore.com of this error"), MB_OK | MB_ICONSTOP | MB_APPLMODAL);
}

clsVScrollBarTemplate::clsVScrollBarTemplate(CActiveGanttVCCtl* oControl)
{
	EnableAutomation();
	AfxOleLockApp();
	mp_oControl = oControl;
	mp_lSmallChange = 0;
	mp_lLargeChange = 0;
	mp_lValue = 0;
	mp_lMin = 0;
	mp_lMax = 0;
	mp_oArrowButtons = new clsButtonState(mp_oControl, "Arrow");
	mp_oThumbButton = new clsButtonState(mp_oControl, "Thumb");
	mp_sStyleIndex = "DS_SCROLLBAR";
	mp_oStyle = mp_oControl->GetStyles()->FItem("DS_SCROLLBAR");
}

clsVScrollBarTemplate::~clsVScrollBarTemplate()
{
	delete mp_oThumbButton;
	mp_oThumbButton = NULL;
	delete mp_oArrowButtons;
	mp_oArrowButtons = NULL;
	AfxOleUnlockApp();
}

void clsVScrollBarTemplate::OnFinalRelease()
{
	CCmdTarget::OnFinalRelease();
}

LONG clsVScrollBarTemplate::GetTimerInterval(void)
{
	return mp_lTimerInterval;
}

void clsVScrollBarTemplate::SetTimerInterval(LONG Value)
{
	mp_lTimerInterval = Value;
}

BOOL clsVScrollBarTemplate::GetEnabled(void)
{
	return mp_bEnabled;
}

void clsVScrollBarTemplate::SetEnabled(BOOL Value)
{
	mp_bEnabled = Value;
}

LONG clsVScrollBarTemplate::GetValue(void)
{
	return mp_lValue;
}

void clsVScrollBarTemplate::SetValue(LONG Value)
{
	mp_lValue = Value;
	if (mp_lValue < mp_lMin)
	{
//		throw new Exception("Value is less than mp_lMin.");
	}
}

LONG clsVScrollBarTemplate::GetMin(void)
{
	return mp_lMin;
}

void clsVScrollBarTemplate::SetMin(LONG Value)
{
	mp_lMin = Value;
}

LONG clsVScrollBarTemplate::GetMax(void)
{
	return mp_lMax;
}

void clsVScrollBarTemplate::SetMax(LONG Value)
{
	mp_lMax = Value;
	if (mp_lMax < mp_lMin)
	{
//		throw new Exception("mp_lMax is less than mp_lMin.");
	}
}

LONG clsVScrollBarTemplate::GetSmallChange(void)
{
	return mp_lSmallChange;
}

void clsVScrollBarTemplate::SetSmallChange(LONG Value)
{
	mp_lSmallChange = Value;
}

LONG clsVScrollBarTemplate::GetLargeChange(void)
{
	return mp_lLargeChange;
}

void clsVScrollBarTemplate::SetLargeChange(LONG Value)
{
	mp_lLargeChange = Value;
}

LONG clsVScrollBarTemplate::GetState(void)
{
	return mp_yState;
}

void clsVScrollBarTemplate::SetState(LONG Value)
{
	mp_yState = Value;
	switch (mp_yState)
	{
	case SS_CANTDISPLAY:
		mp_yState = SS_HIDDEN;
		this->Visible = FALSE;
		break;
	case SS_NOTNEEDED:
		if (mp_oControl->GetScrollBarBehaviour() == SB_DISABLE)
		{
			mp_yState = SS_SHOWN;
			this->mp_bEnabled = FALSE;
			this->Visible = TRUE;
		}
		else
		{
			mp_yState = SS_HIDDEN;
			this->Visible = FALSE;
		}
		break;
	case SS_NEEDED:
		mp_yState = SS_SHOWN;
		this->mp_bEnabled = TRUE;
		this->Visible = TRUE;
		break;
	}
}

void clsVScrollBarTemplate::Initialize(E_SCROLLBAR yType)
{
	mp_lValueBuff = GetValue();
	Width = 17;
	mp_iButton = BTN_NONE;
	mp_iTimerInterval = 50;
	mp_lValueBuff = GetValue();
	mp_yType = yType;
	mp_lValue = 0;
}

void clsVScrollBarTemplate::Draw(void)
{
	if (Visible == FALSE)
	{
		return;
	}
    clsStyle* oArrowButtonUpStyle = GetArrowButtons()->GetNormalStyle();
    clsStyle* oArrowButtonDownStyle = GetArrowButtons()->GetNormalStyle();
    clsStyle* oThumbButtonStyle = GetThumbButton()->GetNormalStyle();
	if ((mp_lMax - mp_lMin) == 0)
	{
		return;
	}
	mp_oControl->clsG->mp_DrawItem(Left, Top, Left + Width - 1, Top + Height - 1, "", "", FALSE, NULL, 0, 0, mp_oStyle);
	CalculateV();
	if (mp_bEnabled == FALSE)
	{
		oThumbButtonStyle = GetThumbButton()->GetDisabledStyle();
		oArrowButtonUpStyle = GetArrowButtons()->GetDisabledStyle();
		oArrowButtonDownStyle = GetArrowButtons()->GetDisabledStyle();
	}
	else if (mp_bMouseDown == TRUE)
	{
		if (mp_iButton == BTN_UP)
		{
			oThumbButtonStyle = GetThumbButton()->GetNormalStyle();
			oArrowButtonUpStyle = GetArrowButtons()->GetPressedStyle();
			oArrowButtonDownStyle = GetArrowButtons()->GetNormalStyle();
		}
		else if (mp_iButton == BTN_DOWN)
		{
            oThumbButtonStyle = GetThumbButton()->GetNormalStyle();
            oArrowButtonUpStyle = GetArrowButtons()->GetNormalStyle();
            oArrowButtonDownStyle = GetArrowButtons()->GetPressedStyle();
		}
		else if (mp_iButton == BTN_BUTTON)
		{
            oThumbButtonStyle = GetThumbButton()->GetPressedStyle();
            oArrowButtonUpStyle = GetArrowButtons()->GetNormalStyle();
            oArrowButtonDownStyle = GetArrowButtons()->GetNormalStyle();
		}
		else
		{
            oThumbButtonStyle = GetThumbButton()->GetNormalStyle();
            oArrowButtonUpStyle = GetArrowButtons()->GetNormalStyle();
            oArrowButtonDownStyle = GetArrowButtons()->GetNormalStyle();
		}
	}
	else
	{
        oArrowButtonUpStyle = GetArrowButtons()->GetNormalStyle();
        oArrowButtonDownStyle = GetArrowButtons()->GetNormalStyle();
        oThumbButtonStyle = GetThumbButton()->GetNormalStyle();
	}
	mp_oControl->clsG->mp_DrawItem(Left, Top, Left + Width - 1, Top + 16, "", "", FALSE, NULL, 0, 0, oArrowButtonUpStyle);
	mp_oControl->clsG->mp_DrawItem(Left, Top + Height - 17, Left + Width - 1, Top + Height - 1, "", "", FALSE, NULL, 0, 0, oArrowButtonDownStyle);
	mp_oControl->clsG->mp_DrawItem(Left + ButtonX1 - 1, Top + ButtonY1, Left + ButtonX2 - 1, Top + ButtonY2 - 2, "", "", FALSE, NULL, 0, 0, oThumbButtonStyle);
    if (oArrowButtonUpStyle->GetScrollBarStyle()->GetDropShadow() == TRUE)
    {
        mp_oControl->clsG->mp_DrawArrow(Left + oArrowButtonUpStyle->GetScrollBarStyle()->GetDropShadowUpX(), Top + oArrowButtonUpStyle->GetScrollBarStyle()->GetDropShadowUpY(), AWD_UP, oArrowButtonUpStyle->GetScrollBarStyle()->GetArrowSize(), oArrowButtonUpStyle->GetScrollBarStyle()->GetDropShadowArrowColor());
    }
    mp_oControl->clsG->mp_DrawArrow(Left + oArrowButtonUpStyle->GetScrollBarStyle()->GetUpX(), Top + oArrowButtonUpStyle->GetScrollBarStyle()->GetUpY(), AWD_UP, oArrowButtonUpStyle->GetScrollBarStyle()->GetArrowSize(), oArrowButtonUpStyle->GetScrollBarStyle()->GetArrowColor());
    if (oArrowButtonDownStyle->GetScrollBarStyle()->GetDropShadow() == TRUE)
    {
        mp_oControl->clsG->mp_DrawArrow(Left + oArrowButtonDownStyle->GetScrollBarStyle()->GetDropShadowDownX(), Top + Height - 17 + oArrowButtonDownStyle->GetScrollBarStyle()->GetDropShadowDownY(), AWD_DOWN, oArrowButtonDownStyle->GetScrollBarStyle()->GetArrowSize(), oArrowButtonDownStyle->GetScrollBarStyle()->GetDropShadowArrowColor());
    }
    mp_oControl->clsG->mp_DrawArrow(Left + oArrowButtonDownStyle->GetScrollBarStyle()->GetDownX(), Top + Height - 17 + oArrowButtonDownStyle->GetScrollBarStyle()->GetDownY(), AWD_DOWN, oArrowButtonDownStyle->GetScrollBarStyle()->GetArrowSize(), oArrowButtonDownStyle->GetScrollBarStyle()->GetArrowColor());
}

void clsVScrollBarTemplate::OnMouseHover(LONG X,LONG Y)
{
	mp_oControl->MouseKeyboardEvents->mp_SetCursor(CT_NORMAL);
}

void clsVScrollBarTemplate::OnMouseDown(LONG X,LONG Y)
{
	if (mp_bEnabled == FALSE)
	{
		return;
	}
	if (Visible == FALSE)
	{
		return;
	}
	mp_oControl->MouseKeyboardEvents->mp_SetCursor(CT_NORMAL);
	mp_bMouseDown = TRUE;
	mp_oControl->SetTimerEx(1000 + mp_yType, (UINT)mp_iTimerInterval);
	ScrollBarClick(X, Y);
}

void clsVScrollBarTemplate::OnMouseMove(LONG X,LONG Y)
{
	if (mp_bEnabled == FALSE)
	{
		return;
	}
	if (Visible == FALSE)
	{
		return;
	}
	ConvertToRelativeCoords(X, Y);
	mp_oControl->MouseKeyboardEvents->mp_SetCursor(CT_NORMAL);
	if (mp_iButton == BTN_BUTTON)
	{
		LONG mp_iProjectedValue = mp_lValue + (InvCalculateV(Y) - InvCalculateV(mp_iMouseYPosition));
		if (mp_iProjectedValue >= mp_lMin && mp_iProjectedValue <= mp_lMax)
		{
			mp_iMouseYPosition = Y;
			mp_lValue = mp_iProjectedValue;
			mp_ValueChanged();
		}
		else if (mp_iProjectedValue < mp_lMin)
		{
			mp_iMouseYPosition = Y;
			mp_lValue = mp_lMin;
			mp_ValueChanged();
		}
		else if (mp_iProjectedValue > mp_lMax)
		{
			mp_iMouseYPosition = Y;
			mp_lValue = mp_lMax;
			mp_ValueChanged();
		}
	}
}

void clsVScrollBarTemplate::OnMouseUp(void)
{
	if (mp_bEnabled == FALSE)
	{
		return;
	}
	if (Visible == FALSE)
	{
		return;
	}
	mp_oControl->MouseKeyboardEvents->mp_SetCursor(CT_NORMAL);
	mp_bMouseDown = FALSE;
	mp_oControl->KillTimerEx();
	mp_iButton = BTN_NONE;
	mp_oControl->Redraw();
}

BOOL clsVScrollBarTemplate::OverControl(LONG X,LONG Y)
{
	if (X >= Left && X <= Left + Width && Y >= Top && Y <= Top + Height)
	{
		return TRUE;
	}
	else
	{
		return FALSE;
	}
}

void clsVScrollBarTemplate::ConvertToRelativeCoords(LONG& X,LONG& Y)
{
	X = X - Left;
	Y = Y - Top;
}

BOOL clsVScrollBarTemplate::ScrollBarClick(LONG X,LONG Y)
{
	if (OverControl(X,Y) == FALSE)
	{
		return FALSE;
	}
	ConvertToRelativeCoords(X, Y);
	CalculateV();
	if (Y < 17)
	{
		mp_iButton = BTN_UP;
		SmallChangeUp();
		return TRUE;
	}
	else if (Y > 17 && Y < ButtonY1)
	{
		mp_iButton = BTN_UPLCHANGE;
		LargeChangeUp();
		return TRUE;
	}
	else if (Y >= ButtonY1 && Y <= ButtonY2)
	{
		mp_iButton = BTN_BUTTON;
		mp_iMouseYPosition = Y;
		return TRUE;
	}
	else if (Y > ButtonY2 && Y < Height - 17)
	{
		mp_iButton = BTN_DOWNLCHANGE;
		LargeChangeDown();
		return TRUE;
	}
	else if (Y > Height - 17 && Y < Height)
	{
		mp_iButton = BTN_DOWN;
		SmallChangeDown();
		return TRUE;
	}
	else
	{
		return FALSE;
	}
}

void clsVScrollBarTemplate::SmallChangeUp(void)
{
	if ((mp_lValue - mp_lSmallChange) >= mp_lMin)
	{
		mp_lValue = mp_lValue - mp_lSmallChange;
	}
	else
	{
		mp_lValue = mp_lMin;
	}
	mp_ValueChanged();
}

void clsVScrollBarTemplate::SmallChangeDown(void)
{
	if ((mp_lValue + mp_lSmallChange) <= mp_lMax)
	{
		mp_lValue = mp_lValue + mp_lSmallChange;
	}
	else
	{
		mp_lValue = mp_lMax;
	}
	mp_ValueChanged();
}

void clsVScrollBarTemplate::LargeChangeUp(void)
{
	if ((mp_lValue - mp_lLargeChange) >= mp_lMin)
	{
		mp_lValue = mp_lValue - mp_lLargeChange;
	}
	else if ((mp_lValue - mp_lLargeChange) < mp_lMin)
	{
		mp_lValue = mp_lMin;
	}
	mp_ValueChanged();
}

void clsVScrollBarTemplate::LargeChangeDown(void)
{
	if ((mp_lValue + mp_lLargeChange) <= mp_lMax)
	{
		mp_lValue = mp_lValue + mp_lLargeChange;
	}
	else if ((mp_lValue + mp_lLargeChange) > mp_lMax)
	{
		mp_lValue = mp_lMax;
	}
	mp_ValueChanged();
}

void clsVScrollBarTemplate::mp_ValueChanged(void)
{	
	if ((mp_lValue - mp_lValueBuff) != 0)
	{
		mp_oControl->VerticalScrollBar_ValueChanged(mp_lValue - mp_lValueBuff);
	}
	mp_lValueBuff = mp_lValue;
}

void clsVScrollBarTemplate::mp_oTimer_Tick(void)
{
	switch (mp_iButton)
	{
	case BTN_UP:
		SmallChangeUp();
		break;
	case BTN_UPLCHANGE:
		LargeChangeUp();
		break;
	case BTN_DOWNLCHANGE:
		LargeChangeDown();
		break;
	case BTN_DOWN:
		SmallChangeDown();
		break;
	}
}

void clsVScrollBarTemplate::CalculateV(void)
{
	LONG lHeight;
	LONG lItemLength = 0;
	lHeight = Height - (16 * 2) - 1;
	if (mp_lLargeChange > 0)
	{
		lItemLength = CInt32((DOUBLE)lHeight / ((((DOUBLE)mp_lMax - (DOUBLE)mp_lMin) / (DOUBLE)mp_lLargeChange) + 1));
	}
	if (lItemLength < 8)
	{
		lItemLength = 8;
	}
	lItemLength = lItemLength + 1;
	lHeight = lHeight - lItemLength;
	ButtonX1 = 1;
	ButtonX2 = ButtonX1 + Width - 1;
	ButtonY1 = CInt32(((((DOUBLE)mp_lValue - (DOUBLE)mp_lMin) / ((DOUBLE)mp_lMax - (DOUBLE)mp_lMin)) * (DOUBLE)lHeight) + 17);
	ButtonY2 = ButtonY1 + lItemLength;
}

LONG clsVScrollBarTemplate::InvCalculateV(LONG Y)
{
	LONG lHeight;
	LONG lItemLength = 0;
	LONG iReturnValue;
	lHeight = Height - (16 * 2) - 1;
	if (mp_lLargeChange > 0)
	{
		lItemLength = CInt32((DOUBLE)lHeight / ((((DOUBLE)mp_lMax - (DOUBLE)mp_lMin) / (DOUBLE)mp_lLargeChange) + 1));
	}
	if (lItemLength < 8)
	{
		lItemLength = 8;
	}
	lItemLength = lItemLength + 1;
	lHeight = lHeight - lItemLength;
	iReturnValue = CInt32(((((DOUBLE)Y - 17) * ((DOUBLE)mp_lMax - (DOUBLE)mp_lMin)) / (DOUBLE)lHeight) + (DOUBLE)mp_lMin);
	return iReturnValue;
}

CString clsVScrollBarTemplate::GetStyleIndex(void)
{
	if (mp_sStyleIndex == "DS_SCROLLBAR") 
	{
		return "";
	}
	else 
	{
		return mp_sStyleIndex;
	}
}

void clsVScrollBarTemplate::SetStyleIndex(CString Value)
{
	Value = Value.Trim();
	if (Value.GetLength() == 0) {Value = "DS_SCROLLBAR";} 
	mp_sStyleIndex = Value;
	mp_oStyle = mp_oControl->GetStyles()->FItem(Value);
}

clsStyle* clsVScrollBarTemplate::GetStyle(void)
{
	return mp_oStyle;
}

clsButtonState* clsVScrollBarTemplate::GetArrowButtons(void)
{
	return mp_oArrowButtons;
}

clsButtonState* clsVScrollBarTemplate::GetThumbButton(void)
{
	return mp_oThumbButton;
}

CString clsVScrollBarTemplate::GetXML(void)
{
	clsXML oXML(mp_oControl, "ScrollBar");
	oXML.InitializeWriter();
	oXML.WriteProperty("StyleIndex", mp_sStyleIndex);
	oXML.WriteObject(mp_oArrowButtons->GetXML());
	oXML.WriteObject(mp_oThumbButton->GetXML());
	return oXML.GetXML();
}

void clsVScrollBarTemplate::SetXML(CString sXML)
{
	clsXML oXML(mp_oControl, "ScrollBar");
	oXML.SetXML(sXML);
	oXML.InitializeReader();
	oXML.ReadProperty("StyleIndex", mp_sStyleIndex);
	SetStyleIndex(mp_sStyleIndex);
	mp_oArrowButtons->SetXML(oXML.ReadObject("ArrowButtonState"));
	mp_oThumbButton->SetXML(oXML.ReadObject("ThumbButtonState"));
}