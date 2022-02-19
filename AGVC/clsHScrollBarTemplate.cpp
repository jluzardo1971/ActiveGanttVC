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
#include "clsHScrollBarTemplate.h"



IMPLEMENT_DYNCREATE(clsHScrollBarTemplate, CCmdTarget)

//{C4E0B52F-E7AE-490B-B781-9CA83ADF9F8F}
static const IID IID_IclsHScrollBarTemplate = { 0xC4E0B52F, 0xE7AE, 0x490B, { 0xB7, 0x81, 0x9C, 0xA8, 0x3A, 0xDF, 0x9F, 0x8F} };

//{9DA609FD-ADC9-4D59-883E-D73FC7E71FF9}
IMPLEMENT_OLECREATE_FLAGS(clsHScrollBarTemplate, "AGVC.clsHScrollBarTemplate", afxRegApartmentThreading, 0x9DA609FD, 0xADC9, 0x4D59, 0x88, 0x3E, 0xD7, 0x3F, 0xC7, 0xE7, 0x1F, 0xF9)

BEGIN_DISPATCH_MAP(clsHScrollBarTemplate, CCmdTarget)
	DISP_PROPERTY_EX_ID(clsHScrollBarTemplate, "TimerInterval", 1, odl_GetTimerInterval, odl_SetTimerInterval, VT_I4)
	DISP_PROPERTY_EX_ID(clsHScrollBarTemplate, "StyleIndex", 2, odl_GetStyleIndex, odl_SetStyleIndex, VT_BSTR) 
	DISP_PROPERTY_EX_ID(clsHScrollBarTemplate, "Style", 3, odl_GetStyle, SetNotSupported, VT_DISPATCH)
	DISP_PROPERTY_EX_ID(clsHScrollBarTemplate, "ArrowButtons", 4, odl_GetArrowButtons, SetNotSupported, VT_DISPATCH) 
	DISP_PROPERTY_EX_ID(clsHScrollBarTemplate, "ThumbButton", 5, odl_GetThumbButton, SetNotSupported, VT_DISPATCH)
	DISP_FUNCTION_ID(clsHScrollBarTemplate, "GetXML", 6, odl_GetXML, VT_BSTR, VTS_NONE) 
	DISP_FUNCTION_ID(clsHScrollBarTemplate, "SetXML", 7, odl_SetXML, VT_EMPTY, VTS_BSTR)
END_DISPATCH_MAP()

BEGIN_INTERFACE_MAP(clsHScrollBarTemplate, CCmdTarget)
	INTERFACE_PART(clsHScrollBarTemplate, IID_IclsHScrollBarTemplate, Dispatch)
END_INTERFACE_MAP()

BEGIN_MESSAGE_MAP(clsHScrollBarTemplate, CCmdTarget)
END_MESSAGE_MAP()

clsHScrollBarTemplate::clsHScrollBarTemplate()
{
	AfxMessageBox(_T("Invalid constructor has been invoked for clsHScrollBarTemplate. Please inform support@sourcecodestore.com of this error"), MB_OK | MB_ICONSTOP | MB_APPLMODAL);
}


clsHScrollBarTemplate::clsHScrollBarTemplate(CActiveGanttVCCtl* oControl)
{
	EnableAutomation();
	AfxOleLockApp();
	mp_oControl = oControl;
	mp_oArrowButtons = new clsButtonState(mp_oControl, "Arrow");
	mp_oThumbButton = new clsButtonState(mp_oControl, "Thumb");
	Clear();
}

clsHScrollBarTemplate::~clsHScrollBarTemplate()
{
	delete mp_oThumbButton;
	mp_oThumbButton = NULL;
	delete mp_oArrowButtons;
	mp_oArrowButtons = NULL;
	AfxOleUnlockApp();
}

void clsHScrollBarTemplate::OnFinalRelease()
{
	CCmdTarget::OnFinalRelease();
}

LONG clsHScrollBarTemplate::GetTimerInterval(void)
{
	return mp_lTimerInterval;
}

void clsHScrollBarTemplate::SetTimerInterval(LONG Value)
{
	mp_lTimerInterval = Value;
}

BOOL clsHScrollBarTemplate::GetEnabled(void)
{
	return mp_bEnabled;
}
void clsHScrollBarTemplate::SetEnabled(BOOL Value)
{
	mp_bEnabled = Value;
}

LONG clsHScrollBarTemplate::GetValue(void)
{
	return mp_lValue;
}

void clsHScrollBarTemplate::SetValue(LONG Value)
{
	mp_lValue = Value;
	if (mp_lValue < mp_lMin)
	{
//		throw new Exception("Value is less than mp_lMin.");
	}
}

LONG clsHScrollBarTemplate::GetMin(void)
{
	return mp_lMin;
}

void clsHScrollBarTemplate::SetMin(LONG Value)
{
	mp_lMin = Value;
}

LONG clsHScrollBarTemplate::GetMax(void)
{
	return mp_lMax;
}

void clsHScrollBarTemplate::SetMax(LONG Value)
{
	mp_lMax = Value;
	if (mp_lMax < mp_lMin)
	{
//		throw new Exception("mp_lMax is less than mp_lMin.");
	}
}

LONG clsHScrollBarTemplate::GetSmallChange(void)
{
	return mp_lSmallChange;
}

void clsHScrollBarTemplate::SetSmallChange(LONG Value)
{
	mp_lSmallChange = Value;
}

LONG clsHScrollBarTemplate::GetLargeChange(void)
{
	return mp_lLargeChange;
}

void clsHScrollBarTemplate::SetLargeChange(LONG Value)
{
	mp_lLargeChange = Value;
}

LONG clsHScrollBarTemplate::GetState(void)
{
	return mp_yState;
}

void clsHScrollBarTemplate::SetState(LONG Value)
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

void clsHScrollBarTemplate::Initialize(E_SCROLLBAR yType)
{
	mp_lSmallChange = 1;
	Height = 17;
	mp_bMouseDown = FALSE;
	mp_iButton = BTN_NONE;
	mp_iTimerInterval = 100;
	mp_lValueBuff = GetValue();
	mp_yType = yType;
	mp_lValue = 0;
}

void clsHScrollBarTemplate::Draw(void)
{
	if (Visible == FALSE)
	{
		return;
	}
    clsStyle* oArrowButtonLeftStyle = GetArrowButtons()->GetNormalStyle();
    clsStyle* oArrowButtonRightStyle = GetArrowButtons()->GetNormalStyle();
    clsStyle* oThumbButtonStyle = GetThumbButton()->GetNormalStyle();
	if ((mp_lMax - mp_lMin) == 0)
	{
		return;
	}
    mp_oControl->clsG->mp_DrawItem(Left, Top, Left + Width - 1, Top + Height - 1, "", "", FALSE, NULL, 0, 0, mp_oStyle);
	CalculateH();
	if (mp_bEnabled == FALSE)
	{
        oThumbButtonStyle = GetThumbButton()->GetDisabledStyle();
        oArrowButtonLeftStyle = GetArrowButtons()->GetDisabledStyle();
        oArrowButtonRightStyle = GetArrowButtons()->GetDisabledStyle();
	}
	else if (mp_bMouseDown == TRUE)
	{
		if (mp_iButton == BTN_LEFT)
		{
            oThumbButtonStyle = GetThumbButton()->GetNormalStyle();
            oArrowButtonLeftStyle = GetArrowButtons()->GetPressedStyle();
            oArrowButtonRightStyle = GetArrowButtons()->GetNormalStyle();
		}
		else if (mp_iButton == BTN_RIGHT)
		{
            oThumbButtonStyle = GetThumbButton()->GetNormalStyle();
            oArrowButtonLeftStyle = GetArrowButtons()->GetNormalStyle();
            oArrowButtonRightStyle = GetArrowButtons()->GetPressedStyle();
		}
		else if (mp_iButton == BTN_BUTTON)
		{
            oThumbButtonStyle = GetThumbButton()->GetPressedStyle();
            oArrowButtonLeftStyle = GetArrowButtons()->GetNormalStyle();
            oArrowButtonRightStyle = GetArrowButtons()->GetNormalStyle();
		}
		else
		{
            oThumbButtonStyle = GetThumbButton()->GetNormalStyle();
            oArrowButtonLeftStyle = GetArrowButtons()->GetNormalStyle();
            oArrowButtonRightStyle = GetArrowButtons()->GetNormalStyle();
		}

	}
	else
	{
        oArrowButtonLeftStyle = GetArrowButtons()->GetNormalStyle();
        oArrowButtonRightStyle = GetArrowButtons()->GetNormalStyle();
        oThumbButtonStyle = GetThumbButton()->GetNormalStyle();
	}
    mp_oControl->clsG->mp_DrawItem(Left, Top, Left + 16, Top + Height - 1, "", "", FALSE, NULL, 0, 0, oArrowButtonLeftStyle);
    mp_oControl->clsG->mp_DrawItem(Left + Width - 17, Top, Left + Width - 1, Top + Height - 1, "", "", FALSE, NULL, 0, 0, oArrowButtonRightStyle);
    mp_oControl->clsG->mp_DrawItem(Left + ButtonX1, Top + ButtonY1 - 1, Left + ButtonX2 - 2, Top + ButtonY2 - 1, "", "", FALSE, NULL, 0, 0, oThumbButtonStyle);
    if (oArrowButtonLeftStyle->GetScrollBarStyle()->GetDropShadow() == TRUE)
    {
        mp_oControl->clsG->mp_DrawArrow(Left + oArrowButtonLeftStyle->GetScrollBarStyle()->GetDropShadowLeftX(), Top + oArrowButtonLeftStyle->GetScrollBarStyle()->GetDropShadowLeftY(), AWD_LEFT, oArrowButtonLeftStyle->GetScrollBarStyle()->GetArrowSize(), oArrowButtonLeftStyle->GetScrollBarStyle()->GetDropShadowArrowColor());
    }
    mp_oControl->clsG->mp_DrawArrow(Left + oArrowButtonLeftStyle->GetScrollBarStyle()->GetLeftX(), Top + oArrowButtonLeftStyle->GetScrollBarStyle()->GetLeftY(), AWD_LEFT, oArrowButtonLeftStyle->GetScrollBarStyle()->GetArrowSize(), oArrowButtonLeftStyle->GetScrollBarStyle()->GetArrowColor());
    if (oArrowButtonRightStyle->GetScrollBarStyle()->GetDropShadow() == TRUE)
    {
        mp_oControl->clsG->mp_DrawArrow(Left + Width - 17 + oArrowButtonRightStyle->GetScrollBarStyle()->GetDropShadowRightX(), Top + oArrowButtonRightStyle->GetScrollBarStyle()->GetDropShadowRightY(), AWD_RIGHT, oArrowButtonRightStyle->GetScrollBarStyle()->GetArrowSize(), oArrowButtonRightStyle->GetScrollBarStyle()->GetDropShadowArrowColor());
    }
    mp_oControl->clsG->mp_DrawArrow(Left + Width - 17 + oArrowButtonRightStyle->GetScrollBarStyle()->GetRightX(), Top + oArrowButtonRightStyle->GetScrollBarStyle()->GetRightY(), AWD_RIGHT, oArrowButtonRightStyle->GetScrollBarStyle()->GetArrowSize(), oArrowButtonRightStyle->GetScrollBarStyle()->GetArrowColor());
}

void clsHScrollBarTemplate::OnMouseHover(LONG X,LONG Y)
{
	mp_oControl->MouseKeyboardEvents->mp_SetCursor(CT_NORMAL);
}

void clsHScrollBarTemplate::OnMouseDown(LONG X,LONG Y)
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

void clsHScrollBarTemplate::OnMouseMove(LONG X,LONG Y)
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
	if (mp_iButton == BTN_BUTTON)
	{
		LONG mp_iProjectedValue = mp_lValue + (InvCalculateH(X) - InvCalculateH(mp_iMouseXPosition));
		if (mp_iProjectedValue >= mp_lMin && mp_iProjectedValue <= mp_lMax)
		{
			mp_iMouseXPosition = X;
			mp_lValue = mp_iProjectedValue;
			mp_ValueChanged();
		}
		else if (mp_iProjectedValue < mp_lMin)
		{
			mp_iMouseXPosition = X;
			mp_lValue = mp_lMin;
			mp_ValueChanged();
		}
		else if (mp_iProjectedValue > mp_lMax)
		{
			mp_iMouseXPosition = X;
			mp_lValue = mp_lMax;
			mp_ValueChanged();
		}
	}
}

void clsHScrollBarTemplate::OnMouseUp(void)
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

BOOL clsHScrollBarTemplate::OverControl(LONG X,LONG Y)
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

void clsHScrollBarTemplate::ConvertToRelativeCoords(LONG& X,LONG& Y)
{
	X = X - Left;
	Y = Y - Top;
}

BOOL clsHScrollBarTemplate::ScrollBarClick(LONG X,LONG Y)
{
	if (OverControl(X,Y) == FALSE)
	{
		return FALSE;
	}
	ConvertToRelativeCoords(X, Y);
	CalculateH();
	if (X < 17)
	{
		mp_iButton = BTN_LEFT;
		SmallChangeLeft();
		return TRUE;
	}
	else if (X > 17 && X < ButtonX1)
	{
		mp_iButton = BTN_LEFTLCHANGE;
		LargeChangeLeft();
		return TRUE;
	}
	else if (X >= ButtonX1 && X <= ButtonX2)
	{
		mp_iButton = BTN_BUTTON;
		mp_iMouseXPosition = X;
		return TRUE;
	}
	else if (X > ButtonY2 && X < Width - 17)
	{
		mp_iButton = BTN_RIGHTLCHANGE;
		LargeChangeRight();
		return TRUE;
	}
	else if (X > Width - 17 && X < Width)
	{
		mp_iButton = BTN_RIGHT;
		SmallChangeRight();
		return TRUE;
	}
	else
	{
		return FALSE;
	}
}

void clsHScrollBarTemplate::SmallChangeLeft(void)
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

void clsHScrollBarTemplate::SmallChangeRight(void)
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

void clsHScrollBarTemplate::LargeChangeLeft(void)
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

void clsHScrollBarTemplate::LargeChangeRight(void)
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

void clsHScrollBarTemplate::mp_ValueChanged(void)
{
	if ((mp_lValue - mp_lValueBuff) != 0)
	{
		if (mp_yType == SCR_HORIZONTAL2)
		{
			mp_oControl->TimeLineScrollBar_ValueChanged(mp_lValue - mp_lValueBuff);
		}
		else if (mp_yType == SCR_HORIZONTAL1)
		{
			mp_oControl->HorizontalScrollBar_ValueChanged(mp_lValue - mp_lValueBuff);
		}
	}
	mp_lValueBuff = mp_lValue;
}

void clsHScrollBarTemplate::mp_oTimer_Tick(void)
{
	switch (mp_iButton)
	{
	case BTN_LEFT:
		SmallChangeLeft();
		break;
	case BTN_LEFTLCHANGE:
		LargeChangeLeft();
		break;
	case BTN_RIGHTLCHANGE:
		LargeChangeRight();
		break;
	case BTN_RIGHT:
		SmallChangeRight();
		break;
	}
}

void clsHScrollBarTemplate::CalculateH(void)
{
	LONG lWidth;
	LONG lItemLength = 0;
	lWidth = Width - (16 * 2) - 1;
	if (mp_lLargeChange > 0)
	{
		lItemLength = CInt32((DOUBLE)lWidth / ((((DOUBLE)mp_lMax - (DOUBLE)mp_lMin) / (DOUBLE)mp_lLargeChange) + 1));
	}
	if (lItemLength < 8)
	{
		lItemLength = 8;
	}
	lItemLength = lItemLength + 1;
	lWidth = lWidth - lItemLength;
	ButtonX1 = CInt32(((((DOUBLE)mp_lValue - (DOUBLE)mp_lMin) / ((DOUBLE)mp_lMax - (DOUBLE)mp_lMin)) * (DOUBLE)lWidth) + 17);
	ButtonX2 = ButtonX1 + lItemLength;
	ButtonY1 = 1;
	ButtonY2 = ButtonY1 + Height - 1;
}

LONG clsHScrollBarTemplate::InvCalculateH(LONG X)
{
	LONG lWidth;
	LONG lItemLength = 0;
	LONG iReturnValue;
	lWidth = Width - (16 * 2) - 1;
	if (mp_lLargeChange > 0)
	{
		lItemLength = CInt32((DOUBLE)lWidth / ((((DOUBLE)mp_lMax - (DOUBLE)mp_lMin) / (DOUBLE)mp_lLargeChange) + 1));
	}
	if (lItemLength < 8)
	{
		lItemLength = 8;
	}
	lItemLength = lItemLength + 1;
	lWidth = lWidth - lItemLength;
	iReturnValue = CInt32(((((DOUBLE)X - 17) * ((DOUBLE)mp_lMax - (DOUBLE)mp_lMin)) / (DOUBLE)lWidth) + (DOUBLE)mp_lMin);
	return iReturnValue;
}

CString clsHScrollBarTemplate::GetStyleIndex(void)
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

void clsHScrollBarTemplate::SetStyleIndex(CString Value)
{
	Value = Value.Trim();
	if (Value.GetLength() == 0) {Value = "DS_SCROLLBAR";} 
	mp_sStyleIndex = Value;
	mp_oStyle = mp_oControl->GetStyles()->FItem(Value);
}

clsStyle* clsHScrollBarTemplate::GetStyle(void)
{
	return mp_oStyle;
}

clsButtonState* clsHScrollBarTemplate::GetArrowButtons(void)
{
	return mp_oArrowButtons;
}

clsButtonState* clsHScrollBarTemplate::GetThumbButton(void)
{
	return mp_oThumbButton;
}

CString clsHScrollBarTemplate::GetXML(void)
{
	clsXML oXML(mp_oControl, "ScrollBar");
	oXML.InitializeWriter();
	oXML.WriteProperty("StyleIndex", mp_sStyleIndex);
	oXML.WriteObject(mp_oArrowButtons->GetXML());
	oXML.WriteObject(mp_oThumbButton->GetXML());
	return oXML.GetXML();
}

void clsHScrollBarTemplate::SetXML(CString sXML)
{
	clsXML oXML(mp_oControl, "ScrollBar");
	oXML.SetXML(sXML);
	oXML.InitializeReader();
	oXML.ReadProperty("StyleIndex", mp_sStyleIndex);
	SetStyleIndex(mp_sStyleIndex);
	mp_oArrowButtons->SetXML(oXML.ReadObject("ArrowButtonState"));
	mp_oThumbButton->SetXML(oXML.ReadObject("ThumbButtonState"));
}

void clsHScrollBarTemplate::Clear(void)
{
	mp_lSmallChange = 0;
	mp_lLargeChange = 0;
	mp_lValue = 0;
	mp_lMin = 0;
	mp_lMax = 0;
	mp_sStyleIndex = "DS_SCROLLBAR";
	mp_oStyle = mp_oControl->GetStyles()->FItem("DS_SCROLLBAR");
}

void clsHScrollBarTemplate::Clone(clsHScrollBarTemplate* oClone)
{
    oClone->mp_lSmallChange = mp_lSmallChange;
    oClone->mp_lLargeChange = mp_lLargeChange;
    oClone->mp_lMin = mp_lMin;
    oClone->mp_lMax = mp_lMax;
    oClone->mp_lValue = mp_lValue;
    oClone->Height = Height;
    oClone->mp_bMouseDown = mp_bMouseDown;
    oClone->mp_iButton = mp_iButton;
    oClone->mp_iTimerInterval = mp_iTimerInterval;
    oClone->mp_lValueBuff = mp_lValueBuff;
    oClone->SetStyleIndex(mp_sStyleIndex);
    GetArrowButtons()->Clone(oClone->GetArrowButtons());
    GetThumbButton()->Clone(oClone->GetThumbButton());
}