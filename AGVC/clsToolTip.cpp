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
#include "clsToolTip.h"



IMPLEMENT_DYNCREATE(clsToolTip, CCmdTarget)

//{516BC300-A07D-40FB-94F8-6FFF191683E8}
static const IID IID_IclsToolTip = { 0x516BC300, 0xA07D, 0x40FB, { 0x94, 0xF8, 0x6F, 0xFF, 0x19, 0x16, 0x83, 0xE8} };

//{CEA1B67A-F369-4DFD-87FD-046AB2FAAA46}
IMPLEMENT_OLECREATE_FLAGS(clsToolTip, "AGVC.clsToolTip", afxRegApartmentThreading, 0xCEA1B67A, 0xF369, 0x4DFD, 0x87, 0xFD, 0x04, 0x6A, 0xB2, 0xFA, 0xAA, 0x46)

BEGIN_DISPATCH_MAP(clsToolTip, CCmdTarget)
	DISP_PROPERTY_EX_ID(clsToolTip, "Font", 1, odl_GetFont, SetNotSupported, VT_DISPATCH) 
	DISP_PROPERTY_EX_ID(clsToolTip, "BackColor", 2, odl_GetBackColor, odl_SetBackColor, VT_COLOR) 
	DISP_PROPERTY_EX_ID(clsToolTip, "ForeColor", 3, odl_GetForeColor, odl_SetForeColor, VT_COLOR) 
	DISP_PROPERTY_EX_ID(clsToolTip, "BorderColor", 4, odl_GetBorderColor, odl_SetBorderColor, VT_COLOR) 
	DISP_PROPERTY_EX_ID(clsToolTip, "Text", 5, odl_GetText, odl_SetText, VT_BSTR) 
	DISP_PROPERTY_EX_ID(clsToolTip, "AutomaticSizing", 6, odl_GetAutomaticSizing, odl_SetAutomaticSizing, VT_BOOL) 
	DISP_PROPERTY_EX_ID(clsToolTip, "Left", 7, odl_GetLeft, odl_SetLeft, VT_I4) 
	DISP_PROPERTY_EX_ID(clsToolTip, "Right", 8, odl_GetRight, SetNotSupported, VT_I4) 
	DISP_PROPERTY_EX_ID(clsToolTip, "Top", 9, odl_GetTop, odl_SetTop, VT_I4) 
	DISP_PROPERTY_EX_ID(clsToolTip, "Bottom", 10, odl_GetBottom, SetNotSupported, VT_I4) 
	DISP_PROPERTY_EX_ID(clsToolTip, "Width", 11, odl_GetWidth, odl_SetWidth, VT_I4) 
	DISP_PROPERTY_EX_ID(clsToolTip, "Height", 12, odl_GetHeight, odl_SetHeight, VT_I4) 
	DISP_PROPERTY_EX_ID(clsToolTip, "Visible", 13, odl_GetVisible, odl_SetVisible, VT_BOOL) 
END_DISPATCH_MAP()

BEGIN_INTERFACE_MAP(clsToolTip, CCmdTarget)
	INTERFACE_PART(clsToolTip, IID_IclsToolTip, Dispatch)
END_INTERFACE_MAP()

BEGIN_MESSAGE_MAP(clsToolTip, CCmdTarget)
END_MESSAGE_MAP()

clsToolTip::clsToolTip()
{
	AfxMessageBox(_T("Invalid constructor has been invoked for clsToolTip. Please inform support@sourcecodestore.com of this error"), MB_OK | MB_ICONSTOP | MB_APPLMODAL);
}

clsToolTip::clsToolTip(CActiveGanttVCCtl* oControl)
{
	EnableAutomation();
	AfxOleLockApp();
	mp_oControl = oControl;
	mp_clrBackColor = Color::LightYellow;
	mp_clrForeColor = Color::Black;
	mp_clrBorderColor = Color::Black;
	mp_lLeft = 0;
	mp_lTop = 0;
	mp_lWidth = 0;
	mp_lHeight = 0;
	mp_lX = 0;
	mp_lY = 0;
	mp_oControl->GetConfiguration()->GetDefaultFont()->Clone(&mp_oFont);
}

clsToolTip::~clsToolTip()
{
	AfxOleUnlockApp();
}

void clsToolTip::OnFinalRelease()
{
	CCmdTarget::OnFinalRelease();
}

clsFont* clsToolTip::GetFont(void)
{
	return &mp_oFont;
}

Gdiplus::Color clsToolTip::GetBackColor(void)
{
	return mp_clrBackColor;
}

void clsToolTip::SetBackColor(Gdiplus::Color Value)
{
	mp_clrBackColor = Value;
}

Gdiplus::Color clsToolTip::GetForeColor(void)
{
	return mp_clrForeColor;
}

void clsToolTip::SetForeColor(Gdiplus::Color Value)
{
	mp_clrForeColor = Value;
}

Gdiplus::Color clsToolTip::GetBorderColor(void)
{
	return mp_clrBorderColor;
}

void clsToolTip::SetBorderColor(Gdiplus::Color Value)
{
	mp_clrBorderColor = Value;
}

CString clsToolTip::GetText(void)
{
	return mp_sText;
}

void clsToolTip::SetText(CString Value)
{
	mp_sText = Value;
	if (mp_bAutomaticSizing == TRUE)
	{
		HDC lHdc;
		lHdc = GetDC(mp_oControl->f_Hwnd());
		Gdiplus::Graphics* lControlGraphics = NULL;
		lControlGraphics = new Gdiplus::Graphics(lHdc);
		Gdiplus::RectF layoutRect(0, 0, 0, 0);
		Gdiplus::RectF boundRect;
		lControlGraphics->MeasureString(mp_sText, mp_sText.GetLength(), mp_oFont.GetFontP(), layoutRect, &boundRect);
		mp_lWidth = CInt32(boundRect.GetRight() - boundRect.GetLeft());
		mp_lHeight = CInt32(boundRect.GetBottom() - boundRect.GetTop());
		ReleaseDC(mp_oControl->f_Hwnd(), lHdc);
		delete lControlGraphics;
		lControlGraphics = NULL;
	}
}

BOOL clsToolTip::GetAutomaticSizing(void)
{
	return mp_bAutomaticSizing;
}

void clsToolTip::SetAutomaticSizing(BOOL Value)
{
	mp_bAutomaticSizing = Value;
}

LONG clsToolTip::GetLeft(void)
{
	return mp_lLeft;
}

void clsToolTip::SetLeft(LONG Value)
{
	mp_lLeft = Value;
}

LONG clsToolTip::GetRight(void)
{
	return mp_lLeft + mp_lWidth;
}

LONG clsToolTip::GetTop(void)
{
	return mp_lTop;
}

void clsToolTip::SetTop(LONG Value)
{
	mp_lTop = Value;
}

LONG clsToolTip::GetBottom(void)
{
	return mp_lTop + mp_lHeight;
}

LONG clsToolTip::GetWidth(void)
{
	return mp_lWidth;
}

void clsToolTip::SetWidth(LONG Value)
{
	mp_lWidth = Value;
}

LONG clsToolTip::GetHeight(void)
{
	return mp_lHeight;
}

void clsToolTip::SetHeight(LONG Value)
{
	mp_lHeight = Value;
}

BOOL clsToolTip::GetVisible(void)
{
	return mp_bVisible;
}

void clsToolTip::SetVisible(BOOL Value)
{

	mp_bVisible = Value;
	if ((mp_bBackupDCActive == TRUE)) 
	{
		RestoreBackupDC();
	}
	if ((mp_lWidth == 0 || mp_lHeight == 0)) 
	{
		mp_bVisible = FALSE;
	}
	if ((mp_bVisible == TRUE)) 
	{
		if (mp_oControl->GetToolTipEventArgs()->GetX() == mp_lX && mp_oControl->GetToolTipEventArgs()->GetY() == mp_lY)
		{
			return;
		}
		HDC lControlHDC;
		lControlHDC = GetDC(mp_oControl->f_Hwnd());
		Gdiplus::Graphics oGraphics(lControlHDC);
		Gdiplus::Bitmap oToolTipBitmap(mp_lWidth,mp_lHeight, PixelFormat32bppARGB);
		mp_oToolTipGraphics = new Gdiplus::Graphics(&oToolTipBitmap);
		Gdiplus::Pen oPen(Color::Black);
		Gdiplus::SolidBrush oBrush(mp_clrBackColor);
		mp_lBackupLeft = mp_lLeft;
		mp_lBackupTop = mp_lTop;
		mp_lBackupRight = mp_lLeft + mp_lWidth;
		mp_lBackupBottom = mp_lTop + mp_lHeight;
		mp_oToolTipGraphics->FillRectangle(&oBrush, 0, 0, mp_lWidth, mp_lHeight);
		mp_oToolTipGraphics->DrawLine(&oPen, 0, 0, mp_lWidth - 1, 0);
		mp_oToolTipGraphics->DrawLine(&oPen, mp_lWidth - 1, 0, mp_lWidth - 1, mp_lHeight - 1);
		mp_oToolTipGraphics->DrawLine(&oPen, 0, 0, 0, mp_lHeight - 1);
		mp_oToolTipGraphics->DrawLine(&oPen, 0, mp_lHeight - 1, mp_lWidth - 1, mp_lHeight - 1);
		mp_oControl->clsG->bToolTipGraphics = TRUE;
		//mp_oControl->clsG->ClipRegion(mp_lLeft, mp_lTop, GetRight() - 1, GetBottom() - 1, FALSE);
		mp_lX = mp_oControl->GetToolTipEventArgs()->GetX();
		mp_lY = mp_oControl->GetToolTipEventArgs()->GetY();
		switch (mp_oControl->GetToolTipEventArgs()->GetToolTipType()) 
		{
		case TPT_HOVER:
			mp_oControl->GetToolTipEventArgs()->SetCustomDraw(FALSE);
			mp_oControl->FireOnMouseHoverToolTipDraw(mp_oControl->GetToolTipEventArgs()->GetEventTarget());
			break;
		case TPT_MOVEMENT:
			mp_oControl->GetToolTipEventArgs()->SetCustomDraw(FALSE);
			mp_oControl->FireOnMouseMoveToolTipDraw(mp_oControl->GetToolTipEventArgs()->GetOperation());
			break;
		}
		//mp_oControl->clsG->ClearClipRegion();
		mp_oControl->clsG->bToolTipGraphics = FALSE;
		if (mp_oControl->GetToolTipEventArgs()->GetCustomDraw() == FALSE) 
		{
			Gdiplus::SolidBrush oForeBrush(mp_clrForeColor);
			//PointF origin((REAL)mp_lLeft, (REAL)mp_lTop);
			PointF origin(0, 0);
			mp_oToolTipGraphics->DrawString(mp_sText, mp_sText.GetLength(), mp_oFont.GetFontP(), origin, &oForeBrush);
		}
		oGraphics.DrawImage(&oToolTipBitmap,mp_lLeft,mp_lTop,mp_lWidth,mp_lHeight);
		ReleaseDC(mp_oControl->f_Hwnd(), lControlHDC);
		
		delete mp_oToolTipGraphics;
		mp_oToolTipGraphics = NULL;
		mp_bBackupDCActive = TRUE;
		
	}
}

void clsToolTip::RestoreBackupDC(void)
{
	HDC lControlHDC;
	lControlHDC = GetDC(mp_oControl->f_Hwnd());
	Gdiplus::Graphics lControlGraphics(lControlHDC);
	if (mp_bVisible == TRUE) 
	{
		Gdiplus::Rect oRect(0, 0, 0, 0);
		oRect.X = mp_lLeft;
		oRect.Y = mp_lTop;
		oRect.Width = mp_lWidth;
		oRect.Height = mp_lHeight;
		lControlGraphics.ExcludeClip(oRect);
	}
	lControlGraphics.DrawImage(mp_oControl->GetBitmap(), 0, 0);
	mp_bBackupDCActive = FALSE;
}