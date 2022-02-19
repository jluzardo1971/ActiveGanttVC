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
#include "clsTextFlags.h"

IMPLEMENT_DYNCREATE(clsTextFlags, CCmdTarget)

//{9B183408-5FF1-4B9C-8730-E61B0C18B5D9}
static const IID IID_IclsTextFlags = { 0x9B183408, 0x5FF1, 0x4B9C, { 0x87, 0x30, 0xE6, 0x1B, 0x0C, 0x18, 0xB5, 0xD9} };

//{9AAF8333-E184-4A01-A6EC-FCEE0D1D7A88}
IMPLEMENT_OLECREATE_FLAGS(clsTextFlags, "AGVC.clsTextFlags", afxRegApartmentThreading, 0x9AAF8333, 0xE184, 0x4A01, 0xA6, 0xEC, 0xFC, 0xEE, 0x0D, 0x1D, 0x7A, 0x88)

BEGIN_DISPATCH_MAP(clsTextFlags, CCmdTarget)
	DISP_PROPERTY_EX_ID(clsTextFlags, "VerticalAlignment", 1, odl_GetVerticalAlignment, odl_SetVerticalAlignment, VT_I4) 
	DISP_PROPERTY_EX_ID(clsTextFlags, "HorizontalAlignment", 2, odl_GetHorizontalAlignment, odl_SetHorizontalAlignment, VT_I4) 
	DISP_PROPERTY_EX_ID(clsTextFlags, "WordWrap", 3, odl_GetWordWrap, odl_SetWordWrap, VT_BOOL) 
	DISP_PROPERTY_EX_ID(clsTextFlags, "RightToLeft", 4, odl_GetRightToLeft, odl_SetRightToLeft, VT_BOOL) 
	DISP_PROPERTY_EX_ID(clsTextFlags, "OffsetBottom", 5, odl_GetOffsetBottom, odl_SetOffsetBottom, VT_I4) 
	DISP_PROPERTY_EX_ID(clsTextFlags, "OffsetLeft", 6, odl_GetOffsetLeft, odl_SetOffsetLeft, VT_I4) 
	DISP_PROPERTY_EX_ID(clsTextFlags, "OffsetRight", 7, odl_GetOffsetRight, odl_SetOffsetRight, VT_I4) 
	DISP_PROPERTY_EX_ID(clsTextFlags, "OffsetTop", 8, odl_GetOffsetTop, odl_SetOffsetTop, VT_I4)
	DISP_FUNCTION_ID(clsTextFlags, "GetXML", 9, odl_GetXML, VT_BSTR, VTS_NONE) 
	DISP_FUNCTION_ID(clsTextFlags, "SetXML", 10, odl_SetXML, VT_EMPTY, VTS_BSTR) 
END_DISPATCH_MAP()

BEGIN_INTERFACE_MAP(clsTextFlags, CCmdTarget)
	INTERFACE_PART(clsTextFlags, IID_IclsTextFlags, Dispatch)
END_INTERFACE_MAP()

BEGIN_MESSAGE_MAP(clsTextFlags, CCmdTarget)
END_MESSAGE_MAP()

clsTextFlags::clsTextFlags()
{
	EnableAutomation();
	AfxOleLockApp();
	mp_oControl = NULL;
	Clear();
}

clsTextFlags::clsTextFlags(CActiveGanttVCCtl* oControl)
{
	EnableAutomation();
	AfxOleLockApp();
	mp_oControl = oControl;
	Clear();
}

clsTextFlags::clsTextFlags(const clsTextFlags& oTextFlags)
{
	EnableAutomation();
	AfxOleLockApp();
	mp_yVerticalAlignment = oTextFlags.mp_yVerticalAlignment;
	mp_yHorizontalAlignment = oTextFlags.mp_yHorizontalAlignment;
	mp_bWordWrap = oTextFlags.mp_bWordWrap;
	mp_bRightToLeft = oTextFlags.mp_bRightToLeft;
	mp_lOffsetBottom = oTextFlags.mp_lOffsetBottom;
	mp_lOffsetLeft = oTextFlags.mp_lOffsetLeft;
	mp_lOffsetRight = oTextFlags.mp_lOffsetRight;
	mp_lOffsetTop = oTextFlags.mp_lOffsetTop;
}

clsTextFlags::~clsTextFlags()
{
	AfxOleUnlockApp();
}

void clsTextFlags::OnFinalRelease()
{
	CCmdTarget::OnFinalRelease();
}

LONG clsTextFlags::GetVerticalAlignment(void)
{
	return mp_yVerticalAlignment;
}

void clsTextFlags::SetVerticalAlignment(LONG Value)
{
	mp_yVerticalAlignment = Value;
}

LONG clsTextFlags::GetHorizontalAlignment(void)
{
	return mp_yHorizontalAlignment;
}

void clsTextFlags::SetHorizontalAlignment(LONG Value)
{
	mp_yHorizontalAlignment = Value;
}

BOOL clsTextFlags::GetWordWrap(void)
{
	return mp_bWordWrap;
}

void clsTextFlags::SetWordWrap(BOOL Value)
{
	mp_bWordWrap = Value;
}

BOOL clsTextFlags::GetRightToLeft(void)
{
	return mp_bRightToLeft;
}

void clsTextFlags::SetRightToLeft(BOOL Value)
{
	mp_bRightToLeft = Value;
}

LONG clsTextFlags::GetOffsetBottom(void)
{
	return mp_lOffsetBottom;
}

void clsTextFlags::SetOffsetBottom(LONG Value)
{
	mp_lOffsetBottom = Value;
}

LONG clsTextFlags::GetOffsetLeft(void)
{
	return mp_lOffsetLeft;
}

void clsTextFlags::SetOffsetLeft(LONG Value)
{
	mp_lOffsetLeft = Value;
}

LONG clsTextFlags::GetOffsetRight(void)
{
	return mp_lOffsetRight;
}

void clsTextFlags::SetOffsetRight(LONG Value)
{
	mp_lOffsetRight = Value;
}

LONG clsTextFlags::GetOffsetTop(void)
{
	return mp_lOffsetTop;
}

void clsTextFlags::SetOffsetTop(LONG Value)
{
	mp_lOffsetTop = Value;
}

Gdiplus::StringFormat* clsTextFlags::GetFlags(void)
{
	Gdiplus::StringFormat* oReturn = ::new StringFormat();
	if (mp_yVerticalAlignment == VAL_TOP) 
	{
		oReturn->SetLineAlignment(StringAlignmentNear);
	}
	if (mp_yVerticalAlignment == VAL_CENTER) 
	{
		oReturn->SetLineAlignment(StringAlignmentCenter);
	}
	if (mp_yVerticalAlignment == VAL_BOTTOM) 
	{
		oReturn->SetLineAlignment(StringAlignmentFar);
	}
	if (mp_yHorizontalAlignment == HAL_LEFT) 
	{
		oReturn->SetAlignment(StringAlignmentNear);
	}
	if (mp_yHorizontalAlignment == HAL_CENTER) 
	{
		oReturn->SetAlignment(StringAlignmentCenter);
	}
	if (mp_yHorizontalAlignment == HAL_RIGHT) 
	{
		oReturn->SetAlignment(StringAlignmentFar);
	}
	if (mp_bWordWrap == FALSE) 
	{
		oReturn->SetFormatFlags(oReturn->GetFormatFlags() | StringFormatFlagsNoWrap);
	}
	if (mp_bRightToLeft == TRUE) 
	{
		oReturn->SetFormatFlags(oReturn->GetFormatFlags() | StringFormatFlagsDirectionRightToLeft);
	}
	return oReturn;
}

CString clsTextFlags::GetXML(void)
{
	clsXML oXML(mp_oControl, "TextFlags");
	oXML.InitializeWriter();
	oXML.WriteProperty("HorizontalAlignment", mp_yHorizontalAlignment);
	oXML.WriteProperty("OffsetBottom", mp_lOffsetBottom);
	oXML.WriteProperty("OffsetLeft", mp_lOffsetLeft);
	oXML.WriteProperty("OffsetRight", mp_lOffsetRight);
	oXML.WriteProperty("OffsetTop", mp_lOffsetTop);
	oXML.WriteProperty("RightToLeft", mp_bRightToLeft);
	oXML.WriteProperty("VerticalAlignment", mp_yVerticalAlignment);
	oXML.WriteProperty("WordWrap", mp_bWordWrap);
	return oXML.GetXML();
}

void clsTextFlags::SetXML(CString sXML)
{
	clsXML oXML(mp_oControl, "TextFlags");
	oXML.SetXML(sXML);
	oXML.InitializeReader();
	oXML.ReadProperty("HorizontalAlignment", mp_yHorizontalAlignment);
	oXML.ReadProperty("OffsetBottom", mp_lOffsetBottom);
	oXML.ReadProperty("OffsetLeft", mp_lOffsetLeft);
	oXML.ReadProperty("OffsetRight", mp_lOffsetRight);
	oXML.ReadProperty("OffsetTop", mp_lOffsetTop);
	oXML.ReadProperty("RightToLeft", mp_bRightToLeft);
	oXML.ReadProperty("VerticalAlignment", mp_yVerticalAlignment);
	oXML.ReadProperty("WordWrap", mp_bWordWrap);
}



void clsTextFlags::Clear()
{
	mp_yVerticalAlignment = VAL_TOP;
	mp_yHorizontalAlignment = HAL_LEFT;
	mp_bWordWrap = 0;
	mp_bRightToLeft = 0;
	mp_lOffsetBottom = 0;
	mp_lOffsetLeft = 0;
	mp_lOffsetRight = 0;
	mp_lOffsetTop = 0;
}
void clsTextFlags::Clone(clsTextFlags* oClone)
{
    oClone->SetHorizontalAlignment(mp_yHorizontalAlignment);
    oClone->SetOffsetBottom(mp_lOffsetBottom);
    oClone->SetOffsetLeft(mp_lOffsetLeft);
    oClone->SetOffsetRight(mp_lOffsetRight);
    oClone->SetOffsetTop(mp_lOffsetTop);
    oClone->SetRightToLeft(mp_bRightToLeft);
    oClone->SetVerticalAlignment(mp_yVerticalAlignment);
    oClone->SetWordWrap(mp_bWordWrap);
}