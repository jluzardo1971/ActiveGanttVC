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
#include "CustomTierDrawEventArgs.h"

IMPLEMENT_DYNCREATE(CustomTierDrawEventArgs, CCmdTarget)

//{646AA6AF-316C-4812-BDF1-2FA3BE0F3D19}
static const IID IID_ICustomTierDrawEventArgs = { 0x646AA6AF, 0x316C, 0x4812, { 0xBD, 0xF1, 0x2F, 0xA3, 0xBE, 0x0F, 0x3D, 0x19} };

//{3A149262-BDB0-4A85-AC3A-E449928EAC80}
IMPLEMENT_OLECREATE_FLAGS(CustomTierDrawEventArgs, "AGVC.CustomTierDrawEventArgs", afxRegApartmentThreading, 0x3A149262, 0xBDB0, 0x4A85, 0xAC, 0x3A, 0xE4, 0x49, 0x92, 0x8E, 0xAC, 0x80)

BEGIN_DISPATCH_MAP(CustomTierDrawEventArgs, CCmdTarget)
	DISP_PROPERTY_EX_ID(CustomTierDrawEventArgs, "Text", 1, odl_GetText, odl_SetText, VT_BSTR) 
	DISP_PROPERTY_EX_ID(CustomTierDrawEventArgs, "CustomDraw", 2, odl_GetCustomDraw, odl_SetCustomDraw, VT_BOOL) 
	DISP_PROPERTY_EX_ID(CustomTierDrawEventArgs, "StyleIndex", 3, odl_GetStyleIndex, odl_SetStyleIndex, VT_BSTR) 
	DISP_PROPERTY_EX_ID(CustomTierDrawEventArgs, "TierPosition", 4, odl_GetTierPosition, odl_SetTierPosition, VT_I4) 
	DISP_PROPERTY_EX_ID(CustomTierDrawEventArgs, "StartDate", 5, odl_GetStartDate, odl_SetStartDate, VT_DATE) 
	DISP_PROPERTY_EX_ID(CustomTierDrawEventArgs, "EndDate", 6, odl_GetEndDate, odl_SetEndDate, VT_DATE) 
	DISP_PROPERTY_EX_ID(CustomTierDrawEventArgs, "Left", 7, odl_GetLeft, odl_SetLeft, VT_I4) 
	DISP_PROPERTY_EX_ID(CustomTierDrawEventArgs, "Top", 8, odl_GetTop, odl_SetTop, VT_I4) 
	DISP_PROPERTY_EX_ID(CustomTierDrawEventArgs, "Right", 9, odl_GetRight, odl_SetRight, VT_I4) 
	DISP_PROPERTY_EX_ID(CustomTierDrawEventArgs, "Bottom", 10, odl_GetBottom, odl_SetBottom, VT_I4) 
	DISP_PROPERTY_EX_ID(CustomTierDrawEventArgs, "LeftTrim", 11, odl_GetLeftTrim, odl_SetLeftTrim, VT_I4) 
	DISP_PROPERTY_EX_ID(CustomTierDrawEventArgs, "RightTrim", 12, odl_GetRightTrim, odl_SetRightTrim, VT_I4) 
	DISP_PROPERTY_EX_ID(CustomTierDrawEventArgs, "Graphics", 13, odl_GetGraphics, SetNotSupported, VT_DISPATCH) 
	DISP_PROPERTY_EX_ID(CustomTierDrawEventArgs, "Interval", 14, odl_GetInterval, odl_SetInterval, VT_I4)
	DISP_PROPERTY_EX_ID(CustomTierDrawEventArgs, "Factor", 15, odl_GetFactor, odl_SetFactor, VT_I4) 
END_DISPATCH_MAP()

BEGIN_INTERFACE_MAP(CustomTierDrawEventArgs, CCmdTarget)
	INTERFACE_PART(CustomTierDrawEventArgs, IID_ICustomTierDrawEventArgs, Dispatch)
END_INTERFACE_MAP()

BEGIN_MESSAGE_MAP(CustomTierDrawEventArgs, CCmdTarget)
END_MESSAGE_MAP()

CustomTierDrawEventArgs::CustomTierDrawEventArgs(void)
{
	EnableAutomation();
	AfxOleLockApp();
	Clear();
}

CustomTierDrawEventArgs::~CustomTierDrawEventArgs()
{
	AfxOleUnlockApp();
}

void CustomTierDrawEventArgs::OnFinalRelease()
{
	CCmdTarget::OnFinalRelease();
}


CString CustomTierDrawEventArgs::GetText(void)
{
	return mp_sText;
}

void CustomTierDrawEventArgs::SetText(CString Value)
{
	mp_sText = Value;
}

BOOL CustomTierDrawEventArgs::GetCustomDraw(void)
{
	return mp_bCustomDraw;
}

void CustomTierDrawEventArgs::SetCustomDraw(BOOL Value)
{
	mp_bCustomDraw = Value;
}

CString CustomTierDrawEventArgs::GetStyleIndex(void)
{
	return mp_sStyleIndex;
}

void CustomTierDrawEventArgs::SetStyleIndex(CString Value)
{
	mp_sStyleIndex = Value;
}

LONG CustomTierDrawEventArgs::GetTierPosition(void)
{
	return mp_yTierPosition;
}

void CustomTierDrawEventArgs::SetTierPosition(LONG Value)
{
	mp_yTierPosition = Value;
}

COleDateTime CustomTierDrawEventArgs::GetStartDate(void)
{
	return mp_dtStartDate;
}

void CustomTierDrawEventArgs::SetStartDate(COleDateTime Value)
{
	mp_dtStartDate = Value;
}

COleDateTime CustomTierDrawEventArgs::GetEndDate(void)
{
	return mp_dtEndDate;
}

void CustomTierDrawEventArgs::SetEndDate(COleDateTime Value)
{
	mp_dtEndDate = Value;
}

LONG CustomTierDrawEventArgs::GetLeft(void)
{
	return mp_lLeft;
}

void CustomTierDrawEventArgs::SetLeft(LONG Value)
{
	mp_lLeft = Value;
}

LONG CustomTierDrawEventArgs::GetTop(void)
{
	return mp_lTop;
}

void CustomTierDrawEventArgs::SetTop(LONG Value)
{
	mp_lTop = Value;
}

LONG CustomTierDrawEventArgs::GetRight(void)
{
	return mp_lRight;
}

void CustomTierDrawEventArgs::SetRight(LONG Value)
{
	mp_lRight = Value;
}

LONG CustomTierDrawEventArgs::GetBottom(void)
{
	return mp_lBottom;
}

void CustomTierDrawEventArgs::SetBottom(LONG Value)
{
	mp_lBottom = Value;
}

LONG CustomTierDrawEventArgs::GetLeftTrim(void)
{
	return mp_lLeftTrim;
}

void CustomTierDrawEventArgs::SetLeftTrim(LONG Value)
{
	mp_lLeftTrim = Value;
}

LONG CustomTierDrawEventArgs::GetRightTrim(void)
{
	return mp_lRightTrim;
}

void CustomTierDrawEventArgs::SetRightTrim(LONG Value)
{
	mp_lRightTrim = Value;
}

clsGDIGraphics* CustomTierDrawEventArgs::GetGraphics(void)
{
	return &mp_oGraphics;
}

LONG CustomTierDrawEventArgs::GetInterval(void)
{
	return mp_yInterval;
}

void CustomTierDrawEventArgs::SetInterval(LONG Value)
{
	mp_yInterval = Value;
}

LONG CustomTierDrawEventArgs::GetFactor(void)
{
	return mp_lFactor;
}

void CustomTierDrawEventArgs::SetFactor(LONG Value)
{
	mp_lFactor = Value;
}

void CustomTierDrawEventArgs::Clear(void)
{
	mp_sText = "";
	mp_bCustomDraw = FALSE;
	mp_sStyleIndex = "";
	mp_yTierPosition = SP_UPPER;
	mp_dtStartDate = (DATE)0;
	mp_dtEndDate = (DATE)0;
	mp_lLeft = 0;
	mp_lTop = 0;
	mp_lRight = 0;
	mp_lBottom = 0;
	mp_lLeftTrim = 0;
	mp_lRightTrim = 0;
	mp_yInterval = IL_SECOND;
	mp_lFactor = 1;
}