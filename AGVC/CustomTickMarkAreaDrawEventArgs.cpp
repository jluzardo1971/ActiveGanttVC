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
#include "CustomTickMarkAreaDrawEventArgs.h"

IMPLEMENT_DYNCREATE(CustomTickMarkAreaDrawEventArgs, CCmdTarget)

//{E96E6EAB-84A6-442F-8964-B5C817A03DC9}
static const IID IID_ICustomTickMarkAreaDrawEventArgs = { 0xE96E6EAB, 0x84A6, 0x442F, { 0x89, 0x64, 0xB5, 0xC8, 0x17, 0xA0, 0x3D, 0xC9} };

//{2F041DC3-43F5-4D37-91E1-40AD2ABD92E3}
IMPLEMENT_OLECREATE_FLAGS(CustomTickMarkAreaDrawEventArgs, "AGVC.CustomTickMarkAreaDrawEventArgs", afxRegApartmentThreading, 0x2F041DC3, 0x43F5, 0x4D37, 0x91, 0xE1, 0x40, 0xAD, 0x2A, 0xBD, 0x92, 0xE3)

BEGIN_DISPATCH_MAP(CustomTickMarkAreaDrawEventArgs, CCmdTarget)
	DISP_PROPERTY_EX_ID(CustomTickMarkAreaDrawEventArgs, "CustomDraw", 1, odl_GetCustomDraw, odl_SetCustomDraw, VT_BOOL)
	DISP_PROPERTY_EX_ID(CustomTickMarkAreaDrawEventArgs, "Graphics", 2, odl_GetGraphics, SetNotSupported, VT_DISPATCH)
	DISP_PROPERTY_EX_ID(CustomTickMarkAreaDrawEventArgs, "Left", 3, odl_GetLeft, SetNotSupported, VT_I4) 
	DISP_PROPERTY_EX_ID(CustomTickMarkAreaDrawEventArgs, "Top", 4, odl_GetTop, SetNotSupported, VT_I4) 
	DISP_PROPERTY_EX_ID(CustomTickMarkAreaDrawEventArgs, "Right", 5, odl_GetRight, SetNotSupported, VT_I4) 
	DISP_PROPERTY_EX_ID(CustomTickMarkAreaDrawEventArgs, "Bottom", 6, odl_GetBottom, SetNotSupported, VT_I4)
	DISP_PROPERTY_EX_ID(CustomTickMarkAreaDrawEventArgs, "dtDate", 7, odl_GetdtDate, SetNotSupported, VT_DATE) 
	DISP_PROPERTY_EX_ID(CustomTickMarkAreaDrawEventArgs, "oTickMark", 8, odl_GetoTickMark, SetNotSupported, VT_DISPATCH)
	DISP_PROPERTY_EX_ID(CustomTickMarkAreaDrawEventArgs, "X", 9, odl_GetX, SetNotSupported, VT_I4)
END_DISPATCH_MAP()

BEGIN_INTERFACE_MAP(CustomTickMarkAreaDrawEventArgs, CCmdTarget)
	INTERFACE_PART(CustomTickMarkAreaDrawEventArgs, IID_ICustomTickMarkAreaDrawEventArgs, Dispatch)
END_INTERFACE_MAP()

BEGIN_MESSAGE_MAP(CustomTickMarkAreaDrawEventArgs, CCmdTarget)
END_MESSAGE_MAP()

CustomTickMarkAreaDrawEventArgs::CustomTickMarkAreaDrawEventArgs(void)
{
	EnableAutomation();
	AfxOleLockApp();
	Clear();
}

CustomTickMarkAreaDrawEventArgs::~CustomTickMarkAreaDrawEventArgs()
{
	AfxOleUnlockApp();
}

void CustomTickMarkAreaDrawEventArgs::OnFinalRelease()
{
	CCmdTarget::OnFinalRelease();
}

void CustomTickMarkAreaDrawEventArgs::Clear(void)
{
	mp_bCustomDraw = FALSE;
	mp_dtDate = (DATE)0;
	mp_lLeft = 0;
	mp_lTop = 0;
	mp_lRight = 0;
	mp_lBottom = 0;
	mp_oTickMark = NULL;
	mp_lX = 0;
}

BOOL CustomTickMarkAreaDrawEventArgs::GetCustomDraw(void)
{
	return mp_bCustomDraw;
}

void CustomTickMarkAreaDrawEventArgs::SetCustomDraw(BOOL value)
{
	mp_bCustomDraw = value;
}

void CustomTickMarkAreaDrawEventArgs::SetLeft(LONG value)
{
	mp_lLeft = value;
}

void CustomTickMarkAreaDrawEventArgs::SetTop(LONG value)
{
	mp_lTop = value;
}

void CustomTickMarkAreaDrawEventArgs::SetRight(LONG value)
{
	mp_lRight = value;
}

void CustomTickMarkAreaDrawEventArgs::SetBottom(LONG value)
{
	mp_lBottom = value;
}

void CustomTickMarkAreaDrawEventArgs::SetoTickMark(clsTickMark* value)
{
	mp_oTickMark = value;
}

void CustomTickMarkAreaDrawEventArgs::SetX(LONG value)
{
	mp_lX = value;
}

void CustomTickMarkAreaDrawEventArgs::SetdtDate(COleDateTime value)
{
	mp_dtDate = value;
}