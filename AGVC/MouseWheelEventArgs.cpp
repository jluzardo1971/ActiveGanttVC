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
#include "MouseWheelEventArgs.h"

IMPLEMENT_DYNCREATE(MouseWheelEventArgs, CCmdTarget)

//{FF6C5F85-5D43-4AA9-8643-6F918560F7D5}
static const IID IID_IMouseWheelEventArgs = { 0xFF6C5F85, 0x5D43, 0x4AA9, { 0x86, 0x43, 0x6F, 0x91, 0x85, 0x60, 0xF7, 0xD5} };

//{3B5A6B8F-CA6A-4E07-9FDB-B1DD06B68592}
IMPLEMENT_OLECREATE_FLAGS(MouseWheelEventArgs, "AGVC.MouseWheelEventArgs", afxRegApartmentThreading, 0x3B5A6B8F, 0xCA6A, 0x4E07, 0x9F, 0xDB, 0xB1, 0xDD, 0x06, 0xB6, 0x85, 0x92)

BEGIN_DISPATCH_MAP(MouseWheelEventArgs, CCmdTarget)
	DISP_PROPERTY_EX_ID(MouseWheelEventArgs, "X", 1, odl_GetX, odl_SetX, VT_I4) 
	DISP_PROPERTY_EX_ID(MouseWheelEventArgs, "Y", 2, odl_GetY, odl_SetY, VT_I4) 
	DISP_PROPERTY_EX_ID(MouseWheelEventArgs, "Button", 3, odl_GetButton, odl_SetButton, VT_I4)
	DISP_PROPERTY_EX_ID(MouseWheelEventArgs, "Delta", 4, odl_GetDelta, SetNotSupported, VT_I4)
END_DISPATCH_MAP()

BEGIN_INTERFACE_MAP(MouseWheelEventArgs, CCmdTarget)
	INTERFACE_PART(MouseWheelEventArgs, IID_IMouseWheelEventArgs, Dispatch)
END_INTERFACE_MAP()

BEGIN_MESSAGE_MAP(MouseWheelEventArgs, CCmdTarget)
END_MESSAGE_MAP()

MouseWheelEventArgs::MouseWheelEventArgs(void)
{
	EnableAutomation();
	AfxOleLockApp();
	Clear();
}

MouseWheelEventArgs::~MouseWheelEventArgs()
{
	AfxOleUnlockApp();
}

void MouseWheelEventArgs::OnFinalRelease()
{
	CCmdTarget::OnFinalRelease();
}

LONG MouseWheelEventArgs::GetX(void)
{
	return mp_lX;
}

void MouseWheelEventArgs::SetX(LONG Value)
{
	mp_lX = Value;
}

LONG MouseWheelEventArgs::GetY(void)
{
	return mp_lY;
}

void MouseWheelEventArgs::SetY(LONG Value)
{
	mp_lY = Value;
}

LONG MouseWheelEventArgs::GetButton(void)
{
	return mp_yButton;
}

void MouseWheelEventArgs::SetButton(LONG Value)
{
	mp_yButton = Value;
}

LONG MouseWheelEventArgs::GetDelta(void)
{
	return mp_lDelta;
}

void MouseWheelEventArgs::SetDelta(LONG Value)
{
	mp_lDelta = Value;
}

void MouseWheelEventArgs::Clear(void)
{
	mp_lX = 0;
	mp_lY = 0;
	mp_yButton = 0;
	mp_lDelta = 0;
}