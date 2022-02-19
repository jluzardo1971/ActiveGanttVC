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
#include "MouseEventArgs.h"

IMPLEMENT_DYNCREATE(MouseEventArgs, CCmdTarget)

//{0840CFEE-B2ED-4C2A-8BBD-B6521B17FDEC}
static const IID IID_IMouseEventArgs = { 0x0840CFEE, 0xB2ED, 0x4C2A, { 0x8B, 0xBD, 0xB6, 0x52, 0x1B, 0x17, 0xFD, 0xEC} };

//{7D374166-8B39-4ADC-BB4A-9D2493790313}
IMPLEMENT_OLECREATE_FLAGS(MouseEventArgs, "AGVC.MouseEventArgs", afxRegApartmentThreading, 0x7D374166, 0x8B39, 0x4ADC, 0xBB, 0x4A, 0x9D, 0x24, 0x93, 0x79, 0x03, 0x13)

BEGIN_DISPATCH_MAP(MouseEventArgs, CCmdTarget)
	DISP_PROPERTY_EX_ID(MouseEventArgs, "X", 1, odl_GetX, odl_SetX, VT_I4) 
	DISP_PROPERTY_EX_ID(MouseEventArgs, "Y", 2, odl_GetY, odl_SetY, VT_I4) 
	DISP_PROPERTY_EX_ID(MouseEventArgs, "EventTarget", 3, odl_GetEventTarget, odl_SetEventTarget, VT_I4) 
	DISP_PROPERTY_EX_ID(MouseEventArgs, "Operation", 4, odl_GetOperation, odl_SetOperation, VT_I4) 
	DISP_PROPERTY_EX_ID(MouseEventArgs, "Button", 5, odl_GetButton, odl_SetButton, VT_I4) 
	DISP_PROPERTY_EX_ID(MouseEventArgs, "Cancel", 6, odl_GetCancel, odl_SetCancel, VT_BOOL) 
END_DISPATCH_MAP()

BEGIN_INTERFACE_MAP(MouseEventArgs, CCmdTarget)
	INTERFACE_PART(MouseEventArgs, IID_IMouseEventArgs, Dispatch)
END_INTERFACE_MAP()

BEGIN_MESSAGE_MAP(MouseEventArgs, CCmdTarget)
END_MESSAGE_MAP()

MouseEventArgs::MouseEventArgs(void)
{
	EnableAutomation();
	AfxOleLockApp();
	Clear();
}

MouseEventArgs::~MouseEventArgs()
{
	AfxOleUnlockApp();
}

void MouseEventArgs::OnFinalRelease()
{
	CCmdTarget::OnFinalRelease();
}

LONG MouseEventArgs::GetX(void)
{
	return mp_lX;
}

void MouseEventArgs::SetX(LONG Value)
{
	mp_lX = Value;
}

LONG MouseEventArgs::GetY(void)
{
	return mp_lY;
}

void MouseEventArgs::SetY(LONG Value)
{
	mp_lY = Value;
}

LONG MouseEventArgs::GetEventTarget(void)
{
	return mp_yEventTarget;
}

void MouseEventArgs::SetEventTarget(LONG Value)
{
	mp_yEventTarget = Value;
}

LONG MouseEventArgs::GetOperation(void)
{
	return mp_yOperation;
}

void MouseEventArgs::SetOperation(LONG Value)
{
	mp_yOperation = Value;
}

LONG MouseEventArgs::GetButton(void)
{
	return mp_yButton;
}

void MouseEventArgs::SetButton(LONG Value)
{
	mp_yButton = Value;
}

BOOL MouseEventArgs::GetCancel(void)
{
	return mp_bCancel;
}

void MouseEventArgs::SetCancel(BOOL Value)
{
	mp_bCancel = Value;
}

void MouseEventArgs::Clear(void)
{
	mp_lX = 0;
	mp_lY = 0;
	mp_yEventTarget = 0;
	mp_yOperation = 0;
	mp_yButton = 0;
	mp_bCancel = FALSE;
}