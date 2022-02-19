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
#include "KeyEventArgs.h"

IMPLEMENT_DYNCREATE(KeyEventArgs, CCmdTarget)

//{88FBA620-8A17-41B4-8F9A-4F88C2C91D62}
static const IID IID_IKeyEventArgs = { 0x88FBA620, 0x8A17, 0x41B4, { 0x8F, 0x9A, 0x4F, 0x88, 0xC2, 0xC9, 0x1D, 0x62} };

//{B1D465D7-89DE-405D-A1B3-1FCCD6B479DB}
IMPLEMENT_OLECREATE_FLAGS(KeyEventArgs, "AGVC.KeyEventArgs", afxRegApartmentThreading, 0xB1D465D7, 0x89DE, 0x405D, 0xA1, 0xB3, 0x1F, 0xCC, 0xD6, 0xB4, 0x79, 0xDB)

BEGIN_DISPATCH_MAP(KeyEventArgs, CCmdTarget)
	DISP_PROPERTY_EX_ID(KeyEventArgs, "KeyCode", 1, odl_GetKeyCode, odl_SetKeyCode, VT_I2) 
	DISP_PROPERTY_EX_ID(KeyEventArgs, "Cancel", 2, odl_GetCancel, odl_SetCancel, VT_BOOL) 
	DISP_PROPERTY_EX_ID(KeyEventArgs, "CharacterCode", 3, odl_GetCharacterCode, odl_SetCharacterCode, VT_I2)
	DISP_PROPERTY_EX_ID(KeyEventArgs, "Alt", 4, odl_GetAlt, SetNotSupported, VT_BOOL)
	DISP_PROPERTY_EX_ID(KeyEventArgs, "Shift", 5, odl_GetShift, SetNotSupported, VT_BOOL)
	DISP_PROPERTY_EX_ID(KeyEventArgs, "Control", 6, odl_GetControl, SetNotSupported, VT_BOOL) 
END_DISPATCH_MAP()

BEGIN_INTERFACE_MAP(KeyEventArgs, CCmdTarget)
	INTERFACE_PART(KeyEventArgs, IID_IKeyEventArgs, Dispatch)
END_INTERFACE_MAP()

BEGIN_MESSAGE_MAP(KeyEventArgs, CCmdTarget)
END_MESSAGE_MAP()

KeyEventArgs::KeyEventArgs(void)
{
	EnableAutomation();
	AfxOleLockApp();
	Clear();
}

KeyEventArgs::~KeyEventArgs()
{
	AfxOleUnlockApp();
}

void KeyEventArgs::OnFinalRelease()
{
	CCmdTarget::OnFinalRelease();
}

SHORT KeyEventArgs::GetKeyCode(void)
{
	return mp_yKeyCode;
}

void KeyEventArgs::SetKeyCode(SHORT Value)
{
	mp_yKeyCode = Value;
}

BOOL KeyEventArgs::GetCancel(void)
{
	return mp_bCancel;
}

void KeyEventArgs::SetCancel(BOOL Value)
{
	mp_bCancel = Value;
}

SHORT KeyEventArgs::GetCharacterCode(void)
{
	return mp_yCharacterCode;
}

void KeyEventArgs::SetCharacterCode(SHORT Value)
{
	mp_yCharacterCode = Value;
}

void KeyEventArgs::Clear(void)
{
	mp_yKeyCode = 0;
	mp_yCharacterCode = 0;
	mp_bCancel = FALSE;
	BOOL mp_bAlt = FALSE;
	BOOL mp_bControl = FALSE;
	BOOL mp_bShift = FALSE;
}

BOOL KeyEventArgs::GetAlt(void)
{
	return mp_bAlt;
}

void KeyEventArgs::SetAlt(BOOL Value)
{
	mp_bAlt = Value;
}

BOOL KeyEventArgs::GetShift(void)
{
	return mp_bShift;
}

void KeyEventArgs::SetShift(BOOL Value)
{
	mp_bShift = Value;
}

BOOL KeyEventArgs::GetControl(void)
{
	return mp_bControl;
}

void KeyEventArgs::SetControl(BOOL Value)
{
	mp_bControl = Value;
}