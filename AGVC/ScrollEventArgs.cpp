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
#include "ScrollEventArgs.h"

IMPLEMENT_DYNCREATE(ScrollEventArgs, CCmdTarget)

//{B7208B5F-B8E0-4A9B-A5A1-E6850492549E}
static const IID IID_IScrollEventArgs = { 0xB7208B5F, 0xB8E0, 0x4A9B, { 0xA5, 0xA1, 0xE6, 0x85, 0x04, 0x92, 0x54, 0x9E} };

//{D42F2C72-9AEA-45C4-A1A2-4B5F6373A839}
IMPLEMENT_OLECREATE_FLAGS(ScrollEventArgs, "AGVC.ScrollEventArgs", afxRegApartmentThreading, 0xD42F2C72, 0x9AEA, 0x45C4, 0xA1, 0xA2, 0x4B, 0x5F, 0x63, 0x73, 0xA8, 0x39)

BEGIN_DISPATCH_MAP(ScrollEventArgs, CCmdTarget)
	DISP_PROPERTY_EX_ID(ScrollEventArgs, "ScrollBarType", 1, odl_GetScrollBarType, odl_SetScrollBarType, VT_I4) 
	DISP_PROPERTY_EX_ID(ScrollEventArgs, "Offset", 2, odl_GetOffset, odl_SetOffset, VT_I4) 
END_DISPATCH_MAP()

BEGIN_INTERFACE_MAP(ScrollEventArgs, CCmdTarget)
	INTERFACE_PART(ScrollEventArgs, IID_IScrollEventArgs, Dispatch)
END_INTERFACE_MAP()

BEGIN_MESSAGE_MAP(ScrollEventArgs, CCmdTarget)
END_MESSAGE_MAP()

ScrollEventArgs::ScrollEventArgs(void)
{
	EnableAutomation();
	AfxOleLockApp();
	Clear();
}

ScrollEventArgs::~ScrollEventArgs()
{
	AfxOleUnlockApp();
}

void ScrollEventArgs::OnFinalRelease()
{
	CCmdTarget::OnFinalRelease();
}



LONG ScrollEventArgs::GetScrollBarType(void)
{
	return mp_yScrollBarType;
}

void ScrollEventArgs::SetScrollBarType(LONG Value)
{
	mp_yScrollBarType = Value;
}

LONG ScrollEventArgs::GetOffset(void)
{
	return mp_lOffset;
}

void ScrollEventArgs::SetOffset(LONG Value)
{
	mp_lOffset = Value;
}

void ScrollEventArgs::Clear(void)
{
	mp_yScrollBarType = 0;
	mp_lOffset = 0;
}