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
#include "TickMarkTextDrawEventArgs.h"

IMPLEMENT_DYNCREATE(TickMarkTextDrawEventArgs, CCmdTarget)

//{DB523317-B81D-47BB-8A47-084C8CC51E58}
static const IID IID_ITickMarkTextDrawEventArgs = { 0xDB523317, 0xB81D, 0x47BB, { 0x8A, 0x47, 0x08, 0x4C, 0x8C, 0xC5, 0x1E, 0x58} };

//{96FE4B60-D1D4-4ABE-ADF3-A82F82587DFE}
IMPLEMENT_OLECREATE_FLAGS(TickMarkTextDrawEventArgs, "AGVC.TickMarkTextDrawEventArgs", afxRegApartmentThreading, 0x96FE4B60, 0xD1D4, 0x4ABE, 0xAD, 0xF3, 0xA8, 0x2F, 0x82, 0x58, 0x7D, 0xFE)

BEGIN_DISPATCH_MAP(TickMarkTextDrawEventArgs, CCmdTarget)
	DISP_PROPERTY_EX_ID(TickMarkTextDrawEventArgs, "CustomDraw", 1, odl_GetCustomDraw, odl_SetCustomDraw, VT_BOOL) 
	DISP_PROPERTY_EX_ID(TickMarkTextDrawEventArgs, "Text", 2, odl_GetText, odl_SetText, VT_BSTR)
	DISP_PROPERTY_EX_ID(TickMarkTextDrawEventArgs, "dtDate", 3, odl_GetdtDate, SetNotSupported, VT_DATE) 
END_DISPATCH_MAP()

BEGIN_INTERFACE_MAP(TickMarkTextDrawEventArgs, CCmdTarget)
	INTERFACE_PART(TickMarkTextDrawEventArgs, IID_ITickMarkTextDrawEventArgs, Dispatch)
END_INTERFACE_MAP()

BEGIN_MESSAGE_MAP(TickMarkTextDrawEventArgs, CCmdTarget)
END_MESSAGE_MAP()

TickMarkTextDrawEventArgs::TickMarkTextDrawEventArgs(void)
{
	EnableAutomation();
	AfxOleLockApp();
	Clear();
}

TickMarkTextDrawEventArgs::~TickMarkTextDrawEventArgs()
{
	AfxOleUnlockApp();
}

void TickMarkTextDrawEventArgs::OnFinalRelease()
{
	CCmdTarget::OnFinalRelease();
}

void TickMarkTextDrawEventArgs::Clear(void)
{
	mp_bCustomDraw = FALSE;
	mp_sText = _T("");
	mp_dtDate = (DATE)0;
}



CString TickMarkTextDrawEventArgs::GetText(void)
{
	return mp_sText;
}

void TickMarkTextDrawEventArgs::SetText(CString Value)
{
	mp_sText = Value;
}

BOOL TickMarkTextDrawEventArgs::GetCustomDraw(void)
{
	return mp_bCustomDraw;
}

void TickMarkTextDrawEventArgs::SetCustomDraw(BOOL Value)
{
	mp_bCustomDraw = Value;
}