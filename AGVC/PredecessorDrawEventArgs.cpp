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
#include "PredecessorDrawEventArgs.h"

IMPLEMENT_DYNCREATE(PredecessorDrawEventArgs, CCmdTarget)

//{B0C7C442-81C0-449C-A916-16CE8C8129B4}
static const IID IID_IPredecessorDrawEventArgs = { 0xB0C7C442, 0x81C0, 0x449C, { 0xA9, 0x16, 0x16, 0xCE, 0x8C, 0x81, 0x29, 0xB4} };

//{217E5088-3784-489D-B07B-2CEE068EE91B}
IMPLEMENT_OLECREATE_FLAGS(PredecessorDrawEventArgs, "AGVC.PredecessorDrawEventArgs", afxRegApartmentThreading, 0x217E5088, 0x3784, 0x489D, 0xB0, 0x7B, 0x2C, 0xEE, 0x06, 0x8E, 0xE9, 0x1B)

BEGIN_DISPATCH_MAP(PredecessorDrawEventArgs, CCmdTarget)
	DISP_PROPERTY_EX_ID(PredecessorDrawEventArgs, "CustomDraw", 1, odl_GetCustomDraw, odl_SetCustomDraw, VT_BOOL) 
	DISP_PROPERTY_EX_ID(PredecessorDrawEventArgs, "Graphics", 2, odl_GetGraphics, SetNotSupported, VT_DISPATCH) 
	DISP_PROPERTY_EX_ID(PredecessorDrawEventArgs, "PredecessorIndex", 3, odl_GetPredecessorIndex, odl_SetPredecessorIndex, VT_I4) 
	DISP_PROPERTY_EX_ID(PredecessorDrawEventArgs, "TaskIndex", 4, odl_GetTaskIndex, odl_SetTaskIndex, VT_I4) 
	DISP_PROPERTY_EX_ID(PredecessorDrawEventArgs, "PredecessorType", 5, odl_GetPredecessorType, odl_SetPredecessorType, VT_I4) 
END_DISPATCH_MAP()

BEGIN_INTERFACE_MAP(PredecessorDrawEventArgs, CCmdTarget)
	INTERFACE_PART(PredecessorDrawEventArgs, IID_IPredecessorDrawEventArgs, Dispatch)
END_INTERFACE_MAP()

BEGIN_MESSAGE_MAP(PredecessorDrawEventArgs, CCmdTarget)
END_MESSAGE_MAP()

PredecessorDrawEventArgs::PredecessorDrawEventArgs(void)
{
	EnableAutomation();
	AfxOleLockApp();
	Clear();
}

PredecessorDrawEventArgs::~PredecessorDrawEventArgs()
{
	AfxOleUnlockApp();
}

void PredecessorDrawEventArgs::OnFinalRelease()
{
	CCmdTarget::OnFinalRelease();
}

BOOL PredecessorDrawEventArgs::GetCustomDraw(void)
{
	return mp_bCustomDraw;
}

void PredecessorDrawEventArgs::SetCustomDraw(BOOL Value)
{
	mp_bCustomDraw = Value;
}

clsGDIGraphics* PredecessorDrawEventArgs::GetGraphics(void)
{
	return &mp_oGraphics;
}

LONG PredecessorDrawEventArgs::GetPredecessorIndex(void)
{
	return mp_lPredecessorIndex;
}

void PredecessorDrawEventArgs::SetPredecessorIndex(LONG Value)
{
	mp_lPredecessorIndex = Value;
}

LONG PredecessorDrawEventArgs::GetTaskIndex(void)
{
	return mp_lTaskIndex;
}

void PredecessorDrawEventArgs::SetTaskIndex(LONG Value)
{
	mp_lTaskIndex = Value;
}

LONG PredecessorDrawEventArgs::GetPredecessorType(void)
{
	return mp_yPredecessorType;
}

void PredecessorDrawEventArgs::SetPredecessorType(LONG Value)
{
	mp_yPredecessorType = Value;
}

void PredecessorDrawEventArgs::Clear(void)
{
	mp_bCustomDraw = FALSE;
	mp_lPredecessorIndex = 0;
	mp_lTaskIndex = 0;
	mp_yPredecessorType = 0;
}