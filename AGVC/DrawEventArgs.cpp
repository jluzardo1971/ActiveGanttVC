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
#include "DrawEventArgs.h"

IMPLEMENT_DYNCREATE(DrawEventArgs, CCmdTarget)

//{B6D3C1FB-C891-4802-88FF-1956E4DA1AFA}
static const IID IID_IDrawEventArgs = { 0xB6D3C1FB, 0xC891, 0x4802, { 0x88, 0xFF, 0x19, 0x56, 0xE4, 0xDA, 0x1A, 0xFA} };

//{DCE73646-554E-4FDF-9DCE-67CD7C6CFEEC}
IMPLEMENT_OLECREATE_FLAGS(DrawEventArgs, "AGVC.DrawEventArgs", afxRegApartmentThreading, 0xDCE73646, 0x554E, 0x4FDF, 0x9D, 0xCE, 0x67, 0xCD, 0x7C, 0x6C, 0xFE, 0xEC)

BEGIN_DISPATCH_MAP(DrawEventArgs, CCmdTarget)
	DISP_PROPERTY_EX_ID(DrawEventArgs, "EventTarget", 1, odl_GetEventTarget, odl_SetEventTarget, VT_I4) 
	DISP_PROPERTY_EX_ID(DrawEventArgs, "CustomDraw", 2, odl_GetCustomDraw, odl_SetCustomDraw, VT_BOOL) 
	DISP_PROPERTY_EX_ID(DrawEventArgs, "ObjectIndex", 3, odl_GetObjectIndex, odl_SetObjectIndex, VT_I4) 
	DISP_PROPERTY_EX_ID(DrawEventArgs, "ParentObjectIndex", 4, odl_GetParentObjectIndex, odl_SetParentObjectIndex, VT_I4) 
	DISP_PROPERTY_EX_ID(DrawEventArgs, "Graphics", 5, odl_GetGraphics, SetNotSupported, VT_DISPATCH) 
END_DISPATCH_MAP()

BEGIN_INTERFACE_MAP(DrawEventArgs, CCmdTarget)
	INTERFACE_PART(DrawEventArgs, IID_IDrawEventArgs, Dispatch)
END_INTERFACE_MAP()

BEGIN_MESSAGE_MAP(DrawEventArgs, CCmdTarget)
END_MESSAGE_MAP()

DrawEventArgs::DrawEventArgs(void)
{
	EnableAutomation();
	AfxOleLockApp();
	Clear();
}

DrawEventArgs::~DrawEventArgs()
{
	AfxOleUnlockApp();
}

void DrawEventArgs::OnFinalRelease()
{
	CCmdTarget::OnFinalRelease();
}

LONG DrawEventArgs::GetEventTarget(void)
{
	return mp_yEventTarget;
}

void DrawEventArgs::SetEventTarget(LONG Value)
{
	mp_yEventTarget = Value;
}

BOOL DrawEventArgs::GetCustomDraw(void)
{
	return mp_bCustomDraw;
}

void DrawEventArgs::SetCustomDraw(BOOL Value)
{
	mp_bCustomDraw = Value;
}

LONG DrawEventArgs::GetObjectIndex(void)
{
	return mp_lObjectIndex;
}

void DrawEventArgs::SetObjectIndex(LONG Value)
{
	mp_lObjectIndex = Value;
}

LONG DrawEventArgs::GetParentObjectIndex(void)
{
	return mp_lParentObjectIndex;
}

void DrawEventArgs::SetParentObjectIndex(LONG Value)
{
	mp_lParentObjectIndex = Value;
}

clsGDIGraphics* DrawEventArgs::GetGraphics(void)
{
	return &mp_oGraphics;
}

void DrawEventArgs::Clear(void)
{
	mp_yEventTarget = 0;
	mp_bCustomDraw = FALSE;
	mp_lObjectIndex = 0;
	mp_lParentObjectIndex = 0;
}