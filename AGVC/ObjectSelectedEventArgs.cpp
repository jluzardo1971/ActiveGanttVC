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
#include "ObjectSelectedEventArgs.h"

IMPLEMENT_DYNCREATE(ObjectSelectedEventArgs, CCmdTarget)

//{5DCF8B10-1EB5-464E-92F3-A7FEECA77EB3}
static const IID IID_IObjectSelectedEventArgs = { 0x5DCF8B10, 0x1EB5, 0x464E, { 0x92, 0xF3, 0xA7, 0xFE, 0xEC, 0xA7, 0x7E, 0xB3} };

//{4F2415FE-F559-4C54-B31D-BCB1DCA7BB1F}
IMPLEMENT_OLECREATE_FLAGS(ObjectSelectedEventArgs, "AGVC.ObjectSelectedEventArgs", afxRegApartmentThreading, 0x4F2415FE, 0xF559, 0x4C54, 0xB3, 0x1D, 0xBC, 0xB1, 0xDC, 0xA7, 0xBB, 0x1F)

BEGIN_DISPATCH_MAP(ObjectSelectedEventArgs, CCmdTarget)
	DISP_PROPERTY_EX_ID(ObjectSelectedEventArgs, "EventTarget", 1, odl_GetEventTarget, odl_SetEventTarget, VT_I4) 
	DISP_PROPERTY_EX_ID(ObjectSelectedEventArgs, "ObjectIndex", 2, odl_GetObjectIndex, odl_SetObjectIndex, VT_I4) 
	DISP_PROPERTY_EX_ID(ObjectSelectedEventArgs, "ParentObjectIndex", 3, odl_GetParentObjectIndex, odl_SetParentObjectIndex, VT_I4) 
END_DISPATCH_MAP()

BEGIN_INTERFACE_MAP(ObjectSelectedEventArgs, CCmdTarget)
	INTERFACE_PART(ObjectSelectedEventArgs, IID_IObjectSelectedEventArgs, Dispatch)
END_INTERFACE_MAP()

BEGIN_MESSAGE_MAP(ObjectSelectedEventArgs, CCmdTarget)
END_MESSAGE_MAP()

ObjectSelectedEventArgs::ObjectSelectedEventArgs(void)
{
	EnableAutomation();
	AfxOleLockApp();
	Clear();
}

ObjectSelectedEventArgs::~ObjectSelectedEventArgs()
{
	AfxOleUnlockApp();
}

void ObjectSelectedEventArgs::OnFinalRelease()
{
	CCmdTarget::OnFinalRelease();
}



LONG ObjectSelectedEventArgs::GetEventTarget(void)
{
	return mp_yEventTarget;
}

void ObjectSelectedEventArgs::SetEventTarget(LONG Value)
{
	mp_yEventTarget = Value;
}

LONG ObjectSelectedEventArgs::GetObjectIndex(void)
{
	return mp_lObjectIndex;
}

void ObjectSelectedEventArgs::SetObjectIndex(LONG Value)
{
	mp_lObjectIndex = Value;
}

LONG ObjectSelectedEventArgs::GetParentObjectIndex(void)
{
	return mp_lParentObjectIndex;
}

void ObjectSelectedEventArgs::SetParentObjectIndex(LONG Value)
{
	mp_lParentObjectIndex = Value;
}

void ObjectSelectedEventArgs::Clear(void)
{
	mp_yEventTarget = 0;
	mp_lObjectIndex = 0;
	mp_lParentObjectIndex = 0;
}