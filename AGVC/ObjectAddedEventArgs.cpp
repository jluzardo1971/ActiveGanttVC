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
#include "ObjectAddedEventArgs.h"

IMPLEMENT_DYNCREATE(ObjectAddedEventArgs, CCmdTarget)

//{BC1197F1-B59B-4BC6-B33F-65461516C17E}
static const IID IID_IObjectAddedEventArgs = { 0xBC1197F1, 0xB59B, 0x4BC6, { 0xB3, 0x3F, 0x65, 0x46, 0x15, 0x16, 0xC1, 0x7E} };

//{72CE6550-58B2-44B9-B93F-8E41525040F7}
IMPLEMENT_OLECREATE_FLAGS(ObjectAddedEventArgs, "AGVC.ObjectAddedEventArgs", afxRegApartmentThreading, 0x72CE6550, 0x58B2, 0x44B9, 0xB9, 0x3F, 0x8E, 0x41, 0x52, 0x50, 0x40, 0xF7)

BEGIN_DISPATCH_MAP(ObjectAddedEventArgs, CCmdTarget)
	DISP_PROPERTY_EX_ID(ObjectAddedEventArgs, "TaskIndex", 1, odl_GetTaskIndex, odl_SetTaskIndex, VT_I4) 
	DISP_PROPERTY_EX_ID(ObjectAddedEventArgs, "PredecessorObjectIndex", 2, odl_GetPredecessorObjectIndex, odl_SetPredecessorObjectIndex, VT_I4) 
	DISP_PROPERTY_EX_ID(ObjectAddedEventArgs, "PredecessorTaskIndex", 3, odl_GetPredecessorTaskIndex, odl_SetPredecessorTaskIndex, VT_I4) 
	DISP_PROPERTY_EX_ID(ObjectAddedEventArgs, "PredecessorType", 4, odl_GetPredecessorType, odl_SetPredecessorType, VT_I4) 
	DISP_PROPERTY_EX_ID(ObjectAddedEventArgs, "TaskKey", 5, odl_GetTaskKey, odl_SetTaskKey, VT_BSTR) 
	DISP_PROPERTY_EX_ID(ObjectAddedEventArgs, "PredecessorTaskKey", 6, odl_GetPredecessorTaskKey, odl_SetPredecessorTaskKey, VT_BSTR) 
	DISP_PROPERTY_EX_ID(ObjectAddedEventArgs, "EventTarget", 7, odl_GetEventTarget, odl_SetEventTarget, VT_I4) 
END_DISPATCH_MAP()

BEGIN_INTERFACE_MAP(ObjectAddedEventArgs, CCmdTarget)
	INTERFACE_PART(ObjectAddedEventArgs, IID_IObjectAddedEventArgs, Dispatch)
END_INTERFACE_MAP()

BEGIN_MESSAGE_MAP(ObjectAddedEventArgs, CCmdTarget)
END_MESSAGE_MAP()

ObjectAddedEventArgs::ObjectAddedEventArgs(void)
{
	EnableAutomation();
	AfxOleLockApp();
	Clear();
}

ObjectAddedEventArgs::~ObjectAddedEventArgs()
{
	AfxOleUnlockApp();
}

void ObjectAddedEventArgs::OnFinalRelease()
{
	CCmdTarget::OnFinalRelease();
}

LONG ObjectAddedEventArgs::GetTaskIndex(void)
{
	return mp_lTaskIndex;
}

void ObjectAddedEventArgs::SetTaskIndex(LONG Value)
{
	mp_lTaskIndex = Value;
}

LONG ObjectAddedEventArgs::GetPredecessorObjectIndex(void)
{
	return mp_lPredecessorObjectIndex;
}

void ObjectAddedEventArgs::SetPredecessorObjectIndex(LONG Value)
{
	mp_lPredecessorObjectIndex = Value;
}

LONG ObjectAddedEventArgs::GetPredecessorTaskIndex(void)
{
	return mp_lPredecessorTaskIndex;
}

void ObjectAddedEventArgs::SetPredecessorTaskIndex(LONG Value)
{
	mp_lPredecessorTaskIndex = Value;
}

LONG ObjectAddedEventArgs::GetPredecessorType(void)
{
	return mp_yPredecessorType;
}

void ObjectAddedEventArgs::SetPredecessorType(LONG Value)
{
	mp_yPredecessorType = Value;
}

CString ObjectAddedEventArgs::GetTaskKey(void)
{
	return mp_sTaskKey;
}

void ObjectAddedEventArgs::SetTaskKey(CString Value)
{
	mp_sTaskKey = Value;
}

CString ObjectAddedEventArgs::GetPredecessorTaskKey(void)
{
	return mp_sPredecessorTaskKey;
}

void ObjectAddedEventArgs::SetPredecessorTaskKey(CString Value)
{
	mp_sPredecessorTaskKey = Value;
}

LONG ObjectAddedEventArgs::GetEventTarget(void)
{
	return mp_yEventTarget;
}

void ObjectAddedEventArgs::SetEventTarget(LONG Value)
{
	mp_yEventTarget = Value;
}

void ObjectAddedEventArgs::Clear(void)
{
	mp_lTaskIndex = 0;
	mp_lPredecessorObjectIndex = 0;
	mp_lPredecessorTaskIndex = 0;
	mp_yPredecessorType = 0;
	mp_sTaskKey = "";
	mp_sPredecessorTaskKey = "";
	mp_yEventTarget = 0;
}