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
#include "ObjectStateChangedEventArgs.h"

IMPLEMENT_DYNCREATE(ObjectStateChangedEventArgs, CCmdTarget)

//{54AD7E85-3885-40D2-BF6A-07EA30563764}
static const IID IID_IObjectStateChangedEventArgs = { 0x54AD7E85, 0x3885, 0x40D2, { 0xBF, 0x6A, 0x07, 0xEA, 0x30, 0x56, 0x37, 0x64} };

//{20B012E9-C84B-47BB-8DDB-32838732F148}
IMPLEMENT_OLECREATE_FLAGS(ObjectStateChangedEventArgs, "AGVC.ObjectStateChangedEventArgs", afxRegApartmentThreading, 0x20B012E9, 0xC84B, 0x47BB, 0x8D, 0xDB, 0x32, 0x83, 0x87, 0x32, 0xF1, 0x48)

BEGIN_DISPATCH_MAP(ObjectStateChangedEventArgs, CCmdTarget)
	DISP_PROPERTY_EX_ID(ObjectStateChangedEventArgs, "EventTarget", 1, odl_GetEventTarget, odl_SetEventTarget, VT_I4) 
	DISP_PROPERTY_EX_ID(ObjectStateChangedEventArgs, "Index", 2, odl_GetIndex, odl_SetIndex, VT_I4) 
	DISP_PROPERTY_EX_ID(ObjectStateChangedEventArgs, "Cancel", 3, odl_GetCancel, odl_SetCancel, VT_BOOL) 
	DISP_PROPERTY_EX_ID(ObjectStateChangedEventArgs, "DestinationIndex", 4, odl_GetDestinationIndex, odl_SetDestinationIndex, VT_I4) 
	DISP_PROPERTY_EX_ID(ObjectStateChangedEventArgs, "InitialRowIndex", 5, odl_GetInitialRowIndex, odl_SetInitialRowIndex, VT_I4) 
	DISP_PROPERTY_EX_ID(ObjectStateChangedEventArgs, "FinalRowIndex", 6, odl_GetFinalRowIndex, odl_SetFinalRowIndex, VT_I4) 
	DISP_PROPERTY_EX_ID(ObjectStateChangedEventArgs, "InitialColumnIndex", 7, odl_GetInitialColumnIndex, odl_SetInitialColumnIndex, VT_I4) 
	DISP_PROPERTY_EX_ID(ObjectStateChangedEventArgs, "FinalColumnIndex", 8, odl_GetFinalColumnIndex, odl_SetFinalColumnIndex, VT_I4)
	DISP_PROPERTY_EX_ID(ObjectStateChangedEventArgs, "StartDate", 9, odl_GetStartDate, SetNotSupported, VT_DATE)
	DISP_PROPERTY_EX_ID(ObjectStateChangedEventArgs, "EndDate", 10, odl_GetStartDate, SetNotSupported, VT_DATE)
	DISP_PROPERTY_EX_ID(ObjectStateChangedEventArgs, "ChangeType", 11, odl_GetChangeType, SetNotSupported, VT_I4)
END_DISPATCH_MAP()

BEGIN_INTERFACE_MAP(ObjectStateChangedEventArgs, CCmdTarget)
	INTERFACE_PART(ObjectStateChangedEventArgs, IID_IObjectStateChangedEventArgs, Dispatch)
END_INTERFACE_MAP()

BEGIN_MESSAGE_MAP(ObjectStateChangedEventArgs, CCmdTarget)
END_MESSAGE_MAP()

ObjectStateChangedEventArgs::ObjectStateChangedEventArgs(void)
{
	EnableAutomation();
	AfxOleLockApp();
	Clear();
}

ObjectStateChangedEventArgs::~ObjectStateChangedEventArgs()
{
	AfxOleUnlockApp();
}

void ObjectStateChangedEventArgs::OnFinalRelease()
{
	CCmdTarget::OnFinalRelease();
}



LONG ObjectStateChangedEventArgs::GetEventTarget(void)
{
	return mp_yEventTarget;
}

void ObjectStateChangedEventArgs::SetEventTarget(LONG Value)
{
	mp_yEventTarget = Value;
}

LONG ObjectStateChangedEventArgs::GetIndex(void)
{
	return mp_lIndex;
}

void ObjectStateChangedEventArgs::SetIndex(LONG Value)
{
	mp_lIndex = Value;
}

BOOL ObjectStateChangedEventArgs::GetCancel(void)
{
	return mp_bCancel;
}

void ObjectStateChangedEventArgs::SetCancel(BOOL Value)
{
	mp_bCancel = Value;
}

LONG ObjectStateChangedEventArgs::GetDestinationIndex(void)
{
	return mp_lDestinationIndex;
}

void ObjectStateChangedEventArgs::SetDestinationIndex(LONG Value)
{
	mp_lDestinationIndex = Value;
}

LONG ObjectStateChangedEventArgs::GetInitialRowIndex(void)
{
	return mp_lInitialRowIndex;
}

void ObjectStateChangedEventArgs::SetInitialRowIndex(LONG Value)
{
	mp_lInitialRowIndex = Value;
}

LONG ObjectStateChangedEventArgs::GetFinalRowIndex(void)
{
	return mp_lFinalRowIndex;
}

void ObjectStateChangedEventArgs::SetFinalRowIndex(LONG Value)
{
	mp_lFinalRowIndex = Value;
}

LONG ObjectStateChangedEventArgs::GetInitialColumnIndex(void)
{
	return mp_lInitialColumnIndex;
}

void ObjectStateChangedEventArgs::SetInitialColumnIndex(LONG Value)
{
	mp_lInitialColumnIndex = Value;
}

LONG ObjectStateChangedEventArgs::GetFinalColumnIndex(void)
{
	return mp_lFinalColumnIndex;
}

void ObjectStateChangedEventArgs::SetFinalColumnIndex(LONG Value)
{
	mp_lFinalColumnIndex = Value;
}

COleDateTime ObjectStateChangedEventArgs::GetStartDate(void)
{
	return mp_dtStartDate;
}

void ObjectStateChangedEventArgs::SetStartDate(COleDateTime Value)
{
	mp_dtStartDate = Value;
}

COleDateTime ObjectStateChangedEventArgs::GetEndDate(void)
{
	return mp_dtEndDate;
}

void ObjectStateChangedEventArgs::SetEndDate(COleDateTime Value)
{
	mp_dtEndDate = Value;
}

LONG ObjectStateChangedEventArgs::GetChangeType(void)
{
	return mp_lChangeType;
}

void ObjectStateChangedEventArgs::SetChangeType(LONG Value)
{
	mp_lChangeType = Value;
}

void ObjectStateChangedEventArgs::Clear(void)
{
	mp_yEventTarget = 0;
	mp_lIndex = 0;
	mp_dtStartDate = (DATE)0;
	mp_dtEndDate = (DATE)0;
	mp_bCancel = FALSE;
	mp_lChangeType = 0;
}