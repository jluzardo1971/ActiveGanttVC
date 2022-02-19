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
#include "ToolTipEventArgs.h"

IMPLEMENT_DYNCREATE(ToolTipEventArgs, CCmdTarget)

//{984AD3A4-0822-4BED-87E0-5EBF27776FBC}
static const IID IID_IToolTipEventArgs = { 0x984AD3A4, 0x0822, 0x4BED, { 0x87, 0xE0, 0x5E, 0xBF, 0x27, 0x77, 0x6F, 0xBC} };

//{AFF74710-C6A8-4F6F-A6D9-B19D5EA3F28B}
IMPLEMENT_OLECREATE_FLAGS(ToolTipEventArgs, "AGVC.ToolTipEventArgs", afxRegApartmentThreading, 0xAFF74710, 0xC6A8, 0x4F6F, 0xA6, 0xD9, 0xB1, 0x9D, 0x5E, 0xA3, 0xF2, 0x8B)

BEGIN_DISPATCH_MAP(ToolTipEventArgs, CCmdTarget)
	DISP_PROPERTY_EX_ID(ToolTipEventArgs, "InitialRowIndex", 1, odl_GetInitialRowIndex, odl_SetInitialRowIndex, VT_I4) 
	DISP_PROPERTY_EX_ID(ToolTipEventArgs, "FinalRowIndex", 2, odl_GetFinalRowIndex, odl_SetFinalRowIndex, VT_I4) 
	DISP_PROPERTY_EX_ID(ToolTipEventArgs, "TaskIndex", 3, odl_GetTaskIndex, odl_SetTaskIndex, VT_I4) 
	DISP_PROPERTY_EX_ID(ToolTipEventArgs, "MilestoneIndex", 4, odl_GetMilestoneIndex, odl_SetMilestoneIndex, VT_I4) 
	DISP_PROPERTY_EX_ID(ToolTipEventArgs, "PercentageIndex", 5, odl_GetPercentageIndex, odl_SetPercentageIndex, VT_I4) 
	DISP_PROPERTY_EX_ID(ToolTipEventArgs, "RowIndex", 6, odl_GetRowIndex, odl_SetRowIndex, VT_I4) 
	DISP_PROPERTY_EX_ID(ToolTipEventArgs, "CellIndex", 7, odl_GetCellIndex, odl_SetCellIndex, VT_I4) 
	DISP_PROPERTY_EX_ID(ToolTipEventArgs, "ColumnIndex", 8, odl_GetColumnIndex, odl_SetColumnIndex, VT_I4) 
	DISP_PROPERTY_EX_ID(ToolTipEventArgs, "InitialStartDate", 9, odl_GetInitialStartDate, odl_SetInitialStartDate, VT_DATE) 
	DISP_PROPERTY_EX_ID(ToolTipEventArgs, "InitialEndDate", 10, odl_GetInitialEndDate, odl_SetInitialEndDate, VT_DATE) 
	DISP_PROPERTY_EX_ID(ToolTipEventArgs, "StartDate", 11, odl_GetStartDate, odl_SetStartDate, VT_DATE) 
	DISP_PROPERTY_EX_ID(ToolTipEventArgs, "EndDate", 12, odl_GetEndDate, odl_SetEndDate, VT_DATE) 
	DISP_PROPERTY_EX_ID(ToolTipEventArgs, "XStart", 13, odl_GetXStart, odl_SetXStart, VT_I4) 
	DISP_PROPERTY_EX_ID(ToolTipEventArgs, "XEnd", 14, odl_GetXEnd, odl_SetXEnd, VT_I4) 
	DISP_PROPERTY_EX_ID(ToolTipEventArgs, "Operation", 15, odl_GetOperation, odl_SetOperation, VT_I4) 
	DISP_PROPERTY_EX_ID(ToolTipEventArgs, "EventTarget", 16, odl_GetEventTarget, odl_SetEventTarget, VT_I4) 
	DISP_PROPERTY_EX_ID(ToolTipEventArgs, "TaskPosition", 17, odl_GetTaskPosition, odl_SetTaskPosition, VT_BSTR) 
	DISP_PROPERTY_EX_ID(ToolTipEventArgs, "PredecessorPosition", 18, odl_GetPredecessorPosition, odl_SetPredecessorPosition, VT_BSTR) 
	DISP_PROPERTY_EX_ID(ToolTipEventArgs, "X", 19, odl_GetX, odl_SetX, VT_I4) 
	DISP_PROPERTY_EX_ID(ToolTipEventArgs, "Y", 20, odl_GetY, odl_SetY, VT_I4) 
	DISP_PROPERTY_EX_ID(ToolTipEventArgs, "X1", 21, odl_GetX1, odl_SetX1, VT_I4) 
	DISP_PROPERTY_EX_ID(ToolTipEventArgs, "Y1", 22, odl_GetY1, odl_SetY1, VT_I4) 
	DISP_PROPERTY_EX_ID(ToolTipEventArgs, "X2", 23, odl_GetX2, odl_SetX2, VT_I4) 
	DISP_PROPERTY_EX_ID(ToolTipEventArgs, "Y2", 24, odl_GetY2, odl_SetY2, VT_I4) 
	DISP_PROPERTY_EX_ID(ToolTipEventArgs, "CustomDraw", 25, odl_GetCustomDraw, odl_SetCustomDraw, VT_BOOL) 
	DISP_PROPERTY_EX_ID(ToolTipEventArgs, "Graphics", 26, odl_GetGraphics, SetNotSupported, VT_DISPATCH) 
	DISP_PROPERTY_EX_ID(ToolTipEventArgs, "ToolTipType", 27, odl_GetToolTipType, odl_SetToolTipType, VT_I4) 
END_DISPATCH_MAP()

BEGIN_INTERFACE_MAP(ToolTipEventArgs, CCmdTarget)
	INTERFACE_PART(ToolTipEventArgs, IID_IToolTipEventArgs, Dispatch)
END_INTERFACE_MAP()

BEGIN_MESSAGE_MAP(ToolTipEventArgs, CCmdTarget)
END_MESSAGE_MAP()

ToolTipEventArgs::ToolTipEventArgs(void)
{
	EnableAutomation();
	AfxOleLockApp();
	Clear();
}

ToolTipEventArgs::~ToolTipEventArgs()
{
	AfxOleUnlockApp();
}

void ToolTipEventArgs::OnFinalRelease()
{
	CCmdTarget::OnFinalRelease();
}

LONG ToolTipEventArgs::GetInitialRowIndex(void)
{
	return mp_lInitialRowIndex;
}

void ToolTipEventArgs::SetInitialRowIndex(LONG Value)
{
	mp_lInitialRowIndex = Value;
}

LONG ToolTipEventArgs::GetFinalRowIndex(void)
{
	return mp_lFinalRowIndex;
}

void ToolTipEventArgs::SetFinalRowIndex(LONG Value)
{
	mp_lFinalRowIndex = Value;
}

LONG ToolTipEventArgs::GetTaskIndex(void)
{
	return mp_lTaskIndex;
}

void ToolTipEventArgs::SetTaskIndex(LONG Value)
{
	mp_lTaskIndex = Value;
}

LONG ToolTipEventArgs::GetMilestoneIndex(void)
{
	return mp_lMilestoneIndex;
}

void ToolTipEventArgs::SetMilestoneIndex(LONG Value)
{
	mp_lMilestoneIndex = Value;
}

LONG ToolTipEventArgs::GetPercentageIndex(void)
{
	return mp_lPercentageIndex;
}

void ToolTipEventArgs::SetPercentageIndex(LONG Value)
{
	mp_lPercentageIndex = Value;
}

LONG ToolTipEventArgs::GetRowIndex(void)
{
	return mp_lRowIndex;
}

void ToolTipEventArgs::SetRowIndex(LONG Value)
{
	mp_lRowIndex = Value;
}

LONG ToolTipEventArgs::GetCellIndex(void)
{
	return mp_lCellIndex;
}

void ToolTipEventArgs::SetCellIndex(LONG Value)
{
	mp_lCellIndex = Value;
}

LONG ToolTipEventArgs::GetColumnIndex(void)
{
	return mp_lColumnIndex;
}

void ToolTipEventArgs::SetColumnIndex(LONG Value)
{
	mp_lColumnIndex = Value;
}

COleDateTime ToolTipEventArgs::GetInitialStartDate(void)
{
	return mp_dtInitialStartDate;
}

void ToolTipEventArgs::SetInitialStartDate(COleDateTime Value)
{
	mp_dtInitialStartDate = Value;
}

COleDateTime ToolTipEventArgs::GetInitialEndDate(void)
{
	return mp_dtInitialEndDate;
}

void ToolTipEventArgs::SetInitialEndDate(COleDateTime Value)
{
	mp_dtInitialEndDate = Value;
}

COleDateTime ToolTipEventArgs::GetStartDate(void)
{
	return mp_dtStartDate;
}

void ToolTipEventArgs::SetStartDate(COleDateTime Value)
{
	mp_dtStartDate = Value;
}

COleDateTime ToolTipEventArgs::GetEndDate(void)
{
	return mp_dtEndDate;
}

void ToolTipEventArgs::SetEndDate(COleDateTime Value)
{
	mp_dtEndDate = Value;
}

LONG ToolTipEventArgs::GetXStart(void)
{
	return mp_lXStart;
}

void ToolTipEventArgs::SetXStart(LONG Value)
{
	mp_lXStart = Value;
}

LONG ToolTipEventArgs::GetXEnd(void)
{
	return mp_lXEnd;
}

void ToolTipEventArgs::SetXEnd(LONG Value)
{
	mp_lXEnd = Value;
}

LONG ToolTipEventArgs::GetOperation(void)
{
	return mp_yOperation;
}

void ToolTipEventArgs::SetOperation(LONG Value)
{
	mp_yOperation = Value;
}

LONG ToolTipEventArgs::GetEventTarget(void)
{
	return mp_yEventTarget;
}

void ToolTipEventArgs::SetEventTarget(LONG Value)
{
	mp_yEventTarget = Value;
}

CString ToolTipEventArgs::GetTaskPosition(void)
{
	return mp_sTaskPosition;
}

void ToolTipEventArgs::SetTaskPosition(CString Value)
{
	mp_sTaskPosition = Value;
}

CString ToolTipEventArgs::GetPredecessorPosition(void)
{
	return mp_sPredecessorPosition;
}

void ToolTipEventArgs::SetPredecessorPosition(CString Value)
{
	mp_sPredecessorPosition = Value;
}

LONG ToolTipEventArgs::GetX(void)
{
	return mp_lX;
}

void ToolTipEventArgs::SetX(LONG Value)
{
	mp_lX = Value;
}

LONG ToolTipEventArgs::GetY(void)
{
	return mp_lY;
}

void ToolTipEventArgs::SetY(LONG Value)
{
	mp_lY = Value;
}

LONG ToolTipEventArgs::GetX1(void)
{
	return mp_lX1;
}

void ToolTipEventArgs::SetX1(LONG Value)
{
	mp_lX1 = Value;
}

LONG ToolTipEventArgs::GetY1(void)
{
	return mp_lY1;
}

void ToolTipEventArgs::SetY1(LONG Value)
{
	mp_lY1 = Value;
}

LONG ToolTipEventArgs::GetX2(void)
{
	return mp_lX2;
}

void ToolTipEventArgs::SetX2(LONG Value)
{
	mp_lX2 = Value;
}

LONG ToolTipEventArgs::GetY2(void)
{
	return mp_lY2;
}

void ToolTipEventArgs::SetY2(LONG Value)
{
	mp_lY2 = Value;
}

BOOL ToolTipEventArgs::GetCustomDraw(void)
{
	return mp_bCustomDraw;
}

void ToolTipEventArgs::SetCustomDraw(BOOL Value)
{
	mp_bCustomDraw = Value;
}

clsGDIGraphics* ToolTipEventArgs::GetGraphics(void)
{
	return &mp_oGraphics;
}

LONG ToolTipEventArgs::GetToolTipType(void)
{
	return mp_yToolTipType;
}

void ToolTipEventArgs::SetToolTipType(LONG Value)
{
	mp_yToolTipType = Value;
}

void ToolTipEventArgs::Clear(void)
{
	mp_lInitialRowIndex = 0;
	mp_lFinalRowIndex = 0;
	mp_lRowIndex = 0;
	mp_lTaskIndex = 0;
	mp_lMilestoneIndex = 0;
	mp_lPercentageIndex = 0;
	mp_lCellIndex = 0;
	mp_lColumnIndex = 0;
	mp_dtStartDate = (DATE)0;
	mp_dtEndDate = (DATE)0;
	mp_dtInitialStartDate = (DATE)0;
	mp_dtInitialEndDate = (DATE)0;
	mp_lXStart = 0;
	mp_lXEnd = 0;
	mp_lX = 0;
	mp_lY = 0;
	mp_lX1 = 0;
	mp_lY1 = 0;
	mp_lX2 = 0;
	mp_lY2 = 0;
	mp_yOperation = 0;
	mp_yEventTarget = 0;
	mp_yToolTipType = 0;
}