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
#include "clsClientArea.h"



IMPLEMENT_DYNCREATE(clsClientArea, CCmdTarget)

//{CD771458-FDE5-4FAA-BA09-0CD6743F0B68}
static const IID IID_IclsClientArea = { 0xCD771458, 0xFDE5, 0x4FAA, { 0xBA, 0x09, 0x0C, 0xD6, 0x74, 0x3F, 0x0B, 0x68} };

//{8217AF7D-97D1-4952-824C-658FF255EF33}
IMPLEMENT_OLECREATE_FLAGS(clsClientArea, "AGVC.clsClientArea", afxRegApartmentThreading, 0x8217AF7D, 0x97D1, 0x4952, 0x82, 0x4C, 0x65, 0x8F, 0xF2, 0x55, 0xEF, 0x33)

BEGIN_DISPATCH_MAP(clsClientArea, CCmdTarget)
	DISP_PROPERTY_EX_ID(clsClientArea, "DetectConflicts", 1, odl_GetDetectConflicts, odl_SetDetectConflicts, VT_BOOL) 
	DISP_PROPERTY_EX_ID(clsClientArea, "MilestoneSelectionOffset", 2, odl_GetMilestoneSelectionOffset, odl_SetMilestoneSelectionOffset, VT_I4) 
	DISP_PROPERTY_EX_ID(clsClientArea, "FirstVisibleRow", 3, odl_GetFirstVisibleRow, odl_SetFirstVisibleRow, VT_I4) 
	DISP_PROPERTY_EX_ID(clsClientArea, "LastVisibleRow", 4, odl_GetLastVisibleRow, SetNotSupported, VT_I4) 
	DISP_PROPERTY_EX_ID(clsClientArea, "ToolTipFormat", 5, odl_GetToolTipFormat, odl_SetToolTipFormat, VT_BSTR) 
	DISP_PROPERTY_EX_ID(clsClientArea, "ToolTipsVisible", 6, odl_GetToolTipsVisible, odl_SetToolTipsVisible, VT_BOOL) 
	DISP_PROPERTY_EX_ID(clsClientArea, "Top", 7, odl_GetTop, SetNotSupported, VT_I4) 
	DISP_PROPERTY_EX_ID(clsClientArea, "Bottom", 8, odl_GetBottom, SetNotSupported, VT_I4) 
	DISP_PROPERTY_EX_ID(clsClientArea, "Left", 9, odl_GetLeft, SetNotSupported, VT_I4) 
	DISP_PROPERTY_EX_ID(clsClientArea, "Right", 10, odl_GetRight, SetNotSupported, VT_I4) 
	DISP_PROPERTY_EX_ID(clsClientArea, "Width", 11, odl_GetWidth, SetNotSupported, VT_I4) 
	DISP_PROPERTY_EX_ID(clsClientArea, "Height", 12, odl_GetHeight, SetNotSupported, VT_I4) 
	DISP_PROPERTY_EX_ID(clsClientArea, "Grid", 13, odl_GetGrid, SetNotSupported, VT_DISPATCH) 
	DISP_FUNCTION_ID(clsClientArea, "GetXML", 14, odl_GetXML, VT_BSTR, VTS_NONE) 
	DISP_FUNCTION_ID(clsClientArea, "SetXML", 15, odl_SetXML, VT_EMPTY, VTS_BSTR)
	DISP_PROPERTY_EX_ID(clsClientArea, "PredecessorSelectionOffset", 16, odl_GetPredecessorSelectionOffset, odl_SetPredecessorSelectionOffset, VT_I4)
	DISP_PROPERTY_EX_ID(clsClientArea, "TaskBorderSelectionOffset", 17, odl_GetTaskBorderSelectionOffset, odl_SetTaskBorderSelectionOffset, VT_I4)
END_DISPATCH_MAP()

BEGIN_INTERFACE_MAP(clsClientArea, CCmdTarget)
	INTERFACE_PART(clsClientArea, IID_IclsClientArea, Dispatch)
END_INTERFACE_MAP()

BEGIN_MESSAGE_MAP(clsClientArea, CCmdTarget)
END_MESSAGE_MAP()

clsClientArea::clsClientArea(CActiveGanttVCCtl* oControl, clsTimeLine* oTimeLine)
{
	EnableAutomation();
	AfxOleLockApp();
	mp_oControl = oControl;
	mp_oTimeLine = oTimeLine;
	mp_lLastVisibleRow = 0;
	mp_oGrid = new clsGrid(mp_oControl, mp_oTimeLine);
	Clear();
}

clsClientArea::clsClientArea()
{
	AfxMessageBox(_T("Invalid constructor has been invoked for clsClientArea. Please inform support@sourcecodestore.com of this error"), MB_OK | MB_ICONSTOP | MB_APPLMODAL);
}

clsClientArea::~clsClientArea()
{
	delete mp_oGrid;
	mp_oGrid = NULL;
	AfxOleUnlockApp();
}

void clsClientArea::OnFinalRelease()
{
	CCmdTarget::OnFinalRelease();
}

BOOL clsClientArea::GetDetectConflicts(void)
{
	return mp_bDetectConflicts;
}

void clsClientArea::SetDetectConflicts(BOOL Value)
{
	mp_bDetectConflicts = Value;
}

LONG clsClientArea::GetMilestoneSelectionOffset(void)
{
	return mp_lMilestoneSelectionOffset;
}

void clsClientArea::SetMilestoneSelectionOffset(LONG Value)
{
	mp_lMilestoneSelectionOffset = Value;
}

LONG clsClientArea::GetPredecessorSelectionOffset(void)
{
	return mp_lPredecessorSelectionOffset;
}

void clsClientArea::SetPredecessorSelectionOffset(LONG Value)
{
	mp_lPredecessorSelectionOffset = Value;
}

LONG clsClientArea::GetTaskBorderSelectionOffset(void)
{
	return mp_lTaskBorderSelectionOffset;
}

void clsClientArea::SetTaskBorderSelectionOffset(LONG Value)
{
	mp_lTaskBorderSelectionOffset = Value;
}

LONG clsClientArea::GetFirstVisibleRow(void)
{
	return mp_oControl->GetVerticalScrollBar()->GetValue();
}

void clsClientArea::SetFirstVisibleRow(LONG Value)
{
	mp_oControl->GetVerticalScrollBar()->SetValue(Value);
}

LONG clsClientArea::GetLastVisibleRow(void)
{
	return mp_lLastVisibleRow;
}

CString clsClientArea::GetToolTipFormat(void)
{
	return mp_sToolTipFormat;
}

void clsClientArea::SetToolTipFormat(CString Value)
{
	mp_sToolTipFormat = Value;
}

BOOL clsClientArea::GetToolTipsVisible(void)
{
	return mp_bToolTipsVisible;
}

void clsClientArea::SetToolTipsVisible(BOOL Value)
{
	mp_bToolTipsVisible = Value;
}

LONG clsClientArea::GetTop(void)
{
	if ((mp_oTimeLine->GetHeight() == 0)) 
	{
		return mp_oControl->Getmt_BorderThickness();
	}
	else 
	{
		return mp_oTimeLine->GetBottom() + 1;
	}
}

LONG clsClientArea::GetBottom(void)
{
	if (mp_oControl->f_TimeLineScrollBar()->GetState() == SS_SHOWN) 
	{
		return mp_oControl->clsG->Height() - mp_oControl->Getmt_BorderThickness() - 1 - mp_oControl->f_TimeLineScrollBar()->GetHeight();
	}
	else 
	{
		return mp_oControl->clsG->Height() - mp_oControl->Getmt_BorderThickness() - 1;
	}
}

LONG clsClientArea::GetLeft(void)
{
	return mp_oTimeLine->Getf_lStart();
}

LONG clsClientArea::GetRight(void)
{
	return mp_oTimeLine->Getf_lEnd();
}

LONG clsClientArea::GetWidth(void)
{
	return GetRight() - GetLeft();
}

LONG clsClientArea::GetHeight(void)
{
	return GetBottom() - GetTop();
}

clsGrid* clsClientArea::GetGrid(void)
{
	return mp_oGrid;
}

void clsClientArea::Setf_LastVisibleRow(LONG Value)
{
	mp_lLastVisibleRow = Value;
}

void clsClientArea::Draw(void)
{
	LONG lRowIndex = 0;
	clsRow* oRow = NULL;
	if (mp_oControl->GetRows()->GetCount() == 0) 
	{
		return;
	}
	mp_oControl->clsG->mp_ClipRegion(mp_oControl->GetSplitter()->GetRight(), mp_oControl->GetCurrentViewObject()->GetClientArea()->GetTop(), mp_oControl->Getmt_RightMargin(), mp_oControl->GetCurrentViewObject()->GetClientArea()->GetBottom(), TRUE);
	for (lRowIndex = mp_oControl->GetVerticalScrollBar()->GetValue(); lRowIndex <= mp_lLastVisibleRow; lRowIndex++) 
	{
		oRow = (clsRow*) mp_oControl->GetRows()->mp_oCollection->m_oReturnArrayElement(lRowIndex);
		mp_oControl->clsG->mp_DrawItemRow(mp_oControl->GetSplitter()->GetRight(), oRow->GetTop(), mp_oControl->Getmt_RightMargin(), oRow->GetBottom(), "", FALSE, NULL, 0, 0, oRow->GetClientAreaStyle(), oRow);
	}
}

CString clsClientArea::GetXML(void)
{
	clsXML oXML(mp_oControl, "ClientArea");
	oXML.InitializeWriter();
	oXML.WriteProperty("DetectConflicts", mp_bDetectConflicts);
	oXML.WriteProperty("MilestoneSelectionOffset", mp_lMilestoneSelectionOffset);
	oXML.WriteProperty("ToolTipFormat", mp_sToolTipFormat);
	oXML.WriteProperty("ToolTipsVisible", mp_bToolTipsVisible);
	oXML.WriteProperty("PredecessorSelectionOffset", mp_lPredecessorSelectionOffset);
	oXML.WriteProperty("TaskBorderSelectionOffset", mp_lTaskBorderSelectionOffset);
	oXML.WriteObject(mp_oGrid->GetXML());
	return oXML.GetXML();
}

void clsClientArea::SetXML(CString sXML)
{
	clsXML oXML(mp_oControl, "ClientArea");
	oXML.SetXML(sXML);
	oXML.InitializeReader();
	oXML.ReadProperty("DetectConflicts", mp_bDetectConflicts);
	oXML.ReadProperty("MilestoneSelectionOffset", mp_lMilestoneSelectionOffset);
	oXML.ReadProperty("ToolTipFormat", mp_sToolTipFormat);
	oXML.ReadProperty("ToolTipsVisible", mp_bToolTipsVisible);
	oXML.ReadProperty("PredecessorSelectionOffset", mp_lPredecessorSelectionOffset);
	oXML.ReadProperty("TaskBorderSelectionOffset", mp_lTaskBorderSelectionOffset);
	mp_oGrid->SetXML(oXML.ReadObject("Grid"));
}

void clsClientArea::Clear()
{
	mp_bDetectConflicts = TRUE;
	mp_lMilestoneSelectionOffset = 5;
	
	mp_sToolTipFormat = "ddddd";
	mp_bToolTipsVisible = TRUE;
	mp_lPredecessorSelectionOffset = 2;
	mp_lTaskBorderSelectionOffset = 2;
}

void clsClientArea::Clone(clsClientArea* oClone)
{
    oClone->SetDetectConflicts(mp_bDetectConflicts);
    oClone->SetMilestoneSelectionOffset(mp_lMilestoneSelectionOffset);
    oClone->SetToolTipFormat(mp_sToolTipFormat);
    oClone->SetToolTipsVisible(mp_bToolTipsVisible);
    oClone->SetPredecessorSelectionOffset(mp_lPredecessorSelectionOffset);
    oClone->SetTaskBorderSelectionOffset(mp_lTaskBorderSelectionOffset);
    GetGrid()->Clone(oClone->GetGrid());
}