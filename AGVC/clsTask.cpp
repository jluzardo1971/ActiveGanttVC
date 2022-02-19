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
#include "clsTask.h"



IMPLEMENT_DYNCREATE(clsTask, clsItemBase)

//{7498F106-3C16-4971-91D0-54984352DC42}
static const IID IID_IclsTask = { 0x7498F106, 0x3C16, 0x4971, { 0x91, 0xD0, 0x54, 0x98, 0x43, 0x52, 0xDC, 0x42} };

//{EEAECDE1-B760-454D-8942-6B56DC7E1877}
IMPLEMENT_OLECREATE_FLAGS(clsTask, "AGVC.clsTask", afxRegApartmentThreading, 0xEEAECDE1, 0xB760, 0x454D, 0x89, 0x42, 0x6B, 0x56, 0xDC, 0x7E, 0x18, 0x77)

BEGIN_DISPATCH_MAP(clsTask, clsItemBase)
	DISP_PROPERTY_EX_ID(clsTask, "Key", 1, odl_GetKey, odl_SetKey, VT_BSTR) 
	DISP_PROPERTY_EX_ID(clsTask, "IncomingPredecessors", 2, odl_GetIncomingPredecessors, odl_SetIncomingPredecessors, VT_BOOL) 
	DISP_PROPERTY_EX_ID(clsTask, "OutgoingPredecessors", 3, odl_GetOutgoingPredecessors, odl_SetOutgoingPredecessors, VT_BOOL) 
	DISP_PROPERTY_EX_ID(clsTask, "AllowStretchLeft", 4, odl_GetAllowStretchLeft, odl_SetAllowStretchLeft, VT_BOOL) 
	DISP_PROPERTY_EX_ID(clsTask, "AllowStretchRight", 5, odl_GetAllowStretchRight, odl_SetAllowStretchRight, VT_BOOL) 
	DISP_PROPERTY_EX_ID(clsTask, "Text", 6, odl_GetText, odl_SetText, VT_BSTR) 
	DISP_PROPERTY_EX_ID(clsTask, "LayerIndex", 7, odl_GetLayerIndex, odl_SetLayerIndex, VT_BSTR) 
	DISP_PROPERTY_EX_ID(clsTask, "Layer", 8, odl_GetLayer, SetNotSupported, VT_DISPATCH) 
	DISP_PROPERTY_EX_ID(clsTask, "Image", 9, odl_GetImage, SetNotSupported, VT_DISPATCH) 
	DISP_PROPERTY_EX_ID(clsTask, "RowKey", 10, odl_GetRowKey, odl_SetRowKey, VT_BSTR) 
	DISP_PROPERTY_EX_ID(clsTask, "Row", 11, odl_GetRow, SetNotSupported, VT_DISPATCH) 
	DISP_PROPERTY_EX_ID(clsTask, "StyleIndex", 12, odl_GetStyleIndex, odl_SetStyleIndex, VT_BSTR) 
	DISP_PROPERTY_EX_ID(clsTask, "Style", 13, odl_GetStyle, SetNotSupported, VT_DISPATCH) 
	DISP_PROPERTY_EX_ID(clsTask, "Tag", 14, odl_GetTag, odl_SetTag, VT_BSTR) 
	DISP_PROPERTY_EX_ID(clsTask, "AllowedMovement", 15, odl_GetAllowedMovement, odl_SetAllowedMovement, VT_I4) 
	DISP_PROPERTY_EX_ID(clsTask, "LeftTrim", 16, odl_GetLeftTrim, SetNotSupported, VT_I4) 
	DISP_PROPERTY_EX_ID(clsTask, "RightTrim", 17, odl_GetRightTrim, SetNotSupported, VT_I4) 
	DISP_PROPERTY_EX_ID(clsTask, "StartDate", 18, odl_GetStartDate, odl_SetStartDate, VT_DATE) 
	DISP_PROPERTY_EX_ID(clsTask, "Left", 19, odl_GetLeft, SetNotSupported, VT_I4) 
	DISP_PROPERTY_EX_ID(clsTask, "Right", 20, odl_GetRight, SetNotSupported, VT_I4) 
	DISP_PROPERTY_EX_ID(clsTask, "EndDate", 21, odl_GetEndDate, odl_SetEndDate, VT_DATE) 
	DISP_PROPERTY_EX_ID(clsTask, "Top", 22, odl_GetTop, SetNotSupported, VT_I4) 
	DISP_PROPERTY_EX_ID(clsTask, "Bottom", 23, odl_GetBottom, SetNotSupported, VT_I4) 
	DISP_PROPERTY_EX_ID(clsTask, "Visible", 24, odl_GetVisible, odl_SetVisible, VT_BOOL) 
	DISP_PROPERTY_EX_ID(clsTask, "Type", 25, odl_GetType, SetNotSupported, VT_I4) 
	DISP_PROPERTY_EX_ID(clsTask, "Index", 26, odl_GetIndex, SetNotSupported, VT_I4)
	DISP_PROPERTY_EX_ID(clsTask, "ImageTag", 27, odl_GetImageTag, odl_SetImageTag, VT_BSTR)
	DISP_PROPERTY_EX_ID(clsTask, "AllowTextEdit", 28, odl_GetAllowTextEdit, odl_SetAllowTextEdit, VT_BOOL)
	DISP_PROPERTY_EX_ID(clsTask, "Warning", 29, odl_GetWarning, SetNotSupported, VT_BOOL)
	DISP_PROPERTY_EX_ID(clsTask, "WarningStyleIndex", 30, odl_GetWarningStyleIndex, odl_SetWarningStyleIndex, VT_BSTR) 
	DISP_PROPERTY_EX_ID(clsTask, "WarningStyle", 31, odl_GetWarningStyle, SetNotSupported, VT_DISPATCH)
	DISP_PROPERTY_EX_ID(clsTask, "TaskType", 32, odl_GetTaskType, odl_SetTaskType, VT_I4)
	DISP_PROPERTY_EX_ID(clsTask, "DurationInterval", 33, odl_GetDurationInterval, odl_SetDurationInterval, VT_I4)
	DISP_PROPERTY_EX_ID(clsTask, "DurationFactor", 34, odl_GetDurationFactor, odl_SetDurationFactor, VT_I4)

	DISP_FUNCTION_ID(clsTask, "InConflict", 35, odl_InConflict, VT_BOOL, VTS_NONE) 
	DISP_FUNCTION_ID(clsTask, "GetXML", 36, odl_GetXML, VT_BSTR, VTS_NONE) 
	DISP_FUNCTION_ID(clsTask, "SetXML", 37, odl_SetXML, VT_EMPTY, VTS_BSTR)
END_DISPATCH_MAP()

BEGIN_INTERFACE_MAP(clsTask, clsItemBase)
	INTERFACE_PART(clsTask, IID_IclsTask, Dispatch)
END_INTERFACE_MAP()

BEGIN_MESSAGE_MAP(clsTask, clsItemBase)
END_MESSAGE_MAP()

clsTask::clsTask()
{
	AfxMessageBox(_T("Invalid constructor has been invoked for clsTask. Please inform support@sourcecodestore.com of this error"), MB_OK | MB_ICONSTOP | MB_APPLMODAL);
}

clsTask::clsTask(CActiveGanttVCCtl* oControl)
{
	EnableAutomation();
	AfxOleLockApp();
	mp_oControl = oControl;
	mp_bIncomingPredecessors = TRUE;
	mp_bOutgoingPredecessors = TRUE;
	mp_bAllowStretchLeft = TRUE;
	mp_bAllowStretchRight = TRUE;
	mp_dtEndDate = (DATE)0;
	mp_dtStartDate = (DATE)0;
	mp_sText = "";
	mp_sLayerIndex = "0";
	mp_oLayer = mp_oControl->GetLayers()->FItem("0");
	mp_sRowKey = "";
	mp_sStyleIndex = "DS_TASK";
	mp_oStyle = mp_oControl->GetStyles()->FItem("DS_TASK");
	mp_sTag = "";
	mp_yAllowedMovement = MT_UNRESTRICTED;
	mp_bVisible = TRUE;
	mp_oImage = new clsImage();
	mp_sImageTag = _T("");
	mp_bAllowTextEdit = FALSE;
	mp_bWarning = FALSE;
	mp_sWarningStyleIndex = _T("");
	mp_yTaskType = TT_START_END;
	mp_yDurationInterval = IL_HOUR;
	mp_lDurationFactor = 0;
}

clsTask::~clsTask()
{
	delete mp_oImage;
	mp_oImage = NULL;
	AfxOleUnlockApp();
}

void clsTask::OnFinalRelease()
{
	clsItemBase::OnFinalRelease();
}

BOOL clsTask::GetAllowTextEdit(void)
{
	return mp_bAllowTextEdit;
}

void clsTask::SetAllowTextEdit(BOOL Value)
{
	mp_bAllowTextEdit = Value;
}


CString clsTask::GetKey(void)
{
	return mp_sKey;
}

void clsTask::SetKey(CString Value)
{
	mp_oControl->GetTasks()->mp_oCollection->mp_SetKey(&mp_sKey, Value, TASKS_SET_KEY);
}

BOOL clsTask::GetIncomingPredecessors(void)
{
	return mp_bIncomingPredecessors;
}

void clsTask::SetIncomingPredecessors(BOOL Value)
{
	mp_bIncomingPredecessors = Value;
}

BOOL clsTask::GetOutgoingPredecessors(void)
{
	return mp_bOutgoingPredecessors;
}

void clsTask::SetOutgoingPredecessors(BOOL Value)
{
	mp_bOutgoingPredecessors = Value;
}

BOOL clsTask::GetAllowStretchLeft(void)
{
	return mp_bAllowStretchLeft;
}

void clsTask::SetAllowStretchLeft(BOOL Value)
{
	mp_bAllowStretchLeft = Value;
}

BOOL clsTask::GetAllowStretchRight(void)
{
	return mp_bAllowStretchRight;
}

void clsTask::SetAllowStretchRight(BOOL Value)
{
	mp_bAllowStretchRight = Value;
}

CString clsTask::GetText(void)
{
	return mp_sText;
}

void clsTask::SetText(CString Value)
{
	mp_sText = Value;
}

CString clsTask::GetLayerIndex(void)
{
	return mp_sLayerIndex;
}

void clsTask::SetLayerIndex(CString Value)
{
	if (Value.Trim() == "") 
	{
		Value = "0";
	}
	mp_sLayerIndex = Value;
	mp_oLayer = mp_oControl->GetLayers()->FItem(Value);
}

clsLayer* clsTask::GetLayer(void)
{
	return mp_oLayer;
}

clsImage* clsTask::GetImage(void)
{
	return mp_oImage;
}

CString clsTask::GetRowKey(void)
{
	return mp_sRowKey;
}

void clsTask::SetRowKey(CString Value)
{
	if (mp_oControl->GetRows()->mp_oCollection->m_bDoesKeyExist(Value) == FALSE) 
	{
		mp_oControl->mp_ErrorReport(INVALID_ROW_KEY, _T("Invalid Row Key (\"")  + Value + _T("\")"), _T("ActiveGanttVCCtl.clsTask.Let RowKey"));
		return;
	}
	mp_sRowKey = Value;
	mp_oRow = mp_oControl->GetRows()->Item(mp_sRowKey);
}

clsRow* clsTask::GetRow(void)
{
	return mp_oRow;
}

CString clsTask::GetStyleIndex(void)
{
	if (mp_sStyleIndex == "DS_TASK") 
	{
		return "";
	}
	else 
	{
		return mp_sStyleIndex;
	}
}

void clsTask::SetStyleIndex(CString Value)
{
	Value = Value.Trim();
	if (Value.GetLength() == 0) {Value = "DS_TASK";} 
	mp_sStyleIndex = Value;
	mp_oStyle = mp_oControl->GetStyles()->FItem(Value);
}

clsStyle* clsTask::GetStyle(void)
{
	return mp_oStyle;
}

CString clsTask::GetTag(void)
{
	return mp_sTag;
}

void clsTask::SetTag(CString Value)
{
	mp_sTag = Value;
}

LONG clsTask::GetAllowedMovement(void)
{
	return mp_yAllowedMovement;
}

void clsTask::SetAllowedMovement(LONG Value)
{
	mp_yAllowedMovement = Value;
}

LONG clsTask::GetLeftTrim(void)
{

	if (GetLeft() < mp_oControl->GetSplitter()->GetRight()) 
	{
		return mp_oControl->GetSplitter()->GetRight();
	}
	else 
	{
		return GetLeft();
	}
}

LONG clsTask::GetRightTrim(void)
{
	if (GetRight() > mp_oControl->Getmt_RightMargin()) 
	{
		return mp_oControl->Getmt_RightMargin();
	}
	else 
	{
		return GetRight();
	}
}

COleDateTime clsTask::GetStartDate(void)
{
	return mp_dtStartDate;
}

void clsTask::SetStartDate(COleDateTime Value)
{
	mp_dtStartDate = Value;
    if (mp_yTaskType == TT_DURATION && mp_lDurationFactor != 0)
    {
        mp_GetDuration();
    }
}

LONG clsTask::GetLeft(void)
{
	if (mp_dtStartDate == mp_dtEndDate) 
	{
		return mp_oControl->GetMathLib()->GetXCoordinateFromDate(GetStartDate()) - mp_oControl->GetCurrentViewObject()->GetClientArea()->GetMilestoneSelectionOffset();
	}
	else 
	{
		return mp_oControl->GetMathLib()->GetXCoordinateFromDate(GetStartDate());
	}
}

LONG clsTask::GetRight(void)
{
	if (mp_dtStartDate == mp_dtEndDate) 
	{
		return mp_oControl->GetMathLib()->GetXCoordinateFromDate(GetEndDate()) + mp_oControl->GetCurrentViewObject()->GetClientArea()->GetMilestoneSelectionOffset();
	}
	else 
	{
		return mp_oControl->GetMathLib()->GetXCoordinateFromDate(GetEndDate());
	}
}

COleDateTime clsTask::GetEndDate(void)
{
	return mp_dtEndDate;
}

void clsTask::SetEndDate(COleDateTime Value)
{
    if (mp_yTaskType == TT_START_END)
    {
        mp_dtEndDate = Value;
    }
}

LONG clsTask::GetTop(void)
{
	if ((mp_oRow->GetHeight() <= -1)) 
	{
		return mp_oRow->GetTop();
	}
	if (mp_oStyle->GetPlacement() == PLC_ROWEXTENTSPLACEMENT) 
	{
		return mp_oRow->GetTop();
	}
	if (mp_oStyle->GetPlacement() == PLC_OFFSETPLACEMENT) 
	{
		return mp_oRow->GetTop() + mp_oStyle->GetOffsetTop();
	}
	return 0;
}

LONG clsTask::GetBottom(void)
{
	if ((mp_oRow->GetHeight() <= -1)) 
	{
		return mp_oRow->GetTop();
	}
	if (mp_oStyle->GetPlacement() == PLC_ROWEXTENTSPLACEMENT) 
	{
		return mp_oRow->GetBottom() - 1;
	}
	if (mp_oStyle->GetPlacement() == PLC_OFFSETPLACEMENT) 
	{
		return mp_oRow->GetTop() + mp_oStyle->GetOffsetTop() + mp_oStyle->GetOffsetBottom();
	}
	return 0;
}

BOOL clsTask::GetVisible(void)
{
	if (mp_oLayer->GetVisible() == FALSE) 
	{
		return FALSE;
	}
	if ((mp_oRow->GetHeight() <= -1)) 
	{
		return FALSE;
	}
	if (mp_oRow->GetVisible() == FALSE) 
	{
		return FALSE;
	}
	return mp_bVisible;
}

void clsTask::SetVisible(BOOL Value)
{
	mp_bVisible = Value;
}

BOOL clsTask::GetInsideVisibleTimeLineArea(void)
{
	if (GetStartDate() > mp_oControl->GetCurrentViewObject()->GetTimeLine()->GetEndDate())
	{
		return FALSE;
	}
	if (GetEndDate() < mp_oControl->GetCurrentViewObject()->GetTimeLine()->GetStartDate())
	{
		return FALSE;
	}
	return TRUE;
}

E_CLIENTAREAVISIBILITY clsTask::GetClientAreaVisiblity(void)
{
    if (GetStartDate() > mp_oControl->GetCurrentViewObject()->GetTimeLine()->GetEndDate())
	{
        return VS_RIGHTOFVISIBLEAREA;
	}
    if (GetEndDate() < mp_oControl->GetCurrentViewObject()->GetTimeLine()->GetStartDate())
	{
        return VS_LEFTOFVISIBLEAREA;
	}
    return VS_INSIDEVISIBLEAREA;
}

LONG clsTask::GetType(void)
{
	if (mp_yTaskType == TT_DURATION && mp_lDurationFactor == 0)
	{
		return OT_MILESTONE;
	}
	if (GetStartDate() == GetEndDate()) 
	{
		return OT_MILESTONE;
	}
	else 
	{
		return OT_TASK;
	}
}

LONG clsTask::GetIndex(void)
{
	return mp_lIndex;
}

void clsTask::SetIndex(LONG Value)
{
	mp_lIndex = Value;
}

BOOL clsTask::Getf_bLeftVisible(void)
{
	if (GetLeftTrim() == GetLeft()) 
	{
		return TRUE;
	}
	else 
	{
		return FALSE;
	}
}

BOOL clsTask::Getf_bRightVisible(void)
{
	if (GetRightTrim() == GetRight()) 
	{
		return TRUE;
	}
	else 
	{
		return FALSE;
	}
}

CString clsTask::GetImageTag(void)
{
	return mp_sImageTag;
}

void clsTask::SetImageTag(CString Value)
{
	mp_sImageTag = Value;
}



BOOL clsTask::InConflict(void)
{
	return mp_oControl->GetMathLib()->DetectConflict(GetStartDate(), GetEndDate(), mp_sRowKey, GetIndex(), mp_sLayerIndex);
}



BOOL clsTask::GetWarning(void)
{
    if (mp_oControl->GetPredecessors()->GetCount() == 0)
    {
        return FALSE;
    }
    else
    {
        return mp_bWarning;
    }
}



CString clsTask::GetWarningStyleIndex(void)
{
	return mp_sWarningStyleIndex;
}

void clsTask::SetWarningStyleIndex(CString Value)
{
	Value = Value.Trim();
    mp_sWarningStyleIndex = Value;
    if (Value.GetLength() > 0)
    {
        mp_oWarningStyle = mp_oControl->GetStyles()->FItem(Value);
    }
    else
    {
        mp_oWarningStyle = NULL;
    }
}

clsStyle* clsTask::GetWarningStyle(void)
{
    if (mp_sWarningStyleIndex.GetLength() == 0)
    {
        return mp_oStyle;
    }
    else
    {
        return mp_oWarningStyle;
    }
}



LONG clsTask::GetTaskType(void)
{
	return mp_yTaskType;
}

void clsTask::SetTaskType(LONG Value)
{
	mp_yTaskType = Value;
	if (mp_yTaskType == TT_DURATION) 
	{
		mp_GetDuration();
	}
}

LONG clsTask::GetDurationInterval(void)
{
	return mp_yDurationInterval;
}

void clsTask::SetDurationInterval(LONG Value)
{
	if (!(Value == IL_SECOND || Value == IL_MINUTE || Value == IL_HOUR || Value == IL_DAY)) 
	{
		mp_oControl->mp_ErrorReport(INVALID_DURATION_INTERVAL, _T("Interval is invalid for a duration"), _T("clsTask.Set_DurationInterval"));
		return;
	}
	mp_yDurationInterval = Value;
	if (mp_yTaskType == TT_DURATION) 
	{
		mp_GetDuration();
	}
}



LONG clsTask::GetDurationFactor(void)
{
	return mp_lDurationFactor;
}

void clsTask::SetDurationFactor(LONG Value)
{
	if (Value < 0) 
	{
		Value = Value * -1;
	}
	mp_lDurationFactor = Value;
	if (mp_yTaskType == TT_DURATION) 
	{
		mp_GetDuration();
	}
}

void clsTask::mp_GetDuration(void)
{
	mp_dtEndDate = mp_oControl->GetMathLib()->GetEndDate(mp_dtStartDate, mp_yDurationInterval, mp_lDurationFactor);
}

CString clsTask::GetXML(void)
{
	clsXML oXML(mp_oControl, "Task");
	oXML.InitializeWriter();
	oXML.WriteProperty("AllowedMovement", mp_yAllowedMovement);
	oXML.WriteProperty("AllowStretchLeft", mp_bAllowStretchLeft);
	oXML.WriteProperty("AllowStretchRight", mp_bAllowStretchRight);
	oXML.WriteProperty("EndDate", mp_dtEndDate);
	oXML.WriteProperty("Image", *mp_oImage);
	oXML.WriteProperty("IncomingPredecessors", mp_bIncomingPredecessors);
	oXML.WriteProperty("Key", mp_sKey);
	oXML.WriteProperty("LayerIndex", mp_sLayerIndex);
	oXML.WriteProperty("OutgoingPredecessors", mp_bOutgoingPredecessors);
	oXML.WriteProperty("RowKey", mp_sRowKey);
	oXML.WriteProperty("StartDate", mp_dtStartDate);
	oXML.WriteProperty("StyleIndex", mp_sStyleIndex);
	oXML.WriteProperty("Tag", mp_sTag);
	oXML.WriteProperty("Text", mp_sText);
	oXML.WriteProperty("Visible", mp_bVisible);
	oXML.WriteProperty("ImageTag", mp_sImageTag);
	oXML.WriteProperty("AllowTextEdit", mp_bAllowTextEdit);
	oXML.WriteProperty("WarningStyleIndex", mp_sWarningStyleIndex);
	oXML.WriteProperty("TaskType", mp_yTaskType);
	oXML.WriteProperty("DurationInterval", mp_yDurationInterval);
	oXML.WriteProperty("DurationFactor", mp_lDurationFactor);
	return oXML.GetXML();
}

void clsTask::SetXML(CString sXML)
{
	clsXML oXML(mp_oControl, "Task");
	oXML.SetXML(sXML);
	oXML.InitializeReader();
	oXML.ReadProperty("AllowedMovement", mp_yAllowedMovement);
	oXML.ReadProperty("AllowStretchLeft", mp_bAllowStretchLeft);
	oXML.ReadProperty("AllowStretchRight", mp_bAllowStretchRight);
	oXML.ReadProperty("EndDate", mp_dtEndDate);
	oXML.ReadProperty("Image", *mp_oImage);
	oXML.ReadProperty("IncomingPredecessors", mp_bIncomingPredecessors);
	oXML.ReadProperty("Key", mp_sKey);
	oXML.ReadProperty("LayerIndex", mp_sLayerIndex);
	mp_oLayer = mp_oControl->mp_oLayers->FItem(mp_sLayerIndex);
	oXML.ReadProperty("OutgoingPredecessors", mp_bOutgoingPredecessors);
	oXML.ReadProperty("RowKey", mp_sRowKey);
	mp_oRow = mp_oControl->mp_oRows->Item(mp_sRowKey);
	oXML.ReadProperty("StartDate", mp_dtStartDate);
	oXML.ReadProperty("StyleIndex", mp_sStyleIndex);
	SetStyleIndex(mp_sStyleIndex);
	oXML.ReadProperty("Tag", mp_sTag);
	oXML.ReadProperty("Text", mp_sText);
	oXML.ReadProperty("Visible", mp_bVisible);
	oXML.ReadProperty("ImageTag", mp_sImageTag);
	oXML.ReadProperty("AllowTextEdit", mp_bAllowTextEdit);
    oXML.ReadProperty("WarningStyleIndex", mp_sWarningStyleIndex);
    SetWarningStyleIndex(mp_sWarningStyleIndex);
    oXML.ReadProperty("TaskType", mp_yTaskType);
    oXML.ReadProperty("DurationInterval", mp_yDurationInterval);
    oXML.ReadProperty("DurationFactor", mp_lDurationFactor);
}