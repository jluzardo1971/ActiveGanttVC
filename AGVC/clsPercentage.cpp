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
#include "clsPercentage.h"



IMPLEMENT_DYNCREATE(clsPercentage, clsItemBase)

//{1FF6240F-B3CC-4861-91E5-5A032166F149}
static const IID IID_IclsPercentage = { 0x1FF6240F, 0xB3CC, 0x4861, { 0x91, 0xE5, 0x5A, 0x03, 0x21, 0x66, 0xF1, 0x49} };

//{C70B902D-B6D7-4104-9D24-C543D210BECA}
IMPLEMENT_OLECREATE_FLAGS(clsPercentage, "AGVC.clsPercentage", afxRegApartmentThreading, 0xC70B902D, 0xB6D7, 0x4104, 0x9D, 0x24, 0xC5, 0x43, 0xD2, 0x10, 0xBE, 0xCA)

BEGIN_DISPATCH_MAP(clsPercentage, clsItemBase)
	DISP_PROPERTY_EX_ID(clsPercentage, "AllowSize", 1, odl_GetAllowSize, odl_SetAllowSize, VT_BOOL) 
	DISP_PROPERTY_EX_ID(clsPercentage, "Key", 2, odl_GetKey, odl_SetKey, VT_BSTR) 
	DISP_PROPERTY_EX_ID(clsPercentage, "Percent", 3, odl_GetPercent, odl_SetPercent, VT_R4) 
	DISP_PROPERTY_EX_ID(clsPercentage, "TaskKey", 4, odl_GetTaskKey, odl_SetTaskKey, VT_BSTR) 
	DISP_PROPERTY_EX_ID(clsPercentage, "Task", 5, odl_GetTask, SetNotSupported, VT_DISPATCH) 
	DISP_PROPERTY_EX_ID(clsPercentage, "Layer", 6, odl_GetLayer, SetNotSupported, VT_DISPATCH) 
	DISP_PROPERTY_EX_ID(clsPercentage, "Format", 7, odl_GetFormat, odl_SetFormat, VT_BSTR) 
	DISP_PROPERTY_EX_ID(clsPercentage, "StyleIndex", 8, odl_GetStyleIndex, odl_SetStyleIndex, VT_BSTR) 
	DISP_PROPERTY_EX_ID(clsPercentage, "Style", 9, odl_GetStyle, SetNotSupported, VT_DISPATCH) 
	DISP_PROPERTY_EX_ID(clsPercentage, "Tag", 10, odl_GetTag, odl_SetTag, VT_BSTR) 
	DISP_PROPERTY_EX_ID(clsPercentage, "LeftTrim", 11, odl_GetLeftTrim, SetNotSupported, VT_I4) 
	DISP_PROPERTY_EX_ID(clsPercentage, "RightTrim", 12, odl_GetRightTrim, SetNotSupported, VT_I4) 
	DISP_PROPERTY_EX_ID(clsPercentage, "Left", 13, odl_GetLeft, SetNotSupported, VT_I4) 
	DISP_PROPERTY_EX_ID(clsPercentage, "Top", 14, odl_GetTop, SetNotSupported, VT_I4) 
	DISP_PROPERTY_EX_ID(clsPercentage, "Bottom", 15, odl_GetBottom, SetNotSupported, VT_I4) 
	DISP_PROPERTY_EX_ID(clsPercentage, "Right", 16, odl_GetRight, SetNotSupported, VT_I4) 
	DISP_PROPERTY_EX_ID(clsPercentage, "Visible", 17, odl_GetVisible, odl_SetVisible, VT_BOOL) 
	DISP_FUNCTION_ID(clsPercentage, "GetXML", 18, odl_GetXML, VT_BSTR, VTS_NONE) 
	DISP_FUNCTION_ID(clsPercentage, "SetXML", 19, odl_SetXML, VT_EMPTY, VTS_BSTR)
	DISP_PROPERTY_EX_ID(clsPercentage, "Index", 20, odl_GetIndex, SetNotSupported, VT_I4)
	DISP_FUNCTION_ID(clsPercentage, "ToString", 21, odl_ToString, VT_BSTR, VTS_NONE) 
END_DISPATCH_MAP()

BEGIN_INTERFACE_MAP(clsPercentage, clsItemBase)
	INTERFACE_PART(clsPercentage, IID_IclsPercentage, Dispatch)
END_INTERFACE_MAP()

BEGIN_MESSAGE_MAP(clsPercentage, clsItemBase)
END_MESSAGE_MAP()

clsPercentage::clsPercentage()
{
	AfxMessageBox(_T("Invalid constructor has been invoked for clsPercentage. Please inform support@sourcecodestore.com of this error"), MB_OK | MB_ICONSTOP | MB_APPLMODAL);
}

clsPercentage::clsPercentage(CActiveGanttVCCtl* oControl)
{
	EnableAutomation();
	AfxOleLockApp();
	mp_oControl = oControl;
	mp_fPercent = 0;
	mp_sTaskKey = _T("");
	mp_sTag = _T("");
	mp_bVisible = TRUE;
	mp_bAllowSize = TRUE;
	mp_sFormat = _T("");
	mp_sStyleIndex = _T("DS_PERCENTAGE");
	mp_oStyle = mp_oControl->GetStyles()->FItem(_T("DS_PERCENTAGE"));
}

clsPercentage::~clsPercentage()
{
	AfxOleUnlockApp();
}

void clsPercentage::OnFinalRelease()
{
	clsItemBase::OnFinalRelease();
}

BOOL clsPercentage::GetAllowSize(void)
{
	return mp_bAllowSize;
}

void clsPercentage::SetAllowSize(BOOL Value)
{
	mp_bAllowSize = Value;
}

CString clsPercentage::GetKey(void)
{
	return mp_sKey;
}

void clsPercentage::SetKey(CString Value)
{
	mp_oControl->GetPercentages()->mp_oCollection->mp_SetKey(&mp_sKey, Value, PERCENTAGES_SET_KEY);
}

FLOAT clsPercentage::GetPercent(void)
{
	return mp_fPercent;
}

void clsPercentage::SetPercent(FLOAT Value)
{
	mp_fPercent = Value;
}

CString clsPercentage::GetTaskKey(void)
{
	return mp_sTaskKey;
}

void clsPercentage::SetTaskKey(CString Value)
{
	mp_sTaskKey = Value;
	mp_oTask = mp_oControl->GetTasks()->Item(Value);
}

clsTask* clsPercentage::GetTask(void)
{
	return mp_oTask;
}

clsLayer* clsPercentage::GetLayer(void)
{
	return mp_oTask->GetLayer();
}

CString clsPercentage::GetTag(void)
{
	return mp_sTag;
}

void clsPercentage::SetTag(CString Value)
{
	mp_sTag = Value;
}

LONG clsPercentage::GetLeftTrim(void)
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

LONG clsPercentage::GetRightTrim(void)
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

LONG clsPercentage::GetLeft(void)
{
	return mp_oTask->GetLeft();
}

LONG clsPercentage::GetTop(void)
{
	if ((mp_oTask->GetRow()->GetHeight() <= -1)) 
	{
		return mp_oTask->GetRow()->GetTop();
	}
	if (mp_oStyle->GetPlacement() == PLC_ROWEXTENTSPLACEMENT) 
	{
		return mp_oTask->GetRow()->GetTop();
	}
	if (mp_oStyle->GetPlacement() == PLC_OFFSETPLACEMENT) 
	{
		return mp_oTask->GetRow()->GetTop() + mp_oStyle->GetOffsetTop();
	}
	return 0;
}

LONG clsPercentage::GetBottom(void)
{
	if ((mp_oTask->GetRow()->GetHeight() <= -1)) 
	{
		return mp_oTask->GetRow()->GetTop();
	}
	if (mp_oStyle->GetPlacement() == PLC_ROWEXTENTSPLACEMENT) 
	{
		return mp_oTask->GetRow()->GetBottom() - 1;
	}
	if (mp_oStyle->GetPlacement() == PLC_OFFSETPLACEMENT) 
	{
		return mp_oTask->GetRow()->GetTop() + mp_oStyle->GetOffsetTop() + mp_oStyle->GetOffsetBottom();
	}
	return 0;
}

LONG clsPercentage::GetRight(void)
{
	LONG lLeft = GetLeft();
	LONG lRight = GetLeft() + CInt32((mp_oTask->GetRight() - mp_oTask->GetLeft()) * mp_fPercent);
	LONG lWidth = CInt32(((DOUBLE)mp_oTask->GetRight() - (DOUBLE)mp_oTask->GetLeft()) * (DOUBLE)mp_fPercent);
	return GetLeft() + CInt32((mp_oTask->GetRight() - mp_oTask->GetLeft()) * mp_fPercent);
}

BOOL clsPercentage::GetVisible(void)
{
	if (mp_oTask->GetRow()->GetVisible() == FALSE) 
	{
		return FALSE;
	}
	if (mp_oTask->GetVisible() == FALSE || mp_oTask->GetType() == OT_MILESTONE) 
	{
		return FALSE;
	}
	return mp_bVisible;
}

void clsPercentage::SetVisible(BOOL Value)
{
	mp_bVisible = Value;
}

LONG clsPercentage::GetIndex(void)
{
	return mp_lIndex;
}

void clsPercentage::SetIndex(LONG Value)
{
	mp_lIndex = Value;
}

BOOL clsPercentage::Getf_bLeftVisible(void)
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

BOOL clsPercentage::Getf_bRightVisible(void)
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

LONG clsPercentage::GetRightSel(void)
{
	if (GetRight() == GetLeft()) 
	{
		return GetLeft() + 15;
	}
	else 
	{
		return GetRight();
	}
}

CString clsPercentage::GetFormat(void)
{
	return mp_sFormat;
}

void clsPercentage::SetFormat(CString Value)
{
	mp_sFormat = Value;
}

CString clsPercentage::GetStyleIndex(void)
{
	if (mp_sStyleIndex == "DS_PERCENTAGE") 
	{
		return "";
	}
	else 
	{
		return mp_sStyleIndex;
	}
}

void clsPercentage::SetStyleIndex(CString Value)
{
	Value = Value.Trim();
	if (Value.GetLength() == 0) {Value = "DS_PERCENTAGE";} 
	mp_sStyleIndex = Value;
	mp_oStyle = mp_oControl->GetStyles()->FItem(Value);
}

clsStyle* clsPercentage::GetStyle(void)
{
	return mp_oStyle;
}

CString clsPercentage::ToString(void)
{
	return g_Format(mp_fPercent, mp_sFormat);
}

CString clsPercentage::GetXML(void)
{

	clsXML oXML(mp_oControl, "Percentage");
	oXML.InitializeWriter();
	oXML.WriteProperty("AllowSize", mp_bAllowSize);
	oXML.WriteProperty("Key", mp_sKey);
	oXML.WriteProperty("Percent", mp_fPercent);
	oXML.WriteProperty("Tag", mp_sTag);
	oXML.WriteProperty("TaskKey", mp_sTaskKey);
	oXML.WriteProperty("Visible", mp_bVisible);
	oXML.WriteProperty("StyleIndex", mp_sStyleIndex);
	oXML.WriteProperty("Format", mp_sFormat);
	return oXML.GetXML();
}

void clsPercentage::SetXML(CString sXML)
{
	clsXML oXML(mp_oControl, "Percentage");
	oXML.SetXML(sXML);
	oXML.InitializeReader();
	oXML.ReadProperty("AllowSize", mp_bAllowSize);
	oXML.ReadProperty("Key", mp_sKey);
	oXML.ReadProperty("Percent", mp_fPercent);
	oXML.ReadProperty("Tag", mp_sTag);
	oXML.ReadProperty("TaskKey", mp_sTaskKey);
	mp_oTask = mp_oControl->mp_oTasks->Item(mp_sTaskKey);
	oXML.ReadProperty("Visible", mp_bVisible);
	oXML.ReadProperty("StyleIndex", mp_sStyleIndex);
	SetStyleIndex(mp_sStyleIndex);
	oXML.ReadProperty("Format", mp_sFormat);
}