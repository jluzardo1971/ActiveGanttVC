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
#include "clsCell.h"



IMPLEMENT_DYNCREATE(clsCell, clsItemBase)

//{A2E1E15A-EC10-4970-A334-10E00C396BB5}
static const IID IID_IclsCell = { 0xA2E1E15A, 0xEC10, 0x4970, { 0xA3, 0x34, 0x10, 0xE0, 0x0C, 0x39, 0x6B, 0xB5} };

//{839E7680-CEEA-4F23-88EC-DCC7D2066A10}
IMPLEMENT_OLECREATE_FLAGS(clsCell, "AGVC.clsCell", afxRegApartmentThreading, 0x839E7680, 0xCEEA, 0x4F23, 0x88, 0xEC, 0xDC, 0xC7, 0xD2, 0x06, 0x6A, 0x10)

BEGIN_DISPATCH_MAP(clsCell, clsItemBase)
	DISP_PROPERTY_EX_ID(clsCell, "Row", 1, odl_GetRow, SetNotSupported, VT_DISPATCH) 
	DISP_PROPERTY_EX_ID(clsCell, "Key", 2, odl_GetKey, odl_SetKey, VT_BSTR) 
	DISP_PROPERTY_EX_ID(clsCell, "Text", 3, odl_GetText, odl_SetText, VT_BSTR) 
	DISP_PROPERTY_EX_ID(clsCell, "Image", 4, odl_GetImage, SetNotSupported, VT_DISPATCH) 
	DISP_PROPERTY_EX_ID(clsCell, "StyleIndex", 5, odl_GetStyleIndex, odl_SetStyleIndex, VT_BSTR) 
	DISP_PROPERTY_EX_ID(clsCell, "Style", 6, odl_GetStyle, SetNotSupported, VT_DISPATCH) 
	DISP_PROPERTY_EX_ID(clsCell, "Tag", 7, odl_GetTag, odl_SetTag, VT_BSTR) 
	DISP_PROPERTY_EX_ID(clsCell, "RowKey", 8, odl_GetRowKey, SetNotSupported, VT_BSTR) 
	DISP_PROPERTY_EX_ID(clsCell, "Left", 9, odl_GetLeft, SetNotSupported, VT_I4) 
	DISP_PROPERTY_EX_ID(clsCell, "Top", 10, odl_GetTop, SetNotSupported, VT_I4) 
	DISP_PROPERTY_EX_ID(clsCell, "Right", 11, odl_GetRight, SetNotSupported, VT_I4) 
	DISP_PROPERTY_EX_ID(clsCell, "Bottom", 12, odl_GetBottom, SetNotSupported, VT_I4) 
	DISP_PROPERTY_EX_ID(clsCell, "LeftTrim", 13, odl_GetLeftTrim, SetNotSupported, VT_I4) 
	DISP_PROPERTY_EX_ID(clsCell, "RightTrim", 14, odl_GetRightTrim, SetNotSupported, VT_I4) 
	DISP_FUNCTION_ID(clsCell, "GetXML", 15, odl_GetXML, VT_BSTR, VTS_NONE) 
	DISP_FUNCTION_ID(clsCell, "SetXML", 16, odl_SetXML, VT_EMPTY, VTS_BSTR)
	DISP_PROPERTY_EX_ID(clsCell, "Index", 17, odl_GetIndex, SetNotSupported, VT_I4)
	DISP_PROPERTY_EX_ID(clsCell, "ImageTag", 18, odl_GetImageTag, odl_SetImageTag, VT_BSTR)
	DISP_PROPERTY_EX_ID(clsCell, "AllowTextEdit", 19, odl_GetAllowTextEdit, odl_SetAllowTextEdit, VT_BOOL) 
END_DISPATCH_MAP()

BEGIN_INTERFACE_MAP(clsCell, clsItemBase)
	INTERFACE_PART(clsCell, IID_IclsCell, Dispatch)
END_INTERFACE_MAP()

BEGIN_MESSAGE_MAP(clsCell, clsItemBase)
END_MESSAGE_MAP()

clsCell::clsCell(CActiveGanttVCCtl* oControl, clsCells* oCells)
{
	EnableAutomation();
	AfxOleLockApp();
	mp_oControl = oControl;
	mp_sText = "";
	mp_sStyleIndex = "DS_CELL";
	mp_oStyle = mp_oControl->GetStyles()->FItem("DS_CELL");
	mp_sTag = "";
	mp_oCells = oCells;
	mp_oImage = new clsImage();
	mp_sImageTag = _T("");
	mp_bAllowTextEdit = FALSE;
}

clsCell::clsCell()
{
	AfxMessageBox(_T("Invalid constructor has been invoked for clsCell. Please inform support@sourcecodestore.com of this error"), MB_OK | MB_ICONSTOP | MB_APPLMODAL);
}

clsCell::~clsCell()
{
	delete mp_oImage;
	mp_oImage = NULL;
	AfxOleUnlockApp();
}

void clsCell::OnFinalRelease()
{
	clsItemBase::OnFinalRelease();
}



BOOL clsCell::GetAllowTextEdit(void)
{
	return mp_bAllowTextEdit;
}

void clsCell::SetAllowTextEdit(BOOL Value)
{
	mp_bAllowTextEdit = Value;
}



clsRow* clsCell::GetRow(void)
{
	return mp_oRow;
}

CString clsCell::GetKey(void)
{
	return mp_sKey;
}

void clsCell::SetKey(CString Value)
{
	mp_oCells->mp_oCollection->mp_SetKey(&mp_sKey, Value, CELLS_SET_KEY);
}

CString clsCell::GetText(void)
{
	return mp_sText;
}

void clsCell::SetText(CString Value)
{
	mp_sText = Value;
}

clsImage* clsCell::GetImage(void)
{
	return mp_oImage;
}

CString clsCell::GetStyleIndex(void)
{
	if (mp_sStyleIndex == "DS_CELL") 
	{
		return "";
	}
	else 
	{
		return mp_sStyleIndex;
	}
}

void clsCell::SetStyleIndex(CString Value)
{
	Value = Value.Trim();
	if (Value.GetLength() == 0) {Value = "DS_CELL";}
	mp_sStyleIndex = Value;
	mp_oStyle = mp_oControl->GetStyles()->FItem(Value);
}

clsStyle* clsCell::GetStyle(void)
{
	return mp_oStyle;
}

CString clsCell::GetTag(void)
{
	return mp_sTag;
}

void clsCell::SetTag(CString Value)
{
	mp_sTag = Value;
}

CString clsCell::GetRowKey(void)
{
	return mp_oCells->Row()->GetKey();
}

LONG clsCell::GetLeft(void)
{
	return mp_oControl->GetColumns()->Item(CStr(mp_lIndex))->GetLeft();
}

LONG clsCell::GetTop(void)
{
	return mp_oCells->Row()->GetTop();
}

LONG clsCell::GetRight(void)
{
	return mp_oControl->GetColumns()->Item(CStr(mp_lIndex))->GetRight();
}

LONG clsCell::GetBottom(void)
{
	return mp_oCells->Row()->GetBottom();
}

LONG clsCell::GetLeftTrim(void)
{
	return mp_oControl->GetColumns()->Item(CStr(mp_lIndex))->GetLeftTrim();
}

LONG clsCell::GetRightTrim(void)
{
	return mp_oControl->GetColumns()->Item(CStr(mp_lIndex))->GetRightTrim();
}

LONG clsCell::GetIndex(void)
{
	return mp_lIndex;
}

void clsCell::SetIndex(LONG Value)
{
	mp_lIndex = Value;
}

CString clsCell::GetImageTag(void)
{
	return mp_sImageTag;
}

void clsCell::SetImageTag(CString Value)
{
	mp_sImageTag = Value;
}



CString clsCell::GetXML(void)
{
	clsXML oXML(mp_oControl, "Cell");
	oXML.InitializeWriter();
	oXML.WriteProperty("Image", *mp_oImage);
	oXML.WriteProperty("Key", mp_sKey);
	oXML.WriteProperty("StyleIndex", mp_sStyleIndex);
	oXML.WriteProperty("Tag", mp_sTag);
	oXML.WriteProperty("Text", mp_sText);
	oXML.WriteProperty("ImageTag", mp_sImageTag);
	oXML.WriteProperty("AllowTextEdit", mp_bAllowTextEdit);
	return oXML.GetXML();
}

void clsCell::SetXML(CString sXML)
{
	clsXML oXML(mp_oControl, "Cell");
	oXML.SetXML(sXML);
	oXML.InitializeReader();
	oXML.ReadProperty("Image", *mp_oImage);
	oXML.ReadProperty("Key", mp_sKey);
	oXML.ReadProperty("StyleIndex", mp_sStyleIndex);
	SetStyleIndex(mp_sStyleIndex);
	oXML.ReadProperty("Tag", mp_sTag);
	oXML.ReadProperty("Text", mp_sText);
	oXML.ReadProperty("ImageTag", mp_sImageTag);
	oXML.ReadProperty("AllowTextEdit", mp_bAllowTextEdit);
}

IDispatch* clsCell::odl_GetRow(void) //Circular Reference
{
	
	return GetRow()->GetIDispatch(TRUE);
}