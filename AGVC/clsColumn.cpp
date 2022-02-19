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
#include "clsColumn.h"



IMPLEMENT_DYNCREATE(clsColumn, clsItemBase)

//{FFBDEC51-A8AF-4287-AB4F-B4BCC7D2C894}
static const IID IID_IclsColumn = { 0xFFBDEC51, 0xA8AF, 0x4287, { 0xAB, 0x4F, 0xB4, 0xBC, 0xC7, 0xD2, 0xC8, 0x94} };

//{D35C6591-D305-488C-B5F2-FAABA93D80FD}
IMPLEMENT_OLECREATE_FLAGS(clsColumn, "AGVC.clsColumn", afxRegApartmentThreading, 0xD35C6591, 0xD305, 0x488C, 0xB5, 0xF2, 0xFA, 0xAB, 0xA9, 0x3D, 0x80, 0xFD)

BEGIN_DISPATCH_MAP(clsColumn, clsItemBase)
	DISP_PROPERTY_EX_ID(clsColumn, "AllowMove", 1, odl_GetAllowMove, odl_SetAllowMove, VT_BOOL) 
	DISP_PROPERTY_EX_ID(clsColumn, "AllowSize", 2, odl_GetAllowSize, odl_SetAllowSize, VT_BOOL) 
	DISP_PROPERTY_EX_ID(clsColumn, "Key", 3, odl_GetKey, odl_SetKey, VT_BSTR) 
	DISP_PROPERTY_EX_ID(clsColumn, "Width", 4, odl_GetWidth, odl_SetWidth, VT_I4) 
	DISP_PROPERTY_EX_ID(clsColumn, "Text", 5, odl_GetText, odl_SetText, VT_BSTR) 
	DISP_PROPERTY_EX_ID(clsColumn, "Image", 6, odl_GetImage, SetNotSupported, VT_DISPATCH) 
	DISP_PROPERTY_EX_ID(clsColumn, "StyleIndex", 7, odl_GetStyleIndex, odl_SetStyleIndex, VT_BSTR) 
	DISP_PROPERTY_EX_ID(clsColumn, "Style", 8, odl_GetStyle, SetNotSupported, VT_DISPATCH) 
	DISP_PROPERTY_EX_ID(clsColumn, "Tag", 9, odl_GetTag, odl_SetTag, VT_BSTR) 
	DISP_PROPERTY_EX_ID(clsColumn, "LeftTrim", 10, odl_GetLeftTrim, SetNotSupported, VT_I4) 
	DISP_PROPERTY_EX_ID(clsColumn, "RightTrim", 11, odl_GetRightTrim, SetNotSupported, VT_I4) 
	DISP_PROPERTY_EX_ID(clsColumn, "Left", 12, odl_GetLeft, SetNotSupported, VT_I4) 
	DISP_PROPERTY_EX_ID(clsColumn, "Top", 13, odl_GetTop, SetNotSupported, VT_I4) 
	DISP_PROPERTY_EX_ID(clsColumn, "Right", 14, odl_GetRight, SetNotSupported, VT_I4) 
	DISP_PROPERTY_EX_ID(clsColumn, "Bottom", 15, odl_GetBottom, SetNotSupported, VT_I4) 
	DISP_PROPERTY_EX_ID(clsColumn, "Visible", 16, odl_GetVisible, SetNotSupported, VT_BOOL) 
	DISP_FUNCTION_ID(clsColumn, "GetXML", 17, odl_GetXML, VT_BSTR, VTS_NONE) 
	DISP_FUNCTION_ID(clsColumn, "SetXML", 18, odl_SetXML, VT_EMPTY, VTS_BSTR)
	DISP_PROPERTY_EX_ID(clsColumn, "Index", 19, odl_GetIndex, SetNotSupported, VT_I4)
	DISP_PROPERTY_EX_ID(clsColumn, "ImageTag", 20, odl_GetImageTag, odl_SetImageTag, VT_BSTR)
	DISP_PROPERTY_EX_ID(clsColumn, "AllowTextEdit", 21, odl_GetAllowTextEdit, odl_SetAllowTextEdit, VT_BOOL) 
END_DISPATCH_MAP()

BEGIN_INTERFACE_MAP(clsColumn, clsItemBase)
	INTERFACE_PART(clsColumn, IID_IclsColumn, Dispatch)
END_INTERFACE_MAP()

BEGIN_MESSAGE_MAP(clsColumn, clsItemBase)
END_MESSAGE_MAP()

clsColumn::clsColumn()
{
	AfxMessageBox(_T("Invalid constructor has been invoked for clsColumn. Please inform support@sourcecodestore.com of this error"), MB_OK | MB_ICONSTOP | MB_APPLMODAL);
}

clsColumn::clsColumn(CActiveGanttVCCtl* oControl)
{
	EnableAutomation();
	AfxOleLockApp();
	mp_oControl = oControl;
	mp_bAllowMove = TRUE;
	mp_bAllowSize = TRUE;
	mp_lWidth = 125;
	mp_sText = "";
	mp_sStyleIndex = "DS_COLUMN";
	mp_oStyle = mp_oControl->GetStyles()->FItem("DS_COLUMN");
	mp_sTag = "";
	mp_lLeft = 0;
	mp_lRight = 0;
	mp_bVisible = FALSE;
	mp_oImage = new clsImage();
	mp_sImageTag = _T("");
	mp_bAllowTextEdit = FALSE;
}

clsColumn::~clsColumn()
{
	delete mp_oImage;
	mp_oImage = NULL;
	AfxOleUnlockApp();
}

void clsColumn::OnFinalRelease()
{
	clsItemBase::OnFinalRelease();
}

BOOL clsColumn::GetAllowTextEdit(void)
{
	return mp_bAllowTextEdit;
}

void clsColumn::SetAllowTextEdit(BOOL Value)
{
	mp_bAllowTextEdit = Value;
}

BOOL clsColumn::GetAllowMove(void)
{
	return mp_bAllowMove;
}

void clsColumn::SetAllowMove(BOOL Value)
{
	mp_bAllowMove = Value;
}

BOOL clsColumn::GetAllowSize(void)
{
	return mp_bAllowSize;
}

void clsColumn::SetAllowSize(BOOL Value)
{
	mp_bAllowSize = Value;
}

CString clsColumn::GetKey(void)
{
	return mp_sKey;
}

void clsColumn::SetKey(CString Value)
{
	mp_oControl->GetColumns()->mp_oCollection->mp_SetKey(&mp_sKey, Value, COLUMNS_SET_KEY);
}

LONG clsColumn::GetWidth(void)
{
	return mp_lWidth;
}

void clsColumn::SetWidth(LONG Value)
{
	mp_lWidth = Value;
}

CString clsColumn::GetText(void)
{
	return mp_sText;
}

void clsColumn::SetText(CString Value)
{
	mp_sText = Value;
}

clsImage* clsColumn::GetImage(void)
{
	return mp_oImage;
}

CString clsColumn::GetStyleIndex(void)
{
	if (mp_sStyleIndex == "DS_COLUMN") 
	{
		return "";
	}
	else 
	{
		return mp_sStyleIndex;
	}
}

void clsColumn::SetStyleIndex(CString Value)
{
	Value = Value.Trim();
	if (Value.GetLength() == 0) {Value = "DS_COLUMN";}
	mp_sStyleIndex = Value;
	mp_oStyle = mp_oControl->GetStyles()->FItem(Value);
}

clsStyle* clsColumn::GetStyle(void)
{
	return mp_oStyle;
}

CString clsColumn::GetTag(void)
{
	return mp_sTag;
}

void clsColumn::SetTag(CString Value)
{
	mp_sTag = Value;
}

LONG clsColumn::GetLeftTrim(void)
{
	if (mp_lLeft < mp_oControl->Getmt_LeftMargin()) 
	{
		return mp_oControl->Getmt_LeftMargin();
	}
	else 
	{
		return mp_lLeft;
	}
}

LONG clsColumn::GetRightTrim(void)
{
	if (mp_lRight > mp_oControl->GetSplitter()->GetLeft()) 
	{
		return mp_oControl->GetSplitter()->GetLeft();
	}
	else 
	{
		return mp_lRight;
	}
}

LONG clsColumn::GetLeft(void)
{
	return mp_lLeft;
}

LONG clsColumn::GetTop(void)
{
	return mp_oControl->Getmt_TopMargin();
}

LONG clsColumn::GetRight(void)
{
	return mp_lRight;
}

LONG clsColumn::GetBottom(void)
{
	return mp_oControl->GetCurrentViewObject()->GetTimeLine()->GetBottom();
}

BOOL clsColumn::GetVisible(void)
{
	return mp_bVisible;
}

LONG clsColumn::GetIndex(void)
{
	return mp_lIndex;
}

void clsColumn::SetIndex(LONG Value)
{
	mp_lIndex = Value;
}

void clsColumn::Setf_lLeft(LONG Value)
{
	mp_lLeft = Value;;
}

void clsColumn::Setf_lRight(LONG Value)
{
	mp_lRight = Value;
}

void clsColumn::Setf_bVisible(BOOL Value)
{
	if (Value == FALSE) 
	{
		mp_lLeft = 0;
		mp_lRight = 0;
	}
	mp_bVisible = Value;
}

CString clsColumn::GetImageTag(void)
{
	return mp_sImageTag;
}

void clsColumn::SetImageTag(CString Value)
{
	mp_sImageTag = Value;
}

CString clsColumn::GetXML(void)
{
	clsXML oXML(mp_oControl, "Column");
	oXML.InitializeWriter();
	oXML.WriteProperty("AllowMove", mp_bAllowMove);
	oXML.WriteProperty("AllowSize", mp_bAllowSize);
	oXML.WriteProperty("Image", *mp_oImage);
	oXML.WriteProperty("Key", mp_sKey);
	oXML.WriteProperty("StyleIndex", mp_sStyleIndex);
	oXML.WriteProperty("Tag", mp_sTag);
	oXML.WriteProperty("Text", mp_sText);
	oXML.WriteProperty("Width", mp_lWidth);
	oXML.WriteProperty("ImageTag", mp_sImageTag);
	oXML.WriteProperty("AllowTextEdit", mp_bAllowTextEdit);
	return oXML.GetXML();
}

void clsColumn::SetXML(CString sXML)
{
	clsXML oXML(mp_oControl, "Column");
	oXML.SetXML(sXML);
	oXML.InitializeReader();
	oXML.ReadProperty("AllowMove", mp_bAllowMove);
	oXML.ReadProperty("AllowSize", mp_bAllowSize);
	oXML.ReadProperty("Image", *mp_oImage);
	oXML.ReadProperty("Key", mp_sKey);
	oXML.ReadProperty("StyleIndex", mp_sStyleIndex);
	SetStyleIndex(mp_sStyleIndex);
	oXML.ReadProperty("Tag", mp_sTag);
	oXML.ReadProperty("Text", mp_sText);
	oXML.ReadProperty("Width", mp_lWidth);
	oXML.ReadProperty("ImageTag", mp_sImageTag);
	oXML.ReadProperty("AllowTextEdit", mp_bAllowTextEdit);
}