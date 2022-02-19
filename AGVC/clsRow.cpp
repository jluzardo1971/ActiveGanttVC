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
#include "clsRow.h"



IMPLEMENT_DYNCREATE(clsRow, clsItemBase)

//{3D5F7B28-8E26-4BF3-AB60-653A1F529F05}
static const IID IID_IclsRow = { 0x3D5F7B28, 0x8E26, 0x4BF3, { 0xAB, 0x60, 0x65, 0x3A, 0x1F, 0x52, 0x9F, 0x05} };

//{6050BA5F-17A5-4936-8930-ACADB94C8BB9}
IMPLEMENT_OLECREATE_FLAGS(clsRow, "AGVC.clsRow", afxRegApartmentThreading, 0x6050BA5F, 0x17A5, 0x4936, 0x89, 0x30, 0xAC, 0xAD, 0xB9, 0x4C, 0x8B, 0xB9)

BEGIN_DISPATCH_MAP(clsRow, clsItemBase)
	DISP_PROPERTY_EX_ID(clsRow, "AllowMove", 1, odl_GetAllowMove, odl_SetAllowMove, VT_BOOL) 
	DISP_PROPERTY_EX_ID(clsRow, "AllowSize", 2, odl_GetAllowSize, odl_SetAllowSize, VT_BOOL) 
	DISP_PROPERTY_EX_ID(clsRow, "Key", 3, odl_GetKey, odl_SetKey, VT_BSTR) 
	DISP_PROPERTY_EX_ID(clsRow, "Container", 4, odl_GetContainer, odl_SetContainer, VT_BOOL) 
	DISP_PROPERTY_EX_ID(clsRow, "UseNodeImages", 5, odl_GetUseNodeImages, odl_SetUseNodeImages, VT_BOOL) 
	DISP_PROPERTY_EX_ID(clsRow, "MergeCells", 6, odl_GetMergeCells, odl_SetMergeCells, VT_BOOL) 
	DISP_PROPERTY_EX_ID(clsRow, "Height", 7, odl_GetHeight, odl_SetHeight, VT_I4) 
	DISP_PROPERTY_EX_ID(clsRow, "Text", 8, odl_GetText, odl_SetText, VT_BSTR) 
	DISP_PROPERTY_EX_ID(clsRow, "Image", 9, odl_GetImage, SetNotSupported, VT_DISPATCH) 
	DISP_PROPERTY_EX_ID(clsRow, "StyleIndex", 10, odl_GetStyleIndex, odl_SetStyleIndex, VT_BSTR) 
	DISP_PROPERTY_EX_ID(clsRow, "Style", 11, odl_GetStyle, SetNotSupported, VT_DISPATCH) 
	DISP_PROPERTY_EX_ID(clsRow, "Tag", 12, odl_GetTag, odl_SetTag, VT_BSTR) 
	DISP_PROPERTY_EX_ID(clsRow, "ClientAreaStyleIndex", 13, odl_GetClientAreaStyleIndex, odl_SetClientAreaStyleIndex, VT_BSTR) 
	DISP_PROPERTY_EX_ID(clsRow, "ClientAreaStyle", 14, odl_GetClientAreaStyle, SetNotSupported, VT_DISPATCH) 
	DISP_PROPERTY_EX_ID(clsRow, "Left", 15, odl_GetLeft, SetNotSupported, VT_I4) 
	DISP_PROPERTY_EX_ID(clsRow, "Top", 16, odl_GetTop, SetNotSupported, VT_I4) 
	DISP_PROPERTY_EX_ID(clsRow, "Right", 17, odl_GetRight, SetNotSupported, VT_I4) 
	DISP_PROPERTY_EX_ID(clsRow, "Bottom", 18, odl_GetBottom, SetNotSupported, VT_I4) 
	DISP_PROPERTY_EX_ID(clsRow, "Visible", 19, odl_GetVisible, SetNotSupported, VT_BOOL) 
	DISP_PROPERTY_EX_ID(clsRow, "Node", 20, odl_GetNode, SetNotSupported, VT_DISPATCH) 
	DISP_PROPERTY_EX_ID(clsRow, "Cells", 21, odl_GetCells, SetNotSupported, VT_DISPATCH) 
	DISP_FUNCTION_ID(clsRow, "GetXML", 22, odl_GetXML, VT_BSTR, VTS_NONE) 
	DISP_FUNCTION_ID(clsRow, "SetXML", 23, odl_SetXML, VT_EMPTY, VTS_BSTR)
	DISP_PROPERTY_EX_ID(clsRow, "Index", 24, odl_GetIndex, SetNotSupported, VT_I4)
	DISP_PROPERTY_EX_ID(clsRow, "ImageTag", 25, odl_GetImageTag, odl_SetImageTag, VT_BSTR) 
	DISP_PROPERTY_EX_ID(clsRow, "AllowTextEdit", 26, odl_GetAllowTextEdit, odl_SetAllowTextEdit, VT_BOOL)
	DISP_PROPERTY_EX_ID(clsRow, "Highlighted", 27, odl_GetHighlighted, odl_SetHighlighted, VT_BOOL) 
END_DISPATCH_MAP()

BEGIN_INTERFACE_MAP(clsRow, clsItemBase)
	INTERFACE_PART(clsRow, IID_IclsRow, Dispatch)
END_INTERFACE_MAP()

BEGIN_MESSAGE_MAP(clsRow, clsItemBase)
END_MESSAGE_MAP()

clsRow::clsRow()
{
	AfxMessageBox(_T("Invalid constructor has been invoked for clsRow. Please inform support@sourcecodestore.com of this error"), MB_OK | MB_ICONSTOP | MB_APPLMODAL);
}

clsRow::clsRow(CActiveGanttVCCtl* oControl)
{
	EnableAutomation();
	AfxOleLockApp();
	mp_oControl = oControl;
	mp_bContainer = TRUE;
	mp_bMergeCells = FALSE;
	mp_lHeight = 40;
	mp_sText = "";
	mp_sStyleIndex = "DS_ROW";
	mp_oStyle = mp_oControl->GetStyles()->FItem("DS_ROW");
	mp_sTag = "";
	mp_oCells = new clsCells(mp_oControl, this);
	mp_sClientAreaStyleIndex = "DS_CLIENTAREA";
	mp_oClientAreaStyle = mp_oControl->GetStyles()->FItem("DS_CLIENTAREA");
	mp_oNode = new clsNode(mp_oControl, this);
	mp_lTop = 0;
	mp_lBottom = 0;
	mp_bUseNodeImages = FALSE;
	mp_bAllowSize = TRUE;
	mp_bAllowMove = TRUE;
	mp_oImage = new clsImage();
	mp_sImageTag = _T("");
	mp_bAllowTextEdit = FALSE;
	mp_bHighlighted = FALSE;
}

clsRow::~clsRow()
{
	delete mp_oCells;
	mp_oCells = NULL;
	delete mp_oNode;
	mp_oNode = NULL;
	delete mp_oImage;
	mp_oImage = NULL;
	AfxOleUnlockApp();
}

void clsRow::OnFinalRelease()
{

	clsItemBase::OnFinalRelease();
}

BOOL clsRow::GetHighlighted(void)
{
	return mp_bHighlighted;
}

void clsRow::SetHighlighted(BOOL Value)
{
	mp_bHighlighted = Value;
}

BOOL clsRow::GetAllowTextEdit(void)
{
	return mp_bAllowTextEdit;
}

void clsRow::SetAllowTextEdit(BOOL Value)
{
	mp_bAllowTextEdit = Value;
}

BOOL clsRow::GetAllowMove(void)
{
	return mp_bAllowMove;
}

void clsRow::SetAllowMove(BOOL Value)
{
	mp_bAllowMove = Value;
}

BOOL clsRow::GetAllowSize(void)
{
	return mp_bAllowSize;
}

void clsRow::SetAllowSize(BOOL Value)
{
	mp_bAllowSize = Value;
}

CString clsRow::GetKey(void)
{
	return mp_sKey;
}

void clsRow::SetKey(CString Value)
{
	mp_oControl->GetRows()->mp_oCollection->mp_SetKey(&mp_sKey, Value, ROWS_SET_KEY);
}

BOOL clsRow::GetContainer(void)
{
	return mp_bContainer;
}

void clsRow::SetContainer(BOOL Value)
{
	mp_bContainer = Value;
}

BOOL clsRow::GetUseNodeImages(void)
{
	return mp_bUseNodeImages;
}

void clsRow::SetUseNodeImages(BOOL Value)
{
	mp_bUseNodeImages = Value;
}

BOOL clsRow::GetMergeCells(void)
{
	if (mp_oControl->GetTreeviewColumnIndex() != 0)
	{
		return FALSE;
	}
	else
	{
		return mp_bMergeCells;
	}
}

void clsRow::SetMergeCells(BOOL Value)
{
	mp_bMergeCells = Value;
}

LONG clsRow::GetHeight(void)
{
	if (mp_oNode->GetHidden() == FALSE) 
	{
		return mp_lHeight;
	}
	else 
	{
		return -1;
	}
}

void clsRow::SetHeight(LONG Value)
{
	mp_lHeight = Value;
}

CString clsRow::GetText(void)
{
	return mp_sText;
}

void clsRow::SetText(CString Value)
{
	mp_sText = Value;
}

clsImage* clsRow::GetImage(void)
{
	if (mp_bUseNodeImages == FALSE) 
	{
		return mp_oImage;
	}
	else 
	{
		clsImage* oImage = NULL;
		if (mp_oNode->GetExpanded() == TRUE && mp_oNode->Children() > 0 && (mp_oNode->GetExpandedImage() != NULL)) 
		{
			oImage = mp_oNode->GetExpandedImage();
		}
		else if (mp_oNode->GetSelected() == TRUE && (mp_oNode->GetSelectedImage() != NULL)) 
		{
			oImage = mp_oNode->GetSelectedImage();
		}
		else if ((mp_oNode->GetImage() != NULL)) 
		{
			oImage = mp_oNode->GetImage();
		}
		return oImage;
	}
}

CString clsRow::GetStyleIndex(void)
{
	if (mp_sStyleIndex == "DS_ROW") 
	{
		return "";
	}
	else 
	{
		return mp_sStyleIndex;
	}
}

void clsRow::SetStyleIndex(CString Value)
{
	Value = Value.Trim();
	if (Value.GetLength() == 0) {Value = "DS_ROW";} 
	mp_sStyleIndex = Value;
	mp_oStyle = mp_oControl->GetStyles()->FItem(Value);
}

clsStyle* clsRow::GetStyle(void)
{
	return mp_oStyle;
}

CString clsRow::GetTag(void)
{
	return mp_sTag;
}

void clsRow::SetTag(CString Value)
{
	mp_sTag = Value;
}

CString clsRow::GetClientAreaStyleIndex(void)
{
	if (mp_sClientAreaStyleIndex == "DS_CLIENTAREA") 
	{
		return "";
	}
	else 
	{
		return mp_sClientAreaStyleIndex;
	}
}

void clsRow::SetClientAreaStyleIndex(CString Value)
{
	Value = Value.Trim();
	if (Value.GetLength() == 0) {Value = "DS_CLIENTAREA";} 
	mp_sClientAreaStyleIndex = Value;
	mp_oClientAreaStyle = mp_oControl->GetStyles()->FItem(Value);
}

clsStyle* clsRow::GetClientAreaStyle(void)
{
	return mp_oClientAreaStyle;
}

LONG clsRow::GetLeft(void)
{
	return mp_oControl->Getmt_LeftMargin();
}

LONG clsRow::GetTop(void)
{
	return mp_lTop;
}

LONG clsRow::GetRight(void)
{
	return mp_oControl->GetSplitter()->GetLeft();
}

LONG clsRow::GetBottom(void)
{
	return mp_lBottom;
}

BOOL clsRow::GetVisible(void)
{

	if (mp_yClientAreaVisibility == VS_INSIDEVISIBLEAREA) 
	{
		return TRUE;
	}
	else 
	{
		return FALSE;
	}
}

LONG clsRow::GetIndex(void)
{
	return mp_lIndex;
}

void clsRow::SetIndex(LONG Value)
{
	mp_lIndex = Value;
}

clsNode* clsRow::GetNode(void)
{
	return mp_oNode;
}

clsCells* clsRow::GetCells(void)
{
	return mp_oCells;
}

void clsRow::Setf_lTop(LONG Value)
{
	mp_lTop = Value;
}

void clsRow::Setf_lBottom(LONG Value)
{
	mp_lBottom = Value;
}

LONG clsRow::GetClientAreaVisibility(void)
{
	return mp_yClientAreaVisibility;
}

void clsRow::SetClientAreaVisibility(LONG Value)
{
	mp_yClientAreaVisibility = Value;
}

CString clsRow::GetImageTag(void)
{
	return mp_sImageTag;
}

void clsRow::SetImageTag(CString Value)
{
	mp_sImageTag = Value;
}

Gdiplus::Color clsRow::Getf_RowBackColor(void)
{
    if (GetHighlighted() == TRUE)
	{
        return mp_oControl->GetConfiguration()->GetRowHighlightedBackColor();
	}
    else
	{
        if (mp_oControl->GetConfiguration()->GetAlternatingRowStyles() == TRUE)
		{
            if (GetIndex() % 2 == 1)
			{
                //Odd
                return mp_oControl->GetConfiguration()->GetRowOddBackColor();
			}
            else
			{
                //Even
                return mp_oControl->GetConfiguration()->GetRowEvenBackColor();
			}
		}
        else
		{
            return GetNode()->GetStyle()->GetBackColor();
		}
	}
}

CString clsRow::GetXML(void)
{
	clsXML oXML(mp_oControl, "Row");
	oXML.InitializeWriter();
	oXML.WriteProperty("AllowMove", mp_bAllowMove);
	oXML.WriteProperty("AllowSize", mp_bAllowSize);
	oXML.WriteProperty("ClientAreaStyleIndex", mp_sClientAreaStyleIndex);
	oXML.WriteProperty("Container", mp_bContainer);
	oXML.WriteProperty("Height", mp_lHeight);
	oXML.WriteProperty("Image", *mp_oImage);
	oXML.WriteProperty("Key", mp_sKey);
	oXML.WriteProperty("MergeCells", mp_bMergeCells);
	oXML.WriteProperty("StyleIndex", mp_sStyleIndex);
	oXML.WriteProperty("Tag", mp_sTag);
	oXML.WriteProperty("Text", mp_sText);
	oXML.WriteProperty("UseNodeImages", mp_bUseNodeImages);
	oXML.WriteProperty("ImageTag", mp_sImageTag);
	oXML.WriteProperty("AllowTextEdit", mp_bAllowTextEdit);
	oXML.WriteProperty("Highlighted", mp_bHighlighted);
	oXML.WriteObject(mp_oCells->GetXML());
	oXML.WriteObject(mp_oNode->GetXML());
	return oXML.GetXML();
}

void clsRow::SetXML(CString sXML)
{
	clsXML oXML(mp_oControl, "Row");
	oXML.SetXML(sXML);
	oXML.InitializeReader();
	oXML.ReadProperty("AllowMove", mp_bAllowMove);
	oXML.ReadProperty("AllowSize", mp_bAllowSize);
	oXML.ReadProperty("ClientAreaStyleIndex", mp_sClientAreaStyleIndex);
	SetClientAreaStyleIndex(mp_sClientAreaStyleIndex);
	oXML.ReadProperty("Container", mp_bContainer);
	oXML.ReadProperty("Height", mp_lHeight);
	oXML.ReadProperty("Image", *mp_oImage);
	oXML.ReadProperty("Key", mp_sKey);
	oXML.ReadProperty("MergeCells", mp_bMergeCells);
	oXML.ReadProperty("StyleIndex", mp_sStyleIndex);
	SetStyleIndex(mp_sStyleIndex);
	oXML.ReadProperty("Tag", mp_sTag);
	oXML.ReadProperty("Text", mp_sText);
	oXML.ReadProperty("UseNodeImages", mp_bUseNodeImages);
	oXML.ReadProperty("ImageTag", mp_sImageTag);
	oXML.ReadProperty("AllowTextEdit", mp_bAllowTextEdit);
	oXML.ReadProperty("Highlighted", mp_bHighlighted);
	mp_oCells->SetXML(oXML.ReadObject("Cells"));
	mp_oNode->SetXML(oXML.ReadObject("Node"));
}