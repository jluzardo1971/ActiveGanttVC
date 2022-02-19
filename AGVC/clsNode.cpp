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
#include "clsNode.h"



IMPLEMENT_DYNCREATE(clsNode, CCmdTarget)

//{AA3EF537-3637-4EB9-8721-40B33F7C333B}
static const IID IID_IclsNode = { 0xAA3EF537, 0x3637, 0x4EB9, { 0x87, 0x21, 0x40, 0xB3, 0x3F, 0x7C, 0x33, 0x3B} };

//{115AF1AC-67AD-43B1-9534-F41AF71F7A4E}
IMPLEMENT_OLECREATE_FLAGS(clsNode, "AGVC.clsNode", afxRegApartmentThreading, 0x115AF1AC, 0x67AD, 0x43B1, 0x95, 0x34, 0xF4, 0x1A, 0xF7, 0x1F, 0x7A, 0x4E)

BEGIN_DISPATCH_MAP(clsNode, CCmdTarget)
	DISP_PROPERTY_EX_ID(clsNode, "Row", 1, odl_GetRow, SetNotSupported, VT_DISPATCH) 
	DISP_PROPERTY_EX_ID(clsNode, "CheckBoxVisible", 2, odl_GetCheckBoxVisible, odl_SetCheckBoxVisible, VT_BOOL) 
	DISP_PROPERTY_EX_ID(clsNode, "ImageVisible", 3, odl_GetImageVisible, odl_SetImageVisible, VT_BOOL) 
	DISP_PROPERTY_EX_ID(clsNode, "Left", 4, odl_GetLeft, SetNotSupported, VT_I4) 
	DISP_PROPERTY_EX_ID(clsNode, "Top", 5, odl_GetTop, SetNotSupported, VT_I4) 
	DISP_PROPERTY_EX_ID(clsNode, "Bottom", 6, odl_GetBottom, SetNotSupported, VT_I4) 
	DISP_PROPERTY_EX_ID(clsNode, "Depth", 7, odl_GetDepth, odl_SetDepth, VT_I4)  
	DISP_PROPERTY_EX_ID(clsNode, "Expanded", 8, odl_GetExpanded, odl_SetExpanded, VT_BOOL) 
	DISP_PROPERTY_EX_ID(clsNode, "Selected", 9, odl_GetSelected, odl_SetSelected, VT_BOOL) 
	DISP_PROPERTY_EX_ID(clsNode, "ExpandedImage", 10, odl_GetExpandedImage, SetNotSupported, VT_DISPATCH) 
	DISP_PROPERTY_EX_ID(clsNode, "SelectedImage", 11, odl_GetSelectedImage, SetNotSupported, VT_DISPATCH) 
	DISP_PROPERTY_EX_ID(clsNode, "Image", 12, odl_GetImage, SetNotSupported, VT_DISPATCH) 
	DISP_PROPERTY_EX_ID(clsNode, "Tag", 13, odl_GetTag, odl_SetTag, VT_BSTR) 
	DISP_PROPERTY_EX_ID(clsNode, "Text", 14, odl_GetText, odl_SetText, VT_BSTR) 
	DISP_PROPERTY_EX_ID(clsNode, "Checked", 15, odl_GetChecked, odl_SetChecked, VT_BOOL) 
	DISP_PROPERTY_EX_ID(clsNode, "Height", 16, odl_GetHeight, odl_SetHeight, VT_I4) 
	DISP_PROPERTY_EX_ID(clsNode, "Hidden", 17, odl_GetHidden, SetNotSupported, VT_BOOL) 
	DISP_FUNCTION_ID(clsNode, "NextSibling", 18, odl_NextSibling, VT_DISPATCH, VTS_NONE) 
	DISP_FUNCTION_ID(clsNode, "PreviousSibling", 19, odl_PreviousSibling, VT_DISPATCH, VTS_NONE) 
	DISP_FUNCTION_ID(clsNode, "IsFirstSibling", 20, odl_IsFirstSibling, VT_BOOL, VTS_NONE) 
	DISP_FUNCTION_ID(clsNode, "FirstSibling", 21, odl_FirstSibling, VT_DISPATCH, VTS_NONE) 
	DISP_FUNCTION_ID(clsNode, "IsLastSibling", 22, odl_IsLastSibling, VT_BOOL, VTS_NONE) 
	DISP_FUNCTION_ID(clsNode, "LastSibling", 23, odl_LastSibling, VT_DISPATCH, VTS_NONE) 
	DISP_FUNCTION_ID(clsNode, "Child", 24, odl_Child, VT_DISPATCH, VTS_NONE) 
	DISP_FUNCTION_ID(clsNode, "Parent", 25, odl_Parent, VT_DISPATCH, VTS_NONE) 
	DISP_FUNCTION_ID(clsNode, "IsRoot", 26, odl_IsRoot, VT_BOOL, VTS_NONE) 
	DISP_FUNCTION_ID(clsNode, "FullPath", 27, odl_FullPath, VT_BSTR, VTS_NONE) 
	DISP_FUNCTION_ID(clsNode, "Children", 28, odl_Children, VT_I4, VTS_NONE) 
	DISP_FUNCTION_ID(clsNode, "GetXML", 29, odl_GetXML, VT_BSTR, VTS_NONE) 
	DISP_FUNCTION_ID(clsNode, "SetXML", 30, odl_SetXML, VT_EMPTY, VTS_BSTR)
	DISP_PROPERTY_EX_ID(clsNode, "ExpandedImageTag", 31, odl_GetExpandedImageTag, odl_SetExpandedImageTag, VT_BSTR) 
	DISP_PROPERTY_EX_ID(clsNode, "SelectedImageTag", 32, odl_GetSelectedImageTag, odl_SetSelectedImageTag, VT_BSTR) 
	DISP_PROPERTY_EX_ID(clsNode, "ImageTag", 33, odl_GetImageTag, odl_SetImageTag, VT_BSTR) 
	DISP_PROPERTY_EX_ID(clsNode, "StyleIndex", 34, odl_GetStyleIndex, odl_SetStyleIndex, VT_BSTR) 
	DISP_PROPERTY_EX_ID(clsNode, "Style", 35, odl_GetStyle, SetNotSupported, VT_DISPATCH) 
	DISP_PROPERTY_EX_ID(clsNode, "AllowTextEdit", 36, odl_GetAllowTextEdit, odl_SetAllowTextEdit, VT_BOOL) 
END_DISPATCH_MAP()

BEGIN_INTERFACE_MAP(clsNode, CCmdTarget)
	INTERFACE_PART(clsNode, IID_IclsNode, Dispatch)
END_INTERFACE_MAP()

BEGIN_MESSAGE_MAP(clsNode, CCmdTarget)
END_MESSAGE_MAP()

clsNode::clsNode()
{
	AfxMessageBox(_T("Invalid constructor has been invoked for clsNode. Please inform support@sourcecodestore.com of this error"), MB_OK | MB_ICONSTOP | MB_APPLMODAL);
}

clsNode::clsNode(CActiveGanttVCCtl* oControl, clsRow* oRow)
{
	EnableAutomation();
	AfxOleLockApp();
	mp_oControl = oControl;
	mp_oRow = oRow;
	mp_lDepth = 0;
	mp_bExpanded = TRUE;
	mp_oImage = new clsImage();
	mp_oExpandedImage = new clsImage();
	mp_oSelectedImage = new clsImage();
	mp_bChecked = FALSE;
	mp_bCheckBoxVisible = FALSE;
	mp_bImageVisible = FALSE;
	mp_sTag = "";
	mp_oParent = NULL;
	mp_sExpandedImageTag = _T("");
	mp_sSelectedImageTag = _T("");
	mp_sImageTag = _T("");
	mp_sStyleIndex = "DS_NODE";
	mp_oStyle = mp_oControl->GetStyles()->FItem("DS_NODE");
	mp_bAllowTextEdit = FALSE;
}

clsNode::~clsNode()
{
	delete mp_oImage;
	mp_oImage = NULL;
	delete mp_oSelectedImage;
	mp_oSelectedImage = NULL;
	delete mp_oExpandedImage;
	mp_oExpandedImage = NULL;
	AfxOleUnlockApp();
}

void clsNode::OnFinalRelease()
{
	CCmdTarget::OnFinalRelease();
}

BOOL clsNode::GetAllowTextEdit(void)
{
	return mp_bAllowTextEdit;
}

void clsNode::SetAllowTextEdit(BOOL Value)
{
	mp_bAllowTextEdit = Value;
}

clsRow* clsNode::GetRow(void)
{
	return mp_oRow;
}

BOOL clsNode::GetCheckBoxVisible(void)
{
	if (mp_oControl->GetTreeview()->GetCheckBoxes() == TRUE) 
	{
		return mp_bCheckBoxVisible;
	}
	else 
	{
		return FALSE;
	}
}

void clsNode::SetCheckBoxVisible(BOOL Value)
{
	mp_bCheckBoxVisible = Value;
}

BOOL clsNode::GetImageVisible(void)
{
	if (mp_oControl->GetTreeview()->GetImages() == TRUE) 
	{
		if (mp_oImage->GetImageP() == NULL) 
		{
			return FALSE;
		}
		else 
		{
			return mp_bImageVisible;
		}
	}
	else 
	{
		return FALSE;
	}
}

void clsNode::SetImageVisible(BOOL Value)
{
	mp_bImageVisible = Value;
}

LONG clsNode::GetLeft(void)
{
	if (GetHidden() == TRUE) 
	{
		return 0;
	}
	else 
	{
		return mp_oControl->GetTreeview()->GetLeft() + mp_oControl->GetTreeview()->GetIndentation() + (GetDepth() * mp_oControl->GetTreeview()->GetIndentation());
	}
}

LONG clsNode::GetTop(void)
{
	return mp_oRow->GetTop();
}

LONG clsNode::GetBottom(void)
{
	return mp_oRow->GetBottom();
}

LONG clsNode::GetDepth(void)
{
	return mp_lDepth;
}

void clsNode::SetDepth(LONG Value)
{
	if (Value < 0)
	{
		mp_oControl->mp_ErrorReport(NODE_INVALID_DEPTH, _T("Depth cannot be less than 0"), _T("clsNode.SetDepth"));
		Value = 0;
	}
	mp_lDepth = Value;
}

BOOL clsNode::GetExpanded(void)
{
	if (Children() > 0) 
	{
		return mp_bExpanded;
	}
	else 
	{
		return TRUE;
	}
}

void clsNode::SetExpanded(BOOL Value)
{
	mp_bExpanded = Value;
	mp_oControl->GetVerticalScrollBar()->Update();
}

BOOL clsNode::GetSelected(void)
{
	return (GetIndex() == mp_oControl->GetSelectedRowIndex());
}

void clsNode::SetSelected(BOOL Value)
{
	if (Value == TRUE) 
	{
		mp_oControl->SetSelectedRowIndex(GetIndex());
	}
	else 
	{
		if (mp_oControl->GetSelectedRowIndex() == GetIndex()) 
		{
			mp_oControl->SetSelectedRowIndex(0);
		}
	}
}

clsImage* clsNode::GetExpandedImage(void)
{
	return mp_oExpandedImage;
}

clsImage* clsNode::GetSelectedImage(void)
{
	return mp_oSelectedImage;
}

clsImage* clsNode::GetImage(void)
{
	return mp_oImage;
}

CString clsNode::GetTag(void)
{
	return mp_sTag;
}

void clsNode::SetTag(CString Value)
{
	mp_sTag = Value;
}

CString clsNode::GetText(void)
{
	return mp_oRow->GetText();
}

void clsNode::SetText(CString Value)
{
	mp_oRow->SetText(Value);
}

BOOL clsNode::GetChecked(void)
{
	return mp_bChecked;
}

void clsNode::SetChecked(BOOL Value)
{
	mp_bChecked = Value;
}

LONG clsNode::GetHeight(void)
{
	return mp_oRow->GetHeight();
}

void clsNode::SetHeight(LONG Value)
{
	mp_oRow->SetHeight(Value);
}

BOOL clsNode::GetHidden(void)
{
	if (mp_lDepth == 0)
	{
		return FALSE;
	}
	BOOL bHidden = FALSE;
	bHidden = FALSE;
	bHidden = RecurseHidden(this, bHidden);
	return bHidden;
}

LONG clsNode::GetYCenter(void)
{
	return GetTop() + (GetHeight() / 2);
}

LONG clsNode::Getmt_TextLeft(void)
{
	if (GetCheckBoxVisible() == FALSE && GetImageVisible() == FALSE) 
	{
		return GetLeft() + 10;
	}
	else if (GetCheckBoxVisible() == TRUE && GetImageVisible() == TRUE) 
	{
		return GetLeft() + 28 + mp_oImage->GetWidth() + 5;
	}
	else if (GetCheckBoxVisible() == TRUE) 
	{
		return GetLeft() + 28;
	}
	else if (GetImageVisible() == TRUE) 
	{
		return GetLeft() + 10 + mp_oImage->GetWidth() + 5;
	}
	return 0;
}

LONG clsNode::Getmt_TextTop(void)
{
	return GetYCenter() - (mp_oControl->clsG->mp_lStrHeight(GetText(), mp_oStyle->GetFont())) + 3;
}

LONG clsNode::Getmt_TextRight(void)
{
	return Getmt_TextLeft() + (mp_oControl->clsG->mp_lStrWidth(GetText(), mp_oStyle->GetFont())) + 10;
}

LONG clsNode::Getmt_TextBottom(void)
{
	return GetYCenter() + (mp_oControl->clsG->mp_lStrHeight(GetText(), mp_oStyle->GetFont())) - 3;
}

LONG clsNode::GetCheckBoxLeft(void)
{
	if (GetCheckBoxVisible() == TRUE && mp_oControl->GetTreeview()->GetCheckBoxes() == TRUE) 
	{
		return GetLeft() + 14;
	}
	else 
	{
		return 0;
	}
}

LONG clsNode::GetImageLeft(void)
{
	if (GetImageVisible() == TRUE && mp_oControl->GetTreeview()->GetImages() == TRUE) 
	{
		return Getmt_TextLeft() - 3 - mp_oImage->GetWidth();
	}
	else 
	{
		return 0;
	}
}

LONG clsNode::GetImageTop(void)
{
	if (GetImageVisible() == TRUE && mp_oControl->GetTreeview()->GetImages() == TRUE) 
	{
		return GetYCenter() - (mp_oImage->GetHeight() / 2);
	}
	else 
	{
		return 0;
	}
}

LONG clsNode::GetImageRight(void)
{
	if (GetImageVisible() == TRUE && mp_oControl->GetTreeview()->GetImages() == TRUE) 
	{
		return GetImageLeft() + mp_oImage->GetWidth();
	}
	else 
	{
		return 0;
	}
}

LONG clsNode::GetImageBottom(void)
{
	if (GetImageVisible() == TRUE && mp_oControl->GetTreeview()->GetImages() == TRUE) 
	{
		return GetImageTop() + mp_oImage->GetHeight();
	}
	else 
	{
		return 0;
	}
}

LONG clsNode::GetIndex(void)
{
	return mp_oRow->GetIndex();
}

clsNode* clsNode::NextSibling(void)
{
	LONG lIndex = GetIndex();
	clsNode* oNode = NULL;
	clsRow* oRow = NULL;
	if ((lIndex + 1) <= mp_oControl->GetRows()->GetCount()) 
	{
		for (lIndex = (GetIndex() + 1); lIndex <= mp_oControl->GetRows()->GetCount(); lIndex++) 
		{
			oRow = (clsRow*)mp_oControl->GetRows()->mp_oCollection->m_oReturnArrayElement(lIndex);
			oNode = oRow->GetNode();
			if (oNode->GetDepth() == GetDepth()) 
			{
				return oNode;
			}
			else if (oNode->GetDepth() == GetDepth() - 1) 
			{
				return NULL;
			}
		}
		return NULL;
	}
	else 
	{
		//This Node is the Last Node
		return NULL;
	}
}

clsNode* clsNode::PreviousSibling(void)
{
	LONG lIndex = GetIndex();
	clsNode* oNode = NULL;
	clsRow* oRow = NULL;
	if (lIndex > 1) 
	{
		for (lIndex = (GetIndex() - 1); lIndex >= 1; lIndex += -1) 
		{
			oRow = (clsRow*)mp_oControl->GetRows()->mp_oCollection->m_oReturnArrayElement(lIndex);
			oNode = oRow->GetNode();
			if (oNode->GetDepth() == GetDepth()) 
			{
				return oNode;
			}
			else if (oNode->GetDepth() == GetDepth() - 1) 
			{
				return NULL;
			}
		}
		return NULL;
	}
	else 
	{
		//This node is the First Node
		return NULL;
	}
}

BOOL clsNode::IsFirstSibling(void)
{
	clsNode* oNode = NULL;
	oNode = PreviousSibling();
	if (oNode == NULL) 
	{
		return TRUE;
	}
	else 
	{
		return FALSE;
	}
}

clsNode* clsNode::FirstSibling(void)
{
	clsNode* oNode = NULL;
	clsNode* oReturnNode = NULL;
	oReturnNode = this;
	oNode = PreviousSibling();
	while ((oNode != NULL)) 
	{
		oReturnNode = oNode;
		oNode = oReturnNode->PreviousSibling();
	}
	return oReturnNode;
}

BOOL clsNode::IsLastSibling(void)
{
	clsNode* oNode = NULL;
	oNode = NextSibling();
	if (oNode == NULL) 
	{
		return TRUE;
	}
	else 
	{
		return FALSE;
	}
}

clsNode* clsNode::LastSibling(void)
{
	clsNode* oNode = NULL;
	clsNode* oReturnNode = NULL;
	oReturnNode = this;
	oNode = NextSibling();
	while ((oNode != NULL)) 
	{
		oReturnNode = oNode;
		oNode = oReturnNode->NextSibling();
	}
	return oReturnNode;
}

clsNode* clsNode::Child(void)
{
	LONG lIndex = GetIndex() + 1;
	clsNode* oNode = NULL;
	clsRow* oRow = NULL;
	if (lIndex <= mp_oControl->GetRows()->GetCount()) 
	{
		oRow = (clsRow*)mp_oControl->GetRows()->mp_oCollection->m_oReturnArrayElement(lIndex);
		oNode = oRow->GetNode();
		if (oNode->GetDepth() == (GetDepth() + 1)) 
		{
			return oNode;
		}
		else 
		{
			return NULL;
		}
	}
	else 
	{
		return NULL;
	}
}

clsNode* clsNode::Parent(void)
{
	return mp_oParent;
}

BOOL clsNode::IsRoot(void)
{
	if (Parent() == NULL) 
	{
		return TRUE;
	}
	else 
	{
		return FALSE;
	}
}

CString clsNode::FullPath(void)
{
	clsNode* oNode = NULL;
	CString sReturn = "";
	oNode = Parent();
	while ((oNode != NULL)) 
	{
		sReturn = oNode->GetText() + mp_oControl->GetTreeview()->GetPathSeparator() + sReturn;
		oNode = oNode->Parent();
	}
	return sReturn + GetText();
}

LONG clsNode::Children(void)
{
	LONG lIndex = 0;
	clsNode* oNode = NULL;
	clsRow* oRow = NULL;
	LONG lReturn = 0;
	for (lIndex = (mp_oRow->GetIndex() + 1); lIndex <= mp_oControl->GetRows()->GetCount(); lIndex++) 
	{
		oRow = (clsRow*)mp_oControl->GetRows()->mp_oCollection->m_oReturnArrayElement(lIndex);
		oNode = oRow->GetNode();
		if (GetDepth() < oNode->GetDepth()) 
		{
			if (GetDepth() + 1 == oNode->GetDepth()) 
			{
				lReturn = lReturn + 1;
			}
		}
		else 
		{
			break; 
		}
	}
	return lReturn;
}

BOOL clsNode::RecurseHidden(clsNode* oNode,BOOL bHidden)
{
	oNode = oNode->Parent();
	if ((oNode != NULL)) 
	{
		if (oNode->mp_bExpanded == FALSE) 
		{
			bHidden = TRUE;
		}
		bHidden = RecurseHidden(oNode, bHidden);
	}
	return bHidden;
}

CString clsNode::GetExpandedImageTag(void)
{
	return mp_sExpandedImageTag;
}

void clsNode::SetExpandedImageTag(CString Value)
{
	mp_sExpandedImageTag = Value;
}

CString clsNode::GetSelectedImageTag(void)
{
	return mp_sSelectedImageTag;
}

void clsNode::SetSelectedImageTag(CString Value)
{
	mp_sSelectedImageTag = Value;
}

CString clsNode::GetImageTag(void)
{
	return mp_sImageTag;
}

void clsNode::SetImageTag(CString Value)
{
	mp_sImageTag = Value;
}

CString clsNode::GetStyleIndex(void)
{
	if (mp_sStyleIndex == "DS_NODE") 
	{
		return "";
	}
	else 
	{
		return mp_sStyleIndex;
	}
}

void clsNode::SetStyleIndex(CString Value)
{
	Value = Value.Trim();
	if (Value.GetLength() == 0) {Value = "DS_NODE";} 
	mp_sStyleIndex = Value;
	mp_oStyle = mp_oControl->GetStyles()->FItem(Value);
}

clsStyle* clsNode::GetStyle(void)
{
	return mp_oStyle;
}

CString clsNode::GetXML(void)
{
	clsXML oXML(mp_oControl, "Node");
	oXML.InitializeWriter();
	oXML.WriteProperty("CheckBoxVisible", mp_bCheckBoxVisible);
	oXML.WriteProperty("Checked", mp_bChecked);
	oXML.WriteProperty("Depth", mp_lDepth);
	oXML.WriteProperty("Expanded", mp_bExpanded);
	oXML.WriteProperty("ExpandedImage", *mp_oExpandedImage);
	oXML.WriteProperty("Image", *mp_oImage);
	oXML.WriteProperty("ImageVisible", mp_bImageVisible);
	oXML.WriteProperty("SelectedImage", *mp_oSelectedImage);
	oXML.WriteProperty("Tag", mp_sTag);
	oXML.WriteProperty("ExpandedImageTag", mp_sExpandedImageTag);
	oXML.WriteProperty("SelectedImageTag", mp_sSelectedImageTag);
	oXML.WriteProperty("ImageTag", mp_sImageTag);
	oXML.WriteProperty("StyleIndex", mp_sStyleIndex);
	oXML.WriteProperty("AllowTextEdit", mp_bAllowTextEdit);
	return oXML.GetXML();
}

void clsNode::SetXML(CString sXML)
{
	clsXML oXML(mp_oControl, "Node");
	oXML.SetXML(sXML);
	oXML.InitializeReader();
	oXML.ReadProperty("CheckBoxVisible", mp_bCheckBoxVisible);
	oXML.ReadProperty("Checked", mp_bChecked);
	oXML.ReadProperty("Depth", mp_lDepth);
	oXML.ReadProperty("Expanded", mp_bExpanded);
	oXML.ReadProperty("ExpandedImage", *mp_oExpandedImage);
	oXML.ReadProperty("Image", *mp_oImage);
	oXML.ReadProperty("ImageVisible", mp_bImageVisible);
	oXML.ReadProperty("SelectedImage", *mp_oSelectedImage);
	oXML.ReadProperty("Tag", mp_sTag);
	oXML.ReadProperty("ExpandedImageTag", mp_sExpandedImageTag);
	oXML.ReadProperty("SelectedImageTag", mp_sSelectedImageTag);
	oXML.ReadProperty("ImageTag", mp_sImageTag);
	oXML.ReadProperty("StyleIndex", mp_sStyleIndex);
	SetStyleIndex(mp_sStyleIndex);
	oXML.ReadProperty("AllowTextEdit", mp_bAllowTextEdit);
}

IDispatch* clsNode::odl_GetRow(void) //Circular Reference
{
	
	return GetRow()->GetIDispatch(TRUE);
}