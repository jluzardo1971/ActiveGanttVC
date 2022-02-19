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
#include "clsTreeview.h"



IMPLEMENT_DYNCREATE(clsTreeview, CCmdTarget)

//{74CA6930-8845-4556-AE74-0F217D7360AF}
static const IID IID_IclsTreeview = { 0x74CA6930, 0x8845, 0x4556, { 0xAE, 0x74, 0x0F, 0x21, 0x7D, 0x73, 0x60, 0xAF} };

//{47FACBE2-D949-40C3-ADB0-D1078B252B46}
IMPLEMENT_OLECREATE_FLAGS(clsTreeview, "AGVC.clsTreeview", afxRegApartmentThreading, 0x47FACBE2, 0xD949, 0x40C3, 0xAD, 0xB0, 0xD1, 0x07, 0x8B, 0x25, 0x2B, 0x46)

BEGIN_DISPATCH_MAP(clsTreeview, CCmdTarget)
	DISP_PROPERTY_EX_ID(clsTreeview, "FirstVisibleNode", 1, odl_GetFirstVisibleNode, odl_SetFirstVisibleNode, VT_I4) 
	DISP_PROPERTY_EX_ID(clsTreeview, "LastVisibleNode", 2, odl_GetLastVisibleNode, SetNotSupported, VT_I4) 
	DISP_PROPERTY_EX_ID(clsTreeview, "Indentation", 3, odl_GetIndentation, odl_SetIndentation, VT_I4) 
	DISP_PROPERTY_EX_ID(clsTreeview, "CheckBoxBorderColor", 4, odl_GetCheckBoxBorderColor, odl_SetCheckBoxBorderColor, VT_COLOR) 
	DISP_PROPERTY_EX_ID(clsTreeview, "CheckBoxBackColor", 5, odl_GetCheckBoxColor, odl_SetCheckBoxColor, VT_COLOR) 
	DISP_PROPERTY_EX_ID(clsTreeview, "CheckBoxMarkColor", 6, odl_GetCheckBoxMarkColor, odl_SetCheckBoxMarkColor, VT_COLOR) 
	DISP_PROPERTY_EX_ID(clsTreeview, "BackColor", 7, odl_GetBackColor, odl_SetBackColor, VT_COLOR) 
	DISP_PROPERTY_EX_ID(clsTreeview, "PathSeparator", 8, odl_GetPathSeparator, odl_SetPathSeparator, VT_BSTR) 
	DISP_PROPERTY_EX_ID(clsTreeview, "TreeLines", 9, odl_GetTreeLines, odl_SetTreeLines, VT_BOOL) 
	DISP_PROPERTY_EX_ID(clsTreeview, "TreeviewMode", 10, odl_GetTreeviewMode, odl_SetTreeviewMode, VT_I4) 
	DISP_PROPERTY_EX_ID(clsTreeview, "Images", 11, odl_GetImages, odl_SetImages, VT_BOOL) 
	DISP_PROPERTY_EX_ID(clsTreeview, "CheckBoxes", 12, odl_GetCheckBoxes, odl_SetCheckBoxes, VT_BOOL) 
	DISP_PROPERTY_EX_ID(clsTreeview, "FullColumnSelect", 13, odl_GetFullColumnSelect, odl_SetFullColumnSelect, VT_BOOL) 
	DISP_PROPERTY_EX_ID(clsTreeview, "ExpansionOnSelection", 14, odl_GetExpansionOnSelection, odl_SetExpansionOnSelection, VT_BOOL) 
	DISP_PROPERTY_EX_ID(clsTreeview, "SelectedBackColor", 15, odl_GetSelectedBackColor, odl_SetSelectedBackColor, VT_COLOR) 
	DISP_PROPERTY_EX_ID(clsTreeview, "SelectedForeColor", 16, odl_GetSelectedForeColor, odl_SetSelectedForeColor, VT_COLOR) 
	DISP_PROPERTY_EX_ID(clsTreeview, "TreeLineColor", 17, odl_GetTreeLineColor, odl_SetTreeLineColor, VT_COLOR) 
	DISP_PROPERTY_EX_ID(clsTreeview, "PlusMinusBorderColor", 18, odl_GetPlusMinusBorderColor, odl_SetPlusMinusBorderColor, VT_COLOR) 
	DISP_PROPERTY_EX_ID(clsTreeview, "PlusMinusSignColor", 19, odl_GetPlusMinusSignColor, odl_SetPlusMinusSignColor, VT_COLOR) 
	DISP_FUNCTION_ID(clsTreeview, "ClearSelections", 20, odl_ClearSelections, VT_EMPTY, VTS_NONE) 
	DISP_FUNCTION_ID(clsTreeview, "SetXML", 21, odl_SetXML, VT_EMPTY, VTS_BSTR) 
	DISP_FUNCTION_ID(clsTreeview, "GetXML", 22, odl_GetXML, VT_BSTR, VTS_NONE)
	DISP_PROPERTY_EX_ID(clsTreeview, "CheckBoxBackgroundMode", 23, odl_GetCheckBoxBackgroundMode, odl_SetCheckBoxBackgroundMode, VT_I4)
	DISP_PROPERTY_EX_ID(clsTreeview, "PlusMinusBackColor", 24, odl_GetPlusMinusBackColor, odl_SetPlusMinusBackColor, VT_COLOR)
	DISP_PROPERTY_EX_ID(clsTreeview, "PlusMinusBackgroundMode", 25, odl_GetPlusMinusBackgroundMode, odl_SetPlusMinusBackgroundMode, VT_I4)
	DISP_PROPERTY_EX_ID(clsTreeview, "ExpandIconForeColor", 26, odl_GetExpandIconForeColor, odl_SetExpandIconForeColor, VT_COLOR)
	DISP_PROPERTY_EX_ID(clsTreeview, "ExpandIconBackColor", 27, odl_GetExpandIconBackColor, odl_SetExpandIconBackColor, VT_COLOR) 
	DISP_PROPERTY_EX_ID(clsTreeview, "ExpandIconDropShadowColor", 28, odl_GetExpandIconDropShadowColor, odl_SetExpandIconDropShadowColor, VT_COLOR) 
	DISP_PROPERTY_EX_ID(clsTreeview, "CollapseIconForeColor", 29,  odl_GetCollapseIconForeColor, odl_SetCollapseIconForeColor, VT_COLOR)
	DISP_PROPERTY_EX_ID(clsTreeview, "CollapseIconDropShadowColor", 30, odl_GetCollapseIconDropShadowColor, odl_SetCollapseIconDropShadowColor, VT_COLOR)


END_DISPATCH_MAP()

BEGIN_INTERFACE_MAP(clsTreeview, CCmdTarget)
	INTERFACE_PART(clsTreeview, IID_IclsTreeview, Dispatch)
END_INTERFACE_MAP()

BEGIN_MESSAGE_MAP(clsTreeview, CCmdTarget)
END_MESSAGE_MAP()

clsTreeview::clsTreeview()
{
	AfxMessageBox(_T("Invalid constructor has been invoked for clsTreeview. Please inform support@sourcecodestore.com of this error"), MB_OK | MB_ICONSTOP | MB_APPLMODAL);
}

clsTreeview::clsTreeview(CActiveGanttVCCtl* oControl)
{
	EnableAutomation();
	AfxOleLockApp();
	mp_oControl = oControl;
	mp_lLastVisibleNode = 0;
	mp_lIndentation = 20;
	mp_clrBackColor = Color::White;
	mp_clrCheckBoxBorderColor = Color::Gray;
	mp_clrCheckBoxBackColor = Color::White;
	mp_clrCheckBoxMarkColor = Color::Black;
	mp_clrSelectedBackColor = Color::Blue;
	mp_clrSelectedForeColor = Color::White;
	mp_clrTreeLineColor = Color::Gray;
	mp_clrPlusMinusBorderColor = Color::Gray;
	mp_clrPlusMinusSignColor = Color::Black;
	mp_clrPlusMinusBackColor = Color::White;
	mp_clrExpandIconForeColor = Color::MakeARGB(255, 0, 0, 0);
	mp_clrExpandIconBackColor = Color::MakeARGB(255, 255, 255, 255);
	mp_clrExpandIconDropShadowColor = Color::MakeARGB(255, 211, 211, 211);
	mp_clrCollapseIconForeColor = Color::MakeARGB(255, 0, 0, 0);
	mp_clrCollapseIconDropShadowColor = Color::MakeARGB(255, 128, 128, 128);
	mp_bCheckBoxes = FALSE;
	mp_bTreeLines = TRUE;
	mp_bImages = TRUE;
	mp_yTreeviewMode = TM_PLUSMINUSSIGNS;
	mp_bFullColumnSelect = FALSE;
	mp_bExpansionOnSelection = FALSE;
	mp_sPathSeparator = "/";
	mp_yOperation = EO_NONE;
	mp_yCheckBoxBackgroundMode = FP_TRANSPARENT;
    mp_yPlusMinusBackgroundMode = FP_TRANSPARENT;
}

clsTreeview::~clsTreeview()
{
	AfxOleUnlockApp();
}

void clsTreeview::OnFinalRelease()
{
	CCmdTarget::OnFinalRelease();
}

LONG clsTreeview::GetFirstVisibleNode(void)
{
	if (mp_oControl->GetRows()->GetCount() == 0) 
	{
		return 0;
	}
	else 
	{
		return mp_oControl->GetRows()->GetRealFirstVisibleRow();
	}
}

void clsTreeview::SetFirstVisibleNode(LONG Value)
{
	if (Value < 1) 
	{
		Value = 1;
	}
	else if (((Value > mp_oControl->GetRows()->GetCount()) && (mp_oControl->GetRows()->GetCount() != 0))) 
	{
		Value = mp_oControl->GetRows()->GetCount();
	}
	mp_oControl->GetVerticalScrollBar()->SetValue(Value);
}

LONG clsTreeview::GetLastVisibleNode(void)
{
	return mp_lLastVisibleNode;
}

LONG clsTreeview::GetIndentation(void)
{
	return mp_lIndentation;
}

void clsTreeview::SetIndentation(LONG Value)
{
	mp_lIndentation = Value;
}

Gdiplus::Color clsTreeview::GetCheckBoxBorderColor(void)
{
	return mp_clrCheckBoxBorderColor;
}

void clsTreeview::SetCheckBoxBorderColor(Gdiplus::Color Value)
{
	mp_clrCheckBoxBorderColor = Value;
}

Gdiplus::Color clsTreeview::GetCheckBoxBackColor(void)
{
	return mp_clrCheckBoxBackColor;
}

void clsTreeview::SetCheckBoxBackColor(Gdiplus::Color Value)
{
	mp_clrCheckBoxBackColor = Value;
}

LONG clsTreeview::GetCheckBoxBackgroundMode(void)
{
	return mp_yCheckBoxBackgroundMode;
}

void clsTreeview::SetCheckBoxBackgroundMode(LONG Value)
{
    if (Value != FP_SOLID || Value != FP_TRANSPARENT)
	{
		Value = FP_SOLID;
	}
	mp_yCheckBoxBackgroundMode = Value;
}

Gdiplus::Color clsTreeview::GetCheckBoxMarkColor(void)
{
	return mp_clrCheckBoxMarkColor;
}

void clsTreeview::SetCheckBoxMarkColor(Gdiplus::Color Value)
{
	mp_clrCheckBoxMarkColor = Value;
}

Gdiplus::Color clsTreeview::GetBackColor(void)
{
	return mp_clrBackColor;
}

void clsTreeview::SetBackColor(Gdiplus::Color Value)
{
	mp_clrBackColor = Value;
}

CString clsTreeview::GetPathSeparator(void)
{
	return mp_sPathSeparator;
}

void clsTreeview::SetPathSeparator(CString Value)
{
	mp_sPathSeparator = Value;
}

BOOL clsTreeview::GetTreeLines(void)
{
	return mp_bTreeLines;
}

void clsTreeview::SetTreeLines(BOOL Value)
{
	mp_bTreeLines = Value;
}

LONG clsTreeview::GetTreeviewMode(void)
{
	return mp_yTreeviewMode;
}

void clsTreeview::SetTreeviewMode(LONG Value)
{
	mp_yTreeviewMode = Value;
}

BOOL clsTreeview::GetImages(void)
{
	return mp_bImages;
}

void clsTreeview::SetImages(BOOL Value)
{
	mp_bImages = Value;
}

BOOL clsTreeview::GetCheckBoxes(void)
{
	return mp_bCheckBoxes;
}

void clsTreeview::SetCheckBoxes(BOOL Value)
{
	mp_bCheckBoxes = Value;
}

BOOL clsTreeview::GetFullColumnSelect(void)
{
	return mp_bFullColumnSelect;
}

void clsTreeview::SetFullColumnSelect(BOOL Value)
{
	mp_bFullColumnSelect = Value;
}

BOOL clsTreeview::GetExpansionOnSelection(void)
{
	return mp_bExpansionOnSelection;
}

void clsTreeview::SetExpansionOnSelection(BOOL Value)
{
	mp_bExpansionOnSelection = Value;
}

Gdiplus::Color clsTreeview::GetSelectedBackColor(void)
{
	return mp_clrSelectedBackColor;
}

void clsTreeview::SetSelectedBackColor(Gdiplus::Color Value)
{
	mp_clrSelectedBackColor = Value;
}

Gdiplus::Color clsTreeview::GetSelectedForeColor(void)
{
	return mp_clrSelectedForeColor;
}

void clsTreeview::SetSelectedForeColor(Gdiplus::Color Value)
{
	mp_clrSelectedForeColor = Value;
}

Gdiplus::Color clsTreeview::GetTreeLineColor(void)
{
	return mp_clrTreeLineColor;
}

void clsTreeview::SetTreeLineColor(Gdiplus::Color Value)
{
	mp_clrTreeLineColor = Value;
}

Gdiplus::Color clsTreeview::GetPlusMinusBackColor(void)
{
	return mp_clrPlusMinusBackColor;
}

void clsTreeview::SetPlusMinusBackColor(Gdiplus::Color Value)
{
	mp_clrPlusMinusBackColor = Value;
}

LONG clsTreeview::GetPlusMinusBackgroundMode(void)
{
	return mp_yPlusMinusBackgroundMode;
}

void clsTreeview::SetPlusMinusBackgroundMode(LONG Value)
{
    if (Value != FP_SOLID || Value != FP_TRANSPARENT)
	{
        Value = FP_SOLID;
	}
    mp_yPlusMinusBackgroundMode = Value;
}

Gdiplus::Color clsTreeview::GetPlusMinusBorderColor(void)
{
	return mp_clrPlusMinusBorderColor;
}

void clsTreeview::SetPlusMinusBorderColor(Gdiplus::Color Value)
{
	mp_clrPlusMinusBorderColor = Value;
}

Gdiplus::Color clsTreeview::GetPlusMinusSignColor(void)
{
	return mp_clrPlusMinusSignColor;
}

void clsTreeview::SetPlusMinusSignColor(Gdiplus::Color Value)
{
	mp_clrPlusMinusSignColor = Value;
}

LONG clsTreeview::GetLeft(void)
{
	if (mp_oControl->GetTreeviewColumnIndex() == 0) 
	{
		return 0;
	}
	return mp_oControl->GetColumns()->Item(CStr(mp_oControl->GetTreeviewColumnIndex()))->GetLeft();
}

LONG clsTreeview::GetRight(void)
{
	if (mp_oControl->GetTreeviewColumnIndex() == 0)
	{
		return 0;
	}
	return mp_oControl->GetColumns()->Item(CStr(mp_oControl->GetTreeviewColumnIndex()))->GetRight();
}

LONG clsTreeview::GetLeftTrim(void)
{
	if (mp_oControl->GetTreeviewColumnIndex() == 0) 
	{
		return 0;
	}
	return mp_oControl->GetColumns()->Item(CStr(mp_oControl->GetTreeviewColumnIndex()))->GetLeftTrim();
}

LONG clsTreeview::GetRightTrim(void)
{
	if (mp_oControl->GetTreeviewColumnIndex() == 0)
	{
		return 0;
	}
	return mp_oControl->GetColumns()->Item(CStr(mp_oControl->GetTreeviewColumnIndex()))->GetRightTrim();
}

LONG clsTreeview::Getf_FirstVisibleNode(void)
{
	if (mp_oControl->GetRows()->GetCount() == 0) 
	{
		return 0;
	}
	else 
	{
		return mp_oControl->GetVerticalScrollBar()->GetValue();
	}
}

void clsTreeview::Setf_LastVisibleNode(LONG Value)
{
	mp_lLastVisibleNode = Value;
}

LONG clsTreeview::Getmt_BorderThickness(void)
{
	return  mp_oControl->Getmt_BorderThickness();
}

BOOL clsTreeview::OverControl(LONG X,LONG Y)
{
	clsRow* oRow = NULL;
	LONG lIndex = 0;
	if (mp_oControl->GetTreeviewColumnIndex() == 0) 
	{
		return FALSE;
	}
	if (!(X >= GetLeftTrim() && X <= GetRightTrim())) 
	{
		return FALSE;
	}
	for (lIndex = 1; lIndex <= mp_oControl->GetRows()->GetCount(); lIndex++) 
	{
		oRow = (clsRow*)mp_oControl->GetRows()->mp_oCollection->m_oReturnArrayElement(lIndex);
		if (oRow->GetVisible() == TRUE) 
		{
			if (Y >= oRow->GetTop() && Y <= oRow->GetBottom()) 
			{
				return TRUE;
			}
		}
	}
	return FALSE;
}

void clsTreeview::OnMouseHover(LONG X,LONG Y)
{
	switch (CursorPosition(X, Y)) 
	{
	case EVT_TREEVIEWCHECKBOX:
		mp_EO_CHECKBOXCLICK(MouseHover, X, Y);
		break;
	case EVT_TREEVIEWSIGN:
		mp_EO_SIGNCLICK(MouseHover, X, Y);
		break;
	case EVT_SELECTEDROW:
		if (mp_bCursorEditTextNode(X, Y) == TRUE)
		{
			return;
		}
		switch (mp_oControl->MouseKeyboardEvents->mp_yRowArea(X, Y)) 
		{
		case EA_BOTTOM:
			mp_EO_ROWSIZING(MouseHover, X, Y);
			break;
		case EA_CENTER:
			mp_EO_ROWMOVEMENT(MouseHover, X, Y);
			break;
		}
		break;
	case EVT_ROW:
		mp_EO_ROWSELECTION(MouseHover, X, Y);
		break;
	}
}

void clsTreeview::OnMouseDown(LONG X,LONG Y)
{
	switch (CursorPosition(X, Y)) 
	{
	case EVT_TREEVIEWCHECKBOX:
		mp_EO_CHECKBOXCLICK(MouseDown, X, Y);
		mp_yOperation = EO_CHECKBOXCLICK;
		break;
	case EVT_TREEVIEWSIGN:
		mp_EO_SIGNCLICK(MouseDown, X, Y);
		mp_yOperation = EO_SIGNCLICK;
		break;
	case EVT_SELECTEDROW:
		if (mp_bShowEditTextNode(X, Y) == TRUE)
		{
			return;
		}
		switch (mp_oControl->MouseKeyboardEvents->mp_yRowArea(X, Y)) 
		{
		case EA_BOTTOM:
			mp_EO_ROWSIZING(MouseDown, X, Y);
			mp_yOperation = EO_ROWSIZING;
			break;
		case EA_CENTER:
			mp_EO_ROWMOVEMENT(MouseDown, X, Y);
			mp_yOperation = EO_ROWMOVEMENT;
			break;
		}
		break;
	case EVT_ROW:
		mp_EO_ROWSELECTION(MouseDown, X, Y);
		mp_yOperation = EO_ROWSELECTION;
		break;
	}
}

void clsTreeview::OnMouseMove(LONG X,LONG Y)
{
	E_OPERATION yOperation = (E_OPERATION)mp_yOperation;
	switch (mp_yOperation) 
	{
	case EO_CHECKBOXCLICK:
		mp_EO_CHECKBOXCLICK(MouseMove, X, Y);
		break;
	case EO_SIGNCLICK:
		mp_EO_SIGNCLICK(MouseMove, X, Y);
		break;
	case EO_ROWMOVEMENT:
		mp_EO_ROWMOVEMENT(MouseMove, X, Y);
		break;
	case EO_ROWSIZING:
		mp_EO_ROWSIZING(MouseMove, X, Y);
		break;
	case EO_ROWSELECTION:
		mp_EO_ROWSELECTION(MouseMove, X, Y);
		break;
	}
}

void clsTreeview::OnMouseUp(LONG X,LONG Y)
{
	switch (mp_yOperation) 
	{
	case EO_CHECKBOXCLICK:
		mp_EO_CHECKBOXCLICK(MouseUp, X, Y);
		break;
	case EO_SIGNCLICK:
		mp_EO_SIGNCLICK(MouseUp, X, Y);
		break;
	case EO_ROWMOVEMENT:
		mp_EO_ROWMOVEMENT(MouseUp, X, Y);
		break;
	case EO_ROWSIZING:
		mp_EO_ROWSIZING(MouseUp, X, Y);
		break;
	case EO_ROWSELECTION:
		mp_EO_ROWSELECTION(MouseUp, X, Y);
		break;
	}
	mp_yOperation = EO_NONE;
}

LONG clsTreeview::CursorPosition(LONG X,LONG Y)
{
	if (mp_bOverCheckBox(X, Y) == TRUE) 
	{
		return EVT_TREEVIEWCHECKBOX;
	}
	else if (mp_bOverPlusMinusSign(X, Y) == TRUE) 
	{
		return EVT_TREEVIEWSIGN;
	}
	else if (mp_oControl->MouseKeyboardEvents->mp_bOverSelectedRow(X, Y) == TRUE) 
	{
		return EVT_SELECTEDROW;
	}
	else if (mp_oControl->MouseKeyboardEvents->mp_bOverRow(X, Y) == TRUE) 
	{
		return EVT_ROW;
	}
	return EVT_NONE;
}

void clsTreeview::mp_EO_CHECKBOXCLICK(LONG yMouseKeyBoardEvent,LONG X,LONG Y)
{
	clsRow* oRow = NULL;
	clsNode* oNode = NULL;
	switch (yMouseKeyBoardEvent) 
	{
	case MouseHover:
		mp_oControl->MouseKeyboardEvents->mp_SetCursor(CT_NORMAL);
		break;
	case MouseDown:
		s_chkCLK.lNodeIndex = mp_oControl->GetMathLib()->GetNodeIndexByCheckBoxPosition(X, Y);
		oRow = (clsRow*)mp_oControl->GetRows()->mp_oCollection->m_oReturnArrayElement(s_chkCLK.lNodeIndex);
		oNode = oRow->GetNode();
		oNode->SetChecked(!oNode->GetChecked());
		break;
	case MouseMove:
		break;
	case MouseUp:
		mp_oControl->GetNodeEventArgs()->Clear();
		mp_oControl->GetNodeEventArgs()->SetIndex(s_chkCLK.lNodeIndex);
		mp_oControl->FireNodeChecked();
		mp_oControl->Redraw();
		break;
	case MouseClick:
		break;
	case MouseDblClick:
		break;
	case MouseWheel:
		break;
	//case KeyDown:
	//	break;
	//case KeyUp:
	//	break;
	//case KeyPress:
	//	break;
	}
}

void clsTreeview::mp_EO_SIGNCLICK(LONG yMouseKeyBoardEvent,LONG X,LONG Y)
{
	clsRow* oRow = NULL;
	clsNode* oNode = NULL;
	switch (yMouseKeyBoardEvent) 
	{
	case MouseHover:
		mp_oControl->MouseKeyboardEvents->mp_SetCursor(CT_NORMAL);
		break;
	case MouseDown:
		s_sgnCLK.lNodeIndex = mp_oControl->GetMathLib()->GetNodeIndexBySignPosition(X, Y);
		oRow = (clsRow*)mp_oControl->GetRows()->mp_oCollection->m_oReturnArrayElement(s_sgnCLK.lNodeIndex);
		oNode = oRow->GetNode();
		oNode->SetExpanded(!oNode->GetExpanded());
		break;
	case MouseMove:
		break;
	case MouseUp:
		mp_oControl->GetNodeEventArgs()->Clear();
		mp_oControl->GetNodeEventArgs()->SetIndex(s_sgnCLK.lNodeIndex);
		mp_oControl->FireNodeExpanded();
		mp_oControl->Redraw();
		break;
	case MouseClick:
		break;
	case MouseDblClick:
		break;
	case MouseWheel:
		break;
	//case KeyDown:
	//	break;
	//case KeyUp:
	//	break;
	//case KeyPress:
	//	break;
	}
}

void clsTreeview::mp_EO_ROWMOVEMENT(LONG yMouseKeyBoardEvent,LONG X,LONG Y)
{
	clsRow* oDestinationRow = NULL;
	if (mp_oControl->GetAllowRowMove() == FALSE) 
	{
		return;
	}
	switch (yMouseKeyBoardEvent) 
	{
	case MouseHover:
		mp_oControl->MouseKeyboardEvents->mp_SetCursor(CT_MOVEROW);
		break;
	case MouseDown:
		s_rowMVT.lRowIndex = mp_oControl->GetMathLib()->GetRowIndexByPosition(Y);
		mp_oControl->GetObjectStateChangedEventArgs()->Clear();
		mp_oControl->GetObjectStateChangedEventArgs()->SetEventTarget((LONG)EVT_ROW);
		mp_oControl->GetObjectStateChangedEventArgs()->SetIndex(s_rowMVT.lRowIndex);
		mp_oControl->GetObjectStateChangedEventArgs()->SetCancel(FALSE);
        mp_oControl->GetObjectStateChangedEventArgs()->SetChangeType(CT_MOVE);
        mp_oControl->FireBeginObjectChange();
		break;
	case MouseMove:
		if (mp_oControl->GetObjectStateChangedEventArgs()->GetCancel() == FALSE) 
		{
			mp_oControl->clsG->mp_EraseReversibleFrames();
			mp_oControl->MouseKeyboardEvents->mp_DynamicRowMove(Y);
			s_rowMVT.lDestinationRowIndex = mp_oControl->GetMathLib()->GetRowIndexByPosition(Y);
			if (s_rowMVT.lDestinationRowIndex >= 1) 
			{
				mp_oControl->clsG->mp_EraseReversibleFrames();
				mp_oControl->GetObjectStateChangedEventArgs()->SetDestinationIndex(s_rowMVT.lDestinationRowIndex);
				mp_oControl->FireObjectChange();
				if (mp_oControl->GetObjectStateChangedEventArgs()->GetCancel() == FALSE) 
				{
					oDestinationRow = (clsRow*)mp_oControl->GetRows()->mp_oCollection->m_oReturnArrayElement(s_rowMVT.lDestinationRowIndex);
					mp_oControl->MouseKeyboardEvents->mp_DrawMovingReversibleFrame(oDestinationRow->GetLeft(), oDestinationRow->GetTop(), oDestinationRow->GetRight(), oDestinationRow->GetBottom(), FCT_NORMAL);
				}
			}
		}

		break;
	case MouseUp:
		if (mp_oControl->GetObjectStateChangedEventArgs()->GetCancel() == FALSE) 
		{
			mp_oControl->clsG->mp_EraseReversibleFrames();
			mp_oControl->FireEndObjectChange();
			if (mp_oControl->GetObjectStateChangedEventArgs()->GetCancel() == FALSE) 
			{
				s_rowMVT.lDestinationRowIndex = mp_oControl->GetMathLib()->GetRowIndexByPosition(Y);
				if (s_rowMVT.lDestinationRowIndex >= 1 && (s_rowMVT.lRowIndex != s_rowMVT.lDestinationRowIndex)) 
				{
					clsRow* oRow = NULL;
					oRow = (clsRow*)mp_oControl->GetRows()->mp_oCollection->m_oReturnArrayElement(s_rowMVT.lRowIndex);
					oDestinationRow = (clsRow*)mp_oControl->GetRows()->mp_oCollection->m_oReturnArrayElement(s_rowMVT.lDestinationRowIndex);
					oRow->GetNode()->SetDepth(oDestinationRow->GetNode()->GetDepth());
					mp_oControl->SetSelectedRowIndex(mp_oControl->GetRows()->mp_oCollection->m_lCopyAndMoveItems(s_rowMVT.lRowIndex, s_rowMVT.lDestinationRowIndex));
					mp_oControl->FireCompleteObjectChange();
				}
			}
		}

		mp_oControl->Redraw();
		mp_oControl->MouseKeyboardEvents->mp_SetCursor(CT_NORMAL);
		break;
	case MouseClick:
		break;
	case MouseDblClick:
		break;
	case MouseWheel:
		break;
	//case KeyDown:
	//	break;
	//case KeyUp:
	//	break;
	//case KeyPress:
	//	break;
	}
}

void clsTreeview::mp_EO_ROWSIZING(LONG yMouseKeyBoardEvent,LONG X,LONG Y)
{
	clsRow* oRow = NULL;
	if (mp_oControl->GetAllowRowSize() == FALSE) 
	{
		return;
	}
	switch (yMouseKeyBoardEvent) 
	{
	case MouseHover:
		mp_oControl->MouseKeyboardEvents->mp_SetCursor(CT_ROWHEIGHT);
		break;
	case MouseDown:
		s_rowSZ.lRowIndex = mp_oControl->GetMathLib()->GetRowIndexByPosition(Y);
		mp_oControl->GetObjectStateChangedEventArgs()->Clear();
		mp_oControl->GetObjectStateChangedEventArgs()->SetEventTarget((LONG)EVT_ROW);
		mp_oControl->GetObjectStateChangedEventArgs()->SetIndex(s_rowSZ.lRowIndex);
		mp_oControl->GetObjectStateChangedEventArgs()->SetCancel(FALSE);
        mp_oControl->GetObjectStateChangedEventArgs()->SetChangeType(CT_SIZE);
        mp_oControl->FireBeginObjectChange();
		break;
	case MouseMove:
		if (mp_oControl->GetObjectStateChangedEventArgs()->GetCancel() == FALSE) 
		{
			mp_oControl->clsG->mp_EraseReversibleFrames();
			mp_oControl->FireBeginObjectChange();
			if (mp_oControl->GetObjectStateChangedEventArgs()->GetCancel() == FALSE) 
			{
				mp_oControl->MouseKeyboardEvents->mp_DrawMovingReversibleFrame(0, Y, mp_oControl->clsG->Width(), Y + 2, FCT_NORMAL);
			}
		}

		break;
	case MouseUp:
		if (mp_oControl->GetObjectStateChangedEventArgs()->GetCancel() == FALSE) 
		{
			mp_oControl->clsG->mp_EraseReversibleFrames();
			mp_oControl->FireEndObjectChange();
			if (mp_oControl->GetObjectStateChangedEventArgs()->GetCancel() == FALSE) 
			{
				oRow = (clsRow*)mp_oControl->GetRows()->mp_oCollection->m_oReturnArrayElement(s_rowSZ.lRowIndex);
				oRow->SetHeight(oRow->GetHeight() + (Y - oRow->GetBottom()));
				if (oRow->GetHeight() < mp_oControl->GetMinRowHeight()) 
				{
					oRow->SetHeight(mp_oControl->GetMinRowHeight());
				}
				mp_oControl->FireCompleteObjectChange();
			}
		}

		mp_oControl->Redraw();
		mp_oControl->MouseKeyboardEvents->mp_SetCursor(CT_NORMAL);
		break;
	case MouseClick:
		break;
	case MouseDblClick:
		break;
	case MouseWheel:
		break;
	//case KeyDown:
	//	break;
	//case KeyUp:
	//	break;
	//case KeyPress:
	//	break;
	}
}

void clsTreeview::mp_EO_ROWSELECTION(LONG yMouseKeyBoardEvent,LONG X,LONG Y)
{
	clsRow* oRow = NULL;
	switch (yMouseKeyBoardEvent) 
	{
	case MouseHover:
		mp_oControl->MouseKeyboardEvents->mp_SetCursor(CT_NORMAL);
		break;
	case MouseDown:
		s_rowSEL.lRowIndex = mp_oControl->GetMathLib()->GetRowIndexByPosition(Y);
		s_rowSEL.lCellIndex = mp_oControl->GetMathLib()->GetCellIndexByPosition(X);
		break;
	case MouseMove:
		break;
	case MouseUp:
		mp_oControl->SetSelectedRowIndex(s_rowSEL.lRowIndex);
		mp_oControl->SetSelectedCellIndex(s_rowSEL.lCellIndex);
		oRow = (clsRow*)mp_oControl->GetRows()->mp_oCollection->m_oReturnArrayElement(mp_oControl->GetSelectedRowIndex());
		mp_oControl->GetObjectSelectedEventArgs()->Clear();
		mp_oControl->GetObjectSelectedEventArgs()->SetEventTarget((LONG)EVT_ROW);
		mp_oControl->GetObjectSelectedEventArgs()->SetObjectIndex(mp_oControl->GetSelectedRowIndex());
		mp_oControl->FireObjectSelected();
		mp_oControl->Redraw();
		mp_oControl->MouseKeyboardEvents->mp_SetCursor(CT_NORMAL);
		break;
	case MouseClick:
		break;
	case MouseDblClick:
		break;
	case MouseWheel:
		break;
	//case KeyDown:
	//	break;
	//case KeyUp:
	//	break;
	//case KeyPress:
	//	break;
	}
}

BOOL clsTreeview::mp_bOverCheckBox(LONG X,LONG Y)
{
	LONG lIndex = 0;
	clsNode* oNode = NULL;
	clsRow* oRow = NULL;
	BOOL bReturn = FALSE;
	if (mp_bCheckBoxes == FALSE) 
	{
		return FALSE;
	}
	bReturn = FALSE;
	for (lIndex = 1; lIndex <= mp_oControl->GetRows()->GetCount(); lIndex++) 
	{
		oRow = (clsRow*)mp_oControl->GetRows()->mp_oCollection->m_oReturnArrayElement(lIndex);
		oNode = oRow->GetNode();
		if (oRow->GetClientAreaVisibility() == VS_INSIDEVISIBLEAREA && X >= (oNode->GetCheckBoxLeft()) && X <= (oNode->GetCheckBoxLeft() + 13) && Y <= (oNode->GetYCenter() + 6) && Y >= (oNode->GetYCenter() - 7)) 
		{
			bReturn = TRUE;
		}
	}
	return bReturn;
}

BOOL clsTreeview::mp_bOverPlusMinusSign(LONG X,LONG Y)
{
	LONG lIndex = 0;
	clsNode* oNode = NULL;
	clsRow* oRow = NULL;
	BOOL bReturn = FALSE;
	if (mp_yTreeviewMode == TM_NONE)
	{
		return FALSE;
	}
	bReturn = FALSE;
	for (lIndex = 1; lIndex <= mp_oControl->GetRows()->GetCount(); lIndex++) 
	{
		oRow = (clsRow*)mp_oControl->GetRows()->mp_oCollection->m_oReturnArrayElement(lIndex);
		oNode = oRow->GetNode();
		if (oRow->GetClientAreaVisibility() == VS_INSIDEVISIBLEAREA && X >= (oNode->GetLeft() - 5) && X <= (oNode->GetLeft() + 5) && Y <= (oNode->GetYCenter() + 5) && Y >= (oNode->GetYCenter() - 5)) 
		{
			bReturn = TRUE;
		}
	}
	return bReturn;
}

void clsTreeview::Draw(void)
{
    if (mp_oControl->GetTreeviewColumnIndex() == 0)
    {
        return;
    }
    if (mp_oControl->GetColumns()->Item(CStr(mp_oControl->GetTreeviewColumnIndex()))->GetVisible() == FALSE)
    {
        return;
    }
    mp_oControl->clsG->mp_ClipRegion(GetLeftTrim(), mp_oControl->GetCurrentViewObject()->GetClientArea()->GetTop(), GetRightTrim(), mp_oControl->clsG->Height() - Getmt_BorderThickness() - 1, FALSE);
    mp_oControl->GetRows()->NodesDrawBackground();
    mp_oControl->clsG->mp_ClipRegion(GetLeftTrim(), mp_oControl->GetCurrentViewObject()->GetClientArea()->GetTop(), GetRightTrim() - 2, mp_oControl->clsG->Height() - Getmt_BorderThickness() - 1, FALSE);
	mp_oControl->GetRows()->NodesDraw();
	mp_oControl->GetRows()->NodesDrawTreeLines();
	mp_oControl->GetRows()->NodesDrawElements();
    mp_oControl->clsG->mp_ClearClipRegion();
}

void clsTreeview::ClearSelections(void)
{
	mp_oControl->SetSelectedRowIndex(0);
}

BOOL clsTreeview::mp_bCursorEditTextNode(int X, int Y)
{
	clsNode* oNode;
	clsRow* oRow;
	oRow = mp_oControl->GetRows()->Item(CStr(mp_oControl->GetSelectedRowIndex()));
	oNode = oRow->GetNode();
	if (oNode->GetAllowTextEdit() == TRUE)
	{
		if (X >= oNode->mp_lTextLeft && X <= oNode->mp_lTextRight)
		{
			if (Y >= oNode->mp_lTextTop && Y <= oNode->mp_lTextBottom)
			{
				mp_oControl->MouseKeyboardEvents->mp_SetCursor(CT_IBEAM);
				return TRUE;
			}
		}
	}
	mp_oControl->MouseKeyboardEvents->mp_SetCursor(CT_NORMAL);
	return FALSE;
}

BOOL clsTreeview::mp_bShowEditTextNode(int X, int Y)
{
	clsNode* oNode;
	clsRow* oRow;
	oRow = mp_oControl->GetRows()->Item(CStr(mp_oControl->GetSelectedRowIndex()));
	oNode = oRow->GetNode();
	if (oNode->GetAllowTextEdit() == TRUE)
	{
		if (X >= oNode->mp_lTextLeft && X <= oNode->mp_lTextRight)
		{
			if (Y >= oNode->mp_lTextTop && Y <= oNode->mp_lTextBottom)
			{
				mp_oControl->mp_oTextBox->Initialize(mp_oControl->GetSelectedRowIndex(), 0, TOT_NODE, X, Y);
				return TRUE;
			}
		}
	}
	return FALSE;
} 

Gdiplus::Color clsTreeview::GetExpandIconForeColor(void)
{
	return mp_clrExpandIconForeColor;
}

void clsTreeview::SetExpandIconForeColor(Gdiplus::Color Value)
{
	mp_clrExpandIconForeColor = Value;
}

Gdiplus::Color clsTreeview::GetExpandIconBackColor(void)
{
	return mp_clrExpandIconBackColor;
}

void clsTreeview::SetExpandIconBackColor(Gdiplus::Color Value)
{
	mp_clrExpandIconBackColor = Value;
}

Gdiplus::Color clsTreeview::GetExpandIconDropShadowColor(void)
{
	return mp_clrExpandIconDropShadowColor;
}

void clsTreeview::SetExpandIconDropShadowColor(Gdiplus::Color Value)
{
	mp_clrExpandIconDropShadowColor = Value;
}

Gdiplus::Color clsTreeview::GetCollapseIconForeColor(void)
{
	return mp_clrCollapseIconForeColor;
}

void clsTreeview::SetCollapseIconForeColor(Gdiplus::Color Value)
{
	mp_clrCollapseIconForeColor = Value;
}

Gdiplus::Color clsTreeview::GetCollapseIconDropShadowColor(void)
{
	return mp_clrCollapseIconDropShadowColor;
}

void clsTreeview::SetCollapseIconDropShadowColor(Gdiplus::Color Value)
{
	mp_clrCollapseIconDropShadowColor = Value;
}

void clsTreeview::SetXML(CString sXML)
{
	clsXML oXML(mp_oControl, "Treeview");
	oXML.SetXML(sXML);
	oXML.InitializeReader();
	oXML.ReadProperty("BackColor", mp_clrBackColor);
	oXML.ReadProperty("CheckBoxBorderColor", mp_clrCheckBoxBorderColor);
	oXML.ReadProperty("CheckBoxBackColor", mp_clrCheckBoxBackColor);
	oXML.ReadProperty("CheckBoxBackgroundMode", mp_yCheckBoxBackgroundMode);
	oXML.ReadProperty("CheckBoxes", mp_bCheckBoxes);
	oXML.ReadProperty("CheckBoxMarkColor", mp_clrCheckBoxMarkColor);
	oXML.ReadProperty("ExpansionOnSelection", mp_bExpansionOnSelection);
	oXML.ReadProperty("FullColumnSelect", mp_bFullColumnSelect);
	oXML.ReadProperty("Images", mp_bImages);
	oXML.ReadProperty("Indentation", mp_lIndentation);
	oXML.ReadProperty("PathSeparator", mp_sPathSeparator);
	oXML.ReadProperty("PlusMinusBorderColor", mp_clrPlusMinusBorderColor);
	oXML.ReadProperty("PlusMinusSignColor", mp_clrPlusMinusSignColor);
	oXML.ReadProperty("PlusMinusBackColor", mp_clrPlusMinusBackColor);
    oXML.ReadProperty("PlusMinusBackgroundMode", mp_yPlusMinusBackgroundMode);
	oXML.ReadProperty("TreeviewMode", mp_yTreeviewMode);
	oXML.ReadProperty("SelectedBackColor", mp_clrSelectedBackColor);
	oXML.ReadProperty("SelectedForeColor", mp_clrSelectedForeColor);
	oXML.ReadProperty("TreeLineColor", mp_clrTreeLineColor);
	oXML.ReadProperty("TreeLines", mp_bTreeLines);
    oXML.ReadProperty("ExpandIconForeColor", mp_clrExpandIconForeColor);
    oXML.ReadProperty("ExpandIconBackColor", mp_clrExpandIconBackColor);
    oXML.ReadProperty("ExpandIconDropShadowColor", mp_clrExpandIconDropShadowColor);
    oXML.ReadProperty("CollapseIconForeColor", mp_clrCollapseIconForeColor);
    oXML.ReadProperty("CollapseIconDropShadowColor", mp_clrCollapseIconDropShadowColor);
}

CString clsTreeview::GetXML(void)
{
	clsXML oXML(mp_oControl, "Treeview");
	oXML.InitializeWriter();
	oXML.WriteProperty("BackColor", mp_clrBackColor);
	oXML.WriteProperty("CheckBoxBorderColor", mp_clrCheckBoxBorderColor);
	oXML.WriteProperty("CheckBoxBackColor", mp_clrCheckBoxBackColor);
	oXML.WriteProperty("CheckBoxBackgroundMode", mp_yCheckBoxBackgroundMode);
	oXML.WriteProperty("CheckBoxes", mp_bCheckBoxes);
	oXML.WriteProperty("CheckBoxMarkColor", mp_clrCheckBoxMarkColor);
	oXML.WriteProperty("ExpansionOnSelection", mp_bExpansionOnSelection);
	oXML.WriteProperty("FullColumnSelect", mp_bFullColumnSelect);
	oXML.WriteProperty("Images", mp_bImages);
	oXML.WriteProperty("Indentation", mp_lIndentation);
	oXML.WriteProperty("PathSeparator", mp_sPathSeparator);
	oXML.WriteProperty("PlusMinusBorderColor", mp_clrPlusMinusBorderColor);
	oXML.WriteProperty("PlusMinusSignColor", mp_clrPlusMinusSignColor);
	oXML.WriteProperty("PlusMinusBackColor", mp_clrPlusMinusBackColor);
    oXML.WriteProperty("PlusMinusBackgroundMode", mp_yPlusMinusBackgroundMode);
	oXML.WriteProperty("TreeviewMode", mp_yTreeviewMode);
	oXML.WriteProperty("SelectedBackColor", mp_clrSelectedBackColor);
	oXML.WriteProperty("SelectedForeColor", mp_clrSelectedForeColor);
	oXML.WriteProperty("TreeLineColor", mp_clrTreeLineColor);
	oXML.WriteProperty("TreeLines", mp_bTreeLines);
    oXML.WriteProperty("ExpandIconForeColor", mp_clrExpandIconForeColor);
    oXML.WriteProperty("ExpandIconBackColor", mp_clrExpandIconBackColor);
    oXML.WriteProperty("ExpandIconDropShadowColor", mp_clrExpandIconDropShadowColor);
    oXML.WriteProperty("CollapseIconForeColor", mp_clrCollapseIconForeColor);
    oXML.WriteProperty("CollapseIconDropShadowColor", mp_clrCollapseIconDropShadowColor);
	return oXML.GetXML();
}