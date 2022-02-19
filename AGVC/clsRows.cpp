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
#include "clsRows.h"



IMPLEMENT_DYNCREATE(clsRows, CCmdTarget)

//{C1A6AEAA-2967-4AFD-ADC4-9EEF49913F82}
static const IID IID_IclsRows = { 0xC1A6AEAA, 0x2967, 0x4AFD, { 0xAD, 0xC4, 0x9E, 0xEF, 0x49, 0x91, 0x3F, 0x82} };

//{925007DE-388B-4553-80D0-F6A575C8A422}
IMPLEMENT_OLECREATE_FLAGS(clsRows, "AGVC.clsRows", afxRegApartmentThreading, 0x925007DE, 0x388B, 0x4553, 0x80, 0xD0, 0xF6, 0xA5, 0x75, 0xC8, 0xA4, 0x22)

BEGIN_DISPATCH_MAP(clsRows, CCmdTarget)
	DISP_PROPERTY_EX_ID(clsRows, "Count", 1, odl_GetCount, SetNotSupported, VT_I4) 
	DISP_PROPERTY_PARAM_ID(clsRows, "Item", 2, odl_Item, SetNotSupported, VT_DISPATCH, VTS_BSTR) 
	DISP_FUNCTION_ID(clsRows, "Clear", 3, odl_Clear, VT_EMPTY, VTS_NONE) 
	DISP_FUNCTION_ID(clsRows, "Remove", 4, odl_Remove, VT_EMPTY, VTS_BSTR) 
	DISP_FUNCTION_ID(clsRows, "MoveRow", 5, odl_MoveRow, VT_EMPTY, VTS_I4 VTS_I4) 
	DISP_FUNCTION_ID(clsRows, "GetXML", 6, odl_GetXML, VT_BSTR, VTS_NONE) 
	DISP_FUNCTION_ID(clsRows, "SetXML", 7, odl_SetXML, VT_EMPTY, VTS_BSTR)
	DISP_FUNCTION_ID(clsRows, "Add", 8, odl_Add, VT_DISPATCH, VTS_BSTR VTS_BSTR VTS_BOOL VTS_BOOL VTS_BSTR)
	DISP_FUNCTION_ID(clsRows, "SortRows", 9, odl_SortRows, VT_EMPTY, VTS_BSTR VTS_BOOL VTS_I4 VTS_I4 VTS_I4)
	DISP_FUNCTION_ID(clsRows, "SortCells", 10, odl_SortCells, VT_EMPTY, VTS_I4 VTS_BSTR VTS_BOOL VTS_I4 VTS_I4 VTS_I4)
	DISP_FUNCTION_ID(clsRows, "UpdateTree", 11, odl_UpdateTree, VT_EMPTY, VTS_NONE)
	DISP_FUNCTION_ID(clsRows, "BeginLoad", 12, odl_BeginLoad, VT_EMPTY, VTS_BOOL) 
	DISP_FUNCTION_ID(clsRows, "Load", 13, odl_Load, VT_DISPATCH, VTS_BSTR) 
	DISP_FUNCTION_ID(clsRows, "EndLoad", 14, odl_EndLoad, VT_EMPTY, VTS_NONE)
	DISP_FUNCTION_ID(clsRows, "VerifyTree", 15, odl_VerifyTree, VT_BOOL, VTS_NONE)
	DISP_DEFVALUE(clsRows, "Item")
END_DISPATCH_MAP()

BEGIN_INTERFACE_MAP(clsRows, CCmdTarget)
	INTERFACE_PART(clsRows, IID_IclsRows, Dispatch)
END_INTERFACE_MAP()

BEGIN_MESSAGE_MAP(clsRows, CCmdTarget)
END_MESSAGE_MAP()

clsRows::clsRows(CActiveGanttVCCtl* oControl)
{
	EnableAutomation();
	AfxOleLockApp();
	mp_oControl = oControl;
	mp_oCollection = new clsCollectionBase(mp_oControl, "Row");
	mp_lLoadIndex = -1;
}

clsRows::clsRows()
{
	AfxMessageBox(_T("Invalid constructor has been invoked for clsRows. Please inform support@sourcecodestore.com of this error"), MB_OK | MB_ICONSTOP | MB_APPLMODAL);
}

clsRows::~clsRows()
{
	delete mp_oCollection;
	mp_oCollection = NULL;
	AfxOleUnlockApp();
}

void clsRows::OnFinalRelease()
{
	CCmdTarget::OnFinalRelease();
}

LONG clsRows::GetCount(void)
{
	return mp_oCollection->m_lCount();
}

clsCollectionBase* clsRows::GetCollection(void)
{
	return mp_oCollection;
}

LONG clsRows::GetTopOffset(void)
{
	return mp_lTopOffset;
}

void clsRows::SetTopOffset(LONG Value)
{
	mp_lTopOffset = Value;
}

LONG clsRows::GetRealFirstVisibleRow(void)
{
	return RealIndex(mp_oControl->GetVerticalScrollBar()->GetValue());
}

clsRow* clsRows::Item(CString Index)
{
	clsRow *oRow;
	oRow = (clsRow*)mp_oCollection->m_oItem(Index, ROWS_ITEM_1, ROWS_ITEM_2, ROWS_ITEM_3, ROWS_ITEM_4);
	return oRow;
}

clsRow* clsRows::Add(CString Key,CString Text,BOOL MergeCells,BOOL Container,CString StyleIndex)
{
	mp_oCollection->SetAddMode(TRUE);
	clsRow* oRow = new clsRow(mp_oControl);
	LONG lIndex = 0;
	oRow->SetKey(Key);
	oRow->SetText(Text);
	oRow->SetMergeCells(MergeCells);
	oRow->SetContainer(Container);
	oRow->SetStyleIndex(StyleIndex);
	mp_oCollection->m_Add(oRow, Key, ROWS_ADD_1, ROWS_ADD_2, TRUE, ROWS_ADD_3);
	for (lIndex = 1; lIndex <= mp_oControl->GetColumns()->GetCount(); lIndex++) 
	{
		oRow = (clsRow*)mp_oControl->GetRows()->mp_oCollection->m_oReturnArrayElement(mp_oControl->GetRows()->GetCount());
		oRow->GetCells()->Add();
	}
	mp_oControl->GetVerticalScrollBar()->Update();
	return oRow;
}

void clsRows::Clear(void)
{
	mp_oControl->GetTasks()->Clear();
	mp_oCollection->m_Clear();
	mp_oControl->SetSelectedRowIndex(0);
	mp_oControl->GetVerticalScrollBar()->Reset();

}

void clsRows::Remove(CString Index)
{
	CString sRIndex = "";
	CString sRKey = "";
	mp_oCollection->m_GetKeyAndIndex(Index, &sRKey, &sRIndex);
	mp_oControl->GetTasks()->mp_oCollection->m_CollectionRemoveWhere("RowKey", sRKey, sRIndex);
	mp_oCollection->m_Remove(Index, ROWS_REMOVE_1, ROWS_REMOVE_2, ROWS_REMOVE_3, ROWS_REMOVE_4);
	mp_oControl->SetSelectedRowIndex(0);
	mp_oControl->SetSelectedTaskIndex(0);
	mp_oControl->GetVerticalScrollBar()->Update();
}

void clsRows::MoveRow(LONG OriginRowIndex,LONG DestRowIndex)
{
	if (OriginRowIndex < 1 || OriginRowIndex > GetCount()) 
	{
		return;
	}
	if (DestRowIndex < 1 || DestRowIndex > GetCount()) 
	{
		return;
	}
	if (DestRowIndex == OriginRowIndex) 
	{
		return;
	}
	mp_oCollection->m_lCopyAndMoveItems(OriginRowIndex, DestRowIndex);
}

void clsRows::SortRows(CString PropertyName,BOOL Descending,LONG SortType,LONG StartIndex,LONG EndIndex)
{
	if (StartIndex == -1) 
	{
		StartIndex = 1;
	}
	if (EndIndex == -1) 
	{
		EndIndex = GetCount();
	}
	if (GetCount() == 0) 
	{
		return;
	}
	if (StartIndex < 1 || StartIndex > GetCount()) 
	{
		return;
	}
	if (EndIndex < 1 || EndIndex > GetCount()) 
	{
		return;
	}
	if (EndIndex == StartIndex) 
	{
		return;
	}
	mp_oCollection->m_Sort(PropertyName, Descending, (E_SORTTYPE)SortType, StartIndex, EndIndex);
}

void clsRows::SortCells(LONG CellIndex,CString PropertyName,BOOL Descending,LONG SortType,LONG StartIndex,LONG EndIndex)
{
	if (StartIndex == -1) 
	{
		StartIndex = 1;
	}
	if (EndIndex == -1) 
	{
		EndIndex = GetCount();
	}
	if (GetCount() == 0) 
	{
		return;
	}
	if (StartIndex < 1 || StartIndex > GetCount()) 
	{
		return;
	}
	if (EndIndex < 1 || EndIndex > GetCount()) 
	{
		return;
	}
	if (EndIndex == StartIndex) 
	{
		return;
	}
	if (CellIndex < 1 || CellIndex > mp_oControl->GetColumns()->GetCount()) 
	{
		return;
	}
	mp_oCollection->m_SortCells(CellIndex, PropertyName, Descending, (E_SORTTYPE)SortType, StartIndex, EndIndex);
}

void clsRows::ClearCells(void)
{
	LONG lIndex = 0;
	clsRow* oRow = NULL;
	for (lIndex = 1; lIndex <= mp_oCollection->m_lCount(); lIndex++) 
	{
		oRow = (clsRow*)mp_oCollection->m_oReturnArrayElement(lIndex);
		oRow->GetCells()->Clear();
	}
}

LONG clsRows::Height(void)
{
	return Height(GetCount());
}

LONG clsRows::Height(LONG lRow)
{
	LONG lBuffer = 0;
	LONG lIndex = 0;
	clsRow* oRow = NULL;
	if (GetCount() == 0 || lRow == 0) 
	{
		return 0;
	}

	BOOL bChildrenHidden = FALSE;
	LONG lCurrentDepth = 0;
	for (lIndex = 1; lIndex <= lRow; lIndex++) 
	{
		BOOL bHidden = FALSE;
		oRow = (clsRow*)mp_oCollection->m_oReturnArrayElement(lIndex);
		if (oRow->GetNode()->GetDepth() == 0)
		{
			bHidden = FALSE;
		}
		if (bChildrenHidden == TRUE)
		{
			bHidden = TRUE;
		}
		if (oRow->GetNode()->GetDepth() < lCurrentDepth)
		{
			lCurrentDepth = oRow->GetNode()->GetDepth();
			bHidden = FALSE;
			bChildrenHidden = FALSE;
		}
		if (bHidden == FALSE)
		{

			lBuffer = lBuffer + oRow->GetHeight() + 1;
		}
		if (oRow->GetNode()->GetExpanded() == FALSE && bChildrenHidden == FALSE)
		{
			lCurrentDepth = oRow->GetNode()->GetDepth() + 1;
			bChildrenHidden = TRUE;
		}
	}
	return lBuffer;
}

LONG clsRows::CalculateHeight(LONG StartIndex,LONG EndIndex)
{
	LONG lBuffer = 0;
	LONG lIndex = 0;
	clsRow* oRow = NULL;
	if (StartIndex == 0) 
	{
		return 0;
	}
	for (lIndex = StartIndex; lIndex <= EndIndex; lIndex++) 
	{
		oRow = (clsRow*)mp_oCollection->m_oReturnArrayElement(lIndex);
		lBuffer = lBuffer + oRow->GetHeight();
	}
	return lBuffer;
}

LONG clsRows::CalculateRows(LONG StartIndex,LONG Height)
{
	LONG lBuffer = 0;
	LONG lIndex = 0;
	clsRow* oRow = NULL;
	LONG lRows = 0;
	lRows = 1;
	if (StartIndex == 0) 
	{
		return lRows;
	}
	for (lIndex = StartIndex; lIndex <= GetCount(); lIndex++) 
	{
		oRow = (clsRow*)mp_oCollection->m_oReturnArrayElement(lIndex);
		lBuffer = lBuffer + oRow->GetHeight();
		if (lBuffer > Height) 
		{
			break; 
		}
		lRows = lRows + 1;
	}
	return lRows;
}

void clsRows::PositionRows(void)
{
	clsRow* oRow = NULL;
	LONG lRowIndex = 0;
	LONG lBottomBuff = 0;
	clsClientArea* oClientArea = NULL;
	oClientArea = mp_oControl->GetCurrentViewObject()->GetClientArea();
	if (GetCount() == 0) 
	{
		oClientArea->Setf_LastVisibleRow(0);
		mp_lTopOffset = mp_oControl->GetCurrentViewObject()->GetClientArea()->GetTop();
		return;
	}
	else 
	{
		mp_lTopOffset = 0;
	}
	for (lRowIndex = (mp_lRealFirstVisibleRow - 1); lRowIndex >= 1; lRowIndex += -1) 
	{
		oRow = (clsRow*)mp_oCollection->m_oReturnArrayElement(lRowIndex);
		if (lRowIndex == (mp_lRealFirstVisibleRow - 1)) 
		{
			oRow->Setf_lBottom(mp_oControl->GetCurrentViewObject()->GetClientArea()->GetTop() - 1);
		}
		else 
		{
			oRow->Setf_lBottom(lBottomBuff - 1);
		}
		oRow->Setf_lTop(oRow->GetBottom() - oRow->GetHeight());
		lBottomBuff = oRow->GetTop();
		oRow->SetClientAreaVisibility(VS_ABOVEVISIBLEAREA);
	}
	for (lRowIndex = mp_lRealFirstVisibleRow; lRowIndex <= GetCount(); lRowIndex++) 
	{
		oRow = (clsRow*)mp_oCollection->m_oReturnArrayElement(lRowIndex);
		if (mp_lTopOffset < mp_oControl->Getmt_TableBottom()) 
		{
			if (lRowIndex == mp_lRealFirstVisibleRow) 
			{
				oRow->Setf_lTop(mp_oControl->GetCurrentViewObject()->GetClientArea()->GetTop());
			}
			else 
			{
				oRow->Setf_lTop(lBottomBuff + 1);
			}
			oRow->Setf_lBottom(oRow->GetTop() + oRow->GetHeight());
			lBottomBuff = oRow->GetBottom();
			mp_lTopOffset = oRow->GetBottom();
			oClientArea->Setf_LastVisibleRow(lRowIndex);
			oRow->SetClientAreaVisibility(VS_INSIDEVISIBLEAREA);
		}
		else 
		{
			break;
		}
	}
	for (lRowIndex = (oClientArea->GetLastVisibleRow() + 1); lRowIndex <= GetCount(); lRowIndex++) 
	{
		oRow = (clsRow*)mp_oCollection->m_oReturnArrayElement(lRowIndex);
		oRow->Setf_lTop(lBottomBuff + 1);
		oRow->Setf_lBottom(oRow->GetTop() + oRow->GetHeight());
		lBottomBuff = oRow->GetBottom();
		oRow->SetClientAreaVisibility(VS_BELOWVISIBLEAREA);
	}
}

LONG clsRows::NodeHeight(LONG lFirstVisibleNode)
{
	LONG lIndex = 0;
	LONG lReturn = 0;
	clsNode* oNode = NULL;
	clsRow* oRow = NULL;
	if (lFirstVisibleNode == 0) 
	{
		return 0;
	}
	for (lIndex = lFirstVisibleNode; lIndex <= GetCount(); lIndex++) 
	{
		oRow = (clsRow*)mp_oCollection->m_oReturnArrayElement(lIndex);
		oNode = oRow->GetNode();
		if (oNode->GetHidden() == FALSE) 
		{
			lReturn = lReturn + oNode->GetHeight() + 1;
		}
	}
	return lReturn;
}

void clsRows::InitializePosition(void)
{
	mp_lRealFirstVisibleRow = GetRealFirstVisibleRow();
}

LONG clsRows::RealIndex(LONG Index)
{
	LONG lIndex = 0;
	clsRow* oRow = NULL;
	clsNode* oNode = NULL;
	LONG lCount = 0;
	for (lIndex = 1; lIndex <= GetCount(); lIndex++) 
	{
		oRow = (clsRow*)mp_oControl->GetRows()->mp_oCollection->m_oReturnArrayElement(lIndex);
		oNode = oRow->GetNode();
		if (oNode->GetHidden() == FALSE) 
		{
			lCount = lCount + 1;
			if (lCount == Index) 
			{
				return lIndex;
			}
		}
	}
	return -1;
}

void clsRows::NodesDrawBackground(void)
{
	LONG lIndex = 0;
	clsNode* oNode = NULL;
	clsRow* oRow = NULL;
	clsCell* oCell = NULL;
	if (GetCount() == 0) 
	{
		return;
	}
	for (lIndex = mp_lRealFirstVisibleRow; lIndex <= GetCount(); lIndex++) 
	{
		oRow = (clsRow*)mp_oCollection->m_oReturnArrayElement(lIndex);
		oCell = (clsCell*)oRow->GetCells()->mp_oCollection->m_oReturnArrayElement(mp_oControl->GetTreeviewColumnIndex());
		oNode = oRow->GetNode();
		if (mp_oControl->GetTreeview()->GetFullColumnSelect() == TRUE)
		{
			mp_oControl->clsG->mp_DrawItemRow(oCell->GetLeft(), oCell->GetTop(), oCell->GetRight() - 1, oCell->GetBottom(), "", (oRow->GetIndex() == mp_oControl->GetSelectedRowIndex() && mp_oControl->GetTreeviewColumnIndex() == mp_oControl->GetSelectedCellIndex()), oCell->GetImage(), oCell->GetLeftTrim(), oCell->GetRightTrim(), oNode->GetStyle(), oRow);
		}
		else
		{
			mp_oControl->clsG->mp_DrawItemRow(oCell->GetLeft(), oCell->GetTop(), oCell->GetRight() - 1, oCell->GetBottom(), "", FALSE, oCell->GetImage(), oCell->GetLeftTrim(), oCell->GetRightTrim(), oNode->GetStyle(), oRow);
	        if (oRow->GetIndex() == mp_oControl->GetSelectedRowIndex() && mp_oControl->GetTreeviewColumnIndex() == mp_oControl->GetSelectedCellIndex() && oNode->GetStyle()->GetSelectionRectangleStyle()->GetVisible() == TRUE)
            {
                mp_oControl->clsG->mp_DrawSelectionRectangle(oNode->Getmt_TextLeft(), oNode->GetTop(), mp_oControl->GetSplitter()->GetLeft(), oNode->GetBottom() - 1, oNode->GetStyle());
            }
		}
	}
}

void clsRows::NodesDrawTreeLines(void)
{
	LONG lIndex = 0;
	clsNode* oNode = NULL;
	clsRow* oRow = NULL;
	clsNode* oRelative = NULL;
	if (GetCount() == 0 || mp_oControl->GetTreeview()->GetTreeLines() == FALSE) 
	{
		return;
	}
	for (lIndex = mp_lRealFirstVisibleRow; lIndex <= GetCount(); lIndex++) 
	{
		oRow = (clsRow*)mp_oCollection->m_oReturnArrayElement(lIndex);
		oNode = oRow->GetNode();
		if ((oRow->GetClientAreaVisibility() == VS_INSIDEVISIBLEAREA || oRow->GetClientAreaVisibility() == VS_BELOWVISIBLEAREA) && oNode->GetHidden() == FALSE) 
		{
			if (lIndex <= mp_oControl->GetCurrentViewObject()->GetClientArea()->GetLastVisibleRow())
			{
				if (oNode->GetCheckBoxVisible() == TRUE)
				{
					mp_oControl->clsG->mp_DrawLine(oNode->GetLeft(), oNode->GetYCenter(), oNode->GetLeft() + 15, oNode->GetYCenter(), LT_NORMAL, mp_oControl->GetTreeview()->GetTreeLineColor(), LDS_SOLID);
				}
				else if (oNode->GetImageVisible() == TRUE)
				{
					mp_oControl->clsG->mp_DrawLine(oNode->GetLeft(), oNode->GetYCenter(), oNode->GetLeft() + 15, oNode->GetYCenter(), LT_NORMAL, mp_oControl->GetTreeview()->GetTreeLineColor(), LDS_SOLID);
				}
				else if (oNode->GetImageVisible() == FALSE && oNode->GetCheckBoxVisible() == FALSE)
				{
					mp_oControl->clsG->mp_DrawLine(oNode->GetLeft(), oNode->GetYCenter(), oNode->Getmt_TextLeft(), oNode->GetYCenter(), LT_NORMAL, mp_oControl->GetTreeview()->GetTreeLineColor(), LDS_SOLID);
				}
			}
			if (oNode->GetIndex() == 1) 
			{
				//Dont Draw anything
			}
			else 
			{
				oRelative = NULL;
				if (oNode->IsFirstSibling()) 
				{
					oRelative = oNode->Parent();
					if (oRelative->GetRow()->GetIndex() > mp_oControl->GetCurrentViewObject()->GetClientArea()->GetLastVisibleRow())
					{
						oRelative = NULL;
					}
					else 
					{
						mp_oControl->clsG->mp_DrawLine(oNode->GetLeft(), oNode->GetYCenter(), oNode->GetLeft(), oRelative->Getmt_TextBottom(), LT_NORMAL, mp_oControl->GetTreeview()->GetTreeLineColor(), LDS_SOLID);
					}
				}
				else if (oNode->IsLastSibling())
				{
					oRelative = oNode->FirstSibling();
					if (oRelative->GetRow()->GetIndex() > mp_oControl->GetCurrentViewObject()->GetClientArea()->GetLastVisibleRow())
					{
						oRelative = NULL;
					}
					else
					{
						mp_oControl->clsG->mp_DrawLine(oNode->GetLeft(), oNode->GetYCenter(), oNode->GetLeft(), oRelative->GetYCenter(), LT_NORMAL, mp_oControl->GetTreeview()->GetTreeLineColor(), LDS_SOLID);
					}
				}
			}
		}
	}
}

void clsRows::NodesDraw(void)
{
	LONG lIndex = 0;
	clsNode* oNode = NULL;
	clsRow* oRow = NULL;
	if (GetCount() == 0) 
	{
		return;
	}
	for (lIndex = mp_lRealFirstVisibleRow; lIndex <= GetCount(); lIndex++) 
	{
		oRow = (clsRow*)mp_oCollection->m_oReturnArrayElement(lIndex);
		oNode = oRow->GetNode();
		if (oRow->GetClientAreaVisibility() == VS_INSIDEVISIBLEAREA && oNode->GetHidden() == FALSE) 
		{
			mp_oControl->clsG->mp_DrawAlignedText(oNode->Getmt_TextLeft(), oNode->Getmt_TextTop(), oNode->Getmt_TextRight(), oNode->Getmt_TextBottom(), oNode->GetText(), HAL_LEFT, VAL_CENTER, oNode->GetStyle()->GetForeColor(), oNode->GetStyle()->GetFont());
            if (oNode->GetText().GetLength() > 0)
            {
                oNode->mp_lTextLeft = mp_oControl->clsG->mp_oTextFinalLayout.Left;
                oNode->mp_lTextTop = mp_oControl->clsG->mp_oTextFinalLayout.Top;
                oNode->mp_lTextRight = mp_oControl->clsG->mp_oTextFinalLayout.Left + mp_oControl->clsG->mp_oTextFinalLayout.Width - 1;
                oNode->mp_lTextBottom = mp_oControl->clsG->mp_oTextFinalLayout.Top + mp_oControl->clsG->mp_oTextFinalLayout.Height - 1;
            }
            else
            {
                oNode->mp_lTextLeft = oNode->Getmt_TextLeft();
                oNode->mp_lTextTop = oNode->Getmt_TextTop();
                oNode->mp_lTextRight = oNode->Getmt_TextRight();
                oNode->mp_lTextBottom = oNode->Getmt_TextBottom();
            }
		}
	}
}

void clsRows::mp_DrawSign(BOOL bExpanded,LONG X,LONG Y, clsRow* oRow)
{
	if (mp_oControl->GetTreeview()->GetTreeviewMode() == TM_NONE) 
	{
		return;
	}
	else if (mp_oControl->GetTreeview()->GetTreeviewMode() == TM_PLUSMINUSSIGNS)
	{
		Y = Y - 1;
		if (mp_oControl->GetTreeview()->GetPlusMinusBackgroundMode() == FP_SOLID)
		{
			mp_oControl->clsG->mp_DrawLine(X - 4, Y - 3, X + 4, Y + 4, LT_FILLED, mp_oControl->GetTreeview()->GetPlusMinusBackColor(), LDS_SOLID);
		}
		else
		{
			mp_oControl->clsG->mp_DrawLine(X - 4, Y - 3, X + 4, Y + 4, LT_FILLED, oRow->Getf_RowBackColor(), LDS_SOLID);
		}
		mp_oControl->clsG->mp_DrawLine(X - 4, Y - 3, X + 4, Y + 5, LT_BORDER, mp_oControl->GetTreeview()->GetPlusMinusBorderColor(), LDS_SOLID);
		mp_oControl->clsG->mp_DrawLine(X - 2, Y + 1, X + 2, Y + 1, LT_NORMAL, mp_oControl->GetTreeview()->GetPlusMinusSignColor(), LDS_SOLID);
		if (bExpanded == FALSE) 
		{
			mp_oControl->clsG->mp_DrawLine(X, Y - 1, X, Y + 3, LT_NORMAL, mp_oControl->GetTreeview()->GetPlusMinusSignColor(), LDS_SOLID);
		}
	}
	else if (mp_oControl->GetTreeview()->GetTreeviewMode() == TM_EXPANDCOLLAPSEICONS)
	{
		if (bExpanded == FALSE) 
		{
			mp_oControl->clsG->mp_DrawExpandIcon(X - 3, Y - 4, mp_oControl->GetTreeview()->GetExpandIconForeColor(), mp_oControl->GetTreeview()->GetExpandIconDropShadowColor(), mp_oControl->GetTreeview()->GetExpandIconBackColor());
		}
		else
		{
			mp_oControl->clsG->mp_DrawCollapseIcon(X - 3, Y - 4, mp_oControl->GetTreeview()->GetCollapseIconForeColor(), mp_oControl->GetTreeview()->GetCollapseIconDropShadowColor());
		}
	}
}

LONG clsRows::HiddenRows(void)
{
	LONG lIndex = 0;
	clsRow* oRow = NULL;
	LONG lReturn = 0;


	BOOL bChildrenHidden = FALSE;
	LONG lCurrentDepth = 0;
	for (lIndex = 1; lIndex <= GetCount(); lIndex++) 
	{
		BOOL bHidden = FALSE;
		oRow = (clsRow*)mp_oControl->GetRows()->mp_oCollection->m_oReturnArrayElement(lIndex);
        if (oRow->GetNode()->GetDepth() == 0)
        {
            bHidden = FALSE;
        }
        if (bChildrenHidden == TRUE)
        {
            bHidden = TRUE;
        }
        if (oRow->GetNode()->GetDepth() < lCurrentDepth)
        {
            lCurrentDepth = oRow->GetNode()->GetDepth();
            bHidden = FALSE;
            bChildrenHidden = FALSE;
        }	
		if (bHidden == TRUE) 
		{
			lReturn = lReturn + 1;
		}
        if (oRow->GetNode()->GetExpanded() == FALSE && bChildrenHidden == FALSE)
        {
            lCurrentDepth = oRow->GetNode()->GetDepth() + 1;
            bChildrenHidden = TRUE;
        }
	}
	return lReturn;
}

void clsRows::mp_DrawCheckBox(clsNode* oNode)
{
	if (mp_oControl->GetTreeview()->GetCheckBoxes() == FALSE) 
	{
		return;
	}
	if (oNode->GetCheckBoxVisible() == TRUE) 
	{
		if (mp_oControl->GetTreeview()->GetCheckBoxBackgroundMode() == FP_SOLID)
		{
			mp_oControl->clsG->mp_DrawLine(oNode->GetCheckBoxLeft() + 1, oNode->GetYCenter() - 5, oNode->GetCheckBoxLeft() + 11, oNode->GetYCenter() + 5, LT_FILLED, mp_oControl->GetTreeview()->GetCheckBoxBackColor(), LDS_SOLID);
		}
        mp_oControl->clsG->mp_DrawLine(oNode->GetCheckBoxLeft() + 1, oNode->GetYCenter() - 5, oNode->GetCheckBoxLeft() + 11, oNode->GetYCenter() + 5, LT_BORDER, mp_oControl->GetTreeview()->GetCheckBoxBorderColor(), LDS_SOLID);
		if (oNode->GetChecked() == TRUE) 
		{
			mp_oControl->clsG->mp_DrawLine(oNode->GetCheckBoxLeft() + 3, oNode->GetYCenter(), oNode->GetCheckBoxLeft() + 3, oNode->GetYCenter() + 2, LT_NORMAL, mp_oControl->GetTreeview()->GetCheckBoxMarkColor(), LDS_SOLID);
			mp_oControl->clsG->mp_DrawLine(oNode->GetCheckBoxLeft() + 4, oNode->GetYCenter() + 1, oNode->GetCheckBoxLeft() + 4, oNode->GetYCenter() + 3, LT_NORMAL, mp_oControl->GetTreeview()->GetCheckBoxMarkColor(), LDS_SOLID);
			mp_oControl->clsG->mp_DrawLine(oNode->GetCheckBoxLeft() + 5, oNode->GetYCenter() + 2, oNode->GetCheckBoxLeft() + 5, oNode->GetYCenter() + 4, LT_NORMAL, mp_oControl->GetTreeview()->GetCheckBoxMarkColor(), LDS_SOLID);
			mp_oControl->clsG->mp_DrawLine(oNode->GetCheckBoxLeft() + 6, oNode->GetYCenter() + 1, oNode->GetCheckBoxLeft() + 6, oNode->GetYCenter() + 3, LT_NORMAL, mp_oControl->GetTreeview()->GetCheckBoxMarkColor(), LDS_SOLID);
			mp_oControl->clsG->mp_DrawLine(oNode->GetCheckBoxLeft() + 7, oNode->GetYCenter(), oNode->GetCheckBoxLeft() + 7, oNode->GetYCenter() + 2, LT_NORMAL, mp_oControl->GetTreeview()->GetCheckBoxMarkColor(), LDS_SOLID);
			mp_oControl->clsG->mp_DrawLine(oNode->GetCheckBoxLeft() + 8, oNode->GetYCenter() - 1, oNode->GetCheckBoxLeft() + 8, oNode->GetYCenter() + 1, LT_NORMAL, mp_oControl->GetTreeview()->GetCheckBoxMarkColor(), LDS_SOLID);
			mp_oControl->clsG->mp_DrawLine(oNode->GetCheckBoxLeft() + 9, oNode->GetYCenter() - 2, oNode->GetCheckBoxLeft() + 9, oNode->GetYCenter(), LT_NORMAL, mp_oControl->GetTreeview()->GetCheckBoxMarkColor(), LDS_SOLID);
		}
	}
}

void clsRows::mp_DrawImage(clsNode* oNode)
{
	clsImage* oImage = NULL;
	LONG lImageWidth = 0;
	LONG lImageHeight = 0;
	if (mp_oControl->GetTreeview()->GetImages() == FALSE) 
	{
		return;
	}
	if (oNode->GetTag() == "NP")
	{
		int y = 0;
	}
	if (oNode->GetImageVisible() == TRUE) 
	{
		if (oNode->GetExpanded() == TRUE && oNode->Children() > 0 && (oNode->GetExpandedImage()->GetImageP() != NULL)) 
		{
			oImage = oNode->GetExpandedImage();
		}
		else if (oNode->GetSelected() == TRUE && (oNode->GetSelectedImage()->GetImageP() != NULL)) 
		{
			oImage = oNode->GetSelectedImage();
		}
		else if ((oNode->GetImage()->GetImageP() != NULL)) 
		{
			oImage = oNode->GetImage();
		}
		if ((oImage->GetImageP() != NULL)) 
		{
			lImageWidth = oImage->GetWidth();
			lImageHeight = oImage->GetHeight();
			mp_oControl->clsG->mp_PaintImage(oImage, oNode->GetImageLeft(), oNode->GetImageTop(), oNode->GetImageRight(), oNode->GetImageBottom(), 0, 0, TRUE);
		}
	}
}

void clsRows::Draw(void)
{
	LONG lCellIndex = 0;
	LONG lRowIndex = 0;
	clsRow* oRow = NULL;
	clsColumn* oColumn = NULL;
	clsCell* oCell = NULL;
	LONG lTableBottom = mp_oControl->Getmt_TableBottom();
	LONG lBottom;
	mp_oControl->clsG->mp_ClipRegion(mp_oControl->Getmt_LeftMargin(), mp_oControl->GetCurrentViewObject()->GetClientArea()->GetTop(), mp_oControl->GetSplitter()->GetLeft(), mp_oControl->Getmt_TableBottom(), FALSE);
	mp_oControl->GetDrawEventArgs()->Clear();
	mp_oControl->GetDrawEventArgs()->SetCustomDraw(FALSE);
	mp_oControl->GetDrawEventArgs()->SetEventTarget((LONG)EVT_GRID);
	mp_oControl->GetDrawEventArgs()->SetObjectIndex(0);
	mp_oControl->GetDrawEventArgs()->SetParentObjectIndex(0);
	mp_oControl->FireDraw();
	if (mp_oControl->GetDrawEventArgs()->GetCustomDraw() == TRUE) 
	{
		return;
	}
	if (GetCount() == 0) 
	{
		return;
	}
	for (lRowIndex = mp_lRealFirstVisibleRow; lRowIndex <= mp_oControl->GetCurrentViewObject()->GetClientArea()->GetLastVisibleRow(); lRowIndex++) 
	{
		oRow = (clsRow*)mp_oCollection->m_oReturnArrayElement(lRowIndex);
		if (oRow->GetVisible() == TRUE && oRow->GetHeight() > -1) 
		{
			if (oRow->GetMergeCells() == TRUE) 
			{
				if (oRow->GetBottom() > lTableBottom)
				{
					lBottom = lTableBottom;
				}
				else
				{
					lBottom = oRow->GetBottom();
				}
				mp_oControl->clsG->mp_ClipRegion(oRow->GetLeft(), oRow->GetTop(), oRow->GetRight(), lBottom, TRUE);
				mp_oControl->GetDrawEventArgs()->Clear();
				mp_oControl->GetDrawEventArgs()->SetCustomDraw(FALSE);
				mp_oControl->GetDrawEventArgs()->SetEventTarget(EVT_ROW);
				mp_oControl->GetDrawEventArgs()->SetObjectIndex(lRowIndex);
				mp_oControl->GetDrawEventArgs()->SetParentObjectIndex(0);
				mp_oControl->FireDraw();
				if (mp_oControl->GetDrawEventArgs()->GetCustomDraw() == FALSE) 
				{
					mp_oControl->clsG->mp_DrawItemRow(oRow->GetLeft(), oRow->GetTop(), oRow->GetRight(), oRow->GetBottom(), oRow->GetText(), (lRowIndex == mp_oControl->GetSelectedRowIndex()), oRow->GetImage(), 0, 0, oRow->GetStyle(), oRow);
					if (oRow->GetText().GetLength() > 0)
					{
						oRow->mp_lTextLeft = mp_oControl->clsG->mp_oTextFinalLayout.Left;
						oRow->mp_lTextTop = mp_oControl->clsG->mp_oTextFinalLayout.Top;
						oRow->mp_lTextRight = mp_oControl->clsG->mp_oTextFinalLayout.Left + mp_oControl->clsG->mp_oTextFinalLayout.Width - 1;
						oRow->mp_lTextBottom = mp_oControl->clsG->mp_oTextFinalLayout.Top + mp_oControl->clsG->mp_oTextFinalLayout.Height - 1;
					}
					else
					{
						oRow->mp_lTextLeft = oRow->GetLeft();
						oRow->mp_lTextTop = oRow->GetTop();
						oRow->mp_lTextRight = oRow->GetRight();
						oRow->mp_lTextBottom = oRow->GetBottom();
					}
				}
			}
			else 
			{
				for (lCellIndex = 1; lCellIndex <= mp_oControl->GetColumns()->GetCount(); lCellIndex++) 
				{
					if (lCellIndex != mp_oControl->GetTreeviewColumnIndex())
					{
						oCell = (clsCell*)oRow->GetCells()->mp_oCollection->m_oReturnArrayElement(lCellIndex);
						oColumn = (clsColumn*)mp_oControl->GetColumns()->mp_oCollection->m_oReturnArrayElement(lCellIndex);
						if (oColumn->GetVisible() == TRUE) 
						{
							if (oCell->GetBottom() > lTableBottom)
							{
								lBottom = lTableBottom;
							}
							else
							{
								lBottom = oCell->GetBottom();
							}
							mp_oControl->clsG->mp_ClipRegion(oCell->GetLeftTrim(), oCell->GetTop(), oCell->GetRightTrim(), lBottom, TRUE);
							mp_oControl->GetDrawEventArgs()->Clear();
							mp_oControl->GetDrawEventArgs()->SetCustomDraw(FALSE);
							mp_oControl->GetDrawEventArgs()->SetEventTarget((LONG)EVT_CELL);
							mp_oControl->GetDrawEventArgs()->SetObjectIndex(lCellIndex);
							mp_oControl->GetDrawEventArgs()->SetParentObjectIndex(lRowIndex);
							mp_oControl->FireDraw();
							if (mp_oControl->GetDrawEventArgs()->GetCustomDraw() == FALSE) 
							{
								mp_oControl->clsG->mp_DrawItemRow(oCell->GetLeft(), oCell->GetTop(), oCell->GetRight() - 1, oCell->GetBottom(), oCell->GetText(), (lRowIndex == mp_oControl->GetSelectedRowIndex() && lCellIndex == mp_oControl->GetSelectedCellIndex()), oCell->GetImage(), oCell->GetLeftTrim(), oCell->GetRightTrim(), oCell->GetStyle(), oRow);
								if (oCell->GetText().GetLength() > 0)
								{
									oCell->mp_lTextLeft = mp_oControl->clsG->mp_oTextFinalLayout.Left;
									oCell->mp_lTextTop = mp_oControl->clsG->mp_oTextFinalLayout.Top;
									oCell->mp_lTextRight = mp_oControl->clsG->mp_oTextFinalLayout.Left + mp_oControl->clsG->mp_oTextFinalLayout.Width - 1;
									oCell->mp_lTextBottom = mp_oControl->clsG->mp_oTextFinalLayout.Top + mp_oControl->clsG->mp_oTextFinalLayout.Height - 1;
								}
								else
								{
									oCell->mp_lTextLeft = oCell->GetLeft();
									oCell->mp_lTextTop = oCell->GetTop();
									oCell->mp_lTextRight = oCell->GetRight();
									oCell->mp_lTextBottom = oCell->GetBottom();
								}
							}
						}
					}
				}
			}
		}
	}
}

CString clsRows::GetXML(void)
{
	LONG lIndex;
	clsRow* oRow; 
	clsXML oXML(mp_oControl, "Rows");
	oXML.InitializeWriter();
	for (lIndex = 1; lIndex <= GetCount(); lIndex++)
	{
		oRow = (clsRow*) mp_oCollection->m_oReturnArrayElement(lIndex);
		oXML.WriteObject(oRow->GetXML());
	}
	return oXML.GetXML();
}

void clsRows::SetXML(CString sXML)
{
	LONG lIndex;
	clsXML oXML(mp_oControl, "Rows");
	oXML.SetXML(sXML);
	oXML.InitializeReader();
	mp_oCollection->m_Clear();
	for (lIndex = 1;lIndex <= oXML.ReadCollectionCount();lIndex++)
	{
		clsRow* oRow = new clsRow(mp_oControl);
		oRow->SetXML(oXML.ReadCollectionObject(lIndex));
		mp_oCollection->SetAddMode(TRUE);
		mp_oCollection->m_Add(oRow, oRow->mp_sKey, ROWS_ADD_1, ROWS_ADD_2, TRUE, ROWS_ADD_3);
	}
	mp_oControl->mp_oVerticalScrollBar->Update();
	mp_oControl->mp_oVerticalScrollBar->SetValue(1);
}

void clsRows::UpdateTree()
{
	if (VerifyTree() == FALSE)
	{
		mp_oControl->mp_ErrorReport(ROWS_INVALID_TREE_STRUCTURE, _T("Invalid treeview structure"), _T("clsRow.UpdateTree"));
		return;
	}
	LONG lIndex = 0;
	clsRow* oRow = NULL;
	mp_oTempNodeList.RemoveAll();
	for (lIndex = 1; lIndex <= GetCount(); lIndex++)
	{
		oRow = (clsRow*)mp_oCollection->m_oReturnArrayElement(lIndex);
		if (oRow->GetNode()->GetDepth() == 0)
		{
			oRow->GetNode()->mp_oParent = NULL;
		}
		else
		{
			oRow->GetNode()->mp_oParent = (clsNode*)mp_oTempNodeList[oRow->GetNode()->GetDepth() - 1];
		}
		if (oRow->GetNode()->GetDepth() > (mp_oTempNodeList.GetSize() - 1))
		{
			mp_oTempNodeList.Add(oRow->GetNode());
		}
		else
		{
			mp_oTempNodeList.SetAt(oRow->GetNode()->GetDepth(), oRow->GetNode());
		}
	}
}

BOOL clsRows::VerifyTree()
{
	LONG lPreviousDepth = 0;
	LONG lIndex = 0;
	clsRow* oRow = NULL;
	for (lIndex = 1; lIndex <= GetCount(); lIndex++)
	{
		oRow = (clsRow*)mp_oCollection->m_oReturnArrayElement(lIndex);
		LONG lDepth = oRow->GetNode()->GetDepth();
		if (lIndex == 1 && lDepth != 0)
		{
			return FALSE;
		}
		if ((lDepth - lPreviousDepth) > 1)
		{
			return FALSE;
		}
		lPreviousDepth = lDepth;
	}
	return TRUE;
}

void clsRows::BeginLoad(BOOL Preserve)
{
    mp_oTempNodeList.RemoveAll();
    if (Preserve == FALSE)
    {
        mp_lLoadIndex = 1;
		mp_oCollection->mp_aoCollection_Clear();
		mp_oCollection->mp_oKeys_Clear();
    }
    else
    {
        mp_lLoadIndex = mp_oCollection->mp_aoCollection.GetCount() + 1;
    }
}

clsRow* clsRows::Load(CString sKey)
{
	if (mp_lLoadIndex == -1)
	{
		return NULL;
	}
	int lIndex = 0;
	clsRow* oRow = new clsRow(mp_oControl);
	oRow->mp_sKey = sKey;
	for (lIndex = 1; lIndex <= mp_oControl->GetColumns()->GetCount(); lIndex++) 
	{
		oRow->GetCells()->Add();
	}
	oRow->mp_lIndex = mp_lLoadIndex;
	mp_oCollection->mp_aoCollection.SetAt(mp_lLoadIndex, oRow);
	int* pIndex;
	pIndex = new int(mp_lLoadIndex);
	mp_oCollection->mp_oKeys.SetAt(sKey, pIndex);
	mp_lLoadIndex = mp_lLoadIndex + 1;
	return oRow;
}

void clsRows::EndLoad(void)
{
    mp_oControl->GetVerticalScrollBar()->Update();
    UpdateTree();
	mp_lLoadIndex = -1;
}

void clsRows::NodesDrawElements()
{
	int lIndex = 0;
	clsNode* oNode = NULL;
	clsRow* oRow = NULL;
	if (GetCount() == 0)
	{
		return;
	}
	for (lIndex = mp_lRealFirstVisibleRow; lIndex <= GetCount(); lIndex++)
	{
		oRow = (clsRow*)mp_oCollection->m_oReturnArrayElement(lIndex);
		oNode = oRow->GetNode();
		if (oRow->GetClientAreaVisibility() == VS_INSIDEVISIBLEAREA && oNode->GetHidden() == FALSE)
		{
			if (oNode->Children() > 0)
			{
				mp_DrawSign(oNode->GetExpanded(), oNode->GetLeft(), oNode->GetYCenter(), oRow);
			}
			mp_DrawCheckBox(oNode);
			mp_DrawImage(oNode);
		}
	}
}