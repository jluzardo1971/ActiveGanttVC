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
#include "clsColumns.h"



IMPLEMENT_DYNCREATE(clsColumns, CCmdTarget)

//{5A5CD870-0C09-4FF9-AAF7-722D226E9E08}
static const IID IID_IclsColumns = { 0x5A5CD870, 0x0C09, 0x4FF9, { 0xAA, 0xF7, 0x72, 0x2D, 0x22, 0x6E, 0x9E, 0x08} };

//{4E43DBB3-EDE0-4CBC-848B-20069B58FC0C}
IMPLEMENT_OLECREATE_FLAGS(clsColumns, "AGVC.clsColumns", afxRegApartmentThreading, 0x4E43DBB3, 0xEDE0, 0x4CBC, 0x84, 0x8B, 0x20, 0x06, 0x9B, 0x58, 0xFC, 0x0C)

BEGIN_DISPATCH_MAP(clsColumns, CCmdTarget)
	DISP_PROPERTY_EX_ID(clsColumns, "Count", 1, odl_GetCount, SetNotSupported, VT_I4) 
	DISP_PROPERTY_PARAM_ID(clsColumns, "Item", 2, odl_Item, SetNotSupported, VT_DISPATCH, VTS_BSTR) 
	DISP_FUNCTION_ID(clsColumns, "Add", 3, odl_Add, VT_DISPATCH, VTS_BSTR VTS_BSTR VTS_I4 VTS_BSTR) 
	DISP_FUNCTION_ID(clsColumns, "Clear", 4, odl_Clear, VT_EMPTY, VTS_NONE) 
	DISP_FUNCTION_ID(clsColumns, "Remove", 5, odl_Remove, VT_EMPTY, VTS_BSTR) 
	DISP_FUNCTION_ID(clsColumns, "MoveColumn", 6, odl_MoveColumn, VT_EMPTY, VTS_I4 VTS_I4) 
	DISP_FUNCTION_ID(clsColumns, "GetXML", 7, odl_GetXML, VT_BSTR, VTS_NONE) 
	DISP_FUNCTION_ID(clsColumns, "SetXML", 8, odl_SetXML, VT_EMPTY, VTS_BSTR)
	DISP_DEFVALUE(clsColumns, "Item")
END_DISPATCH_MAP()

BEGIN_INTERFACE_MAP(clsColumns, CCmdTarget)
	INTERFACE_PART(clsColumns, IID_IclsColumns, Dispatch)
END_INTERFACE_MAP()

BEGIN_MESSAGE_MAP(clsColumns, CCmdTarget)
END_MESSAGE_MAP()

clsColumns::clsColumns(CActiveGanttVCCtl* oControl)
{
	EnableAutomation();
	AfxOleLockApp();
	mp_oControl = oControl;
	mp_oCollection = new clsCollectionBase(mp_oControl, "Column");
}

clsColumns::clsColumns()
{
	AfxMessageBox(_T("Invalid constructor has been invoked for clsColumns. Please inform support@sourcecodestore.com of this error"), MB_OK | MB_ICONSTOP | MB_APPLMODAL);
}

clsColumns::~clsColumns()
{
	delete mp_oCollection;
	mp_oCollection = NULL;
	AfxOleUnlockApp();
}

void clsColumns::OnFinalRelease()
{
	CCmdTarget::OnFinalRelease();
}

LONG clsColumns::GetCount(void)
{
	return mp_oCollection->m_lCount();
}

clsCollectionBase* clsColumns::GetCollection(void)
{
	return mp_oCollection;
}

LONG clsColumns::GetWidth(void)
{
	LONG lIndex;
	LONG lResult = 0;
	for (lIndex = 1;lIndex <= GetCount();lIndex++) 
	{
		clsColumn* oColumn;
		oColumn = (clsColumn*) mp_oCollection->m_oReturnArrayElement(lIndex);
		lResult = lResult + oColumn->GetWidth();
	}
	return lResult;
}

void clsColumns::Finalize(void)
{
}

clsColumn* clsColumns::Item(CString Index)
{
	clsColumn *oColumn;
	oColumn = (clsColumn*)mp_oCollection->m_oItem(Index, COLUMNS_ITEM_1, COLUMNS_ITEM_2, COLUMNS_ITEM_3, COLUMNS_ITEM_4);
	return oColumn;
}

clsColumn* clsColumns::Add(CString Text,CString Key,LONG Width,CString StyleIndex)
{
	mp_oCollection->SetAddMode(TRUE);
	clsColumn* oColumn = new clsColumn(mp_oControl);
	Text = g_Trim(Text);
	oColumn->SetText(Text);
	oColumn->SetWidth(Width);
	oColumn->SetStyleIndex(StyleIndex);
	oColumn->SetKey(Key);
	mp_oCollection->m_Add(oColumn, Key, COLUMNS_ADD_1, COLUMNS_ADD_2, FALSE, COLUMNS_ADD_3);
	LONG lIndex;
	clsRow* oRow;
	for (lIndex = 1;lIndex <= mp_oControl->GetRows()->GetCount();lIndex++) 
	{
		oRow = (clsRow*) mp_oControl->GetRows()->mp_oCollection->m_oReturnArrayElement(lIndex);
		oRow->GetCells()->Add();
	}
	mp_oControl->GetSplitter()->f_AdjustPosition();
	return oColumn;
}

void clsColumns::Clear(void)
{
	mp_oControl->GetRows()->ClearCells();
	mp_oCollection->m_Clear();
	mp_oControl->SetSelectedColumnIndex(0);
	mp_oControl->SetSelectedCellIndex(0);
	mp_oControl->GetSplitter()->f_AdjustPosition();
}

void clsColumns::Remove(CString Index)
{
	LONG lIndex;
	clsRow* oRow;
	for (lIndex = 1;lIndex <= mp_oControl->GetRows()->GetCount();lIndex++) 
	{
		oRow = (clsRow*) mp_oControl->GetRows()->mp_oCollection->m_oReturnArrayElement(lIndex);
		oRow->GetCells()->Remove(Index);
	}
	mp_oCollection->m_Remove(Index, COLUMNS_REMOVE_1, COLUMNS_REMOVE_2, COLUMNS_REMOVE_3, COLUMNS_REMOVE_4);
	mp_oControl->SetSelectedColumnIndex(0);
	mp_oControl->SetSelectedCellIndex(0);
	mp_oControl->GetSplitter()->f_AdjustPosition();
}

void clsColumns::MoveColumn(LONG OriginColumnIndex,LONG DestColumnIndex)
{
	clsColumn* oColumn;
	clsRow* oRow;
	int lIndex = 0;
	if (OriginColumnIndex < 1 || OriginColumnIndex > GetCount())
	{
		return;
	}
	if (DestColumnIndex < 1 || DestColumnIndex > GetCount())
	{
		return;
	}
	if (DestColumnIndex == OriginColumnIndex)
	{
		return;
	}
	if (mp_oControl->GetTreeviewColumnIndex() > 0)
	{
		for (lIndex = 1; lIndex <= mp_oControl->GetColumns()->GetCount(); lIndex++)
		{
			oColumn = (clsColumn*)mp_oControl->GetColumns()->mp_oCollection->m_oReturnArrayElement(lIndex);
			if (lIndex == mp_oControl->GetTreeviewColumnIndex())
			{
				oColumn->mp_bTreeViewColumnIndex = TRUE;
			}
			else
			{
				oColumn->mp_bTreeViewColumnIndex = FALSE;
			}
		}
	}
	mp_oCollection->m_lCopyAndMoveItems(OriginColumnIndex, DestColumnIndex);
	for (lIndex = 1; lIndex <= mp_oControl->GetRows()->GetCount(); lIndex++)
	{
		oRow = (clsRow*)mp_oControl->GetRows()->mp_oCollection->m_oReturnArrayElement(lIndex);
		oRow->GetCells()->mp_oCollection->m_lCopyAndMoveItems(OriginColumnIndex, DestColumnIndex);
	}
	if (mp_oControl->GetTreeviewColumnIndex() > 0)
	{
		for (lIndex = 1; lIndex <= mp_oControl->GetColumns()->GetCount(); lIndex++)
		{
			oColumn = (clsColumn*)mp_oControl->GetColumns()->mp_oCollection->m_oReturnArrayElement(lIndex);
			if (oColumn->mp_bTreeViewColumnIndex == TRUE)
			{
				mp_oControl->SetTreeviewColumnIndex(lIndex);
			}
		}
	}
}

void clsColumns::Position(void)
{
	LONG lIndex;
	clsColumn* oColumn;
	LONG lLeft;
	lLeft = -mp_oControl->GetHorizontalScrollBar()->GetValue() + mp_oControl->Getmt_LeftMargin();
	for (lIndex = 1;lIndex <= GetCount();lIndex++) 
	{
		oColumn = (clsColumn*) mp_oCollection->m_oReturnArrayElement(lIndex);
		oColumn->Setf_lLeft(lLeft);
		oColumn->Setf_lRight(lLeft + oColumn->GetWidth());
		if (oColumn->GetRight() < mp_oControl->Getmt_LeftMargin())
		{
			oColumn->Setf_bVisible(FALSE);
		}
		else if (oColumn->GetLeft() > mp_oControl->GetSplitter()->GetLeft())
		{
			oColumn->Setf_bVisible(FALSE);
		}
		else
		{
			oColumn->Setf_bVisible(TRUE);
		}
		if (oColumn->GetRight() > oColumn->GetLeft())
		{
			oColumn->Setf_bVisible(TRUE);
		}
		else
		{
			oColumn->Setf_bVisible(FALSE);
		}
		lLeft = lLeft + oColumn->GetWidth();
	}
}

void clsColumns::Draw(void)
{
	clsColumn* oColumn;
	LONG lIndex;
	if (GetCount() == 0)
	{
		return;
	}
	if ((mp_oControl->GetCurrentViewObject()->GetTimeLine()->GetHeight() == 0))
	{
		return;
	}
	for (lIndex = 1;lIndex <= GetCount();lIndex++) 
	{
		oColumn = (clsColumn*) mp_oCollection->m_oReturnArrayElement(lIndex);
		if (oColumn->GetVisible() == TRUE)
		{
			mp_oControl->clsG->mp_ClipRegion(oColumn->GetLeftTrim(), oColumn->GetTop(), oColumn->GetRightTrim(), oColumn->GetBottom(), TRUE);
			mp_oControl->GetDrawEventArgs()->Clear();
			mp_oControl->GetDrawEventArgs()->SetCustomDraw(FALSE);
			mp_oControl->GetDrawEventArgs()->SetEventTarget((LONG)EVT_COLUMN);
			mp_oControl->GetDrawEventArgs()->SetObjectIndex(lIndex);
			mp_oControl->GetDrawEventArgs()->SetParentObjectIndex(0);
			mp_oControl->FireDraw();
			if (mp_oControl->GetDrawEventArgs()->GetCustomDraw() == FALSE)
			{
				mp_oControl->clsG->mp_DrawItem(oColumn->GetLeft(), oColumn->GetTop(), oColumn->GetRight() - 1, oColumn->GetBottom(), _T(""), oColumn->GetText(), (lIndex == mp_oControl->GetSelectedColumnIndex()), oColumn->GetImage(), oColumn->GetLeftTrim(), oColumn->GetRightTrim(), oColumn->GetStyle());
				if (oColumn->GetText().GetLength() > 0)
				{
					oColumn->mp_lTextLeft = mp_oControl->clsG->mp_oTextFinalLayout.Left;
					oColumn->mp_lTextTop = mp_oControl->clsG->mp_oTextFinalLayout.Top;
					oColumn->mp_lTextRight = mp_oControl->clsG->mp_oTextFinalLayout.Left + mp_oControl->clsG->mp_oTextFinalLayout.Width - 1;
					oColumn->mp_lTextBottom = mp_oControl->clsG->mp_oTextFinalLayout.Top + mp_oControl->clsG->mp_oTextFinalLayout.Height - 1;
				}
				else
				{
					oColumn->mp_lTextLeft = oColumn->GetLeft();
					oColumn->mp_lTextTop = oColumn->GetTop();
					oColumn->mp_lTextRight = oColumn->GetRight();
					oColumn->mp_lTextBottom = oColumn->GetBottom();
				}
			}
		}
	}
}

CString clsColumns::GetXML(void)
{
	LONG lIndex;
	clsColumn* oColumn; 
	clsXML oXML(mp_oControl, "Columns");
	oXML.InitializeWriter();
	for (lIndex = 1; lIndex <= GetCount(); lIndex++)
	{
		oColumn = (clsColumn*) mp_oCollection->m_oReturnArrayElement(lIndex);
		oXML.WriteObject(oColumn->GetXML());
	}
	return oXML.GetXML();
}

void clsColumns::SetXML(CString sXML)
{
	LONG lIndex;
	clsXML oXML(mp_oControl, "Columns");
	oXML.SetXML(sXML);
	oXML.InitializeReader();
	mp_oCollection->m_Clear();
	for (lIndex = 1;lIndex <= oXML.ReadCollectionCount();lIndex++)
	{
		clsColumn* oColumn = new clsColumn(mp_oControl);
		oColumn->SetXML(oXML.ReadCollectionObject(lIndex));
		mp_oCollection->SetAddMode(TRUE);
		mp_oCollection->m_Add(oColumn, oColumn->mp_sKey, COLUMNS_ADD_1, COLUMNS_ADD_2, FALSE, COLUMNS_ADD_3);
	}
}