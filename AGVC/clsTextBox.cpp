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
#include "clsTextBox.h"



clsTextBox::clsTextBox(CActiveGanttVCCtl* oControl)
{
	mp_oControl = oControl;
	mp_bInitialized = FALSE;
}

clsTextBox::~clsTextBox(void)
{
}

void clsTextBox::SetGDIFont(clsFont* oFont)
{
	LOGFONT lf;
#ifdef _UNICODE
	
	oFont->GetFontP()->GetLogFontW(mp_oControl->clsG->oGraphics(), &lf);
	mp_oFont.CreateFontIndirectW(&lf);
#else
	oFont->GetFontP()->GetLogFontA(mp_oControl->clsG->oGraphics(), &lf);
	mp_oFont.CreateFontIndirectA(&lf);
#endif

	this->SetFont(&mp_oFont);

}

void clsTextBox::Initialize(int lIndex, int lIndex2, E_TEXTOBJECTTYPE yObjectType, int X, int Y)
{
	mp_yObjectType = yObjectType;
	mp_lIndex = lIndex;
	mp_lIndex2 = lIndex2;

	if (!::IsWindow(m_hWnd))
	{
		CRect oRect(0,0,10,10);
		BOOL bCreated = FALSE;
		bCreated = Create(WS_CHILD | WS_VISIBLE | ES_MULTILINE | ES_AUTOHSCROLL | ES_AUTOVSCROLL | ES_LEFT, oRect, mp_oControl, 15620);
	}

	if (mp_bInitialized == TRUE)
	{
		Terminate();
	}
	mp_oControl->MouseKeyboardEvents->mp_yOperation = EO_NONE;
	mp_oControl->MouseKeyboardEvents->mp_SetCursor(CT_NORMAL);

	switch (mp_yObjectType)
	{
	case TOT_COLUMN:
		mp_oColumn = (clsColumn*)mp_oControl->GetColumns()->mp_oCollection->m_oReturnArrayElement(lIndex);
		mp_clrBackColor = mp_oColumn->GetStyle()->GetTextEditBackColor();
		mp_clrForeColor = mp_oColumn->GetStyle()->GetTextEditForeColor();
		SetGDIFont(mp_oColumn->GetStyle()->GetFont());
		this->MoveWindow(mp_oColumn->mp_lTextLeft, mp_oColumn->mp_lTextTop, mp_oColumn->mp_lTextRight - mp_oColumn->mp_lTextLeft + 1, mp_oColumn->mp_lTextBottom - mp_oColumn->mp_lTextTop + 1);
		this->SetWindowTextW(mp_oColumn->GetText());
		break;
	case TOT_NODE:
		mp_oRow = (clsRow*)mp_oControl->GetRows()->mp_oCollection->m_oReturnArrayElement(lIndex);
		mp_oNode = mp_oRow->GetNode();
		mp_clrBackColor = mp_oNode->GetStyle()->GetTextEditBackColor();
		mp_clrForeColor = mp_oNode->GetStyle()->GetTextEditForeColor();
		SetGDIFont(mp_oNode->GetStyle()->GetFont());
		this->MoveWindow(mp_oNode->mp_lTextLeft, mp_oNode->mp_lTextTop, mp_oNode->mp_lTextRight - mp_oNode->mp_lTextLeft + 1, mp_oNode->mp_lTextBottom - mp_oNode->mp_lTextTop + 1);
		this->SetWindowTextW(mp_oNode->GetText());
		break;
	case TOT_ROW:
		mp_oRow = (clsRow*)mp_oControl->GetRows()->mp_oCollection->m_oReturnArrayElement(lIndex);
		mp_clrBackColor = mp_oRow->GetStyle()->GetTextEditBackColor();
		mp_clrForeColor = mp_oRow->GetStyle()->GetTextEditForeColor();
		SetGDIFont(mp_oRow->GetStyle()->GetFont());
		this->MoveWindow(mp_oRow->mp_lTextLeft, mp_oRow->mp_lTextTop, mp_oRow->mp_lTextRight - mp_oRow->mp_lTextLeft + 1, mp_oRow->mp_lTextBottom - mp_oRow->mp_lTextTop + 1);
		this->SetWindowTextW(mp_oRow->GetText());
		break;
	case TOT_CELL:
		mp_oRow = (clsRow*)mp_oControl->GetRows()->mp_oCollection->m_oReturnArrayElement(lIndex);
		mp_oCell = (clsCell*)mp_oRow->GetCells()->mp_oCollection->m_oReturnArrayElement(lIndex2);
		mp_clrBackColor = mp_oCell->GetStyle()->GetTextEditBackColor();
		mp_clrForeColor = mp_oCell->GetStyle()->GetTextEditForeColor();
		SetGDIFont(mp_oCell->GetStyle()->GetFont());
		this->MoveWindow(mp_oCell->mp_lTextLeft, mp_oCell->mp_lTextTop, mp_oCell->mp_lTextRight - mp_oCell->mp_lTextLeft + 1, mp_oCell->mp_lTextBottom - mp_oCell->mp_lTextTop + 1);
		this->SetWindowTextW(mp_oCell->GetText());
		break;
	case TOT_TASK:
		mp_oTask = (clsTask*)mp_oControl->GetTasks()->mp_oCollection->m_oReturnArrayElement(lIndex);
		mp_clrBackColor = mp_oTask->GetStyle()->GetTextEditBackColor();
		mp_clrForeColor = mp_oTask->GetStyle()->GetTextEditForeColor();
		SetGDIFont(mp_oTask->GetStyle()->GetFont());
		this->MoveWindow(mp_oTask->mp_lTextLeft, mp_oTask->mp_lTextTop, mp_oTask->mp_lTextRight - mp_oTask->mp_lTextLeft + 1, mp_oTask->mp_lTextBottom - mp_oTask->mp_lTextTop + 1);
		this->SetWindowTextW(mp_oTask->GetText());
		break;
	}
	
	CString sBuff;
	this->GetWindowTextW(sBuff);
	mp_sText = sBuff;
	mp_oControl->GetTextEditEventArgs()->Clear();
	mp_oControl->GetTextEditEventArgs()->SetObjectType(mp_yObjectType);
	if (mp_yObjectType == TOT_CELL)
	{
		mp_oControl->GetTextEditEventArgs()->SetParentObjectIndex(mp_lIndex);
		mp_oControl->GetTextEditEventArgs()->SetObjectIndex(mp_lIndex2);
	}
	else
	{
		mp_oControl->GetTextEditEventArgs()->SetParentObjectIndex(0);
		mp_oControl->GetTextEditEventArgs()->SetObjectIndex(mp_lIndex);
	}
	mp_oControl->GetTextEditEventArgs()->SetText(sBuff);
	mp_oControl->FireBeginTextEdit();
	if (mp_oControl->GetTextEditEventArgs()->GetText() != sBuff)
	{
	    this->SetWindowTextW(mp_oControl->GetTextEditEventArgs()->GetText());
	}
	mp_bInitialized = TRUE;
	this->ShowWindow(TRUE);
	POINT point;
	GetCursorPos(&point);
	PostMessage(WM_LBUTTONDOWN, MK_LBUTTON, MAKELONG(point.x, point.y));
	PostMessage(WM_LBUTTONUP, MK_LBUTTON, MAKELONG(point.x, point.y));
	//if (mp_sText.GetLength() > 0)
	//{
	//	point.x = X - mp_oHeader->mp_lTextLeft;
	//	point.y = Y - mp_oHeader->mp_lTextTop;
	//	int n = this->CharFromPos(point);
	//	int nLineIndex = HIWORD(n);
	//	int nCharIndex = LOWORD(n);
	//	this->SetSel(0,-1);
	//}
}

void clsTextBox::Terminate()
{
	if (mp_bInitialized == TRUE)
	{
		CString sBuff;
		this->GetWindowTextW(sBuff);
		mp_sText = sBuff;
		mp_oControl->GetTextEditEventArgs()->Clear();
		mp_oControl->GetTextEditEventArgs()->SetObjectType(mp_yObjectType);
	    if (mp_yObjectType == TOT_CELL)
	    {
			mp_oControl->GetTextEditEventArgs()->SetParentObjectIndex(mp_lIndex);
			mp_oControl->GetTextEditEventArgs()->SetObjectIndex(mp_lIndex2);
	    }
	    else
	    {
			mp_oControl->GetTextEditEventArgs()->SetParentObjectIndex(0);
			mp_oControl->GetTextEditEventArgs()->SetObjectIndex(mp_lIndex);
	    }
	    mp_oControl->GetTextEditEventArgs()->SetText(sBuff);
	    mp_oControl->FireEndTextEdit();
	}

	mp_bInitialized = FALSE;
	if (::IsWindow(m_hWnd))
	{
		mp_oFont.Detach();
		this->ShowWindow(FALSE);
	}
}

BEGIN_MESSAGE_MAP(clsTextBox, CEdit)
END_MESSAGE_MAP()

void clsTextBox::OnUpdateEdit()
{
	CString sBuff;
	this->GetWindowTextW(sBuff);
	switch (mp_yObjectType)
	{
	case TOT_COLUMN:
		mp_oColumn->SetText(sBuff);
		mp_oControl->Refresh();
		this->MoveWindow(mp_oColumn->mp_lTextLeft, mp_oColumn->mp_lTextTop, mp_oColumn->mp_lTextRight - mp_oColumn->mp_lTextLeft + 1, mp_oColumn->mp_lTextBottom - mp_oColumn->mp_lTextTop + 1);
		break;
	case TOT_NODE:
		mp_oNode->SetText(sBuff);
		mp_oControl->Refresh();
		this->MoveWindow(mp_oNode->mp_lTextLeft, mp_oNode->mp_lTextTop, mp_oNode->mp_lTextRight - mp_oNode->mp_lTextLeft + 1, mp_oNode->mp_lTextBottom - mp_oNode->mp_lTextTop + 1, TRUE);
		break;
	case TOT_ROW:
		mp_oRow->SetText(sBuff);
		mp_oControl->Refresh();
		this->MoveWindow(mp_oRow->mp_lTextLeft, mp_oRow->mp_lTextTop, mp_oRow->mp_lTextRight - mp_oRow->mp_lTextLeft + 1, mp_oRow->mp_lTextBottom - mp_oRow->mp_lTextTop + 1, TRUE);
		break;
	case TOT_CELL:
		mp_oCell->SetText(sBuff);
		mp_oControl->Refresh();
		this->MoveWindow(mp_oCell->mp_lTextLeft, mp_oCell->mp_lTextTop, mp_oCell->mp_lTextRight - mp_oCell->mp_lTextLeft + 1, mp_oCell->mp_lTextBottom - mp_oCell->mp_lTextTop + 1, TRUE);
		break;
	case TOT_TASK:
		mp_oTask->SetText(sBuff);
		mp_oControl->Refresh();
		this->MoveWindow(mp_oTask->mp_lTextLeft, mp_oTask->mp_lTextTop, mp_oTask->mp_lTextRight - mp_oTask->mp_lTextLeft + 1, mp_oTask->mp_lTextBottom - mp_oTask->mp_lTextTop + 1, TRUE);
		break;
	}
}