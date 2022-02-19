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
#include "activeganttvc.h"
#include "clsMouseKeyboardEvents.h"
#include "ActiveGanttVCCtl.h"
#include "GlobalFunctions.h"

#ifdef _DEBUG
#undef THIS_FILE
static char THIS_FILE[]=__FILE__;
#define new DEBUG_NEW
#endif



clsMouseKeyboardEvents::clsMouseKeyboardEvents(CActiveGanttVCCtl* oControl)
{
	mp_oControl = oControl;
	mp_oToolTip = new clsToolTip(mp_oControl);
	mp_hDefault = AfxGetApp()->LoadCursor(IDCS_DEFAULT);
	mp_hSizeWE = AfxGetApp()->LoadCursor(IDCS_SIZEWE);
	mp_hSizeAll = AfxGetApp()->LoadCursor(IDCS_SIZEALL);
	mp_hCrossHair = AfxGetApp()->LoadCursor(IDCS_CROSSHAIR);
	mp_hHorizontalSplit = AfxGetApp()->LoadCursor(IDCS_HORIZONTALSPLIT);
	mp_hVerticalSplit = AfxGetApp()->LoadCursor(IDCS_VERTICALSPLIT);
	mp_hDragMove = AfxGetApp()->LoadCursor(IDCS_DRAGMOVE);
	mp_hWait = AfxGetApp()->LoadCursor(IDCS_WAIT);
	mp_hPercentage = AfxGetApp()->LoadCursor(IDCS_PERCENTAGE);
	mp_hPredecessor = AfxGetApp()->LoadCursor(IDCS_PREDECESSOR);
	mp_hNoDrop = AfxGetApp()->LoadCursor(IDCS_NODROP);
	mp_hIBeam = AfxGetApp()->LoadCursor(IDCS_IBEAM);
}

clsMouseKeyboardEvents::~clsMouseKeyboardEvents()
{
	delete mp_oToolTip;
	mp_oToolTip = NULL;
}

void clsMouseKeyboardEvents::KeyDown(void)
{
	mp_oControl->FireControlKeyDown();
	if (mp_oControl->GetKeyEventArgs()->GetCancel() == TRUE) 
	{
		return;
	}
	switch (mp_oControl->GetKeyEventArgs()->GetKeyCode()) 
	{
	case KEYS_P:
		mp_yOperationModifier = OPM_PREDECESSORMODE;
		break;
	}
}

void clsMouseKeyboardEvents::KeyUp(void)
{
	mp_oControl->FireControlKeyUp();
	mp_yOperationModifier = OPM_NONE;
}

void clsMouseKeyboardEvents::KeyPress(void)
{
	mp_oControl->FireControlKeyPress();
	if (mp_oControl->GetKeyEventArgs()->GetCancel() == TRUE) 
	{
		return;
	}
}

void clsMouseKeyboardEvents::OnMouseClick(void)
{
	mp_oControl->FireControlClick();
	if (mp_oControl->GetMouseEventArgs()->GetCancel() == TRUE) 
	{
		return;
	}
}

void clsMouseKeyboardEvents::OnMouseDblClick(void)
{
	mp_oControl->FireControlDblClick();
	if (mp_oControl->GetMouseEventArgs()->GetCancel() == TRUE) 
	{
		return;
	}
}

void clsMouseKeyboardEvents::OnMouseMoveGeneral(LONG X,LONG Y)
{
	if (mp_yOperation == EO_NONE) 
	{
		OnMouseHover(X, Y);
	}
	else 
	{
		OnMouseMove(X, Y);
	}
}

void clsMouseKeyboardEvents::OnMouseHover(LONG X,LONG Y)
{

	E_EVENTTARGET yEventTarget = EVT_NONE;
	yEventTarget = (E_EVENTTARGET) CursorPosition(X, Y);
	mp_oControl->GetMouseEventArgs()->Clear();
	mp_oControl->GetMouseEventArgs()->SetX(X);
	mp_oControl->GetMouseEventArgs()->SetY(Y);
	mp_oControl->GetMouseEventArgs()->SetEventTarget(yEventTarget);
	mp_oControl->GetMouseEventArgs()->SetCancel(FALSE);
	mp_oControl->FireControlMouseHover();
	if (mp_oControl->GetMouseEventArgs()->GetCancel() == TRUE) 
	{
		return;
	}
	mp_oControl->GetToolTipEventArgs()->Clear();
	mp_oControl->GetToolTipEventArgs()->SetX(X);
	mp_oControl->GetToolTipEventArgs()->SetY(Y);
	mp_oControl->GetToolTipEventArgs()->SetEventTarget(yEventTarget);
	mp_oControl->FireToolTipOnMouseHover(yEventTarget);
	switch (yEventTarget) 
	{
	case EVT_VSCROLLBAR:
		mp_oControl->GetVerticalScrollBar()->GetScrollBar()->OnMouseHover(X, Y);
		break;
	case EVT_HSCROLLBAR:
		mp_oControl->GetHorizontalScrollBar()->GetScrollBar()->OnMouseHover(X, Y);
		break;
	case EVT_TIMELINESCROLLBAR:
		mp_oControl->f_TimeLineScrollBar()->GetScrollBar()->OnMouseHover(X, Y);
		break;
	case EVT_TREEVIEW:
		mp_oControl->GetTreeview()->OnMouseHover(X, Y);
		break;
	case EVT_SPLITTER:
		mp_EO_SPLITTERMOVEMENT(MouseHover, X, Y);
		break;
	case EVT_TIMELINE:
		mp_EO_TIMELINEMOVEMENT(MouseHover, X, Y);
		break;
	case EVT_SELECTEDCOLUMN:
		if (mp_bCursorEditTextColumn(X, Y) == TRUE)
		{
			return;
		}
		switch (mp_yColumnArea(X, Y)) 
		{
		case EA_RIGHT:
			mp_EO_COLUMNSIZING(MouseHover, X, Y);
			break;
		case EA_CENTER:
			mp_EO_COLUMNMOVEMENT(MouseHover, X, Y);
			break;
		}
		break;
	case EVT_COLUMN:
		mp_EO_COLUMNSELECTION(MouseHover, X, Y);
		break;
	case EVT_SELECTEDROW:
		if (mp_bCursorEditTextRow(X, Y) == TRUE)
		{
			return;
		}
		switch (mp_yRowArea(X, Y)) 
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
	case EVT_SELECTEDPERCENTAGE:
		mp_EO_PERCENTAGESIZING(MouseHover, X, Y);
		break;
	case EVT_PERCENTAGE:
		mp_EO_PERCENTAGESELECTION(MouseHover, X, Y);
		break;
	case EVT_SELECTEDTASK:
		if (mp_yOperationModifier == OPM_PREDECESSORMODE) 
		{
			mp_EO_PREDECESSORADDITION(MouseHover, X, Y);
		}
		else 
		{
			if (mp_bCursorEditTextTask(X, Y) == TRUE)
			{
				return;
			}
			switch (mp_yTaskArea(X, Y)) 
			{
			case EA_CENTER:
				mp_EO_TASKMOVEMENT(MouseHover, X, Y);
				break;
			case EA_LEFT:
				mp_EO_TASKSTRETCHLEFT(MouseHover, X, Y);
				break;
			case EA_RIGHT:
				mp_EO_TASKSTRETCHRIGHT(MouseHover, X, Y);
				break;
			case EA_NONE:
				break;
			}
		}

		break;
	case EVT_TASK:
		mp_EO_TASKSELECTION(MouseHover, X, Y);
		break;
    case EVT_SELECTEDPREDECESSOR:
        //
        break;
    case EVT_PREDECESSOR:
        mp_EO_PREDECESSORSELECTION(MouseHover, X, Y);
        break;
	case EVT_CLIENTAREA:
		if (mp_bCursorEditTextTask(X, Y) == TRUE)
		{
			return;
		}
		mp_EO_TASKADDITION(MouseHover, X, Y);
		break;
	case EVT_NONE:
		mp_SetCursor(CT_NORMAL);
		break;
		//Moving over empty space
	}
}

void clsMouseKeyboardEvents::OnMouseDown(LONG X,LONG Y,LONG Button)
{
	mp_oControl->mp_oTextBox->Terminate();
	E_EVENTTARGET EventTarget = EVT_NONE;
	EventTarget = (E_EVENTTARGET)CursorPosition(X, Y);
	mp_oControl->GetMouseEventArgs()->Clear();
	mp_oControl->GetMouseEventArgs()->SetX(X);
	mp_oControl->GetMouseEventArgs()->SetY(Y);
	mp_oControl->GetMouseEventArgs()->SetEventTarget(EventTarget);
	mp_oControl->GetMouseEventArgs()->SetButton((E_MOUSEBUTTONS)Button);
	mp_oControl->GetMouseEventArgs()->SetCancel(FALSE);
	mp_oControl->FireControlMouseDown();
	if (mp_oControl->GetMouseEventArgs()->GetCancel() == TRUE) 
	{
		return;
	}
	switch (CursorPosition(X, Y)) 
	{
	case EVT_VSCROLLBAR:
		mp_oControl->GetVerticalScrollBar()->GetScrollBar()->OnMouseDown(X, Y);
		mp_yOperation = EO_VERTICALSCROLLBAR;
		break;
	case EVT_HSCROLLBAR:
		mp_oControl->GetHorizontalScrollBar()->GetScrollBar()->OnMouseDown(X, Y);
		mp_yOperation = EO_HORIZONTALSCROLLBAR;
		break;
	case EVT_TIMELINESCROLLBAR:
		mp_oControl->f_TimeLineScrollBar()->GetScrollBar()->OnMouseDown(X, Y);
		mp_yOperation = EO_TIMELINESCROLLBAR;
		break;
	case EVT_TREEVIEW:
		mp_oControl->GetTreeview()->OnMouseDown(X, Y);
		mp_yOperation = EO_TREEVIEW;
		break;
	case EVT_SPLITTER:
		mp_EO_SPLITTERMOVEMENT(MouseDown, X, Y);
		mp_yOperation = EO_SPLITTERMOVEMENT;
		break;
	case EVT_TIMELINE:
		mp_EO_TIMELINEMOVEMENT(MouseDown, X, Y);
		mp_yOperation = EO_TIMELINEMOVEMENT;
		break;
	case EVT_SELECTEDCOLUMN:
		if (mp_bShowEditTextColumn(X, Y) == TRUE)
		{
			return;
		}
		switch (mp_yColumnArea(X, Y)) 
		{
		case EA_RIGHT:
			mp_EO_COLUMNSIZING(MouseDown, X, Y);
			mp_yOperation = EO_COLUMNSIZING;
			break;
		case EA_CENTER:
			mp_EO_COLUMNMOVEMENT(MouseDown, X, Y);
			mp_yOperation = EO_COLUMNMOVEMENT;
			break;
		}
		break;
	case EVT_COLUMN:
		mp_EO_COLUMNSELECTION(MouseDown, X, Y);
		mp_yOperation = EO_COLUMNSELECTION;
		break;
	case EVT_SELECTEDROW:
		if (mp_bShowEditTextRow(X, Y) == TRUE)
		{
			return;
		}
		switch (mp_yRowArea(X, Y)) 
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
	case EVT_SELECTEDPERCENTAGE:
		mp_EO_PERCENTAGESIZING(MouseDown, X, Y);
		mp_yOperation = EO_PERCENTAGESIZING;
		break;
	case EVT_PERCENTAGE:
		mp_EO_PERCENTAGESELECTION(MouseDown, X, Y);
		mp_yOperation = EO_PERCENTAGESELECTION;
		break;
	case EVT_SELECTEDTASK:
		if (mp_yOperationModifier == OPM_PREDECESSORMODE) 
		{
			mp_EO_PREDECESSORADDITION(MouseDown, X, Y);
			mp_yOperation = EO_PREDECESSORADDITION;
		}
		else 
		{
			if (mp_bShowEditTextTask(X, Y) == TRUE)
			{
				return;
			}
			switch (mp_yTaskArea(X, Y)) 
			{
			case EA_CENTER:
				mp_EO_TASKMOVEMENT(MouseDown, X, Y);
				mp_yOperation = EO_TASKMOVEMENT;
				break;
			case EA_LEFT:
				mp_EO_TASKSTRETCHLEFT(MouseDown, X, Y);
				mp_yOperation = EO_TASKSTRETCHLEFT;
				break;
			case EA_RIGHT:
				mp_EO_TASKSTRETCHRIGHT(MouseDown, X, Y);
				mp_yOperation = EO_TASKSTRETCHRIGHT;
				break;
			case EA_NONE:
				mp_yOperation = EO_NONE;
				break;
			}
		}

		break;
	case EVT_TASK:
		mp_EO_TASKSELECTION(MouseDown, X, Y);
		mp_yOperation = EO_TASKSELECTION;
		break;
    case EVT_SELECTEDPREDECESSOR:
        //
        break;
    case EVT_PREDECESSOR:
        mp_EO_PREDECESSORSELECTION(MouseDown, X, Y);
        mp_yOperation = EO_PREDECESSORSELECTION;
        break;
	case EVT_CLIENTAREA:
		if (mp_bShowEditTextTask(X, Y) == TRUE)
		{
			return;
		}
        mp_EO_TASKADDITION(MouseDown, X, Y);
        switch (mp_oControl->GetAddMode())
        {
            case AT_TASKADD:
            case AT_DURATION_TASKADD:
                mp_yOperation = EO_TASKADDITION;
                break;
            case AT_MILESTONEADD:
            case AT_DURATION_MILESTONEADD:
                mp_yOperation = EO_MILESTONEADDITION;
                break;
            case AT_BOTH:
            case AT_DURATION_BOTH:
                mp_yOperation = EO_TASKADDITION;
                break;
        }
        break;
	}
}

void clsMouseKeyboardEvents::OnMouseMove(LONG X,LONG Y)
{
	E_OPERATION yOperation = (E_OPERATION) mp_yOperation;
	mp_oControl->GetMouseEventArgs()->SetX(X);
	mp_oControl->GetMouseEventArgs()->SetY(Y);
	mp_oControl->GetMouseEventArgs()->SetOperation((LONG)mp_yOperation);
	mp_oControl->FireControlMouseMove();
	if (mp_oControl->GetMouseEventArgs()->GetCancel() == TRUE) 
	{
		mp_yOperation = EO_NONE;
		return;
	}
	switch (mp_yOperation) 
	{
	case EO_VERTICALSCROLLBAR:
		mp_oControl->GetVerticalScrollBar()->GetScrollBar()->OnMouseMove(X, Y);
		break;
	case EO_HORIZONTALSCROLLBAR:
		mp_oControl->GetHorizontalScrollBar()->GetScrollBar()->OnMouseMove(X, Y);
		break;
	case EO_TIMELINESCROLLBAR:
		mp_oControl->f_TimeLineScrollBar()->GetScrollBar()->OnMouseMove(X, Y);
		break;
	case EO_TREEVIEW:
		mp_oControl->GetTreeview()->OnMouseMove(X, Y);
		break;
	case EO_SPLITTERMOVEMENT:
		mp_EO_SPLITTERMOVEMENT(MouseMove, X, Y);
		break;
	case EO_TIMELINEMOVEMENT:
		mp_EO_TIMELINEMOVEMENT(MouseMove, X, Y);
		break;
	case EO_COLUMNMOVEMENT:
		mp_EO_COLUMNMOVEMENT(MouseMove, X, Y);
		break;
	case EO_COLUMNSIZING:
		mp_EO_COLUMNSIZING(MouseMove, X, Y);
		break;
	case EO_COLUMNSELECTION:
		mp_EO_COLUMNSELECTION(MouseMove, X, Y);
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
	case EO_PERCENTAGESIZING:
		mp_EO_PERCENTAGESIZING(MouseMove, X, Y);
		break;
	case EO_PERCENTAGESELECTION:
		mp_EO_PERCENTAGESELECTION(MouseMove, X, Y);
		break;
	case EO_PREDECESSORADDITION:
		mp_EO_PREDECESSORADDITION(MouseMove, X, Y);
		break;
	case EO_TASKMOVEMENT:
		mp_EO_TASKMOVEMENT(MouseMove, X, Y);
		break;
	case EO_TASKSTRETCHLEFT:
		mp_EO_TASKSTRETCHLEFT(MouseMove, X, Y);
		break;
	case EO_TASKSTRETCHRIGHT:
		mp_EO_TASKSTRETCHRIGHT(MouseMove, X, Y);
		break;
	case EO_TASKSELECTION:
		mp_EO_TASKSELECTION(MouseMove, X, Y);
		break;
	case EO_TASKADDITION:
		mp_EO_TASKADDITION(MouseMove, X, Y);
		break;
	}
}

void clsMouseKeyboardEvents::OnMouseUp(LONG X,LONG Y)
{
	mp_oControl->GetMouseEventArgs()->SetX(X);
	mp_oControl->GetMouseEventArgs()->SetY(Y);
	mp_oControl->FireControlMouseUp();
	if (mp_oControl->GetMouseEventArgs()->GetCancel() == TRUE) 
	{
		mp_yOperation = EO_NONE;
		return;
	}
	switch (mp_yOperation) 
	{
	case EO_VERTICALSCROLLBAR:
		mp_oControl->GetVerticalScrollBar()->GetScrollBar()->OnMouseUp();
		break;
	case EO_HORIZONTALSCROLLBAR:
		mp_oControl->GetHorizontalScrollBar()->GetScrollBar()->OnMouseUp();
		break;
	case EO_TIMELINESCROLLBAR:
		mp_oControl->f_TimeLineScrollBar()->GetScrollBar()->OnMouseUp();
		break;
	case EO_TREEVIEW:
		mp_oControl->GetTreeview()->OnMouseUp(X, Y);
		break;
	case EO_SPLITTERMOVEMENT:
		mp_EO_SPLITTERMOVEMENT(MouseUp, X, Y);
		break;
	case EO_TIMELINEMOVEMENT:
		mp_EO_TIMELINEMOVEMENT(MouseUp, X, Y);
		break;
	case EO_COLUMNMOVEMENT:
		mp_EO_COLUMNMOVEMENT(MouseUp, X, Y);
		break;
	case EO_COLUMNSIZING:
		mp_EO_COLUMNSIZING(MouseUp, X, Y);
		break;
	case EO_COLUMNSELECTION:
		mp_EO_COLUMNSELECTION(MouseUp, X, Y);
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
	case EO_PERCENTAGESELECTION:
		mp_EO_PERCENTAGESELECTION(MouseUp, X, Y);
		break;
	case EO_PERCENTAGESIZING:
		mp_EO_PERCENTAGESIZING(MouseUp, X, Y);
		break;
	case EO_PREDECESSORADDITION:
		mp_EO_PREDECESSORADDITION(MouseUp, X, Y);
		break;
	case EO_TASKMOVEMENT:
		mp_EO_TASKMOVEMENT(MouseUp, X, Y);
		break;
	case EO_TASKSTRETCHLEFT:
		mp_EO_TASKSTRETCHLEFT(MouseUp, X, Y);
		break;
	case EO_TASKSTRETCHRIGHT:
		mp_EO_TASKSTRETCHRIGHT(MouseUp, X, Y);
		break;
	case EO_TASKSELECTION:
		mp_EO_TASKSELECTION(MouseUp, X, Y);
		break;
	case EO_TASKADDITION:
		mp_EO_TASKADDITION(MouseUp, X, Y);
		break;
    case EO_PREDECESSORSELECTION:
        mp_EO_PREDECESSORSELECTION(MouseUp, X, Y);
        break;
	}
	mp_yOperation = EO_NONE;
}

LONG clsMouseKeyboardEvents::CursorPosition(LONG X,LONG Y)
{

	if (mp_oControl->GetVerticalScrollBar()->GetScrollBar() == NULL)
	{
		return EVT_NONE;
	}

	if (mp_oControl->GetVerticalScrollBar()->GetScrollBar()->OverControl(X, Y) == TRUE) 
	{
		return EVT_VSCROLLBAR;
	}
	else if (mp_oControl->GetHorizontalScrollBar()->GetScrollBar()->OverControl(X, Y) == TRUE) 
	{
		return EVT_HSCROLLBAR;
	}
	else if (mp_oControl->f_TimeLineScrollBar()->GetScrollBar()->OverControl(X, Y) == TRUE) 
	{
		return EVT_TIMELINESCROLLBAR;
	}
	else if (mp_oControl->GetTreeview()->OverControl(X, Y) == TRUE) 
	{
		return EVT_TREEVIEW;
	}
	else if (mp_bOverSplitter(X, Y) == TRUE) 
	{
		return EVT_SPLITTER;
	}
	else if (mp_bOverEmptySpace(Y) == TRUE) 
	{
		return EVT_NONE;
	}
	else if (mp_bOverTimeLine(X, Y) == TRUE) 
	{
		return EVT_TIMELINE;
	}
	else if (mp_bOverSelectedColumn(X, Y) == TRUE) 
	{
		return EVT_SELECTEDCOLUMN;
	}
	else if (mp_bOverColumn(X, Y) == TRUE) 
	{
		return EVT_COLUMN;
	}
	else if (mp_bOverSelectedRow(X, Y) == TRUE) 
	{
		return EVT_SELECTEDROW;
	}
	else if (mp_bOverRow(X, Y) == TRUE) 
	{
		return EVT_ROW;
	}
	else if (mp_bOverSelectedPercentage(X, Y) == TRUE && mp_yOperationModifier != OPM_PREDECESSORMODE) 
	{
		return EVT_SELECTEDPERCENTAGE;
	}
	else if (mp_bOverPercentage(X, Y) == TRUE && mp_yOperationModifier != OPM_PREDECESSORMODE) 
	{
		return EVT_PERCENTAGE;
	}
	else if (mp_bOverSelectedTask(X, Y) == TRUE) 
	{
		return EVT_SELECTEDTASK;
	}
	else if (mp_bOverTask(X, Y) == TRUE) 
	{
		return EVT_TASK;
	}
	else if (mp_bOverSelectedPredecessor(X, Y) == TRUE)
	{
		return EVT_SELECTEDPREDECESSOR;
	}
	else if (mp_bOverPredecessor(X, Y) == TRUE)
	{
		return EVT_PREDECESSOR;
	}
	else if (mp_bOverClientArea(X, Y) == TRUE) 
	{
		return EVT_CLIENTAREA;
	}
	else 
	{
		return EVT_NONE;
	}

}

void clsMouseKeyboardEvents::mp_EO_SPLITTERMOVEMENT(LONG yMouseKeyBoardEvent,LONG X,LONG Y)
{
	if (mp_oControl->GetAllowSplitterMove() == FALSE) 
	{
		return;
	}
	switch (yMouseKeyBoardEvent) 
	{
	case MouseHover:
		mp_SetCursor(CT_MOVESPLITTER);
		break;
	case MouseDown:
		break;
	case MouseMove:
		mp_oControl->clsG->mp_EraseReversibleFrames();
		mp_DrawMovingReversibleFrame(X, 0, X + 2, 0, FCT_VERTICALSPLITTER);
		break;
	case MouseUp:
		LONG lWidth;
		if (X > (mp_oControl->clsG->Width() - 10)) 
		{
			X = mp_oControl->clsG->Width() - 10;
		}

		if (X < 10) 
		{
			X = 10;
		}
		lWidth = mp_oControl->GetColumns()->GetWidth();
		if (X > lWidth) 
		{
			X = lWidth;
			mp_oControl->GetSplitter()->SetPosition(X);
			mp_oControl->GetHorizontalScrollBar()->SetValue(0);
		}
		else if (X > (lWidth - mp_oControl->GetHorizontalScrollBar()->GetValue()))
		{
			X = (lWidth - mp_oControl->GetHorizontalScrollBar()->GetValue());
			mp_oControl->GetSplitter()->SetPosition(X);
		}
		else
		{
			mp_oControl->GetSplitter()->SetPosition(X);
		}

		mp_oControl->clsG->mp_EraseReversibleFrames();
		mp_oControl->Redraw();
		mp_SetCursor(CT_NORMAL);
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

void clsMouseKeyboardEvents::mp_EO_TIMELINEMOVEMENT(LONG yMouseKeyBoardEvent,LONG X,LONG Y)
{
	if (mp_oControl->GetAllowTimeLineScroll() == FALSE) 
	{
		return;
	}
	if (mp_oControl->f_TimeLineScrollBar()->GetEnabled() == TRUE) 
	{
		return;
	}
	switch (yMouseKeyBoardEvent) 
	{
	case MouseHover:
		mp_SetCursor(CT_NORMAL);
		break;
	case MouseDown:
		mp_SetCursor(CT_SCROLLTIMELINE);
		s_tmlnMVT.lX = X;
		break;
	case MouseMove:
		break;
	case MouseUp:
		mp_oControl->GetCurrentViewObject()->GetTimeLine()->Setf_StartDate(mp_oControl->GetMathLib()->DateTimeAdd(mp_oControl->GetCurrentViewObject()->GetInterval(), (s_tmlnMVT.lX - X) * mp_oControl->GetCurrentViewObject()->GetFactor(), mp_oControl->GetCurrentViewObject()->GetTimeLine()->GetStartDate()));
		mp_oControl->Redraw();
		mp_SetCursor(CT_NORMAL);
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

void clsMouseKeyboardEvents::mp_EO_COLUMNMOVEMENT(LONG yMouseKeyBoardEvent,LONG X,LONG Y)
{


	clsColumn* oColumn = NULL;
	clsColumn* oDestinationColumn = NULL;
	clsRow* oRow = NULL;
	int lIndex = 0;
	if (mp_oControl->GetAllowColumnMove() == FALSE) 
	{
		return;
	}
	oColumn = (clsColumn*)mp_oControl->GetColumns()->mp_oCollection->m_oReturnArrayElement(mp_oControl->GetSelectedColumnIndex());
	if (oColumn->GetAllowMove() == FALSE) 
	{
		return;
	}
	switch (yMouseKeyBoardEvent) 
	{
	case MouseHover:
		mp_SetCursor(CT_MOVECOLUMN);
		break;
	case MouseDown:
		s_clmnMVT.lColumnIndex = mp_oControl->GetMathLib()->GetColumnIndexByPosition(X, Y);
		mp_oControl->GetObjectStateChangedEventArgs()->Clear();
		mp_oControl->GetObjectStateChangedEventArgs()->SetEventTarget((LONG)EVT_COLUMN);
		mp_oControl->GetObjectStateChangedEventArgs()->SetIndex(s_clmnMVT.lColumnIndex);
		mp_oControl->GetObjectStateChangedEventArgs()->SetCancel(FALSE);
		mp_oControl->GetObjectStateChangedEventArgs()->SetChangeType(CT_MOVE);
		mp_oControl->FireBeginObjectChange();
		break;
	case MouseMove:
		if (mp_oControl->GetObjectStateChangedEventArgs()->GetCancel() == FALSE) 
		{
			mp_oControl->clsG->mp_EraseReversibleFrames();
			mp_DynamicColumnMove(X);
			s_clmnMVT.lDestinationColumnIndex = mp_oControl->GetMathLib()->GetColumnIndexByPosition(X, Y);
			if (s_clmnMVT.lDestinationColumnIndex >= 1) 
			{
				mp_oControl->clsG->mp_EraseReversibleFrames();
				mp_oControl->GetObjectStateChangedEventArgs()->SetDestinationIndex(s_clmnMVT.lDestinationColumnIndex);
				mp_oControl->FireObjectChange();
				if (mp_oControl->GetObjectStateChangedEventArgs()->GetCancel() == FALSE) 
				{
					oDestinationColumn = (clsColumn*)mp_oControl->GetColumns()->mp_oCollection->m_oReturnArrayElement(s_clmnMVT.lDestinationColumnIndex);
					mp_DrawMovingReversibleFrame(oDestinationColumn->GetLeftTrim(), oDestinationColumn->GetTop(), oDestinationColumn->GetRightTrim(), oDestinationColumn->GetBottom(), FCT_NORMAL);
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
				s_clmnMVT.lDestinationColumnIndex = mp_oControl->GetMathLib()->GetColumnIndexByPosition(X, Y);
				if (s_clmnMVT.lDestinationColumnIndex >= 1 && (s_clmnMVT.lColumnIndex != s_clmnMVT.lDestinationColumnIndex)) 
				{
					mp_oControl->FireCompleteObjectChange();
					if (mp_oControl->GetObjectStateChangedEventArgs()->GetCancel() == FALSE)
					{
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
						mp_oControl->SetSelectedColumnIndex(mp_oControl->GetColumns()->mp_oCollection->m_lCopyAndMoveItems(s_clmnMVT.lColumnIndex, s_clmnMVT.lDestinationColumnIndex));
						for (lIndex = 1; lIndex <= mp_oControl->GetRows()->GetCount(); lIndex++) 
						{
							oRow = (clsRow*)mp_oControl->GetRows()->mp_oCollection->m_oReturnArrayElement(lIndex);
							oRow->GetCells()->mp_oCollection->m_lCopyAndMoveItems(s_clmnMVT.lColumnIndex, s_clmnMVT.lDestinationColumnIndex);
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
				}
			}
		}

		mp_oControl->Redraw();
		mp_SetCursor(CT_NORMAL);
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

void clsMouseKeyboardEvents::mp_EO_COLUMNSIZING(LONG yMouseKeyBoardEvent,LONG X,LONG Y)
{

	clsColumn* oColumn = NULL;
	if (mp_oControl->GetAllowColumnSize() == FALSE) 
	{
		return;
	}
	oColumn = (clsColumn*)mp_oControl->GetColumns()->mp_oCollection->m_oReturnArrayElement(mp_oControl->GetSelectedColumnIndex());
	if (oColumn->GetAllowSize() == FALSE) 
	{
		return;
	}
	switch (yMouseKeyBoardEvent) 
	{
	case MouseHover:
		mp_SetCursor(CT_COLUMNWIDTH);
		break;
	case MouseDown:
		s_clmnSZ.lColumnIndex = mp_oControl->GetMathLib()->GetColumnIndexByPosition(X, Y);
		mp_oControl->GetObjectStateChangedEventArgs()->Clear();
		mp_oControl->GetObjectStateChangedEventArgs()->SetEventTarget((LONG)EVT_COLUMN);
		mp_oControl->GetObjectStateChangedEventArgs()->SetIndex(s_clmnSZ.lColumnIndex);
		mp_oControl->GetObjectStateChangedEventArgs()->SetCancel(FALSE);
        mp_oControl->GetObjectStateChangedEventArgs()->SetChangeType(CT_SIZE);
        mp_oControl->FireBeginObjectChange();
		break;
	case MouseMove:
		if (mp_oControl->GetObjectStateChangedEventArgs()->GetCancel() == FALSE) 
		{
			mp_oControl->clsG->mp_EraseReversibleFrames();
			mp_oControl->FireObjectChange();
			if (mp_oControl->GetObjectStateChangedEventArgs()->GetCancel() == FALSE) 
			{
				mp_DrawMovingReversibleFrame(X, 0, X + 2, mp_oControl->clsG->Height(), FCT_NORMAL);
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
				if (X < mp_oControl->Getmt_BorderThickness()) 
				{
					X = mp_oControl->Getmt_BorderThickness();
				}
				if (X > mp_oControl->GetSplitter()->GetPosition()) 
				{
					mp_oControl->GetSplitter()->SetPosition(X);
				}
				oColumn = (clsColumn*)mp_oControl->GetColumns()->mp_oCollection->m_oReturnArrayElement(s_clmnSZ.lColumnIndex);
				oColumn->SetWidth(oColumn->GetWidth() + (X - oColumn->GetRight()));
				if (oColumn->GetWidth() < mp_oControl->GetMinColumnWidth()) 
				{
					oColumn->SetWidth(mp_oControl->GetMinColumnWidth());
				}
				if (mp_oControl->GetSplitter()->GetPosition() > mp_oControl->GetColumns()->GetWidth()) 
				{
					mp_oControl->GetSplitter()->SetPosition(mp_oControl->GetColumns()->GetWidth());
					mp_oControl->GetHorizontalScrollBar()->SetValue(0);
				}
				mp_oControl->FireCompleteObjectChange();
			}
		}

		mp_oControl->Redraw();
		mp_SetCursor(CT_NORMAL);
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

void clsMouseKeyboardEvents::mp_EO_COLUMNSELECTION(LONG yMouseKeyBoardEvent,LONG X,LONG Y)
{

	switch (yMouseKeyBoardEvent) 
	{
	case MouseHover:
		mp_SetCursor(CT_NORMAL);
		break;
	case MouseDown:
		s_clmnSEL.lColumnIndex = mp_oControl->GetMathLib()->GetColumnIndexByPosition(X, Y);
		break;
	case MouseMove:
		break;
	case MouseUp:
		mp_oControl->SetSelectedColumnIndex(s_clmnSEL.lColumnIndex);
		mp_oControl->GetObjectSelectedEventArgs()->Clear();
		mp_oControl->GetObjectSelectedEventArgs()->SetEventTarget((LONG)EVT_COLUMN);
		mp_oControl->GetObjectSelectedEventArgs()->SetObjectIndex(mp_oControl->GetSelectedColumnIndex());
		mp_oControl->FireObjectSelected();
		mp_oControl->Redraw();
		mp_SetCursor(CT_NORMAL);
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

void clsMouseKeyboardEvents::mp_EO_ROWMOVEMENT(LONG yMouseKeyBoardEvent,LONG X,LONG Y)
{
	clsRow* oRow = NULL;
	clsRow* oDestinationRow = NULL;
	if (mp_oControl->GetAllowRowMove() == FALSE) 
	{
		return;
	}
	oRow = (clsRow*)mp_oControl->GetRows()->mp_oCollection->m_oReturnArrayElement(mp_oControl->GetSelectedRowIndex());
	if (oRow->GetAllowMove() == FALSE) 
	{
		return;
	}
	switch (yMouseKeyBoardEvent) 
	{
	case MouseHover:
		mp_SetCursor(CT_MOVEROW);
		break;
	case MouseDown:
		s_rowMVT.lRowIndex = mp_oControl->GetMathLib()->GetRowIndexByPosition(Y);
		mp_oControl->GetObjectStateChangedEventArgs()->Clear();
		mp_oControl->GetObjectStateChangedEventArgs()->SetEventTarget((LONG)EVT_ROW);
		mp_oControl->GetObjectStateChangedEventArgs()->SetIndex(s_rowMVT.lRowIndex);
		mp_oControl->GetObjectStateChangedEventArgs()->SetInitialRowIndex(s_rowMVT.lRowIndex);
		mp_oControl->GetObjectStateChangedEventArgs()->SetCancel(FALSE);
		mp_oControl->GetObjectStateChangedEventArgs()->SetChangeType(CT_MOVE);
		mp_oControl->FireBeginObjectChange();
		break;
	case MouseMove:
		if (mp_oControl->GetObjectStateChangedEventArgs()->GetCancel() == FALSE) 
		{
			mp_oControl->clsG->mp_EraseReversibleFrames();
			mp_DynamicRowMove(Y);
			s_rowMVT.lDestinationRowIndex = mp_oControl->GetMathLib()->GetRowIndexByPosition(Y);
			if (s_rowMVT.lDestinationRowIndex >= 1) 
			{
				mp_oControl->clsG->mp_EraseReversibleFrames();
				mp_oControl->GetObjectStateChangedEventArgs()->SetDestinationIndex(s_rowMVT.lDestinationRowIndex);
				mp_oControl->GetObjectStateChangedEventArgs()->SetFinalRowIndex(s_rowMVT.lDestinationRowIndex);
				mp_oControl->FireObjectChange();
				if (mp_oControl->GetObjectStateChangedEventArgs()->GetCancel() == FALSE) 
				{
					oDestinationRow = (clsRow*)mp_oControl->GetRows()->mp_oCollection->m_oReturnArrayElement(s_rowMVT.lDestinationRowIndex);
					mp_DrawMovingReversibleFrame(oDestinationRow->GetLeft(), oDestinationRow->GetTop(), oDestinationRow->GetLeft(), oDestinationRow->GetBottom(), FCT_NORMAL);
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
					mp_oControl->SetSelectedRowIndex(mp_oControl->GetRows()->mp_oCollection->m_lCopyAndMoveItems(s_rowMVT.lRowIndex, s_rowMVT.lDestinationRowIndex));
					mp_oControl->FireCompleteObjectChange();
				}
			}
		}

		mp_oControl->Redraw();
		mp_SetCursor(CT_NORMAL);
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

void clsMouseKeyboardEvents::mp_EO_ROWSIZING(LONG yMouseKeyBoardEvent,LONG X,LONG Y)
{

	clsRow* oRow = NULL;
	if (mp_oControl->GetAllowRowSize() == FALSE) 
	{
		return;
	}
	oRow = (clsRow*)mp_oControl->GetRows()->mp_oCollection->m_oReturnArrayElement(mp_oControl->GetSelectedRowIndex());
	if (oRow->GetAllowSize() == FALSE) 
	{
		return;
	}
	switch (yMouseKeyBoardEvent) 
	{
	case MouseHover:
		mp_SetCursor(CT_ROWHEIGHT);
		break;
	case MouseDown:
		s_rowSZ.lRowIndex = mp_oControl->GetMathLib()->GetRowIndexByPosition(Y);
		mp_oControl->GetObjectStateChangedEventArgs()->Clear();
		mp_oControl->GetObjectStateChangedEventArgs()->SetEventTarget((LONG)EVT_ROW);
		mp_oControl->GetObjectStateChangedEventArgs()->SetIndex(s_rowSZ.lRowIndex);
        mp_oControl->GetObjectStateChangedEventArgs()->SetChangeType(CT_SIZE);
        mp_oControl->FireBeginObjectChange();
		break;
	case MouseMove:
		if (mp_oControl->GetObjectStateChangedEventArgs()->GetCancel() == FALSE) 
		{
			mp_oControl->clsG->mp_EraseReversibleFrames();
			mp_oControl->FireObjectChange();
			if (mp_oControl->GetObjectStateChangedEventArgs()->GetCancel() == FALSE) 
			{
				mp_DrawMovingReversibleFrame(0, Y, mp_oControl->clsG->Width(), Y + 2, FCT_NORMAL);
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
		mp_SetCursor(CT_NORMAL);
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

void clsMouseKeyboardEvents::mp_EO_ROWSELECTION(LONG yMouseKeyBoardEvent,LONG X,LONG Y)
{

	clsRow* oRow = NULL;
	switch (yMouseKeyBoardEvent) 
	{
	case MouseHover:
		mp_SetCursor(CT_NORMAL);
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
		if (oRow->GetMergeCells() == TRUE) 
		{
			mp_oControl->GetObjectSelectedEventArgs()->Clear();
			mp_oControl->GetObjectSelectedEventArgs()->SetEventTarget((LONG)EVT_ROW);
			mp_oControl->GetObjectSelectedEventArgs()->SetObjectIndex(mp_oControl->GetSelectedRowIndex());
			mp_oControl->FireObjectSelected();
		}
		else 
		{
			mp_oControl->GetObjectSelectedEventArgs()->Clear();
			mp_oControl->GetObjectSelectedEventArgs()->SetEventTarget((LONG)EVT_CELL);
			mp_oControl->GetObjectSelectedEventArgs()->SetObjectIndex(mp_oControl->GetSelectedCellIndex());
			mp_oControl->GetObjectSelectedEventArgs()->SetParentObjectIndex(mp_oControl->GetSelectedRowIndex());
			mp_oControl->FireObjectSelected();
		}

		mp_oControl->Redraw();
		mp_SetCursor(CT_NORMAL);
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

void clsMouseKeyboardEvents::mp_EO_TASKMOVEMENT(LONG yMouseKeyBoardEvent,LONG X,LONG Y)
{

	clsTask* oTask = NULL;
	clsRow* oRow = NULL;
	if (mp_oControl->GetAllowEdit() == FALSE) 
	{
		return;
	}
	oTask = (clsTask*)mp_oControl->GetTasks()->mp_oCollection->m_oReturnArrayElement(mp_oControl->GetSelectedTaskIndex());
	if (oTask->GetAllowedMovement() == MT_MOVEMENTDISABLED) 
	{
		return;
	}
	switch (yMouseKeyBoardEvent) 
	{
	case MouseHover:
		mp_SetCursor(CT_MOVETASK);
		break;
	case MouseDown:
		s_tskMVT.Clear();
		s_tskMVT.lInitialRowIndex = mp_oControl->GetMathLib()->GetRowIndexByPosition(Y);
		X = mp_fSnapX(X);
		mp_oControl->GetObjectStateChangedEventArgs()->Clear();
		mp_oControl->GetObjectStateChangedEventArgs()->SetEventTarget((LONG)EVT_TASK);
		mp_oControl->GetObjectStateChangedEventArgs()->SetIndex(mp_oControl->GetSelectedTaskIndex());
		mp_oControl->GetObjectStateChangedEventArgs()->SetInitialRowIndex(s_tskMVT.lInitialRowIndex);
		mp_oControl->GetObjectStateChangedEventArgs()->SetStartDate(oTask->GetStartDate());
		mp_oControl->GetObjectStateChangedEventArgs()->SetEndDate(oTask->GetEndDate());
		mp_oControl->GetObjectStateChangedEventArgs()->SetCancel(FALSE);
        mp_oControl->GetObjectStateChangedEventArgs()->SetChangeType(CT_MOVE);
        mp_oControl->FireBeginObjectChange();
		if (mp_oControl->GetObjectStateChangedEventArgs()->GetCancel() == FALSE) 
		{
			s_tskMVT.lDeltaLeft = X - mp_oControl->GetMathLib()->GetXCoordinateFromDate(oTask->GetStartDate());
			s_tskMVT.lDeltaRight = mp_oControl->GetMathLib()->GetXCoordinateFromDate(oTask->GetEndDate()) - X;
			s_tskMVT.dtInitialStartDate = oTask->GetStartDate();
			s_tskMVT.dtInitialEndDate = oTask->GetEndDate();
			s_tskMVT.sInitialRowKey = oTask->GetRowKey();
			s_tskMVT.lDurationFactor = oTask->GetDurationFactor();
		}
		break;
    case MouseMove:
        if (oTask->GetTaskType() == TT_START_END)
        {
            if (mp_oControl->GetObjectStateChangedEventArgs()->GetCancel() == FALSE)
            {
                X = mp_fSnapX(X);
                s_tskMVT.lFinalRowIndex = mp_oControl->GetMathLib()->GetRowIndexByPosition(Y);
                mp_oControl->clsG->mp_EraseReversibleFrames();
                mp_oControl->GetObjectStateChangedEventArgs()->SetFinalRowIndex(s_tskMVT.lFinalRowIndex);
                mp_oControl->GetObjectStateChangedEventArgs()->SetStartDate(mp_oControl->GetMathLib()->GetDateFromXCoordinate(X - s_tskMVT.lDeltaLeft));
                mp_oControl->GetObjectStateChangedEventArgs()->SetEndDate(mp_oControl->GetMathLib()->GetDateFromXCoordinate(X + s_tskMVT.lDeltaRight));
                mp_oControl->FireObjectChange();
                mp_DynamicRowMove(Y);
                if (mp_oControl->GetObjectStateChangedEventArgs()->GetCancel() == FALSE && (s_tskMVT.lFinalRowIndex >= 1 && s_tskMVT.lFinalRowIndex <= mp_oControl->GetRows()->GetCount())) 
                {
                    if (X < 0 || X > mp_oControl->clsG->Width() || Y < 0 || Y > mp_oControl->clsG->Height())
                    {
                        mp_SetCursor(CT_NODROP);
                        mp_oControl->clsG->mp_DrawReversibleFrameEx();
                    }
                    oRow = (clsRow*)mp_oControl->GetRows()->mp_oCollection->m_oReturnArrayElement(s_tskMVT.lFinalRowIndex);
                    if ((s_tskMVT.lInitialRowIndex != s_tskMVT.lFinalRowIndex && oTask->GetAllowedMovement() == MT_RESTRICTEDTOROW) || (oRow->GetContainer() == FALSE) || (mp_oControl->GetMathLib()->DetectConflict(mp_oControl->GetMathLib()->GetDateFromXCoordinate(mp_fSnapX(X - s_tskMVT.lDeltaLeft)), mp_oControl->GetMathLib()->GetDateFromXCoordinate(mp_fSnapX(X + s_tskMVT.lDeltaRight)), oRow->GetKey(), mp_oControl->GetSelectedTaskIndex(), oTask->GetLayerIndex()) == TRUE && mp_oControl->GetCurrentViewObject()->GetClientArea()->GetDetectConflicts() == TRUE)) 
                    {
                        mp_SetCursor(CT_NODROP);
                        mp_oControl->clsG->mp_DrawReversibleFrameEx();
                        return;
                    }
                    mp_SetCursor(CT_MOVETASK);
                    mp_DynamicTimeLineMove(X);
                    mp_oControl->GetToolTipEventArgs()->Clear();
                    mp_oControl->GetToolTipEventArgs()->SetTaskIndex(mp_oControl->GetSelectedTaskIndex());
                    mp_oControl->GetToolTipEventArgs()->SetInitialRowIndex(s_tskMVT.lInitialRowIndex);
                    mp_oControl->GetToolTipEventArgs()->SetFinalRowIndex(s_tskMVT.lFinalRowIndex);
                    mp_oControl->GetToolTipEventArgs()->SetInitialStartDate(s_tskMVT.dtInitialStartDate);
                    mp_oControl->GetToolTipEventArgs()->SetInitialEndDate(s_tskMVT.dtInitialEndDate);
                    mp_oControl->GetToolTipEventArgs()->SetStartDate(mp_oControl->GetMathLib()->GetDateFromXCoordinate(X - s_tskMVT.lDeltaLeft));
                    mp_oControl->GetToolTipEventArgs()->SetEndDate(mp_oControl->GetMathLib()->GetDateFromXCoordinate(X + s_tskMVT.lDeltaRight));
                    mp_oControl->GetToolTipEventArgs()->SetX(X);
                    mp_oControl->GetToolTipEventArgs()->SetY(Y);
                    mp_oControl->FireToolTipOnMouseMove(mp_yOperation);
                    mp_DrawMovingReversibleFrame(X - s_tskMVT.lDeltaLeft, oRow->GetTop(), X + s_tskMVT.lDeltaRight, oRow->GetBottom(), FCT_KEEPLEFTRIGHTBOUNDS);
                }
            }
        }
        else if (oTask->GetTaskType() == TT_DURATION)
        {
            if (mp_oControl->GetObjectStateChangedEventArgs()->GetCancel() == FALSE)
            {
                COleDateTime dtStartDate;
                COleDateTime dtEndDate;
                X = mp_fSnapX(X);
                s_tskMVT.lFinalRowIndex = mp_oControl->GetMathLib()->GetRowIndexByPosition(Y);
                mp_oControl->clsG->mp_EraseReversibleFrames();
                mp_oControl->GetObjectStateChangedEventArgs()->SetFinalRowIndex(s_tskMVT.lFinalRowIndex);
                dtStartDate = mp_oControl->GetMathLib()->GetDateFromXCoordinate(X - s_tskMVT.lDeltaLeft);
                dtEndDate = mp_oControl->GetMathLib()->GetEndDate(dtStartDate, oTask->GetDurationInterval(), oTask->GetDurationFactor());
                mp_oControl->GetObjectStateChangedEventArgs()->SetStartDate(dtStartDate);
                mp_oControl->GetObjectStateChangedEventArgs()->SetEndDate(dtEndDate);
                mp_oControl->FireObjectChange();
                mp_DynamicRowMove(Y);
                if (mp_oControl->GetObjectStateChangedEventArgs()->GetCancel() == FALSE && (s_tskMVT.lFinalRowIndex >= 1 && s_tskMVT.lFinalRowIndex <= mp_oControl->GetRows()->GetCount())) 
                {
                    if (X < 0 || X > mp_oControl->clsG->Width() || Y < 0 || Y > mp_oControl->clsG->Height())
                    {
                        mp_SetCursor(CT_NODROP);
                        mp_oControl->clsG->mp_DrawReversibleFrameEx();
                    }
                    oRow = (clsRow*)mp_oControl->GetRows()->mp_oCollection->m_oReturnArrayElement(s_tskMVT.lFinalRowIndex);
                    if ((s_tskMVT.lInitialRowIndex != s_tskMVT.lFinalRowIndex && oTask->GetAllowedMovement() == MT_RESTRICTEDTOROW) || (oRow->GetContainer() == FALSE) || (mp_oControl->GetMathLib()->DetectConflict(mp_oControl->GetMathLib()->GetDateFromXCoordinate(mp_fSnapX(X - s_tskMVT.lDeltaLeft)), mp_oControl->GetMathLib()->GetDateFromXCoordinate(mp_fSnapX(X + s_tskMVT.lDeltaRight)), oRow->GetKey(), mp_oControl->GetSelectedTaskIndex(), oTask->GetLayerIndex()) == TRUE && mp_oControl->GetCurrentViewObject()->GetClientArea()->GetDetectConflicts() == TRUE))
                    {
                        mp_SetCursor(CT_NODROP);
                        mp_oControl->clsG->mp_DrawReversibleFrameEx();
                        return;
                    }
                    mp_SetCursor(CT_MOVETASK);
                    mp_DynamicTimeLineMove(X);
                    mp_oControl->GetToolTipEventArgs()->Clear();
                    mp_oControl->GetToolTipEventArgs()->SetTaskIndex(mp_oControl->GetSelectedTaskIndex());
                    mp_oControl->GetToolTipEventArgs()->SetInitialRowIndex(s_tskMVT.lInitialRowIndex);
                    mp_oControl->GetToolTipEventArgs()->SetFinalRowIndex(s_tskMVT.lFinalRowIndex);
                    mp_oControl->GetToolTipEventArgs()->SetInitialStartDate(s_tskMVT.dtInitialStartDate);
                    mp_oControl->GetToolTipEventArgs()->SetInitialEndDate(s_tskMVT.dtInitialEndDate);
                    mp_oControl->GetToolTipEventArgs()->SetStartDate(dtStartDate);
                    mp_oControl->GetToolTipEventArgs()->SetEndDate(dtEndDate);
                    mp_oControl->GetToolTipEventArgs()->SetX(X);
                    mp_oControl->GetToolTipEventArgs()->SetY(Y);
                    mp_oControl->FireToolTipOnMouseMove(mp_yOperation);
                    mp_DrawMovingReversibleFrame(mp_oControl->GetMathLib()->GetXCoordinateFromDate(dtStartDate), oRow->GetTop(), mp_oControl->GetMathLib()->GetXCoordinateFromDate(dtEndDate), oRow->GetBottom(), FCT_KEEPLEFTRIGHTBOUNDS);
                }
            }
        }
        break;
    case MouseUp:
        if (oTask->GetTaskType() == TT_START_END)
        {
            if (mp_oControl->GetObjectStateChangedEventArgs()->GetCancel() == FALSE && (s_tskMVT.lFinalRowIndex >= 1 && s_tskMVT.lFinalRowIndex <= mp_oControl->GetRows()->GetCount())) 
            {
                oRow = (clsRow*)mp_oControl->GetRows()->mp_oCollection->m_oReturnArrayElement(s_tskMVT.lFinalRowIndex);
                mp_oControl->clsG->mp_EraseReversibleFrames();
                if ((s_tskMVT.lInitialRowIndex != s_tskMVT.lFinalRowIndex && oTask->GetAllowedMovement() == MT_RESTRICTEDTOROW) || (oRow->GetContainer() == FALSE) || (mp_oControl->GetMathLib()->DetectConflict(mp_oControl->GetMathLib()->GetDateFromXCoordinate(mp_fSnapX(X - s_tskMVT.lDeltaLeft)), mp_oControl->GetMathLib()->GetDateFromXCoordinate(mp_fSnapX(X + s_tskMVT.lDeltaRight)), oRow->GetKey(), mp_oControl->GetSelectedTaskIndex(), oTask->GetLayerIndex()) == TRUE && mp_oControl->GetCurrentViewObject()->GetClientArea()->GetDetectConflicts() == TRUE))
                {
                }
                else
                {
                    X = mp_fSnapX(X);
                    mp_oControl->FireEndObjectChange();
                    if (mp_oControl->GetObjectStateChangedEventArgs()->GetCancel() == FALSE) 
                    {
                        oTask->SetStartDate(mp_oControl->GetMathLib()->GetDateFromXCoordinate(X - s_tskMVT.lDeltaLeft));
                        oTask->SetEndDate(mp_oControl->GetMathLib()->GetDateFromXCoordinate(X + s_tskMVT.lDeltaRight));
                        if (mp_oControl->GetCurrentViewObject()->GetClientArea()->GetGrid()->GetSnapToGrid() == TRUE)
                        {
                            oTask->SetStartDate(mp_oControl->GetMathLib()->RoundDate(mp_oControl->GetCurrentViewObject()->GetClientArea()->GetGrid()->GetInterval(), mp_oControl->GetCurrentViewObject()->GetClientArea()->GetGrid()->GetFactor(), oTask->GetStartDate()));
                            oTask->SetEndDate(mp_oControl->GetMathLib()->RoundDate(mp_oControl->GetCurrentViewObject()->GetClientArea()->GetGrid()->GetInterval(), mp_oControl->GetCurrentViewObject()->GetClientArea()->GetGrid()->GetFactor(), oTask->GetEndDate()));
                        }
                        oTask->SetRowKey(oRow->GetKey());
                        if (oTask->GetStartDate() != s_tskMVT.dtInitialStartDate || oTask->GetEndDate() != s_tskMVT.dtInitialEndDate || oTask->GetRowKey() != s_tskMVT.sInitialRowKey) 
                        {
                            mp_oControl->FireCompleteObjectChange();
                            if (mp_oControl->GetObjectStateChangedEventArgs()->GetCancel() == TRUE)
                            {
                                ////Account for Duration
                                oTask->SetStartDate(s_tskMVT.dtInitialStartDate);
                                oTask->SetEndDate(s_tskMVT.dtInitialEndDate);
                                oTask->SetRowKey(s_tskMVT.sInitialRowKey);
                            }
                        }
                    }
                }
            }
            if (mp_oControl->GetEnforcePredecessors() == TRUE)
            {
                mp_oControl->CheckPredecessors();
            }
            mp_oControl->Redraw();
            mp_SetCursor(CT_NORMAL);
        }
        else if (oTask->GetTaskType() == TT_DURATION)
        {
            if (mp_oControl->GetObjectStateChangedEventArgs()->GetCancel() == FALSE && (s_tskMVT.lFinalRowIndex >= 1 && s_tskMVT.lFinalRowIndex <= mp_oControl->GetRows()->GetCount())) 
            {
                COleDateTime dtStartDate;
                COleDateTime dtEndDate;
                LONG lDuration = 0;
                oRow = (clsRow*)mp_oControl->GetRows()->mp_oCollection->m_oReturnArrayElement(s_tskMVT.lFinalRowIndex);
                mp_oControl->clsG->mp_EraseReversibleFrames();
                if ((s_tskMVT.lInitialRowIndex != s_tskMVT.lFinalRowIndex && oTask->GetAllowedMovement() == MT_RESTRICTEDTOROW) || (oRow->GetContainer() == FALSE) || (mp_oControl->GetMathLib()->DetectConflict(mp_oControl->GetMathLib()->GetDateFromXCoordinate(mp_fSnapX(X - s_tskMVT.lDeltaLeft)), mp_oControl->GetMathLib()->GetDateFromXCoordinate(mp_fSnapX(X + s_tskMVT.lDeltaRight)), oRow->GetKey(), mp_oControl->GetSelectedTaskIndex(), oTask->GetLayerIndex()) == TRUE && mp_oControl->GetCurrentViewObject()->GetClientArea()->GetDetectConflicts() == TRUE))
                {
                }
                else
                {
                    X = mp_fSnapX(X);
                    mp_oControl->FireEndObjectChange();
                    if (mp_oControl->GetObjectStateChangedEventArgs()->GetCancel() == FALSE)
                    {
                        dtStartDate = mp_oControl->GetMathLib()->GetDateFromXCoordinate(X - s_tskMVT.lDeltaLeft);
                        dtEndDate = mp_oControl->GetMathLib()->GetEndDate(dtStartDate, oTask->GetDurationInterval(), oTask->GetDurationFactor());
                        if (mp_oControl->GetCurrentViewObject()->GetClientArea()->GetGrid()->GetSnapToGrid() == TRUE)
                        {
                            dtStartDate = mp_oControl->GetMathLib()->RoundDate(mp_oControl->GetCurrentViewObject()->GetClientArea()->GetGrid()->GetInterval(), mp_oControl->GetCurrentViewObject()->GetClientArea()->GetGrid()->GetFactor(), oTask->GetStartDate());
                        }
                        lDuration = mp_oControl->GetMathLib()->CalculateDuration(dtStartDate, dtEndDate, oTask->GetDurationInterval());
						if (lDuration != s_tskMVT.lDurationFactor)
						{
							mp_oControl->mp_ErrorReport(ERR_DURATION_INCONSISTENT, _T("Inconsistent duration"), _T("clsMouseKeyboardEvents.mp_EO_TASKMOVEMENT"));
						}
                        oTask->SetStartDate(dtStartDate);
                        oTask->SetRowKey(oRow->GetKey());
                        if (oTask->GetStartDate() != s_tskMVT.dtInitialStartDate || oTask->GetEndDate() != s_tskMVT.dtInitialEndDate || oTask->GetRowKey() != s_tskMVT.sInitialRowKey) 
                        {
                            mp_oControl->FireCompleteObjectChange();
                            if (mp_oControl->GetObjectStateChangedEventArgs()->GetCancel() == TRUE)
                            {
                                oTask->SetStartDate(s_tskMVT.dtInitialStartDate);
                                oTask->SetEndDate(s_tskMVT.dtInitialEndDate);
								oTask->SetRowKey(s_tskMVT.sInitialRowKey);
                            }
                        }
                    }
                }
            }
            if (mp_oControl->GetEnforcePredecessors() == TRUE)
            {
                mp_oControl->CheckPredecessors();
            }
            mp_oControl->Redraw();
            mp_SetCursor(CT_NORMAL);
        }
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

void clsMouseKeyboardEvents::mp_EO_TASKSTRETCHLEFT(LONG yMouseKeyBoardEvent,LONG X,LONG Y)
{

	clsTask* oTask = NULL;
	clsRow* oRow = NULL;
	if (mp_oControl->GetAllowEdit() == FALSE) 
	{
		return;
	}
	oTask = (clsTask*)mp_oControl->GetTasks()->mp_oCollection->m_oReturnArrayElement(mp_oControl->GetSelectedTaskIndex());
	if (oTask->GetAllowStretchLeft() == FALSE || oTask->GetTaskType() == TT_DURATION)
	{
		return;
	}
	switch (yMouseKeyBoardEvent) 
	{
	case MouseHover:
		mp_SetCursor(CT_SIZETASK);
		break;
	case MouseDown:
		mp_oControl->GetObjectStateChangedEventArgs()->SetCancel(FALSE);
		s_tskSTL.lRowIndex = mp_oControl->GetMathLib()->GetRowIndexByPosition(Y);
		X = mp_fSnapX(X);
		mp_oControl->GetObjectStateChangedEventArgs()->Clear();
		mp_oControl->GetObjectStateChangedEventArgs()->SetEventTarget((LONG)EVT_TASK);
		mp_oControl->GetObjectStateChangedEventArgs()->SetIndex(mp_oControl->GetSelectedTaskIndex());
		mp_oControl->GetObjectStateChangedEventArgs()->SetCancel(FALSE);
        mp_oControl->GetObjectStateChangedEventArgs()->SetChangeType(CT_SIZE);
        mp_oControl->FireBeginObjectChange();
		if (mp_oControl->GetObjectStateChangedEventArgs()->GetCancel() == FALSE) 
		{
			s_tskSTL.dtInitialStartDate = oTask->GetStartDate();
			s_tskSTL.dtInitialEndDate = oTask->GetEndDate();
		}

		break;
	case MouseMove:
		if (mp_oControl->GetObjectStateChangedEventArgs()->GetCancel() == FALSE) 
		{
			X = mp_fSnapX(X);
			mp_oControl->clsG->mp_EraseReversibleFrames();
			oRow = (clsRow*)mp_oControl->GetRows()->mp_oCollection->m_oReturnArrayElement(s_tskSTL.lRowIndex);
			s_tskSTL.dtFinalStartDate = mp_oControl->GetMathLib()->GetDateFromXCoordinate(X);
			mp_oControl->FireObjectChange();
			if (mp_oControl->GetObjectStateChangedEventArgs()->GetCancel() == FALSE) 
			{
				if (mp_oControl->GetMathLib()->DetectConflict(mp_oControl->GetMathLib()->GetDateFromXCoordinate(X), oTask->GetEndDate(), oRow->GetKey(), mp_oControl->GetSelectedTaskIndex(), oTask->GetLayerIndex()) == TRUE && mp_oControl->GetCurrentViewObject()->GetClientArea()->GetDetectConflicts() == TRUE) 
				{
					mp_SetCursor(CT_NODROP);
				}
				else 
				{
					mp_SetCursor(CT_SIZETASK);
				}
				mp_DynamicTimeLineMove(X);
				mp_oControl->GetToolTipEventArgs()->Clear();
				mp_oControl->GetToolTipEventArgs()->SetTaskIndex(mp_oControl->GetSelectedTaskIndex());
				mp_oControl->GetToolTipEventArgs()->SetRowIndex(s_tskSTL.lRowIndex);
				mp_oControl->GetToolTipEventArgs()->SetInitialStartDate(s_tskSTL.dtInitialStartDate);
				mp_oControl->GetToolTipEventArgs()->SetInitialEndDate(s_tskSTL.dtInitialEndDate);
				mp_oControl->GetToolTipEventArgs()->SetStartDate(s_tskSTL.dtFinalStartDate);
				mp_oControl->GetToolTipEventArgs()->SetEndDate(s_tskSTL.dtInitialEndDate);
				mp_oControl->GetToolTipEventArgs()->SetX(X);
				mp_oControl->GetToolTipEventArgs()->SetY(Y);
				mp_oControl->FireToolTipOnMouseMove(mp_yOperation);
				mp_DrawMovingReversibleFrame(X, oRow->GetTop(), mp_oControl->GetMathLib()->GetXCoordinateFromDate(oTask->GetEndDate()), oRow->GetBottom(), FCT_KEEPLEFTRIGHTBOUNDS);
			}
		}

		break;
	case MouseUp:
		if (s_tskSTL.dtFinalStartDate >= s_tskSTL.dtInitialEndDate) 
		{
			mp_oControl->GetObjectStateChangedEventArgs()->SetCancel(TRUE);
		}

		if (mp_oControl->GetObjectStateChangedEventArgs()->GetCancel() == FALSE) 
		{
			mp_oControl->clsG->mp_EraseReversibleFrames();
			if (s_tskSTL.dtFinalStartDate < mp_oControl->GetCurrentViewObject()->GetTimeLine()->GetStartDate()) 
			{
				s_tskSTL.dtFinalStartDate = mp_oControl->GetCurrentViewObject()->GetTimeLine()->GetStartDate();
			}
			mp_oControl->FireEndObjectChange();
			if (mp_oControl->GetObjectStateChangedEventArgs()->GetCancel() == FALSE) 
			{
				oRow = (clsRow*)mp_oControl->GetRows()->mp_oCollection->m_oReturnArrayElement(s_tskSTL.lRowIndex);
				if (mp_oControl->GetMathLib()->DetectConflict(mp_oControl->GetMathLib()->GetDateFromXCoordinate(X), oTask->GetEndDate(), oRow->GetKey(), mp_oControl->GetSelectedTaskIndex(), oTask->GetLayerIndex()) == TRUE && mp_oControl->GetCurrentViewObject()->GetClientArea()->GetDetectConflicts() == TRUE) 
				{
				}
				else 
				{
					oTask->SetStartDate(s_tskSTL.dtFinalStartDate);
					if (mp_oControl->GetCurrentViewObject()->GetClientArea()->GetGrid()->GetSnapToGrid() == TRUE) 
					{
						oTask->SetStartDate(mp_oControl->GetMathLib()->RoundDate(mp_oControl->GetCurrentViewObject()->GetClientArea()->GetGrid()->GetInterval(), mp_oControl->GetCurrentViewObject()->GetClientArea()->GetGrid()->GetFactor(), oTask->GetStartDate()));
						oTask->SetEndDate(mp_oControl->GetMathLib()->RoundDate(mp_oControl->GetCurrentViewObject()->GetClientArea()->GetGrid()->GetInterval(), mp_oControl->GetCurrentViewObject()->GetClientArea()->GetGrid()->GetFactor(), oTask->GetEndDate()));
					}
					if (oTask->GetStartDate() != s_tskSTL.dtInitialStartDate) 
					{
						mp_oControl->FireCompleteObjectChange();
						if (mp_oControl->GetObjectStateChangedEventArgs()->GetCancel() == TRUE)
						{
							oTask->SetStartDate(s_tskSTL.dtInitialStartDate);
						}
					}
				}
			}
		}
		if (mp_oControl->GetEnforcePredecessors() == TRUE)
		{
			mp_oControl->CheckPredecessors();
		}
		mp_oControl->Redraw();
		mp_SetCursor(CT_NORMAL);
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

void clsMouseKeyboardEvents::mp_EO_TASKSTRETCHRIGHT(LONG yMouseKeyBoardEvent,LONG X,LONG Y)
{

	clsTask* oTask = NULL;
	clsRow* oRow = NULL;
	if (mp_oControl->GetAllowEdit() == FALSE) 
	{
		return;
	}
	oTask = (clsTask*)mp_oControl->GetTasks()->mp_oCollection->m_oReturnArrayElement(mp_oControl->GetSelectedTaskIndex());
	if (oTask->GetAllowStretchRight() == FALSE) 
	{
		return;
	}
	switch (yMouseKeyBoardEvent) 
	{
	case MouseHover:
		mp_SetCursor(CT_SIZETASK);
		break;
	case MouseDown:
		s_tskSTR.lRowIndex = mp_oControl->GetMathLib()->GetRowIndexByPosition(Y);
		X = mp_fSnapX(X);
		mp_oControl->GetObjectStateChangedEventArgs()->Clear();
		mp_oControl->GetObjectStateChangedEventArgs()->SetEventTarget((LONG)EVT_TASK);
		mp_oControl->GetObjectStateChangedEventArgs()->SetIndex(mp_oControl->GetSelectedTaskIndex());
		mp_oControl->GetObjectStateChangedEventArgs()->SetCancel(FALSE);
        mp_oControl->GetObjectStateChangedEventArgs()->SetChangeType(CT_SIZE);
        mp_oControl->FireBeginObjectChange();
		if (mp_oControl->GetObjectStateChangedEventArgs()->GetCancel() == FALSE) 
		{
			s_tskSTR.dtInitialStartDate = oTask->GetStartDate();
			s_tskSTR.dtInitialEndDate = oTask->GetEndDate();
		}

		break;
	case MouseMove:
		if (mp_oControl->GetObjectStateChangedEventArgs()->GetCancel() == FALSE) 
		{
			X = mp_fSnapX(X);
			mp_oControl->clsG->mp_EraseReversibleFrames();
			oRow = (clsRow*)mp_oControl->GetRows()->mp_oCollection->m_oReturnArrayElement(s_tskSTR.lRowIndex);
			s_tskSTR.dtFinalEndDate = mp_oControl->GetMathLib()->GetDateFromXCoordinate(X);
			mp_oControl->FireObjectChange();
			if (mp_oControl->GetObjectStateChangedEventArgs()->GetCancel() == FALSE) 
			{
				if (mp_oControl->GetMathLib()->DetectConflict(oTask->GetStartDate(), mp_oControl->GetMathLib()->GetDateFromXCoordinate(X), oRow->GetKey(), mp_oControl->GetSelectedTaskIndex(), oTask->GetLayerIndex()) == TRUE && mp_oControl->GetCurrentViewObject()->GetClientArea()->GetDetectConflicts() == TRUE) 
				{
					mp_SetCursor(CT_NODROP);
				}
				else 
				{
					mp_SetCursor(CT_SIZETASK);
				}
				mp_DynamicTimeLineMove(X);
				mp_oControl->GetToolTipEventArgs()->Clear();
				mp_oControl->GetToolTipEventArgs()->SetTaskIndex(mp_oControl->GetSelectedTaskIndex());
				mp_oControl->GetToolTipEventArgs()->SetRowIndex(s_tskSTR.lRowIndex);
				mp_oControl->GetToolTipEventArgs()->SetInitialStartDate(s_tskSTR.dtInitialStartDate);
				mp_oControl->GetToolTipEventArgs()->SetInitialEndDate(s_tskSTR.dtInitialEndDate);
				mp_oControl->GetToolTipEventArgs()->SetStartDate(s_tskSTR.dtInitialStartDate);
				mp_oControl->GetToolTipEventArgs()->SetEndDate(s_tskSTR.dtFinalEndDate);
				mp_oControl->GetToolTipEventArgs()->SetX(X);
				mp_oControl->GetToolTipEventArgs()->SetY(Y);
				mp_oControl->FireToolTipOnMouseMove(mp_yOperation);
				mp_DrawMovingReversibleFrame(mp_oControl->GetMathLib()->GetXCoordinateFromDate(oTask->GetStartDate()), oRow->GetTop(), X, oRow->GetBottom(), FCT_KEEPLEFTRIGHTBOUNDS);
			}
		}

		break;
	case MouseUp:
		if (s_tskSTR.dtFinalEndDate <= s_tskSTR.dtInitialStartDate) 
		{
			mp_oControl->GetObjectStateChangedEventArgs()->SetCancel(TRUE);
		}

		if (mp_oControl->GetObjectStateChangedEventArgs()->GetCancel() == FALSE) 
		{
			mp_oControl->clsG->mp_EraseReversibleFrames();
			if (s_tskSTR.dtFinalEndDate > mp_oControl->GetCurrentViewObject()->GetTimeLine()->GetEndDate()) 
			{
				s_tskSTR.dtFinalEndDate = mp_oControl->GetCurrentViewObject()->GetTimeLine()->GetEndDate();
			}
			mp_oControl->FireEndObjectChange();
			if (mp_oControl->GetObjectStateChangedEventArgs()->GetCancel() == FALSE) 
			{
				oRow = (clsRow*)mp_oControl->GetRows()->mp_oCollection->m_oReturnArrayElement(s_tskSTR.lRowIndex);
				if (mp_oControl->GetMathLib()->DetectConflict(oTask->GetStartDate(), mp_oControl->GetMathLib()->GetDateFromXCoordinate(X), oRow->GetKey(), mp_oControl->GetSelectedTaskIndex(), oTask->GetLayerIndex()) == TRUE && mp_oControl->GetCurrentViewObject()->GetClientArea()->GetDetectConflicts() == TRUE) 
				{
				}
				else 
				{
		            if (oTask->GetTaskType() == TT_START_END)
                    {
						oTask->SetEndDate(s_tskSTR.dtFinalEndDate);
						if (mp_oControl->GetCurrentViewObject()->GetClientArea()->GetGrid()->GetSnapToGrid() == TRUE) 
						{
							oTask->SetStartDate(mp_oControl->GetMathLib()->RoundDate(mp_oControl->GetCurrentViewObject()->GetClientArea()->GetGrid()->GetInterval(), mp_oControl->GetCurrentViewObject()->GetClientArea()->GetGrid()->GetFactor(), oTask->GetStartDate()));
							oTask->SetEndDate(mp_oControl->GetMathLib()->RoundDate(mp_oControl->GetCurrentViewObject()->GetClientArea()->GetGrid()->GetInterval(), mp_oControl->GetCurrentViewObject()->GetClientArea()->GetGrid()->GetFactor(), oTask->GetEndDate()));
						}
                    }
                    else if (oTask->GetTaskType() == TT_DURATION)
                    {
                        oTask->SetDurationFactor(mp_oControl->GetMathLib()->CalculateDuration(s_tskSTR.dtInitialStartDate, s_tskSTR.dtFinalEndDate, oTask->GetDurationInterval()));
                    }

					if (oTask->GetEndDate() != s_tskSTR.dtInitialEndDate) 
					{
						mp_oControl->FireCompleteObjectChange();
						if (mp_oControl->GetObjectStateChangedEventArgs()->GetCancel() == TRUE)
						{
							oTask->SetEndDate(s_tskSTR.dtInitialEndDate);
						}
					}
				}
			}
		}
		if (mp_oControl->GetEnforcePredecessors() == TRUE)
		{
			mp_oControl->CheckPredecessors();
		}
		mp_oControl->Redraw();
		mp_SetCursor(CT_NORMAL);
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

void clsMouseKeyboardEvents::mp_EO_TASKSELECTION(LONG yMouseKeyBoardEvent,LONG X,LONG Y)
{

	clsTask* oTask = NULL;
	switch (yMouseKeyBoardEvent) 
	{
	case MouseHover:
		//ToolTipText(EO_TASKSELECTION, mp_oControl->GetMathLib()->GetTaskIndexByPosition(X, Y), X, Y, Nothing, Nothing, "")
		mp_SetCursor(CT_NORMAL);
		break;
	case MouseDown:
		s_tskSEL.lTaskIndex = mp_oControl->GetMathLib()->GetTaskIndexByPosition(X, Y);
		break;
	case MouseMove:
		break;
	case MouseUp:
		mp_oControl->SetSelectedTaskIndex(s_tskSEL.lTaskIndex);
		oTask = (clsTask*)mp_oControl->GetTasks()->mp_oCollection->m_oReturnArrayElement(mp_oControl->GetSelectedTaskIndex());
		if (mp_oControl->GetCurrentViewObject()->GetClientArea()->GetGrid()->GetSnapToGrid() == TRUE && mp_oControl->GetCurrentViewObject()->GetClientArea()->GetGrid()->GetSnapToGridOnSelection() == TRUE) 
		{
			oTask = (clsTask*)mp_oControl->GetTasks()->mp_oCollection->m_oReturnArrayElement(mp_oControl->GetSelectedTaskIndex());
			oTask->SetStartDate(mp_oControl->GetMathLib()->RoundDate(mp_oControl->GetCurrentViewObject()->GetClientArea()->GetGrid()->GetInterval(), mp_oControl->GetCurrentViewObject()->GetClientArea()->GetGrid()->GetFactor(), oTask->GetStartDate()));
			oTask->SetEndDate(mp_oControl->GetMathLib()->RoundDate(mp_oControl->GetCurrentViewObject()->GetClientArea()->GetGrid()->GetInterval(), mp_oControl->GetCurrentViewObject()->GetClientArea()->GetGrid()->GetFactor(), oTask->GetEndDate()));
		}

		mp_oControl->GetObjectSelectedEventArgs()->Clear();
		mp_oControl->GetObjectSelectedEventArgs()->SetEventTarget((LONG)EVT_TASK);
		mp_oControl->GetObjectSelectedEventArgs()->SetObjectIndex(mp_oControl->GetSelectedTaskIndex());
		mp_oControl->FireObjectSelected();
		mp_oControl->Redraw();
		mp_SetCursor(CT_NORMAL);
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

void clsMouseKeyboardEvents::mp_EO_TASKADDITION(LONG yMouseKeyBoardEvent,LONG X,LONG Y)
{

	clsRow* oRow = NULL;
	if (mp_oControl->GetAllowAdd() == FALSE) 
	{
		mp_SetCursor(CT_NORMAL);
		return;
	}
	s_tskADD.lRowIndex = mp_oControl->GetMathLib()->GetRowIndexByPosition(Y);
	if (s_tskADD.lRowIndex <= 0) 
	{
		return;
	}
	oRow = (clsRow*)mp_oControl->GetRows()->mp_oCollection->m_oReturnArrayElement(s_tskADD.lRowIndex);
	if (oRow->GetContainer() == FALSE) 
	{
		return;
	}
	switch (yMouseKeyBoardEvent) 
	{
	case MouseHover:
		mp_SetCursor(CT_CLIENTAREA);
		break;
	case MouseDown:
		s_tskADD.bCancel = FALSE;
		X = mp_fSnapX(X);
		if ((oRow->GetContainer() == FALSE) || (mp_oControl->GetMathLib()->DetectConflict(mp_oControl->GetMathLib()->GetDateFromXCoordinate(X), mp_oControl->GetMathLib()->GetDateFromXCoordinate(X), oRow->GetKey(), 0, mp_oControl->GetCurrentLayer()) == TRUE && mp_oControl->GetCurrentViewObject()->GetClientArea()->GetDetectConflicts() == TRUE)) 
		{
			s_tskADD.bCancel = TRUE;
		}
		else 
		{
			s_tskADD.dtStartDate = mp_oControl->GetMathLib()->GetDateFromXCoordinate(X);
		}

		break;
	case MouseMove:
		if (s_tskADD.bCancel == FALSE) 
		{
			X = mp_fSnapX(X);
			mp_oControl->clsG->mp_EraseReversibleFrames();
			if ((mp_oControl->GetMathLib()->DetectConflict(s_tskADD.dtStartDate, mp_oControl->GetMathLib()->GetDateFromXCoordinate(X), oRow->GetKey(), 0, mp_oControl->GetCurrentLayer()) == TRUE || oRow->GetContainer() == FALSE) && mp_oControl->GetCurrentViewObject()->GetClientArea()->GetDetectConflicts() == TRUE) 
			{
				mp_SetCursor(CT_NODROP);
				s_tskADD.bInConflict = TRUE;
			}
			else 
			{
				s_tskADD.bInConflict = FALSE;
				mp_SetCursor(CT_CLIENTAREA);
				s_tskADD.dtEndDate = mp_oControl->GetMathLib()->GetDateFromXCoordinate(X);
				mp_DynamicTimeLineMove(X);
				mp_oControl->GetToolTipEventArgs()->Clear();
				mp_oControl->GetToolTipEventArgs()->SetRowIndex(s_tskADD.lRowIndex);
				mp_oControl->GetToolTipEventArgs()->SetStartDate(s_tskADD.dtStartDate);
				mp_oControl->GetToolTipEventArgs()->SetEndDate(s_tskADD.dtEndDate);
				mp_oControl->GetToolTipEventArgs()->SetX(X);
				mp_oControl->GetToolTipEventArgs()->SetY(Y);
				mp_oControl->FireToolTipOnMouseMove(mp_yOperation);
				mp_DrawMovingReversibleFrame(mp_oControl->GetMathLib()->GetXCoordinateFromDate(s_tskADD.dtStartDate), oRow->GetTop(), X, oRow->GetBottom(), FCT_ADD);
			}
		}

		break;
	case MouseUp:
        mp_oControl->clsG->mp_EraseReversibleFrames();
        if (s_tskADD.bCancel == FALSE && s_tskADD.bInConflict == FALSE) 
        {
            X = mp_fSnapX(X);
            s_tskADD.dtEndDate = mp_oControl->GetMathLib()->GetDateFromXCoordinate(X);
            if (s_tskADD.dtEndDate == s_tskADD.dtStartDate)
            {
                if (mp_oControl->GetAddMode() == AT_BOTH || mp_oControl->GetAddMode() == AT_MILESTONEADD)
                {
                    mp_oControl->GetTasks()->Add("", oRow->GetKey(), s_tskADD.dtEndDate, s_tskADD.dtStartDate, _T(""), _T("DS_TASK"), mp_oControl->GetCurrentLayer());
                }
                else if (mp_oControl->GetAddMode() == AT_DURATION_BOTH || mp_oControl->GetAddMode() == AT_DURATION_MILESTONEADD)
                {
                    COleDateTime dtStartDate = s_tskADD.dtStartDate;
                    COleDateTime dtTLStartDate = mp_oControl->GetCurrentViewObject()->GetTimeLine()->GetStartDate();
                    COleDateTime dtTLEndDate = mp_oControl->GetCurrentViewObject()->GetTimeLine()->GetEndDate();
                    std::vector <S_TIMEBLOCK> aTimeBlocks;
                    mp_oControl->GetMathLib()->mp_GetTimeBlocks(aTimeBlocks, dtTLStartDate, dtTLEndDate);
                    mp_oControl->GetMathLib()->mp_ValidateStartDate(aTimeBlocks, dtStartDate);
                    mp_oControl->GetTasks()->DAdd(oRow->GetKey(), dtStartDate, mp_oControl->GetAddDurationInterval(), 0, _T(""), _T(""), _T("DS_TASK"), mp_oControl->GetCurrentLayer());
                }
				if (mp_oControl->GetAddMode() == AT_DURATION_BOTH || mp_oControl->GetAddMode() == AT_BOTH || mp_oControl->GetAddMode() == AT_MILESTONEADD || mp_oControl->GetAddMode() == AT_DURATION_MILESTONEADD)
				{
					mp_oControl->SetSelectedTaskIndex(mp_oControl->GetTasks()->GetCount());
					mp_oControl->GetObjectAddedEventArgs()->Clear();
					mp_oControl->GetObjectAddedEventArgs()->SetTaskIndex(mp_oControl->GetTasks()->GetCount());
					mp_oControl->GetObjectAddedEventArgs()->SetEventTarget((LONG)EVT_MILESTONE);
					mp_oControl->FireObjectAdded();
				}
            }
            else
            {
                if (mp_oControl->GetAddMode() == AT_BOTH || mp_oControl->GetAddMode() == AT_TASKADD)
                {
                    if (s_tskADD.dtEndDate < s_tskADD.dtStartDate)
                    {
                        mp_oControl->GetTasks()->Add(_T(""), oRow->GetKey(), s_tskADD.dtEndDate, s_tskADD.dtStartDate, _T(""), _T("DS_TASK"), mp_oControl->GetCurrentLayer());
                    }
                    else
                    {
                        mp_oControl->GetTasks()->Add(_T(""), oRow->GetKey(), s_tskADD.dtStartDate, s_tskADD.dtEndDate, _T(""), _T("DS_TASK"), mp_oControl->GetCurrentLayer());
                    }
                    mp_oControl->SetSelectedTaskIndex(mp_oControl->GetTasks()->GetCount());
                    mp_oControl->GetObjectAddedEventArgs()->Clear();
                    mp_oControl->GetObjectAddedEventArgs()->SetTaskIndex(mp_oControl->GetTasks()->GetCount());
                    mp_oControl->GetObjectAddedEventArgs()->SetEventTarget((LONG)EVT_TASK);
                    mp_oControl->FireObjectAdded();
                }
                else if (mp_oControl->GetAddMode() == AT_DURATION_BOTH || mp_oControl->GetAddMode() == AT_DURATION_TASKADD)
                {
                    int lDuration = 0;
                    COleDateTime dtStartDate;
                    COleDateTime dtEndDate;
                    if (s_tskADD.dtEndDate > s_tskADD.dtStartDate)
                    {
                        dtStartDate = s_tskADD.dtStartDate;
                        dtEndDate = s_tskADD.dtEndDate;
                    }
                    else
                    {
                        dtStartDate = s_tskADD.dtEndDate;
                        dtEndDate = s_tskADD.dtStartDate;
                    }
                    lDuration = mp_oControl->GetMathLib()->CalculateDuration(dtStartDate, dtEndDate, mp_oControl->GetAddDurationInterval());
                    mp_oControl->GetTasks()->DAdd(oRow->GetKey(), dtStartDate, mp_oControl->GetAddDurationInterval(), lDuration, _T(""), _T(""), _T("DS_TASK"), mp_oControl->GetCurrentLayer());
                    mp_oControl->SetSelectedTaskIndex(mp_oControl->GetTasks()->GetCount());
                    mp_oControl->GetObjectAddedEventArgs()->Clear();
                    mp_oControl->GetObjectAddedEventArgs()->SetTaskIndex(mp_oControl->GetTasks()->GetCount());
                    mp_oControl->GetObjectAddedEventArgs()->SetEventTarget((LONG)EVT_TASK);
                    mp_oControl->FireObjectAdded();
                }
            }
        }
        mp_oControl->Redraw();
        mp_SetCursor(CT_NORMAL);
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

void clsMouseKeyboardEvents::mp_EO_PERCENTAGESELECTION(LONG yMouseKeyBoardEvent,LONG X,LONG Y)
{

	clsPercentage* oPercentage = NULL;
	switch (yMouseKeyBoardEvent) 
	{
	case MouseHover:
		mp_SetCursor(CT_NORMAL);
		break;
	case MouseDown:
		s_perSEL.lPercentageIndex = mp_oControl->GetMathLib()->GetPercentageIndexByPosition(X, Y);
		break;
	case MouseMove:
		break;
	case MouseUp:
		if (s_perSEL.lPercentageIndex >= 1)
		{
			mp_oControl->SetSelectedPercentageIndex(s_perSEL.lPercentageIndex);
			oPercentage = (clsPercentage*)mp_oControl->GetPercentages()->mp_oCollection->m_oReturnArrayElement(mp_oControl->GetSelectedPercentageIndex());
			mp_oControl->GetObjectSelectedEventArgs()->Clear();
			mp_oControl->GetObjectSelectedEventArgs()->SetEventTarget((LONG)EVT_PERCENTAGE);
			mp_oControl->GetObjectSelectedEventArgs()->SetObjectIndex(mp_oControl->GetSelectedPercentageIndex());
		}
		else
		{
			mp_oControl->SetSelectedPercentageIndex(0);
			mp_oControl->GetObjectSelectedEventArgs()->Clear();
			mp_oControl->GetObjectSelectedEventArgs()->SetEventTarget((LONG)EVT_TASK);
			mp_oControl->GetObjectSelectedEventArgs()->SetObjectIndex(mp_oControl->GetSelectedTaskIndex());
		}
		mp_oControl->FireObjectSelected();
		mp_oControl->Redraw();
		mp_SetCursor(CT_NORMAL);
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

void clsMouseKeyboardEvents::mp_EO_PERCENTAGESIZING(LONG yMouseKeyBoardEvent,LONG X,LONG Y)
{

	clsPercentage* oPercentage = NULL;
	clsTask* oTask = NULL;
	if (mp_oControl->GetAllowEdit() == FALSE) 
	{
		return;
	}
	oPercentage = (clsPercentage*)mp_oControl->GetPercentages()->Item(CStr(mp_oControl->GetSelectedPercentageIndex()));
	if (oPercentage->GetAllowSize() == FALSE) 
	{
		return;
	}
	switch (yMouseKeyBoardEvent) 
	{
	case MouseHover:
		mp_SetCursor(CT_PERCENTAGE);
		break;
	case MouseDown:
		s_perSZ.bMouseMove = FALSE;
		oTask = (clsTask*)mp_oControl->GetTasks()->Item(oPercentage->GetTaskKey());
		mp_oControl->GetObjectStateChangedEventArgs()->Clear();
		mp_oControl->GetObjectStateChangedEventArgs()->SetEventTarget((LONG)EVT_PERCENTAGE);
		mp_oControl->GetObjectStateChangedEventArgs()->SetIndex(mp_oControl->GetSelectedPercentageIndex());
		mp_oControl->GetObjectStateChangedEventArgs()->SetCancel(FALSE);
		mp_oControl->GetObjectStateChangedEventArgs()->SetChangeType(CT_SIZE);
		mp_oControl->FireBeginObjectChange();
		if (mp_oControl->GetObjectStateChangedEventArgs()->GetCancel() == FALSE) 
		{
			s_perSZ.lTaskIndex = oTask->GetIndex();
		}
		break;
	case MouseMove:
		s_perSZ.bMouseMove = TRUE;
		if (mp_oControl->GetObjectStateChangedEventArgs()->GetCancel() == FALSE) 
		{
			mp_oControl->clsG->mp_EraseReversibleFrames();
			oTask = (clsTask*)mp_oControl->GetTasks()->Item(oPercentage->GetTaskKey());
			mp_oControl->FireObjectChange();
			if (mp_oControl->GetObjectStateChangedEventArgs()->GetCancel() == FALSE) 
			{
				mp_DynamicTimeLineMove(X);
				s_perSZ.lTaskStart = mp_oControl->GetMathLib()->GetXCoordinateFromDate(oTask->GetStartDate());
				s_perSZ.lTaskEnd = mp_oControl->GetMathLib()->GetXCoordinateFromDate(oTask->GetEndDate());
				if (X < s_perSZ.lTaskStart) 
				{
					X = s_perSZ.lTaskStart;
				}
				if (X > s_perSZ.lTaskEnd) 
				{
					X = s_perSZ.lTaskEnd;
				}
				float fPercent = 0;
				fPercent = mp_oControl->GetMathLib()->PercentageComplete(s_perSZ.lTaskStart, s_perSZ.lTaskEnd, X);
				fPercent = (FLOAT)CInt32(fPercent * (FLOAT)100);
				mp_oControl->GetToolTipEventArgs()->Clear();
				mp_oControl->GetToolTipEventArgs()->SetPercentageIndex(mp_oControl->GetSelectedPercentageIndex());
				mp_oControl->GetToolTipEventArgs()->SetXStart(s_perSZ.lTaskStart);
				mp_oControl->GetToolTipEventArgs()->SetXEnd(s_perSZ.lTaskEnd);
				mp_oControl->GetToolTipEventArgs()->SetTaskIndex(oTask->GetIndex());
				mp_oControl->GetToolTipEventArgs()->SetRowIndex(mp_oControl->GetRows()->Item(oTask->GetRowKey())->GetIndex());
				mp_oControl->GetToolTipEventArgs()->SetStartDate(oTask->GetStartDate());
				mp_oControl->GetToolTipEventArgs()->SetEndDate(oTask->GetEndDate());
				mp_oControl->GetToolTipEventArgs()->SetX(X);
				mp_oControl->GetToolTipEventArgs()->SetY(Y);
				mp_oControl->FireToolTipOnMouseMove(mp_yOperation);
				mp_DrawMovingReversibleFrame(s_perSZ.lTaskStart, oPercentage->GetTop(), X, oPercentage->GetBottom(), FCT_KEEPLEFTRIGHTBOUNDS);
				s_perSZ.lX = X;
			}
		}

		break;
	case MouseUp:
		if (s_perSZ.bMouseMove == FALSE) 
		{
			return;
		}

		if (mp_oControl->GetObjectStateChangedEventArgs()->GetCancel() == FALSE) 
		{
			mp_oControl->clsG->mp_EraseReversibleFrames();
			mp_oControl->FireEndObjectChange();
			if (mp_oControl->GetObjectStateChangedEventArgs()->GetCancel() == FALSE) 
			{
				s_perSZ.lTaskEnd = s_perSZ.lTaskEnd - s_perSZ.lTaskStart;
				s_perSZ.lX = s_perSZ.lX - s_perSZ.lTaskStart;
				if (s_perSZ.lX == 0) 
				{
					oPercentage->SetPercent(0);
				}
				else if (s_perSZ.lX == s_perSZ.lTaskEnd) 
				{
					oPercentage->SetPercent(1);
				}
				else 
				{
					oPercentage->SetPercent((FLOAT)s_perSZ.lX / (FLOAT)s_perSZ.lTaskEnd);
				}
				mp_oControl->FireCompleteObjectChange();
			}
		}

		mp_oControl->Redraw();
		mp_SetCursor(CT_NORMAL);
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

void clsMouseKeyboardEvents::mp_EO_PREDECESSORSELECTION(LONG yMouseKeyBoardEvent, LONG X, LONG Y)
{
	switch (yMouseKeyBoardEvent) 
	{
	case MouseHover:
		mp_SetCursor(CT_NORMAL);
		break;
	case MouseDown:
		s_preSEL.lPredecessorIndex = mp_oControl->GetMathLib()->GetPredecessorIndexByPosition(X, Y);
		break;
	case MouseMove:
		break;
	case MouseUp:
		mp_oControl->SetSelectedPredecessorIndex(s_preSEL.lPredecessorIndex);
		mp_oControl->GetObjectSelectedEventArgs()->Clear();
		mp_oControl->GetObjectSelectedEventArgs()->SetEventTarget(EVT_PREDECESSOR);
		mp_oControl->GetObjectSelectedEventArgs()->SetObjectIndex(mp_oControl->GetSelectedPredecessorIndex());
		mp_oControl->FireObjectSelected();
		mp_oControl->Redraw();
		mp_SetCursor(CT_NORMAL);
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

void clsMouseKeyboardEvents::mp_EO_PREDECESSORADDITION(LONG yMouseKeyBoardEvent,LONG X,LONG Y)
{


	clsTask* oTask = NULL;
	clsTask* oPredecessor = NULL;
	CString sType = "";
	E_CONSTRAINTTYPE mp_yType = PCT_END_TO_START;
	if (mp_oControl->GetAllowPredecessorAdd() == FALSE) 
	{
		return;
	}
	switch (yMouseKeyBoardEvent) 
	{
	case MouseHover:
		mp_SetCursor(CT_PREDECESSOR);
		break;
	case MouseDown:
		s_preADD.bCancel = FALSE;
		s_preADD.lXStart = X;
		s_preADD.lYStart = Y;
		s_preADD.lPredecessorIndex = mp_oControl->GetMathLib()->GetTaskIndexByPosition(X, Y);
		oPredecessor = mp_oControl->GetTasks()->Item(CStr(s_preADD.lPredecessorIndex));
		if ((oPredecessor->GetKey() == "" || oPredecessor->GetOutgoingPredecessors() == FALSE)) 
		{
			s_preADD.bCancel = TRUE;
			return;
		}

		s_preADD.sPredecessorKey = oPredecessor->GetKey();
		if ((X <= oPredecessor->GetLeft() + ((oPredecessor->GetRight() - oPredecessor->GetLeft()) / 2))) 
		{
			s_preADD.sPredecessorPosition = "S";
		}
		else 
		{
			s_preADD.sPredecessorPosition = "E";
		}

		break;
	case MouseMove:
		if (s_preADD.bCancel == FALSE) 
		{
			mp_oControl->clsG->mp_EraseReversibleFrames();
			mp_oControl->clsG->mp_EraseReversibleLines();
			//mp_DynamicRowMove(Y)
			//mp_DynamicTimeLineMove(X)
			mp_oControl->clsG->mp_DrawReversibleLine(s_preADD.lXStart, s_preADD.lYStart, X, Y);
			s_preADD.lTaskIndex = mp_oControl->GetMathLib()->GetTaskIndexByPosition(X, Y);
			if (s_preADD.lTaskIndex > 0) 
			{
				oTask = (clsTask*)mp_oControl->GetTasks()->Item(CStr(s_preADD.lTaskIndex));
				if (oTask->GetIncomingPredecessors() == FALSE) 
				{
					mp_SetCursor(CT_NODROP);
					s_preADD.bCantAccept = TRUE;
				}
				else 
				{
					s_preADD.bCantAccept = FALSE;
					mp_SetCursor(CT_PREDECESSOR);
					if ((X <= oTask->GetLeft() + ((oTask->GetRight() - oTask->GetLeft()) / 2))) 
					{
						s_preADD.sTaskPosition = "S";
					}
					else 
					{
						s_preADD.sTaskPosition = "E";
					}
					mp_oControl->GetToolTipEventArgs()->Clear();
					mp_oControl->GetToolTipEventArgs()->SetPredecessorPosition(s_preADD.sPredecessorPosition);
					mp_oControl->GetToolTipEventArgs()->SetTaskPosition(s_preADD.sTaskPosition);
					mp_oControl->clsG->mp_EraseReversibleLines();
					mp_oControl->FireToolTipOnMouseMove(mp_yOperation);
					mp_oControl->clsG->mp_DrawReversibleLine(s_preADD.lXStart, s_preADD.lYStart, X, Y);
					mp_DrawMovingReversibleFrame(oTask->GetLeft(), oTask->GetTop(), oTask->GetRight(), oTask->GetBottom(), FCT_KEEPLEFTRIGHTBOUNDS);
				}
			}
		}

		break;
	case MouseUp:

		mp_oControl->clsG->mp_EraseReversibleFrames();
		mp_oControl->clsG->mp_EraseReversibleLines();
		if (s_preADD.bCancel == FALSE && s_preADD.bCantAccept == FALSE) 
		{
			if (s_preADD.lTaskIndex > 0) 
			{
				oTask = (clsTask*)mp_oControl->GetTasks()->Item(CStr(s_preADD.lTaskIndex));
				if ((X <= oTask->GetLeft() + ((oTask->GetRight() - oTask->GetLeft()) / 2))) 
				{
					s_preADD.sTaskPosition = "S";
				}
				else 
				{
					s_preADD.sTaskPosition = "E";
				}
				sType = s_preADD.sPredecessorPosition + s_preADD.sTaskPosition;
				if (sType == "EE") 
				{
					mp_yType = PCT_END_TO_END;
					mp_oControl->GetPredecessors()->Add(oTask->GetKey(), s_preADD.sPredecessorKey, PCT_END_TO_END, "", "DS_PREDECESSOR");
				}
				else if (sType == "SS") 
				{
					mp_yType = PCT_START_TO_START;
					mp_oControl->GetPredecessors()->Add(oTask->GetKey(), s_preADD.sPredecessorKey, PCT_START_TO_START, "", "DS_PREDECESSOR");
				}
				else if (sType == "ES") 
				{
					mp_yType = PCT_END_TO_START;
					mp_oControl->GetPredecessors()->Add(oTask->GetKey(), s_preADD.sPredecessorKey, PCT_END_TO_START, "", "DS_PREDECESSOR");
				}
				else if (sType == "SE") 
				{
					mp_yType = PCT_START_TO_END;
					mp_oControl->GetPredecessors()->Add(oTask->GetKey(), s_preADD.sPredecessorKey, PCT_START_TO_END, "", "DS_PREDECESSOR");
				}
				if (mp_oControl->GetEnforcePredecessors() == TRUE)
				{
					mp_oControl->CheckPredecessors();
				}
				oPredecessor = mp_oControl->GetTasks()->Item(s_preADD.sPredecessorKey);
				mp_oControl->GetObjectAddedEventArgs()->Clear();
				mp_oControl->GetObjectAddedEventArgs()->SetTaskIndex(oTask->GetIndex());
				mp_oControl->GetObjectAddedEventArgs()->SetTaskKey(oTask->GetKey());
				mp_oControl->GetObjectAddedEventArgs()->SetPredecessorTaskIndex(oPredecessor->GetIndex());
				mp_oControl->GetObjectAddedEventArgs()->SetPredecessorTaskKey(oPredecessor->GetKey());
				mp_oControl->GetObjectAddedEventArgs()->SetPredecessorObjectIndex(mp_oControl->GetPredecessors()->GetCount());
				mp_oControl->GetObjectAddedEventArgs()->SetPredecessorType(mp_yType);
				mp_oControl->GetObjectAddedEventArgs()->SetEventTarget((LONG)EVT_PREDECESSOR);
				mp_oControl->FireObjectAdded();
			}
		}

		mp_oControl->Redraw();
		mp_SetCursor(CT_NORMAL);
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

BOOL clsMouseKeyboardEvents::mp_bOverSplitter(LONG X,LONG Y)
{
    if (mp_oControl->GetSplitter()->GetWidth() == 0)
    {
        return FALSE;
    }
	if (X >= (mp_oControl->GetSplitter()->GetRight() - mp_oControl->GetSplitter()->GetWidth()) && X <= mp_oControl->GetSplitter()->GetRight() && Y < mp_oControl->clsG->Height()) 
	{
		return TRUE;
	}
	else 
	{
		return FALSE;
	}
}

BOOL clsMouseKeyboardEvents::mp_bOverEmptySpace(LONG Y)
{
	if (Y > mp_oControl->GetRows()->GetTopOffset()) 
	{
		return TRUE;
	}
	else 
	{
		return FALSE;
	}
}

BOOL clsMouseKeyboardEvents::mp_bOverTimeLine(LONG X,LONG Y)
{
	if (X >= mp_oControl->GetCurrentViewObject()->GetTimeLine()->Getf_lStart() && X <= mp_oControl->GetCurrentViewObject()->GetTimeLine()->Getf_lEnd() && Y <= mp_oControl->GetCurrentViewObject()->GetTimeLine()->GetBottom() && Y >= mp_oControl->Getmt_TopMargin()) 
	{
		return TRUE;
	}
	else 
	{
		return FALSE;
	}
}

BOOL clsMouseKeyboardEvents::mp_bOverSelectedColumn(LONG X,LONG Y)
{
	clsColumn* oColumn = NULL;
	if (mp_oControl->GetSelectedColumnIndex() < 1 || mp_oControl->GetColumns()->GetCount() == 0) 
	{
		return FALSE;
	}
	oColumn = (clsColumn*)mp_oControl->GetColumns()->mp_oCollection->m_oReturnArrayElement(mp_oControl->GetSelectedColumnIndex());
	if (X >= oColumn->GetLeft() && X <= oColumn->GetRight() && Y >= oColumn->GetTop() && Y <= oColumn->GetBottom()) 
	{
		return TRUE;
	}
	else 
	{
		return FALSE;
	}
}

BOOL clsMouseKeyboardEvents::mp_bOverColumn(LONG X,LONG Y)
{
	clsColumn* oColumn = NULL;
	int lIndex = 0;
	if (!(X <= mp_oControl->GetSplitter()->GetLeft() && Y <= mp_oControl->GetCurrentViewObject()->GetTimeLine()->GetBottom())) 
	{
		return FALSE;
	}
	for (lIndex = 1; lIndex <= mp_oControl->GetColumns()->GetCount(); lIndex++) 
	{
		oColumn = (clsColumn*)mp_oControl->GetColumns()->mp_oCollection->m_oReturnArrayElement(lIndex);
		if (oColumn->GetVisible() == TRUE) 
		{
			if (X >= oColumn->GetLeft() && X <= oColumn->GetRight() && Y >= oColumn->GetTop() && Y <= oColumn->GetBottom()) 
			{
				return TRUE;
			}
		}
	}
	return FALSE;
}

BOOL clsMouseKeyboardEvents::mp_bOverSelectedRow(LONG X,LONG Y)
{
	clsRow* oRow = NULL;
	if (mp_oControl->GetSelectedRowIndex() < 1 || mp_oControl->GetRows()->GetCount() == 0) 
	{
		return FALSE;
	}
	oRow = (clsRow*)mp_oControl->GetRows()->mp_oCollection->m_oReturnArrayElement(mp_oControl->GetSelectedRowIndex());
	if (oRow->GetMergeCells() == TRUE) 
	{
		if (X >= oRow->GetLeft() && X <= oRow->GetRight() && Y >= oRow->GetTop() && Y <= oRow->GetBottom()) 
		{
			return TRUE;
		}
		else 
		{
			return FALSE;
		}
	}
	else 
	{
		if (X >= oRow->GetLeft() && X <= oRow->GetRight() && Y >= oRow->GetTop() && Y <= oRow->GetBottom()) 
		{
			if (mp_oControl->GetSelectedCellIndex() == mp_oControl->GetMathLib()->GetCellIndexByPosition(X)) 
			{
				return TRUE;
			}
			else 
			{
				return FALSE;
			}
		}
		else 
		{
			return FALSE;
		}
	}
}

BOOL clsMouseKeyboardEvents::mp_bOverRow(LONG X,LONG Y)
{
	clsRow* oRow = NULL;
	int lIndex = 0;
	if (!(X <= mp_oControl->GetCurrentViewObject()->GetTimeLine()->Getf_lStart() && Y > mp_oControl->GetCurrentViewObject()->GetTimeLine()->GetBottom())) 
	{
		return FALSE;
	}
	for (lIndex = 1; lIndex <= mp_oControl->GetRows()->GetCount(); lIndex++) 
	{
		oRow = (clsRow*)mp_oControl->GetRows()->mp_oCollection->m_oReturnArrayElement(lIndex);
		if (oRow->GetVisible() == TRUE) 
		{
			if (X >= oRow->GetLeft() && X <= oRow->GetRight() && Y >= oRow->GetTop() && Y <= oRow->GetBottom()) 
			{
				return TRUE;
			}
		}
	}
	return FALSE;
}

BOOL clsMouseKeyboardEvents::mp_bOverSelectedTask(LONG X,LONG Y)
{
	clsTask* oSelectedTask = NULL;
	if (X < mp_oControl->GetCurrentViewObject()->GetTimeLine()->Getf_lStart()) 
	{
		return FALSE;
	}
	if (X > mp_oControl->GetCurrentViewObject()->GetTimeLine()->Getf_lEnd()) 
	{
		return FALSE;
	}
	if (mp_oControl->GetSelectedTaskIndex() < 1 || mp_oControl->GetTasks()->GetCount() == 0) 
	{
		return FALSE;
	}
	oSelectedTask = (clsTask*)mp_oControl->GetTasks()->mp_oCollection->m_oReturnArrayElement(mp_oControl->GetSelectedTaskIndex());
	if (X >= oSelectedTask->GetLeft() && X <= oSelectedTask->GetRight() && Y >= oSelectedTask->GetTop() && Y <= oSelectedTask->GetBottom() && mp_oControl->GetMathLib()->InCurrentLayer(oSelectedTask->GetLayerIndex())) 
	{
		return TRUE;
	}
	else 
	{
		return FALSE;
	}
}

BOOL clsMouseKeyboardEvents::mp_bOverSelectedPredecessor(LONG X,LONG Y)
{
    clsPredecessor* oSelectedPredecessor = NULL;
	if (X < mp_oControl->GetCurrentViewObject()->GetTimeLine()->Getf_lStart()) 
	{
		return FALSE;
	}
	if (X > mp_oControl->GetCurrentViewObject()->GetTimeLine()->Getf_lEnd()) 
	{
		return FALSE;
	}
	if (mp_oControl->GetSelectedPredecessorIndex() < 1 || mp_oControl->GetPredecessors()->GetCount() == 0) 
	{
		return FALSE;
	}
	oSelectedPredecessor = (clsPredecessor*)mp_oControl->GetPredecessors()->mp_oCollection->m_oReturnArrayElement(mp_oControl->GetSelectedPredecessorIndex());
    if (oSelectedPredecessor->HitTest(X, Y) == TRUE)
    {
        return TRUE;
    }
    else
    {
        return FALSE;
    }
}

BOOL clsMouseKeyboardEvents::mp_bOverSelectedPercentage(LONG X,LONG Y)
{
	clsPercentage* oSelectedPercentage = NULL;
	if (X < mp_oControl->GetCurrentViewObject()->GetTimeLine()->Getf_lStart()) 
	{
		return FALSE;
	}
	if (X > mp_oControl->GetCurrentViewObject()->GetTimeLine()->Getf_lEnd()) 
	{
		return FALSE;
	}
	if (mp_oControl->GetSelectedPercentageIndex() < 1 || mp_oControl->GetPercentages()->GetCount() == 0)
	{
		return FALSE;
	}
	oSelectedPercentage = (clsPercentage*)mp_oControl->GetPercentages()->mp_oCollection->m_oReturnArrayElement(mp_oControl->GetSelectedPercentageIndex());
	if (X >= oSelectedPercentage->GetLeft() && X <= oSelectedPercentage->GetRightSel() && Y >= oSelectedPercentage->GetTop() && Y <= oSelectedPercentage->GetBottom()) 
	{
		return TRUE;
	}
	else 
	{
		return FALSE;
	}
}

LONG clsMouseKeyboardEvents::mp_yTaskArea(LONG X,LONG Y)
{
	clsTask* oSelectedTask = NULL;
	oSelectedTask = (clsTask*)mp_oControl->GetTasks()->mp_oCollection->m_oReturnArrayElement(mp_oControl->GetSelectedTaskIndex());
	if (X >= oSelectedTask->GetLeft() && X <= oSelectedTask->GetRight() && Y >= oSelectedTask->GetTop() && Y <= oSelectedTask->GetBottom() && mp_oControl->GetMathLib()->InCurrentLayer(oSelectedTask->GetLayerIndex())) 
	{
		if (X >= oSelectedTask->GetLeft() && X <= oSelectedTask->GetLeft() + mp_oControl->GetCurrentViewObject()->GetClientArea()->GetTaskBorderSelectionOffset()) 
		{
			if (oSelectedTask->Getf_bLeftVisible() == TRUE) 
			{
				return EA_LEFT;
			}
			else 
			{
				return EA_CENTER;
			}
		}
		if (X >= oSelectedTask->GetRight() - mp_oControl->GetCurrentViewObject()->GetClientArea()->GetTaskBorderSelectionOffset() && X <= oSelectedTask->GetRight()) 
		{
			if (oSelectedTask->Getf_bRightVisible() == TRUE) 
			{
				return EA_RIGHT;
			}
			else 
			{
				return EA_CENTER;
			}
		}
		return EA_CENTER;
	}
	return EA_NONE;
}

LONG clsMouseKeyboardEvents::mp_yRowArea(LONG X,LONG Y)
{
	clsRow* oSelectedRow = NULL;
	oSelectedRow = (clsRow*)mp_oControl->GetRows()->mp_oCollection->m_oReturnArrayElement(mp_oControl->GetSelectedRowIndex());
	if (Y >= oSelectedRow->GetBottom() && Y <= oSelectedRow->GetBottom() + 3) 
	{
		return EA_BOTTOM;
	}
	else 
	{
		return EA_CENTER;
	}
}

LONG clsMouseKeyboardEvents::mp_yColumnArea(LONG X,LONG Y)
{
	clsColumn* oSelectedColumn = NULL;
	oSelectedColumn = (clsColumn*)mp_oControl->GetColumns()->mp_oCollection->m_oReturnArrayElement(mp_oControl->GetSelectedColumnIndex());
	if (X >= (oSelectedColumn->GetRight() - 3) && X <= oSelectedColumn->GetRight()) 
	{
		return EA_RIGHT;
	}
	else 
	{
		return EA_CENTER;
	}
}

void clsMouseKeyboardEvents::mp_DynamicColumnMove(LONG v_X)
{
	if (v_X < mp_oControl->Getmt_LeftMargin()) 
	{
		if (mp_oControl->GetHorizontalScrollBar()->GetValue() > 20) 
		{
			mp_oControl->GetHorizontalScrollBar()->SetValue(mp_oControl->GetHorizontalScrollBar()->GetValue() - 20);
		}
		else 
		{
			mp_oControl->GetHorizontalScrollBar()->SetValue(0);
		}
		mp_oControl->Redraw();
		return;
	}
	if (v_X > mp_oControl->GetSplitter()->GetLeft()) 
	{
		if (mp_oControl->GetHorizontalScrollBar()->GetValue() < (mp_oControl->GetHorizontalScrollBar()->GetMax() - 20)) 
		{
			mp_oControl->GetHorizontalScrollBar()->SetValue(mp_oControl->GetHorizontalScrollBar()->GetValue() + 20);
		}
		else 
		{
			mp_oControl->GetHorizontalScrollBar()->SetValue(mp_oControl->GetHorizontalScrollBar()->GetMax());
		}
		mp_oControl->Redraw();
		return;
	}
}

void clsMouseKeyboardEvents::mp_DynamicRowMove(LONG v_Y)
{
	if (v_Y < mp_oControl->GetCurrentViewObject()->GetTimeLine()->GetBottom()) 
	{
		if (mp_oControl->GetCurrentViewObject()->GetClientArea()->GetFirstVisibleRow() > 1) 
		{
			mp_oControl->GetCurrentViewObject()->GetClientArea()->SetFirstVisibleRow(mp_oControl->GetCurrentViewObject()->GetClientArea()->GetFirstVisibleRow() - 1);
			mp_oControl->GetVerticalScrollBar()->SetValue(mp_oControl->GetVerticalScrollBar()->GetValue() - 1);
			mp_oControl->Redraw();
			return;
		}
	}
	if (v_Y > mp_oControl->GetCurrentViewObject()->GetClientArea()->GetBottom()) 
	{
		if (mp_oControl->GetVerticalScrollBar()->GetValue() < mp_oControl->GetVerticalScrollBar()->GetMax()) 
		{
			mp_oControl->GetCurrentViewObject()->GetClientArea()->SetFirstVisibleRow(mp_oControl->GetCurrentViewObject()->GetClientArea()->GetFirstVisibleRow() + 1);
			mp_oControl->GetVerticalScrollBar()->SetValue(mp_oControl->GetVerticalScrollBar()->GetValue() + 1);
			mp_oControl->Redraw();
		}
	}
}

void clsMouseKeyboardEvents::mp_DynamicTimeLineMove(LONG v_X)
{
	if (v_X > mp_oControl->GetCurrentViewObject()->GetTimeLine()->Getf_lEnd()) 
	{
		if (mp_oControl->f_TimeLineScrollBar()->GetEnabled() == FALSE) 
		{
			mp_oControl->GetCurrentViewObject()->GetTimeLine()->Setf_StartDate(mp_oControl->GetMathLib()->DateTimeAdd(mp_oControl->GetCurrentViewObject()->Getf_ScrollInterval(), mp_oControl->GetCurrentViewObject()->GetFactor(), mp_oControl->GetCurrentViewObject()->GetTimeLine()->GetStartDate()));
		}
		else 
		{
			if (mp_oControl->f_TimeLineScrollBar()->GetValue() < mp_oControl->f_TimeLineScrollBar()->GetMax()) 
			{
				mp_oControl->f_TimeLineScrollBar()->SetValue(mp_oControl->f_TimeLineScrollBar()->GetValue() + 1);
			}
		}
		mp_oControl->Redraw();
	}
	if (v_X < mp_oControl->GetCurrentViewObject()->GetTimeLine()->Getf_lStart()) 
	{
		if (mp_oControl->f_TimeLineScrollBar()->GetEnabled() == FALSE) 
		{
			mp_oControl->GetCurrentViewObject()->GetTimeLine()->Setf_StartDate(mp_oControl->GetMathLib()->DateTimeAdd(mp_oControl->GetCurrentViewObject()->Getf_ScrollInterval(), -mp_oControl->GetCurrentViewObject()->GetFactor(), mp_oControl->GetCurrentViewObject()->GetTimeLine()->GetStartDate()));
		}
		else 
		{
			if (mp_oControl->f_TimeLineScrollBar()->GetValue() > 0) 
			{
				mp_oControl->f_TimeLineScrollBar()->SetValue(mp_oControl->f_TimeLineScrollBar()->GetValue() - 1);
			}
		}
		mp_oControl->Redraw();
	}
}

BOOL clsMouseKeyboardEvents::mp_bOverTask(LONG X,LONG Y)
{
	clsTask* oTask = NULL;
	int lIndex = 0;
	for (lIndex = mp_oControl->GetTasks()->GetCount(); lIndex >= 1; lIndex += -1) 
	{
		oTask = (clsTask*)mp_oControl->GetTasks()->mp_oCollection->m_oReturnArrayElement(lIndex);
		if (oTask->GetVisible() == TRUE && mp_oControl->GetMathLib()->InCurrentLayer(oTask->GetLayerIndex())) 
		{
			if (X >= oTask->GetLeft() && X <= oTask->GetRight() && Y >= oTask->GetTop() && Y <= oTask->GetBottom()) 
			{
				return TRUE;
			}
		}
	}
	return FALSE;
}

BOOL clsMouseKeyboardEvents::mp_bOverPredecessor(LONG X, LONG Y)
{
	clsPredecessor* oPredecessor = NULL;
	int lIndex = 0;
	for (lIndex = mp_oControl->GetPredecessors()->GetCount(); lIndex >= 1; lIndex += -1)
	{
		oPredecessor = (clsPredecessor*)mp_oControl->GetPredecessors()->mp_oCollection->m_oReturnArrayElement(lIndex);
		if (oPredecessor->GetVisible() == TRUE)
		{
			if (oPredecessor->HitTest(X, Y) == TRUE)
			{
				return TRUE;
			}
		}
	}
	return FALSE;
}

BOOL clsMouseKeyboardEvents::mp_bOverPercentage(LONG X,LONG Y)
{
	clsPercentage* oPercentage = NULL;
	int lIndex = 0;
	for (lIndex = mp_oControl->GetPercentages()->GetCount(); lIndex >= 1; lIndex += -1) 
	{
		oPercentage = (clsPercentage*)mp_oControl->GetPercentages()->mp_oCollection->m_oReturnArrayElement(lIndex);
		if (oPercentage->GetVisible() == TRUE) 
		{
			if (X >= oPercentage->GetLeft() && X <= oPercentage->GetRightSel() && Y >= oPercentage->GetTop() && Y <= oPercentage->GetBottom()) 
			{
				return TRUE;
			}
		}
	}
	return FALSE;
}

BOOL clsMouseKeyboardEvents::mp_bOverClientArea(LONG X,LONG Y)
{
	if (X >= mp_oControl->GetCurrentViewObject()->GetTimeLine()->Getf_lStart() && X <= mp_oControl->GetCurrentViewObject()->GetTimeLine()->Getf_lEnd() && Y >= mp_oControl->GetCurrentViewObject()->GetClientArea()->GetTop()) 
	{
		return TRUE;
	}
	else 
	{
		return FALSE;
	}
}

LONG clsMouseKeyboardEvents::mp_fSnapX(LONG X)
{
	COleDateTime dtDate;
	if (mp_oControl->GetCurrentViewObject()->GetClientArea()->GetGrid()->GetSnapToGrid() == FALSE) 
	{
		return X;
	}
	dtDate = mp_oControl->GetMathLib()->GetDateFromXCoordinate(X);
	dtDate = mp_oControl->GetMathLib()->RoundDate(mp_oControl->GetCurrentViewObject()->GetClientArea()->GetGrid()->GetInterval(), mp_oControl->GetCurrentViewObject()->GetClientArea()->GetGrid()->GetFactor(), dtDate);
	return mp_oControl->GetMathLib()->GetXCoordinateFromDate(dtDate);
}

void clsMouseKeyboardEvents::mp_DrawMovingReversibleFrame(LONG v_X1,LONG v_Y1,LONG v_X2,LONG v_Y2,LONG v_yFocusType)
{
	mp_oControl->clsG->Setf_FocusLeft(v_X1);
	mp_oControl->clsG->Setf_FocusTop(v_Y1);
	mp_oControl->clsG->Setf_FocusRight(v_X2);
	mp_oControl->clsG->Setf_FocusBottom(v_Y2);

	switch (v_yFocusType) 
	{
	case FCT_NORMAL:
		break;
	case FCT_KEEPLEFTRIGHTBOUNDS:
		if (mp_oControl->clsG->Getf_FocusLeft() < mp_oControl->GetCurrentViewObject()->GetTimeLine()->Getf_lStart()) 
		{
			mp_oControl->clsG->Setf_FocusLeft(mp_oControl->GetCurrentViewObject()->GetTimeLine()->Getf_lStart());
		}

		if (mp_oControl->clsG->Getf_FocusRight() > mp_oControl->GetCurrentViewObject()->GetTimeLine()->Getf_lEnd()) 
		{
			mp_oControl->clsG->Setf_FocusRight(mp_oControl->GetCurrentViewObject()->GetTimeLine()->Getf_lEnd());
		}

		break;
	case FCT_ADD:
		if (mp_oControl->clsG->Getf_FocusLeft() < mp_oControl->GetCurrentViewObject()->GetTimeLine()->Getf_lStart()) 
		{
			mp_oControl->clsG->Setf_FocusLeft(mp_oControl->GetCurrentViewObject()->GetTimeLine()->Getf_lStart());
		}

		if (mp_oControl->clsG->Getf_FocusRight() < mp_oControl->clsG->Getf_FocusLeft()) 
		{
			mp_oControl->clsG->Setf_FocusRight(mp_oControl->clsG->Getf_FocusLeft());
			mp_oControl->clsG->Setf_FocusLeft(v_X2);
		}


		break;
	case FCT_VERTICALSPLITTER:
		if (mp_oControl->clsG->Getf_FocusLeft() >= mp_oControl->GetSplitter()->GetRight()) 
		{
			mp_oControl->clsG->Setf_FocusBottom(mp_oControl->GetCurrentViewObject()->GetClientArea()->GetBottom());
		}
		else 
		{
			mp_oControl->clsG->Setf_FocusBottom(mp_oControl->Getmt_TableBottom());
		}

		break;
	}
	mp_oControl->clsG->mp_DrawReversibleFrameEx();
}

void clsMouseKeyboardEvents::OnMouseLeave(void)
{
	mp_oToolTip->SetVisible(FALSE);
}

BOOL clsMouseKeyboardEvents::mp_bCursorEditTextColumn(int X, int Y)
{
	clsColumn* oColumn;
	oColumn = mp_oControl->GetColumns()->Item(CStr(mp_oControl->GetSelectedColumnIndex()));
	if (oColumn->GetAllowTextEdit() == TRUE)
	{
		if (X >= oColumn->mp_lTextLeft && X <= oColumn->mp_lTextRight)
		{
			if (Y >= oColumn->mp_lTextTop && Y <= oColumn->mp_lTextBottom)
			{
				mp_SetCursor(CT_IBEAM);
				return TRUE;
			}
		}
	}
	mp_SetCursor(CT_NORMAL);
	return FALSE;
}

BOOL clsMouseKeyboardEvents::mp_bShowEditTextColumn(int X, int Y)
{
	clsColumn* oColumn;
	oColumn = mp_oControl->GetColumns()->Item(CStr(mp_oControl->GetSelectedColumnIndex()));
	if (oColumn->GetAllowTextEdit() == TRUE)
	{
		if (X >= oColumn->mp_lTextLeft && X <= oColumn->mp_lTextRight)
		{
			if (Y >= oColumn->mp_lTextTop && Y <= oColumn->mp_lTextBottom)
			{
				mp_oControl->mp_oTextBox->Initialize(mp_oControl->GetSelectedColumnIndex(), 0, TOT_COLUMN, X, Y);
				return TRUE;
			}
		}
	}
	return FALSE;
}

BOOL clsMouseKeyboardEvents::mp_bCursorEditTextRow(int X, int Y)
{
	clsRow* oRow;
	oRow = mp_oControl->GetRows()->Item(CStr(mp_oControl->GetSelectedRowIndex()));
	if (oRow->GetMergeCells() == TRUE)
	{
		if (oRow->GetAllowTextEdit() == TRUE)
		{
			if (X >= oRow->mp_lTextLeft && X <= oRow->mp_lTextRight)
			{
				if (Y >= oRow->mp_lTextTop && Y <= oRow->mp_lTextBottom)
				{
					mp_SetCursor(CT_IBEAM);
					return TRUE;
				}
			}
		}
	}
	else
	{
		clsCell* oCell;
		int lCellIndex = 0;
		clsColumn* oColumn;
		for (lCellIndex = 1; lCellIndex <= mp_oControl->GetColumns()->GetCount(); lCellIndex++)
		{
			oColumn = mp_oControl->GetColumns()->Item(CStr(lCellIndex));
			if (oColumn->GetVisible() == TRUE)
			{
				oCell = oRow->GetCells()->Item(CStr(lCellIndex));
				if (oCell->GetAllowTextEdit() == TRUE)
				{
					if (X >= oCell->mp_lTextLeft && X <= oCell->mp_lTextRight)
					{
						if (Y >= oCell->mp_lTextTop && Y <= oCell->mp_lTextBottom)
						{
							mp_SetCursor(CT_IBEAM);
							return TRUE;
						}
					}
				}
			}
		}
	}
	mp_SetCursor(CT_NORMAL);
	return FALSE;
}

BOOL clsMouseKeyboardEvents::mp_bShowEditTextRow(int X, int Y)
{
	clsRow* oRow;
	oRow = mp_oControl->GetRows()->Item(CStr(mp_oControl->GetSelectedRowIndex()));
	if (oRow->GetMergeCells() == TRUE)
	{
		if (oRow->GetAllowTextEdit() == TRUE)
		{
			if (X >= oRow->mp_lTextLeft && X <= oRow->mp_lTextRight)
			{
				if (Y >= oRow->mp_lTextTop && Y <= oRow->mp_lTextBottom)
				{
					mp_oControl->mp_oTextBox->Initialize(mp_oControl->GetSelectedRowIndex(), 0, TOT_ROW, X, Y);
					return TRUE;
				}
			}
		}
	}
	else
	{
		clsCell* oCell;
		int lCellIndex = 0;
		clsColumn* oColumn;
		for (lCellIndex = 1; lCellIndex <= mp_oControl->GetColumns()->GetCount(); lCellIndex++)
		{
			oColumn = mp_oControl->GetColumns()->Item(CStr(lCellIndex));
			if (oColumn->GetVisible() == TRUE)
			{
				oCell = oRow->GetCells()->Item(CStr(lCellIndex));
				if (oCell->GetAllowTextEdit() == TRUE)
				{
					if (X >= oCell->mp_lTextLeft && X <= oCell->mp_lTextRight)
					{
						if (Y >= oCell->mp_lTextTop && Y <= oCell->mp_lTextBottom)
						{
							mp_oControl->mp_oTextBox->Initialize(mp_oControl->GetSelectedRowIndex(), lCellIndex, TOT_CELL, X, Y);
							return TRUE;
						}
					}
				}
			}
		}
	}
	return FALSE;
}

BOOL clsMouseKeyboardEvents::mp_bCursorEditTextTask(int X, int Y)
{
	clsTask* oTask;
	if (mp_oControl->GetSelectedTaskIndex() <= 0)
	{
		mp_SetCursor(CT_NORMAL);
		return FALSE;
	}
	oTask = mp_oControl->GetTasks()->Item(CStr(mp_oControl->GetSelectedTaskIndex()));
	if (oTask->GetAllowTextEdit() == TRUE)
	{
		if (X >= oTask->mp_lTextLeft && X <= oTask->mp_lTextRight)
		{
			if (Y >= oTask->mp_lTextTop && Y <= oTask->mp_lTextBottom)
			{
				mp_SetCursor(CT_IBEAM);
				return TRUE;
			}
		}
	}
	mp_SetCursor(CT_NORMAL);
	return FALSE;
}

BOOL clsMouseKeyboardEvents::mp_bShowEditTextTask(int X, int Y)
{
	clsTask* oTask;
	if (mp_oControl->GetSelectedTaskIndex() <= 0)
	{
		return FALSE;
	}
	oTask = mp_oControl->GetTasks()->Item(CStr(mp_oControl->GetSelectedTaskIndex()));
	if (oTask->GetAllowTextEdit() == TRUE)
	{
		if (X >= oTask->mp_lTextLeft && X <= oTask->mp_lTextRight)
		{
			if (Y >= oTask->mp_lTextTop && Y <= oTask->mp_lTextBottom)
			{
				mp_oControl->mp_oTextBox->Initialize(mp_oControl->GetSelectedTaskIndex(), 0, TOT_TASK, X, Y);
				return TRUE;
			}
		}
	}
	return FALSE;
}

void clsMouseKeyboardEvents::mp_SetCursor(LONG v_iCursorType)
{
	switch (v_iCursorType) 
	{
	case CT_NORMAL:
		::SetCursor(mp_hDefault);
		break;
	case CT_SIZETASK:
		::SetCursor(mp_hSizeWE);
		break;
	case CT_MOVETASK:
		::SetCursor(mp_hSizeAll);
		break;
	case CT_MOVEMILESTONE:
		::SetCursor(mp_hSizeAll);
		break;
	case CT_CLIENTAREA:
		::SetCursor(mp_hCrossHair);
		break;
	case CT_MOVESPLITTER:
		::SetCursor(mp_hSizeWE);
		break;
	case CT_IBEAM:
		::SetCursor(mp_hIBeam);
		break;
	case CT_ROWHEIGHT:
		::SetCursor(mp_hHorizontalSplit);
		break;
	case CT_COLUMNWIDTH:
		::SetCursor(mp_hVerticalSplit);
		break;
	case CT_MOVEROW:
		::SetCursor(mp_hDragMove);
		break;
	case CT_MOVECOLUMN:
		::SetCursor(mp_hDragMove);
		break;
	case CT_SCROLLTIMELINE:
		::SetCursor(mp_hWait);
		break;
	case CT_PERCENTAGE:
		::SetCursor(mp_hPercentage);
		break;
	case CT_PREDECESSOR:
		::SetCursor(mp_hPredecessor);
		break;
	case CT_NODROP:
		::SetCursor(mp_hNoDrop);
		break;
	}
}