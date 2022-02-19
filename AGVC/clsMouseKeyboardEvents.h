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
#pragma once


class CActiveGanttVCCtl;
class clsToolTip;

class clsMouseKeyboardEvents  
{
public:
	CActiveGanttVCCtl* mp_oControl;

	clsMouseKeyboardEvents(CActiveGanttVCCtl* oControl);
	virtual ~clsMouseKeyboardEvents();

	LONG mp_yOperation;
	LONG mp_yOperationModifier;
	clsToolTip* mp_oToolTip;
	LONG mp_lCursorType;

protected:
	//Cursors:
	HCURSOR mp_hDefault;
	HCURSOR mp_hSizeWE;
	HCURSOR mp_hSizeAll;
	HCURSOR mp_hCrossHair;
	HCURSOR mp_hHorizontalSplit;
	HCURSOR mp_hVerticalSplit;
	HCURSOR mp_hDragMove;
	HCURSOR mp_hWait;
	HCURSOR mp_hPercentage;
	HCURSOR mp_hPredecessor;
	HCURSOR mp_hNoDrop;
	HCURSOR mp_hIBeam;

//Internal Methods
public:
	void KeyDown(void);
	void KeyUp(void);
	void KeyPress(void);
	void OnMouseClick(void);
	void OnMouseDblClick(void);
	void OnMouseMoveGeneral(LONG X,LONG Y);
	void OnMouseHover(LONG X,LONG Y);
	void OnMouseDown(LONG X,LONG Y,LONG Button);
	void OnMouseMove(LONG X,LONG Y);
	void OnMouseUp(LONG X,LONG Y);
	LONG CursorPosition(LONG X,LONG Y);
	void mp_EO_SPLITTERMOVEMENT(LONG yMouseKeyBoardEvent,LONG X,LONG Y);
	void mp_EO_TIMELINEMOVEMENT(LONG yMouseKeyBoardEvent,LONG X,LONG Y);
	void mp_EO_COLUMNMOVEMENT(LONG yMouseKeyBoardEvent,LONG X,LONG Y);
	void mp_EO_COLUMNSIZING(LONG yMouseKeyBoardEvent,LONG X,LONG Y);
	void mp_EO_COLUMNSELECTION(LONG yMouseKeyBoardEvent,LONG X,LONG Y);
	void mp_EO_ROWMOVEMENT(LONG yMouseKeyBoardEvent,LONG X,LONG Y);
	void mp_EO_ROWSIZING(LONG yMouseKeyBoardEvent,LONG X,LONG Y);
	void mp_EO_ROWSELECTION(LONG yMouseKeyBoardEvent,LONG X,LONG Y);
	void mp_EO_TASKMOVEMENT(LONG yMouseKeyBoardEvent,LONG X,LONG Y);
	void mp_EO_TASKSTRETCHLEFT(LONG yMouseKeyBoardEvent,LONG X,LONG Y);
	void mp_EO_TASKSTRETCHRIGHT(LONG yMouseKeyBoardEvent,LONG X,LONG Y);
	void mp_EO_TASKSELECTION(LONG yMouseKeyBoardEvent,LONG X,LONG Y);
	void mp_EO_TASKADDITION(LONG yMouseKeyBoardEvent,LONG X,LONG Y);
	void mp_EO_PERCENTAGESELECTION(LONG yMouseKeyBoardEvent,LONG X,LONG Y);
	void mp_EO_PERCENTAGESIZING(LONG yMouseKeyBoardEvent,LONG X,LONG Y);
	void mp_EO_PREDECESSORSELECTION(LONG yMouseKeyBoardEvent, LONG X, LONG Y);
	void mp_EO_PREDECESSORADDITION(LONG yMouseKeyBoardEvent,LONG X,LONG Y);
	BOOL mp_bOverSplitter(LONG X,LONG Y);
	BOOL mp_bOverEmptySpace(LONG Y);
	BOOL mp_bOverTimeLine(LONG X,LONG Y);
	BOOL mp_bOverSelectedColumn(LONG X,LONG Y);
	BOOL mp_bOverColumn(LONG X,LONG Y);
	BOOL mp_bOverSelectedRow(LONG X,LONG Y);
	BOOL mp_bOverRow(LONG X,LONG Y);
	BOOL mp_bOverSelectedTask(LONG X,LONG Y);
	BOOL mp_bOverSelectedPredecessor(LONG X,LONG Y);
	BOOL mp_bOverSelectedPercentage(LONG X,LONG Y);
	LONG mp_yTaskArea(LONG X,LONG Y);
	LONG mp_yRowArea(LONG X,LONG Y);
	LONG mp_yColumnArea(LONG X,LONG Y);
	void mp_DynamicColumnMove(LONG v_X);
	void mp_DynamicRowMove(LONG v_Y);
	void mp_DynamicTimeLineMove(LONG v_X);
	BOOL mp_bOverTask(LONG X,LONG Y);
	BOOL mp_bOverPredecessor(LONG X, LONG Y);
	BOOL mp_bOverPercentage(LONG X,LONG Y);
	BOOL mp_bOverClientArea(LONG X,LONG Y);
	LONG mp_fSnapX(LONG X);
	void mp_DrawMovingReversibleFrame(LONG v_X1,LONG v_Y1,LONG v_X2,LONG v_Y2,LONG v_yFocusType);
	void OnMouseLeave(void);

	BOOL mp_bCursorEditTextColumn(int X, int Y);
	BOOL mp_bShowEditTextColumn(int X, int Y);
	BOOL mp_bCursorEditTextRow(int X, int Y);
	BOOL mp_bShowEditTextRow(int X, int Y);
	BOOL mp_bCursorEditTextTask(int X, int Y);
	BOOL mp_bShowEditTextTask(int X, int Y);

	void mp_SetCursor(LONG v_iCursorType);

	struct S_TIMELINEMOVEMENT
	{
		LONG lX;
		void Clear()
		{
			lX = 0;
		}
	};

	struct S_COLUMNMOVEMENT
	{
		LONG lColumnIndex;
		LONG lDestinationColumnIndex;
		void Clear()
		{
			lColumnIndex = 0;
			lDestinationColumnIndex = 0;
		}
	};

	struct S_COLUMNSIZING
	{
		LONG lColumnIndex;
		void Clear()
		{
			lColumnIndex = 0;
		}
	};

	struct S_COLUMNSELECTION
	{
		LONG lColumnIndex;
		void Clear()
		{
			lColumnIndex = 0;
		}
	};

	struct S_ROWMOVEMENT
	{
		LONG lRowIndex;
		LONG lDestinationRowIndex;
		void Clear()
		{
			lRowIndex = 0;
			lDestinationRowIndex = 0;
		}
	};

	struct S_ROWSIZING
	{
		LONG lRowIndex;
		void Clear()
		{
			lRowIndex = 0;
		}
	};

	struct S_ROWSELECTION
	{
		LONG lRowIndex;
		LONG lCellIndex;
		void Clear()
		{
			lRowIndex = 0;
			lCellIndex = 0;
		}
	};

	struct S_TASKMOVEMENT
	{
		LONG lInitialRowIndex;
		CString sInitialRowKey;
		LONG lFinalRowIndex;
		COleDateTime dtInitialStartDate;
		COleDateTime dtInitialEndDate;
		LONG lDeltaLeft;
		LONG lDeltaRight;
		LONG lDurationFactor;
		void Clear()
		{
			lInitialRowIndex = 0;
			sInitialRowKey = "";
			lFinalRowIndex = 0;
			dtInitialStartDate = (DATE)0;
			dtInitialEndDate = (DATE)0;
			lDeltaLeft = 0;
			lDeltaRight = 0;
			lDurationFactor = 0;
		}
	};

	struct S_TASKSTRETCHLEFT
	{
		COleDateTime dtInitialStartDate;
		COleDateTime dtInitialEndDate;
		COleDateTime dtFinalStartDate;
		LONG lRowIndex;
		void Clear()
		{
			dtInitialStartDate = (DATE)0;
			dtInitialEndDate = (DATE)0;
			dtFinalStartDate = (DATE)0;
			lRowIndex = 0;
		}
	};

	struct S_TASKSTRETCHRIGHT
	{
		COleDateTime dtInitialStartDate;
		COleDateTime dtInitialEndDate;
		COleDateTime dtFinalEndDate;
		LONG lRowIndex;
		void Clear()
		{
			dtInitialStartDate = (DATE)0;
			dtInitialEndDate = (DATE)0;
			dtFinalEndDate = (DATE)0;
			lRowIndex = 0;
		}
	};

	struct S_TASKSELECTION
	{
		LONG lTaskIndex;
		void Clear()
		{
			lTaskIndex = 0;
		}
	};

	struct S_TASKADDITION
	{
		BOOL bCancel;
		BOOL bInConflict;
		COleDateTime dtStartDate;
		COleDateTime dtEndDate;
		LONG lRowIndex;
		void Clear()
		{
			bCancel = FALSE;
			bInConflict = FALSE;
			dtStartDate = (DATE)0;
			dtEndDate = (DATE)0;
			lRowIndex = 0;
		}
	};

	struct S_PERCENTAGESELECTION
	{
		LONG lPercentageIndex;
		void Clear()
		{
			lPercentageIndex = 0;
		}
	};

	struct S_PERCENTAGESIZING
	{
		LONG lX;
		LONG lTaskStart;
		LONG lTaskEnd;
		LONG lTaskIndex;
		BOOL bMouseMove;
		void Clear()
		{
			lX = 0;
			lTaskStart = 0;
			lTaskEnd = 0;
			lTaskIndex = 0;
			bMouseMove = TRUE;
		}
	};

	struct S_PREDECESSORSELECTION
	{
		LONG lPredecessorIndex;
		void Clear()
		{
			lPredecessorIndex = 0;
		}
	};

	struct S_PREDECESSORADDITION
	{
		LONG lXStart;
		LONG lYStart;
		LONG lTaskIndex;
		CString sTaskPosition;
		LONG lPredecessorIndex;
		CString sPredecessorKey;
		CString sPredecessorPosition;
		BOOL bCancel;
		BOOL bCantAccept;
		void Clear()
		{
			lXStart = 0;
			lYStart = 0;
			lTaskIndex = 0;
			sTaskPosition = "";
			lPredecessorIndex = 0;
			sPredecessorKey = "";
			sPredecessorPosition = "";
			bCancel = FALSE;
			bCantAccept = FALSE;
		}
	};

	typedef enum E_OPERATIONMODIFIER
	{
		OPM_NONE = 0,
		OPM_PREDECESSORMODE = 1
	}E_OPERATIONMODIFIER;

	S_TIMELINEMOVEMENT s_tmlnMVT;
	S_COLUMNMOVEMENT s_clmnMVT;
	S_COLUMNSIZING s_clmnSZ;
	S_COLUMNSELECTION s_clmnSEL;
	S_ROWMOVEMENT s_rowMVT;
	S_ROWSIZING s_rowSZ;
	S_ROWSELECTION s_rowSEL;
	S_TASKMOVEMENT s_tskMVT;
	S_TASKSTRETCHLEFT s_tskSTL;
	S_TASKSTRETCHRIGHT s_tskSTR;
	S_TASKSELECTION s_tskSEL;
	S_TASKADDITION s_tskADD;
	S_PERCENTAGESELECTION s_perSEL;
	S_PERCENTAGESIZING s_perSZ;
	S_PREDECESSORADDITION s_preADD;
	S_PREDECESSORSELECTION s_preSEL;



};