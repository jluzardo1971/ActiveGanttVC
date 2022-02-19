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
#include "afxwin.h"

class clsTextBox : public CEdit
{
public:
	CActiveGanttVCCtl* mp_oControl;

	clsTextBox(CActiveGanttVCCtl* oControl);
	~clsTextBox(void);

	void Initialize(int lIndex, int lIndex2, E_TEXTOBJECTTYPE yObjectType, int X, int Y);
	void Terminate();
	void SetGDIFont(clsFont* oFont);
	void OnUpdateEdit();

// Member Variables
public:
	clsColumn* mp_oColumn;
    clsRow* mp_oRow;
    clsCell* mp_oCell;
    clsNode* mp_oNode;
    clsTask* mp_oTask;
    LONG mp_yObjectType;
    CString mp_sText;
    LONG mp_lIndex;
    LONG mp_lIndex2;
	BOOL mp_bInitialized;
	CFont mp_oFont;
	Gdiplus::Color mp_clrBackColor;
	Gdiplus::Color mp_clrForeColor;

	


protected:
	DECLARE_MESSAGE_MAP()
public:
};