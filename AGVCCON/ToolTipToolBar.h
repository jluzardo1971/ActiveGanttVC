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

// CToolTipToolBar

class CToolTipToolBar : public CToolBar
{
	DECLARE_DYNAMIC(CToolTipToolBar)

public:
	CToolTipToolBar();
	virtual ~CToolTipToolBar();

	vCString m_asToolTips;
	int mp_lCounter;

	void InsertToolTip(CString sToolTip);
	CString GetToolTip(int nID);

	void AddButton(int iImage, CString sToolTipText);
	void AddSeparator();
	void AddSeparator(int lWidth);
	void AddLabel(int lWidth);

	int GetItemLeft(int lIndex);
	int GetItemWidth(int lIndex);

protected:
	DECLARE_MESSAGE_MAP()
	virtual BOOL OnNotify(WPARAM wParam, LPARAM lParam, LRESULT* pResult);
};