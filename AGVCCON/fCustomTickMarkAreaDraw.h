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


// fCustomTickMarkAreaDraw dialog

class fCustomTickMarkAreaDraw : public CDialog
{
	DECLARE_DYNAMIC(fCustomTickMarkAreaDraw)

public:
	fCustomTickMarkAreaDraw(CWnd* pParent = NULL);   // standard constructor
	virtual ~fCustomTickMarkAreaDraw();

	CActiveGanttVCCtl	ActiveGanttVCCtl1;

// Dialog Data
	enum { IDD = IDD_FCUSTOMTICKMARKAREADRAW };

protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support

	DECLARE_MESSAGE_MAP()
public:
	DECLARE_EVENTSINK_MAP()
	void CustomTickMarkAreaDrawActiveganttvcctl1(LPDISPATCH e);
	void CustomTickMarkDrawActiveganttvcctl1(LPDISPATCH e);
	afx_msg void OnSize(UINT nType, int cx, int cy);
	virtual BOOL OnInitDialog();
};