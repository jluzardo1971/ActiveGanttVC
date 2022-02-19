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


// fRCT_DAY dialog

class fRCT_DAY : public CDialog
{
	DECLARE_DYNAMIC(fRCT_DAY)

public:
	fRCT_DAY(CWnd* pParent = NULL);   // standard constructor
	virtual ~fRCT_DAY();

	CActiveGanttVCCtl	ActiveGanttVCCtl1;

// Dialog Data
	enum { IDD = IDD_FRCT_DAY };

protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support

	DECLARE_MESSAGE_MAP()
public:
	virtual BOOL OnInitDialog();
	afx_msg void OnSize(UINT nType, int cx, int cy);
};