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


// fRCT_MONTH dialog

class fRCT_MONTH : public CDialog
{
	DECLARE_DYNAMIC(fRCT_MONTH)

public:
	fRCT_MONTH(CWnd* pParent = NULL);   // standard constructor
	virtual ~fRCT_MONTH();

	CActiveGanttVCCtl	ActiveGanttVCCtl1;

// Dialog Data
	enum { IDD = IDD_FRCT_MONTH };

protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support

	DECLARE_MESSAGE_MAP()
public:
	virtual BOOL OnInitDialog();
	afx_msg void OnSize(UINT nType, int cx, int cy);
};