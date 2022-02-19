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


// fSTL_MSP2007 dialog

class fSTL_MSP2007 : public CDialog
{
	DECLARE_DYNAMIC(fSTL_MSP2007)

public:
	fSTL_MSP2007(CWnd* pParent = NULL);   // standard constructor
	virtual ~fSTL_MSP2007();

	CActiveGanttVCCtl	ActiveGanttVCCtl1;
	CString mp_sFontName;

// Dialog Data
	enum { IDD = IDD_FSTL_MSP2007 };

protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support

	DECLARE_MESSAGE_MAP()
public:
	virtual BOOL OnInitDialog();
	afx_msg void OnSize(UINT nType, int cx, int cy);
};