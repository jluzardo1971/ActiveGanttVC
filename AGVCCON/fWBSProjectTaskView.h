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

class fWBSProject;

class fWBSProjectTaskView : public CDialog
{
	DECLARE_DYNAMIC(fWBSProjectTaskView)

public:
	fWBSProjectTaskView(CWnd* pParent = NULL);   
	virtual ~fWBSProjectTaskView();

	CMSFlexGrid oGrid;
	fWBSProject* mp_oParent;
	LONG mp_lTaskIndex;

	CclsTask GetPredecessorTask(CclsPredecessor* oPredecessor);
	CString GetPredecessorType(LONG lType);


	enum { IDD = IDD_FWBSPROJECTTASKVIEW };

protected:
	virtual void DoDataExchange(CDataExchange* pDX);    

	DECLARE_MESSAGE_MAP()
public:
	CEdit lblTaskName;
	virtual BOOL OnInitDialog();
};