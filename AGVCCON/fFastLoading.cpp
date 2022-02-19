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
#include "AGVCCON.h"
#include "fFastLoading.h"

IMPLEMENT_DYNAMIC(fFastLoading, CDialog)

fFastLoading::fFastLoading(CWnd* pParent /*=NULL*/)
	: CDialog(fFastLoading::IDD, pParent)
{

}

fFastLoading::~fFastLoading()
{
}

void fFastLoading::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	DDX_Control(pDX, IDC_ACTIVEGANTTVCCTL1, ActiveGanttVCCtl1);
}


BEGIN_MESSAGE_MAP(fFastLoading, CDialog)
	ON_WM_SIZE()
END_MESSAGE_MAP()

BOOL fFastLoading::OnInitDialog()
{
	CDialog::OnInitDialog();
	g_MaximizeWindowsDim(this);
	CWnd::ShowWindow(SW_SHOWMAXIMIZED);

	//If you open the form: Styles And Templates -> Available Templates in the main menu (fTemplates.vb)
	//you can preview all available Templates.
	//Or you can also build your own:
	//fMSProject11.vb shows you how to build a Solid Template in the InitializeAG Method.
	//fMSProject14.vb shows you how to build a Gradient Template in the InitializeAG Method.
	ActiveGanttVCCtl1.ApplyTemplate(STC_CH_VGRAD_ANAKIWA_BLUE, STO_DEFAULT);

	int i = 0;
	ActiveGanttVCCtl1.GetColumns().Add(_T("Tasks"), _T(""), 200, _T(""));
	ActiveGanttVCCtl1.SetTreeviewColumnIndex(1);
    ActiveGanttVCCtl1.GetRows().BeginLoad(FALSE);
    ActiveGanttVCCtl1.GetTasks().BeginLoad(FALSE);
    int lCurrentDepth = 0;
    for (i = 1; i <= 5000; i++)
    {
        CclsRow oRow;
        CclsTask oTask;
        oRow = ActiveGanttVCCtl1.GetRows().Load(_T("K") + CStr(i));
        oTask = ActiveGanttVCCtl1.GetTasks().Load(_T("K") + CStr(i), _T("K") + CStr(i));
		oRow.SetHeight(20);
        oRow.SetText(_T("Task K") + CStr(i));
        oRow.SetMergeCells(TRUE);
        oRow.GetNode().SetDepth(lCurrentDepth);
        oTask.SetText(_T("K") + CStr(i));
		oTask.SetStartDate(ActiveGanttVCCtl1.GetMathLib().DateTimeAdd(IL_HOUR, GetRnd(0, 5), GetCurrentDateTime()));
        oTask.SetEndDate(ActiveGanttVCCtl1.GetMathLib().DateTimeAdd(IL_HOUR, GetRnd(2, 7), oTask.GetStartDate()));
        lCurrentDepth = lCurrentDepth + GetRnd(-1, 1);
        if (lCurrentDepth < 0)
        {
            lCurrentDepth = 0;
        }
    }
    ActiveGanttVCCtl1.GetTasks().EndLoad();
    ActiveGanttVCCtl1.GetRows().EndLoad();
    ActiveGanttVCCtl1.GetRows().BeginLoad(TRUE);
    ActiveGanttVCCtl1.GetTasks().BeginLoad(TRUE);
    for (i = 5001; i <= 10000; i++)
    {
        CclsRow oRow;
        CclsTask oTask;
        oRow = ActiveGanttVCCtl1.GetRows().Load(_T("KL") + CStr(i));
        oTask = ActiveGanttVCCtl1.GetTasks().Load(_T("KL") + CStr(i), _T("KL") + CStr(i));
		oRow.SetHeight(20);
        oRow.SetText(_T("Task KL") + CStr(i));
        oRow.SetMergeCells(TRUE);
        oTask.SetText(_T("KL") + CStr(i));
        oTask.SetStartDate(ActiveGanttVCCtl1.GetMathLib().DateTimeAdd(IL_HOUR, GetRnd(0, 5), GetCurrentDateTime()));
        oTask.SetEndDate(ActiveGanttVCCtl1.GetMathLib().DateTimeAdd(IL_HOUR, GetRnd(2, 7), oTask.GetStartDate()));
    }
    ActiveGanttVCCtl1.GetTasks().EndLoad();
    ActiveGanttVCCtl1.GetRows().EndLoad();
    ActiveGanttVCCtl1.Redraw();



	return TRUE;
}
BEGIN_EVENTSINK_MAP(fFastLoading, CDialog)
END_EVENTSINK_MAP()



void fFastLoading::OnSize(UINT nType, int cx, int cy)
{
	CDialog::OnSize(nType, cx, cy);
	g_Resize(this, &ActiveGanttVCCtl1);
}