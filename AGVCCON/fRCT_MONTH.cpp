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
#include "fRCT_MONTH.h"

IMPLEMENT_DYNAMIC(fRCT_MONTH, CDialog)

fRCT_MONTH::fRCT_MONTH(CWnd* pParent /*=NULL*/)
	: CDialog(fRCT_MONTH::IDD, pParent)
{

}

fRCT_MONTH::~fRCT_MONTH()
{
}

void fRCT_MONTH::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	DDX_Control(pDX, IDC_ACTIVEGANTTVCCTL1, ActiveGanttVCCtl1);
}


BEGIN_MESSAGE_MAP(fRCT_MONTH, CDialog)
	ON_WM_SIZE()
END_MESSAGE_MAP()

BOOL fRCT_MONTH::OnInitDialog()
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

    CclsView oView;
    oView = ActiveGanttVCCtl1.GetViews().Add(IL_HOUR, 1, ST_MONTH, ST_NOT_VISIBLE, ST_DAYOFWEEK, _T("View1"));
    oView.GetTimeLine().GetTickMarkArea().SetVisible(FALSE);
    ActiveGanttVCCtl1.SetTierFormatScope(OS_CONTROL);
    ActiveGanttVCCtl1.GetTierFormat().SetDayOfWeekIntervalFormat(_T("dd"));

    ActiveGanttVCCtl1.SetCurrentView(_T("View1"));

    int i = 0;
    for (i = 1; i <= 50; i++)
    {
        ActiveGanttVCCtl1.GetRows().Add(_T("K") + CStr(i), _T(""), FALSE, TRUE, _T(""));
    }

    CclsTimeBlock oTimeBlock;
    oTimeBlock = ActiveGanttVCCtl1.GetTimeBlocks().Add(_T("TimeBlock1"));
    oTimeBlock.SetBaseDate(GetDateTime(2000, 1, 1, 0, 0, 0));
    oTimeBlock.SetDurationInterval(IL_HOUR);
    oTimeBlock.SetDurationFactor(-48);
    oTimeBlock.SetTimeBlockType(TBT_RECURRING);
    oTimeBlock.SetRecurringType(RCT_MONTH);

	return TRUE;
}

void fRCT_MONTH::OnSize(UINT nType, int cx, int cy)
{
	CDialog::OnSize(nType, cx, cy);
	g_Resize(this, &ActiveGanttVCCtl1);
}