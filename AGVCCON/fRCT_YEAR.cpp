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
#include "fRCT_YEAR.h"

IMPLEMENT_DYNAMIC(fRCT_YEAR, CDialog)

fRCT_YEAR::fRCT_YEAR(CWnd* pParent /*=NULL*/)
	: CDialog(fRCT_YEAR::IDD, pParent)
{

}

fRCT_YEAR::~fRCT_YEAR()
{
}

void fRCT_YEAR::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	DDX_Control(pDX, IDC_ACTIVEGANTTVCCTL1, ActiveGanttVCCtl1);
}


BEGIN_MESSAGE_MAP(fRCT_YEAR, CDialog)
	ON_WM_SIZE()
END_MESSAGE_MAP()

BOOL fRCT_YEAR::OnInitDialog()
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
    oView = ActiveGanttVCCtl1.GetViews().Add(IL_DAY, 1, ST_YEAR, ST_NOT_VISIBLE, ST_MONTH, _T("View1"));
    oView.GetTimeLine().GetTickMarkArea().SetVisible(FALSE);

    ActiveGanttVCCtl1.SetTierFormatScope(OS_CONTROL);
    ActiveGanttVCCtl1.GetTierFormat().SetMonthIntervalFormat(_T("MM"));

    ActiveGanttVCCtl1.SetCurrentView(_T("View1"));

    int i = 0;
    for (i = 1; i <= 50; i++)
    {
        ActiveGanttVCCtl1.GetRows().Add(_T("K") + CStr(i), _T(""), FALSE, TRUE, _T(""));
    }

    CclsTimeBlock oTimeBlock;
    COleDateTime dtDate;
    dtDate = GetDateTime(2000, 12, 23, 0, 0, 0);



    oTimeBlock = ActiveGanttVCCtl1.GetTimeBlocks().Add(_T("TimeBlock1"));
    oTimeBlock.SetBaseDate(dtDate);
    oTimeBlock.SetDurationInterval(IL_DAY);
    oTimeBlock.SetDurationFactor(15);
    oTimeBlock.SetTimeBlockType(TBT_RECURRING);
    oTimeBlock.SetRecurringType(RCT_YEAR);

	return TRUE;
}

void fRCT_YEAR::OnSize(UINT nType, int cx, int cy)
{
	CDialog::OnSize(nType, cx, cy);
	g_Resize(this, &ActiveGanttVCCtl1);
}