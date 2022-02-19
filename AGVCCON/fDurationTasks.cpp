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
#include "fDurationTasks.h"

IMPLEMENT_DYNAMIC(fDurationTasks, CDialog)

fDurationTasks::fDurationTasks(CWnd* pParent /*=NULL*/)
	: CDialog(fDurationTasks::IDD, pParent)
{

}

fDurationTasks::~fDurationTasks()
{
}

void fDurationTasks::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	DDX_Control(pDX, IDC_ACTIVEGANTTVCCTL1, ActiveGanttVCCtl1);
}


BEGIN_MESSAGE_MAP(fDurationTasks, CDialog)
	ON_WM_SIZE()
END_MESSAGE_MAP()

BOOL fDurationTasks::OnInitDialog()
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


    ActiveGanttVCCtl1.SetAddMode(AT_DURATION_BOTH);
    ActiveGanttVCCtl1.SetAddDurationInterval(IL_HOUR);

    CclsView oView;
    oView = ActiveGanttVCCtl1.GetViews().Add(IL_MINUTE, 10, ST_MONTH, ST_NOT_VISIBLE, ST_DAYOFWEEK, _T("View1"));
    oView.GetTimeLine().GetTickMarkArea().SetVisible(FALSE);

    ActiveGanttVCCtl1.SetCurrentView(_T("View1"));

    LONG i = 0;
    for (i = 0; i <= 110; i++)
    {
        ActiveGanttVCCtl1.GetRows().Add(_T("K") + CStr(i), _T(""), FALSE, TRUE, _T(""));
    }

    CclsTimeBlock oTimeBlock;

    //Note: non-working overlapping TimeBlock objects are combined for duration calculation purposes.


    // TimeBlock starts at 6:00pm and ends on 7:00am next day (13 Hours)
    // This TimeBlock is repeated every day.
    oTimeBlock = ActiveGanttVCCtl1.GetTimeBlocks().Add(_T("TB_OutOfOfficeHours"));
    oTimeBlock.SetNonWorking(TRUE);
    oTimeBlock.SetBaseDate(GetDateTime(2000, 1, 1, 18, 0, 0));
    oTimeBlock.SetDurationInterval(IL_HOUR);
    oTimeBlock.SetDurationFactor(13);
    oTimeBlock.SetTimeBlockType(TBT_RECURRING);
    oTimeBlock.SetRecurringType(RCT_DAY);

    // TimeBlock starts at 12:00pm (noon) and ends at 1:30pm (90 Minutes)
    // This TimeBlock is repeated every day. 
    oTimeBlock = ActiveGanttVCCtl1.GetTimeBlocks().Add(_T("TB_LunchBreak"));
    oTimeBlock.SetNonWorking(TRUE);
    oTimeBlock.SetBaseDate(GetDateTime(2000, 1, 1, 12, 0, 0));
    oTimeBlock.SetDurationInterval(IL_MINUTE);
    oTimeBlock.SetDurationFactor(90);
    oTimeBlock.SetTimeBlockType(TBT_RECURRING);
    oTimeBlock.SetRecurringType(RCT_DAY);

    // Timeblock starts at 12:00am Saturday and ends on 12:00am Monday (48 Hours)
    // This TimeBlock is repeated every week.
    oTimeBlock = ActiveGanttVCCtl1.GetTimeBlocks().Add(_T("TB_Weekend"));
    oTimeBlock.SetNonWorking(TRUE);
    oTimeBlock.SetBaseDate(GetDateTime(2000, 1, 1, 0, 0, 0));
    oTimeBlock.SetDurationInterval(IL_HOUR);
    oTimeBlock.SetDurationFactor(48);
    oTimeBlock.SetTimeBlockType(TBT_RECURRING);
    oTimeBlock.SetRecurringType(RCT_WEEK);
    oTimeBlock.SetBaseWeekDay(WD_SATURDAY);

    // Arbitrary holiday that starts at 12:00am January 8th and ends on 12:00am January 9th (24 hours)
    // This TimeBlock is repeated every year.
    oTimeBlock = ActiveGanttVCCtl1.GetTimeBlocks().Add(_T("TB_Jan8"));
    oTimeBlock.SetNonWorking(TRUE);
    oTimeBlock.SetBaseDate(GetDateTime(2000, 1, 8, 0, 0, 0));
    oTimeBlock.SetDurationInterval(IL_HOUR);
    oTimeBlock.SetDurationFactor(24);
    oTimeBlock.SetTimeBlockType(TBT_RECURRING);
    oTimeBlock.SetRecurringType(RCT_YEAR);

    ActiveGanttVCCtl1.GetTimeBlocks().SetIntervalStart(GetDateTime(2012, 1, 1));
    ActiveGanttVCCtl1.GetTimeBlocks().SetIntervalEnd(GetDateTime(2023, 6, 1));
    ActiveGanttVCCtl1.GetTimeBlocks().SetIntervalType(TBIT_MANUAL);
    ActiveGanttVCCtl1.GetTimeBlocks().CalculateInterval();


    CclsTask oTask;
    for (i = 0; i <= 100; i++)
    {
        oTask = ActiveGanttVCCtl1.GetTasks().DAdd(_T("K") + CStr(i), GetDateTime(2013, 1, 1, 0, 0, 0), IL_HOUR, i, CStr(i), _T(""), _T(""), _T(""));
    }
    ActiveGanttVCCtl1.GetCurrentViewObject().GetTimeLine().Position(GetDateTime(2013, 1, 1, 0, 0, 0));
	ActiveGanttVCCtl1.Redraw();
	return TRUE;
}

void fDurationTasks::OnSize(UINT nType, int cx, int cy)
{
	CDialog::OnSize(nType, cx, cy);
	g_Resize(this, &ActiveGanttVCCtl1);
}