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
#include "fSTL_MSP2003.h"
#include "afxdialogex.h"

IMPLEMENT_DYNAMIC(fSTL_MSP2003, CDialog)

fSTL_MSP2003::fSTL_MSP2003(CWnd* pParent /*=NULL*/)
	: CDialog(fSTL_MSP2003::IDD, pParent)
{
	mp_sFontName = _T("Tahoma");
}

fSTL_MSP2003::~fSTL_MSP2003()
{
}

void fSTL_MSP2003::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	DDX_Control(pDX, IDC_ACTIVEGANTTVCCTL1, ActiveGanttVCCtl1);
}


BEGIN_MESSAGE_MAP(fSTL_MSP2003, CDialog)
	ON_WM_SIZE()
END_MESSAGE_MAP()

BOOL fSTL_MSP2003::OnInitDialog()
{
	CDialog::OnInitDialog();
	g_MaximizeWindowsDim(this);
	CWnd::ShowWindow(SW_SHOWMAXIMIZED);

	CclsStyle oStyle;
	CclsView oView;

	oStyle = ActiveGanttVCCtl1.GetStyles().Add(_T("TimeLineTiers"));
	oStyle.GetFont().SetFontName(mp_sFontName);
	oStyle.GetFont().SetSize(7);
	oStyle.GetFont().SetStyle(FS_FONTSTYLEREGULAR);
	oStyle.SetAppearance(SA_RAISED);
	oStyle.SetBorderColor(CLR_DARKGRAY);
	oStyle.SetBorderStyle(SBR_SINGLE);

	oStyle = ActiveGanttVCCtl1.GetStyles().Add(_T("TaskStyle"));
	oStyle.SetAppearance(SA_FLAT);
	oStyle.SetPlacement(PLC_OFFSETPLACEMENT);
	oStyle.SetBackColor(CLR_BLUE);
	oStyle.SetBorderColor(CLR_BLUE);
	oStyle.SetBorderStyle(SBR_SINGLE);
	oStyle.GetSelectionRectangleStyle().SetOffsetTop(0);
	oStyle.GetSelectionRectangleStyle().SetOffsetLeft(0);
	oStyle.GetSelectionRectangleStyle().SetOffsetRight(0);
	oStyle.GetSelectionRectangleStyle().SetOffsetBottom(0);
	oStyle.SetTextPlacement(SCP_EXTERIORPLACEMENT);
	oStyle.SetTextAlignmentHorizontal(HAL_RIGHT);
	oStyle.SetTextXMargin(10);
	oStyle.SetOffsetTop(5);
	oStyle.SetOffsetBottom(10);
	oStyle.SetBackgroundMode(FP_HATCH);
	oStyle.SetHatchBackColor(CLR_WHITE);
	oStyle.SetHatchForeColor(CLR_BLUE);
	oStyle.SetHatchStyle(HS_PERCENT50);
	oStyle.GetPredecessorStyle().SetLineColor(CLR_BLACK);
	oStyle.GetMilestoneStyle().SetShapeIndex(FT_DIAMOND);
    oStyle.GetMilestoneStyle().SetFillColor(CLR_BLUE);
    oStyle.GetMilestoneStyle().SetBorderColor(CLR_BLUE);
	oStyle.GetPredecessorStyle().SetXOffset(4);
	oStyle.GetPredecessorStyle().SetYOffset(4);
	oStyle.GetFont().SetStyle(FS_FONTSTYLEBOLD);

	oStyle = ActiveGanttVCCtl1.GetStyles().Add(_T("SummaryStyle"));
	oStyle.SetAppearance(SA_FLAT);
	oStyle.SetPlacement(PLC_OFFSETPLACEMENT);
	oStyle.SetBackColor(CLR_GREEN);
	oStyle.SetBorderColor(CLR_GREEN);
	oStyle.SetBorderStyle(SBR_SINGLE);
	oStyle.GetSelectionRectangleStyle().SetVisible(FALSE);
	oStyle.SetTextPlacement(SCP_EXTERIORPLACEMENT);
	oStyle.SetTextAlignmentHorizontal(HAL_RIGHT);
	oStyle.SetTextXMargin(10);
	oStyle.GetTaskStyle().SetStartShapeIndex(FT_PROJECTDOWN);
	oStyle.GetTaskStyle().SetEndShapeIndex(FT_PROJECTDOWN);
	oStyle.GetTaskStyle().SetEndFillColor(CLR_GREEN);
	oStyle.GetTaskStyle().SetEndBorderColor(CLR_GREEN);
	oStyle.GetTaskStyle().SetStartFillColor(CLR_GREEN);
	oStyle.GetTaskStyle().SetStartBorderColor(CLR_GREEN);
	oStyle.SetFillMode(FM_UPPERHALFFILLED);

	oStyle = ActiveGanttVCCtl1.GetStyles().Add(_T("CellStyleKeyColumn"));
	oStyle.SetAppearance(SA_FLAT);
	oStyle.SetBackColor(CLR_WHITE);
	oStyle.SetBorderColor(FromArgb(255, 128, 128, 128));
	oStyle.SetBorderStyle(SBR_CUSTOM);
	oStyle.GetCustomBorderStyle().SetTop(FALSE);
	oStyle.GetCustomBorderStyle().SetLeft(FALSE);
	oStyle.SetTextAlignmentHorizontal(HAL_RIGHT);
	oStyle.SetTextXMargin(4);

	ActiveGanttVCCtl1.GetSplitter().SetPosition(255);
	ActiveGanttVCCtl1.GetTreeview().SetImages(TRUE);
	ActiveGanttVCCtl1.GetTreeview().SetCheckBoxes(TRUE);
	ActiveGanttVCCtl1.GetTreeview().SetTreeLines(FALSE);

	CclsColumn oColumn;

	oColumn = ActiveGanttVCCtl1.GetColumns().Add(_T("ID"), _T(""), 30, _T(""));
	oColumn.SetAllowTextEdit(TRUE);

	oColumn = ActiveGanttVCCtl1.GetColumns().Add(_T("Task Name"), _T(""), 255, _T(""));
	oColumn.SetAllowTextEdit(TRUE);

	ActiveGanttVCCtl1.SetTreeviewColumnIndex(2);

	ActiveGanttVCCtl1.GetTreeview().SetImages(TRUE);
	ActiveGanttVCCtl1.GetTreeview().SetCheckBoxes(TRUE);

	int i;
	CclsRow oRow;
	for (i = 1; i <= 100; i++)
	{
		oRow = ActiveGanttVCCtl1.GetRows().Add(_T("K") + CStr(i), _T(""), FALSE, TRUE, _T(""));
		oRow.SetMergeCells(FALSE);
		oRow.GetNode().SetText(_T("Row: ") + CStr(i));
        oRow.GetCells().Item(_T("1")).SetText(CStr(i));
        oRow.GetCells().Item(_T("1")).SetStyleIndex(_T("CellStyleKeyColumn"));
		oRow.SetHeight(20);
		if (oRow.GetIndex() % 5 == 0 || oRow.GetIndex() == 1)
		{
			oRow.GetNode().SetDepth(0);
		}
		else
		{
			oRow.GetNode().SetCheckBoxVisible(TRUE);
			oRow.GetNode().SetDepth(1);
		}
	}

	ActiveGanttVCCtl1.GetRows().UpdateTree();

	ActiveGanttVCCtl1.GetTimeLineScrollBar().SetStartDate(GetDateTime(2013, 1, 1));
	ActiveGanttVCCtl1.GetTimeLineScrollBar().SetInterval(IL_HOUR);
	ActiveGanttVCCtl1.GetTimeLineScrollBar().SetFactor(1);
	ActiveGanttVCCtl1.GetTimeLineScrollBar().SetSmallChange(6);
	ActiveGanttVCCtl1.GetTimeLineScrollBar().SetLargeChange(480);
	ActiveGanttVCCtl1.GetTimeLineScrollBar().SetMax(4000);
	ActiveGanttVCCtl1.GetTimeLineScrollBar().SetValue(0);
	ActiveGanttVCCtl1.GetTimeLineScrollBar().SetEnabled(TRUE);
	ActiveGanttVCCtl1.GetTimeLineScrollBar().SetVisible(TRUE);

	ActiveGanttVCCtl1.GetTierFormat().SetQuarterIntervalFormat(_T("yyyy Q\\Q"));
	ActiveGanttVCCtl1.GetTierFormat().SetMonthIntervalFormat(_T("MMM"));

	oView = ActiveGanttVCCtl1.GetViews().Add(IL_HOUR, 24, ST_QUARTER, ST_NOT_VISIBLE, ST_MONTH, _T(""));
	oView.GetTimeLine().GetTierArea().GetUpperTier().SetBackgroundMode(ETBGM_STYLE);
	oView.GetTimeLine().GetTierArea().GetUpperTier().SetStyleIndex(_T("TimeLineTiers"));
	oView.GetTimeLine().GetTierArea().GetLowerTier().SetBackgroundMode(ETBGM_STYLE);
	oView.GetTimeLine().GetTierArea().GetLowerTier().SetStyleIndex(_T("TimeLineTiers"));
	oView.GetTimeLine().GetTickMarkArea().SetVisible(FALSE);

	ActiveGanttVCCtl1.SetCurrentView(_T("1"));

	ActiveGanttVCCtl1.GetTasks().Add(_T(""), _T("K1"), GetDateTime(2013, 3, 1), GetDateTime(2014, 3, 1), _T("SS"), _T("SummaryStyle"), _T("0"));

    ActiveGanttVCCtl1.GetTasks().Add(_T(""), _T("K5"), GetDateTime(2013, 3, 1), GetDateTime(2014, 3, 1), _T("TS"), _T("TaskStyle"), _T("0"));

	return TRUE;
}


void fSTL_MSP2003::OnSize(UINT nType, int cx, int cy)
{
	CDialog::OnSize(nType, cx, cy);
	g_Resize(this, &ActiveGanttVCCtl1);
}