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
#include "fSTL_MSP2007.h"
#include "afxdialogex.h"

IMPLEMENT_DYNAMIC(fSTL_MSP2007, CDialog)

fSTL_MSP2007::fSTL_MSP2007(CWnd* pParent /*=NULL*/)
	: CDialog(fSTL_MSP2007::IDD, pParent)
{
	mp_sFontName = _T("Tahoma");
}

fSTL_MSP2007::~fSTL_MSP2007()
{
}

void fSTL_MSP2007::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	DDX_Control(pDX, IDC_ACTIVEGANTTVCCTL1, ActiveGanttVCCtl1);
}


BEGIN_MESSAGE_MAP(fSTL_MSP2007, CDialog)
	ON_WM_SIZE()
END_MESSAGE_MAP()

BOOL fSTL_MSP2007::OnInitDialog()
{
	CDialog::OnInitDialog();
	g_MaximizeWindowsDim(this);
	CWnd::ShowWindow(SW_SHOWMAXIMIZED);

	CclsStyle oStyle;
	CclsView oView;

	oStyle = ActiveGanttVCCtl1.GetStyles().Add(_T("ScrollBar"));
	oStyle.SetAppearance(SA_FLAT);
	oStyle.SetBackgroundMode(FP_SOLID);
	oStyle.SetBackColor(FromArgb(255, 122, 151, 193));
	oStyle.SetBorderStyle(SBR_SINGLE);
	oStyle.SetBorderColor(FromArgb(255, 139, 144, 150));

	oStyle = ActiveGanttVCCtl1.GetStyles().Add(_T("ArrowButtons"));
	oStyle.SetAppearance(SA_FLAT);
	oStyle.SetBackgroundMode(FP_SOLID);
	oStyle.SetBackColor(FromArgb(255, 122, 151, 193));
	oStyle.SetBorderStyle(SBR_SINGLE);
	oStyle.SetBorderColor(FromArgb(255, 150, 158, 168));

	oStyle = ActiveGanttVCCtl1.GetStyles().Add(_T("ThumbButtonH"));
	oStyle.SetAppearance(SA_FLAT);
	oStyle.SetBackgroundMode(FP_GRADIENT);
	oStyle.SetGradientFillMode(GDT_VERTICAL);
	oStyle.SetStartGradientColor(FromArgb(255, 240, 240, 240));
	oStyle.SetEndGradientColor(FromArgb(255, 165, 186, 207));
	oStyle.SetBorderStyle(SBR_SINGLE);
	oStyle.SetBorderColor(FromArgb(255, 138, 145, 153));

	oStyle = ActiveGanttVCCtl1.GetStyles().Add(_T("ThumbButtonV"));
	oStyle.SetAppearance(SA_FLAT);
	oStyle.SetBackgroundMode(FP_GRADIENT);
	oStyle.SetGradientFillMode(GDT_HORIZONTAL);
	oStyle.SetStartGradientColor(FromArgb(255, 240, 240, 240));
	oStyle.SetEndGradientColor(FromArgb(255, 165, 186, 207));
	oStyle.SetBorderStyle(SBR_SINGLE);
	oStyle.SetBorderColor(FromArgb(255, 138, 145, 153));

	oStyle = ActiveGanttVCCtl1.GetStyles().Add(_T("TimeLineTiers"));
	oStyle.GetFont().SetFontName(mp_sFontName);
	oStyle.GetFont().SetSize(7);
	oStyle.GetFont().SetStyle(FS_FONTSTYLEREGULAR);
	oStyle.SetAppearance(SA_FLAT);
	oStyle.SetBackgroundMode(FP_TRANSPARENT);
	oStyle.GetCustomBorderStyle().SetLeft(TRUE);
	oStyle.GetCustomBorderStyle().SetTop(FALSE);
	oStyle.GetCustomBorderStyle().SetRight(FALSE);
	oStyle.GetCustomBorderStyle().SetBottom(TRUE);
	oStyle.SetBorderStyle(SBR_CUSTOM);
	oStyle.SetBorderColor(FromArgb(255, 197, 206, 216));

	oStyle = ActiveGanttVCCtl1.GetStyles().Add(_T("TimeLine"));
	oStyle.SetAppearance(SA_FLAT);
	oStyle.SetBackgroundMode(FP_GRADIENT);
	oStyle.SetStartGradientColor(FromArgb(255, 179, 206, 235));
	oStyle.SetEndGradientColor(FromArgb(255, 161, 193, 232));
	oStyle.SetGradientFillMode(GDT_VERTICAL);
	oStyle.SetBackColor(CLR_WHITE);
	oStyle.SetBorderStyle(SBR_NONE);

	oStyle = ActiveGanttVCCtl1.GetStyles().Add(_T("ColumnStyle"));
	oStyle.SetAppearance(SA_FLAT);
	oStyle.SetBackgroundMode(FP_GRADIENT);
	oStyle.SetStartGradientColor(FromArgb(255, 179, 206, 235));
	oStyle.SetEndGradientColor(FromArgb(255, 161, 193, 232));
	oStyle.SetGradientFillMode(GDT_VERTICAL);
	oStyle.SetBackColor(CLR_WHITE);
	oStyle.SetBorderStyle(SBR_CUSTOM);
	oStyle.GetCustomBorderStyle().SetLeft(FALSE);
	oStyle.GetCustomBorderStyle().SetTop(FALSE);
	oStyle.GetCustomBorderStyle().SetRight(TRUE);
	oStyle.GetCustomBorderStyle().SetBottom(TRUE);
	oStyle.SetBorderColor(FromArgb(255, 197, 206, 216));

	oStyle = ActiveGanttVCCtl1.GetStyles().Add(_T("TaskStyle"));
	oStyle.SetAppearance(SA_FLAT);
	oStyle.SetBackgroundMode(FP_GRADIENT);
	oStyle.SetStartGradientColor(FromArgb(255, 240, 240, 240));
	oStyle.SetEndGradientColor(FromArgb(255, 0, 0, 255));
	oStyle.SetGradientFillMode(GDT_VERTICAL);
	oStyle.SetBorderStyle(SBR_SINGLE);
	oStyle.SetBorderColor(FromArgb(255, 148, 152, 179));
	oStyle.SetPlacement(PLC_OFFSETPLACEMENT);
	oStyle.GetSelectionRectangleStyle().SetOffsetTop(0);
	oStyle.GetSelectionRectangleStyle().SetOffsetLeft(0);
	oStyle.GetSelectionRectangleStyle().SetOffsetRight(0);
	oStyle.GetSelectionRectangleStyle().SetOffsetBottom(0);
	oStyle.SetTextPlacement(SCP_EXTERIORPLACEMENT);
	oStyle.SetTextAlignmentHorizontal(HAL_RIGHT);
	oStyle.SetTextXMargin(10);
	oStyle.SetOffsetTop(5);
	oStyle.SetOffsetBottom(10);
	oStyle.GetPredecessorStyle().SetLineColor(FromArgb(255, 160, 160, 160));
	oStyle.GetMilestoneStyle().SetShapeIndex(FT_DIAMOND);
    oStyle.GetMilestoneStyle().SetFillColor(CLR_BLUE);
    oStyle.GetMilestoneStyle().SetBorderColor(CLR_BLUE);
	oStyle.GetPredecessorStyle().SetXOffset(4);
	oStyle.GetPredecessorStyle().SetYOffset(4);

	oStyle = ActiveGanttVCCtl1.GetStyles().Add(_T("SummaryStyle"));
	oStyle.SetAppearance(SA_FLAT);
	oStyle.SetBackgroundMode(FP_GRADIENT);
	oStyle.SetStartGradientColor(FromArgb(255, 0, 0, 0));
	oStyle.SetEndGradientColor(FromArgb(255, 240, 240, 240));
	oStyle.SetGradientFillMode(GDT_VERTICAL);
	oStyle.SetBackColor(CLR_BLACK);
	oStyle.SetBorderStyle(SBR_SINGLE);
	oStyle.SetBorderColor(CLR_BLACK);
	oStyle.SetFillMode(FM_UPPERHALFFILLED);
	oStyle.SetPlacement(PLC_OFFSETPLACEMENT);
	oStyle.SetOffsetTop(5);
	oStyle.SetOffsetBottom(10);
	oStyle.GetSelectionRectangleStyle().SetOffsetTop(0);
	oStyle.GetSelectionRectangleStyle().SetOffsetLeft(0);
	oStyle.GetSelectionRectangleStyle().SetOffsetRight(0);
	oStyle.GetSelectionRectangleStyle().SetOffsetBottom(0);
	oStyle.SetTextPlacement(SCP_EXTERIORPLACEMENT);
	oStyle.SetTextAlignmentHorizontal(HAL_RIGHT);
	oStyle.SetTextXMargin(10);
	oStyle.GetPredecessorStyle().SetLineColor(CLR_BLACK);
	oStyle.GetTaskStyle().SetStartShapeIndex(FT_PROJECTDOWN);
	oStyle.GetTaskStyle().SetEndShapeIndex(FT_PROJECTDOWN);

	oStyle = ActiveGanttVCCtl1.GetStyles().Add(_T("NodeStyle"));
	oStyle.SetAppearance(SA_FLAT);
	oStyle.SetBackColor(CLR_WHITE);
	oStyle.SetBorderColor(FromArgb(255, 197, 206, 216));
	oStyle.SetBorderStyle(SBR_CUSTOM);
	oStyle.GetCustomBorderStyle().SetTop(FALSE);
	oStyle.GetCustomBorderStyle().SetLeft(FALSE);

	oStyle = ActiveGanttVCCtl1.GetStyles().Add(_T("CellStyleKeyColumn"));
	oStyle.SetAppearance(SA_FLAT);
	oStyle.SetBackColor(CLR_WHITE);
	oStyle.SetBorderColor(FromArgb(255, 197, 206, 216));
	oStyle.SetBorderStyle(SBR_CUSTOM);
	oStyle.GetCustomBorderStyle().SetTop(FALSE);
	oStyle.GetCustomBorderStyle().SetLeft(FALSE);
	oStyle.SetTextAlignmentHorizontal(HAL_RIGHT);
	oStyle.SetTextXMargin(4);

	oStyle = ActiveGanttVCCtl1.GetStyles().Add(_T("ClientAreaStyle"));
	oStyle.SetAppearance(SA_FLAT);
	oStyle.SetBackColor(CLR_WHITE);
	oStyle.SetBorderStyle(SBR_NONE);

	ActiveGanttVCCtl1.GetTreeview().SetImages(TRUE);
	ActiveGanttVCCtl1.GetTreeview().SetCheckBoxes(TRUE);
	ActiveGanttVCCtl1.GetTreeview().SetTreeLines(FALSE);

	CclsColumn oColumn;

	oColumn = ActiveGanttVCCtl1.GetColumns().Add(_T("ID"), _T(""), 30, _T(""));
	oColumn.SetStyleIndex(_T("ColumnStyle"));
	oColumn.SetAllowTextEdit(TRUE);

	oColumn = ActiveGanttVCCtl1.GetColumns().Add(_T("Task Name"), _T(""), 255, _T(""));
	oColumn.SetStyleIndex(_T("ColumnStyle"));
	oColumn.SetAllowTextEdit(TRUE);

	ActiveGanttVCCtl1.SetTreeviewColumnIndex(2);

	ActiveGanttVCCtl1.GetScrollBarSeparator().GetStyle().SetAppearance(SA_FLAT);
	ActiveGanttVCCtl1.GetScrollBarSeparator().GetStyle().SetBackgroundMode(FP_SOLID);
	ActiveGanttVCCtl1.GetScrollBarSeparator().GetStyle().SetBackColor(FromArgb(255, 164, 196, 237));

	ActiveGanttVCCtl1.GetVerticalScrollBar().GetScrollBar().SetStyleIndex(_T("ScrollBar"));
	ActiveGanttVCCtl1.GetVerticalScrollBar().GetScrollBar().GetArrowButtons().SetNormalStyleIndex(_T("ArrowButtons"));
	ActiveGanttVCCtl1.GetVerticalScrollBar().GetScrollBar().GetArrowButtons().SetPressedStyleIndex(_T("ArrowButtons"));
	ActiveGanttVCCtl1.GetVerticalScrollBar().GetScrollBar().GetArrowButtons().SetDisabledStyleIndex(_T("ArrowButtons"));
	ActiveGanttVCCtl1.GetVerticalScrollBar().GetScrollBar().GetThumbButton().SetNormalStyleIndex(_T("ThumbButtonV"));
	ActiveGanttVCCtl1.GetVerticalScrollBar().GetScrollBar().GetThumbButton().SetPressedStyleIndex(_T("ThumbButtonV"));
	ActiveGanttVCCtl1.GetVerticalScrollBar().GetScrollBar().GetThumbButton().SetDisabledStyleIndex(_T("ThumbButtonV"));

	ActiveGanttVCCtl1.GetHorizontalScrollBar().GetScrollBar().SetStyleIndex(_T("ScrollBar"));
	ActiveGanttVCCtl1.GetHorizontalScrollBar().GetScrollBar().GetArrowButtons().SetNormalStyleIndex(_T("ArrowButtons"));
	ActiveGanttVCCtl1.GetHorizontalScrollBar().GetScrollBar().GetArrowButtons().SetPressedStyleIndex(_T("ArrowButtons"));
	ActiveGanttVCCtl1.GetHorizontalScrollBar().GetScrollBar().GetArrowButtons().SetDisabledStyleIndex(_T("ArrowButtons"));
	ActiveGanttVCCtl1.GetHorizontalScrollBar().GetScrollBar().GetThumbButton().SetNormalStyleIndex(_T("ThumbButtonH"));
	ActiveGanttVCCtl1.GetHorizontalScrollBar().GetScrollBar().GetThumbButton().SetPressedStyleIndex(_T("ThumbButtonH"));
	ActiveGanttVCCtl1.GetHorizontalScrollBar().GetScrollBar().GetThumbButton().SetDisabledStyleIndex(_T("ThumbButtonH"));

	ActiveGanttVCCtl1.GetTimeLineScrollBar().GetScrollBar().SetStyleIndex(_T("ScrollBar"));
	ActiveGanttVCCtl1.GetTimeLineScrollBar().GetScrollBar().GetArrowButtons().SetNormalStyleIndex(_T("ArrowButtons"));
	ActiveGanttVCCtl1.GetTimeLineScrollBar().GetScrollBar().GetArrowButtons().SetPressedStyleIndex(_T("ArrowButtons"));
	ActiveGanttVCCtl1.GetTimeLineScrollBar().GetScrollBar().GetArrowButtons().SetDisabledStyleIndex(_T("ArrowButtons"));
	ActiveGanttVCCtl1.GetTimeLineScrollBar().GetScrollBar().GetThumbButton().SetNormalStyleIndex(_T("ThumbButtonH"));
	ActiveGanttVCCtl1.GetTimeLineScrollBar().GetScrollBar().GetThumbButton().SetPressedStyleIndex(_T("ThumbButtonH"));
	ActiveGanttVCCtl1.GetTimeLineScrollBar().GetScrollBar().GetThumbButton().SetDisabledStyleIndex(_T("ThumbButtonH"));

	ActiveGanttVCCtl1.GetStyle().SetAppearance(SA_FLAT);
	ActiveGanttVCCtl1.GetStyle().SetBorderStyle(SBR_CUSTOM);
	ActiveGanttVCCtl1.GetStyle().SetBorderColor(FromArgb(255, 122, 151, 193));
	ActiveGanttVCCtl1.GetStyle().SetBorderWidth(1);

	ActiveGanttVCCtl1.GetSplitter().SetType(SA_USERDEFINED);
	ActiveGanttVCCtl1.GetSplitter().SetWidth(4);
	ActiveGanttVCCtl1.GetSplitter().SetColor(1, FromArgb(255, 197, 206, 216));
	ActiveGanttVCCtl1.GetSplitter().SetColor(2, CLR_WHITE);
	ActiveGanttVCCtl1.GetSplitter().SetColor(3, CLR_WHITE);
	ActiveGanttVCCtl1.GetSplitter().SetColor(4, FromArgb(255, 197, 206, 216));
	ActiveGanttVCCtl1.GetSplitter().SetPosition(285);

	int i;
	CclsRow oRow;
	for (i = 1; i <= 100; i++)
	{
		oRow = ActiveGanttVCCtl1.GetRows().Add(_T("K") + CStr(i), _T(""), FALSE, TRUE, _T(""));
		oRow.SetHeight(20);
		oRow.GetNode().SetText(_T("Row: ") + CStr(i));
        oRow.GetCells().Item(_T("1")).SetText(CStr(i));
        oRow.GetCells().Item(_T("1")).SetStyleIndex(_T("CellStyleKeyColumn"));
		if (oRow.GetIndex() % 5 == 0 || oRow.GetIndex() == 1)
		{
			oRow.GetNode().SetDepth(0);
		}
		else
		{
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
	oView.GetTimeLine().SetStyleIndex(_T("TimeLine"));

	ActiveGanttVCCtl1.SetCurrentView(_T("1"));

	ActiveGanttVCCtl1.GetTasks().Add(_T(""), _T("K1"), GetDateTime(2013, 3, 1), GetDateTime(2014, 3, 1), _T("SS"), _T("SummaryStyle"), _T("0"));

    ActiveGanttVCCtl1.GetTasks().Add(_T(""), _T("K5"), GetDateTime(2013, 3, 1), GetDateTime(2014, 3, 1), _T("TS"), _T("TaskStyle"), _T("0"));

	return TRUE; 
}


void fSTL_MSP2007::OnSize(UINT nType, int cx, int cy)
{
	CDialog::OnSize(nType, cx, cy);
	g_Resize(this, &ActiveGanttVCCtl1);
}