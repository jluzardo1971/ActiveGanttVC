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
#include "fSTL_CR.h"
#include "afxdialogex.h"

IMPLEMENT_DYNAMIC(fSTL_CR, CDialog)

fSTL_CR::fSTL_CR(CWnd* pParent /*=NULL*/)
	: CDialog(fSTL_CR::IDD, pParent)
{
	mp_sFontName = _T("Tahoma");
}

fSTL_CR::~fSTL_CR()
{
}

void fSTL_CR::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	DDX_Control(pDX, IDC_ACTIVEGANTTVCCTL1, ActiveGanttVCCtl1);
}


BEGIN_MESSAGE_MAP(fSTL_CR, CDialog)
	ON_WM_SIZE()
END_MESSAGE_MAP()


BOOL fSTL_CR::OnInitDialog()
{
	CDialog::OnInitDialog();
	g_MaximizeWindowsDim(this);
	CWnd::ShowWindow(SW_SHOWMAXIMIZED);

	CclsStyle oStyle;
	CclsView oView;
	CclsTimeBlock oTimeBlock;

	oStyle = ActiveGanttVCCtl1.GetStyles().Add(_T("ScrollBar"));
	oStyle.SetAppearance(SA_FLAT);
	oStyle.SetBackgroundMode(FP_SOLID);
	oStyle.SetBackColor(CLR_WHITE);
	oStyle.SetBorderStyle(SBR_SINGLE);
	oStyle.SetBorderColor(FromArgb(255, 150, 158, 168));

	oStyle = ActiveGanttVCCtl1.GetStyles().Add(_T("ArrowButtons"));
	oStyle.SetAppearance(SA_FLAT);
	oStyle.SetBackgroundMode(FP_SOLID);
	oStyle.SetBackColor(CLR_WHITE);
	oStyle.SetBorderStyle(SBR_SINGLE);
	oStyle.SetBorderColor(FromArgb(255, 150, 158, 168));

	oStyle = ActiveGanttVCCtl1.GetStyles().Add(_T("ThumbButton"));
	oStyle.SetAppearance(SA_FLAT);
	oStyle.SetBackgroundMode(FP_SOLID);
	oStyle.SetBackColor(CLR_WHITE);
	oStyle.SetBorderStyle(SBR_SINGLE);
	oStyle.SetBorderColor(FromArgb(255, 138, 145, 153));

	oStyle = ActiveGanttVCCtl1.GetStyles().Add(_T("SplitterStyle"));
	oStyle.SetAppearance(SA_FLAT);
	oStyle.SetBackgroundMode(FP_GRADIENT);
	oStyle.SetGradientFillMode(GDT_VERTICAL);
	oStyle.SetStartGradientColor(FromArgb(255, 109, 122, 136));
	oStyle.SetEndGradientColor(FromArgb(255, 220, 220, 220));

	oStyle = ActiveGanttVCCtl1.GetStyles().Add(_T("Columns"));
	oStyle.GetFont().SetFontName(mp_sFontName);
	oStyle.GetFont().SetSize(8);
	oStyle.GetFont().SetStyle(FS_FONTSTYLEBOLD);
	oStyle.SetAppearance(SA_FLAT);
	oStyle.SetBackgroundMode(FP_GRADIENT);
	oStyle.SetGradientFillMode(GDT_VERTICAL);
	oStyle.SetStartGradientColor(FromArgb(255, 148, 164, 189));
	oStyle.SetEndGradientColor(FromArgb(255, 178, 199, 228));
	oStyle.SetForeColor(CLR_WHITE);
	oStyle.SetBorderColor(CLR_BLACK);
	oStyle.SetBorderStyle(SBR_CUSTOM);
	oStyle.GetCustomBorderStyle().SetLeft(FALSE);
	oStyle.GetCustomBorderStyle().SetTop(FALSE);
	oStyle.SetTextAlignmentVertical(VAL_BOTTOM);

	oStyle = ActiveGanttVCCtl1.GetStyles().Add(_T("TimeLine"));
	oStyle.GetFont().SetFontName(mp_sFontName);
	oStyle.GetFont().SetSize(7);
	oStyle.GetFont().SetStyle(FS_FONTSTYLEREGULAR);
	oStyle.SetAppearance(SA_FLAT);
	oStyle.SetBackgroundMode(FP_GRADIENT);
	oStyle.SetGradientFillMode(GDT_VERTICAL);
	oStyle.SetStartGradientColor(FromArgb(255, 148, 164, 189));
	oStyle.SetEndGradientColor(FromArgb(255, 178, 199, 228));
	oStyle.SetForeColor(CLR_WHITE);
	oStyle.SetBorderColor(CLR_BLACK);
	oStyle.GetCustomBorderStyle().SetLeft(TRUE);
	oStyle.GetCustomBorderStyle().SetTop(TRUE);
	oStyle.GetCustomBorderStyle().SetRight(FALSE);
	oStyle.GetCustomBorderStyle().SetBottom(TRUE);
	oStyle.SetBorderStyle(SBR_CUSTOM);

	oStyle = ActiveGanttVCCtl1.GetStyles().Add(_T("TimeLineVA"));
	oStyle.GetFont().SetFontName(mp_sFontName);
	oStyle.GetFont().SetSize(7);
	oStyle.GetFont().SetStyle(FS_FONTSTYLEREGULAR);
	oStyle.SetAppearance(SA_FLAT);
	oStyle.SetBackgroundMode(FP_GRADIENT);
	oStyle.SetGradientFillMode(GDT_VERTICAL);
	oStyle.SetStartGradientColor(FromArgb(255, 148, 164, 189));
	oStyle.SetEndGradientColor(FromArgb(255, 178, 199, 228));
	oStyle.SetForeColor(CLR_WHITE);
	oStyle.SetBorderColor(CLR_BLACK);
	oStyle.SetBorderStyle(SBR_SINGLE);
	oStyle.SetDrawTextInVisibleArea(TRUE);

	oStyle = ActiveGanttVCCtl1.GetStyles().Add(_T("Branch"));
	oStyle.GetFont().SetFontName(mp_sFontName);
	oStyle.GetFont().SetSize(9);
	oStyle.GetFont().SetStyle(FS_FONTSTYLEREGULAR);
	oStyle.SetAppearance(SA_FLAT);
	oStyle.SetBackgroundMode(FP_GRADIENT);
	oStyle.SetGradientFillMode(GDT_VERTICAL);
	oStyle.SetStartGradientColor(FromArgb(255, 109, 122, 136));
	oStyle.SetEndGradientColor(FromArgb(255, 179, 199, 229));
	oStyle.SetTextAlignmentHorizontal(HAL_LEFT);
	oStyle.SetTextAlignmentVertical(VAL_TOP);
	oStyle.SetTextXMargin(5);
	oStyle.SetTextYMargin(5);
	oStyle.SetForeColor(CLR_WHITE);
	oStyle.SetBorderColor(CLR_BLACK);
	oStyle.SetBorderStyle(SBR_SINGLE);
	oStyle.SetImageAlignmentHorizontal(HAL_RIGHT);
	oStyle.SetImageAlignmentVertical(VAL_BOTTOM);
	oStyle.SetImageXMargin(5);
	oStyle.SetImageYMargin(5);
	oStyle.SetUseMask(FALSE);

	oStyle = ActiveGanttVCCtl1.GetStyles().Add(_T("BranchCA"));
	oStyle.SetAppearance(SA_FLAT);
	oStyle.SetBackgroundMode(FP_GRADIENT);
	oStyle.SetGradientFillMode(GDT_VERTICAL);
	oStyle.SetStartGradientColor(FromArgb(255, 109, 122, 136));
	oStyle.SetEndGradientColor(FromArgb(255, 179, 199, 229));
	oStyle.SetForeColor(CLR_WHITE);

	oStyle = ActiveGanttVCCtl1.GetStyles().Add(_T("Weekend"));
	oStyle.SetAppearance(SA_FLAT);
	oStyle.SetBackgroundMode(FP_GRADIENT);
	oStyle.SetGradientFillMode(GDT_HORIZONTAL);
	oStyle.SetStartGradientColor(FromArgb(255, 133, 143, 154));
	oStyle.SetEndGradientColor(FromArgb(255, 172, 183, 194));

	oStyle = ActiveGanttVCCtl1.GetStyles().Add(_T("Reservation"));
	oStyle.GetFont().SetFontName(mp_sFontName);
	oStyle.GetFont().SetSize(7);
	oStyle.GetFont().SetStyle(FS_FONTSTYLEREGULAR);
	oStyle.SetForeColor(CLR_WHITE);
	oStyle.SetTextAlignmentHorizontal(HAL_LEFT);
	oStyle.SetTextAlignmentVertical(VAL_TOP);
	oStyle.SetTextXMargin(5);
	oStyle.GetSelectionRectangleStyle().SetOffsetTop(0);
	oStyle.GetSelectionRectangleStyle().SetOffsetBottom(0);
	oStyle.GetSelectionRectangleStyle().SetOffsetLeft(0);
	oStyle.GetSelectionRectangleStyle().SetOffsetRight(0);
	oStyle.SetAppearance(SA_FLAT);
	oStyle.SetBackgroundMode(FP_GRADIENT);
	oStyle.SetGradientFillMode(GDT_HORIZONTAL);
	oStyle.SetStartGradientColor(FromArgb(255, 109, 122, 136));
	oStyle.SetEndGradientColor(FromArgb(255, 179, 199, 229));

	oStyle = ActiveGanttVCCtl1.GetStyles().Add(_T("Rental"));
	oStyle.GetFont().SetFontName(mp_sFontName);
	oStyle.GetFont().SetSize(7);
	oStyle.GetFont().SetStyle(FS_FONTSTYLEREGULAR);
	oStyle.SetForeColor(CLR_WHITE);
	oStyle.SetTextAlignmentHorizontal(HAL_LEFT);
	oStyle.SetTextAlignmentVertical(VAL_TOP);
	oStyle.SetTextXMargin(5);
	oStyle.GetSelectionRectangleStyle().SetOffsetTop(0);
	oStyle.GetSelectionRectangleStyle().SetOffsetBottom(0);
	oStyle.GetSelectionRectangleStyle().SetOffsetLeft(0);
	oStyle.GetSelectionRectangleStyle().SetOffsetRight(0);
	oStyle.SetAppearance(SA_FLAT);
	oStyle.SetBackgroundMode(FP_GRADIENT);
	oStyle.SetGradientFillMode(GDT_HORIZONTAL);
	oStyle.SetStartGradientColor(FromArgb(255, 162, 78, 50));
	oStyle.SetEndGradientColor(FromArgb(255, 215, 92, 54));

	oStyle = ActiveGanttVCCtl1.GetStyles().Add(_T("Maintenance"));
	oStyle.GetFont().SetFontName(mp_sFontName);
	oStyle.GetFont().SetSize(7);
	oStyle.GetFont().SetStyle(FS_FONTSTYLEREGULAR);
	oStyle.SetForeColor(CLR_WHITE);
	oStyle.SetTextAlignmentHorizontal(HAL_LEFT);
	oStyle.SetTextAlignmentVertical(VAL_TOP);
	oStyle.SetTextXMargin(5);
	oStyle.GetSelectionRectangleStyle().SetOffsetTop(0);
	oStyle.GetSelectionRectangleStyle().SetOffsetBottom(0);
	oStyle.GetSelectionRectangleStyle().SetOffsetLeft(0);
	oStyle.GetSelectionRectangleStyle().SetOffsetRight(0);
	oStyle.SetAppearance(SA_FLAT);
	oStyle.SetBackgroundMode(FP_GRADIENT);
	oStyle.SetGradientFillMode(GDT_HORIZONTAL);
	oStyle.SetStartGradientColor(FromArgb(255, 255, 77, 1));
	oStyle.SetEndGradientColor(FromArgb(255, 244, 172, 43));

	ActiveGanttVCCtl1.GetColumns().Add(_T(""), _T(""), 45, _T("Columns"));
	ActiveGanttVCCtl1.GetColumns().Add(_T(""), _T(""), 95, _T("Columns"));
	ActiveGanttVCCtl1.GetColumns().Add(_T(""), _T(""), 250, _T("Columns"));

	ActiveGanttVCCtl1.GetSplitter().SetPosition(340);
	ActiveGanttVCCtl1.GetSplitter().SetType(SA_STYLE);
	ActiveGanttVCCtl1.GetSplitter().SetWidth(6);
	ActiveGanttVCCtl1.GetSplitter().SetStyleIndex(_T("SplitterStyle"));

	ActiveGanttVCCtl1.GetScrollBarSeparator().SetStyleIndex(_T("ScrollBar"));

	ActiveGanttVCCtl1.GetVerticalScrollBar().GetScrollBar().SetStyleIndex(_T("ScrollBar"));
	ActiveGanttVCCtl1.GetVerticalScrollBar().GetScrollBar().GetArrowButtons().SetNormalStyleIndex(_T("ArrowButtons"));
	ActiveGanttVCCtl1.GetVerticalScrollBar().GetScrollBar().GetArrowButtons().SetPressedStyleIndex(_T("ArrowButtons"));
	ActiveGanttVCCtl1.GetVerticalScrollBar().GetScrollBar().GetArrowButtons().SetDisabledStyleIndex(_T("ArrowButtons"));
	ActiveGanttVCCtl1.GetVerticalScrollBar().GetScrollBar().GetThumbButton().SetNormalStyleIndex(_T("ThumbButton"));
	ActiveGanttVCCtl1.GetVerticalScrollBar().GetScrollBar().GetThumbButton().SetPressedStyleIndex(_T("ThumbButton"));
	ActiveGanttVCCtl1.GetVerticalScrollBar().GetScrollBar().GetThumbButton().SetDisabledStyleIndex(_T("ThumbButton"));

	ActiveGanttVCCtl1.GetHorizontalScrollBar().GetScrollBar().SetStyleIndex(_T("ScrollBar"));
	ActiveGanttVCCtl1.GetHorizontalScrollBar().GetScrollBar().GetArrowButtons().SetNormalStyleIndex(_T("ArrowButtons"));
	ActiveGanttVCCtl1.GetHorizontalScrollBar().GetScrollBar().GetArrowButtons().SetPressedStyleIndex(_T("ArrowButtons"));
	ActiveGanttVCCtl1.GetHorizontalScrollBar().GetScrollBar().GetArrowButtons().SetDisabledStyleIndex(_T("ArrowButtons"));
	ActiveGanttVCCtl1.GetHorizontalScrollBar().GetScrollBar().GetThumbButton().SetNormalStyleIndex(_T("ThumbButton"));
	ActiveGanttVCCtl1.GetHorizontalScrollBar().GetScrollBar().GetThumbButton().SetPressedStyleIndex(_T("ThumbButton"));
	ActiveGanttVCCtl1.GetHorizontalScrollBar().GetScrollBar().GetThumbButton().SetDisabledStyleIndex(_T("ThumbButton"));

	oTimeBlock = ActiveGanttVCCtl1.GetTimeBlocks().Add(_T(""));
    oTimeBlock.SetTimeBlockType(TBT_RECURRING);
    oTimeBlock.SetRecurringType(RCT_WEEK);
    oTimeBlock.SetBaseWeekDay(WD_SATURDAY);
	COleDateTime dtBaseDate;
	dtBaseDate.SetDateTime(2013, 1, 1, 0, 0, 0);
	oTimeBlock.SetBaseDate(dtBaseDate);
    oTimeBlock.SetDurationFactor(48);
    oTimeBlock.SetDurationInterval(IL_HOUR);
    oTimeBlock.SetStyleIndex(_T("Weekend"));

	oView = ActiveGanttVCCtl1.GetViews().Add(IL_MINUTE, 30, ST_MONTH, ST_DAY, ST_NOT_VISIBLE, _T(""));
	oView.GetTimeLine().GetTierArea().GetUpperTier().SetHeight(17);
    oView.GetTimeLine().GetTierArea().GetUpperTier().SetBackgroundMode(ETBGM_STYLE);
    oView.GetTimeLine().GetTierArea().GetUpperTier().SetStyleIndex(_T("TimeLineVA"));
	oView.GetTimeLine().GetTierArea().GetMiddleTier().SetHeight(17);
    oView.GetTimeLine().GetTierArea().GetMiddleTier().SetBackgroundMode(ETBGM_STYLE);
    oView.GetTimeLine().GetTierArea().GetMiddleTier().SetStyleIndex(_T("TimeLineVA"));
	oView.GetTimeLine().GetTierArea().GetLowerTier().SetHeight(17);
	oView.GetTimeLine().GetTierArea().GetLowerTier().SetBackgroundMode(ETBGM_STYLE);
    oView.GetTimeLine().GetTierArea().GetLowerTier().SetStyleIndex(_T("TimeLineVA"));
	oView.GetTimeLine().GetTickMarkArea().SetVisible(FALSE);
	oView.GetTimeLine().GetTickMarkArea().SetStyleIndex(_T("TimeLine"));
	oView.GetTimeLine().GetStyle().SetAppearance(SA_FLAT);
	oView.GetTimeLine().GetStyle().SetBackgroundMode(FP_SOLID);
	oView.GetTimeLine().GetStyle().SetBackColor(CLR_BLACK);
	oView.GetClientArea().GetGrid().SetVerticalLines(TRUE);
	oView.GetClientArea().GetGrid().SetSnapToGrid(TRUE);

	ActiveGanttVCCtl1.SetCurrentView(_T("1"));

	int i;
	CclsRow oRow;
	for (i = 1; i <= 100; i++)
	{
		oRow = ActiveGanttVCCtl1.GetRows().Add(_T("K") + CStr(i), _T(""), FALSE, TRUE, _T(""));
		if (oRow.GetIndex() % 5 == 0 || oRow.GetIndex() == 1)
		{
			oRow.GetNode().SetDepth(0);
			oRow.SetMergeCells(TRUE);
			oRow.SetStyleIndex(_T("Branch"));
			oRow.SetClientAreaStyleIndex(_T("BranchCA"));
		}
		else
		{
			oRow.GetNode().SetDepth(1);
		}
	}

	ActiveGanttVCCtl1.GetTasks().Add(_T("Reservation"), _T("K2"), GetDateTime(2013, 1, 2, 0, 0, 0), GetDateTime(2013, 1, 7, 0, 0, 0), _T("RES"), _T("Reservation"), _T("0"));

    ActiveGanttVCCtl1.GetTasks().Add(_T("Rental"), _T("K3"), GetDateTime(2013, 1, 2, 0, 0, 0), GetDateTime(2013, 1, 7, 0, 0, 0), _T("REN"), _T("Rental"), _T("0"));

    ActiveGanttVCCtl1.GetTasks().Add(_T("Maintenance"), _T("K4"), GetDateTime(2013, 1, 2, 0, 0, 0), GetDateTime(2013, 1, 7, 0, 0, 0), _T("MAI"), _T("Maintenance"), _T("0"));

    ActiveGanttVCCtl1.GetCurrentViewObject().GetTimeLine().Position(GetDateTime(2013, 1, 1, 0, 0, 0));

	return TRUE;
}


void fSTL_CR::OnSize(UINT nType, int cx, int cy)
{
	CDialog::OnSize(nType, cx, cy);
	g_Resize(this, &ActiveGanttVCCtl1);
}