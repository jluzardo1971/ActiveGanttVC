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
#include "fSTL_Anakiwa_Blue.h"
#include "afxdialogex.h"

IMPLEMENT_DYNAMIC(fSTL_Anakiwa_Blue, CDialog)

fSTL_Anakiwa_Blue::fSTL_Anakiwa_Blue(CWnd* pParent /*=NULL*/)
	: CDialog(fSTL_Anakiwa_Blue::IDD, pParent)
{
	mp_sFontName = _T("Tahoma");
}

fSTL_Anakiwa_Blue::~fSTL_Anakiwa_Blue()
{
}

void fSTL_Anakiwa_Blue::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	DDX_Control(pDX, IDC_ACTIVEGANTTVCCTL1, ActiveGanttVCCtl1);
}


BEGIN_MESSAGE_MAP(fSTL_Anakiwa_Blue, CDialog)
	ON_WM_SIZE()
END_MESSAGE_MAP()

BOOL fSTL_Anakiwa_Blue::OnInitDialog()
{
	CDialog::OnInitDialog();
	g_MaximizeWindowsDim(this);
	CWnd::ShowWindow(SW_SHOWMAXIMIZED);

	CclsStyle oStyle;
	CclsView oView;

	oStyle = ActiveGanttVCCtl1.GetStyles().Add(_T("ControlStyle"));
	oStyle.SetAppearance(SA_FLAT);
	oStyle.SetBorderStyle(SBR_CUSTOM);
	oStyle.SetBorderColor(FromArgb(255, 100, 145, 204));
	oStyle.SetBackColor(FromArgb(255, 240, 240, 240));

	oStyle = ActiveGanttVCCtl1.GetStyles().Add(_T("ScrollBar"));
	oStyle.SetAppearance(SA_FLAT);
	oStyle.SetBackgroundMode(FP_SOLID);
	oStyle.SetBackColor(CLR_WHITE);
	oStyle.SetBorderStyle(SBR_SINGLE);
	oStyle.SetBorderColor(FromArgb(255, 150, 150, 150));

	oStyle = ActiveGanttVCCtl1.GetStyles().Add(_T("ArrowButtons"));
	oStyle.SetAppearance(SA_FLAT);
	oStyle.SetBackgroundMode(FP_SOLID);
	oStyle.SetBackColor(CLR_WHITE);
	oStyle.SetBorderStyle(SBR_SINGLE);
	oStyle.SetBorderColor(FromArgb(255, 150, 150, 150));

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

	oStyle = ActiveGanttVCCtl1.GetStyles().Add(_T("ThumbButtonHP"));
	oStyle.SetAppearance(SA_FLAT);
	oStyle.SetBackgroundMode(FP_GRADIENT);
	oStyle.SetGradientFillMode(GDT_VERTICAL);
	oStyle.SetStartGradientColor(FromArgb(255, 165, 186, 207));
	oStyle.SetEndGradientColor(FromArgb(255, 240, 240, 240));
	oStyle.SetBorderStyle(SBR_SINGLE);
	oStyle.SetBorderColor(FromArgb(255, 138, 145, 153));

	oStyle = ActiveGanttVCCtl1.GetStyles().Add(_T("ThumbButtonVP"));
	oStyle.SetAppearance(SA_FLAT);
	oStyle.SetBackgroundMode(FP_GRADIENT);
	oStyle.SetGradientFillMode(GDT_HORIZONTAL);
	oStyle.SetStartGradientColor(FromArgb(255, 165, 186, 207));
	oStyle.SetEndGradientColor(FromArgb(255, 240, 240, 240));
	oStyle.SetBorderStyle(SBR_SINGLE);
	oStyle.SetBorderColor(FromArgb(255, 138, 145, 153));

	oStyle = ActiveGanttVCCtl1.GetStyles().Add(_T("ColumnStyle"));
	oStyle.SetAppearance(SA_FLAT);
	oStyle.SetBackgroundMode(FP_GRADIENT);
	oStyle.SetStartGradientColor(FromArgb(255, 179, 206, 235));
	oStyle.SetEndGradientColor(FromArgb(255, 161, 193, 232));
	oStyle.SetGradientFillMode(GDT_VERTICAL);
	oStyle.GetCustomBorderStyle().SetLeft(FALSE);
	oStyle.GetCustomBorderStyle().SetTop(FALSE);
	oStyle.GetCustomBorderStyle().SetRight(TRUE);
	oStyle.GetCustomBorderStyle().SetBottom(TRUE);
	oStyle.SetBorderStyle(SBR_CUSTOM);
	oStyle.SetBorderColor(FromArgb(255, 100, 145, 204));

	oStyle = ActiveGanttVCCtl1.GetStyles().Add(_T("ScrollBarSeparatorStyle"));
	oStyle.SetAppearance(SA_FLAT);
	oStyle.SetBackgroundMode(FP_SOLID);
	oStyle.SetBackColor(CLR_WHITE);
	oStyle.SetBorderStyle(SBR_SINGLE);
	oStyle.SetBorderColor(FromArgb(255, 150, 150, 150));

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

	oStyle = ActiveGanttVCCtl1.GetStyles().Add(_T("NodeRegular"));
	oStyle.GetFont().SetFontName(mp_sFontName);
	oStyle.GetFont().SetSize(8);
	oStyle.GetFont().SetStyle(FS_FONTSTYLEREGULAR);
	oStyle.SetAppearance(SA_FLAT);
	oStyle.SetBackgroundMode(FP_SOLID);
	oStyle.SetBackColor(CLR_WHITE);
    oStyle.SetBorderColor(FromArgb(255, 192, 192, 192));
    oStyle.SetBorderStyle(SBR_CUSTOM);
    oStyle.GetCustomBorderStyle().SetTop(FALSE);
    oStyle.GetCustomBorderStyle().SetLeft(FALSE);

	oStyle = ActiveGanttVCCtl1.GetStyles().Add(_T("NodeBold"));
	oStyle.GetFont().SetFontName(mp_sFontName);
	oStyle.GetFont().SetSize(8);
	oStyle.GetFont().SetStyle(FS_FONTSTYLEBOLD);
	oStyle.SetAppearance(SA_FLAT);
	oStyle.SetBackgroundMode(FP_SOLID);
	oStyle.SetBackColor(CLR_WHITE);
	oStyle.SetBorderColor(FromArgb(255, 192, 192, 192));
    oStyle.SetBorderStyle(SBR_CUSTOM);
    oStyle.GetCustomBorderStyle().SetTop(FALSE);
    oStyle.GetCustomBorderStyle().SetLeft(FALSE);

	oStyle = ActiveGanttVCCtl1.GetStyles().Add(_T("NormalTask"));
	oStyle.SetAppearance(SA_FLAT);
	oStyle.SetPlacement(PLC_OFFSETPLACEMENT);
	oStyle.SetBackColor(FromArgb(255, 100, 145, 204));
	oStyle.SetBorderColor(CLR_BLUE);
	oStyle.SetBorderStyle(SBR_SINGLE);
	oStyle.GetSelectionRectangleStyle().SetOffsetTop(0);
	oStyle.GetSelectionRectangleStyle().SetOffsetLeft(0);
	oStyle.GetSelectionRectangleStyle().SetOffsetRight(0);
	oStyle.GetSelectionRectangleStyle().SetOffsetBottom(0);
	oStyle.SetOffsetTop(5);
	oStyle.SetOffsetBottom(10);
	oStyle.SetBackgroundMode(FP_GRADIENT);
	oStyle.SetStartGradientColor(CLR_WHITE);
	oStyle.SetEndGradientColor(FromArgb(255, 100, 145, 204));
	oStyle.SetGradientFillMode(GDT_VERTICAL);
	oStyle.GetPredecessorStyle().SetLineColor(FromArgb(255, 100, 145, 204));
	oStyle.GetMilestoneStyle().SetShapeIndex(FT_DIAMOND);

	oStyle = ActiveGanttVCCtl1.GetStyles().Add(_T("GreenSummary"));
	oStyle.SetAppearance(SA_FLAT);
	oStyle.SetPlacement(PLC_OFFSETPLACEMENT);
	oStyle.SetBackColor(CLR_GREEN);
	oStyle.SetBorderColor(CLR_GREEN);
	oStyle.SetBorderStyle(SBR_SINGLE);
	oStyle.GetSelectionRectangleStyle().SetOffsetTop(0);
	oStyle.GetSelectionRectangleStyle().SetOffsetLeft(0);
	oStyle.GetSelectionRectangleStyle().SetOffsetRight(0);
	oStyle.GetSelectionRectangleStyle().SetOffsetBottom(0);
	oStyle.SetOffsetTop(5);
	oStyle.SetOffsetBottom(10);
	oStyle.SetBackgroundMode(FP_GRADIENT);
	oStyle.SetStartGradientColor(CLR_WHITE);
	oStyle.SetEndGradientColor(CLR_GREEN);
	oStyle.SetGradientFillMode(GDT_VERTICAL);
	oStyle.GetMilestoneStyle().SetShapeIndex(FT_DIAMOND);
	oStyle.GetSelectionRectangleStyle().SetVisible(FALSE);
	oStyle.GetTaskStyle().SetEndFillColor(CLR_GREEN);
	oStyle.GetTaskStyle().SetEndBorderColor(CLR_GREEN);
	oStyle.GetTaskStyle().SetStartFillColor(CLR_GREEN);
	oStyle.GetTaskStyle().SetStartBorderColor(CLR_GREEN);
	oStyle.GetTaskStyle().SetStartShapeIndex(FT_PROJECTDOWN);
	oStyle.GetTaskStyle().SetEndShapeIndex(FT_PROJECTDOWN);
	oStyle.SetFillMode(FM_UPPERHALFFILLED);

	oStyle = ActiveGanttVCCtl1.GetStyles().Add(_T("RedSummary"));
	oStyle.SetAppearance(SA_FLAT);
	oStyle.SetPlacement(PLC_OFFSETPLACEMENT);
	oStyle.SetBackColor(CLR_RED);
	oStyle.SetBorderColor(CLR_RED);
	oStyle.SetBorderStyle(SBR_SINGLE);
	oStyle.SetOffsetTop(5);
	oStyle.SetOffsetBottom(10);
	oStyle.SetBackgroundMode(FP_GRADIENT);
	oStyle.SetStartGradientColor(CLR_WHITE);
	oStyle.SetEndGradientColor(CLR_RED);
	oStyle.SetGradientFillMode(GDT_VERTICAL);
	oStyle.GetMilestoneStyle().SetShapeIndex(FT_DIAMOND);
	oStyle.GetSelectionRectangleStyle().SetVisible(FALSE);
	oStyle.GetTaskStyle().SetEndFillColor(CLR_RED);
	oStyle.GetTaskStyle().SetEndBorderColor(CLR_RED);
	oStyle.GetTaskStyle().SetStartFillColor(CLR_RED);
	oStyle.GetTaskStyle().SetStartBorderColor(CLR_RED);
	oStyle.GetTaskStyle().SetStartShapeIndex(FT_PROJECTDOWN);
	oStyle.GetTaskStyle().SetEndShapeIndex(FT_PROJECTDOWN);
	oStyle.SetFillMode(FM_UPPERHALFFILLED);

	oStyle = ActiveGanttVCCtl1.GetStyles().Add(_T("BluePercentages"));
	oStyle.SetAppearance(SA_FLAT);
	oStyle.SetPlacement(PLC_OFFSETPLACEMENT);
	oStyle.SetBackColor(CLR_BLUE);
	oStyle.SetBorderColor(CLR_BLUE);
	oStyle.SetBorderStyle(SBR_SINGLE);
	oStyle.GetSelectionRectangleStyle().SetOffsetTop(0);
	oStyle.GetSelectionRectangleStyle().SetOffsetLeft(0);
	oStyle.GetSelectionRectangleStyle().SetOffsetRight(0);
	oStyle.GetSelectionRectangleStyle().SetOffsetBottom(0);
	oStyle.SetOffsetTop(8);
	oStyle.SetOffsetBottom(4);
	oStyle.GetSelectionRectangleStyle().SetVisible(TRUE);
	oStyle.SetTextVisible(FALSE);
	oStyle.SetBackgroundMode(FP_SOLID);

	oStyle = ActiveGanttVCCtl1.GetStyles().Add(_T("GreenPercentages"));
	oStyle.SetAppearance(SA_FLAT);
	oStyle.SetPlacement(PLC_OFFSETPLACEMENT);
	oStyle.SetBackColor(CLR_GREEN);
	oStyle.SetBorderColor(CLR_GREEN);
	oStyle.SetBorderStyle(SBR_SINGLE);
	oStyle.GetSelectionRectangleStyle().SetOffsetTop(0);
	oStyle.GetSelectionRectangleStyle().SetOffsetLeft(0);
	oStyle.GetSelectionRectangleStyle().SetOffsetRight(0);
	oStyle.GetSelectionRectangleStyle().SetOffsetBottom(0);
	oStyle.SetOffsetTop(5);
	oStyle.SetOffsetBottom(5);
	oStyle.GetSelectionRectangleStyle().SetVisible(FALSE);
	oStyle.SetTextVisible(FALSE);
	oStyle.SetBackgroundMode(FP_SOLID);

	oStyle = ActiveGanttVCCtl1.GetStyles().Add(_T("RedPercentages"));
	oStyle.SetAppearance(SA_FLAT);
	oStyle.SetPlacement(PLC_OFFSETPLACEMENT);
	oStyle.SetBackColor(CLR_RED);
	oStyle.SetBorderColor(CLR_RED);
	oStyle.SetBorderStyle(SBR_SINGLE);
	oStyle.GetSelectionRectangleStyle().SetOffsetTop(0);
	oStyle.GetSelectionRectangleStyle().SetOffsetLeft(0);
	oStyle.GetSelectionRectangleStyle().SetOffsetRight(0);
	oStyle.GetSelectionRectangleStyle().SetOffsetBottom(0);
	oStyle.SetOffsetTop(5);
	oStyle.SetOffsetBottom(5);
	oStyle.GetSelectionRectangleStyle().SetVisible(FALSE);
	oStyle.SetTextVisible(FALSE);
	oStyle.SetBackgroundMode(FP_SOLID);

	oStyle = ActiveGanttVCCtl1.GetStyles().Add(_T("ClientAreaStyle"));
	oStyle.SetAppearance(SA_FLAT);
	oStyle.SetBackColor(CLR_WHITE);
	oStyle.SetBorderColor(FromArgb(255, 192, 192, 192));
	oStyle.SetBorderStyle(SBR_CUSTOM);
    oStyle.GetCustomBorderStyle().SetTop(FALSE);
    oStyle.GetCustomBorderStyle().SetLeft(FALSE);
    oStyle.GetCustomBorderStyle().SetRight(FALSE);

	oStyle = ActiveGanttVCCtl1.GetStyles().Add(_T("CellStyleKeyColumn"));
	oStyle.SetAppearance(SA_FLAT);
	oStyle.SetBackColor(CLR_WHITE);
    oStyle.SetBorderColor(FromArgb(255, 192, 192, 192));
    oStyle.SetBorderStyle(SBR_CUSTOM);
    oStyle.GetCustomBorderStyle().SetTop(FALSE);
    oStyle.GetCustomBorderStyle().SetLeft(FALSE);
	oStyle.SetTextAlignmentHorizontal(HAL_RIGHT);
	oStyle.SetTextXMargin(4);           

	oStyle = ActiveGanttVCCtl1.GetStyles().Add(_T("CellStyle"));
	oStyle.SetAppearance(SA_FLAT);
	oStyle.SetBackColor(CLR_WHITE);
    oStyle.SetBorderColor(FromArgb(255, 192, 192, 192));
    oStyle.SetBorderStyle(SBR_CUSTOM);
    oStyle.GetCustomBorderStyle().SetTop(FALSE);
    oStyle.GetCustomBorderStyle().SetLeft(FALSE);

	ActiveGanttVCCtl1.SetStyleIndex(_T("ControlStyle"));
	ActiveGanttVCCtl1.GetScrollBarSeparator().SetStyleIndex(_T("ScrollBarSeparatorStyle"));

	CclsColumn oColumn;

	oColumn = ActiveGanttVCCtl1.GetColumns().Add(_T("ID"), _T(""), 30, _T(""));
    oColumn.SetStyleIndex(_T("ColumnStyle"));
    oColumn.SetAllowTextEdit(TRUE);

	oColumn = ActiveGanttVCCtl1.GetColumns().Add(_T("Task Name"), _T(""), 300, _T(""));
    oColumn.SetStyleIndex(_T("ColumnStyle"));
    oColumn.SetAllowTextEdit(TRUE);

	oColumn = ActiveGanttVCCtl1.GetColumns().Add(_T("StartDate"), _T(""), 125, _T(""));
    oColumn.SetStyleIndex(_T("ColumnStyle"));
    oColumn.SetAllowTextEdit(TRUE);

	oColumn = ActiveGanttVCCtl1.GetColumns().Add(_T("EndDate"), _T(""), 125, _T(""));
    oColumn.SetStyleIndex(_T("ColumnStyle"));
    oColumn.SetAllowTextEdit(TRUE);

	ActiveGanttVCCtl1.SetTreeviewColumnIndex(2);

	ActiveGanttVCCtl1.GetTreeview().SetImages(TRUE);
	ActiveGanttVCCtl1.GetTreeview().SetCheckBoxes(TRUE);
	ActiveGanttVCCtl1.GetTreeview().SetFullColumnSelect(TRUE);
	ActiveGanttVCCtl1.GetTreeview().SetPlusMinusBorderColor(FromArgb(255, 100, 145, 204));
	ActiveGanttVCCtl1.GetTreeview().SetPlusMinusSignColor(FromArgb(255, 100, 145, 204));
	ActiveGanttVCCtl1.GetTreeview().SetCheckBoxBorderColor(FromArgb(255, 100, 145, 204));
	ActiveGanttVCCtl1.GetTreeview().SetTreeLineColor(FromArgb(255, 100, 145, 204));

	ActiveGanttVCCtl1.GetToolTip().GetFont().SetFontName(mp_sFontName);
	ActiveGanttVCCtl1.GetToolTip().GetFont().SetSize(8);
	ActiveGanttVCCtl1.GetToolTip().GetFont().SetStyle(FS_FONTSTYLEREGULAR);

	ActiveGanttVCCtl1.GetSplitter().SetType(SA_USERDEFINED);
	ActiveGanttVCCtl1.GetSplitter().SetWidth(1);
	ActiveGanttVCCtl1.GetSplitter().SetColor(1, FromArgb(255, 100, 145, 204));
	ActiveGanttVCCtl1.GetSplitter().SetPosition(255);

	ActiveGanttVCCtl1.GetVerticalScrollBar().GetScrollBar().SetStyleIndex(_T("ScrollBar"));
	ActiveGanttVCCtl1.GetVerticalScrollBar().GetScrollBar().GetArrowButtons().SetNormalStyleIndex(_T("ArrowButtons"));
	ActiveGanttVCCtl1.GetVerticalScrollBar().GetScrollBar().GetArrowButtons().SetPressedStyleIndex(_T("ArrowButtons"));
	ActiveGanttVCCtl1.GetVerticalScrollBar().GetScrollBar().GetArrowButtons().SetDisabledStyleIndex(_T("ArrowButtons"));
	ActiveGanttVCCtl1.GetVerticalScrollBar().GetScrollBar().GetThumbButton().SetNormalStyleIndex(_T("ThumbButtonV"));
	ActiveGanttVCCtl1.GetVerticalScrollBar().GetScrollBar().GetThumbButton().SetPressedStyleIndex(_T("ThumbButtonVP"));
	ActiveGanttVCCtl1.GetVerticalScrollBar().GetScrollBar().GetThumbButton().SetDisabledStyleIndex(_T("ThumbButtonV"));

	ActiveGanttVCCtl1.GetHorizontalScrollBar().GetScrollBar().SetStyleIndex(_T("ScrollBar"));
	ActiveGanttVCCtl1.GetHorizontalScrollBar().GetScrollBar().GetArrowButtons().SetNormalStyleIndex(_T("ArrowButtons"));
	ActiveGanttVCCtl1.GetHorizontalScrollBar().GetScrollBar().GetArrowButtons().SetPressedStyleIndex(_T("ArrowButtons"));
	ActiveGanttVCCtl1.GetHorizontalScrollBar().GetScrollBar().GetArrowButtons().SetDisabledStyleIndex(_T("ArrowButtons"));
	ActiveGanttVCCtl1.GetHorizontalScrollBar().GetScrollBar().GetThumbButton().SetNormalStyleIndex(_T("ThumbButtonH"));
	ActiveGanttVCCtl1.GetHorizontalScrollBar().GetScrollBar().GetThumbButton().SetPressedStyleIndex(_T("ThumbButtonHP"));
	ActiveGanttVCCtl1.GetHorizontalScrollBar().GetScrollBar().GetThumbButton().SetDisabledStyleIndex(_T("ThumbButtonH"));

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
		if (i == 1)
		{
			oRow.GetNode().GetImage().FromFile(g_GetAppLocation() + _T("\\res\\modules.bmp"), -1);
		}
		if (oRow.GetIndex() % 5 == 0 || oRow.GetIndex() == 1)
		{
			oRow.GetNode().SetStyleIndex(_T("NodeBold"));
			oRow.GetNode().SetDepth(0);
		}
		else
		{
            oRow.GetNode().SetStyleIndex(_T("NodeRegular"));
            oRow.GetNode().SetCheckBoxVisible(TRUE);
			oRow.GetNode().SetDepth(1);
		}
		oRow.SetClientAreaStyleIndex(_T("ClientAreaStyle"));
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
	ActiveGanttVCCtl1.GetTimeLineScrollBar().GetScrollBar().SetStyleIndex(_T("ScrollBar"));
	ActiveGanttVCCtl1.GetTimeLineScrollBar().GetScrollBar().GetArrowButtons().SetNormalStyleIndex(_T("ArrowButtons"));
	ActiveGanttVCCtl1.GetTimeLineScrollBar().GetScrollBar().GetArrowButtons().SetPressedStyleIndex(_T("ArrowButtons"));
	ActiveGanttVCCtl1.GetTimeLineScrollBar().GetScrollBar().GetArrowButtons().SetDisabledStyleIndex(_T("ArrowButtons"));
	ActiveGanttVCCtl1.GetTimeLineScrollBar().GetScrollBar().GetThumbButton().SetNormalStyleIndex(_T("ThumbButtonH"));
	ActiveGanttVCCtl1.GetTimeLineScrollBar().GetScrollBar().GetThumbButton().SetPressedStyleIndex(_T("ThumbButtonHP"));
	ActiveGanttVCCtl1.GetTimeLineScrollBar().GetScrollBar().GetThumbButton().SetDisabledStyleIndex(_T("ThumbButtonH"));

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

	ActiveGanttVCCtl1.GetTasks().Add(_T(""), _T("K1"), GetDateTime(2013, 3, 1), GetDateTime(2014, 3, 1), _T("RS"), _T("RedSummary"), _T("0"));

    ActiveGanttVCCtl1.GetTasks().Add(_T(""), _T("K5"), GetDateTime(2013, 3, 1), GetDateTime(2014, 3, 1), _T("GS"), _T("GreenSummary"), _T("0"));

	ActiveGanttVCCtl1.GetTasks().Add(_T(""), _T("K11"), GetDateTime(2013, 3, 1), GetDateTime(2014, 3, 1), _T("NT"), _T("NormalTask"), _T("0"));

	return TRUE; 
}


void fSTL_Anakiwa_Blue::OnSize(UINT nType, int cx, int cy)
{
	CDialog::OnSize(nType, cx, cy);
	g_Resize(this, &ActiveGanttVCCtl1);
}