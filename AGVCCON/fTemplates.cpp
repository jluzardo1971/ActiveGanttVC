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
#include "fTemplates.h"
#include "afxdialogex.h"

IMPLEMENT_DYNAMIC(fTemplates, CDialog)

fTemplates::fTemplates(CWnd* pParent /*=NULL*/)
	: CDialog(fTemplates::IDD, pParent)
{
    mp_aControlTemplates.Add(_T("STC_NONE"));
    mp_aControlTemplates.Add(_T("STC_CH_SOLID_WHITE"));
    mp_aControlTemplates.Add(_T("STC_CH_SOLID_DARK_BLUE"));
    mp_aControlTemplates.Add(_T("STC_CH_SOLID_VIOLET"));
    mp_aControlTemplates.Add(_T("STC_CH_SOLID_GREEN"));
    mp_aControlTemplates.Add(_T("STC_CH_SOLID_RED"));
    mp_aControlTemplates.Add(_T("STC_CH_SOLID_LIGHT_BLUE"));
    mp_aControlTemplates.Add(_T("STC_CH_SOLID_GREY"));
    mp_aControlTemplates.Add(_T("STC_CH_SOLID_LIGHT_STEEL_BLUE"));
    mp_aControlTemplates.Add(_T("STC_CH_VGRAD_YELLOW"));
    mp_aControlTemplates.Add(_T("STC_CH_VGRAD_ORANGE"));
    mp_aControlTemplates.Add(_T("STC_CH_VGRAD_RED"));
    mp_aControlTemplates.Add(_T("STC_CH_VGRAD_CRIMSON"));
    mp_aControlTemplates.Add(_T("STC_CH_VGRAD_MAGENTA"));
    mp_aControlTemplates.Add(_T("STC_CH_VGRAD_MULBERRY"));
    mp_aControlTemplates.Add(_T("STC_CH_VGRAD_BLUE_VIOLET"));
    mp_aControlTemplates.Add(_T("STC_CH_VGRAD_ANAKIWA_BLUE"));
    mp_aControlTemplates.Add(_T("STC_CH_VGRAD_BLUE_BELL"));
    mp_aControlTemplates.Add(_T("STC_CH_VGRAD_BLUE"));
    mp_aControlTemplates.Add(_T("STC_CH_VGRAD_AERO"));
    mp_aControlTemplates.Add(_T("STC_CH_HGRAD_YELLOW"));
    mp_aControlTemplates.Add(_T("STC_CH_HGRAD_ORANGE"));
    mp_aControlTemplates.Add(_T("STC_CH_HGRAD_RED"));
    mp_aControlTemplates.Add(_T("STC_CH_HGRAD_CRIMSON"));
    mp_aControlTemplates.Add(_T("STC_CH_HGRAD_MAGENTA"));
    mp_aControlTemplates.Add(_T("STC_CH_HGRAD_MULBERRY"));
    mp_aControlTemplates.Add(_T("STC_CH_HGRAD_BLUE_VIOLET"));
    mp_aControlTemplates.Add(_T("STC_CH_HGRAD_ANAKIWA_BLUE"));
    mp_aControlTemplates.Add(_T("STC_CH_HGRAD_BLUE_BELL"));
    mp_aControlTemplates.Add(_T("STC_CH_HGRAD_BLUE"));
    mp_aControlTemplates.Add(_T("STC_CH_HGRAD_AERO"));


	mp_aObjectTemplates.Add(_T("STO_NONE"));
    mp_aObjectTemplates.Add(_T("STO_BW_HATCH"));
    mp_aObjectTemplates.Add(_T("STO_COLOR_HATCH"));
    mp_aObjectTemplates.Add(_T("STO_GREY_SOLID"));
    mp_aObjectTemplates.Add(_T("STO_COLORS_SOLID"));
    mp_aObjectTemplates.Add(_T("STO_DEFAULT"));
    mp_aObjectTemplates.Add(_T("STO_VARIATION"));
}

fTemplates::~fTemplates()
{
}

void fTemplates::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	DDX_Control(pDX, IDC_ACTIVEGANTTVCCTL1, ActiveGanttVCCtl1);
}


BEGIN_MESSAGE_MAP(fTemplates, CDialog)
	ON_WM_SIZE()
	ON_WM_CREATE()

	ON_CBN_SELCHANGE(DRP_CONTROLTEMPLATES, cboControlTemplates_SelChange)
	ON_CBN_SELCHANGE(DRP_OBJECTTEMPLATES, cboObjectTemplates_SelChange)

	//Separator ID_MYTOOLBAR_BUTTONFIRST + 0
	ON_COMMAND(ID_MYTOOLBAR_BUTTONFIRST + 1, cmdCopy)
END_MESSAGE_MAP()

int fTemplates::OnCreate(LPCREATESTRUCT lpCreateStruct)
{
	if (CDialog::OnCreate(lpCreateStruct) == -1)
		return -1;

	mp_oFont.CreatePointFont(8, _T("MS Sans Serif"));
	oToolBar1.CreateEx(this, TBSTYLE_FLAT, WS_CHILD | WS_VISIBLE | CBRS_TOP | CBRS_GRIPPER | CBRS_TOOLTIPS | CBRS_FLYBY | CBRS_SIZE_DYNAMIC);
	CRect rect;

	rect.top = 1;
	rect.left = 0;
	rect.right = 100;
	rect.bottom = rect.top + 250;
	lblControlTemplates.Create(_T("Control Template:"), SS_LEFT | WS_VISIBLE, rect, &oToolBar1, LBL_CONTROLTEMPLATES);
	lblControlTemplates.SetFont(&mp_oFont);


	rect.top = 1;
	rect.left = 100;
	rect.right = 350;
	rect.bottom = rect.top + 250;
	cboControlTemplates.Create(CBS_DROPDOWNLIST | WS_VSCROLL | WS_VISIBLE, rect, &oToolBar1, DRP_CONTROLTEMPLATES);
	cboControlTemplates.SetFont(&mp_oFont);

	rect.top = 1;
	rect.left = 360;
	rect.right = 460;
	rect.bottom = rect.top + 250;
	lblObjectTemplates.Create(_T("Object Template:"), SS_LEFT | WS_VISIBLE, rect, &oToolBar1, LBL_OBJECTTEMPLATES);
	lblObjectTemplates.SetFont(&mp_oFont);

	rect.top = 1;
	rect.left = 460;
	rect.right = 660;
	rect.bottom = rect.top + 250;
	cboObjectTemplates.Create(CBS_DROPDOWNLIST | WS_VSCROLL | WS_VISIBLE, rect, &oToolBar1, DRP_OBJECTTEMPLATES);
	cboObjectTemplates.SetFont(&mp_oFont);

	oToolBarBitmap.LoadBitmapW(IDB_FTEMPLATES);
	oToolBar1.SetBitmap((HBITMAP)oToolBarBitmap.GetSafeHandle());
	oToolBar1.SetButtons(NULL, 2);  //1 Based

	oToolBar1.AddSeparator(670);
	oToolBar1.AddButton(0, _T("Copy"));


	return 0;
}



BOOL fTemplates::OnInitDialog()
{
	CDialog::OnInitDialog();
	g_MaximizeWindowsDim(this);
	CWnd::ShowWindow(SW_SHOWMAXIMIZED);

	// Reposition the Toolbar
	CWnd::RepositionBars(AFX_IDW_CONTROLBAR_FIRST, AFX_IDW_CONTROLBAR_LAST,0);

	//If you open the form: Styles And Templates -> Available Templates in the main menu (fTemplates.vb)
	//you can preview all available Templates.
	//Or you can also build your own:
	//fMSProject11.vb shows you how to build a Solid Template in the InitializeAG Method.
	//fMSProject14.vb shows you how to build a Gradient Template in the InitializeAG Method.
	ActiveGanttVCCtl1.ApplyTemplate(STC_CH_VGRAD_ANAKIWA_BLUE, STO_DEFAULT);

	int i;
	for (i = 0; i <= mp_aControlTemplates.GetCount() - 1; i++)
	{
		cboControlTemplates.AddString(mp_aControlTemplates[i]);
	}
	cboControlTemplates.SetCurSel((int)ActiveGanttVCCtl1.GetControlTemplate());

	for (i = 0; i <= mp_aObjectTemplates.GetCount() - 1; i++)
	{
		cboObjectTemplates.AddString(mp_aObjectTemplates[i]);
	}
	cboObjectTemplates.SetCurSel((int)ActiveGanttVCCtl1.GetObjectTemplate());

	CclsColumn oColumn;
	oColumn = ActiveGanttVCCtl1.GetColumns().Add(_T("Object"), _T(""), 70, _T(""));
	oColumn = ActiveGanttVCCtl1.GetColumns().Add(_T("Style"), _T(""), 50, _T(""));
	oColumn = ActiveGanttVCCtl1.GetColumns().Add(_T("Nodes"), _T(""), 125, _T(""));

	int lIndex = 0;
	for (lIndex = 1; lIndex <= 100; lIndex++) 
	{
		CclsRow oRow;
		oRow = ActiveGanttVCCtl1.GetRows().Add(_T("K") + CStr(lIndex), _T(""), FALSE, TRUE, _T(""));
		oRow.SetText(_T("Row ") + CStr(lIndex));
		if (oRow.GetIndex() % 5 == 0 || oRow.GetIndex() == 1) 
		{
			oRow.GetNode().SetDepth(0);
		}
		else 
		{
			oRow.GetNode().SetDepth(1);
		}
		oRow.GetNode().SetCheckBoxVisible(TRUE);
		oRow.SetHeight(20);
	}
	ActiveGanttVCCtl1.GetTreeview().SetCheckBoxes(TRUE);
	ActiveGanttVCCtl1.SetTreeviewColumnIndex(3);
	ActiveGanttVCCtl1.GetRows().UpdateTree();

	CclsTask oTask;
	CclsPercentage oPercentage;
	CclsPredecessor oPredecessor;
	CclsTimeBlock oTimeBlock;
	mp_CaptionRow(_T("K1"), _T("Task"), _T("T1"));
	mp_CaptionRow(_T("K3"), _T("Task"), _T("T2"));
	mp_CaptionRow(_T("K5"), _T("Task"), _T("T3"));
	mp_CaptionRow(_T("K7"), _T("Task"), _T("T4"));
	mp_CaptionRow(_T("K9"), _T("Task"), _T("S1"));
	mp_CaptionRow(_T("K11"), _T("Task"), _T("S2"));
	mp_CaptionRow(_T("K13"), _T("Task"), _T("S3"));
	mp_CaptionRow(_T("K15"), _T("Task"), _T("S4"));
	mp_CaptionRow(_T("K17"), _T("Percentage"), _T("P1"));
	mp_CaptionRow(_T("K19"), _T("Percentage"), _T("P2"));
	mp_CaptionRow(_T("K21"), _T("Percentage"), _T("P3"));
	mp_CaptionRow(_T("K23"), _T("Percentage"), _T("P4"));
	mp_CaptionRow(_T("K25"), _T("Percentage"), _T("SP1"));
	mp_CaptionRow(_T("K27"), _T("Percentage"), _T("SP2"));
	mp_CaptionRow(_T("K29"), _T("Percentage"), _T("SP3"));
	mp_CaptionRow(_T("K31"), _T("Percentage"), _T("SP4"));
	mp_CaptionRow(_T("K33"), _T("Task"), _T("RET1"));
	mp_CaptionRow(_T("K35"), _T("Task"), _T("RET2"));
	mp_CaptionRow(_T("K37"), _T("Task"), _T("RET3"));
	mp_CaptionRow(_T("K39"), _T("Task"), _T("RET4"));

	oTask = ActiveGanttVCCtl1.GetTasks().Add(_T("T1"), _T("K1"), GetDateTime(2013, 1, 1, 1, 0, 0), GetDateTime(2013, 1, 1, 5, 0, 0), _T("T1"), _T(""), _T("0"));
	oTask = ActiveGanttVCCtl1.GetTasks().Add(_T(""), _T("K2"), GetDateTime(2013, 1, 1, 8, 0, 0), GetDateTime(2013, 1, 1, 8, 0, 0), _T("M1"), _T(""), _T("0"));
	oPredecessor = ActiveGanttVCCtl1.GetPredecessors().Add(_T("M1"), _T("T1"), PCT_END_TO_START, _T("PD1"), _T(""));

	oTask = ActiveGanttVCCtl1.GetTasks().Add(_T("T2"), _T("K3"), GetDateTime(2013, 1, 1, 1, 0, 0), GetDateTime(2013, 1, 1, 5, 0, 0), _T("T2"), _T(""), _T("0"));
	oTask = ActiveGanttVCCtl1.GetTasks().Add(_T(""), _T("K4"), GetDateTime(2013, 1, 1, 8, 0, 0), GetDateTime(2013, 1, 1, 8, 0, 0), _T("M2"), _T(""), _T("0"));
	oPredecessor = ActiveGanttVCCtl1.GetPredecessors().Add(_T("M2"), _T("T2"), PCT_END_TO_START, _T("PD2"), _T(""));

	oTask = ActiveGanttVCCtl1.GetTasks().Add(_T("T3"), _T("K5"), GetDateTime(2013, 1, 1, 1, 0, 0), GetDateTime(2013, 1, 1, 5, 0, 0), _T("T3"), _T(""), _T("0"));
	oTask = ActiveGanttVCCtl1.GetTasks().Add(_T(""), _T("K6"), GetDateTime(2013, 1, 1, 8, 0, 0), GetDateTime(2013, 1, 1, 8, 0, 0), _T("M3"), _T(""), _T("0"));
	oPredecessor = ActiveGanttVCCtl1.GetPredecessors().Add(_T("M3"), _T("T3"), PCT_END_TO_START, _T("PD3"), _T(""));

	oTask = ActiveGanttVCCtl1.GetTasks().Add(_T("T4"), _T("K7"), GetDateTime(2013, 1, 1, 1, 0, 0), GetDateTime(2013, 1, 1, 5, 0, 0), _T("T4"), _T(""), _T("0"));
	oTask = ActiveGanttVCCtl1.GetTasks().Add(_T(""), _T("K8"), GetDateTime(2013, 1, 1, 8, 0, 0), GetDateTime(2013, 1, 1, 8, 0, 0), _T("M4"), _T(""), _T("0"));
	oPredecessor = ActiveGanttVCCtl1.GetPredecessors().Add(_T("M4"), _T("T4"), PCT_END_TO_START, _T("PD4"), _T(""));

	oTask = ActiveGanttVCCtl1.GetTasks().Add(_T("S1"), _T("K9"), GetDateTime(2013, 1, 1, 1, 0, 0), GetDateTime(2013, 1, 1, 5, 0, 0), _T("S1"), _T(""), _T("0"));
	oTask = ActiveGanttVCCtl1.GetTasks().Add(_T("S2"), _T("K11"), GetDateTime(2013, 1, 1, 1, 0, 0), GetDateTime(2013, 1, 1, 5, 0, 0), _T("S2"), _T(""), _T("0"));
	oTask = ActiveGanttVCCtl1.GetTasks().Add(_T("S3"), _T("K13"), GetDateTime(2013, 1, 1, 1, 0, 0), GetDateTime(2013, 1, 1, 5, 0, 0), _T("S3"), _T(""), _T("0"));
	oTask = ActiveGanttVCCtl1.GetTasks().Add(_T("S4"), _T("K15"), GetDateTime(2013, 1, 1, 1, 0, 0), GetDateTime(2013, 1, 1, 5, 0, 0), _T("S4"), _T(""), _T("0"));
	oTask = ActiveGanttVCCtl1.GetTasks().Add(_T(""), _T("K17"), GetDateTime(2013, 1, 1, 1, 0, 0), GetDateTime(2013, 1, 1, 5, 0, 0), _T("T1P1"), _T(""), _T("0"));
	oTask = ActiveGanttVCCtl1.GetTasks().Add(_T(""), _T("K19"), GetDateTime(2013, 1, 1, 1, 0, 0), GetDateTime(2013, 1, 1, 5, 0, 0), _T("T2P2"), _T(""), _T("0"));
	oTask = ActiveGanttVCCtl1.GetTasks().Add(_T(""), _T("K21"), GetDateTime(2013, 1, 1, 1, 0, 0), GetDateTime(2013, 1, 1, 5, 0, 0), _T("T3P3"), _T(""), _T("0"));
	oTask = ActiveGanttVCCtl1.GetTasks().Add(_T(""), _T("K23"), GetDateTime(2013, 1, 1, 1, 0, 0), GetDateTime(2013, 1, 1, 5, 0, 0), _T("T4P4"), _T(""), _T("0"));
	oTask = ActiveGanttVCCtl1.GetTasks().Add(_T(""), _T("K25"), GetDateTime(2013, 1, 1, 1, 0, 0), GetDateTime(2013, 1, 1, 5, 0, 0), _T("S1SP1"), _T(""), _T("0"));
	oTask = ActiveGanttVCCtl1.GetTasks().Add(_T(""), _T("K27"), GetDateTime(2013, 1, 1, 1, 0, 0), GetDateTime(2013, 1, 1, 5, 0, 0), _T("S2SP2"), _T(""), _T("0"));
	oTask = ActiveGanttVCCtl1.GetTasks().Add(_T(""), _T("K29"), GetDateTime(2013, 1, 1, 1, 0, 0), GetDateTime(2013, 1, 1, 5, 0, 0), _T("S3SP3"), _T(""), _T("0"));
	oTask = ActiveGanttVCCtl1.GetTasks().Add(_T(""), _T("K31"), GetDateTime(2013, 1, 1, 1, 0, 0), GetDateTime(2013, 1, 1, 5, 0, 0), _T("S4SP4"), _T(""), _T("0"));
	oPercentage = ActiveGanttVCCtl1.GetPercentages().Add(_T("T1P1"), _T(""), 0.5, _T("P1"));
	oPercentage = ActiveGanttVCCtl1.GetPercentages().Add(_T("T2P2"), _T(""), 0.5, _T("P2"));
	oPercentage = ActiveGanttVCCtl1.GetPercentages().Add(_T("T3P3"), _T(""), 0.5, _T("P3"));
	oPercentage = ActiveGanttVCCtl1.GetPercentages().Add(_T("T4P4"), _T(""), 0.5, _T("P4"));
	oPercentage = ActiveGanttVCCtl1.GetPercentages().Add(_T("S1SP1"), _T(""), 0.5, _T("SP1"));
	oPercentage = ActiveGanttVCCtl1.GetPercentages().Add(_T("S2SP2"), _T(""), 0.5, _T("SP2"));
	oPercentage = ActiveGanttVCCtl1.GetPercentages().Add(_T("S3SP3"), _T(""), 0.5, _T("SP3"));
	oPercentage = ActiveGanttVCCtl1.GetPercentages().Add(_T("S4SP4"), _T(""), 0.5, _T("SP4"));

	oTask = ActiveGanttVCCtl1.GetTasks().Add(_T("RET1"), _T("K33"), GetDateTime(2013, 1, 1, 1, 0, 0), GetDateTime(2013, 1, 1, 5, 0, 0), _T("RET1"), _T(""), _T("0"));
	oTask = ActiveGanttVCCtl1.GetTasks().Add(_T("RET2"), _T("K35"), GetDateTime(2013, 1, 1, 1, 0, 0), GetDateTime(2013, 1, 1, 5, 0, 0), _T("RET2"), _T(""), _T("0"));
	oTask = ActiveGanttVCCtl1.GetTasks().Add(_T("RET3"), _T("K37"), GetDateTime(2013, 1, 1, 1, 0, 0), GetDateTime(2013, 1, 1, 5, 0, 0), _T("RET3"), _T(""), _T("0"));
	oTask = ActiveGanttVCCtl1.GetTasks().Add(_T("RET4"), _T("K39"), GetDateTime(2013, 1, 1, 1, 0, 0), GetDateTime(2013, 1, 1, 5, 0, 0), _T("RET4"), _T(""), _T("0"));

	oTimeBlock = ActiveGanttVCCtl1.GetTimeBlocks().Add(_T("TB1"));
	oTimeBlock.SetBaseDate(GetDateTime(2013, 1, 1, 9, 0, 0));
	oTimeBlock.SetDurationInterval(IL_HOUR);
	oTimeBlock.SetDurationFactor(3);

	mp_SetStyles();

	ActiveGanttVCCtl1.GetCurrentViewObject().GetTimeLine().Position(GetDateTime(2013, 1, 1));

	return TRUE; 
}

void fTemplates::mp_CaptionRow(CString sRowKey, CString sObject, CString sStyle)
{
	ActiveGanttVCCtl1.GetRows().Item(sRowKey).GetCells().Item(_T("1")).SetText(sObject);
	ActiveGanttVCCtl1.GetRows().Item(sRowKey).GetCells().Item(_T("2")).SetText(sStyle);
}


void fTemplates::OnSize(UINT nType, int cx, int cy)
{
	CDialog::OnSize(nType, cx, cy);
	g_Resize(this, &ActiveGanttVCCtl1);

}


void fTemplates::mp_SetStyles()
{
	//Setting a template invokes the Clear method on the Styles Collection
	//and resets all StyleIndex properties of every object to ""
	if (ActiveGanttVCCtl1.GetStyles().ContainsKey(_T("T1")) == TRUE) 
	{
		ActiveGanttVCCtl1.GetTasks().Item(_T("T1")).SetStyleIndex(_T("T1"));
		ActiveGanttVCCtl1.GetTasks().Item(_T("T2")).SetStyleIndex(_T("T2"));
		ActiveGanttVCCtl1.GetTasks().Item(_T("T3")).SetStyleIndex(_T("T3"));
		ActiveGanttVCCtl1.GetTasks().Item(_T("T4")).SetStyleIndex(_T("T4"));
		ActiveGanttVCCtl1.GetTasks().Item(_T("M1")).SetStyleIndex(_T("T1"));
		ActiveGanttVCCtl1.GetTasks().Item(_T("M2")).SetStyleIndex(_T("T2"));
		ActiveGanttVCCtl1.GetTasks().Item(_T("M3")).SetStyleIndex(_T("T3"));
		ActiveGanttVCCtl1.GetTasks().Item(_T("M4")).SetStyleIndex(_T("T4"));
		ActiveGanttVCCtl1.GetTasks().Item(_T("T1P1")).SetStyleIndex(_T("T1"));
		ActiveGanttVCCtl1.GetTasks().Item(_T("T2P2")).SetStyleIndex(_T("T2"));
		ActiveGanttVCCtl1.GetTasks().Item(_T("T3P3")).SetStyleIndex(_T("T3"));
		ActiveGanttVCCtl1.GetTasks().Item(_T("T4P4")).SetStyleIndex(_T("T4"));
		ActiveGanttVCCtl1.GetPredecessors().Item(_T("PD1")).SetStyleIndex(_T("T1"));
		ActiveGanttVCCtl1.GetPredecessors().Item(_T("PD2")).SetStyleIndex(_T("T2"));
		ActiveGanttVCCtl1.GetPredecessors().Item(_T("PD3")).SetStyleIndex(_T("T3"));
		ActiveGanttVCCtl1.GetPredecessors().Item(_T("PD4")).SetStyleIndex(_T("T4"));

		ActiveGanttVCCtl1.GetTasks().Item(_T("T1")).SetVisible(TRUE);
		ActiveGanttVCCtl1.GetTasks().Item(_T("T2")).SetVisible(TRUE);
		ActiveGanttVCCtl1.GetTasks().Item(_T("T3")).SetVisible(TRUE);
		ActiveGanttVCCtl1.GetTasks().Item(_T("T4")).SetVisible(TRUE);
		ActiveGanttVCCtl1.GetTasks().Item(_T("M1")).SetVisible(TRUE);
		ActiveGanttVCCtl1.GetTasks().Item(_T("M2")).SetVisible(TRUE);
		ActiveGanttVCCtl1.GetTasks().Item(_T("M3")).SetVisible(TRUE);
		ActiveGanttVCCtl1.GetTasks().Item(_T("M4")).SetVisible(TRUE);
		ActiveGanttVCCtl1.GetTasks().Item(_T("T1P1")).SetVisible(TRUE);
		ActiveGanttVCCtl1.GetTasks().Item(_T("T2P2")).SetVisible(TRUE);
		ActiveGanttVCCtl1.GetTasks().Item(_T("T3P3")).SetVisible(TRUE);
		ActiveGanttVCCtl1.GetTasks().Item(_T("T4P4")).SetVisible(TRUE);
		ActiveGanttVCCtl1.GetPredecessors().Item(_T("PD1")).SetVisible(TRUE);
		ActiveGanttVCCtl1.GetPredecessors().Item(_T("PD2")).SetVisible(TRUE);
		ActiveGanttVCCtl1.GetPredecessors().Item(_T("PD3")).SetVisible(TRUE);
		ActiveGanttVCCtl1.GetPredecessors().Item(_T("PD4")).SetVisible(TRUE);
	}
	else 
	{
		ActiveGanttVCCtl1.GetTasks().Item(_T("T1")).SetVisible(FALSE);
		ActiveGanttVCCtl1.GetTasks().Item(_T("T2")).SetVisible(FALSE);
		ActiveGanttVCCtl1.GetTasks().Item(_T("T3")).SetVisible(FALSE);
		ActiveGanttVCCtl1.GetTasks().Item(_T("T4")).SetVisible(FALSE);
		ActiveGanttVCCtl1.GetTasks().Item(_T("M1")).SetVisible(FALSE);
		ActiveGanttVCCtl1.GetTasks().Item(_T("M2")).SetVisible(FALSE);
		ActiveGanttVCCtl1.GetTasks().Item(_T("M3")).SetVisible(FALSE);
		ActiveGanttVCCtl1.GetTasks().Item(_T("M4")).SetVisible(FALSE);
		ActiveGanttVCCtl1.GetTasks().Item(_T("T1P1")).SetVisible(FALSE);
		ActiveGanttVCCtl1.GetTasks().Item(_T("T2P2")).SetVisible(FALSE);
		ActiveGanttVCCtl1.GetTasks().Item(_T("T3P3")).SetVisible(FALSE);
		ActiveGanttVCCtl1.GetTasks().Item(_T("T4P4")).SetVisible(FALSE);
		ActiveGanttVCCtl1.GetPredecessors().Item(_T("PD1")).SetVisible(FALSE);
		ActiveGanttVCCtl1.GetPredecessors().Item(_T("PD2")).SetVisible(FALSE);
		ActiveGanttVCCtl1.GetPredecessors().Item(_T("PD3")).SetVisible(FALSE);
		ActiveGanttVCCtl1.GetPredecessors().Item(_T("PD4")).SetVisible(FALSE);
	}

	if (ActiveGanttVCCtl1.GetStyles().ContainsKey(_T("P1")) == TRUE) 
	{
		ActiveGanttVCCtl1.GetPercentages().Item(_T("P1")).SetStyleIndex(_T("P1"));
		ActiveGanttVCCtl1.GetPercentages().Item(_T("P2")).SetStyleIndex(_T("P2"));
		ActiveGanttVCCtl1.GetPercentages().Item(_T("P3")).SetStyleIndex(_T("P3"));
		ActiveGanttVCCtl1.GetPercentages().Item(_T("P4")).SetStyleIndex(_T("P4"));

		ActiveGanttVCCtl1.GetPercentages().Item(_T("P1")).SetVisible(TRUE);
		ActiveGanttVCCtl1.GetPercentages().Item(_T("P2")).SetVisible(TRUE);
		ActiveGanttVCCtl1.GetPercentages().Item(_T("P3")).SetVisible(TRUE);
		ActiveGanttVCCtl1.GetPercentages().Item(_T("P4")).SetVisible(TRUE);
	}
	else 
	{
		ActiveGanttVCCtl1.GetPercentages().Item(_T("P1")).SetVisible(FALSE);
		ActiveGanttVCCtl1.GetPercentages().Item(_T("P2")).SetVisible(FALSE);
		ActiveGanttVCCtl1.GetPercentages().Item(_T("P3")).SetVisible(FALSE);
		ActiveGanttVCCtl1.GetPercentages().Item(_T("P4")).SetVisible(FALSE);
	}

	if (ActiveGanttVCCtl1.GetStyles().ContainsKey(_T("S1")) == TRUE) 
	{
		ActiveGanttVCCtl1.GetTasks().Item(_T("S1")).SetStyleIndex(_T("S1"));
		ActiveGanttVCCtl1.GetTasks().Item(_T("S2")).SetStyleIndex(_T("S2"));
		ActiveGanttVCCtl1.GetTasks().Item(_T("S3")).SetStyleIndex(_T("S3"));
		ActiveGanttVCCtl1.GetTasks().Item(_T("S4")).SetStyleIndex(_T("S4"));
		ActiveGanttVCCtl1.GetTasks().Item(_T("S1SP1")).SetStyleIndex(_T("S1"));
		ActiveGanttVCCtl1.GetTasks().Item(_T("S2SP2")).SetStyleIndex(_T("S2"));
		ActiveGanttVCCtl1.GetTasks().Item(_T("S3SP3")).SetStyleIndex(_T("S3"));
		ActiveGanttVCCtl1.GetTasks().Item(_T("S4SP4")).SetStyleIndex(_T("S4"));

		ActiveGanttVCCtl1.GetTasks().Item(_T("S1")).SetVisible(TRUE);
		ActiveGanttVCCtl1.GetTasks().Item(_T("S2")).SetVisible(TRUE);
		ActiveGanttVCCtl1.GetTasks().Item(_T("S3")).SetVisible(TRUE);
		ActiveGanttVCCtl1.GetTasks().Item(_T("S4")).SetVisible(TRUE);
		ActiveGanttVCCtl1.GetTasks().Item(_T("S1SP1")).SetVisible(TRUE);
		ActiveGanttVCCtl1.GetTasks().Item(_T("S2SP2")).SetVisible(TRUE);
		ActiveGanttVCCtl1.GetTasks().Item(_T("S3SP3")).SetVisible(TRUE);
		ActiveGanttVCCtl1.GetTasks().Item(_T("S4SP4")).SetVisible(TRUE);
	}
	else 
	{
		ActiveGanttVCCtl1.GetTasks().Item(_T("S1")).SetVisible(FALSE);
		ActiveGanttVCCtl1.GetTasks().Item(_T("S2")).SetVisible(FALSE);
		ActiveGanttVCCtl1.GetTasks().Item(_T("S3")).SetVisible(FALSE);
		ActiveGanttVCCtl1.GetTasks().Item(_T("S4")).SetVisible(FALSE);
		ActiveGanttVCCtl1.GetTasks().Item(_T("S1SP1")).SetVisible(FALSE);
		ActiveGanttVCCtl1.GetTasks().Item(_T("S2SP2")).SetVisible(FALSE);
		ActiveGanttVCCtl1.GetTasks().Item(_T("S3SP3")).SetVisible(FALSE);
		ActiveGanttVCCtl1.GetTasks().Item(_T("S4SP4")).SetVisible(FALSE);
	}

	if (ActiveGanttVCCtl1.GetStyles().ContainsKey(_T("SP1")) == TRUE) 
	{
		ActiveGanttVCCtl1.GetPercentages().Item(_T("SP1")).SetStyleIndex(_T("SP1"));
		ActiveGanttVCCtl1.GetPercentages().Item(_T("SP2")).SetStyleIndex(_T("SP2"));
		ActiveGanttVCCtl1.GetPercentages().Item(_T("SP3")).SetStyleIndex(_T("SP3"));
		ActiveGanttVCCtl1.GetPercentages().Item(_T("SP4")).SetStyleIndex(_T("SP4"));

		ActiveGanttVCCtl1.GetPercentages().Item(_T("SP1")).SetVisible(TRUE);
		ActiveGanttVCCtl1.GetPercentages().Item(_T("SP2")).SetVisible(TRUE);
		ActiveGanttVCCtl1.GetPercentages().Item(_T("SP3")).SetVisible(TRUE);
		ActiveGanttVCCtl1.GetPercentages().Item(_T("SP4")).SetVisible(TRUE);
	}
	else 
	{
		ActiveGanttVCCtl1.GetPercentages().Item(_T("SP1")).SetVisible(FALSE);
		ActiveGanttVCCtl1.GetPercentages().Item(_T("SP2")).SetVisible(FALSE);
		ActiveGanttVCCtl1.GetPercentages().Item(_T("SP3")).SetVisible(FALSE);
		ActiveGanttVCCtl1.GetPercentages().Item(_T("SP4")).SetVisible(FALSE);
	}

	if (ActiveGanttVCCtl1.GetStyles().ContainsKey(_T("RET1")) == TRUE) 
	{
		ActiveGanttVCCtl1.GetTasks().Item(_T("RET1")).SetStyleIndex(_T("RET1"));
		ActiveGanttVCCtl1.GetTasks().Item(_T("RET2")).SetStyleIndex(_T("RET2"));
		ActiveGanttVCCtl1.GetTasks().Item(_T("RET3")).SetStyleIndex(_T("RET3"));
		ActiveGanttVCCtl1.GetTasks().Item(_T("RET4")).SetStyleIndex(_T("RET4"));

		ActiveGanttVCCtl1.GetTasks().Item(_T("RET1")).SetVisible(TRUE);
		ActiveGanttVCCtl1.GetTasks().Item(_T("RET2")).SetVisible(TRUE);
		ActiveGanttVCCtl1.GetTasks().Item(_T("RET3")).SetVisible(TRUE);
		ActiveGanttVCCtl1.GetTasks().Item(_T("RET4")).SetVisible(TRUE);
	}
	else 
	{
		ActiveGanttVCCtl1.GetTasks().Item(_T("RET1")).SetVisible(FALSE);
		ActiveGanttVCCtl1.GetTasks().Item(_T("RET2")).SetVisible(FALSE);
		ActiveGanttVCCtl1.GetTasks().Item(_T("RET3")).SetVisible(FALSE);
		ActiveGanttVCCtl1.GetTasks().Item(_T("RET4")).SetVisible(FALSE);
	}
}

void fTemplates::cboControlTemplates_SelChange()
{
	if ((int)ActiveGanttVCCtl1.GetControlTemplate() != cboControlTemplates.GetCurSel()) 
	{
		ActiveGanttVCCtl1.ApplyTemplate((E_CONTROLTEMPLATE)cboControlTemplates.GetCurSel(), (E_OBJECTTEMPLATE)cboObjectTemplates.GetCurSel());
		mp_SetStyles();
		ActiveGanttVCCtl1.Redraw();
	}
}

void fTemplates::cboObjectTemplates_SelChange()
{
	if ((int)ActiveGanttVCCtl1.GetObjectTemplate() != cboObjectTemplates.GetCurSel()) 
	{
		ActiveGanttVCCtl1.ApplyTemplate((E_CONTROLTEMPLATE)cboControlTemplates.GetCurSel(), (E_OBJECTTEMPLATE)cboObjectTemplates.GetCurSel());
		mp_SetStyles();
		ActiveGanttVCCtl1.Redraw();
	}
}

void fTemplates::cmdCopy()
{
	OpenClipboard();
	EmptyClipboard();
	CString sText = _T("ActiveGanttVCCtl1.ApplyTemplate(") + mp_aControlTemplates[cboControlTemplates.GetCurSel()] + _T(", ") + mp_aObjectTemplates[cboObjectTemplates.GetCurSel()] + _T(");");
	const size_t StringSize = 4000;
	size_t CharactersConverted = 0;
	char aText[StringSize];
	wcstombs_s(&CharactersConverted, aText, sText.GetLength() + 1, sText,  _TRUNCATE);
	HGLOBAL hGlob = GlobalAlloc(GMEM_FIXED, sText.GetLength() + 1);
	strcpy_s((char*)hGlob, sText.GetLength() + 1, aText);
	::SetClipboardData(CF_TEXT, hGlob);
	CloseClipboard();
	MessageBox(_T("\"") + sText + _T("\" has been copied to the clipboard."), _T("Templates"));
}