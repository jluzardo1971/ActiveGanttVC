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
#include "fWBSProject.h"
#include "fWBSProjectTaskView.h"
#include "fWBSPProperties.h"
#include "fLoadXML.h"
#include "fPrintDialog.h"

IMPLEMENT_DYNAMIC(fWBSProject, CDialog)

void fWBSProject::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	DDX_Control(pDX, IDC_ACTIVEGANTTVCCTL1, ActiveGanttVCCtl1);
}


BEGIN_MESSAGE_MAP(fWBSProject, CDialog)
	ON_WM_SIZE()
	ON_WM_CREATE()

	ON_COMMAND(ID_MYTOOLBAR_BUTTONFIRST + 0, cmdSaveXML)
	ON_COMMAND(ID_MYTOOLBAR_BUTTONFIRST + 1, cmdLoadXML)
	ON_COMMAND(ID_MYTOOLBAR_BUTTONFIRST + 2, cmdPrint_Click)

	ON_COMMAND(ID_MYTOOLBAR_BUTTONFIRST + 4, cmdZoomIn_Click)
	ON_COMMAND(ID_MYTOOLBAR_BUTTONFIRST + 5, cmdZoomOut_Click)

	ON_COMMAND(ID_MYTOOLBAR_BUTTONFIRST + 7, cmdBluePercentages_Click)
	ON_COMMAND(ID_MYTOOLBAR_BUTTONFIRST + 8, cmdGreenPercentages_Click)
	ON_COMMAND(ID_MYTOOLBAR_BUTTONFIRST + 9, cmdRedPercentages_Click)

	ON_COMMAND(ID_MYTOOLBAR_BUTTONFIRST + 11, cmdProperties_Click)
	ON_COMMAND(ID_MYTOOLBAR_BUTTONFIRST + 12, cmdCheckPredecessors_Click)

	ON_COMMAND(ID_MYTOOLBAR_BUTTONFIRST + 14, cmdToolTips_Click)
	ON_COMMAND(ID_MYTOOLBAR_BUTTONFIRST + 15, cmdHelp_Click)

	ON_COMMAND(ID_FILE_SAVEASXML_WBSP, mnuSaveXML_Click)
	ON_COMMAND(ID_FILE_LOADFROMXML_WBSP, mnuLoadXML_Click)
	ON_COMMAND(ID_FILE_PRINT_WBSP, mnuPrint_Click)
	ON_COMMAND(ID_FILE_CLOSE_WBSP, mnuClose_Click)
	ON_COMMAND(ID_TREEVIEWPROPERTIES_CHECKBOXES, mnuCheckboxes_Click)
	ON_COMMAND(ID_TREEVIEWPROPERTIES_IMAGES, mnuPictures_Click)
	ON_COMMAND(ID_TREEVIEWPROPERTIES_FULLCOLUMNSELECT, mnuFullColumnSelect_Click)

	ON_WM_CLOSE()
END_MESSAGE_MAP()

BEGIN_EVENTSINK_MAP(fWBSProject, CDialog)
	ON_EVENT(fWBSProject, IDC_ACTIVEGANTTVCCTL1, 20, fWBSProject::CompleteObjectChangeActiveganttvcctl1, VTS_DISPATCH)
	ON_EVENT(fWBSProject, IDC_ACTIVEGANTTVCCTL1, 3, fWBSProject::ControlMouseDownActiveganttvcctl1, VTS_DISPATCH)
	ON_EVENT(fWBSProject, IDC_ACTIVEGANTTVCCTL1, 31, fWBSProject::ControlMouseWheelActiveganttvcctl1, VTS_DISPATCH)
	ON_EVENT(fWBSProject, IDC_ACTIVEGANTTVCCTL1, 22, fWBSProject::NodeCheckedActiveganttvcctl1, VTS_DISPATCH)
	ON_EVENT(fWBSProject, IDC_ACTIVEGANTTVCCTL1, 15, fWBSProject::ObjectAddedActiveganttvcctl1, VTS_DISPATCH)
	ON_EVENT(fWBSProject, IDC_ACTIVEGANTTVCCTL1, 27, fWBSProject::ToolTipOnMouseHoverActiveganttvcctl1, VTS_DISPATCH)
	ON_EVENT(fWBSProject, IDC_ACTIVEGANTTVCCTL1, 30, fWBSProject::OnMouseMoveToolTipDrawActiveganttvcctl1, VTS_DISPATCH)
	ON_EVENT(fWBSProject, IDC_ACTIVEGANTTVCCTL1, 28, fWBSProject::ToolTipOnMouseMoveActiveganttvcctl1, VTS_DISPATCH)
	ON_EVENT(fWBSProject, IDC_ACTIVEGANTTVCCTL1, 29, fWBSProject::OnMouseHoverToolTipDrawActiveganttvcctl1, VTS_DISPATCH)
	ON_EVENT(fWBSProject, IDC_ACTIVEGANTTVCCTL1, 33, fWBSProject::EndTextEditActiveganttvcctl1, VTS_DISPATCH)
	ON_EVENT(fWBSProject, IDC_ACTIVEGANTTVCCTL1, 38, fWBSProject::PagePrintActiveganttvcctl1, VTS_DISPATCH)
END_EVENTSINK_MAP()


fWBSProject::fWBSProject(CWnd* pParent /*=NULL*/)
	: CDialog(fWBSProject::IDD, pParent)
{

	mp_oConn.Open(NULL, FALSE, FALSE, g_DST_ACCESS_GetConnectionString());

	m_hIcon = AfxGetApp()->LoadIcon(IDR_MAINFRAME);
	mp_sFontName = _T("Tahoma");
	mp_bBluePercentagesVisible = TRUE;
	mp_bGreenPercentagesVisible = TRUE;
	mp_bRedPercentagesVisible = TRUE;
}

fWBSProject::~fWBSProject()
{
	mp_oConn.Close();
}


int fWBSProject::OnCreate(LPCREATESTRUCT lpCreateStruct)
{
	if (CDialog::OnCreate(lpCreateStruct) == -1)
		return -1;

	//Create ToolBar:
	oToolBar1.CreateEx(this, TBSTYLE_FLAT, WS_CHILD | WS_VISIBLE | CBRS_TOP | CBRS_GRIPPER | CBRS_TOOLTIPS | CBRS_FLYBY | CBRS_SIZE_DYNAMIC);
	oToolBarBitmap.LoadBitmapW(IDB_FWBSPROJECT);
	oToolBar1.SetBitmap((HBITMAP)oToolBarBitmap.GetSafeHandle());
	oToolBar1.SetButtons(NULL, 16); //1 Based

	oToolBar1.AddButton(0, _T("Save as XML"));
	oToolBar1.AddButton(1, _T("Load XML"));
	oToolBar1.AddButton(2, _T("Print"));
	oToolBar1.AddSeparator();
	oToolBar1.AddButton(3, _T("Zoom in"));
	oToolBar1.AddButton(4, _T("Zoom out"));
	oToolBar1.AddSeparator();
	oToolBar1.AddButton(5, _T("Toggle blue percentages"));
	oToolBar1.AddButton(6, _T("Toggle green percentages"));
	oToolBar1.AddButton(7, _T("Toggle red percentages"));
	oToolBar1.AddSeparator();
	oToolBar1.AddButton(8, _T("Properties"));
	oToolBar1.AddButton(9, _T("CheckPredecessors"));
	oToolBar1.AddSeparator();
	oToolBar1.AddButton(10, _T("Toggle tooltips"));
	oToolBar1.AddButton(11, _T("Help"));

	return 0;
}

BOOL fWBSProject::OnInitDialog()
{
	CDialog::OnInitDialog();
	g_MaximizeWindowsDim(this);
	CWnd::ShowWindow(SW_SHOWMAXIMIZED);

	SetIcon(m_hIcon, TRUE);			// Set big icon
	SetIcon(m_hIcon, FALSE);		// Set small icon

	// Reposition the Toolbar
	CWnd::RepositionBars(AFX_IDW_CONTROLBAR_FIRST, AFX_IDW_CONTROLBAR_LAST,0);

	CString sCaption = _T("Work Breakdown Structure (WBS) Project Management Example - ");
	sCaption = sCaption + _T("Microsoft Access data source (32bit compatible only) - ");
	sCaption = sCaption + _T("ActiveGanttVC Version: ") + ActiveGanttVCCtl1.GetVersion();
	this->SetWindowTextW(sCaption);

	CclsStyle oStyle;
	CclsView oView;

	//If you open the form: Styles And Templates -> Available Templates in the main menu (fTemplates.vb)
	//you can preview all available Templates.
	//Or you can also build your own:
	//fMSProject11.vb shows you how to build a Solid Template in the InitializeAG Method.
	//fMSProject14.vb shows you how to build a Gradient Template in the InitializeAG Method.
	ActiveGanttVCCtl1.ApplyTemplate(STC_CH_VGRAD_ANAKIWA_BLUE, STO_DEFAULT);

    oStyle = ActiveGanttVCCtl1.GetConfiguration().GetDefaultPercentageStyle().Clone(_T("InvisiblePercentages"));
    oStyle = ActiveGanttVCCtl1.GetStyles().Add(_T("SelectedPredecessor"));
    oStyle.GetPredecessorStyle().SetLineColor(CLR_GREEN);

	ActiveGanttVCCtl1.SetControlTag(_T("WBSProject"));
	ActiveGanttVCCtl1.SetAllowRowMove(TRUE);
	ActiveGanttVCCtl1.SetAllowRowSize(TRUE);
	ActiveGanttVCCtl1.SetAddMode(AT_BOTH);
	ActiveGanttVCCtl1.GetConfiguration().SetRowHighlightedBackColor(CLR_LIGHTBLUE);

	CclsColumn oColumn;

	oColumn = ActiveGanttVCCtl1.GetColumns().Add(_T("ID"), _T(""), 30, _T(""));
    oColumn.SetAllowTextEdit(TRUE);

	oColumn = ActiveGanttVCCtl1.GetColumns().Add(_T("Task Name"), _T(""), 300, _T(""));
    oColumn.SetAllowTextEdit(TRUE);

	oColumn = ActiveGanttVCCtl1.GetColumns().Add(_T("StartDate"), _T(""), 125, _T(""));
    oColumn.SetAllowTextEdit(TRUE);

	oColumn = ActiveGanttVCCtl1.GetColumns().Add(_T("EndDate"), _T(""), 125, _T(""));
    oColumn.SetAllowTextEdit(TRUE);

	ActiveGanttVCCtl1.SetTreeviewColumnIndex(2);

	ActiveGanttVCCtl1.GetTreeview().SetImages(TRUE);
	ActiveGanttVCCtl1.GetTreeview().SetCheckBoxes(TRUE);
	ActiveGanttVCCtl1.GetTreeview().SetTreeviewMode(TM_EXPANDCOLLAPSEICONS);
	ActiveGanttVCCtl1.GetTreeview().SetTreeLines(FALSE);
	ActiveGanttVCCtl1.GetToolTip().GetFont().SetFontName(mp_sFontName);
	ActiveGanttVCCtl1.GetToolTip().GetFont().SetSize(8);
	ActiveGanttVCCtl1.GetToolTip().GetFont().SetStyle(FS_FONTSTYLEREGULAR);
	ActiveGanttVCCtl1.GetSplitter().SetPosition(255);

	Access_LoadTasks();
	ActiveGanttVCCtl1.GetRows().UpdateTree();

	ActiveGanttVCCtl1.SetTimeLineScrollBarScope(OS_CONTROL);

	//// Start one month before the first task:
	ActiveGanttVCCtl1.GetTimeLineScrollBar().SetStartDate(ActiveGanttVCCtl1.GetMathLib().DateTimeAdd(IL_MONTH, -1, mp_dtStartDate));
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
	oView.GetTimeLine().GetTierArea().GetUpperTier().SetHeight(17);
	oView.GetTimeLine().GetTierArea().GetLowerTier().SetBackgroundMode(ETBGM_STYLE);
	oView.GetTimeLine().GetTierArea().GetLowerTier().SetHeight(17);
	oView.GetTimeLine().GetTickMarkArea().SetVisible(FALSE);
	oView.GetClientArea().SetDetectConflicts(FALSE);
	
	oView = oView.Clone(_T(""));
    oView.SetInterval(IL_HOUR);
    oView.SetFactor(12);

	oView = oView.Clone(_T(""));
    oView.SetInterval(IL_HOUR);
    oView.SetFactor(6);

	ActiveGanttVCCtl1.SetCurrentView(_T("2"));

	if (ActiveGanttVCCtl1.GetPrinterErrorMessage().GetLength() == 0)
	{
		ActiveGanttVCCtl1.GetPrinter().SetOrientation(OT_LANDSCAPE);
		ActiveGanttVCCtl1.GetPrinter().SetMarginLeft(50); //0.5"
		ActiveGanttVCCtl1.GetPrinter().SetMarginTop(50); //0.5"
		ActiveGanttVCCtl1.GetPrinter().SetMarginRight(50); //0.5"
		ActiveGanttVCCtl1.GetPrinter().SetMarginBottom(200); //2.0"
		ActiveGanttVCCtl1.GetPrinter().SetHeadingsInEveryPage(TRUE);
		ActiveGanttVCCtl1.GetPrinter().SetKeepRowsTogether(TRUE);
		ActiveGanttVCCtl1.GetPrinter().SetColumnsInEveryPage(TRUE);
		ActiveGanttVCCtl1.GetPrinter().SetColumns(1);
		ActiveGanttVCCtl1.GetPrinter().SetKeepColumnsTogether(TRUE);
		ActiveGanttVCCtl1.GetPrinter().SetKeepTimeIntervalsTogether(TRUE);
		ActiveGanttVCCtl1.GetPrinter().SetInterval(IL_MONTH);
		ActiveGanttVCCtl1.GetPrinter().SetFactor(1);
	}

	return TRUE;
}


void fWBSProject::OnSize(UINT nType, int cx, int cy)
{
	CDialog::OnSize(nType, cx, cy);
	g_Resize(this, &ActiveGanttVCCtl1);
}

void fWBSProject::CompleteObjectChangeActiveganttvcctl1(LPDISPATCH e)
{
	CObjectStateChangedEventArgs oE(e);
	switch (oE.GetEventTarget())
	{
	case EVT_TASK:
		{
			UpdateTask(oE.GetIndex());
		}
		break;
	case EVT_PERCENTAGE:
		{
			int lTaskIndex;
			lTaskIndex = ActiveGanttVCCtl1.GetTasks().Item(ActiveGanttVCCtl1.GetPercentages().Item(CStr(oE.GetIndex())).GetTaskKey()).GetIndex();
			UpdateTask(lTaskIndex);
		}
		break;
	}
}

void fWBSProject::ControlMouseDownActiveganttvcctl1(LPDISPATCH e)
{
	CMouseEventArgs oE(e);
	if ((oE.GetEventTarget() == EVT_TASK || oE.GetEventTarget() == EVT_SELECTEDTASK) && oE.GetButton() == BTN_RIGHT)
	{
		fWBSProjectTaskView oForm;
		oForm.mp_lTaskIndex = ActiveGanttVCCtl1.GetMathLib().GetTaskIndexByPosition(oE.GetX(), oE.GetY());
		oForm.mp_oParent = this;
		oForm.DoModal();
		oE.SetCancel(TRUE);
	}
}

void fWBSProject::ControlMouseWheelActiveganttvcctl1(LPDISPATCH e)
{
	CMouseWheelEventArgs oE(e);
	if ((oE.GetDelta() == 0) || (ActiveGanttVCCtl1.GetVerticalScrollBar().GetVisible() == FALSE))
	{
		return;
	}
	int lDelta;
	lDelta = CInt32(-(oE.GetDelta() / 50));
    int lInitialValue;
    lInitialValue = ActiveGanttVCCtl1.GetVerticalScrollBar().GetValue();
	if ((ActiveGanttVCCtl1.GetVerticalScrollBar().GetValue() + lDelta) < 1)
	{
		ActiveGanttVCCtl1.GetVerticalScrollBar().SetValue(1);
	}
	else if ((ActiveGanttVCCtl1.GetVerticalScrollBar().GetValue() + lDelta) > ActiveGanttVCCtl1.GetVerticalScrollBar().GetMax())
	{
		ActiveGanttVCCtl1.GetVerticalScrollBar().SetValue(ActiveGanttVCCtl1.GetVerticalScrollBar().GetMax());
	}
	else
	{
		ActiveGanttVCCtl1.GetVerticalScrollBar().SetValue(ActiveGanttVCCtl1.GetVerticalScrollBar().GetValue() + lDelta);
	}
    ActiveGanttVCCtl1.Redraw();
}

void fWBSProject::NodeCheckedActiveganttvcctl1(LPDISPATCH e)
{
	CNodeEventArgs oE(e);

	CclsRow oRow;
	oRow = ActiveGanttVCCtl1.GetRows().Item(CStr(oE.GetIndex()));
	oRow.SetHighlighted(oRow.GetNode().GetChecked());
}

void fWBSProject::ObjectAddedActiveganttvcctl1(LPDISPATCH e)
{
	CObjectAddedEventArgs oE(e);
	switch (oE.GetEventTarget())
	{
	case EVT_TASK:
	case EVT_MILESTONE:
		{
            CclsTask oTask;
            oTask = GetTaskByRowKey(ActiveGanttVCCtl1.GetTasks().Item(CStr(oE.GetTaskIndex())).GetRowKey());
            oTask.SetStartDate(ActiveGanttVCCtl1.GetTasks().Item(CStr(oE.GetTaskIndex())).GetStartDate());
            oTask.SetEndDate(ActiveGanttVCCtl1.GetTasks().Item(CStr(oE.GetTaskIndex())).GetEndDate());
			UpdateTask(oTask.GetIndex());
            ActiveGanttVCCtl1.GetTasks().Remove(CStr(oE.GetTaskIndex()));
		}
		break;
	case EVT_PREDECESSOR:
		{
			ActiveGanttVCCtl1.GetPredecessors().Item(CStr(oE.GetPredecessorObjectIndex())).SetStyleIndex(_T("T1"));
			ActiveGanttVCCtl1.GetPredecessors().Item(CStr(oE.GetPredecessorObjectIndex())).SetWarningStyleIndex(_T("TW1"));
			ActiveGanttVCCtl1.GetPredecessors().Item(CStr(oE.GetPredecessorObjectIndex())).SetSelectedStyleIndex(_T("SelectedPredecessor"));
            InsertPredecessor(oE.GetPredecessorObjectIndex());
		}
		break;
	}
}

void fWBSProject::ToolTipOnMouseHoverActiveganttvcctl1(LPDISPATCH e)
{
	CToolTipEventArgs oE(e);
	switch (oE.GetEventTarget())
	{
	case EVT_TASK:
	case EVT_SELECTEDTASK:
	case EVT_PERCENTAGE:
	case EVT_SELECTEDPERCENTAGE:
		{
            TaskToolTipCalculateDim(&oE);
            return;
		}
		break;
	}
	ActiveGanttVCCtl1.GetToolTip().SetVisible(FALSE);
}

void fWBSProject::OnMouseHoverToolTipDrawActiveganttvcctl1(LPDISPATCH e)
{
	CToolTipEventArgs oE(e);
	switch (oE.GetEventTarget())
	{
	case EVT_TASK:
	case EVT_SELECTEDTASK:
	case EVT_PERCENTAGE:
	case EVT_SELECTEDPERCENTAGE:
		{
            TaskToolTipDraw(&oE);
			oE.SetCustomDraw(TRUE);
            return;
		}
		break;
	}
}

void fWBSProject::ToolTipOnMouseMoveActiveganttvcctl1(LPDISPATCH e)
{
	CToolTipEventArgs oE(e);
	switch (oE.GetOperation())
	{
	case EO_PERCENTAGESIZING:
	case EO_TASKMOVEMENT:
	case EO_TASKSTRETCHLEFT:
	case EO_TASKSTRETCHRIGHT:
		{
            TaskToolTipCalculateDim(&oE);
            return;
		}
		break;
	}
	ActiveGanttVCCtl1.GetToolTip().SetVisible(FALSE);
}

void fWBSProject::OnMouseMoveToolTipDrawActiveganttvcctl1(LPDISPATCH e)
{
	CToolTipEventArgs oE(e);
	switch (oE.GetOperation())
	{
	case EO_PERCENTAGESIZING:
	case EO_TASKMOVEMENT:
	case EO_TASKSTRETCHLEFT:
	case EO_TASKSTRETCHRIGHT:
		{
            TaskToolTipDraw(&oE);
			oE.SetCustomDraw(TRUE);
            return;
		}
		break;
	}
}

void fWBSProject::TaskToolTipCalculateDim(CToolTipEventArgs* oE)
{
	int Index;
    int lHeight;
    int lWidth;
	CclsToolTip oToolTip;
	CclsClientArea oClientArea;
	CString sRowKey;
	CString sRowText;

    Index = ActiveGanttVCCtl1.GetMathLib().GetTaskIndexByPosition(oE->GetX(), oE->GetY());
    oToolTip = ActiveGanttVCCtl1.GetToolTip();
	oClientArea = ActiveGanttVCCtl1.GetCurrentViewObject().GetClientArea();
	if (Index == -1)
	{
	    sRowKey = ActiveGanttVCCtl1.GetRows().Item(CStr(ActiveGanttVCCtl1.GetMathLib().GetRowIndexByPosition(oE->GetY()))).GetKey();
	}
	else
	{
	    sRowKey = ActiveGanttVCCtl1.GetTasks().Item(CStr(Index)).GetRowKey();
	}
    sRowText = ActiveGanttVCCtl1.GetRows().Item(sRowKey).GetText();
    ActiveGanttVCCtl1.GetToolTipEventArgs().GetGraphics().SetLayoutRect(0, 0, 0, 0);
    lHeight = ActiveGanttVCCtl1.GetToolTipEventArgs().GetGraphics().StringHeight(sRowText, ActiveGanttVCCtl1.GetToolTip().GetFont(), (long)this->m_hWnd);
    lWidth = ActiveGanttVCCtl1.GetToolTipEventArgs().GetGraphics().StringWidth(sRowText, ActiveGanttVCCtl1.GetToolTip().GetFont(), (long)this->m_hWnd);
	if (lWidth > 250)
	{
		lHeight = lHeight * 2;
	}
    oToolTip.SetAutomaticSizing(FALSE);
    oToolTip.SetLeft(oE->GetX() + 20);
    oToolTip.SetTop(oE->GetY() - (lHeight + 60) - 20);
    oToolTip.SetWidth(300);
    oToolTip.SetHeight(lHeight + 60);
    if (oClientArea.GetWidth() < oToolTip.GetWidth())
	{
        oToolTip.SetVisible(FALSE);
		return;
    }
    if (oClientArea.GetHeight() < oToolTip.GetHeight())
	{
        oToolTip.SetVisible(FALSE);
		return;
    }
    if (oToolTip.GetLeft() < oClientArea.GetLeft())
	{
		oToolTip.SetLeft(oClientArea.GetLeft());
    }
    if (oToolTip.GetTop() < oClientArea.GetTop())
	{
		oToolTip.SetTop(oClientArea.GetTop());
    }
    if ((oToolTip.GetLeft() + oToolTip.GetWidth()) > oClientArea.GetRight())
	{
		oToolTip.SetLeft(oClientArea.GetRight() - oToolTip.GetWidth());
    }
    if ((oToolTip.GetTop() + oToolTip.GetHeight()) > oClientArea.GetBottom())
	{
		oToolTip.SetTop(oClientArea.GetBottom() - oToolTip.GetHeight());
    }
    oToolTip.SetVisible(TRUE);
}

void fWBSProject::TaskToolTipDraw(CToolTipEventArgs* oE)
{
    int Index;
    CString sRowKey;
    CString sTaskKey;
    COleDateTime dtStartDate;
    COleDateTime dtEndDate;
    float fPercentage = 0;
    CclsPercentage oPercentage;
    oPercentage = NULL;
    CclsTask oTask;
	CString sStartDate;
	CString sEndDate;
	CString sFrom;
	CString sDuration;
	CString sRowText;
	CString sPercentage = _T("");
	CclsImage oImage;
	CString sImagePath;
    int lHeight;
	if (oE->GetToolTipType() == TPT_HOVER)
	{
        Index = ActiveGanttVCCtl1.GetMathLib().GetTaskIndexByPosition(oE->GetX(), oE->GetY());
		if (Index < 1)
		{
			return;
        }
		oTask = ActiveGanttVCCtl1.GetTasks().Item(CStr(Index));
		sRowKey = oTask.GetRowKey();
        dtStartDate = oTask.GetStartDate();
        dtEndDate = oTask.GetEndDate();
        sTaskKey = oTask.GetKey();
		oPercentage = GetPercentageByTaskKey(sTaskKey);
        if (oPercentage != NULL)
		{
            fPercentage = GetPercentageByTaskKey(sTaskKey).GetPercent() * 100;
		}
	}
	else
	{
        Index = oE->GetTaskIndex();
        if (oE->GetOperation() == EO_TASKMOVEMENT )
		{
            sRowKey = ActiveGanttVCCtl1.GetRows().Item(CStr(oE->GetInitialRowIndex())).GetKey();
		}
		else
		{

            sRowKey = ActiveGanttVCCtl1.GetRows().Item(CStr(oE->GetRowIndex())).GetKey();
		}
        dtStartDate = oE->GetStartDate();
        dtEndDate = oE->GetEndDate();
        sTaskKey = ActiveGanttVCCtl1.GetTasks().Item(CStr(Index)).GetKey();
        if (oE->GetOperation() == EO_PERCENTAGESIZING)
		{
            fPercentage = (float) (oE->GetX() - oE->GetXStart()) / (oE->GetXEnd() - oE->GetXStart()) * 100;
		}
		else
		{
            if (oPercentage != NULL)
			{

                fPercentage = oPercentage.GetPercent() * 100;
			}
		}
	}
    sStartDate = FormatDate(dtStartDate, _T("%A %B %d, %Y")); //"ddd MMM d, yyyy"
    sEndDate = FormatDate(dtEndDate, _T("%A %B %d, %Y")); //"ddd MMM d, yyyy"
    sFrom = _T("From: ") + sStartDate + _T(" To ") + sEndDate;
    sDuration = _T("Duration: ") + CStr(ActiveGanttVCCtl1.GetMathLib().DateTimeDiff(IL_DAY, dtStartDate, dtEndDate)) + _T(" days");
    sRowText = ActiveGanttVCCtl1.GetRows().Item(sRowKey).GetText();
	if (oPercentage != NULL)
	{
		sPercentage = oPercentage.ToString();
	}
    oImage = ActiveGanttVCCtl1.GetRows().Item(sRowKey).GetNode().GetImage();

    oE->GetGraphics().SetLayoutRect(0, 0, oE->GetX2() - oE->GetX1() - 50, oE->GetY2() - oE->GetY1());
    lHeight = ActiveGanttVCCtl1.GetToolTipEventArgs().GetGraphics().StringHeight(sRowText, ActiveGanttVCCtl1.GetToolTip().GetFont(), (long)this->m_hWnd);
    ActiveGanttVCCtl1.GetDrawing().PaintImage(oImage, 2, 2, 17, 17);
    ActiveGanttVCCtl1.GetDrawing().GetTextFlags().SetHorizontalAlignment(HAL_CENTER);
    ActiveGanttVCCtl1.GetDrawing().GetTextFlags().SetVerticalAlignment(VAL_TOP);
    ActiveGanttVCCtl1.GetDrawing().GetTextFlags().SetWordWrap(TRUE);
    ActiveGanttVCCtl1.GetDrawing().DrawText(18, 0, oE->GetX2() - oE->GetX1(), oE->GetY2() - oE->GetY1(), sRowText, CLR_BLACK, ActiveGanttVCCtl1.GetToolTip().GetFont());
    ActiveGanttVCCtl1.GetDrawing().DrawLine(0, lHeight + 6, oE->GetX2(), lHeight + 6, CLR_BLACK, LDS_SOLID, 2);
    ActiveGanttVCCtl1.GetDrawing().DrawAlignedText(2, lHeight + 10, 300, lHeight + 25, sDuration, HAL_LEFT, VAL_CENTER, CLR_BLACK, ActiveGanttVCCtl1.GetToolTip().GetFont());
	ActiveGanttVCCtl1.GetDrawing().DrawAlignedText(2, lHeight + 25, 300, lHeight + 40, sFrom, HAL_LEFT, VAL_CENTER, CLR_BLACK, ActiveGanttVCCtl1.GetToolTip().GetFont());
    ActiveGanttVCCtl1.GetDrawing().DrawAlignedText(2, lHeight + 40, 300, lHeight + 55, _T("Percent Completed: ") + sPercentage + _T("%"), HAL_LEFT, VAL_CENTER, CLR_BLACK, ActiveGanttVCCtl1.GetToolTip().GetFont());
}

// Operations

void fWBSProject::InsertPredecessor(LONG PredecessorObjectIndex)
{
	CclsPredecessor oPredecessor;
	CString sPredecessorTaskKey;
	CString sSuccessorTaskKey;
	oPredecessor = ActiveGanttVCCtl1.GetPredecessors().Item(CStr(PredecessorObjectIndex));
	sPredecessorTaskKey = Replace(oPredecessor.GetPredecessorTask().GetKey(), _T("T"), _T(""));
	sSuccessorTaskKey = Replace(oPredecessor.GetSuccessorTask().GetKey(), _T("T"), _T(""));
	mp_oConn.ExecuteSQL(_T("INSERT INTO tb_GuysStThomas_Predecessors (lPredecessorID, lSuccessorID, lType, lLagFactor, lLagInterval) VALUES (" + sPredecessorTaskKey + _T(",") + sSuccessorTaskKey + _T(",") + CStr((int)oPredecessor.GetPredecessorType()) + _T(",") + CStr((int)oPredecessor.GetLagFactor())) + _T(",") + CStr((int)oPredecessor.GetLagInterval()) + _T(")"));
}

void fWBSProject::UpdateTask(LONG Index)
{
    CclsPercentage oPercentage = GetPercentageByTaskKey(ActiveGanttVCCtl1.GetTasks().Item(CStr(Index)).GetKey());
	CclsTask oTask = ActiveGanttVCCtl1.GetTasks().Item(CStr(Index));
	SetTaskGridColumns(&oTask);
	CString sRowKey = oTask.GetRowKey();
	COleDateTime dtStartDate = oTask.GetStartDate();
	COleDateTime dtEndDate = oTask.GetEndDate();
    CclsNode oNode = ActiveGanttVCCtl1.GetRows().Item(sRowKey).GetNode();
	CString sSQL;
	sSQL = _T("UPDATE tb_GuysStThomas SET dtStartDate = ") +
	g_DST_ACCESS_ConvertDate(dtStartDate) +
	_T(", dtEndDate = ") +
	g_DST_ACCESS_ConvertDate(dtEndDate) +
	_T(", fPercentCompleted = ") +
	CStr(oPercentage.GetPercent()) +
	_T(" WHERE lTaskID = ") + Replace(sRowKey, _T("K"), _T(""));
	mp_oConn.ExecuteSQL(sSQL);
    UpdateSummary(&oNode);
}

void fWBSProject::UpdateSummary(CclsNode* oNode)
{
    CString sSQL;
    CclsNode oParentNode;
    CclsTask oSummaryTask;
    CclsPercentage oSummaryPercentage;
	CclsRow oRow;
    oParentNode = oNode->Parent();
	while (oParentNode != NULL)
	{
		oRow = oParentNode.GetRow();
		oSummaryTask = GetTaskByRowKey(oRow.GetKey());
        oSummaryPercentage = GetPercentageByTaskKey(oSummaryTask.GetKey());
		if (oSummaryTask != NULL)
		{
            CclsTask oChildTask;
            CclsPercentage oChildPercentage; 
            CclsNode oChildNode;
            COleDateTime dtSumStartDate;
            COleDateTime dtSumEndDate;
            int lPercentagesCount = 0;
            FLOAT fPercentagesSum = 0;
            FLOAT fPercentageAvg = 0;
            oChildNode = oParentNode.Child();
			while (oChildNode != NULL)
			{
				oRow = oChildNode.GetRow();
                oChildTask = GetTaskByRowKey(oRow.GetKey());
                oChildPercentage = GetPercentageByTaskKey(oChildTask.GetKey());
                lPercentagesCount = lPercentagesCount + 1;
                fPercentagesSum = fPercentagesSum + oChildPercentage.GetPercent();
				if (oChildTask != NULL)
				{
					if (dtSumStartDate == 0)
					{
						dtSumStartDate = oChildTask.GetStartDate();
					}
					else
					{
                        if (oChildTask.GetStartDate() < dtSumStartDate)
						{
                            dtSumStartDate = oChildTask.GetStartDate();
						}
					}
					if (dtSumEndDate == 0)
					{
						dtSumEndDate = oChildTask.GetEndDate();
					}
					else
					{
                        if (oChildTask.GetEndDate() > dtSumEndDate)
						{
                            dtSumEndDate = oChildTask.GetEndDate();
						}
					}
				}
                oChildNode = oChildNode.NextSibling();
			}
            fPercentageAvg = fPercentagesSum / lPercentagesCount;
            oSummaryTask.SetStartDate(dtSumStartDate);
            oSummaryTask.SetEndDate(dtSumEndDate);
			SetTaskGridColumns(&oSummaryTask);
            oSummaryPercentage.SetPercent(fPercentageAvg);
			CString sRowKey = oSummaryTask.GetRow().GetKey();
			sRowKey.Replace(_T("K"), _T(""));
			sSQL = _T("UPDATE tb_GuysStThomas SET dtStartDate = ") +
			g_DST_ACCESS_ConvertDate(dtSumStartDate) +
			_T(", dtEndDate = ") +
			g_DST_ACCESS_ConvertDate(dtSumEndDate) +
			_T(", fPercentCompleted = ") +
			CStr(oSummaryPercentage.GetPercent()) +
			_T(" WHERE lTaskID = ") + sRowKey;
			mp_oConn.ExecuteSQL(sSQL);
		}
		oParentNode = oParentNode.Parent();
	}
}

CclsTask fWBSProject::GetTaskByRowKey(CString sRowKey)
{
    int i;
    CclsTask oTask;
	for (i = 1; i <= ActiveGanttVCCtl1.GetTasks().GetCount(); i++)
	{
		oTask = ActiveGanttVCCtl1.GetTasks().Item(CStr(i));
        if (oTask.GetRowKey() == sRowKey)
		{
            return oTask;
		}
	}
	return NULL;
}

CclsPercentage fWBSProject::GetPercentageByTaskKey(CString sTaskKey)
{
	int i;
	CclsPercentage oPercentage;
	for (i = 1; i <= ActiveGanttVCCtl1.GetPercentages().GetCount(); i++)
	{
		oPercentage = ActiveGanttVCCtl1.GetPercentages().Item(CStr(i));
		if (oPercentage.GetTaskKey() == sTaskKey)
		{
            return oPercentage;
		}
	}
	
	return NULL;
}

void fWBSProject::mnuSaveXML_Click()
{
	SaveXML();
}

void fWBSProject::mnuLoadXML_Click()
{
	LoadXML();
}

void fWBSProject::mnuPrint_Click()
{
	Print();
}

void fWBSProject::mnuClose_Click()
{
	CDialog::EndDialog(IDOK);
}

void fWBSProject::mnuCheckboxes_Click()
{
	CMenu* oMenu = CWnd::GetMenu();
	UINT lState = oMenu->GetMenuState(ID_TREEVIEWPROPERTIES_CHECKBOXES, MF_BYCOMMAND);
	if (lState & MF_CHECKED)
	{
		oMenu->CheckMenuItem(ID_TREEVIEWPROPERTIES_CHECKBOXES, MF_UNCHECKED | MF_BYCOMMAND);
		ActiveGanttVCCtl1.GetTreeview().SetCheckBoxes(FALSE);
	}
	else
	{
		oMenu->CheckMenuItem(ID_TREEVIEWPROPERTIES_CHECKBOXES, MF_CHECKED | MF_BYCOMMAND);
		ActiveGanttVCCtl1.GetTreeview().SetCheckBoxes(TRUE);
	}
	ActiveGanttVCCtl1.Redraw();
}

void fWBSProject::mnuPictures_Click()
{
	CMenu* oMenu = CWnd::GetMenu();
	UINT lState = oMenu->GetMenuState(ID_TREEVIEWPROPERTIES_IMAGES, MF_BYCOMMAND);
	if (lState & MF_CHECKED)
	{
		oMenu->CheckMenuItem(ID_TREEVIEWPROPERTIES_IMAGES, MF_UNCHECKED | MF_BYCOMMAND);
		ActiveGanttVCCtl1.GetTreeview().SetImages(FALSE);
	}
	else
	{
		oMenu->CheckMenuItem(ID_TREEVIEWPROPERTIES_IMAGES, MF_CHECKED | MF_BYCOMMAND);
		ActiveGanttVCCtl1.GetTreeview().SetImages(TRUE);
	}
	ActiveGanttVCCtl1.Redraw();
}

void fWBSProject::mnuFullColumnSelect_Click()
{
	CMenu* oMenu = CWnd::GetMenu();
	UINT lState = oMenu->GetMenuState(ID_TREEVIEWPROPERTIES_FULLCOLUMNSELECT, MF_BYCOMMAND);
	if (lState & MF_CHECKED)
	{
		oMenu->CheckMenuItem(ID_TREEVIEWPROPERTIES_FULLCOLUMNSELECT, MF_UNCHECKED | MF_BYCOMMAND);
		ActiveGanttVCCtl1.GetTreeview().SetFullColumnSelect(FALSE);
	}
	else
	{
		oMenu->CheckMenuItem(ID_TREEVIEWPROPERTIES_FULLCOLUMNSELECT, MF_CHECKED | MF_BYCOMMAND);
		ActiveGanttVCCtl1.GetTreeview().SetFullColumnSelect(TRUE);
	}
	ActiveGanttVCCtl1.Redraw();
}

void fWBSProject::cmdSaveXML()
{
	SaveXML();
}

void fWBSProject::cmdLoadXML()
{
	LoadXML();
}

void fWBSProject::cmdPrint_Click()
{
	Print();
}

void fWBSProject::cmdZoomIn_Click()
{
	if (CInt32(ActiveGanttVCCtl1.GetCurrentView()) < 3)
	{
		ActiveGanttVCCtl1.SetCurrentView(CStr(CInt32(ActiveGanttVCCtl1.GetCurrentView()) + 1));
		ActiveGanttVCCtl1.Redraw();
	}
}

void fWBSProject::cmdZoomOut_Click()
{
	if (CInt32(ActiveGanttVCCtl1.GetCurrentView()) > 1)
	{
		ActiveGanttVCCtl1.SetCurrentView(CStr(CInt32(ActiveGanttVCCtl1.GetCurrentView()) - 1));
		ActiveGanttVCCtl1.Redraw();
	}
}

void fWBSProject::cmdBluePercentages_Click()
{
	long i;
	CclsPercentage oPercentage;
	if (mp_bBluePercentagesVisible == TRUE)
	{
		mp_bBluePercentagesVisible = FALSE;
	}
	else
	{
		mp_bBluePercentagesVisible = TRUE;
	}
	for (i = 1; i <= ActiveGanttVCCtl1.GetPercentages().GetCount(); i++)
	{
		oPercentage = ActiveGanttVCCtl1.GetPercentages().Item(CStr(i));
		if (oPercentage.GetStyleIndex() == _T("P1"))
		{
			oPercentage.SetVisible(mp_bBluePercentagesVisible);
		}
	}
    ActiveGanttVCCtl1.Redraw();
}

void fWBSProject::cmdGreenPercentages_Click()
{
	long i;
	CclsPercentage oPercentage;
	if (mp_bGreenPercentagesVisible == TRUE)
	{
		mp_bGreenPercentagesVisible = FALSE;
	}
	else
	{
		mp_bGreenPercentagesVisible = TRUE;
	}
	for (i = 1; i <= ActiveGanttVCCtl1.GetPercentages().GetCount(); i++)
	{
		oPercentage = ActiveGanttVCCtl1.GetPercentages().Item(CStr(i));
		if (oPercentage.GetStyleIndex() == _T("SP2"))
		{
			oPercentage.SetVisible(mp_bGreenPercentagesVisible);
		}
	}
    ActiveGanttVCCtl1.Redraw();
}

void fWBSProject::cmdRedPercentages_Click()
{
	long i;
	CclsPercentage oPercentage;
	if (mp_bRedPercentagesVisible == TRUE)
	{
		mp_bRedPercentagesVisible = FALSE;
	}
	else
	{
		mp_bRedPercentagesVisible = TRUE;
	}
	for (i = 1; i <= ActiveGanttVCCtl1.GetPercentages().GetCount(); i++)
	{
		oPercentage = ActiveGanttVCCtl1.GetPercentages().Item(CStr(i));
		if (oPercentage.GetStyleIndex() == _T("SP3"))
		{
			oPercentage.SetVisible(mp_bRedPercentagesVisible);
		}
	}
    ActiveGanttVCCtl1.Redraw();
}

void fWBSProject::cmdProperties_Click()
{
	fWBSPProperties oForm(this);
	oForm.DoModal();
}

void fWBSProject::cmdCheckPredecessors_Click()
{
	ActiveGanttVCCtl1.CheckPredecessors();
	ActiveGanttVCCtl1.Redraw();
}

void fWBSProject::cmdToolTips_Click()
{
	if (ActiveGanttVCCtl1.GetCurrentViewObject().GetClientArea().GetToolTipsVisible() == FALSE)
	{
		ActiveGanttVCCtl1.GetCurrentViewObject().GetClientArea().SetToolTipsVisible(TRUE);
	}
	else
	{
		ActiveGanttVCCtl1.GetCurrentViewObject().GetClientArea().SetToolTipsVisible(FALSE);
	}
    ActiveGanttVCCtl1.Redraw();
}

void fWBSProject::cmdHelp_Click()
{
	g_ShowInBrowser(_T("http://www.sourcecodestore.com/Article.aspx?ID=17"));
}

// Toolbar And Menu Common Functions

void fWBSProject::SaveXML()
{
	TCHAR sDefaultFileName[MAX_PATH];
	wcscpy_s(sDefaultFileName, MAX_PATH, _T("AGVC_WBSP.xml"));
	CFileDialog oDialog(FALSE, NULL, NULL, OFN_OVERWRITEPROMPT,_T("XML File|*.xml||"), this, NULL);
	oDialog.m_ofn.lpstrTitle = _T("Save As XML");
	oDialog.m_ofn.lpstrFile = sDefaultFileName;
	oDialog.m_ofn.nMaxFile = MAX_PATH;
	if (oDialog.DoModal() == IDOK)
	{
		CString sFilePath = oDialog.GetPathName();
		ActiveGanttVCCtl1.WriteXML(sFilePath);
	}
}

void fWBSProject::LoadXML()
{
	fLoadXML oForm;
	oForm.DoModal();
}

void fWBSProject::Print()
{
	if (ActiveGanttVCCtl1.GetPrinterErrorMessage().GetLength() == 0)
	{
		fPrintDialog oForm(GetDateTime(2006, 8, 1, 0, 0, 0), GetDateTime(2010, 1, 1, 0, 0, 0));
		oForm.mp_oControl = &ActiveGanttVCCtl1;
		oForm.DoModal();
	}
}

// Data Loading

void fWBSProject::Access_LoadTasks()
{
	CclsRow oRow;
	CclsTask oTask;
	CRecordset oRecordset(&mp_oConn);
	oRecordset.Open(CRecordset::forwardOnly, _T("SELECT * FROM tb_GuysStThomas"), CRecordset::readOnly);
	while(oRecordset.IsEOF() == FALSE)
	{
		CString sDescription, sTaskType;
		int lID, lDepth;
		COleDateTime dtStartDate, dtEndDate;
		BOOL bSummary, bHasTasks;
		float fPercentCompleted;
		sDescription = GetStrField(oRecordset, _T("sDescription"));
		lID = GetIntField(oRecordset, _T("lTaskID"));
		sTaskType = GetStrField(oRecordset, _T("sTaskType"));
		lDepth = GetIntField(oRecordset, _T("lDepth"));
		dtStartDate = GetDateField(oRecordset, _T("dtStartDate"));
		dtEndDate = GetDateField(oRecordset, _T("dtEndDate"));
		bSummary = GetBoolField(oRecordset, _T("bSummary"));
		bHasTasks = GetBoolField(oRecordset, _T("bHasTasks"));
		fPercentCompleted = GetFloatField(oRecordset, _T("fPercentCompleted"));

		oRow = ActiveGanttVCCtl1.GetRows().Add(_T("K") + CStr(lID), sDescription, FALSE, TRUE, _T(""));
		oRow.GetCells().Item(_T("1")).SetText(CStr(lID));
		oRow.GetCells().Item(_T("1")).SetStyleIndex(_T("CH"));
		oRow.SetHeight(20);
		oRow.GetNode().SetAllowTextEdit(TRUE);
		if (sTaskType == _T("F"))
		{
			if (lDepth == 0)
			{
				oRow.GetNode().GetImage().FromFile(g_GetAppLocation() + _T("\\res\\folderclosed.bmp"), -1);
				oRow.GetNode().GetExpandedImage().FromFile(g_GetAppLocation() + _T("\\res\\folderopen.bmp"), -1);
				oRow.GetNode().SetStyleIndex(_T("NB"));
			}
			else
			{
				oRow.GetNode().GetImage().FromFile(g_GetAppLocation() + _T("\\res\\modules.bmp"), -1);
				oRow.GetNode().SetStyleIndex(_T("NR"));
			}
		}
		else if (sTaskType == _T("A"))
		{
			oRow.GetNode().SetStyleIndex(_T("NR"));
			oRow.GetNode().GetImage().FromFile(g_GetAppLocation() + _T("\\res\\task.bmp"), -1);
			oRow.GetNode().SetCheckBoxVisible(TRUE);
		}
		oRow.GetNode().SetDepth(lDepth);
		oRow.GetNode().SetImageVisible(TRUE);
		oRow.GetNode().SetAllowTextEdit(TRUE);

		if (mp_dtStartDate.m_dt == (DATE)0)
		{
			mp_dtStartDate = dtStartDate;
		}
		else
		{
			if (dtStartDate < mp_dtStartDate)
			{
				mp_dtStartDate = dtStartDate;
			}
		}
		if (mp_dtEndDate.m_dt == (DATE)0)
		{
			mp_dtEndDate = dtEndDate;
		}
		else
		{
			if (dtEndDate > mp_dtEndDate)
			{
				mp_dtEndDate = dtEndDate;
			}
		}
		oTask = ActiveGanttVCCtl1.GetTasks().Add(_T(""), _T("K") + CStr(lID), dtStartDate, dtEndDate, _T("T") + CStr(lID), _T(""), _T(""));
		SetTaskGridColumns(&oTask);
		if (bSummary == TRUE)
		{
			// Prevent user from moving/sizing summary tasks
			oTask.SetAllowedMovement(MT_MOVEMENTDISABLED);
			oTask.SetAllowStretchLeft(FALSE);
			oTask.SetAllowStretchRight(FALSE);
			// Prevent user from adding tasks in these Rows
			oRow.SetContainer(FALSE);
			// Apply Summary Style
			if (oRow.GetNode().GetDepth() == 0)
			{
				oTask.SetStyleIndex(_T("S3"));
				ActiveGanttVCCtl1.GetPercentages().Add(_T("T") + CStr(lID), _T("SP3"), fPercentCompleted, _T(""));
			}
			else if (oRow.GetNode().GetDepth() == 1)
			{
				oTask.SetStyleIndex(_T("S2"));
				ActiveGanttVCCtl1.GetPercentages().Add(_T("T") + CStr(lID), _T("SP2"), fPercentCompleted, _T(""));
			}

			ActiveGanttVCCtl1.GetPercentages().Item(CStr(ActiveGanttVCCtl1.GetPercentages().GetCount())).SetAllowSize(FALSE);
		}
		else
		{
			oTask.SetAllowedMovement(MT_RESTRICTEDTOROW);
			oTask.SetStyleIndex(_T("T1"));
			oTask.SetWarningStyleIndex(_T("TW1"));
			if (bHasTasks == FALSE)
			{
				oTask.SetVisible(FALSE);
				// Prevent user from adding tasks in these rows
				oRow.SetContainer(FALSE);
				ActiveGanttVCCtl1.GetPercentages().Add(_T("T") + CStr(lID), _T("InvisiblePercentages"), fPercentCompleted, _T(""));
				ActiveGanttVCCtl1.GetPercentages().Item(CStr(ActiveGanttVCCtl1.GetPercentages().GetCount())).SetAllowSize(FALSE);
			}
			else
			{
				ActiveGanttVCCtl1.GetPercentages().Add(_T("T") + CStr(lID), _T("P1"), fPercentCompleted, _T(""));
			}
		}
		oRecordset.MoveNext();
	}
	oRecordset.Close();
	oRecordset.Open(CRecordset::forwardOnly, _T("SELECT * FROM tb_GuysStThomas_Predecessors"), CRecordset::readOnly);
	while(oRecordset.IsEOF() == FALSE)
	{
		int lSuccessorTaskID, lPredecessorTaskID, lType, lLagFactor, lLagInterval;
		lSuccessorTaskID = GetIntField(oRecordset, _T("lSuccessorTaskID"));
		lPredecessorTaskID = GetIntField(oRecordset, _T("lPredecessorTaskID"));
		lType = GetIntField(oRecordset, _T("lType"));
		lLagFactor = GetIntField(oRecordset, _T("lLagFactor"));
		lLagInterval = GetIntField(oRecordset, _T("lLagInterval"));
		CclsPredecessor oPredecessor;
		oPredecessor = ActiveGanttVCCtl1.GetPredecessors().Add(_T("T") + CStr(lSuccessorTaskID), _T("T") + CStr(lPredecessorTaskID), (E_CONSTRAINTTYPE)lType, _T(""), _T("T1"));
        oPredecessor.SetLagFactor(lLagFactor);
        oPredecessor.SetLagInterval((E_INTERVAL)lLagInterval);
        oPredecessor.SetWarningStyleIndex(_T("TW1"));
        oPredecessor.SetSelectedStyleIndex(_T("SelectedPredecessor"));
		oRecordset.MoveNext();
	}
	oRecordset.Close();
}

void fWBSProject::AddPredecessor(int lPredecessorID, int lSuccessorID, int lType, int lLagFactor, int yLagInterval)
{
	CclsPredecessor oPredecessor;
	oPredecessor = ActiveGanttVCCtl1.GetPredecessors().Add(_T("T") + CStr(lSuccessorID), _T("T") + CStr(lPredecessorID), (E_CONSTRAINTTYPE)lType, _T(""), _T("T1"));
    oPredecessor.SetWarningStyleIndex(_T("TW1"));
    oPredecessor.SetSelectedStyleIndex(_T("SelectedPredecessor"));
    oPredecessor.SetLagFactor(lLagFactor);
    oPredecessor.SetLagInterval((E_INTERVAL)yLagInterval);

}

void fWBSProject::AddRow_Task(int lID, int lDepth, CString sTaskType, CString sDescription, COleDateTime dtStartDate, COleDateTime dtEndDate, float fPercentCompleted, BOOL bSummary, BOOL bHasTasks)
{
	CclsRow oRow;
	CclsTask oTask;

    oRow = ActiveGanttVCCtl1.GetRows().Add(_T("K") + CStr(lID), sDescription, TRUE, TRUE, _T(""));
	oRow.GetCells().Item(_T("1")).SetText(CStr(lID));
	oRow.GetCells().Item(_T("1")).SetStyleIndex(_T("CH"));
	oRow.SetHeight(20);
	oRow.GetNode().SetAllowTextEdit(TRUE);
	if (sTaskType == _T("F"))
	{
		if (lDepth == 0)
		{
			oRow.GetNode().GetImage().FromFile(g_GetAppLocation() + _T("\\folderclosed.bmp"), -1);
			oRow.GetNode().GetExpandedImage().FromFile(g_GetAppLocation() + _T("\\folderopen.bmp"), -1);
			oRow.GetNode().SetStyleIndex(_T("NB"));
		}
		else
		{
			oRow.GetNode().GetImage().FromFile(g_GetAppLocation() + _T("\\modules.bmp"), -1);
			oRow.GetNode().SetStyleIndex(_T("NR"));
		}
	}
	else if (sTaskType == _T("A"))
	{
		oRow.GetNode().SetStyleIndex(_T("NR"));
        oRow.GetNode().GetImage().FromFile(g_GetAppLocation() + _T("\\task.bmp"), -1);
		oRow.GetNode().SetCheckBoxVisible(TRUE);
	}
    oRow.GetNode().SetDepth(lDepth);
    oRow.GetNode().SetImageVisible(TRUE);
	oRow.GetNode().SetAllowTextEdit(TRUE);
	if (mp_dtStartDate.m_dt == (DATE)0)
	{
		mp_dtStartDate = dtStartDate;
	}
	else
	{
		if (dtStartDate < mp_dtStartDate)
		{
			mp_dtStartDate = dtStartDate;
		}
	}
	if (mp_dtEndDate.m_dt == (DATE)0)
	{
		mp_dtEndDate = dtEndDate;
	}
	else
	{
		if (dtEndDate > mp_dtEndDate)
		{
			mp_dtEndDate = dtEndDate;
		}
	}
    oTask = ActiveGanttVCCtl1.GetTasks().Add(_T(""), _T("K") + CStr(lID), dtStartDate, dtEndDate, _T("T") + CStr(lID), _T(""), _T(""));
	SetTaskGridColumns(&oTask);
	if (bSummary == TRUE)
	{
	    // Prevent user from moving/sizing summary tasks
	    oTask.SetAllowedMovement(MT_MOVEMENTDISABLED);
	    oTask.SetAllowStretchLeft(FALSE);
	    oTask.SetAllowStretchRight(FALSE);
	    // Prevent user from adding tasks in these Rows
	    oRow.SetContainer(FALSE);
	    // Apply Summary Style
		if (oRow.GetNode().GetDepth() == 0)
		{
			oTask.SetStyleIndex(_T("S3"));
			ActiveGanttVCCtl1.GetPercentages().Add(_T("T") + CStr(lID), _T("SP3"), fPercentCompleted, _T(""));
		}
		else if (oRow.GetNode().GetDepth() == 1)
		{
	        oTask.SetStyleIndex(_T("S2"));
	        ActiveGanttVCCtl1.GetPercentages().Add(_T("T") + CStr(lID), _T("SP2"), fPercentCompleted, _T(""));
		}

		ActiveGanttVCCtl1.GetPercentages().Item(CStr(ActiveGanttVCCtl1.GetPercentages().GetCount())).SetAllowSize(FALSE);
	}
	else
	{
		oTask.SetAllowedMovement(MT_RESTRICTEDTOROW);
		oTask.SetStyleIndex(_T("T1"));
		oTask.SetWarningStyleIndex(_T("TW1"));
		if (bHasTasks == FALSE)
		{
			oTask.SetVisible(FALSE);
			// Prevent user from adding tasks in these rows
			oRow.SetContainer(FALSE);
			ActiveGanttVCCtl1.GetPercentages().Add(_T("T") + CStr(lID), _T("InvisiblePercentages"), fPercentCompleted, _T(""));
			ActiveGanttVCCtl1.GetPercentages().Item(CStr(ActiveGanttVCCtl1.GetPercentages().GetCount())).SetAllowSize(FALSE);
		}
		else
		{
	        ActiveGanttVCCtl1.GetPercentages().Add(_T("T") + CStr(lID), _T("P1"), fPercentCompleted, _T(""));
		}
	}
}

void fWBSProject::EndTextEditActiveganttvcctl1(LPDISPATCH e)
{
		CTextEditEventArgs oE(e);
		if (oE.GetObjectType() == TOT_NODE)
		{
			CclsRow oRow;
			oRow = ActiveGanttVCCtl1.GetRows().Item(CStr(oE.GetObjectIndex()));
			CString sRowKey = oRow.GetKey();
			sRowKey.Replace(_T("K"), _T(""));
			mp_oConn.ExecuteSQL(_T("UPDATE tb_GuysStThomas SET sDescription='") + oE.GetText() + _T("' WHERE lTaskID = ") + sRowKey);
		}
}

void fWBSProject::SetTaskGridColumns(CclsTask* oTask)
{
    oTask->GetRow().GetCells().Item(_T("3")).SetText(FormatDate(oTask->GetStartDate(), _T("%m/%d/%Y"))); //"MM/dd/yyyy"
    oTask->GetRow().GetCells().Item(_T("4")).SetText(FormatDate(oTask->GetEndDate(), _T("%m/%d/%Y"))); //"MM/dd/yyyy"
}


void fWBSProject::PagePrintActiveganttvcctl1(LPDISPATCH e)
{
		CPagePrintEventArgs oE(e);

		CclsFont oFont;
		oFont.CreateDispatch(_T("AGVC.clsFont"));
		oFont.SetFontName(_T("Arial"));
		oFont.SetSize(8);


        ULONG oColor;
        LONG X2;
        oColor = ActiveGanttVCCtl1.GetConfiguration().GetDefaultControlStyle().GetBorderColor();
		if ((oE.GetX2() - oE.GetX1()) < CInt32(CDbl(oE.GetPageWidth()) * 0.85))
		{
			X2 = oE.GetPageWidth() - ActiveGanttVCCtl1.GetPrinter().GetMarginRight();
		}
		else
		{
			X2 = oE.GetX2();
		}


        ActiveGanttVCCtl1.GetDrawing().DrawLine(oE.GetX1(), oE.GetPageHeight() - 190, X2, oE.GetPageHeight() - 190, oColor, LDS_SOLID, 1);
        ActiveGanttVCCtl1.GetDrawing().DrawAlignedText(oE.GetX1() + 10, oE.GetPageHeight() - 190, oE.GetX1() + 290, oE.GetPageHeight() - 70, _T("ActiveGanttVC Scheduler Component\r\nWBS Project Example"), HAL_LEFT, VAL_CENTER, CLR_BLACK, oFont);

        int lLine = oE.GetPageHeight() - 180;
        ActiveGanttVCCtl1.GetDrawing().DrawAlignedText(oE.GetX1() + 310, lLine, oE.GetX1() + 400, lLine + 13, _T("Red Summary"), HAL_LEFT, VAL_TOP, CLR_BLACK, oFont);
        ActiveGanttVCCtl1.GetDrawing().DrawObject(oE.GetX1() + 450, lLine, oE.GetX1() + 600, lLine + 10, _T("S3"), _T(""), FALSE, NULL, DO_GENERAL);
        ActiveGanttVCCtl1.GetDrawing().DrawObject(oE.GetX1() + 450, lLine, oE.GetX1() + 525, lLine + 5, _T("SP3"), _T(""), FALSE, NULL, DO_GENERAL);

        lLine = oE.GetPageHeight() - 160;
        ActiveGanttVCCtl1.GetDrawing().DrawAlignedText(oE.GetX1() + 310, lLine, oE.GetX1() + 400, lLine + 13, _T("Green Summary"), HAL_LEFT, VAL_TOP, CLR_BLACK, oFont);
        ActiveGanttVCCtl1.GetDrawing().DrawObject(oE.GetX1() + 450, lLine, oE.GetX1() + 600, lLine + 10, _T("S2"), _T(""), FALSE, NULL, DO_GENERAL);
        ActiveGanttVCCtl1.GetDrawing().DrawObject(oE.GetX1() + 450, lLine, oE.GetX1() + 525, lLine + 5, _T("SP2"), _T(""), FALSE, NULL, DO_GENERAL);

        lLine = oE.GetPageHeight() - 140;
        ActiveGanttVCCtl1.GetDrawing().DrawAlignedText(oE.GetX1() + 310, lLine, oE.GetX1() + 400, lLine + 13, _T("Task"), HAL_LEFT, VAL_TOP, CLR_BLACK, oFont);
        ActiveGanttVCCtl1.GetDrawing().DrawObject(oE.GetX1() + 445, lLine, oE.GetX1() + 605, lLine + 10, _T("T1"), _T(""), FALSE, NULL, DO_GENERAL);
        ActiveGanttVCCtl1.GetDrawing().DrawObject(oE.GetX1() + 445, lLine + 3, oE.GetX1() + 525, lLine + 7, _T("P1"), _T(""), FALSE, NULL, DO_GENERAL);

        lLine = oE.GetPageHeight() - 120;
        ActiveGanttVCCtl1.GetDrawing().DrawAlignedText(oE.GetX1() + 310, lLine, oE.GetX1() + 400, lLine + 13, _T("Milestone"), HAL_LEFT, VAL_TOP, CLR_BLACK, oFont);
        ActiveGanttVCCtl1.GetDrawing().DrawObject(oE.GetX1() + 450, lLine, oE.GetX1() + 450, lLine + 10, _T("T1"), _T(""), FALSE, NULL, DO_MILESTONE);

        ActiveGanttVCCtl1.GetDrawing().DrawLine(oE.GetX1() + 300, oE.GetPageHeight() - 190, oE.GetX1() + 300, oE.GetPageHeight() - 70, oColor, LDS_SOLID, 1);

        ActiveGanttVCCtl1.GetDrawing().DrawLine(oE.GetX1(), oE.GetPageHeight() - 70, X2, oE.GetPageHeight() - 70, oColor, LDS_SOLID, 1);
        ActiveGanttVCCtl1.GetDrawing().DrawAlignedText(oE.GetX1(), oE.GetPageHeight() - 70, X2, oE.GetPageHeight() - 50, _T("Page ") + CStr(oE.GetPageNumber()), HAL_CENTER, VAL_CENTER, CLR_BLACK, oFont);

        ActiveGanttVCCtl1.GetDrawing().DrawBorder(oE.GetX1(), oE.GetY1(), X2, oE.GetPageHeight() - 50, oColor, LDS_SOLID, 1);


}


void fWBSProject::OnClose()
{
	//if (mp_dtStartDate.m_lpDispatch != NULL)
	//{
	//	mp_dtStartDate.ReleaseDispatch();
	//}
	//if (mp_dtEndDate.m_lpDispatch != NULL)
	//{
	//	mp_dtEndDate.ReleaseDispatch();
	//}
	CDialog::OnClose();
}