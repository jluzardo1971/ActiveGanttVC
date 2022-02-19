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
#include "fCarRental.h"
#include "fCarRentalVehicle.h"
#include "fCarRentalBranch.h"
#include "fCarRentalReservation.h"
#include "fLoadXML.h"
#include "fPrintDialog.h"

IMPLEMENT_DYNAMIC(fCarRental, CDialog)

void fCarRental::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	DDX_Control(pDX, IDC_ACTIVEGANTTVCCTL1, ActiveGanttVCCtl1);
}

BEGIN_MESSAGE_MAP(fCarRental, CDialog)
	ON_WM_SIZE()
	ON_WM_CREATE()

	ON_COMMAND(ID_MYTOOLBAR_BUTTONFIRST + 0, cmdSaveXML)
	ON_COMMAND(ID_MYTOOLBAR_BUTTONFIRST + 1, cmdLoadXML)
	ON_COMMAND(ID_MYTOOLBAR_BUTTONFIRST + 2, cmdPrint_Click)
	//Separator ID_MYTOOLBAR_BUTTONFIRST + 3
	ON_COMMAND(ID_MYTOOLBAR_BUTTONFIRST + 4, cmdZoomIn_Click)
	ON_COMMAND(ID_MYTOOLBAR_BUTTONFIRST + 5, cmdZoomOut_Click)
	//Separator ID_MYTOOLBAR_BUTTONFIRST + 6
	ON_COMMAND(ID_MYTOOLBAR_BUTTONFIRST + 7, cmdAddVehicle_Click)
	ON_COMMAND(ID_MYTOOLBAR_BUTTONFIRST + 8, cmdAddBranch_Click)
	//Separator ID_MYTOOLBAR_BUTTONFIRST + 9
	ON_COMMAND(ID_MYTOOLBAR_BUTTONFIRST + 10, cmdBack2_Click)
	ON_COMMAND(ID_MYTOOLBAR_BUTTONFIRST + 11, cmdBack1_Click)
	ON_COMMAND(ID_MYTOOLBAR_BUTTONFIRST + 12, cmdBack0_Click)
	ON_COMMAND(ID_MYTOOLBAR_BUTTONFIRST + 13, cmdFwd0_Click)
	ON_COMMAND(ID_MYTOOLBAR_BUTTONFIRST + 14, cmdFwd1_Click)
	ON_COMMAND(ID_MYTOOLBAR_BUTTONFIRST + 15, cmdFwd2_Click)
	//Separator ID_MYTOOLBAR_BUTTONFIRST + 16
	ON_COMMAND(ID_MYTOOLBAR_BUTTONFIRST + 17, cmdHelp_Click)

	ON_COMMAND(ID_FILE_SAVEASXML_CR, mnuSaveXML_Click)
	ON_COMMAND(ID_FILE_LOADFROMXML_CR, mnuLoadXML_Click)
	ON_COMMAND(ID_FILE_PRINT_CR, mnuPrint_Click)
	ON_COMMAND(ID_FILE_CLOSE_CR, mnuClose_Click)

	ON_COMMAND(ID_EDIT_ROW_CR, mnuEditRow_Click)
	ON_COMMAND(ID_DELETE_ROW_CR, mnuDeleteRow_Click)

	ON_COMMAND(ID_EDIT_TASK_CR, mnuEditTask_Click)
	ON_COMMAND(ID_DELETE_TASK_CR, mnuDeleteTask_Click)
	ON_COMMAND(ID_CONVERT_TO_RENTAL_CR, mnuConvertToRental_Click)

	ON_WM_CLOSE()
END_MESSAGE_MAP()

BEGIN_EVENTSINK_MAP(fCarRental, CDialog)
	ON_EVENT(fCarRental, IDC_ACTIVEGANTTVCCTL1, 13, fCarRental::CustomTierDrawActiveganttvcctl1, VTS_DISPATCH)
	ON_EVENT(fCarRental, IDC_ACTIVEGANTTVCCTL1, 15, fCarRental::ObjectAddedActiveganttvcctl1, VTS_DISPATCH)
	ON_EVENT(fCarRental, IDC_ACTIVEGANTTVCCTL1, 3, fCarRental::ControlMouseDownActiveganttvcctl1, VTS_DISPATCH)
	ON_EVENT(fCarRental, IDC_ACTIVEGANTTVCCTL1, 20, fCarRental::CompleteObjectChangeActiveganttvcctl1, VTS_DISPATCH)
	ON_EVENT(fCarRental, IDC_ACTIVEGANTTVCCTL1, 31, fCarRental::ControlMouseWheelActiveganttvcctl1, VTS_DISPATCH)
	ON_EVENT(fCarRental, IDC_ACTIVEGANTTVCCTL1, 26, fCarRental::ControlDrawActiveganttvcctl1, VTS_NONE)
END_EVENTSINK_MAP()


fCarRental::fCarRental(CWnd* pParent /*=NULL*/) : CDialog(fCarRental::IDD, pParent)
{
	mp_oConn.Open(NULL, FALSE, FALSE, g_DST_ACCESS_GetConnectionString());

	m_hIcon = AfxGetApp()->LoadIcon(IDR_MAINFRAME);
	oBrushReservation.CreateSolidBrush(RGB(153, 170, 194));
	oBrushRental.CreateSolidBrush(RGB(162, 78, 50));
	oBrushMaintenance.CreateSolidBrush(RGB(255, 77, 1));
}

fCarRental::~fCarRental()
{
	mp_oConn.Close();
}

int fCarRental::OnCreate(LPCREATESTRUCT lpCreateStruct)
{
	if (CDialog::OnCreate(lpCreateStruct) == -1)
		return -1;

	oToolBar1.CreateEx(this, TBSTYLE_FLAT, WS_CHILD | WS_VISIBLE | CBRS_TOP | CBRS_GRIPPER | CBRS_TOOLTIPS | CBRS_FLYBY | CBRS_SIZE_DYNAMIC);
	CRect rect;
	rect.top = 1;
	rect.left = 400;
	rect.right = 600;
	rect.bottom = rect.top + 250;
	drpMode.Create(CBS_DROPDOWNLIST | WS_VSCROLL | WS_VISIBLE, rect, &oToolBar1, DRP_MODE);
	mp_oFont.CreatePointFont(8, _T("MS Sans Serif"));
	drpMode.SetFont(&mp_oFont);

	oToolBarBitmap.LoadBitmapW(IDB_FCARRENTAL);
	oToolBar1.SetBitmap((HBITMAP)oToolBarBitmap.GetSafeHandle());
	oToolBar1.SetButtons(NULL, 18);  //1 Based

	oToolBar1.AddButton(0, _T("Save as XML"));
	oToolBar1.AddButton(1, _T("Load XML"));
	oToolBar1.AddButton(2, _T("Print"));
	oToolBar1.AddSeparator();
	oToolBar1.AddButton(3, _T("Zoom in"));
	oToolBar1.AddButton(4, _T("Zoom out"));
	oToolBar1.AddSeparator();
	oToolBar1.AddButton(5, _T("Add Vehicle"));
	oToolBar1.AddButton(6, _T("Add Branch"));
	oToolBar1.AddSeparator();
	oToolBar1.AddButton(7, _T(""));
	oToolBar1.AddButton(8, _T(""));
	oToolBar1.AddButton(9, _T(""));
	oToolBar1.AddButton(10, _T(""));
	oToolBar1.AddButton(11, _T(""));
	oToolBar1.AddButton(12, _T(""));
	oToolBar1.AddSeparator();
	oToolBar1.AddButton(13, _T("Help"));

	return 0;
}

void fCarRental::OnToolBar(UINT nID)
{
	nID = nID - ID_MYTOOLBAR_BUTTONFIRST;
	switch (nID)
	{
	case 0: 
		cmdSaveXML();
		break;
	}
}

BOOL fCarRental::OnInitDialog()
{
	CDialog::OnInitDialog();
	g_MaximizeWindowsDim(this);
	CWnd::ShowWindow(SW_SHOWMAXIMIZED);

	mp_lDeleteTask = 0;
	mp_yAddMode = AM_RENTAL;

	SetIcon(m_hIcon, TRUE);			// Set big icon
	SetIcon(m_hIcon, FALSE);		// Set small icon

	// Reposition the Toolbar
	CWnd::RepositionBars(AFX_IDW_CONTROLBAR_FIRST, AFX_IDW_CONTROLBAR_LAST,0);
	

	CString sCaption = _T("Vehicle Rental/Fleet Control Roster Example - ");
	sCaption = sCaption + _T("Microsoft Access data source (32bit compatible only) - ");
	sCaption = sCaption + _T("ActiveGanttVC Version: ") + ActiveGanttVCCtl1.GetVersion();
	this->SetWindowTextW(sCaption);


	CclsStyle oStyle;
	CclsView oView;
	CclsTimeBlock oTimeBlock;

	//If you open the form: Styles And Templates -> Available Templates in the main menu (fTemplates.vb)
	//you can preview all available Templates.
	//Or you can also build your own:
	//fMSProject11.vb shows you how to build a Solid Template in the InitializeAG Method.
	//fMSProject14.vb shows you how to build a Gradient Template in the InitializeAG Method.
	ActiveGanttVCCtl1.ApplyTemplate(STC_CH_HGRAD_BLUE_BELL, STO_VARIATION_1);

    oStyle = ActiveGanttVCCtl1.GetStyles().Item(_T("RET1"));
    oStyle.SetTextAlignmentHorizontal(HAL_LEFT);
    oStyle.SetTextAlignmentVertical(VAL_TOP);

    oStyle = ActiveGanttVCCtl1.GetStyles().Item(_T("RET2"));
    oStyle.SetTextAlignmentHorizontal(HAL_LEFT);
    oStyle.SetTextAlignmentVertical(VAL_TOP);

    oStyle = ActiveGanttVCCtl1.GetStyles().Item(_T("RET3"));
    oStyle.SetTextAlignmentHorizontal(HAL_LEFT);
    oStyle.SetTextAlignmentVertical(VAL_TOP);



    oStyle = ActiveGanttVCCtl1.GetConfiguration().GetDefaultColumnStyle().Clone(_T("Branch"));
    oStyle.SetTextAlignmentHorizontal(HAL_LEFT);
    oStyle.SetTextAlignmentVertical(VAL_TOP);
    oStyle.SetTextXMargin(5);
    oStyle.SetTextYMargin(5);
    oStyle.SetBorderColor(CLR_BLACK);
    oStyle.SetBorderStyle(SBR_SINGLE);
    oStyle.SetImageAlignmentHorizontal(HAL_RIGHT);
    oStyle.SetImageAlignmentVertical(VAL_BOTTOM);
    oStyle.SetImageXMargin(5);
    oStyle.SetImageYMargin(5);

    oStyle = ActiveGanttVCCtl1.GetConfiguration().GetDefaultColumnStyle().Clone(_T("BranchCA"));

    oStyle = ActiveGanttVCCtl1.GetStyles().Add(_T("SplitterStyle"));
    oStyle.SetAppearance(SA_FLAT);
    oStyle.SetBackgroundMode(FP_GRADIENT);
    oStyle.SetGradientFillMode(GDT_VERTICAL);
    oStyle.SetStartGradientColor(FromArgb(255, 109, 122, 136));
    oStyle.SetEndGradientColor(FromArgb(255, 220, 220, 220));

	ActiveGanttVCCtl1.SetControlTag(_T("CarRental"));
	ActiveGanttVCCtl1.GetColumns().Add(_T(""), _T(""), 60, _T(""));
	ActiveGanttVCCtl1.GetColumns().Add(_T(""), _T(""), 95, _T(""));
	ActiveGanttVCCtl1.GetColumns().Add(_T(""), _T(""), 250, _T(""));

	ActiveGanttVCCtl1.GetSplitter().SetPosition(380);
	ActiveGanttVCCtl1.GetSplitter().SetType(SA_STYLE);
	ActiveGanttVCCtl1.GetSplitter().SetWidth(6);
	ActiveGanttVCCtl1.GetSplitter().SetStyleIndex(_T("SplitterStyle"));

    oTimeBlock = ActiveGanttVCCtl1.GetTimeBlocks().Add(_T(""));
    oTimeBlock.SetTimeBlockType(TBT_RECURRING);
    oTimeBlock.SetRecurringType(RCT_WEEK);
    oTimeBlock.SetBaseWeekDay(WD_SATURDAY);
	COleDateTime dtBaseDate;
	dtBaseDate.SetDateTime(2013, 1, 1, 0, 0, 0);
    oTimeBlock.SetBaseDate(dtBaseDate.m_dt);
    oTimeBlock.SetDurationFactor(48);
    oTimeBlock.SetDurationInterval(IL_HOUR);

	oView = ActiveGanttVCCtl1.GetViews().Add(IL_MINUTE, 30, ST_CUSTOM, ST_CUSTOM, ST_CUSTOM, _T(""));
	oView.GetTimeLine().GetTierArea().GetUpperTier().SetHeight(17);
	oView.GetTimeLine().GetTierArea().GetUpperTier().SetInterval(IL_MONTH);
	oView.GetTimeLine().GetTierArea().GetUpperTier().SetFactor(1);
	oView.GetTimeLine().GetTierArea().GetMiddleTier().SetHeight(17);
	oView.GetTimeLine().GetTierArea().GetMiddleTier().SetInterval(IL_DAY);
	oView.GetTimeLine().GetTierArea().GetMiddleTier().SetFactor(1);
	oView.GetTimeLine().GetTierArea().GetLowerTier().SetInterval(IL_HOUR);
	oView.GetTimeLine().GetTierArea().GetLowerTier().SetFactor(12);
	oView.GetTimeLine().GetTierArea().GetLowerTier().SetHeight(17);
	oView.GetTimeLine().GetTickMarkArea().SetVisible(FALSE);
	oView.GetTimeLine().GetStyle().SetAppearance(SA_FLAT);
	oView.GetTimeLine().GetStyle().SetBackgroundMode(FP_SOLID);
	oView.GetTimeLine().GetStyle().SetBackColor(CLR_BLACK);
	oView.GetClientArea().GetGrid().SetVerticalLines(TRUE);
	oView.GetClientArea().GetGrid().SetSnapToGrid(TRUE);
	ActiveGanttVCCtl1.SetCurrentView(CStr(oView.GetIndex()));
	SetZoom(5);

	Access_LoadRowsAndTasks();
	ActiveGanttVCCtl1.GetRows().UpdateTree();

	drpMode.AddString(_T("Add Reservation Mode"));
	drpMode.AddString(_T("Add Rental Mode"));
	drpMode.AddString(_T("Add Maintenance Mode"));
	drpMode.SetCurSel(0);
	
	if (ActiveGanttVCCtl1.GetPrinterErrorMessage().GetLength() == 0)
	{
		ActiveGanttVCCtl1.GetPrinter().SetOrientation(OT_LANDSCAPE);
		ActiveGanttVCCtl1.GetPrinter().SetMarginLeft(50); //0.5"
		ActiveGanttVCCtl1.GetPrinter().SetMarginTop(50); //0.5"
		ActiveGanttVCCtl1.GetPrinter().SetMarginRight(50); //0.5"
		ActiveGanttVCCtl1.GetPrinter().SetMarginBottom(50); //0.5"
		ActiveGanttVCCtl1.GetPrinter().SetHeadingsInEveryPage(FALSE);
		ActiveGanttVCCtl1.GetPrinter().SetKeepRowsTogether(TRUE);
		ActiveGanttVCCtl1.GetPrinter().SetColumnsInEveryPage(FALSE);
		ActiveGanttVCCtl1.GetPrinter().SetKeepColumnsTogether(TRUE);
		ActiveGanttVCCtl1.GetPrinter().SetKeepTimeIntervalsTogether(FALSE);
	}

	COleDateTime dtStart;

	dtStart.SetDateTime(2009,6,9,0,0,0);

	ActiveGanttVCCtl1.GetCurrentViewObject().GetTimeLine().Position(dtStart);

	
	return TRUE;
}

void fCarRental::OnClose()
{
	CDialog::OnClose();
}

void fCarRental::OnSize(UINT nType, int cx, int cy)
{
	CDialog::OnSize(nType, cx, cy);
	g_Resize(this, &ActiveGanttVCCtl1);
}

void fCarRental::CustomTierDrawActiveganttvcctl1(LPDISPATCH e)
{
	CCustomTierDrawEventArgs oE(e);

	if (oE.GetInterval() == IL_HOUR && oE.GetFactor() == 12)
	{
		oE.SetText(UCase(FormatDate(oE.GetStartDate(), _T("%p"))));
	}
	if (oE.GetInterval() == IL_MONTH && oE.GetFactor() == 1)
	{
		oE.SetText(FormatDate(oE.GetStartDate(), _T("%B %Y")));
	}
	if (oE.GetInterval() == IL_DAY && oE.GetFactor() == 1)
	{
		oE.SetText(FormatDate(oE.GetStartDate(), _T("%a %d")));
	}

}

void fCarRental::ObjectAddedActiveganttvcctl1(LPDISPATCH e)
{
	CObjectAddedEventArgs oE(e);
	if (oE.GetEventTarget() == EVT_TASK)
	{
		CclsTask oTask;
		oTask = ActiveGanttVCCtl1.GetTasks().Item(CStr(oE.GetTaskIndex()));
		oTask.SetStyleIndex(GetAddModeStyle());
		if (GetMode() == AM_RESERVATION)
		{
			oTask.SetTag(_T("AM_RESERVATION"));
		}
		else if (GetMode() == AM_RENTAL)
		{
			oTask.SetTag(_T("AM_RENTAL"));
		}
		else if (GetMode() == AM_MAINTENANCE)
		{
			oTask.SetTag(_T("AM_MAINTENANCE"));
		}
		oTask.SetKey(_T("K") + CStr(ActiveGanttVCCtl1.GetTasks().GetCount()) + _T("New"));
		if (GetMode() == AM_RESERVATION)
		{
			fCarRentalReservation oForm(DM_ADD, this, _T(""));
			ReleaseCapture();
			oForm.DoModal();

		}
		else if (GetMode() == AM_RENTAL)
		{
			fCarRentalReservation oForm(DM_ADD, this, _T(""));
			ReleaseCapture();
			oForm.DoModal();
		}
		else if (GetMode() == AM_MAINTENANCE)
		{
			clsDB oDB(mp_oConn);
			oDB.AddParameter(_T("lRowID"), CInt32(Replace(oTask.GetRowKey(), _T("K"), _T(""))), PT_NUMERIC);
			oDB.AddParameter(_T("lMode"), 2, PT_NUMERIC);
			oDB.AddParameter(_T("dtPickUp"), oTask.GetStartDate(), PT_DATE);
			oDB.AddParameter(_T("dtReturn"), oTask.GetEndDate(), PT_DATE);
			oTask.SetKey(_T("K") + oDB.ExecuteInsert(_T("tb_CR_Rentals")));
			oTask.SetText(_T("Scheduled Maintenance"));
			oTask.SetTag(_T("AM_MAINTENANCE"));
		}
	}
}

void fCarRental::ControlMouseDownActiveganttvcctl1(LPDISPATCH e)
{
	CMouseEventArgs oE(e);
	int lIndex;

	CMenu oMenu;
	CWnd* pActiveGanttVCCtl1 = GetDlgItem(IDC_ACTIVEGANTTVCCTL1);
	RECT oRect;
	pActiveGanttVCCtl1->GetWindowRect(&oRect);


	switch (oE.GetEventTarget())
	{
	case EVT_SELECTEDROW:
	case EVT_ROW:
		{
			lIndex = ActiveGanttVCCtl1.GetMathLib().GetRowIndexByPosition(oE.GetY());
			if (oE.GetButton() == BTN_LEFT)
			{
				if ((oE.GetX() > (ActiveGanttVCCtl1.GetSplitter().GetPosition() - 20)) && (oE.GetX() < (ActiveGanttVCCtl1.GetSplitter().GetPosition() - 5)) && (oE.GetY() < (ActiveGanttVCCtl1.GetRows().Item(CStr(lIndex)).GetBottom() - 5)) && (oE.GetY() > (ActiveGanttVCCtl1.GetRows().Item(CStr(lIndex)).GetBottom() - 20)))
				{
					if (ActiveGanttVCCtl1.GetRows().Item(CStr(lIndex)).GetNode().GetExpanded() == TRUE)
					{
						ActiveGanttVCCtl1.GetRows().Item(CStr(lIndex)).GetNode().SetExpanded(FALSE);
					}
					else
					{
						ActiveGanttVCCtl1.GetRows().Item(CStr(lIndex)).GetNode().SetExpanded(TRUE);
					}
					ActiveGanttVCCtl1.Redraw();
					oE.SetCancel(TRUE);
				}
			}
			else if (oE.GetButton() == BTN_RIGHT)
			{
				mp_sEditRowKey = ActiveGanttVCCtl1.GetRows().Item(CStr(lIndex)).GetKey();
				oE.SetCancel(TRUE);
				oMenu.CreatePopupMenu();
				oMenu.AppendMenu(MF_STRING, ID_EDIT_ROW_CR, _T("Edit Row"));
				oMenu.AppendMenu(MF_SEPARATOR);
				oMenu.AppendMenu(MF_STRING, ID_DELETE_ROW_CR, _T("Delete Row"));
				::SetForegroundWindow(m_hWnd);
				TrackPopupMenu(oMenu, TPM_LEFTALIGN | TPM_RIGHTBUTTON, oE.GetX() + oRect.left, oE.GetY() + oRect.top, NULL, m_hWnd, NULL);
				::SetForegroundWindow(m_hWnd);
			}
		}
		break;
	case EVT_SELECTEDTASK:
	case EVT_TASK:
		{

			if (oE.GetButton() == BTN_RIGHT)
			{
				oE.SetCancel(TRUE);
				CString sTag;
				lIndex = ActiveGanttVCCtl1.GetMathLib().GetTaskIndexByPosition(oE.GetX(), oE.GetY());
				mp_sEditTaskKey = ActiveGanttVCCtl1.GetTasks().Item(CStr(lIndex)).GetKey();
				sTag = ActiveGanttVCCtl1.GetTasks().Item(CStr(lIndex)).GetTag();

				oMenu.CreatePopupMenu();
				if (sTag == _T("AM_RESERVATION"))
				{
					oMenu.AppendMenu(MF_STRING, ID_EDIT_TASK_CR, _T("Edit Task"));
					oMenu.AppendMenu(MF_STRING, ID_CONVERT_TO_RENTAL_CR, _T("Convert To Rental"));
					oMenu.AppendMenu(MF_SEPARATOR);
					oMenu.AppendMenu(MF_STRING, ID_DELETE_TASK_CR, _T("Delete Task"));
				}
				else if (sTag == _T("AM_RENTAL"))
				{
					oMenu.AppendMenu(MF_STRING, ID_EDIT_TASK_CR, _T("Edit Task"));
					oMenu.AppendMenu(MF_SEPARATOR);
					oMenu.AppendMenu(MF_STRING, ID_DELETE_TASK_CR, _T("Delete Task"));
				}
				else if (sTag == _T("AM_MAINTENANCE"))
				{
					oMenu.AppendMenu(MF_STRING, ID_DELETE_TASK_CR, _T("Delete Task"));
				}
				::SetForegroundWindow(m_hWnd);
				TrackPopupMenu(oMenu, TPM_LEFTALIGN | TPM_RIGHTBUTTON, oE.GetX() + oRect.left, oE.GetY() + oRect.top, NULL, m_hWnd, NULL);
				::SetForegroundWindow(m_hWnd);
			}
		}
		break;
	}

}

void fCarRental::CompleteObjectChangeActiveganttvcctl1(LPDISPATCH e)
{
	CObjectStateChangedEventArgs oE(e);
	switch (oE.GetEventTarget())
	{
	case EVT_TASK:
		{
			CclsTask oTask;
			oTask = ActiveGanttVCCtl1.GetTasks().Item(CStr(oE.GetIndex()));
			CalculateRate(&oTask);
		}
		break;
	case EVT_ROW:
		{
			if (oE.GetChangeType() == CT_MOVE) 
			{
				int i;
				CclsRow oRow;
				for (i = 1; i <= ActiveGanttVCCtl1.GetRows().GetCount(); i++)
				{
					oRow = ActiveGanttVCCtl1.GetRows().Item(CStr(i));
					mp_oConn.ExecuteSQL(_T("UPDATE tb_CR_Rows SET lOrder = ") + CStr(i) + _T(" WHERE lRowID = ") + Replace(oRow.GetKey(), _T("K"), _T("")));
				}
			}
		}
		break;
	}
}

void fCarRental::ControlMouseWheelActiveganttvcctl1(LPDISPATCH e)
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

void fCarRental::ControlDrawActiveganttvcctl1()
{
	if (mp_lDeleteTask != 0)
	{
		ActiveGanttVCCtl1.GetTasks().Remove(CStr(mp_lDeleteTask));
		mp_lDeleteTask = 0;
	}
}

HPE_ADDMODE fCarRental::GetMode()
{
	mp_yAddMode = (HPE_ADDMODE)drpMode.GetCurSel();
	return mp_yAddMode;
}

CString fCarRental::GetAddModeStyle()
{
	switch (drpMode.GetCurSel())
	{
	case 0:
		return _T("RET2");
		break;
	case 1:
		return _T("RET1");
		break;
	case 2:
		return _T("RET3");
		break;
	}
	return _T("Error");
}

int fCarRental::GetZoom()
{
	return mp_lZoom;
}

void fCarRental::SetZoom(int NewValue)
{
	if ((NewValue > 5) || (NewValue < 1))
	{
		return;
	}
    mp_lZoom = NewValue;
    CclsView oView;
	oView = ActiveGanttVCCtl1.GetCurrentViewObject();
	switch (mp_lZoom)
	{
	case 5:
		{
			oView.SetInterval(IL_MINUTE);
			oView.SetFactor(30);
			oView.GetClientArea().GetGrid().SetInterval(IL_HOUR);
			oView.GetClientArea().GetGrid().SetFactor(12);
			oView.GetTimeLine().GetTickMarkArea().SetVisible(FALSE);
		}
		break;
	case 4:
		{
			oView.SetInterval(IL_MINUTE);
			oView.SetFactor(15);
			oView.GetClientArea().GetGrid().SetInterval(IL_HOUR);
			oView.GetClientArea().GetGrid().SetFactor(6);
			oView.GetTimeLine().GetTickMarkArea().SetVisible(FALSE);
		}
		break;
	case 3:
		{
			oView.SetInterval(IL_MINUTE);
			oView.SetFactor(10);
			oView.GetClientArea().GetGrid().SetInterval(IL_HOUR);
			oView.GetClientArea().GetGrid().SetFactor(3);
			oView.GetTimeLine().GetTickMarkArea().SetVisible(FALSE);
		}
		break;
	case 2:
		{
			oView.SetInterval(IL_MINUTE);
			oView.SetFactor(5);
			oView.GetClientArea().GetGrid().SetInterval(IL_HOUR);
			oView.GetClientArea().GetGrid().SetFactor(2);
			oView.GetTimeLine().GetTickMarkArea().SetVisible(TRUE);
			oView.GetTimeLine().GetTickMarkArea().SetHeight(30);
			oView.GetTimeLine().GetTickMarkArea().GetTickMarks().Clear();
			oView.GetTimeLine().GetTickMarkArea().GetTickMarks().Add(IL_HOUR, 6, TLT_BIG, TRUE, _T("hh:mmtt"), _T(""));
			oView.GetTimeLine().GetTickMarkArea().GetTickMarks().Add(IL_HOUR, 1, TLT_SMALL, FALSE, _T(""), _T(""));
		}
		break;
	case 1:
		{
			oView.SetInterval(IL_MINUTE);
			oView.SetFactor(1);
			oView.GetClientArea().GetGrid().SetInterval(IL_MINUTE);
			oView.GetClientArea().GetGrid().SetFactor(15);
			oView.GetTimeLine().GetTickMarkArea().SetVisible(TRUE);
			oView.GetTimeLine().GetTickMarkArea().SetHeight(30);
			oView.GetTimeLine().GetTickMarkArea().GetTickMarks().Clear();
			oView.GetTimeLine().GetTickMarkArea().GetTickMarks().Add(IL_HOUR, 1, TLT_BIG, TRUE, _T("hh:mmtt"), _T(""));
		}
		break;
	}
	ActiveGanttVCCtl1.Redraw();

}

CString fCarRental::GetDescription(LONG lCarTypeID)
{
	CRecordset oRecordset(&mp_oConn);
	CString sReturn;
	oRecordset.Open(CRecordset::forwardOnly, _T("SELECT sDescription FROM tb_CR_Car_Types WHERE lCarTypeID = ") + CStr(lCarTypeID), CRecordset::readOnly);
	if (oRecordset.IsEOF() == FALSE)
	{
		sReturn = GetStrField(oRecordset, _T("sDescription"));
	}
	oRecordset.Close();
	return sReturn;
}

void fCarRental::CalculateRate(CclsTask* oTask)
{
	DOUBLE cFactor;
    CString* sRowTag = NULL;
    DOUBLE cRate;
    DOUBLE cSubTotal;
    DOUBLE cOptions;
    DOUBLE cSurcharge;
    DOUBLE cTax;
    DOUBLE cALI;
    DOUBLE cCRF;
	DOUBLE cERF;
    DOUBLE cGPS;
    DOUBLE cLDW;
    DOUBLE cPAI;
    DOUBLE cPEP;
    DOUBLE cRCFC;
    DOUBLE cVLF;
    DOUBLE cWTB;
    BOOL bGPS;
    BOOL bLDW;
    BOOL bPAI;
    BOOL bPEP;
    BOOL bALI;
    CString sCustomerName;
    CString sPhone;
    //
    CString sEstimatedTotal;
    DOUBLE cEstimatedTotal;

	CRecordset oRecordset(&mp_oConn);
	oRecordset.Open(CRecordset::forwardOnly, _T("SELECT * FROM tb_CR_Rentals WHERE lTaskID = ") + Replace(oTask->GetKey(), _T("K"), _T("")), CRecordset::readOnly);
	if (oRecordset.IsEOF() == FALSE)
	{
		sCustomerName = GetStrField(oRecordset, _T("sCustomerName"));
		sPhone = GetStrField(oRecordset, _T("sPhone"));

		bGPS = GetBoolField(oRecordset, _T("bGPS"));
		cGPS = GetDblField(oRecordset, _T("cGPS"));
		bLDW = GetBoolField(oRecordset, _T("bLDW"));
		cLDW = GetDblField(oRecordset, _T("cLDW"));
		bPAI = GetBoolField(oRecordset, _T("bPAI"));
		cPAI = GetDblField(oRecordset, _T("cPAI"));
		bPEP = GetBoolField(oRecordset, _T("bPEP"));
		cPEP = GetDblField(oRecordset, _T("cPEP"));
		bALI = GetBoolField(oRecordset, _T("bALI"));
		cALI = GetDblField(oRecordset, _T("cALI"));
		cERF = GetDblField(oRecordset, _T("cERF"));
		cWTB = GetDblField(oRecordset, _T("cWTB"));
		cRCFC = GetDblField(oRecordset, _T("cRCFC"));
		cVLF = GetDblField(oRecordset, _T("cVLF"));
		cCRF = GetDblField(oRecordset, _T("cCRF"));
	}
	oRecordset.Close();


    cFactor = ActiveGanttVCCtl1.GetMathLib().DateTimeDiff(IL_HOUR, oTask->GetStartDate(), oTask->GetEndDate()) / 24;

	if (bGPS == TRUE)
	{
		cGPS = cGPS * cFactor;
	}
	else
	{
		cGPS = 0;
	}
	if (bLDW == TRUE)
	{
		cLDW = cLDW * cFactor;
	}
	else
	{
		cLDW = 0;
	}
	if (bPAI == TRUE)
	{
		cPAI = cPAI * cFactor;
	}
	else
	{
		cPAI = 0;
	}
	if (bPEP == TRUE)
	{
		cPEP = cPEP * cFactor;
	}
	else
	{
		cPEP = 0;
	}
	if (bALI == TRUE)
	{
		cALI = cALI * cFactor;
	}
	else
	{
		cALI = 0;
	}

    sRowTag = Split(oTask->GetRow().GetTag(), _T("|"));
    cRate = CDbl(sRowTag[1]);
	delete [] sRowTag;
    cERF = cERF * cFactor;
    cWTB = cWTB * cFactor;
    cRCFC = cRCFC * cFactor;
    cVLF = cVLF * cFactor;
    cCRF = cCRF * cRate * cFactor;
    cSurcharge = cERF + cWTB + cRCFC + cVLF + cCRF;
    cOptions = cGPS + cLDW + cPAI + cPEP + cALI;
    cSubTotal = cSurcharge + (cRate * cFactor);
    CString sState;
    cTax = cSubTotal * GetStateTax(oTask, &sState);
    cEstimatedTotal = cSubTotal + cTax + cOptions;
    sEstimatedTotal = CStr(cEstimatedTotal);
	if ((oTask->GetTag() = _T("AM_RESERVATION")) || (oTask->GetTag() = _T("AM_RENTAL")))
	{
		oTask->SetText(sCustomerName + _T("\r\n") + _T("Phone: ") + sPhone + _T("\r\n") + _T("Estimated Total: ") + sEstimatedTotal + _T(" USD"));
	}
	else
	{
		cEstimatedTotal = 0;
		cRate = 0;
	}

	clsDB oDB(mp_oConn);
	oDB.AddParameter(_T("dtPickUp"), oTask->GetStartDate(), PT_DATE);
	oDB.AddParameter(_T("dtReturn"), oTask->GetStartDate(), PT_DATE);
	oDB.AddParameter(_T("cRate"), cRate, PT_NUMERIC);
	oDB.AddParameter(_T("cEstimatedTotal"), cEstimatedTotal, PT_NUMERIC);
	oDB.ExecuteUpdate(_T("tb_CR_Rentals"), _T("lTaskID = ") + Replace(oTask->GetKey(), _T("K"), _T("")));
}

DOUBLE fCarRental::GetStateTax(CclsTask* oTask, CString* sState)
{
	CclsNode oNode;
	
    DOUBLE cTax = 0;
    oNode = oTask->GetRow().GetNode().Parent();

	if (oNode == NULL)
	{
		return 0.1;
	}
	else
	{
		CclsRow oRow(oNode.GetRow());

		
		CRecordset oRecordset(&mp_oConn);
		oRecordset.Open(CRecordset::forwardOnly, _T("SELECT sStateAbr FROM tb_CR_Rows WHERE lRowID = ") + Replace(oRow.GetKey(), _T("K"), _T("")), CRecordset::readOnly);
		if (oRecordset.IsEOF() == FALSE)
		{
			*sState = GetStrField(oRecordset, _T("sStateAbr"));
		}
		oRecordset.Close();
		oRecordset.Open(CRecordset::forwardOnly, _T("SELECT cCarRentalTax FROM tb_CR_US_States WHERE sStateAbr = '") + *sState + _T("'"), CRecordset::readOnly);
		if (oRecordset.IsEOF() == FALSE)
		{
			cTax = GetDblField(oRecordset, _T("cCarRentalTax"));
		}
		oRecordset.Close();
	}
	return cTax;
}

void fCarRental::cmdSaveXML()
{
	SaveXML();
}

void fCarRental::cmdLoadXML()
{
	LoadXML();
}

void fCarRental::cmdPrint_Click()
{
	Print();
}

void fCarRental::cmdZoomIn_Click()
{
	SetZoom(GetZoom() - 1);
}

void fCarRental::cmdZoomOut_Click()
{
	SetZoom(GetZoom() + 1);
}

void fCarRental::cmdAddVehicle_Click()
{
	fCarRentalVehicle oForm;
	oForm.mp_yDialogMode = DM_ADD;
	oForm.mp_oParent = this;
	oForm.DoModal();
}

void fCarRental::cmdAddBranch_Click()
{
	fCarRentalBranch oForm;
	oForm.mp_yDialogMode = DM_ADD;
	oForm.mp_oParent = this;
	oForm.DoModal();
}

void fCarRental::cmdBack2_Click()
{
	ActiveGanttVCCtl1.GetCurrentViewObject().GetTimeLine().Position(ActiveGanttVCCtl1.GetMathLib().DateTimeAdd(ActiveGanttVCCtl1.GetCurrentViewObject().GetClientArea().GetGrid().GetInterval(), -10 * ActiveGanttVCCtl1.GetCurrentViewObject().GetClientArea().GetGrid().GetFactor(), ActiveGanttVCCtl1.GetCurrentViewObject().GetTimeLine().GetStartDate()));
	ActiveGanttVCCtl1.Redraw();
}

void fCarRental::cmdBack1_Click()
{
	ActiveGanttVCCtl1.GetCurrentViewObject().GetTimeLine().Position(ActiveGanttVCCtl1.GetMathLib().DateTimeAdd(ActiveGanttVCCtl1.GetCurrentViewObject().GetClientArea().GetGrid().GetInterval(), -5 * ActiveGanttVCCtl1.GetCurrentViewObject().GetClientArea().GetGrid().GetFactor(), ActiveGanttVCCtl1.GetCurrentViewObject().GetTimeLine().GetStartDate()));
	ActiveGanttVCCtl1.Redraw();
}

void fCarRental::cmdBack0_Click()
{
	ActiveGanttVCCtl1.GetCurrentViewObject().GetTimeLine().Position(ActiveGanttVCCtl1.GetMathLib().DateTimeAdd(ActiveGanttVCCtl1.GetCurrentViewObject().GetClientArea().GetGrid().GetInterval(), -1 * ActiveGanttVCCtl1.GetCurrentViewObject().GetClientArea().GetGrid().GetFactor(), ActiveGanttVCCtl1.GetCurrentViewObject().GetTimeLine().GetStartDate()));
	ActiveGanttVCCtl1.Redraw();
}

void fCarRental::cmdFwd0_Click()
{
	ActiveGanttVCCtl1.GetCurrentViewObject().GetTimeLine().Position(ActiveGanttVCCtl1.GetMathLib().DateTimeAdd(ActiveGanttVCCtl1.GetCurrentViewObject().GetClientArea().GetGrid().GetInterval(), 1 * ActiveGanttVCCtl1.GetCurrentViewObject().GetClientArea().GetGrid().GetFactor(), ActiveGanttVCCtl1.GetCurrentViewObject().GetTimeLine().GetStartDate()));
	ActiveGanttVCCtl1.Redraw();
}

void fCarRental::cmdFwd1_Click()
{
	ActiveGanttVCCtl1.GetCurrentViewObject().GetTimeLine().Position(ActiveGanttVCCtl1.GetMathLib().DateTimeAdd(ActiveGanttVCCtl1.GetCurrentViewObject().GetClientArea().GetGrid().GetInterval(), 5 * ActiveGanttVCCtl1.GetCurrentViewObject().GetClientArea().GetGrid().GetFactor(), ActiveGanttVCCtl1.GetCurrentViewObject().GetTimeLine().GetStartDate()));
	ActiveGanttVCCtl1.Redraw();
}

void fCarRental::cmdFwd2_Click()
{
	ActiveGanttVCCtl1.GetCurrentViewObject().GetTimeLine().Position(ActiveGanttVCCtl1.GetMathLib().DateTimeAdd(ActiveGanttVCCtl1.GetCurrentViewObject().GetClientArea().GetGrid().GetInterval(), 10 * ActiveGanttVCCtl1.GetCurrentViewObject().GetClientArea().GetGrid().GetFactor(), ActiveGanttVCCtl1.GetCurrentViewObject().GetTimeLine().GetStartDate()));
	ActiveGanttVCCtl1.Redraw();
}

void fCarRental::cmdHelp_Click()
{
	g_ShowInBrowser(_T("http://www.sourcecodestore.com/Article.aspx?ID=16"));
}

void fCarRental::mnuEditRow_Click()
{
	CclsRow oRow;
	oRow = ActiveGanttVCCtl1.GetRows().Item(mp_sEditRowKey);

	if (oRow.GetNode().GetDepth() == 1)
	{
		fCarRentalVehicle oForm;
		oForm.mp_yDialogMode = DM_EDIT;
		oForm.mp_sRowID = Replace(mp_sEditRowKey, _T("K"), _T(""));
		oForm.mp_oParent = this;
		oForm.DoModal();
	}
	else if (oRow.GetNode().GetDepth() == 0)
	{
		fCarRentalBranch oForm;
		oForm.mp_yDialogMode = DM_EDIT;
		oForm.mp_sRowID = Replace(mp_sEditRowKey, _T("K"), _T(""));
		oForm.mp_oParent = this;
		oForm.DoModal();
	}
}

void fCarRental::mnuDeleteRow_Click()
{
	if (MessageBox(_T("Are you sure you want to delete this item?"), _T("Delete Row"), MB_YESNO) == IDYES)
	{
		mp_oConn.ExecuteSQL(_T("DELETE * FROM tb_CR_Rows WHERE lRowID = ") + Replace(mp_sEditRowKey, _T("K"), _T("")));
		mp_oConn.ExecuteSQL(_T("DELETE * FROM tb_CR_Rentals WHERE lRowID = ") + Replace(mp_sEditRowKey, _T("K"), _T("")));
		ActiveGanttVCCtl1.GetRows().Remove(mp_sEditRowKey);
		ActiveGanttVCCtl1.GetRows().UpdateTree();
        ActiveGanttVCCtl1.Redraw();
	}
}

void fCarRental::mnuEditTask_Click()
{
	fCarRentalReservation oForm(DM_EDIT, this, Replace(mp_sEditTaskKey, _T("K"), _T("")));
	ReleaseCapture();
	oForm.DoModal();
}

void fCarRental::mnuDeleteTask_Click()
{
	if (MessageBox(_T("Are you sure you want to delete this item?"), _T("Delete Task"), MB_YESNO) == IDYES)
	{
		mp_oConn.ExecuteSQL(_T("DELETE * FROM tb_CR_Rentals WHERE lTaskID = ") + Replace(mp_sEditTaskKey, _T("K"), _T("")));
		ActiveGanttVCCtl1.GetTasks().Remove(mp_sEditTaskKey);
        ActiveGanttVCCtl1.Redraw();

	}
}

void fCarRental::mnuConvertToRental_Click()
{
	CclsTask oTask;
	oTask = ActiveGanttVCCtl1.GetTasks().Item(mp_sEditTaskKey);
	mp_oConn.ExecuteSQL(_T("UPDATE tb_CR_Rentals SET [lMode] = 1 WHERE lTaskID = ") + Replace(mp_sEditTaskKey, _T("K"), _T("")));
	oTask.SetTag(_T("AM_RENTAL"));
    oTask.SetStyleIndex(_T("RET1"));
    ActiveGanttVCCtl1.Redraw();
}

void fCarRental::mnuSaveXML_Click()
{
	SaveXML();
}

void fCarRental::mnuLoadXML_Click()
{
	LoadXML();
}

void fCarRental::mnuPrint_Click()
{
	Print();
}

void fCarRental::mnuClose_Click()
{
	CDialog::EndDialog(IDOK);
}

void fCarRental::SaveXML()
{
	TCHAR sDefaultFileName[MAX_PATH];
	wcscpy_s(sDefaultFileName, MAX_PATH, _T("AGVC_CR.xml"));
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

void fCarRental::LoadXML()
{
	fLoadXML oForm;
	oForm.DoModal();
}

void fCarRental::Print()
{
	if (ActiveGanttVCCtl1.GetPrinterErrorMessage().GetLength() == 0)
	{
		fPrintDialog oForm(GetDateTime(2009, 6, 1, 0, 0, 0), GetDateTime(2009, 6, 30, 0, 0, 0));
		oForm.mp_oControl = &ActiveGanttVCCtl1;
		oForm.DoModal();
	}
}

void fCarRental::Access_LoadRowsAndTasks()
{
	CRecordset oRecordset(&mp_oConn);
	CString sRowID;
	CclsTask oTask;
	CclsRow oRow;
	
	
	oRecordset.Open(CRecordset::forwardOnly, _T("SELECT * FROM tb_CR_Rows ORDER BY lOrder"), CRecordset::readOnly);
	while(oRecordset.IsEOF() == FALSE)
	{
		
		int lRowID, lDepth, lCarTypeID;
		CString sBranchName, sState, sPhone, sLicensePlates, sACRISSCode;
		double cRate;
		lRowID = GetIntField(oRecordset, _T("lRowID"));
		lDepth = GetIntField(oRecordset, _T("lDepth"));
		sBranchName = GetStrField(oRecordset, _T("sBranchName"));
		sState = GetStrField(oRecordset, _T("sStateAbr"));
		sPhone = GetStrField(oRecordset, _T("sPhone"));
		lCarTypeID = GetIntField(oRecordset, _T("lCarTypeID"));
		sLicensePlates = GetStrField(oRecordset, _T("sLicensePlates"));
		sACRISSCode = GetStrField(oRecordset, _T("sACRISSCode"));
		cRate = GetDblField(oRecordset, _T("cRate"));

		sRowID = _T("K") + CStr(lRowID);
		oRow = ActiveGanttVCCtl1.GetRows().Add(sRowID, _T(""), FALSE, TRUE, _T(""));
		if (lDepth == 0)
		{
			oRow.SetText(sBranchName + _T(", ") + sState + _T("\r\n") + _T("Phone: ") + sPhone);
            oRow.SetMergeCells(TRUE);
            oRow.SetContainer(FALSE);
            oRow.SetStyleIndex(_T("Branch"));
            oRow.SetClientAreaStyleIndex(_T("BranchCA"));
            oRow.GetNode().SetDepth(0);
            oRow.SetUseNodeImages(TRUE);
            oRow.GetNode().GetExpandedImage().FromFile(g_GetAppLocation() + _T("\\CarRental\\minus.jpg"), TRUE);
            oRow.GetNode().GetImage().FromFile(g_GetAppLocation() + _T("\\CarRental\\plus.jpg"), TRUE);
            oRow.SetAllowMove(FALSE);
            oRow.SetAllowSize(FALSE);
		}
		else if (lDepth == 1)
		{
            CString sDescription;
            sDescription = GetDescription(lCarTypeID);
            oRow.GetCells().Item(_T("1")).SetText(sLicensePlates);
            oRow.GetCells().Item(_T("2")).GetImage().FromFile(g_GetAppLocation() + _T("\\CarRental\\Small\\") + sDescription + _T(".jpg"), TRUE);
            oRow.GetCells().Item(_T("3")).SetText(sDescription + _T("\r\n") + sACRISSCode + _T(" - ") + CStr(cRate) + _T(" USD"));
            oRow.GetNode().SetDepth(1);
            oRow.SetTag(sACRISSCode + _T("|") + CStr(cRate) + _T("|") + CStr(lCarTypeID));
		}
		oRecordset.MoveNext();
	}
	oRecordset.Close();
	oRecordset.Open(CRecordset::forwardOnly, _T("SELECT * FROM tb_CR_Rentals"), CRecordset::readOnly);
	while(oRecordset.IsEOF() == FALSE)
	{
		int lRowID, lTaskID;
		COleDateTime dtPickUp, dtReturn;
		HPE_ADDMODE lMode;
		CString sName, sPhone;
		double cEstimatedTotal;
		

		lRowID = GetIntField(oRecordset, _T("lRowID"));
		dtPickUp = GetDateField(oRecordset, _T("dtPickUp"));
		dtReturn = GetDateField(oRecordset, _T("dtReturn"));
		lTaskID = GetIntField(oRecordset, _T("lTaskID"));
		lMode = (HPE_ADDMODE)GetIntField(oRecordset, _T("lMode"));
		sName = GetStrField(oRecordset, _T("sCustomerName"));
		sPhone = GetStrField(oRecordset, _T("sPhone"));
		cEstimatedTotal = GetDblField(oRecordset, _T("cEstimatedTotal"));

		if (lMode == AM_RESERVATION)
		{
			oTask.SetTag(_T("AM_RESERVATION"));
		}
		else if (lMode == AM_RENTAL)
		{
			oTask.SetTag(_T("AM_RENTAL"));
		}
		else if (lMode == AM_MAINTENANCE)
		{
			oTask.SetTag(_T("AM_MAINTENANCE"));
		}
		oTask = ActiveGanttVCCtl1.GetTasks().Add(_T(""), _T("K") + CStr(lRowID), dtPickUp, dtReturn, _T("K") + CStr(lTaskID), _T(""), _T(""));
		if (lMode == AM_MAINTENANCE)
		{
			oTask.SetText(_T("Scheduled Maintenance"));
			oTask.SetStyleIndex(_T("RET3"));
		}
		else
		{
			oTask.SetText(sName + _T("\r\n") + _T("Phone: ") + sPhone + _T("\r\n") + _T("Estimated Total: ") + CStr(cEstimatedTotal) + _T(" USD"));
			if (lMode == AM_RESERVATION)
			{
				oTask.SetStyleIndex(_T("RET2"));
			}
			else if (lMode == AM_RENTAL)
			{
				oTask.SetStyleIndex(_T("RET1"));
			}
		}
		
		oRecordset.MoveNext();
	}
	oRecordset.Close();
}