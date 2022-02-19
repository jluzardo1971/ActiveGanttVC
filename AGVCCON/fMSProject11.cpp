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
#include "fMSProject11.h"
#include "fPrintDialog.h"
using namespace MSP2003;

IMPLEMENT_DYNAMIC(fMSProject11, CDialog)

fMSProject11::fMSProject11(CWnd* pParent /*=NULL*/)
	: CDialog(fMSProject11::IDD, pParent)
{
	m_hIcon = AfxGetApp()->LoadIcon(IDR_MAINFRAME);
	mp_sFontName = _T("Tahoma");
}

fMSProject11::~fMSProject11()
{
	
}

void fMSProject11::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	DDX_Control(pDX, IDC_ACTIVEGANTTVCCTL1, ActiveGanttVCCtl1);
	DDX_Control(pDX, IDC_MSP2003CTRL1, MSP2003Ctl1);
}


BEGIN_MESSAGE_MAP(fMSProject11, CDialog)
	ON_WM_CREATE()
	ON_WM_SIZE()
	ON_COMMAND_RANGE(ID_MYTOOLBAR_BUTTONFIRST, ID_MYTOOLBAR_BUTTONLAST, ToolBar_Click)
	ON_COMMAND(ID_FILE_LOADXML, mnuLoadXML_Click)
	ON_COMMAND(ID_FILE_SAVEXML, mnuSaveXML_Click)
	ON_COMMAND(ID_FILE_PRINT, mnuPrint_Click)
	ON_COMMAND(ID_FILE_CLOSE, mnuClose_Click)

	ON_WM_CLOSE()
END_MESSAGE_MAP()

void fMSProject11::ToolBar_Click(UINT nID)
{
	UINT lButtonFirst = ID_MYTOOLBAR_BUTTONFIRST;
	nID = nID - lButtonFirst;
	switch (nID)
	{
	case 0: //Load XML
		{
			cmdLoadXML();
		}
		break;
	case 1: //Save as XML
		{
			cmdSaveXML();
		}
		break;
	case 2: //Print
		{
			cmdPrint_Click();
		}
		break;
	case 4: //Indent
		{
			cmdIndent_Click();
		}
		break;
	case 6: //Zoom in
		{
			cmdZoomIn_Click();
		}
		break;
	case 7: //Zoom out
		{
			cmdZoomOut_Click();
		}
		break;
	}

}

int fMSProject11::OnCreate(LPCREATESTRUCT lpCreateStruct)
{
	if (CDialog::OnCreate(lpCreateStruct) == -1)
		return -1;

	// TODO:  Add your specialized creation code here
	oToolBar1.CreateEx(this, TBSTYLE_FLAT, WS_CHILD | WS_VISIBLE | CBRS_TOP | CBRS_GRIPPER | CBRS_TOOLTIPS | CBRS_FLYBY | CBRS_SIZE_DYNAMIC);
	oToolBarBitmap.LoadBitmapW(IDB_FMSPROJECT);
	oToolBar1.SetBitmap((HBITMAP)oToolBarBitmap.GetSafeHandle());
	oToolBar1.SetButtons(NULL, 8); //1 Based
	oToolBar1.AddButton(0, _T("Load XML"));
	oToolBar1.AddButton(1, _T("Save as XML"));
	oToolBar1.AddButton(2, _T("Print"));
	oToolBar1.AddSeparator();
	oToolBar1.AddButton(3, _T("Indent"));
	oToolBar1.AddSeparator();
	oToolBar1.AddButton(4, _T("Zoom in"));
	oToolBar1.AddButton(5, _T("Zoom out"));


	return 0;
}

BOOL fMSProject11::OnInitDialog()
{
	CDialog::OnInitDialog();
	g_MaximizeWindowsDim(this);
	CWnd::ShowWindow(SW_SHOWMAXIMIZED);

	SetIcon(m_hIcon, TRUE);			// Set big icon
	SetIcon(m_hIcon, FALSE);		// Set small icon

	CWnd::RepositionBars(AFX_IDW_CONTROLBAR_FIRST, AFX_IDW_CONTROLBAR_LAST,0);

	InitializeAG();
	ActiveGanttVCCtl1.Redraw();

	return TRUE;
}

void fMSProject11::OnSize(UINT nType, int cx, int cy)
{
	CDialog::OnSize(nType, cx, cy);
	g_Resize(this, &ActiveGanttVCCtl1);
}

void fMSProject11::InitializeAG()
{
	CclsStyle oStyle;
	CclsView oView;

    CControlTemplateSolid oTemplate;
	oTemplate.CreateDispatch(_T("AGVC.ControlTemplateSolid"));
    oTemplate.SetControlBorderColor(FromArgb(255, 120, 120, 120));
    oTemplate.SetHeadingBorderColor(CLR_WHITE);
    oTemplate.SetHeadingForeColor(CLR_WHITE);
    oTemplate.SetRowForeColor(CLR_BLACK);
    oTemplate.SetHeadingBackColor(FromArgb(255, 111, 120, 156));
    oTemplate.SetTreeviewColor(CLR_BLACK);
    oTemplate.SetRowBorderColor(FromArgb(255, 192, 192, 192));
    oTemplate.SetTimeBlockBackColor(FromArgb(255, 115, 119, 135));

    ActiveGanttVCCtl1.ApplyTemplateSolid(oTemplate, STO_COLOR_HATCH);

	oStyle = ActiveGanttVCCtl1.GetStyles().Item(_T("T1"));
	oStyle.SetTextPlacement(SCP_EXTERIORPLACEMENT);
	oStyle.SetTextAlignmentHorizontal(HAL_RIGHT);
	oStyle.SetTextXMargin(30);

	ActiveGanttVCCtl1.SetAllowRowMove(TRUE);
	ActiveGanttVCCtl1.SetAllowRowSize(TRUE);
	ActiveGanttVCCtl1.SetAddMode(AT_BOTH);
	ActiveGanttVCCtl1.GetSplitter().SetPosition(285);
	ActiveGanttVCCtl1.GetTreeview().SetImages(TRUE);
	ActiveGanttVCCtl1.GetTreeview().SetCheckBoxes(TRUE);
	ActiveGanttVCCtl1.GetTreeview().SetTreeLines(FALSE);
	ActiveGanttVCCtl1.GetToolTip().GetFont().SetFontName(mp_sFontName);
	ActiveGanttVCCtl1.GetToolTip().GetFont().SetSize(8);
	ActiveGanttVCCtl1.GetToolTip().GetFont().SetStyle(FS_FONTSTYLEREGULAR);

	CclsColumn oColumn;

	oColumn = ActiveGanttVCCtl1.GetColumns().Add(_T("ID"), _T(""), 30, _T(""));
	oColumn.SetAllowTextEdit(TRUE);

	oColumn = ActiveGanttVCCtl1.GetColumns().Add(_T("Task Name"), _T(""), 255, _T(""));
	oColumn.SetAllowTextEdit(TRUE);

	ActiveGanttVCCtl1.SetTreeviewColumnIndex(2);

	TRACE("Clone 1 \r\n");
	oView = ActiveGanttVCCtl1.GetCurrentViewObject().Clone(_T(""));
	oView.SetInterval(IL_HOUR);
	oView.SetFactor(12);
	oView.GetTimeLine().GetTierArea().GetUpperTier().SetTierType(ST_CUSTOM);
	oView.GetTimeLine().GetTierArea().GetUpperTier().SetInterval(IL_QUARTER);
	oView.GetTimeLine().GetTierArea().GetUpperTier().SetFactor(1);
	oView.GetTimeLine().GetTierArea().GetUpperTier().SetHeight(17);
	oView.GetTimeLine().GetTierArea().GetMiddleTier().SetTierType(ST_NOT_VISIBLE);
	oView.GetTimeLine().GetTierArea().GetLowerTier().SetTierType(ST_CUSTOM);
	oView.GetTimeLine().GetTierArea().GetLowerTier().SetInterval(IL_MONTH);
	oView.GetTimeLine().GetTierArea().GetLowerTier().SetFactor(1);
	oView.GetTimeLine().GetTierArea().GetLowerTier().SetHeight(17);
	oView.GetTimeLine().GetTickMarkArea().SetVisible(FALSE);
	oView.GetTimeLine().GetTimeLineScrollBar().SetStartDate(GetDateTimeNow());
	oView.GetTimeLine().GetTimeLineScrollBar().SetInterval(IL_HOUR);
	oView.GetTimeLine().GetTimeLineScrollBar().SetFactor(1);
	oView.GetTimeLine().GetTimeLineScrollBar().SetSmallChange(48);
	oView.GetTimeLine().GetTimeLineScrollBar().SetLargeChange(2880);
	oView.GetTimeLine().GetTimeLineScrollBar().SetMax(24000);
	oView.GetTimeLine().GetTimeLineScrollBar().SetValue(0);
	oView.GetTimeLine().GetTimeLineScrollBar().SetEnabled(TRUE);
	oView.GetTimeLine().GetTimeLineScrollBar().SetVisible(TRUE);
	oView.GetClientArea().SetDetectConflicts(FALSE);
	oView.GetClientArea().GetGrid().SetColor(FromArgb(255, 197, 206, 216));

	TRACE("Clone 2 \r\n");
	oView = oView.Clone(_T(""));
    oView.SetInterval(IL_HOUR);
    oView.SetFactor(12);

	TRACE("Clone 3 \r\n");
	oView = oView.Clone(_T(""));
    oView.SetInterval(IL_HOUR);
    oView.SetFactor(6);

	TRACE("Clone 4 \r\n");
	oView = oView.Clone(_T(""));
    oView.SetInterval(IL_HOUR);
    oView.SetFactor(3);

	TRACE("Clone 5 \r\n");
	oView = oView.Clone(_T(""));
    oView.SetInterval(IL_HOUR);
    oView.SetFactor(1);

	if (ActiveGanttVCCtl1.GetPrinterErrorMessage().GetLength() == 0)
	{
		ActiveGanttVCCtl1.GetPrinter().SetOrientation(OT_LANDSCAPE);
		ActiveGanttVCCtl1.GetPrinter().SetMarginLeft(50); //0.5"
		ActiveGanttVCCtl1.GetPrinter().SetMarginTop(50); //0.5"
		ActiveGanttVCCtl1.GetPrinter().SetMarginRight(50); //0.5"
		ActiveGanttVCCtl1.GetPrinter().SetMarginBottom(50); //0.5"
		ActiveGanttVCCtl1.GetPrinter().SetHeadingsInEveryPage(TRUE);
		ActiveGanttVCCtl1.GetPrinter().SetKeepRowsTogether(TRUE);
		ActiveGanttVCCtl1.GetPrinter().SetColumnsInEveryPage(TRUE);
		ActiveGanttVCCtl1.GetPrinter().SetColumns(1);
		ActiveGanttVCCtl1.GetPrinter().SetKeepColumnsTogether(TRUE);
		ActiveGanttVCCtl1.GetPrinter().SetKeepTimeIntervalsTogether(TRUE);
		ActiveGanttVCCtl1.GetPrinter().SetInterval(IL_DAY);
		ActiveGanttVCCtl1.GetPrinter().SetFactor(1);
	}

	ActiveGanttVCCtl1.SetTimeLineScrollBarScope(OS_VIEW);

	ActiveGanttVCCtl1.SetCurrentView(_T("1"));
}



void fMSProject11::MP11_To_AG()
{
	MSP2003::CMP11 oMP11;
	oMP11 = MSP2003Ctl1.GetMP11();
	CclsTask oAGTask;
	CclsRow oAGRow;
	MSP2003::CTask oMPTask;
	int i;
	int j;
	
	for (i = 1; i <= oMP11.GetoTasks().GetCount(); i++)
	{
		oMPTask = oMP11.GetoTasks().Item(CStr(i));
		oAGRow = ActiveGanttVCCtl1.GetRows().Add(_T("K") + CStr(oMPTask.GetlUID()), _T(""), FALSE, TRUE, _T(""));	
		oAGRow.GetCells().Item(_T("1")).SetText(CStr(oMPTask.GetlUID()));
		oAGRow.GetCells().Item(_T("1")).SetStyleIndex(_T("CH"));
		oAGRow.SetHeight(20);
		oAGTask = ActiveGanttVCCtl1.GetTasks().Add(_T(""), _T("K") + CStr(oMPTask.GetlUID()), oMPTask.GetdtStart(), oMPTask.GetdtFinish(), _T(""), _T(""), _T(""));
        oAGTask.SetKey(_T("K") + CStr(oMPTask.GetlUID()));
        oAGTask.SetAllowedMovement(MT_RESTRICTEDTOROW);
		oAGTask.SetAllowTextEdit(TRUE);
		if (oMPTask.GetdtStart() < mp_dtStartDate || mp_dtStartDate.m_dt == (DATE)0)
		{
			mp_dtStartDate = oMPTask.GetdtStart();
		}
		if (oMPTask.GetdtFinish() > mp_dtEndDate || mp_dtEndDate.m_dt == (DATE)0)
		{
			mp_dtEndDate = oMPTask.GetdtFinish();
		}
		if (oAGTask.GetStartDate() == oAGTask.GetEndDate())
		{
			oAGTask.SetText(FormatDate(oAGTask.GetStartDate(), _T("%m/%d"))); //"M/d"
		}
		oAGRow.GetNode().SetDepth(oMPTask.GetlOutlineLevel());
		oAGRow.GetNode().SetText(oMPTask.GetsName());
		oAGRow.GetNode().SetAllowTextEdit(TRUE);
		if (oMPTask.GetsNotes().GetLength() > 0)
		{
			oAGRow.GetNode().GetImage().FromFile(g_GetAppLocation() + _T("\\res\\projectnote.gif"), TRUE);
			oAGRow.GetNode().SetImageVisible(TRUE);
		}
	}

	ActiveGanttVCCtl1.GetRows().UpdateTree();

    // Indent & set Predecessors
	for (i = 1; i <= oMP11.GetoTasks().GetCount(); i++)
	{
		oMPTask = oMP11.GetoTasks().Item(CStr(i));
		oAGRow = ActiveGanttVCCtl1.GetRows().Item(CStr(i));
		oAGTask = ActiveGanttVCCtl1.GetTasks().Item(CStr(i));
		if (oAGRow.GetNode().Children() > 0)
		{
            oAGTask.SetStyleIndex(_T("S1"));
		}
		else
		{
			oAGTask.SetStyleIndex(_T("T1"));
		}
		for (j = 1; j <= oMPTask.GetoPredecessorLink_C().GetCount(); j++)
		{
			MSP2003::CTaskPredecessorLink oMPPredecessor;
			oMPPredecessor = oMPTask.GetoPredecessorLink_C().Item(CStr(j));
			ActiveGanttVCCtl1.GetPredecessors().Add(_T("K") + CStr(oMPTask.GetlUID()), _T("K") + CStr(oMPPredecessor.GetlPredecessorUID()), GetAGPredecessorType((E_TYPE_3)oMPPredecessor.GetyType()), _T(""), _T("T1"));
		}
	}


    // Assignments
	for (i = 1; i <= oMP11.GetoAssignments().GetCount(); i++)
	{
		MSP2003::CAssignment oAssignment;
		oAssignment = oMP11.GetoAssignments().Item(CStr(i));
		oAGTask = ActiveGanttVCCtl1.GetTasks().Item(_T("K") + CStr(oAssignment.GetlTaskUID()));
		if (oAGTask.GetStartDate() != oAGTask.GetEndDate())
		{
			if (oAssignment.GetlResourceUID() > 0)
			{
				if (oAGTask.GetText().GetLength() == 0)
				{
					oAGTask.SetText(oMP11.GetoResources().Item(_T("K") + CStr(oAssignment.GetlResourceUID())).GetsName());
				}
				else
				{
					oAGTask.SetText(oAGTask.GetText() + _T(", ") + oMP11.GetoResources().Item(_T("K") + CStr(oAssignment.GetlResourceUID())).GetsName());
				}
			}
		}
	}
    mp_dtStartDate = ActiveGanttVCCtl1.GetMathLib().DateTimeAdd(IL_HOUR, -3, mp_dtStartDate);
    AGSetStartDate(mp_dtStartDate);
}

ActiveGanttVC::E_CONSTRAINTTYPE fMSProject11::GetAGPredecessorType(MSP2003::E_TYPE_3 MPPredecessorType)
{
	switch (MPPredecessorType)
	{
	case T_3_FF:
		{
			return PCT_END_TO_END;
		}
		break;
	case T_3_FS:
		{
			return PCT_END_TO_START;
		}
		break;
	case T_3_SF:
		{
			return PCT_START_TO_END;
		}
		break;
	case T_3_SS:
		{
			return PCT_START_TO_START;
		}
		break;
	}
	return (ActiveGanttVC::E_CONSTRAINTTYPE)0;
}

void fMSProject11::AGSetStartDate(COleDateTime dtDate)
{
	int i;
	for (i = 1; i <= ActiveGanttVCCtl1.GetViews().GetCount(); i++)
	{
		ActiveGanttVCCtl1.GetViews().Item(CStr(i)).GetTimeLine().GetTimeLineScrollBar().SetStartDate(dtDate);
	}
}

void fMSProject11::cmdSaveXML()
{
	CFileDialog oDialog(FALSE, NULL, NULL, OFN_OVERWRITEPROMPT,_T("XML File|*.xml||"), this, NULL);
	oDialog.m_ofn.lpstrTitle = _T("Save As MS-Project 2003 XML File");
	if (oDialog.DoModal() == IDOK)
	{
		//AG_To_MP11();
		CString sFilePath = oDialog.GetPathName();
		MSP2003Ctl1.GetMP11().WriteXML(sFilePath);
	}
}

void fMSProject11::cmdLoadXML()
{
	TCHAR sDefaultFileName[MAX_PATH * 5];
	TCHAR sInitialDir[MAX_PATH * 5];
	wcscpy_s(sDefaultFileName, MAX_PATH * 5, _T(""));
	wcscpy_s(sInitialDir, MAX_PATH * 5, g_GetAppLocation() + _T("\\MSP2003"));
	CFileDialog oDialog(TRUE, NULL, NULL, OFN_HIDEREADONLY | OFN_OVERWRITEPROMPT,_T("XML File (*.xml)|*.xml|All Files (*.*)|*.*||"), this, NULL);
	oDialog.m_ofn.lpstrTitle = _T("Load MS-Project 2003 XML File");
	oDialog.m_ofn.lpstrFile = sDefaultFileName;
	oDialog.m_ofn.lpstrInitialDir = sInitialDir;
	oDialog.m_ofn.nMaxFile = MAX_PATH * 5;
	if (oDialog.DoModal() == IDOK)
	{
		CWaitCursor oWaitCursor;
		CString sFilePath = oDialog.GetPathName();
        ActiveGanttVCCtl1.Clear();
		MSP2003Ctl1.Clear();
        MSP2003Ctl1.GetMP11().ReadXML(sFilePath);
		oWaitCursor.Restore();
        InitializeAG();
        MP11_To_AG();
        ActiveGanttVCCtl1.Redraw();
	}
}

void fMSProject11::cmdPrint_Click()
{
	if (ActiveGanttVCCtl1.GetPrinterErrorMessage().GetLength() == 0)
	{
		if (ActiveGanttVCCtl1.GetRows().GetCount() == 0)
		{
			return;
		}
		if (mp_dtStartDate.m_dt == (DATE)0 || mp_dtEndDate.m_dt == (DATE)0)
		{
			return;
		}
		ActiveGanttVCCtl1.GetPrinter().SetInterval(ActiveGanttVCCtl1.GetCurrentViewObject().GetTimeLine().GetTierArea().GetLowerTier().GetInterval());
		ActiveGanttVCCtl1.GetPrinter().SetFactor(ActiveGanttVCCtl1.GetCurrentViewObject().GetTimeLine().GetTierArea().GetLowerTier().GetFactor());
		fPrintDialog oForm(mp_dtStartDate, mp_dtEndDate);
		oForm.mp_oControl = &ActiveGanttVCCtl1;
		oForm.DoModal();
	}
}

void fMSProject11::cmdIndent_Click()
{
}

void fMSProject11::cmdZoomIn_Click()
{
	if (CInt32(ActiveGanttVCCtl1.GetCurrentView()) < ActiveGanttVCCtl1.GetViews().GetCount())
	{
		ActiveGanttVCCtl1.SetCurrentView(CStr(CInt32(ActiveGanttVCCtl1.GetCurrentView()) + 1));
        ActiveGanttVCCtl1.Redraw();
	}
}

void fMSProject11::cmdZoomOut_Click()
{
	if (CInt32(ActiveGanttVCCtl1.GetCurrentView()) > 1)
	{
        ActiveGanttVCCtl1.SetCurrentView(CStr(CInt32(ActiveGanttVCCtl1.GetCurrentView()) - 1));
        ActiveGanttVCCtl1.Redraw();
	}
}

void fMSProject11::mnuLoadXML_Click()
{
	cmdLoadXML();
}

void fMSProject11::mnuSaveXML_Click()
{
	cmdSaveXML();
}

void fMSProject11::mnuPrint_Click()
{
	cmdPrint_Click();
}

void fMSProject11::mnuClose_Click()
{
	CDialog::EndDialog(IDOK);
}

BEGIN_EVENTSINK_MAP(fMSProject11, CDialog)
	ON_EVENT(fMSProject11, IDC_ACTIVEGANTTVCCTL1, 13, fMSProject11::CustomTierDrawActiveganttvcctl1, VTS_DISPATCH)
	ON_EVENT(fMSProject11, IDC_ACTIVEGANTTVCCTL1, 35, fMSProject11::ControlMouseWheelActiveganttvcctl1, VTS_DISPATCH)
END_EVENTSINK_MAP()

void fMSProject11::CustomTierDrawActiveganttvcctl1(LPDISPATCH e)
{
	CCustomTierDrawEventArgs oE(e);
	if (oE.GetTierPosition() == SP_UPPER)
	{
		if (CInt32(ActiveGanttVCCtl1.GetCurrentView()) <= 4)
		{
			oE.SetText(FormatDate(oE.GetStartDate(), _T("%Y")) + _T(" Q") + Quarter(oE.GetStartDate()));
		}
		else
		{
			oE.SetText(FormatDate(oE.GetStartDate(), _T("%b")));
		}
	}
	else if (oE.GetTierPosition() == SP_LOWER)
	{
		if (CInt32(ActiveGanttVCCtl1.GetCurrentView()) <= 4)
		{
			oE.SetText(FormatDate(oE.GetStartDate(), _T("%b")));
		}
		else
		{
			oE.SetText(FormatDate(oE.GetStartDate(), _T("%a")));
		}
	}
}

void fMSProject11::ControlMouseWheelActiveganttvcctl1(LPDISPATCH e)
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

void fMSProject11::OnClose()
{
	CDialog::OnClose();
}