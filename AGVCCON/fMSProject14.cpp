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
#include "fMSProject14.h"
#include "fPrintDialog.h"
using namespace MSP2010;

IMPLEMENT_DYNAMIC(fMSProject14, CDialog)

fMSProject14::fMSProject14(CWnd* pParent /*=NULL*/)
	: CDialog(fMSProject14::IDD, pParent)
{
	m_hIcon = AfxGetApp()->LoadIcon(IDR_MAINFRAME);
	mp_sFontName = _T("Tahoma");
}

fMSProject14::~fMSProject14()
{
}

void fMSProject14::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	DDX_Control(pDX, IDC_ACTIVEGANTTVCCTL1, ActiveGanttVCCtl1);
	DDX_Control(pDX, IDC_MSP2010CTRL1, MSP2010Ctl1);
}


BEGIN_MESSAGE_MAP(fMSProject14, CDialog)
	ON_WM_CREATE()
	ON_WM_SIZE()
	ON_COMMAND_RANGE(ID_MYTOOLBAR_BUTTONFIRST, ID_MYTOOLBAR_BUTTONLAST, ToolBar_Click)
	ON_COMMAND(ID_FILE_LOADXML, mnuLoadXML_Click)
	ON_COMMAND(ID_FILE_SAVEXML, mnuSaveXML_Click)
	ON_COMMAND(ID_FILE_PRINT, mnuPrint_Click)
	ON_COMMAND(ID_FILE_CLOSE, mnuClose_Click)

	ON_WM_CLOSE()
END_MESSAGE_MAP()

void fMSProject14::ToolBar_Click(UINT nID)
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

int fMSProject14::OnCreate(LPCREATESTRUCT lpCreateStruct)
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

BOOL fMSProject14::OnInitDialog()
{
	CDialog::OnInitDialog();
	g_MaximizeWindowsDim(this);
	CWnd::ShowWindow(SW_SHOWMAXIMIZED);

	SetIcon(m_hIcon, TRUE);			// Set big icon
	SetIcon(m_hIcon, FALSE);		// Set small icon



	// Reposition the Toolbar
	CWnd::RepositionBars(AFX_IDW_CONTROLBAR_FIRST, AFX_IDW_CONTROLBAR_LAST,0);

	InitializeAG();
	ActiveGanttVCCtl1.Redraw();

	return TRUE;
}

void fMSProject14::OnSize(UINT nType, int cx, int cy)
{
	CDialog::OnSize(nType, cx, cy);
	g_Resize(this, &ActiveGanttVCCtl1);
}

void fMSProject14::InitializeAG()
{

	CclsStyle oStyle;
	CclsView oView;

	CControlTemplateGradient oTemplate;
	oTemplate.CreateDispatch(_T("AGVC.ControlTemplateGradient"));
    oTemplate.SetControlBorderColor(FromArgb(255, 100, 145, 204));
    oTemplate.SetHeadingBorderColor(FromArgb(255, 197, 206, 216));
    oTemplate.SetHeadingForeColor(CLR_BLACK);
    oTemplate.SetRowForeColor(CLR_BLACK);
	oTemplate.SetGradientFillMode(GDT_VERTICAL);
    oTemplate.SetStartGradientColor(FromArgb(255, 210, 210, 210));
    oTemplate.SetEndGradientColor(FromArgb(255, 3, 167, 208));
    oTemplate.SetTreeviewColor(FromArgb(255, 100, 145, 204));
    oTemplate.SetRowBorderColor(FromArgb(255, 192, 192, 192));

    ActiveGanttVCCtl1.ApplyTemplateGradient(oTemplate, STO_DEFAULT);

	oStyle = ActiveGanttVCCtl1.GetStyles().Item(_T("T1"));
	oStyle.SetTextPlacement(SCP_EXTERIORPLACEMENT);
	oStyle.SetTextAlignmentHorizontal(HAL_RIGHT);
	oStyle.SetTextXMargin(30);



	ActiveGanttVCCtl1.SetAllowRowMove(TRUE);
	ActiveGanttVCCtl1.SetAllowRowSize(TRUE);
	ActiveGanttVCCtl1.SetAllowAdd(FALSE);
	ActiveGanttVCCtl1.GetStyle().SetAppearance(SA_FLAT);
	ActiveGanttVCCtl1.GetStyle().SetBorderStyle(SBR_CUSTOM);
	ActiveGanttVCCtl1.GetStyle().SetBorderColor(FromArgb(255, 197, 206, 216));
	ActiveGanttVCCtl1.GetStyle().SetBorderWidth(1);

	ActiveGanttVCCtl1.GetSplitter().SetType(SA_USERDEFINED);
	ActiveGanttVCCtl1.GetSplitter().SetWidth(1);
	ActiveGanttVCCtl1.GetSplitter().SetColor(1, FromArgb(255, 197, 206, 216));
	ActiveGanttVCCtl1.GetSplitter().SetPosition(285);

	ActiveGanttVCCtl1.GetTreeview().SetImages(TRUE);
	ActiveGanttVCCtl1.GetTreeview().SetCheckBoxes(TRUE);
	ActiveGanttVCCtl1.GetTreeview().SetPlusMinusBorderColor(FromArgb(255, 191, 205, 219));
	ActiveGanttVCCtl1.GetTreeview().SetTreeLines(FALSE);

	CclsColumn oColumn;

	oColumn = ActiveGanttVCCtl1.GetColumns().Add(_T("ID"), _T(""), 30, _T(""));
	oColumn.SetAllowTextEdit(TRUE);

	oColumn = ActiveGanttVCCtl1.GetColumns().Add(_T("Task Name"), _T(""), 300, _T(""));
	oColumn.SetAllowTextEdit(TRUE);

	ActiveGanttVCCtl1.SetTreeviewColumnIndex(2);

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

	oView = oView.Clone(_T(""));
    oView.SetInterval(IL_HOUR);
    oView.SetFactor(12);

	oView = oView.Clone(_T(""));
    oView.SetInterval(IL_HOUR);
    oView.SetFactor(6);

	oView = oView.Clone(_T(""));
    oView.SetInterval(IL_HOUR);
    oView.SetFactor(3);

	oView = oView.Clone(_T(""));
    oView.SetInterval(IL_HOUR);
    oView.SetFactor(1);

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
		ActiveGanttVCCtl1.GetPrinter().SetInterval(IL_DAY);
		ActiveGanttVCCtl1.GetPrinter().SetFactor(1);
	}

	ActiveGanttVCCtl1.SetTimeLineScrollBarScope(OS_VIEW);

	ActiveGanttVCCtl1.SetCurrentView(_T("1"));

}



void fMSProject14::MP14_To_AG()
{
	MSP2010::CMP14 oMP14;
	oMP14 = MSP2010Ctl1.GetMP14();
	CclsTask oAGTask;
	CclsRow oAGRow;
	MSP2010::CTask oMPTask;
	int i;
	int j;
	
	for (i = 1; i <= oMP14.GetoTasks().GetCount(); i++)
	{
		oMPTask = oMP14.GetoTasks().Item(CStr(i));
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
	for (i = 1; i <= oMP14.GetoTasks().GetCount(); i++)
	{
		oMPTask = oMP14.GetoTasks().Item(CStr(i));
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
			MSP2010::CTaskPredecessorLink oMPPredecessor;
			oMPPredecessor = oMPTask.GetoPredecessorLink_C().Item(CStr(j));
			ActiveGanttVCCtl1.GetPredecessors().Add(_T("K") + CStr(oMPTask.GetlUID()), _T("K") + CStr(oMPPredecessor.GetlPredecessorUID()), GetAGPredecessorType((MSP2010::E_TYPE_5)oMPPredecessor.GetyType()), _T(""), _T("T1"));
		}
	}


    // Assignments
	for (i = 1; i <= oMP14.GetoAssignments().GetCount(); i++)
	{
		MSP2010::CAssignment oAssignment;
		oAssignment = oMP14.GetoAssignments().Item(CStr(i));
		oAGTask = ActiveGanttVCCtl1.GetTasks().Item(_T("K") + CStr(oAssignment.GetlTaskUID()));
		if (oAGTask.GetStartDate() != oAGTask.GetEndDate())
		{
			if (oAssignment.GetlResourceUID() > 0)
			{
				if (oAGTask.GetText().GetLength() == 0)
				{
					oAGTask.SetText(oMP14.GetoResources().Item(_T("K") + CStr(oAssignment.GetlResourceUID())).GetsName());
				}
				else
				{
					oAGTask.SetText(oAGTask.GetText() + _T(", ") + oMP14.GetoResources().Item(_T("K") + CStr(oAssignment.GetlResourceUID())).GetsName());
				}
			}
		}
	}
    mp_dtStartDate = ActiveGanttVCCtl1.GetMathLib().DateTimeAdd(IL_HOUR, -3, mp_dtStartDate);
    AGSetStartDate(mp_dtStartDate);
}

ActiveGanttVC::E_CONSTRAINTTYPE fMSProject14::GetAGPredecessorType(MSP2010::E_TYPE_5 MPPredecessorType)
{
	switch (MPPredecessorType)
	{
	case T_5_FF:
		{
			return PCT_END_TO_END;
		}
		break;
	case T_5_FS:
		{
			return PCT_END_TO_START;
		}
		break;
	case T_5_SF:
		{
			return PCT_START_TO_END;
		}
		break;
	case T_5_SS:
		{
			return PCT_START_TO_START;
		}
		break;
	}
	return (ActiveGanttVC::E_CONSTRAINTTYPE)0;
}

void fMSProject14::AGSetStartDate(COleDateTime dtDate)
{
	int i;
	for (i = 1; i <= ActiveGanttVCCtl1.GetViews().GetCount(); i++)
	{
		ActiveGanttVCCtl1.GetViews().Item(CStr(i)).GetTimeLine().GetTimeLineScrollBar().SetStartDate(dtDate);
	}
}

void fMSProject14::cmdSaveXML()
{
	CFileDialog oDialog(FALSE, NULL, NULL, OFN_OVERWRITEPROMPT,_T("XML File|*.xml||"), this, NULL);
	oDialog.m_ofn.lpstrTitle = _T("Save As MS-Project 2010 XML File");
	if (oDialog.DoModal() == IDOK)
	{
		//AG_To_MP11();
		CString sFilePath = oDialog.GetPathName();
		MSP2010Ctl1.GetMP14().WriteXML(sFilePath);
	}
}

void fMSProject14::cmdLoadXML()
{
	TCHAR sDefaultFileName[MAX_PATH * 5];
	TCHAR sInitialDir[MAX_PATH * 5];
	wcscpy_s(sDefaultFileName, MAX_PATH * 5, _T(""));
	wcscpy_s(sInitialDir, MAX_PATH * 5, g_GetAppLocation() + _T("\\MSP2010"));
	CFileDialog oDialog(TRUE, NULL, NULL, OFN_HIDEREADONLY | OFN_OVERWRITEPROMPT,_T("XML File (*.xml)|*.xml|All Files (*.*)|*.*||"), this, NULL);
	oDialog.m_ofn.lpstrTitle = _T("Load MS-Project 2010 XML File");
	oDialog.m_ofn.lpstrFile = sDefaultFileName;
	oDialog.m_ofn.lpstrInitialDir = sInitialDir;
	oDialog.m_ofn.nMaxFile = MAX_PATH * 5;
	if (oDialog.DoModal() == IDOK)
	{
		CWaitCursor oWaitCursor;
		CString sFilePath = oDialog.GetPathName();
		ActiveGanttVCCtl1.Clear();
		MSP2010Ctl1.Clear();
        MSP2010Ctl1.GetMP14().ReadXML(sFilePath);
		oWaitCursor.Restore();
        InitializeAG();
        MP14_To_AG();
        ActiveGanttVCCtl1.Redraw();
	}
}
void fMSProject14::cmdPrint_Click()
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
void fMSProject14::cmdIndent_Click()
{
}

void fMSProject14::cmdZoomIn_Click()
{
	if (CInt32(ActiveGanttVCCtl1.GetCurrentView()) < ActiveGanttVCCtl1.GetViews().GetCount())
	{
		ActiveGanttVCCtl1.SetCurrentView(CStr(CInt32(ActiveGanttVCCtl1.GetCurrentView()) + 1));
        ActiveGanttVCCtl1.Redraw();
	}
}
void fMSProject14::cmdZoomOut_Click()
{
	if (CInt32(ActiveGanttVCCtl1.GetCurrentView()) > 1)
	{
        ActiveGanttVCCtl1.SetCurrentView(CStr(CInt32(ActiveGanttVCCtl1.GetCurrentView()) - 1));
        ActiveGanttVCCtl1.Redraw();
	}
}

void fMSProject14::mnuLoadXML_Click()
{
	cmdLoadXML();
}

void fMSProject14::mnuSaveXML_Click()
{
	cmdSaveXML();
}

void fMSProject14::mnuPrint_Click()
{
	cmdPrint_Click();
}

void fMSProject14::mnuClose_Click()
{
	CDialog::EndDialog(IDOK);
}
BEGIN_EVENTSINK_MAP(fMSProject14, CDialog)
	ON_EVENT(fMSProject14, IDC_ACTIVEGANTTVCCTL1, 13, fMSProject14::CustomTierDrawActiveganttvcctl1, VTS_DISPATCH)
	ON_EVENT(fMSProject14, IDC_ACTIVEGANTTVCCTL1, 35, fMSProject14::ControlMouseWheelActiveganttvcctl1, VTS_DISPATCH)
	ON_EVENT(fMSProject14, IDC_ACTIVEGANTTVCCTL1, 42, fMSProject14::PagePrintActiveganttvcctl1, VTS_DISPATCH)
END_EVENTSINK_MAP()

void fMSProject14::CustomTierDrawActiveganttvcctl1(LPDISPATCH e)
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

void fMSProject14::ControlMouseWheelActiveganttvcctl1(LPDISPATCH e)
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

void fMSProject14::PagePrintActiveganttvcctl1(LPDISPATCH e)
{
		CPagePrintEventArgs oE(e);

		CclsFont oFont;
		oFont.CreateDispatch(_T("AGVC.clsFont"));
		oFont.SetFontName(_T("Arial"));
		oFont.SetSize(8);

        ULONG oColor;
        LONG X2;
        oColor = FromArgb(255, 150, 158, 168);
		if ((oE.GetX2() - oE.GetX1()) < CInt32(CDbl(oE.GetPageWidth()) * 0.85))
		{
			X2 = oE.GetPageWidth() - ActiveGanttVCCtl1.GetPrinter().GetMarginRight();
		}
		else
		{
			X2 = oE.GetX2();
		}


        ActiveGanttVCCtl1.GetDrawing().DrawLine(oE.GetX1(), oE.GetPageHeight() - 190, X2, oE.GetPageHeight() - 190, oColor, LDS_SOLID, 1);
        ActiveGanttVCCtl1.GetDrawing().DrawAlignedText(oE.GetX1() + 10, oE.GetPageHeight() - 190, oE.GetX1() + 290, oE.GetPageHeight() - 70, _T("ActiveGanttVC Scheduler Component\r\nMicrosoft Project 2010 Integration"), HAL_LEFT, VAL_CENTER, CLR_BLACK, oFont);

        int lLine = oE.GetPageHeight() - 180;
        ActiveGanttVCCtl1.GetDrawing().DrawAlignedText(oE.GetX1() + 310, lLine, oE.GetX1() + 400, lLine + 13, _T("Summary"), HAL_LEFT, VAL_TOP, CLR_BLACK, oFont);
        ActiveGanttVCCtl1.GetDrawing().DrawObject(oE.GetX1() + 450, lLine, oE.GetX1() + 600, lLine + 10, _T("S1"), _T(""), FALSE, NULL, DO_GENERAL);

        lLine = oE.GetPageHeight() - 160;
        ActiveGanttVCCtl1.GetDrawing().DrawAlignedText(oE.GetX1() + 310, lLine, oE.GetX1() + 400, lLine + 13, _T("Task"), HAL_LEFT, VAL_TOP, CLR_BLACK, oFont);
        ActiveGanttVCCtl1.GetDrawing().DrawObject(oE.GetX1() + 450, lLine, oE.GetX1() + 600, lLine + 10, _T("T1"), _T(""), FALSE, NULL, DO_GENERAL);

        lLine = oE.GetPageHeight() - 140;
        ActiveGanttVCCtl1.GetDrawing().DrawAlignedText(oE.GetX1() + 310, lLine, oE.GetX1() + 400, lLine + 13, _T("Milestone"), HAL_LEFT, VAL_TOP, CLR_BLACK, oFont);
        ActiveGanttVCCtl1.GetDrawing().DrawObject(oE.GetX1() + 445, lLine, oE.GetX1() + 605, lLine + 10, _T("T1"), _T(""), FALSE, NULL, DO_GENERAL);

        ActiveGanttVCCtl1.GetDrawing().DrawLine(oE.GetX1() + 300, oE.GetPageHeight() - 190, oE.GetX1() + 300, oE.GetPageHeight() - 70, oColor, LDS_SOLID, 1);

        ActiveGanttVCCtl1.GetDrawing().DrawLine(oE.GetX1(), oE.GetPageHeight() - 70, X2, oE.GetPageHeight() - 70, oColor, LDS_SOLID, 1);
        ActiveGanttVCCtl1.GetDrawing().DrawAlignedText(oE.GetX1(), oE.GetPageHeight() - 70, X2, oE.GetPageHeight() - 50, _T("Page ") + CStr(oE.GetPageNumber()), HAL_CENTER, VAL_CENTER, CLR_BLACK, oFont);

        ActiveGanttVCCtl1.GetDrawing().DrawBorder(oE.GetX1(), oE.GetY1(), X2, oE.GetPageHeight() - 50, oColor, LDS_SOLID, 1);

}


void fMSProject14::OnClose()
{
	CDialog::OnClose();
}