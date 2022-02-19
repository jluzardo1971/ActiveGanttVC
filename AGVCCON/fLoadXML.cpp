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
#include "fLoadXML.h"

IMPLEMENT_DYNAMIC(fLoadXML, CDialog)

fLoadXML::fLoadXML(CWnd* pParent /*=NULL*/)
	: CDialog(fLoadXML::IDD, pParent)
{
	m_hIcon = AfxGetApp()->LoadIcon(IDR_MAINFRAME);
	bLoaded = FALSE;
}

fLoadXML::~fLoadXML()
{
}

void fLoadXML::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	DDX_Control(pDX, IDC_ACTIVEGANTTVCCTL1, ActiveGanttVCCtl1);
}


BEGIN_MESSAGE_MAP(fLoadXML, CDialog)
	ON_WM_CREATE()
	ON_WM_SIZE()

	ON_COMMAND(ID_MYTOOLBAR_BUTTONFIRST + 0, cmdLoadXML)
	ON_COMMAND(ID_MYTOOLBAR_BUTTONFIRST + 1, cmdSaveXML)

	ON_COMMAND(ID_FILE_LOADXML, mnuLoadXML_Click)
	ON_COMMAND(ID_FILE_SAVEXML, mnuSaveXML_Click)
	ON_COMMAND(ID_FILE_CLOSE, mnuClose_Click)

END_MESSAGE_MAP()

int fLoadXML::OnCreate(LPCREATESTRUCT lpCreateStruct)
{
	if (CDialog::OnCreate(lpCreateStruct) == -1)
		return -1;

	// TODO:  Add your specialized creation code here
	oToolBar1.CreateEx(this, TBSTYLE_FLAT, WS_CHILD | WS_VISIBLE | CBRS_TOP | CBRS_GRIPPER | CBRS_TOOLTIPS | CBRS_FLYBY | CBRS_SIZE_DYNAMIC);
	oToolBarBitmap.LoadBitmapW(IDB_FLOADXML);
	oToolBar1.SetBitmap((HBITMAP)oToolBarBitmap.GetSafeHandle());
	oToolBar1.SetButtons(NULL, 2); //1 Based
	oToolBar1.AddButton(0, _T("Load From XML"));
	oToolBar1.AddButton(1, _T("Save As XML"));

	return 0;
}

BOOL fLoadXML::OnInitDialog()
{
	CDialog::OnInitDialog();
	g_MaximizeWindowsDim(this);
	CWnd::ShowWindow(SW_SHOWMAXIMIZED);

	SetIcon(m_hIcon, TRUE);			// Set big icon
	SetIcon(m_hIcon, FALSE);		// Set small icon

	CWnd::RepositionBars(AFX_IDW_CONTROLBAR_FIRST, AFX_IDW_CONTROLBAR_LAST,0);

	return TRUE;  
}

void fLoadXML::OnSize(UINT nType, int cx, int cy)
{
	CDialog::OnSize(nType, cx, cy);
	g_Resize(this, &ActiveGanttVCCtl1);
}

void fLoadXML::cmdLoadXML()
{
	LoadXML();
}

void fLoadXML::cmdSaveXML()
{
	SaveXML();
}

void fLoadXML::mnuLoadXML_Click()
{
	LoadXML();
}

void fLoadXML::mnuSaveXML_Click()
{
	SaveXML();
}

void fLoadXML::mnuClose_Click()
{
	CDialog::EndDialog(IDOK);
}

void fLoadXML::SaveXML()
{
	TCHAR sDefaultFileName_WBSP[MAX_PATH];
	TCHAR sDefaultFileName_CR[MAX_PATH];
	wcscpy_s(sDefaultFileName_WBSP, MAX_PATH, _T("AGVC_WBSP.xml"));
	wcscpy_s(sDefaultFileName_CR, MAX_PATH, _T("AGVC_CR.xml"));
	CFileDialog oDialog(FALSE, NULL, NULL, OFN_OVERWRITEPROMPT,_T("XML File|*.xml||"), this, NULL);
	oDialog.m_ofn.lpstrTitle = _T("Save As XML");
	if (ActiveGanttVCCtl1.GetControlTag() == "WBSProject")
	{
		oDialog.m_ofn.lpstrFile = sDefaultFileName_WBSP; 
	}
	else
	{
		oDialog.m_ofn.lpstrFile = sDefaultFileName_CR;
	}
	oDialog.m_ofn.nMaxFile = MAX_PATH;
	if (oDialog.DoModal() == IDOK)
	{
		CString sFilePath = oDialog.GetPathName();
		ActiveGanttVCCtl1.WriteXML(sFilePath);
	}
}

void fLoadXML::LoadXML()
{
	//CFileDialog oDialog(TRUE);
	CFileDialog oDialog(TRUE, NULL, NULL, OFN_HIDEREADONLY | OFN_OVERWRITEPROMPT, _T("XML File (*.xml)|*.xml|All Files (*.*)|*.*||"), this, 0, TRUE);
	oDialog.m_ofn.lpstrTitle = _T("Load From XML");
	if (oDialog.DoModal() == IDOK)
	{
		CString sFilePath = oDialog.GetPathName();
		ActiveGanttVCCtl1.ReadXML(sFilePath);
		bLoaded = TRUE;
		ActiveGanttVCCtl1.Redraw();
	}
}

BEGIN_EVENTSINK_MAP(fLoadXML, CDialog)
ON_EVENT(fLoadXML, IDC_ACTIVEGANTTVCCTL1, 13, fLoadXML::CustomTierDrawActiveganttvcctl1, VTS_DISPATCH)
END_EVENTSINK_MAP()

void fLoadXML::CustomTierDrawActiveganttvcctl1(LPDISPATCH e)
{
	CCustomTierDrawEventArgs oE(e);
	if (bLoaded == FALSE)
	{
		return;
	}
	if (ActiveGanttVCCtl1.GetControlTag() == _T("CarRental"))
	{
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
}