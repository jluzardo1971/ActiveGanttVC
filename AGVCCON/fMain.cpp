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
#include "fMain.h"
#include "fWBSProject.h"
#include "fCarRental.h"
#include "fMSProject11.h"
#include "fMSProject12.h"
#include "fMSProject14.h"
#include "fFastLoading.h"
#include "fCustomDrawing.h"
#include "fSortRows.h"
#include "fDurationTasks.h"
#include "fRCT_DAY.h"
#include "fRCT_WEEK.h"
#include "fRCT_MONTH.h"
#include "fRCT_YEAR.h"
#include "fCustomTickMarkTextDraw.h"
#include "fCustomTickMarkAreaDraw.h"
#include "fTemplates.h"
#include "fSTL_Anakiwa_Blue.h"
#include "fSTL_MSP2003.h"
#include "fSTL_MSP2007.h"
#include "fSTL_MSP2010.h"
#include "fSTL_CR.h"

IMPLEMENT_DYNAMIC(fMain, CDialog)

fMain::fMain(CWnd* pParent /*=NULL*/)
	: CDialog(fMain::IDD, pParent)
{
	m_hIcon = AfxGetApp()->LoadIcon(IDR_MAINFRAME);
	mp_olblCopyrightBrush.CreateSolidBrush(RGB(192, 192, 255));
}

fMain::~fMain()
{
}

void fMain::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	DDX_Control(pDX, IDC_TREE1, TreeView1);
}


BEGIN_MESSAGE_MAP(fMain, CDialog)
	ON_WM_PAINT()
	ON_WM_QUERYDRAGICON()
	ON_WM_CTLCOLOR()
	ON_NOTIFY(NM_DBLCLK, IDC_TREE1, &fMain::OnNMDblclkTree1)
END_MESSAGE_MAP()

BOOL fMain::OnInitDialog()
{
	CDialog::OnInitDialog();

	ASSERT((IDM_ABOUTBOX & 0xFFF0) == IDM_ABOUTBOX);
	ASSERT(IDM_ABOUTBOX < 0xF000);

	CMenu* pSysMenu = GetSystemMenu(FALSE);
	if (pSysMenu != NULL)
	{
		CString strAboutMenu;
		strAboutMenu.LoadString(IDS_ABOUTBOX);
		if (!strAboutMenu.IsEmpty())
		{
			pSysMenu->AppendMenu(MF_SEPARATOR);
			pSysMenu->AppendMenu(MF_STRING, IDM_ABOUTBOX, strAboutMenu);
		}
	}

	SetIcon(m_hIcon, TRUE);			// Set big icon
	SetIcon(m_hIcon, FALSE);		// Set small icon

	COLORREF clrTransparent;
	clrTransparent = RGB(255,255,255);

	ImageList1.Create(16, 16, ILC_MASK | ILC_COLOR32, 0, 0);
	CBitmap oFolderClosed;
	CBitmap oFolderOpen;
	CBitmap oAG;
	CBitmap oInternet;
	CBitmap oBlueFolderClosed;
	CBitmap oBlueFolderOpen;
	CBitmap oOnlineDocumentation;
	CBitmap oLocalCHMDocumentation;
	CBitmap oGettingStarted;
	oFolderOpen.LoadBitmapW(IDB_FOLDEROPEN);
	oFolderClosed.LoadBitmapW(IDB_FOLDERCLOSED);
	oAG.LoadBitmapW(IDB_AG);
	oInternet.LoadBitmapW(IDB_INTERNET);
	oBlueFolderClosed.LoadBitmapW(IDB_BLUEFOLDERCLOSED);
	oBlueFolderOpen.LoadBitmapW(IDB_BLUEFOLDEROPEN);
	oOnlineDocumentation.LoadBitmapW(IDB_ONLINEDOCUMENTATION);
	oLocalCHMDocumentation.LoadBitmapW(IDB_LOCALCHMDOCUMENTATION);
	oGettingStarted.LoadBitmapW(IDB_GETTINGSTARTED);
	ImageList1.Add(&oFolderClosed, clrTransparent);
	ImageList1.Add(&oFolderOpen, clrTransparent);
	ImageList1.Add(&oAG, clrTransparent);
	ImageList1.Add(&oInternet, clrTransparent);
	ImageList1.Add(&oBlueFolderClosed, clrTransparent);
	ImageList1.Add(&oBlueFolderOpen, clrTransparent);
	ImageList1.Add(&oOnlineDocumentation, clrTransparent);
	ImageList1.Add(&oLocalCHMDocumentation, clrTransparent);
	ImageList1.Add(&oGettingStarted, clrTransparent);


	TreeView1.SetImageList(&ImageList1, TVSIL_NORMAL);

	COleDateTime dtCurrentDate = COleDateTime::GetCurrentTime();

	CStatic* lblCopyright = (CStatic*) GetDlgItem(IDC_LBLCOPYRIGHT);
	lblCopyright->SetWindowTextW(_T("Copyright ©2002-") + CStr(dtCurrentDate.GetYear()) + _T(" The Source Code Store LLC. All Rights Reserved. All trademarks are property of their legal owner."));


    AddTitleNode(_T("AGEX"), _T("ActiveGantt Examples:"), 4, 5);

    AddNode(_T("AGEX"), _T("GanttCharts"), _T("Gantt Charts:"), 4, 5);

    AddNode(_T("GanttCharts"), _T("WBS"), _T("Work Breakdown Structure (WBS) Project Management Examples:"), 4, 5);
    AddNode(_T("WBS"), _T("WBSProject"), _T("Microsoft Access data source"), 2, 2);

    AddNode(_T("GanttCharts"), _T("MSPI"), _T("Microsoft Project Integration Examples:"), 4, 5);
    AddNode(_T("MSPI"), _T("Project2003"), _T("Demonstrates how ActiveGantt integrates with MS Project 2003 (using XML Files and the MSP2003.ocx)"), 2, 2);
    AddNode(_T("MSPI"), _T("Project2007"), _T("Demonstrates how ActiveGantt integrates with MS Project 2007 (using XML Files and the MSP2007.ocx)"), 2, 2);
    AddNode(_T("MSPI"), _T("Project2010"), _T("Demonstrates how ActiveGantt integrates with MS Project 2010 (using XML Files and the MSP2010.ocx)"), 2, 2);

    AddNode(_T("AGEX"), _T("Schedules"), _T("Schedules and Rosters:"), 4, 5);
    AddNode(_T("Schedules"), _T("VRFC"), _T("Vehicle Rental/Fleet Control Roster Examples:"), 4, 5);
    AddNode(_T("VRFC"), _T("CarRental"), _T("Microsoft Access data source"), 2, 2);

	AddNode(_T("AGEX"), _T("TEMPL"), _T("Styles and Templates:"), 4, 5);
    AddNode(_T("TEMPL"), _T("Templates"), _T("Available Templates"), 2, 2);
    AddNode(_T("TEMPL"), _T("Anakiwa_Blue_Style"), _T("Anakiwa Blue Style without using Templates"), 2, 2);
    AddNode(_T("TEMPL"), _T("MSP2003Style"), _T("MS Project 2003 Style without using Templates"), 2, 2);
    AddNode(_T("TEMPL"), _T("MSP2007Style"), _T("MS Project 2007 Style without using Templates"), 2, 2);
    AddNode(_T("TEMPL"), _T("MSP2010Style"), _T("MS Project 2010 Style without using Templates"), 2, 2);
    AddNode(_T("TEMPL"), _T("CRS"), _T("Car Rental Style without using Templates"), 2, 2);

	AddNode(_T("AGEX"), _T("OTHER"), _T("Other examples:"), 4, 5);
	AddNode(_T("OTHER"), _T("FastLoad"), _T("Fast Loading of Row and Task objects"), 2, 2);
	AddNode(_T("OTHER"), _T("CustomDrawing"), _T("Custom Drawing"), 2, 2);
	AddNode(_T("OTHER"), _T("SortRows"), _T("Sort Rows"), 2, 2);
	AddNode(_T("OTHER"), _T("5MillisecondInterval"), _T("5 Millisecond Interval View"), 2, 2);
	AddNode(_T("OTHER"), _T("CustomTickMarkArea"), _T("Custom TickMarkArea Drawing"), 2, 2);
	AddNode(_T("OTHER"), _T("CustomTickMarkText"), _T("Custom TickMark Text Drawing"), 2, 2);

	AddNode(_T("OTHER"), _T("TimeBlocks"), _T("TimeBlocks and Duration Tasks:"), 4, 5);
	AddNode(_T("TimeBlocks"), _T("RCT_DAY"), _T("Daily Recurrent TimeBlocks"), 2, 2);
	AddNode(_T("TimeBlocks"), _T("RCT_WEEK"), _T("Weekly Recurrent TimeBlocks"), 2, 2);
	AddNode(_T("TimeBlocks"), _T("RCT_MONTH"), _T("Monthly Recurrent TimeBlocks"), 2, 2);
	AddNode(_T("TimeBlocks"), _T("RCT_YEAR"), _T("Yearly Recurrent TimeBlocks"), 2, 2);
	AddNode(_T("TimeBlocks"), _T("DurationTasks"), _T("Duration Tasks (can skip over non-working TimeBlock intervals)"), 2, 2);

	ExpandAll();
	Collapse(_T("OTHER"));

	return TRUE;
}

void fMain::OnPaint()
{
	if (IsIconic())
	{
		CPaintDC dc(this);

		SendMessage(WM_ICONERASEBKGND, reinterpret_cast<WPARAM>(dc.GetSafeHdc()), 0);

		// Center icon in client rectangle
		int cxIcon = GetSystemMetrics(SM_CXICON);
		int cyIcon = GetSystemMetrics(SM_CYICON);
		CRect rect;
		GetClientRect(&rect);
		int x = (rect.Width() - cxIcon + 1) / 2;
		int y = (rect.Height() - cyIcon + 1) / 2;

		// Draw the icon
		dc.DrawIcon(x, y, m_hIcon);
	}
	else
	{
		CDialog::OnPaint();
	}
}

HCURSOR fMain::OnQueryDragIcon()
{
	return static_cast<HCURSOR>(m_hIcon);
}

HBRUSH fMain::OnCtlColor(CDC* pDC, CWnd* pWnd, UINT nCtlColor)
{
	HBRUSH hbr = CDialog::OnCtlColor(pDC, pWnd, nCtlColor);

	if (pWnd->GetDlgCtrlID() == IDC_LBLCOPYRIGHT)
	{
		pDC->SetTextColor(RGB(0, 0, 0));
		pDC->SetBkColor(RGB(192, 192, 255));
		
		return (HBRUSH)mp_olblCopyrightBrush.GetSafeHandle();
	}
	return hbr;
}

void fMain::AddTitleNode(CString sKey, CString sText, long ImageIndex, long SelectedImageIndex)
{
	HTREEITEM hItem;
	hItem = TreeView1.InsertItem(sText, ImageIndex, SelectedImageIndex, TVI_ROOT);
	oKeys.SetAt(sKey, hItem);
}

void fMain::AddNode(CString sParentKey, CString sKey, CString sText, long ImageIndex, long SelectedImageIndex)
{
	HTREEITEM hParentItem;
	HTREEITEM hItem;
	t_TreeKeys::CPair* oPair;
	oPair = oKeys.PLookup(sParentKey);
	hParentItem = (HTREEITEM)oPair->value;
	hItem = TreeView1.InsertItem(sText, ImageIndex, SelectedImageIndex, hParentItem);
	oKeys.SetAt(sKey, hItem);

}

void fMain::ExpandAll()
{
	HTREEITEM hItem;
	hItem = TreeView1.GetFirstVisibleItem();
	while (hItem != NULL)
	{
		TreeView1.Expand(hItem, TVE_EXPAND);
		hItem = TreeView1.GetNextItem(hItem, TVGN_NEXTVISIBLE);
	}
}

void fMain::Collapse(CString sKey)
{
	t_TreeKeys::CPair* oCurrentPair = NULL;

	oCurrentPair = oKeys.PGetFirstAssoc();

	while (oCurrentPair != NULL)
    {
		CString sItemKey;
		HTREEITEM hItem;
		sItemKey = oCurrentPair->key;
		hItem = oCurrentPair->value;
		if (sItemKey == sKey)
		{
			TreeView1.Expand(hItem, TVE_COLLAPSE);
		}
        oCurrentPair = oKeys.PGetNextAssoc(oCurrentPair);
    }


}

void fMain::OnNMDblclkTree1(NMHDR *pNMHDR, LRESULT *pResult)
{
	// TODO: Add your control notification handler code here

	HTREEITEM hSelectedItem;

	CPoint oCursorPos;
	GetCursorPos(&oCursorPos);
	TreeView1.ScreenToClient(&oCursorPos);
	hSelectedItem = TreeView1.HitTest(oCursorPos);
	CString sSelectedKey;

	t_TreeKeys::CPair* oCurrentPair = NULL;

	oCurrentPair = oKeys.PGetFirstAssoc();

	while (oCurrentPair != NULL)
    {
		CString sKey;
		HTREEITEM hItem;
		sKey = oCurrentPair->key;
		hItem = oCurrentPair->value;
		if (hItem == hSelectedItem)
		{
			sSelectedKey = sKey;
			break;
		}
        oCurrentPair = oKeys.PGetNextAssoc(oCurrentPair);
    }

    if (sSelectedKey == _T("WBSProject"))
	{
		fWBSProject oForm;
		oForm.DoModal();
	}
	else if (sSelectedKey == _T("CarRental"))
	{
		fCarRental oForm;
		oForm.DoModal();
	}
	else if (sSelectedKey == _T("Project2003"))
	{
		fMSProject11 oForm;
		oForm.DoModal();
	}
	else if (sSelectedKey == _T("Project2007"))
	{
		fMSProject12 oForm;
		oForm.DoModal();
	}
	else if (sSelectedKey == _T("Project2010"))
	{
		fMSProject14 oForm;
		oForm.DoModal();
	}
	else if (sSelectedKey == _T("FastLoad"))
	{
		fFastLoading oForm;
		oForm.DoModal();
	}
	else if (sSelectedKey == _T("CustomDrawing"))
	{
		fCustomDrawing oForm;
		oForm.DoModal();
	}
	else if (sSelectedKey == _T("SortRows"))
	{
		fSortRows oForm;
		oForm.DoModal();
	}
	else if (sSelectedKey == _T("DurationTasks"))
	{
		fDurationTasks oForm;
		oForm.DoModal();
	}
	else if (sSelectedKey == _T("RCT_DAY"))
	{
		fRCT_DAY oForm;
		oForm.DoModal();
	}
	else if (sSelectedKey == _T("RCT_WEEK"))
	{
		fRCT_WEEK oForm;
		oForm.DoModal();
	}
	else if (sSelectedKey == _T("RCT_MONTH"))
	{
		fRCT_MONTH oForm;
		oForm.DoModal();
	}
	else if (sSelectedKey == _T("RCT_YEAR"))
	{
		fRCT_YEAR oForm;
		oForm.DoModal();
	}
	else if (sSelectedKey == _T("CustomTickMarkArea"))
	{
		fCustomTickMarkAreaDraw oForm;
		oForm.DoModal();
	}
	else if (sSelectedKey == _T("CustomTickMarkText"))
	{
		fCustomTickMarkTextDraw oForm;
		oForm.DoModal();
	}
	else if (sSelectedKey == _T("Templates"))
	{
		fTemplates oForm;
		oForm.DoModal();
	}
	else if (sSelectedKey == _T("Anakiwa_Blue_Style"))
	{
		fSTL_Anakiwa_Blue oForm;
		oForm.DoModal();
	}
	else if (sSelectedKey == _T("MSP2003Style"))
	{
		fSTL_MSP2003 oForm;
		oForm.DoModal();
	}
	else if (sSelectedKey == _T("MSP2007Style"))
	{
		fSTL_MSP2007 oForm;
		oForm.DoModal();
	}
	else if (sSelectedKey == _T("MSP2010Style"))
	{
		fSTL_MSP2010 oForm;
		oForm.DoModal();
	}
	else if (sSelectedKey == _T("CRS"))
	{
		fSTL_CR oForm;
		oForm.DoModal();
	}
	*pResult = 0;
}