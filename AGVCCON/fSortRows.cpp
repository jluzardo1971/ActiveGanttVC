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
#include "fSortRows.h"

IMPLEMENT_DYNAMIC(fSortRows, CDialog)

fSortRows::fSortRows(CWnd* pParent /*=NULL*/)
	: CDialog(fSortRows::IDD, pParent)
{
	mp_bDescending = FALSE;
}

fSortRows::~fSortRows()
{
}

void fSortRows::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	DDX_Control(pDX, IDC_ACTIVEGANTTVCCTL1, ActiveGanttVCCtl1);
}


BEGIN_MESSAGE_MAP(fSortRows, CDialog)
	ON_BN_CLICKED(IDC_CMDSORTROWS, &fSortRows::OnBnClickedCmdsortrows)
	ON_WM_SIZE()
END_MESSAGE_MAP()

BOOL fSortRows::OnInitDialog()
{
	CDialog::OnInitDialog();
	g_MaximizeWindowsDim(this);
	CWnd::ShowWindow(SW_SHOWMAXIMIZED);

	//If you open the form: Styles And Templates -> Available Templates in the main menu (fTemplates.vb)
	//you can preview all available Templates.
	//Or you can also build your own:
	//fMSProject11.vb shows you how to build a Solid Template in the InitializeAG Method.
	//fMSProject14.vb shows you how to build a Gradient Template in the InitializeAG Method.
	ActiveGanttVCCtl1.ApplyTemplate(STC_CH_VGRAD_ANAKIWA_BLUE, STO_DEFAULT);

	int i;
	ActiveGanttVCCtl1.GetColumns().Add(_T(""), _T("C1"), 125, _T(""));
	for (i = 1; i <= 10; i++)
	{
		CString si;
		si = CStr(i);
		while (si.GetLength() < 2)
		{
			si = _T("0") + si;
		}
		ActiveGanttVCCtl1.GetRows().Add(_T("K") + si, _T("K") + si, TRUE, TRUE, _T(""));
	}
	return TRUE;  
}

void fSortRows::OnBnClickedCmdsortrows()
{
	mp_bDescending = !mp_bDescending;
	ActiveGanttVCCtl1.GetRows().SortRows(_T("Text"), mp_bDescending, ES_STRING, -1, -1);
    ActiveGanttVCCtl1.Redraw();
}

void fSortRows::OnSize(UINT nType, int cx, int cy)
{
	CDialog::OnSize(nType, cx, cy);
	g_Resize(this, &ActiveGanttVCCtl1);
}