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
#include "fCustomTickMarkTextDraw.h"
#include "afxdialogex.h"

IMPLEMENT_DYNAMIC(fCustomTickMarkTextDraw, CDialog)

fCustomTickMarkTextDraw::fCustomTickMarkTextDraw(CWnd* pParent /*=NULL*/)
	: CDialog(fCustomTickMarkTextDraw::IDD, pParent)
{

}

fCustomTickMarkTextDraw::~fCustomTickMarkTextDraw()
{
}

void fCustomTickMarkTextDraw::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	DDX_Control(pDX, IDC_ACTIVEGANTTVCCTL1, ActiveGanttVCCtl1);
}


BEGIN_MESSAGE_MAP(fCustomTickMarkTextDraw, CDialog)
	ON_WM_SIZE()
END_MESSAGE_MAP()

BEGIN_EVENTSINK_MAP(fCustomTickMarkTextDraw, CDialog)
	ON_EVENT(fCustomTickMarkTextDraw, IDC_ACTIVEGANTTVCCTL1, 41, fCustomTickMarkTextDraw::TickMarkTextDrawActiveganttvcctl1, VTS_DISPATCH)
END_EVENTSINK_MAP()

BOOL fCustomTickMarkTextDraw::OnInitDialog()
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

	return TRUE;  
}

void fCustomTickMarkTextDraw::OnSize(UINT nType, int cx, int cy)
{
	CDialog::OnSize(nType, cx, cy);
	g_Resize(this, &ActiveGanttVCCtl1);
}


void fCustomTickMarkTextDraw::TickMarkTextDrawActiveganttvcctl1(LPDISPATCH e)
{
	CTickMarkTextDrawEventArgs oE(e);
	oE.SetCustomDraw(TRUE);
    oE.SetText(FormatDate(oE.GetdtDate(), _T("%H")) + _T("H"));
}