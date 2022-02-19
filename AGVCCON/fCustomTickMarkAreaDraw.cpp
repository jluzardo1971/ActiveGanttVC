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
#include "fCustomTickMarkAreaDraw.h"
#include "afxdialogex.h"

IMPLEMENT_DYNAMIC(fCustomTickMarkAreaDraw, CDialog)

fCustomTickMarkAreaDraw::fCustomTickMarkAreaDraw(CWnd* pParent /*=NULL*/)
	: CDialog(fCustomTickMarkAreaDraw::IDD, pParent)
{

}

fCustomTickMarkAreaDraw::~fCustomTickMarkAreaDraw()
{
}

void fCustomTickMarkAreaDraw::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	DDX_Control(pDX, IDC_ACTIVEGANTTVCCTL1, ActiveGanttVCCtl1);
}


BEGIN_MESSAGE_MAP(fCustomTickMarkAreaDraw, CDialog)
	ON_WM_SIZE()
END_MESSAGE_MAP()

BEGIN_EVENTSINK_MAP(fCustomTickMarkAreaDraw, CDialog)
	ON_EVENT(fCustomTickMarkAreaDraw, IDC_ACTIVEGANTTVCCTL1, 39, fCustomTickMarkAreaDraw::CustomTickMarkAreaDrawActiveganttvcctl1, VTS_DISPATCH)
	ON_EVENT(fCustomTickMarkAreaDraw, IDC_ACTIVEGANTTVCCTL1, 40, fCustomTickMarkAreaDraw::CustomTickMarkDrawActiveganttvcctl1, VTS_DISPATCH)
END_EVENTSINK_MAP()

void fCustomTickMarkAreaDraw::OnSize(UINT nType, int cx, int cy)
{
	CDialog::OnSize(nType, cx, cy);
	g_Resize(this, &ActiveGanttVCCtl1);
}


BOOL fCustomTickMarkAreaDraw::OnInitDialog()
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

void fCustomTickMarkAreaDraw::CustomTickMarkAreaDrawActiveganttvcctl1(LPDISPATCH e)
{
	CCustomTickMarkAreaDrawEventArgs oE(e);
	ActiveGanttVCCtl1.GetDrawing().DrawRectangle(oE.GetLeft(), oE.GetTop(), oE.GetRight(), oE.GetBottom(), CLR_GHOSTWHITE, LDS_SOLID, 1); 
	oE.SetCustomDraw(TRUE);
}


void fCustomTickMarkAreaDraw::CustomTickMarkDrawActiveganttvcctl1(LPDISPATCH e)
{
	CCustomTickMarkAreaDrawEventArgs oE(e);

	int y1 = (oE.GetBottom() - oE.GetTop()) / 4;
    switch (oE.GetoTickMark().GetTickMarkType())
    {
        case TLT_BIG:
			ActiveGanttVCCtl1.GetDrawing().DrawLine(oE.GetX(), oE.GetTop(), oE.GetX(), oE.GetBottom(), CLR_RED, LDS_SOLID, 1);
            break;
        case TLT_MEDIUM:
			ActiveGanttVCCtl1.GetDrawing().DrawLine(oE.GetX(), oE.GetTop() + y1, oE.GetX(), oE.GetBottom(), CLR_GREEN, LDS_SOLID, 1);
            break;
        case TLT_SMALL:
			ActiveGanttVCCtl1.GetDrawing().DrawLine(oE.GetX(), oE.GetTop() + (2 * y1), oE.GetX(), oE.GetBottom(), CLR_GREEN, LDS_SOLID, 1);
            break;
    }
    oE.SetCustomDraw(TRUE);
}