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
#include "fCustomDrawing.h"


IMPLEMENT_DYNAMIC(fCustomDrawing, CDialog)

fCustomDrawing::fCustomDrawing(CWnd* pParent /*=NULL*/)
	: CDialog(fCustomDrawing::IDD, pParent)
{

}

fCustomDrawing::~fCustomDrawing()
{
}

void fCustomDrawing::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	DDX_Control(pDX, IDC_ACTIVEGANTTVCCTL1, ActiveGanttVCCtl1);
}


BEGIN_MESSAGE_MAP(fCustomDrawing, CDialog)
	ON_WM_SIZE()
END_MESSAGE_MAP()

BOOL fCustomDrawing::OnInitDialog()
{
	CDialog::OnInitDialog();
	g_MaximizeWindowsDim(this);
	CWnd::ShowWindow(SW_SHOWMAXIMIZED);

	ActiveGanttVCCtl1.ApplyTemplate(STC_CH_HGRAD_BLUE_BELL, STO_VARIATION_1);

	this->SetWindowTextW(_T("Custom Drawing"));

	int i = 0;
	ActiveGanttVCCtl1.GetColumns().Add(_T("Column 1"), _T(""), 125, _T(""));
	ActiveGanttVCCtl1.GetColumns().Add(_T("Column 2"), _T(""), 100, _T(""));
	for (i = 1; i <= 10; i++)
	{
		ActiveGanttVCCtl1.GetRows().Add(_T("K") + CStr(i), _T("Row ") + CStr(i) + _T(" (Key: K") + CStr(i) + _T(")"), TRUE, TRUE, _T(""));
	}

	ActiveGanttVCCtl1.GetCurrentViewObject().GetTimeLine().Position(GetDateTime(2011, 11, 21, 0, 0, 0));

	CclsTask oTask;

	oTask = ActiveGanttVCCtl1.GetTasks().Add(_T("Task 1"), _T("K1"), GetDateTime(2011, 11, 21, 0, 0, 0), GetDateTime(2011, 11, 21, 3, 0, 0), _T(""), _T("RET1"), _T(""));
	oTask.SetAllowedMovement(MT_RESTRICTEDTOROW);
	oTask.SetAllowStretchLeft(TRUE);
	oTask.SetAllowStretchRight(TRUE);
	ActiveGanttVCCtl1.GetTasks().Add(_T("Task 2"), _T("K2"), GetDateTime(2011, 11, 21, 1, 0, 0), GetDateTime(2011, 11, 21, 4, 0, 0), _T(""), _T("RET1"), _T(""));
	ActiveGanttVCCtl1.GetTasks().Add(_T("Task 3"), _T("K3"), GetDateTime(2011, 11, 21, 2, 0, 0), GetDateTime(2011, 11, 21, 5, 0, 0), _T(""), _T("RET1"), _T(""));

	ActiveGanttVCCtl1.Redraw();

	return TRUE;

}

BEGIN_EVENTSINK_MAP(fCustomDrawing, CDialog)
	ON_EVENT(fCustomDrawing, IDC_ACTIVEGANTTVCCTL1, 11, fCustomDrawing::DrawActiveganttvcctl1, VTS_DISPATCH)
END_EVENTSINK_MAP()

void fCustomDrawing::DrawActiveganttvcctl1(LPDISPATCH e)
{
	CDrawEventArgs oE(e);
    if (oE.GetEventTarget() == EVT_TASK)
    {
        if (ActiveGanttVCCtl1.GetSelectedTaskIndex() == oE.GetObjectIndex())
        {
            oE.SetCustomDraw(TRUE);
            CclsTask oTask;
            oTask = ActiveGanttVCCtl1.GetTasks().Item(CStr(oE.GetObjectIndex()));

			CclsFont oFont;
			oFont.CreateDispatch(_T("AGVC.clsFont"));
			oFont.SetFontName(_T("Arial"));
			oFont.SetSize(7);

			//CclsFont oFont = ActiveGanttVCCtl1.GetDrawing().GetFont();
			//oFont.SetFontName(_T("Arial"));
			//oFont.SetSize(7);

			CclsTextFlags oTextFlags = ActiveGanttVCCtl1.GetDrawing().GetTextFlags();
			oTextFlags.SetHorizontalAlignment(HAL_CENTER);
			oTextFlags.SetVerticalAlignment(VAL_CENTER);
			CclsImage oImage = ActiveGanttVCCtl1.GetDrawing().GetImage();
			oImage.FromFile(g_GetAppLocation() + _T("\\Images\\sampleimage.jpg"), TRUE);
            ActiveGanttVCCtl1.GetDrawing().PaintImage(oImage.m_lpDispatch, oTask.GetLeft() + 40, oTask.GetTop() + 10, oTask.GetLeft() + 64, oTask.GetTop() + 34);
            ActiveGanttVCCtl1.GetDrawing().DrawLine(oTask.GetLeft(), CInt32((oTask.GetBottom() - oTask.GetTop()) / 2) + oTask.GetTop(), oTask.GetRight(), CInt32((oTask.GetBottom() - oTask.GetTop()) / 2) + oTask.GetTop(), CLR_GREEN, LDS_SOLID, 1);
            ActiveGanttVCCtl1.GetDrawing().DrawRectangle(oTask.GetLeft(), oTask.GetTop(), oTask.GetLeft() + 10, oTask.GetTop() + 10, CLR_RED, LDS_SOLID, 1);
            ActiveGanttVCCtl1.GetDrawing().DrawBorder(oTask.GetLeft(), oTask.GetTop(), oTask.GetRight(), oTask.GetBottom(), CLR_BLUE, LDS_SOLID, 2);
            ActiveGanttVCCtl1.GetDrawing().DrawAlignedText(oTask.GetLeft(), oTask.GetTop(), oTask.GetRight(), oTask.GetBottom(), oTask.GetText() + _T(" Is Selected"), HAL_RIGHT, VAL_BOTTOM, CLR_BLUE, oFont.m_lpDispatch);
            ActiveGanttVCCtl1.GetDrawing().DrawText(oTask.GetLeft(), oTask.GetTop(), oTask.GetRight(), oTask.GetBottom(), _T("Draw Text"), CLR_RED, oFont.m_lpDispatch);
        }
    }
}

void fCustomDrawing::OnSize(UINT nType, int cx, int cy)
{
	CDialog::OnSize(nType, cx, cy);
	g_Resize(this, &ActiveGanttVCCtl1);
}