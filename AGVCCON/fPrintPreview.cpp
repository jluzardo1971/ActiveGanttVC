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
#include "fPrintPreview.h"
#include "fPrintDialog.h"

IMPLEMENT_DYNAMIC(fPrintPreview, CDialog)

fPrintPreview::fPrintPreview(CWnd* pParent /*=NULL*/)
	: CDialog(fPrintPreview::IDD, pParent)
{
	mp_lPage = 1;
	mp_lHSB_SmallChange = 25;
	mp_lHSB_LargeChange = 100;
	mp_lVSB_SmallChange = 25;
	mp_lVSB_LargeChange = 100;
	mp_lHScroll_Value = 0;
	mp_lVScroll_Value = 0;
	mp_fScale = 1.0f;
}

fPrintPreview::~fPrintPreview()
{
}

void fPrintPreview::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	DDX_Control(pDX, IDC_HSCROLLBAR, m_HScroll);
	DDX_Control(pDX, IDC_VSCROLLBAR, m_VScroll);
	DDX_Control(pDX, IDC_SCROLLBARSSEP, cmdScrollBarsSeparator);
}

BEGIN_MESSAGE_MAP(fPrintPreview, CDialog)
	ON_WM_SIZE()
	ON_WM_PAINT()
	ON_WM_HSCROLL()
	ON_WM_VSCROLL()
	ON_WM_CREATE()
	ON_WM_ERASEBKGND()
	ON_COMMAND_RANGE(ID_MYTOOLBAR_BUTTONFIRST, ID_MYTOOLBAR_BUTTONLAST, ToolBar_Click)
END_MESSAGE_MAP()

int fPrintPreview::OnCreate(LPCREATESTRUCT lpCreateStruct)
{
	if (CDialog::OnCreate(lpCreateStruct) == -1)
		return -1;

	//Create ToolBar:
	oToolBar1.CreateEx(this, TBSTYLE_FLAT, WS_CHILD | WS_VISIBLE | CBRS_TOP | CBRS_GRIPPER | CBRS_TOOLTIPS | CBRS_FLYBY | CBRS_SIZE_DYNAMIC);
	oToolBarBitmap.LoadBitmapW(IDB_FPRINTPREVIEW2);
	oToolBar1.SetBitmap((HBITMAP)oToolBarBitmap.GetSafeHandle());
	oToolBar1.SetButtons(NULL, 11); //1 Based



	oToolBar1.AddButton(0, _T("Left")); 
	oToolBar1.AddButton(1, _T("Up"));
	oToolBar1.AddButton(2, _T("Down"));
	oToolBar1.AddButton(3, _T("Right"));
	oToolBar1.AddSeparator();
	oToolBar1.AddButton(4, _T("Page Back"));
	oToolBar1.AddButton(5, _T("Page Forward"));
	oToolBar1.AddSeparator(90);
	oToolBar1.AddSeparator();
	oToolBar1.AddButton(6, _T("Zoom In"));
	oToolBar1.AddButton(7, _T("Zoom Out"));


	CRect rect;
	rect.top = 4;
	rect.left = 155;
	rect.right = 600;
	rect.bottom = rect.top + 20;
	lblPages.Create(_T("Page 999 of 999"), SS_SIMPLE | WS_VISIBLE, rect, &oToolBar1);
	mp_oFont.CreatePointFont(8, _T("MS Sans Serif"));
	lblPages.SetFont(&mp_oFont);



	return 0;
}

void fPrintPreview::ToolBar_Click(UINT nID)
{
	UINT lButtonFirst = ID_MYTOOLBAR_BUTTONFIRST;
	nID = nID - lButtonFirst;
	switch (nID)
	{
	case 0: //Left
		{
			mp_oParent->mp_oControl->GetPrinter().GetPagePosition(mp_lPage, &mp_lColumn, &mp_lRow);
			if (mp_lColumn > 1)
			{
				mp_lColumn = mp_lColumn - 1;
				mp_lPage = mp_oParent->mp_oControl->GetPrinter().GetPageNumber(mp_lColumn, mp_lRow);
				mp_UpdatePageNumber();
				this->Invalidate();

			}
		}
		break;
	case 1: //Up
		{
			mp_oParent->mp_oControl->GetPrinter().GetPagePosition(mp_lPage, &mp_lColumn, &mp_lRow);
			if (mp_lRow > 1)
			{
				mp_lRow = mp_lRow - 1;
				mp_lPage = mp_oParent->mp_oControl->GetPrinter().GetPageNumber(mp_lColumn, mp_lRow);
				mp_UpdatePageNumber();
				this->Invalidate();
			}
		}
		break;
	case 2: //Down
		{
			mp_oParent->mp_oControl->GetPrinter().GetPagePosition(mp_lPage, &mp_lColumn, &mp_lRow);
			if (mp_lRow < mp_oParent->mp_oControl->GetPrinter().GetYAxisPages())
			{
				mp_lRow = mp_lRow + 1;
				mp_lPage = mp_oParent->mp_oControl->GetPrinter().GetPageNumber(mp_lColumn, mp_lRow);
				mp_UpdatePageNumber();
				this->Invalidate();
			}
		}
		break;
	case 3: //Right
		{
			mp_oParent->mp_oControl->GetPrinter().GetPagePosition(mp_lPage, &mp_lColumn, &mp_lRow);
			if (mp_lColumn < mp_oParent->mp_oControl->GetPrinter().GetXAxisPages())
			{
				mp_lColumn = mp_lColumn + 1;
				mp_lPage = mp_oParent->mp_oControl->GetPrinter().GetPageNumber(mp_lColumn, mp_lRow);
				mp_UpdatePageNumber();
				this->Invalidate();
			}
		}
		break;
	case 5: //Page Back
		{
			if (mp_lPage > 1)
			{
				mp_lPage = mp_lPage - 1;
				mp_UpdatePageNumber();
				this->Invalidate();
			}
		}
		break;
	case 6: //Page Forward
		{
			if (mp_lPage < mp_oParent->mp_oControl->GetPrinter().GetPages())
			{
				mp_lPage = mp_lPage + 1;
				mp_UpdatePageNumber();
				this->Invalidate();
			}
		}
		break;
	case 9: //Zoom In
		{
			if (mp_fScale < 2.0f)
			{
				mp_fScale = mp_fScale + 0.1f;
				this->Invalidate();
			}
		}
		break;
	case 10: //Zoom out
		{
			if (mp_fScale > 0.1f)
			{
				mp_fScale = mp_fScale - 0.1f;
				this->Invalidate();
			}
		}
		break;
	}

}

BOOL fPrintPreview::OnInitDialog()
{
	CDialog::OnInitDialog();
	g_MaximizeWindowsDim(this);
	CWnd::ShowWindow(SW_SHOWMAXIMIZED);

	// Reposition the Toolbar
	CWnd::RepositionBars(AFX_IDW_CONTROLBAR_FIRST, AFX_IDW_CONTROLBAR_LAST,0);

	m_HScroll.SetScrollRange(0, 1200, TRUE);
	m_HScroll.EnableScrollBar(ESB_ENABLE_BOTH);
	
	m_VScroll.SetScrollRange(0, 1200, TRUE);
	m_VScroll.EnableScrollBar(ESB_ENABLE_BOTH);


	mp_UpdatePageNumber();



	return TRUE; 
}

void fPrintPreview::OnSize(UINT nType, int cx, int cy)
{
	CDialog::OnSize(nType, cx, cy);

	if (IsIconic() == FALSE)
	{
		if (IsWindow(m_HScroll.m_hWnd) == TRUE && IsWindow(m_VScroll.m_hWnd) == TRUE)
		{
			PositionScrollBars();
			Invalidate();
		}
	}
}

void fPrintPreview::PositionScrollBars()
{
	RECT oClientRect;
	RECT oToolbarRect;

	GetClientRect(&oClientRect);
	oToolBar1.GetWindowRect(&oToolbarRect);
	int lDialogWidth = (oClientRect.right - oClientRect.left);
	int lDialogHeight = (oClientRect.bottom - oClientRect.top);
	int lToolbarHeight = (oToolbarRect.bottom);

	// Position Scroll Bars
	m_VScroll.MoveWindow(lDialogWidth - 17, lToolbarHeight, 17, lDialogHeight - lToolbarHeight - 17);
	m_HScroll.MoveWindow(0, lDialogHeight - 17, lDialogWidth - 17, 17);
	cmdScrollBarsSeparator.MoveWindow(lDialogWidth - 17, lDialogHeight - 17, 17, 17);
}

void fPrintPreview::OnPaint()
{
	CPaintDC dc(this);

	RECT oClientRect;
	RECT oToolbarRect;

	GetClientRect(&oClientRect);
	oToolBar1.GetWindowRect(&oToolbarRect);
	int lDialogWidth = (oClientRect.right - oClientRect.left);
	int lDialogHeight = (oClientRect.bottom - oClientRect.top);


	CDC pBuff;
	CBitmap oBuff;
	CBitmap* oldBitmap;
	pBuff.CreateCompatibleDC(&dc);
	oBuff.CreateCompatibleBitmap(&dc, lDialogWidth, lDialogHeight);
	oldBitmap = pBuff.SelectObject(&oBuff);

	pBuff.FillSolidRect(0,0,lDialogWidth,lDialogHeight, RGB(132,130,132));
	
	mp_oParent->mp_oControl->GetPrinter().PreviewPage((LONG)pBuff.m_hDC, mp_lPage, mp_fScale, 100 - mp_lHScroll_Value, 100 - mp_lVScroll_Value);

	dc.BitBlt(0, 0, (oClientRect.right - oClientRect.left),  (oClientRect.bottom - oClientRect.top), &pBuff, 0, 0, SRCCOPY);
	pBuff.SelectObject(oldBitmap);
	pBuff.DeleteDC();
}

void fPrintPreview::mp_UpdatePageNumber()
{
	lblPages.SetWindowTextW(_T("Page ") + CStr(mp_lPage) + _T(" of ") + CStr(mp_oParent->mp_oControl->GetPrinter().GetPages()));
	oToolBar1.Invalidate();
}

void fPrintPreview::OnHScroll(UINT nSBCode, UINT nPos, CScrollBar* pScrollBar)
{
	switch (nSBCode)
	{
	case SB_LINEUP:
		m_HScroll.SetScrollPos(m_HScroll.GetScrollPos() - mp_lHSB_SmallChange);
		break;
	case SB_PAGEUP:
		m_HScroll.SetScrollPos(m_HScroll.GetScrollPos() - mp_lHSB_LargeChange);
		break;
	case SB_PAGEDOWN:
		m_HScroll.SetScrollPos(m_HScroll.GetScrollPos() + mp_lHSB_LargeChange);
		break;
	case SB_LINEDOWN:
		m_HScroll.SetScrollPos(m_HScroll.GetScrollPos() + mp_lHSB_SmallChange);
		break;
	case SB_THUMBTRACK:
		m_HScroll.SetScrollPos(nPos);
		break;
	}
	mp_lHScroll_Value = m_HScroll.GetScrollPos();
	this->Invalidate();

	CDialog::OnHScroll(nSBCode, nPos, pScrollBar);
}

void fPrintPreview::OnVScroll(UINT nSBCode, UINT nPos, CScrollBar* pScrollBar)
{
	switch (nSBCode)
	{
	case SB_LINEUP:
		m_VScroll.SetScrollPos(m_VScroll.GetScrollPos() - mp_lVSB_SmallChange);
		break;
	case SB_PAGEUP:
		m_VScroll.SetScrollPos(m_VScroll.GetScrollPos() - mp_lVSB_LargeChange);
		break;
	case SB_PAGEDOWN:
		m_VScroll.SetScrollPos(m_VScroll.GetScrollPos() + mp_lVSB_LargeChange);
		break;
	case SB_LINEDOWN:
		m_VScroll.SetScrollPos(m_VScroll.GetScrollPos() + mp_lVSB_SmallChange);
		break;
	case SB_THUMBTRACK:
		m_VScroll.SetScrollPos(nPos);
		break;
	}
	mp_lVScroll_Value = m_VScroll.GetScrollPos();
	this->Invalidate();

	CDialog::OnVScroll(nSBCode, nPos, pScrollBar);
}

BOOL fPrintPreview::OnEraseBkgnd(CDC* pDC)
{
	// Prevent flickering
	return TRUE;
}