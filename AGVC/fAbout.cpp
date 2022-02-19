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
#include "ActiveGanttVC.h"
#include "fAbout.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

/////////////////////////////////////////////////////////////////////////////
// fAbout dialog


fAbout::fAbout(LPCTSTR sVersion, CWnd* pParent /*=NULL*/)
	: CDialog(fAbout::IDD, pParent)
{
	m_sVersion = _T("ActiveGanttVC Scheduler Component, Version ");
	m_sVersion += sVersion;
	//{{AFX_DATA_INIT(fAbout)
		// NOTE: the ClassWizard will add member initialization here
	//}}AFX_DATA_INIT
}


void fAbout::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	//{{AFX_DATA_MAP(fAbout)
		// NOTE: the ClassWizard will add DDX and DDV calls here
	//}}AFX_DATA_MAP
}


BEGIN_MESSAGE_MAP(fAbout, CDialog)
	//{{AFX_MSG_MAP(fAbout)
	ON_WM_LBUTTONDOWN()
	ON_WM_MOUSEMOVE()
	ON_WM_PAINT()
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()

void fAbout::OnLButtonDown(UINT nFlags, CPoint point) 
{
	if ((point.x >= mp_udtURL.left) && (point.x <= mp_udtURL.right) && (point.y >= mp_udtURL.top) && (point.y <= mp_udtURL.bottom) )
	{
		ShellExecute(0, _T("open"), _T("http://www.sourcecodestore.com"), 0, 0, SW_SHOWNORMAL);
	}
	else if ((point.x >= mp_udtStore.left) && (point.x <= mp_udtStore.right) && (point.y >= mp_udtStore.top) && (point.y <= mp_udtStore.bottom) )
	{
		ShellExecute(0, _T("open"), _T("http://www.sourcecodestore.com/OnlineStore/default.aspx"), 0, 0, SW_SHOWNORMAL);
	}
	else if ((point.x >= mp_udtSales.left) && (point.x <= mp_udtSales.right) && (point.y >= mp_udtSales.top) && (point.y <= mp_udtSales.bottom) )
	{
		ShellExecute(0, _T("open"), _T("mailto:sales@sourcecodestore.com"), 0, 0, SW_SHOWNORMAL);
	}
	else if ((point.x >= mp_udtSupport.left) && (point.x <= mp_udtSupport.right) && (point.y >= mp_udtSupport.top) && (point.y <= mp_udtSupport.bottom) )
	{
		ShellExecute(0, _T("open"), _T("http://www.sourcecodestore.com/Support/default.aspx"), 0, 0, SW_SHOWNORMAL);
	}	
	
	CDialog::OnLButtonDown(nFlags, point);
}

BOOL fAbout::OnInitDialog() 
{
	CDialog::OnInitDialog();
//	mp_hBrowserCursor = AfxGetApp()->LoadCursor(IDC_WEBBROWSER);
	m_hIcon = AfxGetApp()->LoadIcon(IDI_ABOUT);

	SetIcon(m_hIcon, TRUE);
	SetIcon(m_hIcon, FALSE);
	
	return TRUE;  // return TRUE unless you set the focus to a control
	              // EXCEPTION: OCX Property Pages should return FALSE
}

void fAbout::OnMouseMove(UINT nFlags, CPoint point) 
{
	if ((point.x >= mp_udtURL.left) && (point.x <= mp_udtURL.right) && (point.y >= mp_udtURL.top) && (point.y <= mp_udtURL.bottom) )
	{
		::SetCursor(mp_hBrowserCursor);
	}
	else if ((point.x >= mp_udtStore.left) && (point.x <= mp_udtStore.right) && (point.y >= mp_udtStore.top) && (point.y <= mp_udtStore.bottom) )
	{
		::SetCursor(mp_hBrowserCursor);
	}
	else if ((point.x >= mp_udtSales.left) && (point.x <= mp_udtSales.right) && (point.y >= mp_udtSales.top) && (point.y <= mp_udtSales.bottom) )
	{
		::SetCursor(mp_hBrowserCursor);
	}
	else if ((point.x >= mp_udtSupport.left) && (point.x <= mp_udtSupport.right) && (point.y >= mp_udtSupport.top) && (point.y <= mp_udtSupport.bottom) )
	{
		::SetCursor(mp_hBrowserCursor);
	}		
	
	CDialog::OnMouseMove(nFlags, point);
}

void fAbout::OnPaint() 
{
	CPaintDC dc(this); // device context for painting
	CFont oFontBold;
	CFont oFontBoldUnder;
	CFont oFont;
	RECT udtTitle;
	CFont* hHoldFont;


	oFontBold.CreateFont(12,0,0,0, FW_BOLD, FALSE, FALSE, 0, ANSI_CHARSET, OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS, DEFAULT_QUALITY, DEFAULT_PITCH | FF_SWISS, _T("MS Sans Serif"));
	oFontBoldUnder.CreateFont(12,0,0,0, FW_BOLD, FALSE, TRUE, 0, ANSI_CHARSET, OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS, DEFAULT_QUALITY, DEFAULT_PITCH | FF_SWISS, _T("MS Sans Serif"));
	oFont.CreateFont(12,0,0,0, FW_NORMAL, FALSE, TRUE, 0, ANSI_CHARSET, OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS, DEFAULT_QUALITY, DEFAULT_PITCH | FF_SWISS, _T("MS Sans Serif"));
	dc.SetBkMode(TRANSPARENT);
	hHoldFont = dc.SelectObject(&oFontBold);
	
	udtTitle.left = 54;
	udtTitle.top = 25;
	udtTitle.right = 0;
	udtTitle.bottom = 0;
	
	dc.DrawText(m_sVersion, m_sVersion.GetLength(), &udtTitle, DT_CALCRECT);
	dc.DrawText(m_sVersion, m_sVersion.GetLength(), &udtTitle, 0);

	dc.SelectObject(&oFont);
	dc.SetTextColor(RGB(0,0,255));

	mp_udtURL.left = 103;
	mp_udtURL.top = 80;
	mp_udtURL.right = 0;
	mp_udtURL.bottom = 0;
	
	dc.DrawText(_T("http://www.sourcecodestore.com"),30, &mp_udtURL, DT_CALCRECT);
	dc.DrawText(_T("http://www.sourcecodestore.com"),30, &mp_udtURL, 0);
	

	dc.SelectObject(&oFontBoldUnder);

	mp_udtStore.left = 135;
	mp_udtStore.top = 122;
	mp_udtStore.right = 0;
	mp_udtStore.bottom = 0;
	
	dc.DrawText(_T("secure online order form"),24, &mp_udtStore, DT_CALCRECT);
	dc.DrawText(_T("secure online order form"),24, &mp_udtStore, 0);

	dc.SelectObject(&oFont);

	mp_udtSupport.left = 135;
	mp_udtSupport.top = 138;
	mp_udtSupport.right = 0;
	mp_udtSupport.bottom = 0;
	
	dc.DrawText(_T("Technical Support Page"),22, &mp_udtSupport, DT_CALCRECT);
	dc.DrawText(_T("Technical Support Page"),22, &mp_udtSupport, 0);

	mp_udtSales.left = 135;
	mp_udtSales.top = 154;
	mp_udtSales.right = 0;
	mp_udtSales.bottom = 0;
	
	dc.DrawText(_T("sales@sourcecodestore.com"),25, &mp_udtSales, DT_CALCRECT);
	dc.DrawText(_T("sales@sourcecodestore.com"),25, &mp_udtSales, 0);


	dc.SelectObject(hHoldFont);
	oFontBold.DeleteObject();
	oFontBoldUnder.DeleteObject();
	oFont.DeleteObject();
}