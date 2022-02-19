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
#include "ToolTipToolBar.h"

IMPLEMENT_DYNAMIC(CToolTipToolBar, CToolBar)

CToolTipToolBar::CToolTipToolBar()
{
	mp_lCounter = 0;
}

CToolTipToolBar::~CToolTipToolBar()
{
	m_asToolTips.clear();
}


BEGIN_MESSAGE_MAP(CToolTipToolBar, CToolBar)
END_MESSAGE_MAP()

BOOL CToolTipToolBar::OnNotify(WPARAM wParam, LPARAM lParam, LRESULT* pResult)
{
    LPNMHDR pNMHDR = (LPNMHDR) lParam;
    LPTOOLTIPTEXT oToolTipText = (LPTOOLTIPTEXT) pNMHDR;
    if(pNMHDR->code == TTN_NEEDTEXTW) 
    {  
        UINT nID = (UINT) pNMHDR->idFrom;
        if(!(oToolTipText->uFlags & TTF_IDISHWND))
        {
            if ((nID >= ID_MYTOOLBAR_BUTTONFIRST) || (nID <= ID_MYTOOLBAR_BUTTONLAST))
            {
				CString sToolTip = GetToolTip(nID);
				wcscpy_s(oToolTipText->lpszText, sToolTip.GetLength() + 1, sToolTip);
                return TRUE;
            }
        }
    }

	return CToolBar::OnNotify(wParam, lParam, pResult);
}

void CToolTipToolBar::InsertToolTip(CString sToolTip)
{
    m_asToolTips.push_back(sToolTip); //Adds an element to the end of the vector
}

CString CToolTipToolBar::GetToolTip(int nID)
{
	UINT btn1 = ID_MYTOOLBAR_BUTTONFIRST;
    return m_asToolTips[nID - btn1];
}

void CToolTipToolBar::AddButton(int iImage, CString sToolTipText)
{
	InsertToolTip(sToolTipText);
	SetButtonInfo(mp_lCounter, ID_MYTOOLBAR_BUTTONFIRST + mp_lCounter, TBBS_BUTTON | TBBS_AUTOSIZE, iImage);
	mp_lCounter = mp_lCounter + 1;
}

void CToolTipToolBar::AddSeparator()
{
	InsertToolTip(_T(""));
	SetButtonInfo(mp_lCounter, ID_MYTOOLBAR_BUTTONFIRST + mp_lCounter, TBBS_SEPARATOR, 0);
	mp_lCounter = mp_lCounter + 1;
}

void CToolTipToolBar::AddSeparator(int lWidth)
{
	InsertToolTip(_T(""));
	SetButtonInfo(mp_lCounter, ID_MYTOOLBAR_BUTTONFIRST + mp_lCounter, TBBS_SEPARATOR, lWidth);
	mp_lCounter = mp_lCounter + 1;
}

void CToolTipToolBar::AddLabel(int lWidth)
{
	InsertToolTip(_T(""));
	SetButtonInfo(mp_lCounter, ID_MYTOOLBAR_BUTTONFIRST + mp_lCounter, TBSTYLE_AUTOSIZE, lWidth);
	SetButtonText(mp_lCounter, _T("Hello"));
	mp_lCounter = mp_lCounter + 1;
}


int CToolTipToolBar::GetItemLeft(int lIndex)
{
	RECT oRect;
	GetItemRect(lIndex, &oRect);
	return oRect.left;
}

int CToolTipToolBar::GetItemWidth(int lIndex)
{
	RECT oRect;
	GetItemRect(lIndex, &oRect);
	return oRect.right - oRect.left;
}