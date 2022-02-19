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
#pragma once
#include "afxcmn.h"

typedef CMap<CString, LPCTSTR, HTREEITEM, HTREEITEM> t_TreeKeys;

class fMain : public CDialog
{
	DECLARE_DYNAMIC(fMain)

public:
	fMain(CWnd* pParent = NULL);   
	virtual ~fMain();
	HICON m_hIcon;


	enum { IDD = IDD_FMAIN };

protected:
	virtual void DoDataExchange(CDataExchange* pDX);    

	DECLARE_MESSAGE_MAP()
public:
	virtual BOOL OnInitDialog();
	afx_msg void OnPaint();
	afx_msg HCURSOR OnQueryDragIcon();
	afx_msg HBRUSH OnCtlColor(CDC* pDC, CWnd* pWnd, UINT nCtlColor);
	CTreeCtrl TreeView1;
	void AddTitleNode(CString sKey, CString sText, long ImageIndex, long SelectedImageIndex);
	void AddNode(CString sParentKey, CString sKey, CString sText, long ImageIndex, long SelectedImageIndex);
	t_TreeKeys oKeys;
	void ExpandAll();
	void Collapse(CString sKey);
	CImageList ImageList1;
	afx_msg void OnNMDblclkTree1(NMHDR *pNMHDR, LRESULT *pResult);
	CBrush mp_olblCopyrightBrush;
};