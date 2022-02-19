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
#include "fWBSPProperties.h"
#include "fWBSProject.h"

IMPLEMENT_DYNAMIC(fWBSPProperties, CDialog)

fWBSPProperties::fWBSPProperties(fWBSProject* oParent)
	: CDialog(fWBSPProperties::IDD, oParent)
{
	mp_oParent = oParent;
}

fWBSPProperties::~fWBSPProperties()
{
}

void fWBSPProperties::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	DDX_Control(pDX, IDC_CHKENFORCEPREDECESSORS, chkEnforcePredecessors);
	DDX_Control(pDX, IDC_CBOPREDECESSORMODE, cboPredecessorMode);
}


BEGIN_MESSAGE_MAP(fWBSPProperties, CDialog)
END_MESSAGE_MAP()

BOOL fWBSPProperties::OnInitDialog()
{
	CDialog::OnInitDialog();
	if (mp_oParent->ActiveGanttVCCtl1.GetEnforcePredecessors() == TRUE)
	{
		chkEnforcePredecessors.SetCheck(BST_CHECKED);
	}
	else
	{
		chkEnforcePredecessors.SetCheck(BST_UNCHECKED);
	}
	cboPredecessorMode.AddString(_T("PM_FORCE"));
	cboPredecessorMode.AddString(_T("PM_CREATEWARNINGFLAG"));
	cboPredecessorMode.AddString(_T("PM_RAISEEVENT"));
	cboPredecessorMode.SetCurSel(mp_oParent->ActiveGanttVCCtl1.GetPredecessorMode());

	return TRUE;
}

void fWBSPProperties::OnOK()
{
	if (chkEnforcePredecessors.GetCheck() == BST_CHECKED)
	{
		mp_oParent->ActiveGanttVCCtl1.SetEnforcePredecessors(TRUE);
	}
	else
	{
		mp_oParent->ActiveGanttVCCtl1.SetEnforcePredecessors(FALSE);
	}
	mp_oParent->ActiveGanttVCCtl1.SetPredecessorMode((E_PREDECESSORMODE)cboPredecessorMode.GetCurSel());

	CDialog::OnOK();
}