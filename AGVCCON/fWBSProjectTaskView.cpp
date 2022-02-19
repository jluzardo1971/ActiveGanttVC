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
#include "fWBSProjectTaskView.h"
#include "fWBSProject.h"

IMPLEMENT_DYNAMIC(fWBSProjectTaskView, CDialog)

fWBSProjectTaskView::fWBSProjectTaskView(CWnd* pParent /*=NULL*/)
	: CDialog(fWBSProjectTaskView::IDD, pParent)
{

}

fWBSProjectTaskView::~fWBSProjectTaskView()
{
}

void fWBSProjectTaskView::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	DDX_Control(pDX, IDC_MSFLEXGRID1, oGrid);
	DDX_Control(pDX, IDC_LBLTASKNAME, lblTaskName);
}


BEGIN_MESSAGE_MAP(fWBSProjectTaskView, CDialog)
END_MESSAGE_MAP()

BOOL fWBSProjectTaskView::OnInitDialog()
{
	CDialog::OnInitDialog();

	int i;
	CclsRow oRow;
	CclsPredecessor oPredecessor;
	CclsTask oTask;
	oRow = mp_oParent->ActiveGanttVCCtl1.GetTasks().Item(CStr(mp_lTaskIndex)).GetRow();
	lblTaskName.SetWindowTextW(oRow.GetText());
	this->SetWindowTextW(GetText(lblTaskName));
	oGrid.SetCols(3);
	oGrid.SetColWidth(0, 800);
	oGrid.SetColWidth(1, 5000);
	oGrid.SetColWidth(2, 2000);
	oGrid.SetFixedCols(0);
	oGrid.SetRow(0);
	oGrid.SetCol(0);
	oGrid.SetText(_T("ID"));
	oGrid.SetCol(1);
	oGrid.SetText(_T("Predecessor Task Name"));
	oGrid.SetCol(2);
	oGrid.SetText(_T("Type"));

	for (i = 1; i <= mp_oParent->ActiveGanttVCCtl1.GetPredecessors().GetCount(); i++)
	{
		oPredecessor = mp_oParent->ActiveGanttVCCtl1.GetPredecessors().Item(CStr(i));
		oTask = mp_oParent->ActiveGanttVCCtl1.GetTasks().Item(CStr(mp_lTaskIndex));
		if (oPredecessor.GetSuccessorKey() == oTask.GetKey())
		{
			oTask = GetPredecessorTask(&oPredecessor);
			oRow = oTask.GetRow();
			oGrid.AddItem(CStr(oTask.GetIndex()) + _T("\t") + oRow.GetText() + _T("\t") + GetPredecessorType(oPredecessor.GetPredecessorType()), CVar(oGrid.GetRows()));
			oGrid.SetRow(oGrid.GetRows() - 1);
			oGrid.SetCol(0);
			oGrid.SetCellAlignment(1);
			oGrid.SetCol(1);
			oGrid.SetCellAlignment(1);
		}
	}

	if (oGrid.GetRows() > 2)
	{
		oGrid.RemoveItem(1);
	}

	return TRUE;
}

CclsTask fWBSProjectTaskView::GetPredecessorTask(CclsPredecessor* oPredecessor)
{

    int i;
	CclsTask oTask;
	for (i = 1; i <= mp_oParent->ActiveGanttVCCtl1.GetTasks().GetCount(); i++)
	{
		if (mp_oParent->ActiveGanttVCCtl1.GetTasks().Item(CStr(i)).GetKey() == oPredecessor->GetPredecessorKey())
		{
			oTask = mp_oParent->ActiveGanttVCCtl1.GetTasks().Item(CStr(i));
			return oTask;
		}
	}
	return NULL;
}

CString fWBSProjectTaskView::GetPredecessorType(LONG lType)
{
	switch (lType)
	{
	case 0:
		return _T("End-To-Start (ES)");
		break;
	case 1:
		return _T("Start-To-Start (SS)");
		break;
	case 2:
		return _T("End-To-End (EE)");
		break;
	case 3:
		return _T("Start-To-End (SE)");
		break;
	}
	return _T("");
}