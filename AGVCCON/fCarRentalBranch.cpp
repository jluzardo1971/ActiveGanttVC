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
#include "fCarRentalBranch.h"
#include "fCarRental.h"

IMPLEMENT_DYNAMIC(fCarRentalBranch, CDialog)

fCarRentalBranch::fCarRentalBranch(CWnd* pParent /*=NULL*/)
	: CDialog(fCarRentalBranch::IDD, pParent)
{

}

fCarRentalBranch::~fCarRentalBranch()
{
}

void fCarRentalBranch::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	DDX_Control(pDX, IDC_TXTBRANCHNAME, txtBranchName);
	DDX_Control(pDX, IDC_TXTADDRESS, txtAddress);
	DDX_Control(pDX, IDC_TXTCITY, txtCity);
	DDX_Control(pDX, IDC_TXTSTATE, txtState);
	DDX_Control(pDX, IDC_TXTZIP, txtZIP);
	DDX_Control(pDX, IDC_TXTPHONE, txtPhone);
	DDX_Control(pDX, IDC_TXTMANAGERNAME, txtManagerName);
	DDX_Control(pDX, IDC_TXTMANAGERMOBILE, txtManagerMobile);
}


BEGIN_MESSAGE_MAP(fCarRentalBranch, CDialog)

END_MESSAGE_MAP()

BOOL fCarRentalBranch::OnInitDialog()
{
	CDialog::OnInitDialog();

	if (mp_yDialogMode == DM_ADD)
	{
		this->SetWindowTextW(_T("Add New Branch"));
		CString sCityName;
		CString sStateName;
		int lID;
		g_GenerateRandomCity(sCityName, sStateName, lID, mp_oParent->mp_oConn);
		txtCity.SetWindowTextW(sCityName);
        txtBranchName.SetWindowTextW(sCityName);
        txtState.SetWindowTextW(sStateName);
        txtPhone.SetWindowTextW(g_GenerateRandomPhone(_T("")));
        txtManagerName.SetWindowTextW(g_GenerateRandomName(FALSE, mp_oParent->mp_oConn));
        txtManagerMobile.SetWindowTextW(g_GenerateRandomPhone(GetText(txtPhone).Mid(0,5)));
        txtAddress.SetWindowTextW(g_GenerateRandomAddress(mp_oParent->mp_oConn));
        txtZIP.SetWindowTextW(g_GenerateRandomZIP());

	}
	else
	{
		this->SetWindowTextW(_T("Edit Branch"));
		clsDB oDB(mp_oParent->mp_oConn);
		oDB.InitReader(_T("SELECT * FROM tb_CR_Rows WHERE lRowID = ") + mp_sRowID);
        oDB.ReadTextBox(txtCity, _T("sCity"));
        oDB.ReadTextBox(txtBranchName, _T("sBranchName"));
        oDB.ReadTextBox(txtState, _T("sStateAbr"));
        oDB.ReadTextBox(txtPhone, _T("sPhone"));
        oDB.ReadTextBox(txtManagerName, _T("sManagerName"));
        oDB.ReadTextBox(txtManagerMobile, _T("sManagerMobile"));
        oDB.ReadTextBox(txtAddress, _T("sAddress"));
        oDB.ReadTextBox(txtZIP, _T("sZIP"));
        oDB.CloseReader();

	}

	return TRUE;
}



void fCarRentalBranch::OnOK()
{
	CclsRow oRow;
	clsDB oDB(mp_oParent->mp_oConn);
	oDB.AddParameter(_T("lDepth"), 0, PT_NUMERIC);
	oDB.AddParameter(_T("sCity"), GetText(txtCity), PT_STRING);
	oDB.AddParameter(_T("sBranchName"), GetText(txtBranchName), PT_STRING);
	oDB.AddParameter(_T("sStateAbr"), GetText(txtState), PT_STRING);
	oDB.AddParameter(_T("sPhone"), GetText(txtPhone), PT_STRING);
	oDB.AddParameter(_T("sManagerName"), GetText(txtManagerName), PT_STRING);
	oDB.AddParameter(_T("sManagerMobile"), GetText(txtManagerMobile), PT_STRING);
	oDB.AddParameter(_T("sAddress"), GetText(txtAddress), PT_STRING);
	oDB.AddParameter(_T("sZIP"), GetText(txtZIP), PT_STRING);
	if (mp_yDialogMode == DM_ADD)
	{
		oDB.AddParameter(_T("lOrder"), mp_oParent->ActiveGanttVCCtl1.GetRows().GetCount() + 1, PT_NUMERIC);
        mp_sRowID = _T("K") + oDB.ExecuteInsert(_T("tb_CR_Rows"));
        oRow = mp_oParent->ActiveGanttVCCtl1.GetRows().Add(mp_sRowID, _T(""), FALSE, TRUE, _T(""));
		oRow.GetNode().SetDepth(0);
		mp_oParent->ActiveGanttVCCtl1.GetRows().UpdateTree();
	}
	else
	{
        oDB.ExecuteUpdate(_T("tb_CR_Rows"), _T("lRowID = ") + mp_sRowID);
        oRow = mp_oParent->ActiveGanttVCCtl1.GetRows().Item(_T("K") + mp_sRowID);
	}
    oRow.SetText(GetText(txtBranchName) + _T(", ") + GetText(txtState) + _T("\r\n") + _T("Phone: ") + GetText(txtPhone));
    oRow.SetMergeCells(TRUE);
    oRow.SetContainer(FALSE);
    oRow.SetStyleIndex(_T("Branch"));
	oRow.SetClientAreaStyleIndex(_T("BranchCA"));
    oRow.SetUseNodeImages(TRUE);
	oRow.GetNode().GetExpandedImage().FromFile(g_GetAppLocation() + _T("\\CarRental\\minus.jpg"), TRUE);
	oRow.GetNode().GetImage().FromFile(g_GetAppLocation() + _T("\\CarRental\\plus.jpg"), TRUE);
    oRow.SetAllowMove(FALSE);
    oRow.SetAllowSize(FALSE);
	if (mp_yDialogMode == DM_ADD)
	{
		int l;
		l = mp_oParent->ActiveGanttVCCtl1.GetCurrentViewObject().GetClientArea().GetHeight() / 41;
		if ((mp_oParent->ActiveGanttVCCtl1.GetRows().GetCount() - l + 2) > 0)
		{
			mp_oParent->ActiveGanttVCCtl1.GetVerticalScrollBar().SetValue(mp_oParent->ActiveGanttVCCtl1.GetRows().GetCount() - l + 2);
		}
	}
    mp_oParent->ActiveGanttVCCtl1.Redraw();

	CDialog::OnOK();
}