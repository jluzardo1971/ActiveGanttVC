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
#include "fCarRentalVehicle.h"
#include "fCarRental.h"

IMPLEMENT_DYNAMIC(fCarRentalVehicle, CDialog)

fCarRentalVehicle::fCarRentalVehicle(CWnd* pParent /*=NULL*/)
	: CDialog(fCarRentalVehicle::IDD, pParent)
{

}

fCarRentalVehicle::~fCarRentalVehicle()
{

}

void fCarRentalVehicle::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	DDX_Control(pDX, IDC_DRPCARTYPEID, drpCarTypeID);
	DDX_Control(pDX, IDC_TXTLICENSEPLATES, txtLicensePlates);
	DDX_Control(pDX, IDC_DRPACRISS1, drpACRISS1);
	DDX_Control(pDX, IDC_DRPACRISS2, drpACRISS2);
	DDX_Control(pDX, IDC_DRPACRISS3, drpACRISS3);
	DDX_Control(pDX, IDC_DRPACRISS4, drpACRISS4);
	DDX_Control(pDX, IDC_TXTRATE, txtRate);
	DDX_Control(pDX, IDC_LBLACRISS1, lblACRISS1);
	DDX_Control(pDX, IDC_LBLACRISS2, lblACRISS2);
	DDX_Control(pDX, IDC_LBLACRISS3, lblACRISS3);
	DDX_Control(pDX, IDC_LBLACRISS4, lblACRISS4);
	DDX_Control(pDX, IDC_TXTNOTES, txtNotes);
	DDX_Control(pDX, IDC_PICMAKE, picMake);
}


BEGIN_MESSAGE_MAP(fCarRentalVehicle, CDialog)
	ON_CBN_SELCHANGE(IDC_DRPCARTYPEID, &fCarRentalVehicle::OnCbnSelchangeDrpcartypeid)
	ON_CBN_SELCHANGE(IDC_DRPACRISS1, &fCarRentalVehicle::OnCbnSelchangeDrpacriss1)
	ON_CBN_SELCHANGE(IDC_DRPACRISS2, &fCarRentalVehicle::OnCbnSelchangeDrpacriss2)
	ON_CBN_SELCHANGE(IDC_DRPACRISS3, &fCarRentalVehicle::OnCbnSelchangeDrpacriss3)
	ON_CBN_SELCHANGE(IDC_DRPACRISS4, &fCarRentalVehicle::OnCbnSelchangeDrpacriss4)
END_MESSAGE_MAP()

BOOL fCarRentalVehicle::OnInitDialog()
{
	CDialog::OnInitDialog();
	CString sACRISSCode;

	
	g_DST_ACCESS_FillComboBox(&drpCarTypeID, _T("SELECT * FROM tb_CR_Car_Types"), _T("lCarTypeID"), _T("sDescription"), mp_oParent->mp_oConn);
	g_DST_ACCESS_FillComboBoxString(&drpACRISS1, _T("SELECT * FROM tb_CR_ACRISS_Codes WHERE [lPosition] = 1"), _T("sLetter"), _T("sDescription"), mp_oParent->mp_oConn);
	g_DST_ACCESS_FillComboBoxString(&drpACRISS2, _T("SELECT * FROM tb_CR_ACRISS_Codes WHERE [lPosition] = 2"), _T("sLetter"), _T("sDescription"), mp_oParent->mp_oConn);
	g_DST_ACCESS_FillComboBoxString(&drpACRISS3, _T("SELECT * FROM tb_CR_ACRISS_Codes WHERE [lPosition] = 3"), _T("sLetter"), _T("sDescription"), mp_oParent->mp_oConn);
	g_DST_ACCESS_FillComboBoxString(&drpACRISS4, _T("SELECT * FROM tb_CR_ACRISS_Codes WHERE [lPosition] = 4"), _T("sLetter"), _T("sDescription"), mp_oParent->mp_oConn);

	if (mp_yDialogMode == DM_ADD)
	{
		this->SetWindowTextW(_T("Add New Vehicle"));
		txtLicensePlates.SetWindowTextW(g_GenerateRandomLicense());
		drpCarTypeID.SetCurSel(GetRnd(0, 46));
		OnCbnSelchangeDrpcartypeid();
	}
	else
	{
		this->SetWindowTextW(_T("Edit Vehicle"));
		
		clsDB oDB(mp_oParent->mp_oConn);
		oDB.InitReader(_T("SELECT * FROM tb_CR_Rows WHERE lRowID = ") + mp_sRowID);
		oDB.ReadTextBox(txtLicensePlates, _T("sLicensePlates"));
		oDB.ReadComboBox(drpCarTypeID, _T("lCarTypeID"));
		oDB.ReadTextBox(txtNotes, _T("sNotes"));
		oDB.ReadTextBox(txtRate, _T("cRate"));
		sACRISSCode = oDB.ReadFieldString(_T("sACRISSCode"));
		oDB.CloseReader();
        UpdatePicture();
        UpdateACRISSCode(sACRISSCode);
	}

	return TRUE;
}

void fCarRentalVehicle::UpdateACRISSCode(CString sACRISSCode)
{
	drpACRISS1.SetCurSel(GetListIndex(&drpACRISS1, sACRISSCode.Mid(0, 1)));
	drpACRISS2.SetCurSel(GetListIndex(&drpACRISS2, sACRISSCode.Mid(1, 1)));
	drpACRISS3.SetCurSel(GetListIndex(&drpACRISS3, sACRISSCode.Mid(2, 1)));
	drpACRISS4.SetCurSel(GetListIndex(&drpACRISS4, sACRISSCode.Mid(3, 1)));
	lblACRISS1.SetWindowTextW(sACRISSCode.Mid(0, 1));
	lblACRISS2.SetWindowTextW(sACRISSCode.Mid(1, 1));
	lblACRISS3.SetWindowTextW(sACRISSCode.Mid(2, 1));
	lblACRISS4.SetWindowTextW(sACRISSCode.Mid(3, 1));
}



void fCarRentalVehicle::UpdatePicture()
{
	if (mp_oImage != NULL)
	{
		mp_oImage.Destroy();
	}
	if (mp_oImage.Load(g_GetAppLocation() + _T("\\CarRental\\Big\\") + GetSelText(drpCarTypeID) + _T(".jpg")) == S_OK)
	{
		picMake.SetBitmap((HBITMAP)mp_oImage);
		picMake.Invalidate();
	}
}

void fCarRentalVehicle::OnCbnSelchangeDrpcartypeid()
{
	CString sACRISSCode;
	UpdatePicture();

	CRecordset oRecordset(&mp_oParent->mp_oConn);
	CString sReturn;
	CString sTest = CStr((int) drpCarTypeID.GetItemData(drpCarTypeID.GetCurSel()));

	oRecordset.Open(CRecordset::forwardOnly, _T("SELECT sACRISSCode, cStdRate FROM tb_CR_Car_Types WHERE lCarTypeID = ") + CStr((int) drpCarTypeID.GetItemData(drpCarTypeID.GetCurSel())), CRecordset::readOnly);
	if (oRecordset.IsEOF() == FALSE)
	{
		double sRate;
		sACRISSCode = GetStrField(oRecordset, _T("sACRISSCode"));
		UpdateACRISSCode(sACRISSCode);
		sRate = GetDblField(oRecordset, _T("cStdRate"));
		txtRate.SetWindowTextW(CStr(sRate));
	}
}

void fCarRentalVehicle::OnOK()
{


    CclsRow oRow;
	
    clsDB oDB(mp_oParent->mp_oConn);
    oDB.AddParameter(_T("lDepth"), 1, PT_NUMERIC);
    oDB.AddParameter(_T("sLicensePlates"), GetText(txtLicensePlates), PT_STRING);
    oDB.AddParameter(_T("lCarTypeID"), (int) drpCarTypeID.GetItemData(drpCarTypeID.GetCurSel()), PT_NUMERIC);
    oDB.AddParameter(_T("sNotes"), GetText(txtNotes), PT_STRING);
    oDB.AddParameter(_T("cRate"), CDbl(GetText(txtRate)), PT_NUMERIC);
    oDB.AddParameter(_T("sACRISSCode"), GetText(lblACRISS1) + GetText(lblACRISS2) + GetText(lblACRISS3) + GetText(lblACRISS4), PT_STRING);
	if (mp_yDialogMode == DM_ADD)
	{
        oDB.AddParameter(_T("lOrder"), mp_oParent->ActiveGanttVCCtl1.GetRows().GetCount() + 1, PT_NUMERIC);
		mp_sRowID = _T("K") + oDB.ExecuteInsert(_T("tb_CR_Rows"));
		oRow = mp_oParent->ActiveGanttVCCtl1.GetRows().Add(mp_sRowID, _T(""), FALSE, TRUE, _T(""));
		oRow.GetNode().SetDepth(1);
		mp_oParent->ActiveGanttVCCtl1.GetRows().UpdateTree();
	}
	else
	{
		oDB.ExecuteUpdate(_T("tb_CR_Rows"), _T("lRowID = ") + mp_sRowID);
		oRow = mp_oParent->ActiveGanttVCCtl1.GetRows().Item(_T("K") + mp_sRowID);
	}
	oRow.GetCells().Item(_T("1")).SetText(GetText(txtLicensePlates));
	oRow.GetCells().Item(_T("2")).GetImage().FromFile(g_GetAppLocation() + _T("\\CarRental\\Small\\") + GetSelText(drpCarTypeID) + _T(".jpg"), TRUE);
	oRow.GetCells().Item(_T("3")).SetText(GetSelText(drpCarTypeID) + _T("\r\n") + GetText(lblACRISS1) + GetText(lblACRISS2) + GetText(lblACRISS3) + GetText(lblACRISS4) + _T(" - ") + GetText(txtRate) + _T(" USD"));
    
    oRow.SetTag( GetText(lblACRISS1) + GetText(lblACRISS2) + GetText(lblACRISS3) + GetText(lblACRISS4) + _T("|") + GetText(txtRate) + _T("|") + CStr((int)drpCarTypeID.GetItemData(drpCarTypeID.GetCurSel())));
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


	DeleteComboBoxItems();
	CDialog::OnOK();
}

void fCarRentalVehicle::OnCancel()
{
	DeleteComboBoxItems();
	CDialog::OnCancel();
}

void fCarRentalVehicle::DeleteComboBoxItems()
{
	int lIndex;
	for (lIndex = 0; lIndex <= (drpACRISS1.GetCount() - 1); lIndex++)
	{
		CString* pLetter = (CString*) drpACRISS1.GetItemDataPtr(lIndex);
		delete pLetter;
	}
	for (lIndex = 0; lIndex <= (drpACRISS2.GetCount() - 1); lIndex++)
	{
		CString* pLetter = (CString*) drpACRISS2.GetItemDataPtr(lIndex);
		delete pLetter;
	}
	for (lIndex = 0; lIndex <= (drpACRISS3.GetCount() - 1); lIndex++)
	{
		CString* pLetter = (CString*) drpACRISS3.GetItemDataPtr(lIndex);
		delete pLetter;
	}
	for (lIndex = 0; lIndex <= (drpACRISS4.GetCount() - 1); lIndex++)
	{
		CString* pLetter = (CString*) drpACRISS4.GetItemDataPtr(lIndex);
		delete pLetter;
	}
}

void fCarRentalVehicle::OnCbnSelchangeDrpacriss1()
{
	lblACRISS1.SetWindowTextW(GetSelText(drpACRISS1).Mid(0,1));
}

void fCarRentalVehicle::OnCbnSelchangeDrpacriss2()
{
	lblACRISS2.SetWindowTextW(GetSelText(drpACRISS2).Mid(0,1));
}

void fCarRentalVehicle::OnCbnSelchangeDrpacriss3()
{
	lblACRISS3.SetWindowTextW(GetSelText(drpACRISS3).Mid(0,1));
}

void fCarRentalVehicle::OnCbnSelchangeDrpacriss4()
{
	lblACRISS4.SetWindowTextW(GetSelText(drpACRISS4).Mid(0,1));
}