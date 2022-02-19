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
#include "fCarRentalReservation.h"
#include "fCarRental.h"


IMPLEMENT_DYNAMIC(fCarRentalReservation, CDialog)

void fCarRentalReservation::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	DDX_Control(pDX, IDC_LBLMODE, lblMode);
	DDX_Control(pDX, IDC_TXTNAME, txtName);
	DDX_Control(pDX, IDC_TXTADDRESS, txtAddress);
	DDX_Control(pDX, IDC_TXTCITY, txtCity);
	DDX_Control(pDX, IDC_TXTSTATE, txtState);
	DDX_Control(pDX, IDC_TXTZIP, txtZIP);
	DDX_Control(pDX, IDC_TXTPHONE, txtPhone);
	DDX_Control(pDX, IDC_TXTMOBILE, txtMobile);
	DDX_Control(pDX, IDC_TXTPICKUP, txtPickup);
	DDX_Control(pDX, IDC_DTPPICKUP, dtpPickup);
	DDX_Control(pDX, IDC_TXTRETURN, txtReturn);
	DDX_Control(pDX, IDC_DTPRETURN, dtpReturn);
	DDX_Control(pDX, IDC_TXTRATE, txtRate);
	DDX_Control(pDX, IDC_TXTDESCRIPTION, txtDescription);
	DDX_Control(pDX, IDC_TXTACRISS1, txtACRISS1);
	DDX_Control(pDX, IDC_TXTACRISS2, txtACRISS2);
	DDX_Control(pDX, IDC_TXTACRISS3, txtACRISS3);
	DDX_Control(pDX, IDC_TXTACRISS4, txtACRISS4);
	DDX_Control(pDX, IDC_CHKGPS, chkGPS);
	DDX_Control(pDX, IDC_CHKFSO, chkFSO);
	DDX_Control(pDX, IDC_CHKLDW, chkLDW);
	DDX_Control(pDX, IDC_CHKPAI, chkPAI);
	DDX_Control(pDX, IDC_CHKPEP, chkPEP);
	DDX_Control(pDX, IDC_CHKALI, chkALI);
	DDX_Control(pDX, IDC_TXTGPS, txtGPS);
	DDX_Control(pDX, IDC_TXTLDW, txtLDW);
	DDX_Control(pDX, IDC_TXTPAI, txtPAI);
	DDX_Control(pDX, IDC_TXTPEP, txtPEP);
	DDX_Control(pDX, IDC_TXTALI, txtALI);
	DDX_Control(pDX, IDC_TXTCRF, txtCRF);
	DDX_Control(pDX, IDC_TXTRCFC, txtRCFC);
	DDX_Control(pDX, IDC_TXTERF, txtERF);
	DDX_Control(pDX, IDC_TXTVLF, txtVLF);
	DDX_Control(pDX, IDC_TXTWTB, txtWTB);
	DDX_Control(pDX, IDC_TXTSURCHARGE, txtSurcharge);
	DDX_Control(pDX, IDC_TXTTAX, txtTax);
	DDX_Control(pDX, IDC_TXTESTIMATEDTOTAL, txtEstimatedTotal);
	DDX_Control(pDX, IDC_LBLTAX, lblTax);
	DDX_Control(pDX, IDC_PICTURE, mp_oPicture);
}

BEGIN_MESSAGE_MAP(fCarRentalReservation, CDialog)
	ON_WM_CTLCOLOR()
	ON_BN_CLICKED(IDC_CHKALI, &fCarRentalReservation::OnBnClickedChkali)
	ON_BN_CLICKED(IDC_CHKGPS, &fCarRentalReservation::OnBnClickedChkgps)
	ON_BN_CLICKED(IDC_CHKLDW, &fCarRentalReservation::OnBnClickedChkldw)
	ON_BN_CLICKED(IDC_CHKPAI, &fCarRentalReservation::OnBnClickedChkpai)
	ON_BN_CLICKED(IDC_CHKPEP, &fCarRentalReservation::OnBnClickedChkpep)
END_MESSAGE_MAP()

fCarRentalReservation::fCarRentalReservation(PRG_DIALOGMODE yDialogMode, CWnd* pParent, CString sTaskID) : CDialog(fCarRentalReservation::IDD, pParent)
{
	mp_yDialogMode = yDialogMode;
	mp_oParent = (fCarRental*)pParent;
	mp_sTaskID = sTaskID;
}

fCarRentalReservation::~fCarRentalReservation()
{

}

BOOL fCarRentalReservation::OnInitDialog()
{
	CDialog::OnInitDialog();

	CclsTask oTask = GetCurrentTask();
	CString* sRowTag;

	if (mp_yDialogMode == DM_ADD)
	{
		CString sCityName;
		CString sStateName;
		int lID;
		if (mp_oParent->GetMode() == AM_RESERVATION)
		{
			this->SetWindowTextW(_T("Add Reservation"));
	        lblMode.SetWindowTextW(_T("Add Reservation"));
	        lblModeBackColor = RGB(153, 170, 194);
		}
		else
		{
			this->SetWindowTextW(_T("Add Rental"));
            lblMode.SetWindowTextW(_T("Add Reservation"));
            lblModeBackColor = RGB(162, 78, 50);
		}
		mp_olblModeBrush.CreateSolidBrush(lblModeBackColor);

        g_GenerateRandomCity(sCityName, sStateName, lID, mp_oParent->mp_oConn);
		oTask = mp_oParent->ActiveGanttVCCtl1.GetTasks().Item(CStr(mp_oParent->ActiveGanttVCCtl1.GetTasks().GetCount()));
        txtCity.SetWindowTextW(sCityName);
        txtName.SetWindowTextW(g_GenerateRandomName(FALSE, mp_oParent->mp_oConn));
        txtState.SetWindowTextW(sStateName);
        txtPhone.SetWindowTextW(g_GenerateRandomPhone(_T("")));
        txtMobile.SetWindowTextW(g_GenerateRandomPhone(GetText(txtPhone).Mid(0,5)));
        txtAddress.SetWindowTextW(g_GenerateRandomAddress(mp_oParent->mp_oConn));
        txtZIP.SetWindowTextW(g_GenerateRandomZIP());
        dtpPickup.SetTime(oTask.GetStartDate());
        dtpReturn.SetTime(oTask.GetEndDate());

		CRecordset oRecordset(&mp_oParent->mp_oConn);
		oRecordset.Open(CRecordset::forwardOnly, _T("SELECT * FROM tb_CR_Taxes_Surcharges_Options"), CRecordset::readOnly);
		while(oRecordset.IsEOF() == FALSE)
		{
			CString sID, sDescription;
			DOUBLE cRate;

			sID = GetStrField(oRecordset, _T("sID"));
			sDescription = GetStrField(oRecordset, _T("sDescription"));
			cRate = GetDblField(oRecordset, _T("cRate"));


			if (sID == _T("GPS"))
			{
				chkGPS.SetWindowTextW(sDescription);
				chkGPS_Tag = cRate;
			}
			if (sID == _T("LDW"))
			{
          
				chkLDW.SetWindowTextW(sDescription);
	            chkLDW_Tag = cRate;
			}
			if (sID == _T("PAI"))
			{
				chkPAI.SetWindowTextW(sDescription);
	            chkPAI_Tag = cRate;
			}
			if (sID == _T("PEP"))
			{
				chkPEP.SetWindowTextW(sDescription);
	            chkPEP_Tag = cRate;
			}
			if (sID == _T("ALI"))
			{
				chkALI.SetWindowTextW(sDescription);
	            chkALI_Tag = cRate;
			}
			if (sID == _T("ERF"))
			{
	            txtERF_Tag = cRate;
			}
			if (sID == _T("CRF"))
			{
	            txtCRF_Tag = cRate;
			}
			if (sID == _T("RCFC"))
			{
	            txtRCFC_Tag = cRate;
			}
			if (sID == _T("WTB"))
			{
	            txtWTB_Tag = cRate;
			}
			if (sID == _T("VLF"))
			{
	            txtVLF_Tag = cRate;
			}
			oRecordset.MoveNext();
		}
		oRecordset.Close();
	}
	else
	{
		oTask = mp_oParent->ActiveGanttVCCtl1.GetTasks().Item(_T("K") + mp_sTaskID);
		if (oTask.GetTag() == _T("AM_RESERVATION"))
		{
			this->SetWindowTextW(_T("Edit Reservation"));
			lblMode.SetWindowTextW(_T("Edit Reservation"));
			lblModeBackColor = RGB(153, 170, 194);
		}
		else if (oTask.GetTag() == _T("AM_RENTAL"))
		{
			this->SetWindowTextW(_T("Edit Rental"));
			lblMode.SetWindowTextW(_T("Edit Rental"));
			lblModeBackColor = RGB(162, 78, 50);
		}
		mp_olblModeBrush.CreateSolidBrush(lblModeBackColor);

        clsDB oDB(mp_oParent->mp_oConn);
        oDB.InitReader(_T("SELECT * FROM tb_CR_Rentals WHERE lTaskID = ") + mp_sTaskID);
        oDB.ReadTextBox(txtCity, _T("sCity"));
        oDB.ReadTextBox(txtName, _T("sCustomerName"));
        oDB.ReadTextBox(txtState, _T("sStateAbr"));
        oDB.ReadTextBox(txtPhone, _T("sPhone"));
        oDB.ReadTextBox(txtMobile, _T("sMobile"));
        oDB.ReadTextBox(txtAddress, _T("sAddress"));
        oDB.ReadTextBox(txtZIP, _T("sZIP"));
		dtpPickup.SetTime(oDB.ReadFieldDate(_T("dtPickUp")));
        dtpReturn.SetTime(oDB.ReadFieldDate(_T("dtReturn")));
        oDB.ReadTextBox(txtRate, _T("cRate"));
		chkGPS_Tag = oDB.ReadFieldDouble(_T("cGPS"));
        chkLDW_Tag = oDB.ReadFieldDouble(_T("cLDW"));
        chkPAI_Tag = oDB.ReadFieldDouble(_T("cPAI"));
        chkPEP_Tag = oDB.ReadFieldDouble(_T("cPEP"));
        chkALI_Tag = oDB.ReadFieldDouble(_T("cALI"));
        txtERF_Tag = oDB.ReadFieldDouble(_T("cERF"));
        txtCRF_Tag = oDB.ReadFieldDouble(_T("cCRF"));
        txtRCFC_Tag = oDB.ReadFieldDouble(_T("cRCFC"));
        txtWTB_Tag = oDB.ReadFieldDouble(_T("cWTB"));
        txtVLF_Tag = oDB.ReadFieldDouble(_T("cVLF"));
        lblTax_Tag = oDB.ReadFieldDouble(_T("cTax"));
		txtEstimatedTotal_Tag = oDB.ReadFieldDouble(_T("cEstimatedTotal"));
        oDB.ReadCheckBox(chkGPS, _T("bGPS"));
        oDB.ReadCheckBox(chkFSO, _T("bFSO"));
        oDB.ReadCheckBox(chkLDW, _T("bLDW"));
        oDB.ReadCheckBox(chkPAI, _T("bPAI"));
        oDB.ReadCheckBox(chkPEP, _T("bPEP"));
        oDB.ReadCheckBox(chkALI, _T("bALI"));
        oDB.CloseReader();
	}
    GetStateTax();
    txtPickup.SetWindowTextW(FormatDate(oTask.GetStartDate(), _T("%#c")));
    txtReturn.SetWindowTextW(FormatDate(oTask.GetEndDate(), _T("%#c")));
    sRowTag = Split(oTask.GetRow().GetTag(), _T("|"));
    txtDescription.SetWindowTextW(mp_oParent->GetDescription(CInt32(sRowTag[2])));
	if (mp_oImage.Load(g_GetAppLocation() + _T("\\CarRental\\Small\\") + mp_oParent->GetDescription(CInt32(sRowTag[2])) + _T(".jpg")) == S_OK)
	{
		mp_oPicture.SetBitmap((HBITMAP)mp_oImage);
		mp_oPicture.Invalidate();
	}
    txtRate.SetWindowTextW(sRowTag[1]);
    txtACRISS1.SetWindowTextW(GetACRISSDescription(1, sRowTag[0].Mid(0, 1)));
	txtACRISS2.SetWindowTextW(GetACRISSDescription(2, sRowTag[0].Mid(1, 1)));
	txtACRISS3.SetWindowTextW(GetACRISSDescription(3, sRowTag[0].Mid(2, 1)));
	txtACRISS4.SetWindowTextW(GetACRISSDescription(4, sRowTag[0].Mid(3, 1)));
	delete [] sRowTag;
    CalculateRate();
	
	return TRUE;
}

CclsTask fCarRentalReservation::GetCurrentTask()
{
	CclsTask oReturnTask;
	if (mp_yDialogMode == DM_EDIT)
	{
		oReturnTask = mp_oParent->ActiveGanttVCCtl1.GetTasks().Item(_T("K") + mp_sTaskID);
	}
	else if (mp_yDialogMode == DM_ADD)
	{
		oReturnTask = mp_oParent->ActiveGanttVCCtl1.GetTasks().Item(CStr(mp_oParent->ActiveGanttVCCtl1.GetTasks().GetCount()));
	}
	return oReturnTask;
}

void fCarRentalReservation::GetStateTax()
{
	CclsTask oTask = GetCurrentTask();
    CString sState;
    DOUBLE cTax;
    cTax = mp_oParent->GetStateTax(&oTask, &sState);
    lblTax.SetWindowTextW(sState + _T(" State Tax (") + CStr(cTax * 100, 0) + _T("%):"));
    lblTax_Tag = cTax;
}

void fCarRentalReservation::CalculateRate()
{
	CclsTask oTask = GetCurrentTask();
	 
	DOUBLE cFactor;
    CString* sRowTag;
    DOUBLE cRate;
    DOUBLE cSubTotal;
    DOUBLE cOptions;
    cFactor = mp_oParent->ActiveGanttVCCtl1.GetMathLib().DateTimeDiff(IL_HOUR, GetTime(dtpPickup), GetTime(dtpReturn)) / 24;
	if (chkGPS.GetCheck() == TRUE)
	{
	    txtGPS.SetWindowTextW(CStr(chkGPS_Tag * cFactor));
	    txtGPS_Tag = chkGPS_Tag * cFactor;
	}
	else
	{
		txtGPS.SetWindowTextW(_T(""));
		txtGPS_Tag = 0;
	}
	if (chkLDW.GetCheck() == TRUE)
	{
        txtLDW.SetWindowTextW(CStr(chkLDW_Tag * cFactor));
        txtLDW_Tag = chkLDW_Tag * cFactor;
	}
	else
	{
	    txtLDW.SetWindowTextW(_T(""));
	    txtLDW_Tag = 0;
	}
	if (chkPAI.GetCheck() == TRUE)
	{
        txtPAI.SetWindowTextW(CStr(chkPAI_Tag * cFactor));
        txtPAI_Tag = chkPAI_Tag * cFactor;
	}
	else
	{
        txtPAI.SetWindowTextW(_T(""));
        txtPAI_Tag = 0;
	}
	if (chkPEP.GetCheck() == TRUE)
	{
        txtPEP.SetWindowTextW(CStr(chkPEP_Tag * cFactor));
        txtPEP_Tag = chkPEP_Tag * cFactor;
	}
	else
	{
        txtPEP.SetWindowTextW(_T(""));
        txtPEP_Tag = 0;
	}
	if (chkALI.GetCheck() == TRUE)
	{
        txtALI.SetWindowTextW(CStr(chkALI_Tag * cFactor));
        txtALI_Tag = chkALI_Tag * cFactor;
	}
	else
	{
        txtALI.SetWindowTextW(_T(""));
        txtALI_Tag = 0;
	}

    sRowTag = Split(oTask.GetRow().GetTag(), _T("|"));
    cRate = CDbl(sRowTag[1]);
    txtERF.SetWindowTextW(CStr(txtERF_Tag * cFactor));
    txtWTB.SetWindowTextW(CStr(txtWTB_Tag * cFactor));
    txtRCFC.SetWindowTextW(CStr(txtRCFC_Tag * cFactor));
    txtVLF.SetWindowTextW(CStr(txtVLF_Tag * cFactor));
    txtCRF.SetWindowTextW(CStr(txtCRF_Tag * cRate * cFactor));
    txtSurcharge_Tag = (txtERF_Tag * cFactor) + (txtWTB_Tag * cFactor) + (txtRCFC_Tag * cFactor) + (txtVLF_Tag * cFactor) + (txtCRF_Tag * cRate * cFactor);
    txtSurcharge.SetWindowTextW(CStr(txtSurcharge_Tag));

    cOptions = txtGPS_Tag + txtLDW_Tag + txtPAI_Tag + txtPEP_Tag + txtALI_Tag;
    cSubTotal = txtSurcharge_Tag + (cRate * cFactor);

	txtTax_Tag = cSubTotal * lblTax_Tag;
    txtTax.SetWindowTextW(CStr(txtTax_Tag));

    txtEstimatedTotal_Tag = cSubTotal + txtTax_Tag + cOptions;
    txtEstimatedTotal.SetWindowTextW(CStr(txtEstimatedTotal_Tag));
	delete [] sRowTag;
}

CString fCarRentalReservation::GetACRISSDescription(LONG sPosition, CString sLetter)
{
	CString sReturn = _T("GetACRISSDescription Error");
	CRecordset oRecordset(&mp_oParent->mp_oConn);
	oRecordset.Open(CRecordset::forwardOnly, _T("SELECT * FROM tb_CR_ACRISS_Codes WHERE [sLetter] = '") + sLetter + _T("' AND [lPosition] = ") + CStr(sPosition), CRecordset::readOnly);
	if (oRecordset.IsEOF() == FALSE)
	{
		sReturn = GetStrField(oRecordset, _T("sDescription"));
	}
	oRecordset.Close();
	return sReturn;
}

HBRUSH fCarRentalReservation::OnCtlColor(CDC* pDC, CWnd* pWnd, UINT nCtlColor)
{
	HBRUSH hbr = CDialog::OnCtlColor(pDC, pWnd, nCtlColor);

	if (pWnd->GetDlgCtrlID() == IDC_LBLMODE)
	{
		pDC->SetTextColor(RGB(0, 0, 0));
		pDC->SetBkColor(lblModeBackColor);
		return (HBRUSH)mp_olblModeBrush.GetSafeHandle();
	}
	return hbr;
}

void fCarRentalReservation::OnBnClickedChkali()
{
	CalculateRate();
}

void fCarRentalReservation::OnBnClickedChkgps()
{
	CalculateRate();
}

void fCarRentalReservation::OnBnClickedChkldw()
{
	CalculateRate();
}

void fCarRentalReservation::OnBnClickedChkpai()
{
	CalculateRate();
}

void fCarRentalReservation::OnBnClickedChkpep()
{
	CalculateRate();
}

void fCarRentalReservation::OnOK()
{
	CclsTask oTask = GetCurrentTask();

        clsDB oDB(mp_oParent->mp_oConn);
        oDB.AddParameter(_T("lRowID"), CInt32(Replace(oTask.GetRowKey(), _T("K"), _T(""))), PT_NUMERIC);
        oDB.AddParameter(_T("lMode"), (int)mp_oParent->GetMode(), PT_NUMERIC);
        oDB.AddParameter(_T("sCity"), GetText(txtCity), PT_STRING);
        oDB.AddParameter(_T("sCustomerName"), GetText(txtName), PT_STRING);
        oDB.AddParameter(_T("sStateAbr"), GetText(txtState), PT_STRING);
        oDB.AddParameter(_T("sPhone"), GetText(txtPhone), PT_STRING);
        oDB.AddParameter(_T("sMobile"), GetText(txtMobile), PT_STRING);
        oDB.AddParameter(_T("sAddress"), GetText(txtAddress), PT_STRING);
        oDB.AddParameter(_T("sZIP"), GetText(txtZIP), PT_STRING);
        oDB.AddParameter(_T("dtPickUp"), GetTime(dtpPickup), PT_DATE);
        oDB.AddParameter(_T("dtReturn"), GetTime(dtpReturn), PT_DATE);
        oDB.AddParameter(_T("cRate"), CDbl(GetText(txtRate)), PT_NUMERIC);
        oDB.AddParameter(_T("cGPS"), chkGPS_Tag, PT_NUMERIC);
        oDB.AddParameter(_T("cLDW"), chkLDW_Tag, PT_NUMERIC);
        oDB.AddParameter(_T("cPAI"), chkPAI_Tag, PT_NUMERIC);
        oDB.AddParameter(_T("cPEP"), chkPEP_Tag, PT_NUMERIC);
        oDB.AddParameter(_T("cALI"), chkALI_Tag, PT_NUMERIC);
        oDB.AddParameter(_T("cERF"), txtERF_Tag, PT_NUMERIC);
        oDB.AddParameter(_T("cCRF"), txtCRF_Tag, PT_NUMERIC);
        oDB.AddParameter(_T("cRCFC"), txtRCFC_Tag, PT_NUMERIC);
        oDB.AddParameter(_T("cWTB"), txtWTB_Tag, PT_NUMERIC);
        oDB.AddParameter(_T("cVLF"), txtVLF_Tag, PT_NUMERIC);
        oDB.AddParameter(_T("cTax"), lblTax_Tag, PT_NUMERIC);
        oDB.AddParameter(_T("cEstimatedTotal"), txtEstimatedTotal_Tag, PT_NUMERIC);
        oDB.AddParameter(_T("bGPS"), chkGPS.GetCheck(), PT_BOOL);
        oDB.AddParameter(_T("bFSO"), chkFSO.GetCheck(), PT_BOOL);
        oDB.AddParameter(_T("bLDW"), chkLDW.GetCheck(), PT_BOOL);
        oDB.AddParameter(_T("bPAI"), chkPAI.GetCheck(), PT_BOOL);
        oDB.AddParameter(_T("bPEP"), chkPEP.GetCheck(), PT_BOOL);
        oDB.AddParameter(_T("bALI"), chkALI.GetCheck(), PT_BOOL);
		if (mp_yDialogMode == DM_ADD)
		{
			oTask.SetKey(_T("K") + oDB.ExecuteInsert(_T("tb_CR_Rentals")));
		}
		else
		{
			oDB.ExecuteUpdate(_T("tb_CR_Rentals"), _T("lTaskID = ") + Replace(oTask.GetKey(), _T("K"), _T("")));
		}

    oTask.SetText(GetText(txtName) + _T("\r\n") + _T("Phone: ") + GetText(txtPhone) + _T("\r\n") + _T("Estimated Total: ") + GetText(txtEstimatedTotal) + _T(" USD"));
    mp_oParent->ActiveGanttVCCtl1.Redraw();

	CDialog::OnOK();
}

void fCarRentalReservation::OnCancel()
{
	if (mp_yDialogMode == DM_ADD)
	{
		mp_oParent->mp_lDeleteTask = mp_oParent->ActiveGanttVCCtl1.GetTasks().GetCount();
	}
	CDialog::OnCancel();
}