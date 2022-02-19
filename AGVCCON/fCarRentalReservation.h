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
#include "afxwin.h"
#include "afxdtctl.h"
#include "atlimage.h" //CImage

class fCarRental;

class fCarRentalReservation : public CDialog
{
	DECLARE_DYNAMIC(fCarRentalReservation)

public:
	fCarRentalReservation(PRG_DIALOGMODE yDialogMode, CWnd* pParent, CString sTaskID);
	virtual ~fCarRentalReservation();

	PRG_DIALOGMODE mp_yDialogMode;
	CString mp_sTaskID;
	fCarRental* mp_oParent;

	void GetStateTax();
	void CalculateRate();
	CString GetACRISSDescription(LONG sPosition, CString sLetter);
	CclsTask GetCurrentTask();

	DOUBLE chkGPS_Tag;
	DOUBLE chkLDW_Tag;
	DOUBLE chkPAI_Tag;
	DOUBLE chkPEP_Tag;
	DOUBLE chkALI_Tag;

	DOUBLE txtGPS_Tag;
	DOUBLE txtLDW_Tag;
	DOUBLE txtPAI_Tag;
	DOUBLE txtPEP_Tag;
	DOUBLE txtALI_Tag;

	DOUBLE txtSurcharge_Tag;
	DOUBLE txtTax_Tag;

	DOUBLE txtERF_Tag;
	DOUBLE txtCRF_Tag;
	DOUBLE txtRCFC_Tag;
	DOUBLE txtWTB_Tag;
	DOUBLE txtVLF_Tag;
	DOUBLE lblTax_Tag;
	DOUBLE txtEstimatedTotal_Tag;
	COLORREF lblModeBackColor;
	CImage mp_oImage;
	CStatic mp_oPicture;

	CBrush mp_olblModeBrush;


	enum { IDD = IDD_FCARRENTALRESERVATION };

protected:
	virtual void DoDataExchange(CDataExchange* pDX);    

	DECLARE_MESSAGE_MAP()
public:
	CStatic lblMode;
	CEdit txtName;
	CEdit txtAddress;
	CEdit txtCity;
	CEdit txtState;
	CEdit txtZIP;
	CEdit txtPhone;
	CEdit txtMobile;
	CEdit txtPickup;
	CDateTimeCtrl dtpPickup;
	CEdit txtReturn;
	CDateTimeCtrl dtpReturn;
	CEdit txtRate;
	CEdit txtDescription;
	CEdit txtACRISS1;
	CEdit txtACRISS2;
	CEdit txtACRISS3;
	CEdit txtACRISS4;
	CButton chkGPS;
	CButton chkFSO;
	CButton chkLDW;
	CButton chkPAI;
	CButton chkPEP;
	CButton chkALI;
	CEdit txtGPS;
	CEdit txtLDW;
	CEdit txtPAI;
	CEdit txtPEP;
	CEdit txtALI;
	CEdit txtCRF;
	CEdit txtRCFC;
	CEdit txtERF;
	CEdit txtVLF;
	CEdit txtWTB;
	CEdit txtSurcharge;
	CEdit txtTax;
	CEdit txtEstimatedTotal;
	CStatic lblTax;
	virtual BOOL OnInitDialog();
protected:
	virtual void OnOK();
	virtual void OnCancel();
public:
	afx_msg HBRUSH OnCtlColor(CDC* pDC, CWnd* pWnd, UINT nCtlColor);
	afx_msg void OnBnClickedChkali();
	afx_msg void OnBnClickedChkgps();
	afx_msg void OnBnClickedChkldw();
	afx_msg void OnBnClickedChkpai();
	afx_msg void OnBnClickedChkpep();
	
};