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
#include "afxcmn.h"


// fPrintDialog dialog

class fPrintDialog : public CDialog
{
	DECLARE_DYNAMIC(fPrintDialog)

public:
	fPrintDialog(COleDateTime dtStartDate, COleDateTime dtEndDate, CWnd* pParent = NULL);   // standard constructor
	virtual ~fPrintDialog();

	CActiveGanttVCCtl* mp_oControl;
	COleDateTime mp_dtInitialStartDate;
	COleDateTime mp_dtInitialEndDate;
	BOOL mp_bLoaded;


	void SetStartDateCtrl(COleDateTime value);
	COleDateTime GetStartDateCtrl();
	void SetEndDateCtrl(COleDateTime value);
	COleDateTime GetEndDateCtrl();

	void mp_InitDateTimeCtrls(void);

	void mp_UpdateVisibility(void);


	void mp_GetNumberOfPages(void);
	CString GetDefaultPrinterName();



// Dialog Data
	enum { IDD = IDD_FPRINTDIALOG };

protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support

	DECLARE_MESSAGE_MAP()
public:
	CComboBox cboPrinterName;
	CStatic lblNumberOfPages;
	CEdit txtStartPage;
	CEdit txtEndPage;
	CEdit txtPageWidth;
	CEdit txtPageHeight;
	CEdit txtScale;
	CComboBox cboOrientation;
	CEdit txtSDYear;
	CEdit txtEDYear;
	CEdit txtSDMonth;
	CEdit txtEDMonth;
	CEdit txtSDDay;
	CEdit txtEDDay;
	CEdit txtSDHour;
	CEdit txtEDHour;
	CEdit txtSDMinute;
	CEdit txtEDMinute;
	CEdit txtSDSecond;
	CEdit txtEDSecond;
	CButton cmdPreview;
	CSpinButtonCtrl numScale;
	CEdit txtDPIX;
	CSpinButtonCtrl numDPIX;
	CEdit txtDPIY;
	CSpinButtonCtrl numDPIY;
	CSpinButtonCtrl numSDYear;
	CSpinButtonCtrl numSDMonth;
	CSpinButtonCtrl numSDDay;
	CSpinButtonCtrl numSDHour;
	CSpinButtonCtrl numSDMinute;
	CSpinButtonCtrl numSDSecond;
	CSpinButtonCtrl numEDYear;
	CSpinButtonCtrl numEDMonth;
	CSpinButtonCtrl numEDDay;
	CSpinButtonCtrl numEDHour;
	CSpinButtonCtrl numEDMinute;
	CSpinButtonCtrl numEDSecond;
	CEdit txtStartRow;
	CSpinButtonCtrl numStartRow;
	CEdit txtEndRow;
	CSpinButtonCtrl numEndRow;
	CEdit txtMarginLeft;
	CSpinButtonCtrl numMarginLeft;
	CEdit txtMarginTop;
	CSpinButtonCtrl numMarginTop;
	CEdit txtMarginRight;
	CSpinButtonCtrl numMarginRight;
	CEdit txtMarginBottom;
	CSpinButtonCtrl numMarginBottom;
	CEdit txtColumns;
	CSpinButtonCtrl numColumns;
	CEdit txtFactor;
	CSpinButtonCtrl numFactor;
	CComboBox cboResolution;
	CComboBox cboInterval;
	CButton chkKeepRowsTogether;
	CButton chkKeepColumnsTogether;
	CButton chkKeepTimeIntervalsTogether;
	CButton chkHeadingsInEveryPage;
	CButton chkColumnsInEveryPage;
	CButton chkShowAllColumns;
	CButton optAll;
	CButton optFrom;
	CStatic lblDPIX;
	CStatic lblDPIY;




	virtual BOOL OnInitDialog();
protected:
	virtual void OnOK();
public:
	afx_msg void OnBnClickedCmdpreview();
	afx_msg void OnCbnSelchangeCboprintername();
	afx_msg void OnCbnSelchangeCboresolution();
	afx_msg void OnCbnSelchangeCbointerval();
	afx_msg void OnEnChangeTxtfactor();
	afx_msg void OnEnChangeTxtdpix();
	afx_msg void OnEnChangeTxtdpiy();
	afx_msg void OnEnChangeTxtscale();
	afx_msg void OnBnClickedChkkeeprowstogether();
	afx_msg void OnBnClickedChkkeepcolumnstogether();
	afx_msg void OnBnClickedChkkeeptimeintervalstogether();
	afx_msg void OnBnClickedChkheadingsineverypage();
	afx_msg void OnBnClickedChkcolumnsineverypage();
	afx_msg void OnEnChangeTxtcolumns();
	afx_msg void OnBnClickedChkshowallcolumns();
	afx_msg void OnEnSetfocusTxtstartpage();
	afx_msg void OnEnSetfocusTxtendpage();
	afx_msg void OnEnChangeTxtmarginleft();
	afx_msg void OnEnChangeTxtmargintop();
	afx_msg void OnEnChangeTxtmarginright();
	afx_msg void OnEnChangeTxtmarginbottom();
	afx_msg void OnEnChangeTxtstartrow();
	afx_msg void OnEnChangeTxtendrow();
	afx_msg void OnDeltaposNummarginleft(NMHDR *pNMHDR, LRESULT *pResult);
	afx_msg void OnDeltaposNummarginright(NMHDR *pNMHDR, LRESULT *pResult);
	afx_msg void OnDeltaposNummargintop(NMHDR *pNMHDR, LRESULT *pResult);
	afx_msg void OnDeltaposNummarginbottom(NMHDR *pNMHDR, LRESULT *pResult);
	afx_msg void OnCbnSelchangeCboorientation();
	CStatic lblInterval;
	CStatic lblFactor;
};