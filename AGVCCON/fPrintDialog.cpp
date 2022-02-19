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
#include "fPrintDialog.h"
#include "winspool.h"
#include <math.h>
#include "fPrintPreview.h"

#undef GetPrinter

BEGIN_MESSAGE_MAP(fPrintDialog, CDialog)
	ON_BN_CLICKED(IDC_CMDPREVIEW, &fPrintDialog::OnBnClickedCmdpreview)
	ON_CBN_SELCHANGE(IDC_CBOPRINTERNAME, &fPrintDialog::OnCbnSelchangeCboprintername)
	ON_CBN_SELCHANGE(IDC_CBORESOLUTION, &fPrintDialog::OnCbnSelchangeCboresolution)
	ON_CBN_SELCHANGE(IDC_CBOINTERVAL, &fPrintDialog::OnCbnSelchangeCbointerval)
	ON_EN_CHANGE(IDC_TXTFACTOR, &fPrintDialog::OnEnChangeTxtfactor)
	ON_EN_CHANGE(IDC_TXTDPIX, &fPrintDialog::OnEnChangeTxtdpix)
	ON_EN_CHANGE(IDC_TXTDPIY, &fPrintDialog::OnEnChangeTxtdpiy)
	ON_EN_CHANGE(IDC_TXTSCALE, &fPrintDialog::OnEnChangeTxtscale)
	ON_BN_CLICKED(IDC_CHKKEEPROWSTOGETHER, &fPrintDialog::OnBnClickedChkkeeprowstogether)
	ON_BN_CLICKED(IDC_CHKKEEPCOLUMNSTOGETHER, &fPrintDialog::OnBnClickedChkkeepcolumnstogether)
	ON_BN_CLICKED(IDC_CHKKEEPTIMEINTERVALSTOGETHER, &fPrintDialog::OnBnClickedChkkeeptimeintervalstogether)
	ON_BN_CLICKED(IDC_CHKHEADINGSINEVERYPAGE, &fPrintDialog::OnBnClickedChkheadingsineverypage)
	ON_BN_CLICKED(IDC_CHKCOLUMNSINEVERYPAGE, &fPrintDialog::OnBnClickedChkcolumnsineverypage)
	ON_EN_CHANGE(IDC_TXTCOLUMNS, &fPrintDialog::OnEnChangeTxtcolumns)
	ON_BN_CLICKED(IDC_CHKSHOWALLCOLUMNS, &fPrintDialog::OnBnClickedChkshowallcolumns)
	ON_EN_SETFOCUS(IDC_TXTSTARTPAGE, &fPrintDialog::OnEnSetfocusTxtstartpage)
	ON_EN_SETFOCUS(IDC_TXTENDPAGE, &fPrintDialog::OnEnSetfocusTxtendpage)
	ON_EN_CHANGE(IDC_TXTMARGINLEFT, &fPrintDialog::OnEnChangeTxtmarginleft)
	ON_EN_CHANGE(IDC_TXTMARGINTOP, &fPrintDialog::OnEnChangeTxtmargintop)
	ON_EN_CHANGE(IDC_TXTMARGINRIGHT, &fPrintDialog::OnEnChangeTxtmarginright)
	ON_EN_CHANGE(IDC_TXTMARGINBOTTOM, &fPrintDialog::OnEnChangeTxtmarginbottom)
	ON_EN_CHANGE(IDC_TXTSTARTROW, &fPrintDialog::OnEnChangeTxtstartrow)
	ON_EN_CHANGE(IDC_TXTENDROW, &fPrintDialog::OnEnChangeTxtendrow)
	ON_NOTIFY(UDN_DELTAPOS, IDC_NUMMARGINLEFT, &fPrintDialog::OnDeltaposNummarginleft)
	ON_NOTIFY(UDN_DELTAPOS, IDC_NUMMARGINRIGHT, &fPrintDialog::OnDeltaposNummarginright)
	ON_NOTIFY(UDN_DELTAPOS, IDC_NUMMARGINTOP, &fPrintDialog::OnDeltaposNummargintop)
	ON_NOTIFY(UDN_DELTAPOS, IDC_NUMMARGINBOTTOM, &fPrintDialog::OnDeltaposNummarginbottom)
	ON_CBN_SELCHANGE(IDC_CBOORIENTATION, &fPrintDialog::OnCbnSelchangeCboorientation)
END_MESSAGE_MAP()

IMPLEMENT_DYNAMIC(fPrintDialog, CDialog)

void fPrintDialog::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	DDX_Control(pDX, IDC_CBOPRINTERNAME, cboPrinterName);
	DDX_Control(pDX, IDC_LBLNUMBEROFPAGES, lblNumberOfPages);
	DDX_Control(pDX, IDC_TXTSTARTPAGE, txtStartPage);
	DDX_Control(pDX, IDC_TXTENDPAGE, txtEndPage);
	DDX_Control(pDX, IDC_TXTSCALE, txtScale);
	DDX_Control(pDX, IDC_CBOORIENTATION, cboOrientation);
	DDX_Control(pDX, IDC_TXTSDYEAR, txtSDYear);
	DDX_Control(pDX, IDC_TXTEDYEAR, txtEDYear);
	DDX_Control(pDX, IDC_TXTSDMONTH, txtSDMonth);
	DDX_Control(pDX, IDC_TXTEDMONTH, txtEDMonth);
	DDX_Control(pDX, IDC_TXTSDDAY, txtSDDay);
	DDX_Control(pDX, IDC_TXTEDDAY, txtEDDay);
	DDX_Control(pDX, IDC_TXTSDHOUR, txtSDHour);
	DDX_Control(pDX, IDC_TXTEDHOUR, txtEDHour);
	DDX_Control(pDX, IDC_TXTSDMINUTE, txtSDMinute);
	DDX_Control(pDX, IDC_TXTEDMINUTE, txtEDMinute);
	DDX_Control(pDX, IDC_TXTSDSECOND, txtSDSecond);
	DDX_Control(pDX, IDC_TXTEDSECOND, txtEDSecond);
	DDX_Control(pDX, IDC_CMDPREVIEW, cmdPreview);
	DDX_Control(pDX, IDC_NUMSCALE, numScale);
	DDX_Control(pDX, IDC_TXTDPIX, txtDPIX);
	DDX_Control(pDX, IDC_NUMDPIX, numDPIX);
	DDX_Control(pDX, IDC_TXTDPIY, txtDPIY);
	DDX_Control(pDX, IDC_NUMDPIY, numDPIY);
	DDX_Control(pDX, IDC_NUMSDYEAR, numSDYear);
	DDX_Control(pDX, IDC_NUMSDMONTH, numSDMonth);
	DDX_Control(pDX, IDC_NUMSDDAY, numSDDay);
	DDX_Control(pDX, IDC_NUMSDHOUR, numSDHour);
	DDX_Control(pDX, IDC_NUMSDMINUTE, numSDMinute);
	DDX_Control(pDX, IDC_NUMSDSECOND, numSDSecond);
	DDX_Control(pDX, IDC_NUMEDYEAR, numEDYear);
	DDX_Control(pDX, IDC_NUMEDMONTH, numEDMonth);
	DDX_Control(pDX, IDC_NUMEDDAY, numEDDay);
	DDX_Control(pDX, IDC_NUMEDHOUR, numEDHour);
	DDX_Control(pDX, IDC_NUMEDMINUTE, numEDMinute);
	DDX_Control(pDX, IDC_NUMEDSECOND, numEDSecond);
	DDX_Control(pDX, IDC_TXTSTARTROW, txtStartRow);
	DDX_Control(pDX, IDC_NUMSTARTROW, numStartRow);
	DDX_Control(pDX, IDC_TXTENDROW, txtEndRow);
	DDX_Control(pDX, IDC_NUMENDROW, numEndRow);
	DDX_Control(pDX, IDC_TXTMARGINLEFT, txtMarginLeft);
	DDX_Control(pDX, IDC_NUMMARGINLEFT, numMarginLeft);
	DDX_Control(pDX, IDC_TXTMARGINTOP, txtMarginTop);
	DDX_Control(pDX, IDC_NUMMARGINTOP, numMarginTop);
	DDX_Control(pDX, IDC_TXTMARGINRIGHT, txtMarginRight);
	DDX_Control(pDX, IDC_NUMMARGINRIGHT, numMarginRight);
	DDX_Control(pDX, IDC_TXTMARGINBOTTOM, txtMarginBottom);
	DDX_Control(pDX, IDC_NUMMARGINBOTTOM, numMarginBottom);
	DDX_Control(pDX, IDC_TXTCOLUMNS, txtColumns);
	DDX_Control(pDX, IDC_NUMCOLUMNS, numColumns);
	DDX_Control(pDX, IDC_TXTFACTOR, txtFactor);
	DDX_Control(pDX, IDC_NUMFACTOR, numFactor);
	DDX_Control(pDX, IDC_CBORESOLUTION, cboResolution);
	DDX_Control(pDX, IDC_CBOINTERVAL, cboInterval);
	DDX_Control(pDX, IDC_CHKKEEPROWSTOGETHER, chkKeepRowsTogether);
	DDX_Control(pDX, IDC_CHKKEEPCOLUMNSTOGETHER, chkKeepColumnsTogether);
	DDX_Control(pDX, IDC_CHKKEEPTIMEINTERVALSTOGETHER, chkKeepTimeIntervalsTogether);
	DDX_Control(pDX, IDC_CHKHEADINGSINEVERYPAGE, chkHeadingsInEveryPage);
	DDX_Control(pDX, IDC_CHKCOLUMNSINEVERYPAGE, chkColumnsInEveryPage);
	DDX_Control(pDX, IDC_CHKSHOWALLCOLUMNS, chkShowAllColumns);
	DDX_Control(pDX, IDC_OPTALL, optAll);
	DDX_Control(pDX, IDC_LBLDPIX, lblDPIX);
	DDX_Control(pDX, IDC_LBLDPIY, lblDPIY);
	DDX_Control(pDX, IDCANCEL, optFrom);
	DDX_Control(pDX, IDC_LBLINTERVAL, lblInterval);
	DDX_Control(pDX, IDC_LBLFACTOR, lblFactor);
}

fPrintDialog::fPrintDialog(COleDateTime dtStartDate, COleDateTime dtEndDate, CWnd* pParent)
	: CDialog(fPrintDialog::IDD, pParent)
{
	mp_dtInitialStartDate = dtStartDate;
	mp_dtInitialEndDate = dtEndDate;
	mp_bLoaded = FALSE;
}

fPrintDialog::~fPrintDialog()
{
	mp_oControl->GetPrinter().Terminate();
}

//// Form Load/Closed

BOOL fPrintDialog::OnInitDialog()
{
	CDialog::OnInitDialog();

	mp_InitDateTimeCtrls();

	// Printer Name
	CString sDefaultPrinterName;
	DWORD dwSizeNeeded;
	DWORD dwNumItems;
	DWORD dwItem;
	LPPRINTER_INFO_2 lpInfo = NULL;

	sDefaultPrinterName = GetDefaultPrinterName();

	EnumPrinters(PRINTER_ENUM_LOCAL, NULL, 2, NULL, 0, &dwSizeNeeded, &dwNumItems);
	lpInfo = (LPPRINTER_INFO_2) HeapAlloc(GetProcessHeap(), HEAP_ZERO_MEMORY, dwSizeNeeded);
	if (EnumPrinters(PRINTER_ENUM_LOCAL, NULL, 2, (LPBYTE)lpInfo, dwSizeNeeded, &dwSizeNeeded, &dwNumItems) == 0)
	{
		MessageBox(_T("Before attempting to print a printer must be installed via the control panel"),_T("No printer installed"), MB_OK);
		return TRUE;
	}
	for ( dwItem = 0; dwItem < dwNumItems; dwItem++ )
	{
		CString sPrinterName = lpInfo[dwItem].pPrinterName;
		CString sPortName = lpInfo[dwItem].pPortName;
		cboPrinterName.AddString(sPrinterName);
		if (sPrinterName == sDefaultPrinterName)
		{
			cboPrinterName.SetCurSel(dwItem);
		}
	}
	HeapFree (GetProcessHeap(), 0, lpInfo );


	// Orientation
	cboOrientation.AddString(_T("OT_PORTRAIT"));
	cboOrientation.AddString(_T("OT_LANDSCAPE"));
	cboOrientation.SetCurSel(mp_oControl->GetPrinter().GetOrientation());

	// Resolutions
	cboResolution.AddString(_T("PR_HIGH"));
	cboResolution.AddString(_T("PR_MEDIUM"));
	cboResolution.AddString(_T("PR_LOW"));
	cboResolution.AddString(_T("PR_DRAFT"));
	cboResolution.AddString(_T("PR_CUSTOM"));
	cboResolution.SetCurSel(mp_oControl->GetPrinter().GetPrinterResolution());
	OnCbnSelchangeCboresolution();

	// Intervals
	cboInterval.AddString(_T("IL_NANOSECOND"));
	cboInterval.AddString(_T("IL_MICROSECOND"));
	cboInterval.AddString(_T("IL_MILLISECOND"));
	cboInterval.AddString(_T("IL_SECOND"));
	cboInterval.AddString(_T("IL_MINUTE"));
	cboInterval.AddString(_T("IL_HOUR"));
	cboInterval.AddString(_T("IL_DAY"));
	cboInterval.AddString(_T("IL_WEEK"));
	cboInterval.AddString(_T("IL_MONTH"));
	cboInterval.AddString(_T("IL_QUARTER"));
	cboInterval.AddString(_T("IL_YEAR"));
	cboInterval.SetCurSel(mp_oControl->GetPrinter().GetInterval());

	InitNumCtrl(numScale, txtScale, 1, 200, CInt32(mp_oControl->GetPrinter().GetScale() * 100));

	chkKeepRowsTogether.SetCheck(mp_oControl->GetPrinter().GetKeepRowsTogether());
	chkKeepColumnsTogether.SetCheck(mp_oControl->GetPrinter().GetKeepColumnsTogether());
	chkKeepTimeIntervalsTogether.SetCheck(mp_oControl->GetPrinter().GetKeepTimeIntervalsTogether());
	chkHeadingsInEveryPage.SetCheck(mp_oControl->GetPrinter().GetHeadingsInEveryPage());

	chkColumnsInEveryPage.SetCheck(mp_oControl->GetPrinter().GetColumnsInEveryPage());
    InitNumCtrl(numColumns, txtColumns, 0, mp_oControl->GetColumns().GetCount(), mp_oControl->GetPrinter().GetColumns());

	chkShowAllColumns.SetCheck(mp_oControl->GetPrinter().GetShowAllColumns());

	InitNumCtrlDiv100(numMarginLeft, txtMarginLeft, 0, 300, mp_oControl->GetPrinter().GetMarginLeft());
	InitNumCtrlDiv100(numMarginTop, txtMarginTop, 0, 300, mp_oControl->GetPrinter().GetMarginTop());
	InitNumCtrlDiv100(numMarginRight, txtMarginRight, 0, 300, mp_oControl->GetPrinter().GetMarginRight());
	InitNumCtrlDiv100(numMarginBottom, txtMarginBottom, 0, 300, mp_oControl->GetPrinter().GetMarginBottom());

	InitNumCtrl(numStartRow, txtStartRow, 1, mp_oControl->GetRows().GetCount(), 1);
    InitNumCtrl(numEndRow, txtEndRow, 1, mp_oControl->GetRows().GetCount(), mp_oControl->GetRows().GetCount());

	InitNumCtrl(numFactor, txtFactor, 1, 1000, mp_oControl->GetPrinter().GetFactor());

	optAll.SetCheck(TRUE);

	txtStartPage.SetWindowTextW(_T("1"));

	
    if (mp_oControl->GetPrinter().Initialize(GetStartDateCtrl(), GetEndDateCtrl(), numStartRow.GetPos32(), numEndRow.GetPos32()) == TRUE)
	{
        mp_GetNumberOfPages();
        mp_bLoaded = TRUE;
	}
	mp_UpdateVisibility();
	return TRUE;
}

//// Properties

COleDateTime fPrintDialog::GetStartDateCtrl()
{
	COleDateTime dtReturn;
	dtReturn.SetDateTime(numSDYear.GetPos32(), numSDMonth.GetPos32(), numSDDay.GetPos32(), numSDHour.GetPos32(), numSDMinute.GetPos32(), numSDSecond.GetPos32());
	return dtReturn;
}

void fPrintDialog::SetStartDateCtrl(COleDateTime value)
{
	numSDYear.SetPos32(value.GetYear());
	numSDMonth.SetPos32(value.GetMonth());
	numSDDay.SetPos32(value.GetDay());
	numSDHour.SetPos32(value.GetHour());
	numSDMinute.SetPos32(value.GetMinute());
	numSDSecond.SetPos32(value.GetSecond());
}

COleDateTime fPrintDialog::GetEndDateCtrl()
{
	COleDateTime dtReturn;
	dtReturn.SetDateTime(numEDYear.GetPos32(), numEDMonth.GetPos32(), numEDDay.GetPos32(), numEDHour.GetPos32(), numEDMinute.GetPos32(), numEDSecond.GetPos32());
	return dtReturn;
}

void fPrintDialog::SetEndDateCtrl(COleDateTime value)
{
	numEDYear.SetPos32(value.GetYear());
	numEDMonth.SetPos32(value.GetMonth());
	numEDDay.SetPos32(value.GetDay());
	numEDHour.SetPos32(value.GetHour());
	numEDMinute.SetPos32(value.GetMinute());
	numEDSecond.SetPos32(value.GetSecond());
}

void fPrintDialog::mp_InitDateTimeCtrls(void)
{
	InitNumCtrl(numSDYear, txtSDYear, 0, 3000, mp_dtInitialStartDate.GetYear());
	InitNumCtrl(numSDMonth, txtSDMonth, 1, 12, mp_dtInitialStartDate.GetMonth());
	InitNumCtrl(numSDDay, txtSDDay, 1, 12, mp_dtInitialStartDate.GetDay());
	InitNumCtrl(numSDHour, txtSDHour, 0, 23, mp_dtInitialStartDate.GetHour());
	InitNumCtrl(numSDMinute, txtSDMinute, 0, 59, mp_dtInitialStartDate.GetMinute());
	InitNumCtrl(numSDSecond, txtSDSecond, 0, 59, mp_dtInitialStartDate.GetSecond());

	InitNumCtrl(numEDYear, txtEDYear, 0, 3000, mp_dtInitialEndDate.GetYear());
	InitNumCtrl(numEDMonth, txtEDMonth, 1, 12, mp_dtInitialEndDate.GetMonth());
	InitNumCtrl(numEDDay, txtEDDay, 1, 12, mp_dtInitialEndDate.GetDay());
	InitNumCtrl(numEDHour, txtEDHour, 0, 23, mp_dtInitialEndDate.GetHour());
	InitNumCtrl(numEDMinute, txtEDMinute, 0, 59, mp_dtInitialEndDate.GetMinute());
	InitNumCtrl(numEDSecond, txtEDSecond, 0, 59, mp_dtInitialEndDate.GetSecond());

}

//// Control Updating

void fPrintDialog::mp_UpdateVisibility(void)
{
    if (chkKeepTimeIntervalsTogether.GetCheck() == FALSE)
    {
        lblInterval.ShowWindow(SW_HIDE);
        cboInterval.ShowWindow(SW_HIDE);
        lblFactor.ShowWindow(SW_HIDE);
        numFactor.ShowWindow(SW_HIDE);
		txtFactor.ShowWindow(SW_HIDE);
    }
    if (chkKeepColumnsTogether.GetCheck() == FALSE)
    {
        chkKeepTimeIntervalsTogether.ShowWindow(SW_HIDE);
        lblInterval.ShowWindow(SW_HIDE);
        cboInterval.ShowWindow(SW_HIDE);
        lblFactor.ShowWindow(SW_HIDE);
        numFactor.ShowWindow(SW_HIDE);
		txtFactor.ShowWindow(SW_HIDE);
    }
    if (chkKeepTimeIntervalsTogether.GetCheck() == TRUE && chkKeepColumnsTogether.GetCheck() == TRUE)
    {
        chkKeepTimeIntervalsTogether.ShowWindow(SW_SHOW);
        lblInterval.ShowWindow(SW_SHOW);
        cboInterval.ShowWindow(SW_SHOW);
        lblFactor.ShowWindow(SW_SHOW);
        numFactor.ShowWindow(SW_SHOW);
		txtFactor.ShowWindow(SW_SHOW);
    }
}

void fPrintDialog::OnCbnSelchangeCboprintername()
{
	CString sSelectedText;
	cboPrinterName.GetLBText(cboPrinterName.GetCurSel(), sSelectedText);
	if (mp_oControl->GetPrinter().GetPrinterName() != sSelectedText)
	{
		mp_oControl->GetPrinter().SetPrinterName(sSelectedText);
		if (mp_oControl->GetPrinter().Initialize(GetStartDateCtrl(), GetEndDateCtrl(), numStartRow.GetPos32(), numEndRow.GetPos32()) == TRUE)
		{
			cboOrientation.SetCurSel(mp_oControl->GetPrinter().GetOrientation());
			cboResolution.SetCurSel(mp_oControl->GetPrinter().GetPrinterResolution());
		}
	}
}

void fPrintDialog::OnCbnSelchangeCboorientation()
{
	if (mp_oControl->GetPrinter().GetOrientation() != cboOrientation.GetCurSel())
	{
		mp_oControl->GetPrinter().SetOrientation((GRE_ORIENTATION)cboOrientation.GetCurSel());
		mp_GetNumberOfPages();
	}
}

void fPrintDialog::OnCbnSelchangeCboresolution()
{
	if (cboResolution.GetCurSel() == PR_CUSTOM)
	{
		lblDPIX.ShowWindow(SW_SHOW);
		numDPIX.ShowWindow(SW_SHOW);
		txtDPIX.ShowWindow(SW_SHOW);
		lblDPIY.ShowWindow(SW_SHOW);
		numDPIY.ShowWindow(SW_SHOW);
		txtDPIY.ShowWindow(SW_SHOW);
		InitNumCtrl(numDPIX, txtDPIX, 50, 6000, mp_oControl->GetPrinter().GetHorizontalDPI());
		InitNumCtrl(numDPIY, txtDPIY, 50, 6000, mp_oControl->GetPrinter().GetVerticalDPI());
	}
	else
	{
		lblDPIX.ShowWindow(SW_HIDE);
		numDPIX.ShowWindow(SW_HIDE);
		txtDPIX.ShowWindow(SW_HIDE);
		lblDPIY.ShowWindow(SW_HIDE);
		numDPIY.ShowWindow(SW_HIDE);
		txtDPIY.ShowWindow(SW_HIDE);
	}

	if (mp_oControl->GetPrinter().GetPrinterResolution() != cboResolution.GetCurSel())
	{
		mp_oControl->GetPrinter().SetPrinterResolution((GRE_PRINTERRESOLUTION)cboResolution.GetCurSel());
	}
}

void fPrintDialog::OnCbnSelchangeCbointerval()
{
	if (mp_oControl->GetPrinter().GetInterval() != cboInterval.GetCurSel())
	{
		mp_oControl->GetPrinter().SetInterval((E_INTERVAL)cboInterval.GetCurSel());
	}
}

void fPrintDialog::OnEnChangeTxtfactor()
{
	if (mp_oControl->GetPrinter().GetFactor() != numFactor.GetPos32())
	{
		mp_oControl->GetPrinter().SetFactor(numFactor.GetPos32());
		mp_GetNumberOfPages();
	}
}

void fPrintDialog::OnEnChangeTxtdpix()
{

	if (mp_oControl->GetPrinter().GetHorizontalDPI() != numDPIX.GetPos32())
	{
		mp_oControl->GetPrinter().SetHorizontalDPI(numDPIX.GetPos32());
	}
}

void fPrintDialog::OnEnChangeTxtdpiy()
{
	if (mp_oControl->GetPrinter().GetVerticalDPI() != numDPIY.GetPos32())
	{
		mp_oControl->GetPrinter().SetVerticalDPI(numDPIY.GetPos32());
	}
}

void fPrintDialog::OnEnChangeTxtscale()
{
	if (mp_oControl->GetPrinter().GetScale() != (CSng(numScale.GetPos32()) / CSng(100)))
	{
		mp_oControl->GetPrinter().SetScale(CSng(numScale.GetPos32()) / CSng(100));
		mp_GetNumberOfPages();
	}
}

void fPrintDialog::OnBnClickedChkkeeprowstogether()
{
	if (mp_oControl->GetPrinter().GetKeepRowsTogether() != chkKeepRowsTogether.GetCheck())
	{
		mp_oControl->GetPrinter().SetKeepRowsTogether(chkKeepRowsTogether.GetCheck());
		mp_GetNumberOfPages();
	}
}

void fPrintDialog::OnBnClickedChkkeepcolumnstogether()
{
	if (mp_oControl->GetPrinter().GetKeepColumnsTogether() != chkKeepColumnsTogether.GetCheck())
	{
		mp_oControl->GetPrinter().SetKeepColumnsTogether(chkKeepColumnsTogether.GetCheck());
		mp_GetNumberOfPages();
	}
	mp_UpdateVisibility();
}

void fPrintDialog::OnBnClickedChkkeeptimeintervalstogether()
{
	if (mp_oControl->GetPrinter().GetKeepTimeIntervalsTogether() != chkKeepTimeIntervalsTogether.GetCheck())
	{
		mp_oControl->GetPrinter().SetKeepTimeIntervalsTogether(chkKeepTimeIntervalsTogether.GetCheck());
		mp_GetNumberOfPages();
	}
	mp_UpdateVisibility();
}

void fPrintDialog::OnBnClickedChkheadingsineverypage()
{
	if (mp_oControl->GetPrinter().GetHeadingsInEveryPage() != chkHeadingsInEveryPage.GetCheck())
	{
		mp_oControl->GetPrinter().SetHeadingsInEveryPage(chkHeadingsInEveryPage.GetCheck());
		mp_GetNumberOfPages();
	}
}

void fPrintDialog::OnBnClickedChkcolumnsineverypage()
{
	if (mp_oControl->GetPrinter().GetColumnsInEveryPage() != chkColumnsInEveryPage.GetCheck())
	{
		mp_oControl->GetPrinter().SetColumnsInEveryPage(chkColumnsInEveryPage.GetCheck());
		mp_GetNumberOfPages();
	}
}

void fPrintDialog::OnEnChangeTxtcolumns()
{
	if (mp_oControl->GetPrinter().GetColumns() != numColumns.GetPos32())
	{
		mp_oControl->GetPrinter().SetColumns(numColumns.GetPos32());
		mp_GetNumberOfPages();
	}
}

void fPrintDialog::OnBnClickedChkshowallcolumns()
{
	if (mp_oControl->GetPrinter().GetShowAllColumns() != chkShowAllColumns.GetCheck())
	{
		mp_oControl->GetPrinter().SetShowAllColumns(chkShowAllColumns.GetCheck());
		mp_GetNumberOfPages();
	}
}

void fPrintDialog::OnEnSetfocusTxtstartpage()
{
	optFrom.SetCheck(TRUE);
}

void fPrintDialog::OnEnSetfocusTxtendpage()
{
	optFrom.SetCheck(TRUE);
}

void fPrintDialog::OnEnChangeTxtmarginleft()
{
	if (mp_bLoaded == FALSE) { return; }
	if (mp_oControl->GetPrinter().GetMarginLeft() != numMarginLeft.GetPos32())
	{

		mp_oControl->GetPrinter().SetMarginLeft(numMarginLeft.GetPos32());
		mp_GetNumberOfPages();
	}
}

void fPrintDialog::OnEnChangeTxtmargintop()
{
	if (mp_bLoaded == FALSE) { return; }
	if (mp_oControl->GetPrinter().GetMarginTop() != numMarginTop.GetPos32())
	{
		mp_oControl->GetPrinter().SetMarginTop(numMarginTop.GetPos32());
		mp_GetNumberOfPages();
	}
}

void fPrintDialog::OnEnChangeTxtmarginright()
{
	if (mp_bLoaded == FALSE) { return; }
	if (mp_oControl->GetPrinter().GetMarginRight() != numMarginRight.GetPos32())
	{
		mp_oControl->GetPrinter().SetMarginRight(numMarginRight.GetPos32());
		mp_GetNumberOfPages();
	}
}

void fPrintDialog::OnEnChangeTxtmarginbottom()
{
	if (mp_bLoaded == FALSE) { return; }
	if (mp_oControl->GetPrinter().GetMarginBottom() != numMarginBottom.GetPos32())
	{
		mp_oControl->GetPrinter().SetMarginBottom(numMarginBottom.GetPos32());
		mp_GetNumberOfPages();
	}
}

void fPrintDialog::OnEnChangeTxtstartrow()
{
	if (mp_bLoaded == FALSE) { return; }
	if (mp_oControl->GetPrinter().Initialize(GetStartDateCtrl(), GetEndDateCtrl(), numStartRow.GetPos32(), numEndRow.GetPos32()) == TRUE)
	{
		mp_GetNumberOfPages();
	}
}

void fPrintDialog::OnEnChangeTxtendrow()
{
	if (mp_bLoaded == FALSE) { return; }
	if (mp_oControl->GetPrinter().Initialize(GetStartDateCtrl(), GetEndDateCtrl(), numStartRow.GetPos32(), numEndRow.GetPos32()) == TRUE)
	{
		mp_GetNumberOfPages();
	}

}

//// Functions

void fPrintDialog::mp_GetNumberOfPages(void)
{
	mp_oControl->GetPrinter().Calculate();
	txtEndPage.SetWindowTextW(CStr(mp_oControl->GetPrinter().GetPages()));
	lblNumberOfPages.SetWindowTextW(_T("Total Pages: ") + CStr(mp_oControl->GetPrinter().GetPages()));
}

CString fPrintDialog::GetDefaultPrinterName()
{
	CString sReturn;
	TCHAR szPrinterName[MAX_PATH] = _T("");
	DWORD nLen = MAX_PATH ;

	GetDefaultPrinterW(szPrinterName, &nLen);
	sReturn = (TCHAR*)szPrinterName;
	return sReturn;
}

//// Buttons

void fPrintDialog::OnOK()
{
	if (mp_oControl->GetPrinter().Initialize(GetStartDateCtrl(), GetEndDateCtrl(), numStartRow.GetPos32(), numEndRow.GetPos32()) == TRUE)
	{
		if (optAll.GetCheck() == TRUE)
		{
			BeginWaitCursor();
			mp_oControl->GetPrinter().PrintAll();
			EndWaitCursor();
		}
		else
		{
			CString sStartPage;
			CString sEndPage;
			txtStartPage.GetWindowTextW(sStartPage);
			txtEndPage.GetWindowTextW(sEndPage);
			if (sStartPage.GetLength() == 0 || sEndPage.GetLength() == 0)
			{
				return;
			}
			LONG lStartPage = CInt32(sStartPage);
			LONG lEndPage = CInt32(sEndPage);
			if (lEndPage < lStartPage)
			{
				return;
			}
			mp_oControl->GetPrinter().PrintRange(lStartPage, lEndPage);
		}
		CDialog::OnOK();
	}
}

void fPrintDialog::OnBnClickedCmdpreview()
{
	if (mp_oControl->GetPrinter().Initialize(GetStartDateCtrl(), GetEndDateCtrl(), numStartRow.GetPos32(), numEndRow.GetPos32()) == TRUE)
	{
		fPrintPreview oForm;
		oForm.mp_oParent = this;
		oForm.DoModal();
	}
}

//// C++ Specific

void fPrintDialog::OnDeltaposNummarginleft(NMHDR *pNMHDR, LRESULT *pResult)
{
	LPNMUPDOWN pNMUpDown = reinterpret_cast<LPNMUPDOWN>(pNMHDR);

	CString sWindowText = CStr((double)numMarginLeft.GetPos32() / 100, 2);
	txtMarginLeft.SetWindowTextW(sWindowText);

	*pResult = 0;
}

void fPrintDialog::OnDeltaposNummargintop(NMHDR *pNMHDR, LRESULT *pResult)
{
	LPNMUPDOWN pNMUpDown = reinterpret_cast<LPNMUPDOWN>(pNMHDR);

	CString sWindowText = CStr((double)numMarginTop.GetPos32() / 100, 2);
	txtMarginTop.SetWindowTextW(sWindowText);

	*pResult = 0;
}

void fPrintDialog::OnDeltaposNummarginright(NMHDR *pNMHDR, LRESULT *pResult)
{
	LPNMUPDOWN pNMUpDown = reinterpret_cast<LPNMUPDOWN>(pNMHDR);

	CString sWindowText = CStr((double)numMarginRight.GetPos32() / 100, 2);
	txtMarginRight.SetWindowTextW(sWindowText);

	*pResult = 0;
}

void fPrintDialog::OnDeltaposNummarginbottom(NMHDR *pNMHDR, LRESULT *pResult)
{
	LPNMUPDOWN pNMUpDown = reinterpret_cast<LPNMUPDOWN>(pNMHDR);

	CString sWindowText = CStr((double)numMarginBottom.GetPos32() / 100, 2);
	txtMarginBottom.SetWindowTextW(sWindowText);

	*pResult = 0;
}