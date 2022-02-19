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
#include "ActiveGanttVC.h"
#include "ActiveGanttVCCtl.h"
#include "clsXML.h"
#include "clsPrinter.h"
#include "winspool.h"
#include "math.h"



IMPLEMENT_DYNCREATE(clsPrinter, CCmdTarget)

//{26F0D4B7-133A-4FBE-949E-2ED42CAC7298}
static const IID IID_IclsPrinter = { 0x26F0D4B7, 0x133A, 0x4FBE, { 0x94, 0x9E, 0x2E, 0xD4, 0x2C, 0xAC, 0x72, 0x98} };

//{E49E898B-1327-4689-8051-BE52C4EF2BC4}
IMPLEMENT_OLECREATE_FLAGS(clsPrinter, "AGVC.clsPrinter", afxRegApartmentThreading, 0xE49E898B, 0x1327, 0x4689, 0x80, 0x51, 0xBE, 0x52, 0xC4, 0xEF, 0x2B, 0xC4)

BEGIN_DISPATCH_MAP(clsPrinter, CCmdTarget)
	DISP_PROPERTY_EX_ID(clsPrinter, "PrinterName", 1, odl_GetPrinterName, odl_SetPrinterName, VT_BSTR)
	DISP_PROPERTY_EX_ID(clsPrinter, "PrinterResolution", 2, odl_GetPrinterResolution, odl_SetPrinterResolution, VT_I4)
	DISP_PROPERTY_EX_ID(clsPrinter, "HorizontalDPI", 3, odl_GetHorizontalDPI, odl_SetHorizontalDPI, VT_I4)
	DISP_PROPERTY_EX_ID(clsPrinter, "VerticalDPI", 4, odl_GetVerticalDPI, odl_SetVerticalDPI, VT_I4)
	DISP_PROPERTY_EX_ID(clsPrinter, "PaperType", 5, odl_GetPaperType, odl_SetPaperType, VT_I4)
	DISP_PROPERTY_EX_ID(clsPrinter, "Orientation", 6, odl_GetOrientation, odl_SetOrientation, VT_I4)
	DISP_PROPERTY_EX_ID(clsPrinter, "MarginLeft", 7, odl_GetMarginLeft, odl_SetMarginLeft, VT_I4)
	DISP_PROPERTY_EX_ID(clsPrinter, "MarginTop", 8, odl_GetMarginTop, odl_SetMarginTop, VT_I4)
	DISP_PROPERTY_EX_ID(clsPrinter, "MarginRight", 9, odl_GetMarginRight, odl_SetMarginRight, VT_I4)
	DISP_PROPERTY_EX_ID(clsPrinter, "MarginBottom", 10, odl_GetMarginBottom, odl_SetMarginBottom, VT_I4)
	DISP_PROPERTY_EX_ID(clsPrinter, "Scale", 11, odl_GetScale, odl_SetScale, VT_R4)
	DISP_PROPERTY_EX_ID(clsPrinter, "HeadingsInEveryPage", 12, odl_GetHeadingsInEveryPage, odl_SetHeadingsInEveryPage, VT_BOOL)
	DISP_PROPERTY_EX_ID(clsPrinter, "ColumnsInEveryPage", 13, odl_GetColumnsInEveryPage, odl_SetColumnsInEveryPage, VT_BOOL)
	DISP_PROPERTY_EX_ID(clsPrinter, "Columns", 14, odl_GetColumns, odl_SetColumns, VT_I4)
	DISP_PROPERTY_EX_ID(clsPrinter, "KeepRowsTogether", 15, odl_GetKeepRowsTogether, odl_SetKeepRowsTogether, VT_BOOL)
	DISP_PROPERTY_EX_ID(clsPrinter, "KeepColumnsTogether", 16, odl_GetKeepColumnsTogether, odl_SetKeepColumnsTogether, VT_BOOL)
	DISP_PROPERTY_EX_ID(clsPrinter, "KeepTimeIntervalsTogether", 17, odl_GetKeepTimeIntervalsTogether, odl_SetKeepTimeIntervalsTogether, VT_BOOL)
	DISP_PROPERTY_EX_ID(clsPrinter, "ShowAllColumns", 18, odl_GetShowAllColumns, odl_SetShowAllColumns, VT_BOOL)
	DISP_PROPERTY_EX_ID(clsPrinter, "Interval", 19, odl_GetInterval, odl_SetInterval, VT_I4)
	DISP_PROPERTY_EX_ID(clsPrinter, "Factor", 20, odl_GetFactor, odl_SetFactor, VT_I4)
	DISP_PROPERTY_EX_ID(clsPrinter, "Pages", 21, odl_GetPages, SetNotSupported, VT_I4)
	DISP_PROPERTY_EX_ID(clsPrinter, "XAxisPages", 22, odl_GetXAxisPages, SetNotSupported, VT_I4)
	DISP_PROPERTY_EX_ID(clsPrinter, "YAxisPages", 23, odl_GetYAxisPages, SetNotSupported, VT_I4)
	DISP_PROPERTY_EX_ID(clsPrinter, "PrintAreaEndDate", 24, odl_GetPrintAreaEndDate, SetNotSupported, VT_DATE)
	DISP_PROPERTY_EX_ID(clsPrinter, "PrintAreaStartDate", 25, odl_GetPrintAreaStartDate, SetNotSupported, VT_DATE)
	DISP_PROPERTY_EX_ID(clsPrinter, "StartRow", 26, odl_GetStartRow, SetNotSupported, VT_I4)
	DISP_PROPERTY_EX_ID(clsPrinter, "EndRow", 27, odl_GetEndRow, SetNotSupported, VT_I4)
	DISP_PROPERTY_EX_ID(clsPrinter, "PrintAreaWidth", 28, odl_GetPrintAreaWidth, SetNotSupported, VT_I4)
	DISP_PROPERTY_EX_ID(clsPrinter, "PrintAreaHeight", 29, odl_GetPrintAreaHeight, SetNotSupported, VT_I4)
	DISP_FUNCTION_ID(clsPrinter, "Clear", 30, odl_Clear, VT_EMPTY, VTS_NONE)
	DISP_FUNCTION_ID(clsPrinter, "GetPagePosition", 31, odl_GetPagePosition, VT_EMPTY, VTS_I4 VTS_PI4 VTS_PI4) 
	DISP_FUNCTION_ID(clsPrinter, "GetPageNumber", 32, odl_GetPageNumber, VT_I4, VTS_I4 VTS_I4) 
	DISP_FUNCTION_ID(clsPrinter, "Initialize", 33, odl_Initialize, VT_BOOL, VTS_DISPATCH VTS_DISPATCH VTS_I4 VTS_I4) 
	DISP_FUNCTION_ID(clsPrinter, "Terminate", 34, odl_Terminate, VT_EMPTY, VTS_NONE) 
	DISP_FUNCTION_ID(clsPrinter, "Calculate", 35, odl_Calculate, VT_EMPTY, VTS_NONE) 
	DISP_FUNCTION_ID(clsPrinter, "PrintAll", 36, odl_PrintAll, VT_EMPTY, VTS_NONE) 
	DISP_FUNCTION_ID(clsPrinter, "PrintRange", 37, odl_PrintRange, VT_EMPTY, VTS_I4 VTS_I4) 
	DISP_FUNCTION_ID(clsPrinter, "PrintPage", 38, odl_PrintPage, VT_EMPTY, VTS_I4) 
	DISP_FUNCTION_ID(clsPrinter, "PreviewPage", 39, odl_PreviewPage, VT_EMPTY, VTS_HANDLE VTS_I4 VTS_R4 VTS_I4 VTS_I4)
	DISP_PROPERTY_EX_ID(clsPrinter, "PScale", 40, odl_GetScale, odl_SetScale, VT_R4)
END_DISPATCH_MAP()

BEGIN_INTERFACE_MAP(clsPrinter, CCmdTarget)
	INTERFACE_PART(clsPrinter, IID_IclsPrinter, Dispatch)
END_INTERFACE_MAP()

BEGIN_MESSAGE_MAP(clsPrinter, CCmdTarget)
END_MESSAGE_MAP()

clsPrinter::clsPrinter()
{
	AfxMessageBox(_T("Invalid constructor has been invoked for clsPrinter. Please inform support@sourcecodestore.com of this error"), MB_OK | MB_ICONSTOP | MB_APPLMODAL);
}

clsPrinter::clsPrinter(CActiveGanttVCCtl* oControl)
{
	EnableAutomation();
	AfxOleLockApp();
	mp_oControl = oControl;
	Clear();
}

clsPrinter::~clsPrinter()
{
	AfxOleUnlockApp();
}

void clsPrinter::OnFinalRelease()
{
	CCmdTarget::OnFinalRelease();
}

void clsPrinter::Finalize(void)
{
}


// Properties:

CString clsPrinter::GetPrinterName(void)
{
	return mp_sPrinterName;
}

void clsPrinter::SetPrinterName(CString value)
{
	mp_sPrinterName = value;
	mp_QueryPrinter();
}

LONG clsPrinter::GetPrinterResolution(void)
{
	return mp_yPrinterResolution;
}

void clsPrinter::SetPrinterResolution(LONG value)
{
	mp_yPrinterResolution = value;
}

LONG clsPrinter::GetHorizontalDPI(void)
{
	return mp_lHorizontalDPI;
}

void clsPrinter::SetHorizontalDPI(LONG value)
{
	mp_lHorizontalDPI = value;
}

LONG clsPrinter::GetVerticalDPI(void)
{
	return mp_lVerticalDPI;
}

void clsPrinter::SetVerticalDPI(LONG value)
{
	mp_lVerticalDPI = value;
}

LONG clsPrinter::GetPaperType(void)
{
	return mp_yPaperType;
}

void clsPrinter::SetPaperType(LONG value)
{
	mp_yPaperType = value;

	if (mp_sPrinterName.GetLength() == 0)
	{
		mp_sPrinterName = mp_sGetDefaultPrinterName();
	}
	LPDEVMODE oDEVMODE;
	HANDLE hPrinter;
	DWORD dwNeeded, dwRet;
	LPWSTR sPrinterName = mp_sPrinterName.GetBuffer();
	OpenPrinter(sPrinterName, &hPrinter, NULL);
	dwNeeded = DocumentProperties(NULL, hPrinter, sPrinterName, NULL, NULL, 0);
	oDEVMODE = (LPDEVMODE)malloc(dwNeeded);
	dwRet = DocumentProperties(NULL, hPrinter, sPrinterName, oDEVMODE, NULL, DM_OUT_BUFFER);
	//Change Paper Type
	oDEVMODE->dmPaperSize = (short)mp_yPaperType;
	dwRet = DocumentProperties(NULL, hPrinter, sPrinterName, oDEVMODE, oDEVMODE, DM_IN_BUFFER |  DM_OUT_BUFFER);
	ClosePrinter(hPrinter);
	free(oDEVMODE);
	mp_sPrinterName.ReleaseBuffer();

	mp_QueryPrinter();

}

LONG clsPrinter::GetOrientation(void)
{
	return mp_yOrientation;
}

void clsPrinter::SetOrientation(LONG value)
{
	mp_yOrientation = value;
}

LONG clsPrinter::GetMarginLeft(void)
{
	return mp_lMarginLeft;
}

void clsPrinter::SetMarginLeft(LONG value)
{
	mp_lMarginLeft = value;
}

LONG clsPrinter::GetMarginTop(void)
{
	return mp_lMarginTop;
}

void clsPrinter::SetMarginTop(LONG value)
{
	mp_lMarginTop = value;
}

LONG clsPrinter::GetMarginRight(void)
{
	return mp_lMarginRight;
}

void clsPrinter::SetMarginRight(LONG value)
{
	mp_lMarginRight = value;
}

LONG clsPrinter::GetMarginBottom(void)
{
	return mp_lMarginBottom;
}

void clsPrinter::SetMarginBottom(LONG value)
{
	mp_lMarginBottom = value;
}

FLOAT clsPrinter::GetScale(void)
{
	return mp_fScale;
}

void clsPrinter::SetScale(FLOAT value)
{
	mp_fScale = value;
}

BOOL clsPrinter::GetHeadingsInEveryPage(void)
{
	return mp_bHeadingsInEveryPage;
}

void clsPrinter::SetHeadingsInEveryPage(BOOL value)
{
	mp_bHeadingsInEveryPage = value;
	mp_Calculate();
}

BOOL clsPrinter::GetColumnsInEveryPage(void)
{
	return mp_bColumnsInEveryPage;
}

void clsPrinter::SetColumnsInEveryPage(BOOL value)
{
	mp_bColumnsInEveryPage = value;
	mp_Calculate();
}

LONG clsPrinter::GetColumns(void)
{
	return mp_lColumns;
}

void clsPrinter::SetColumns(LONG value)
{
	if (value < 0)
	{
		value = 0;
	}
	mp_lColumns = value;
	mp_Calculate();
}

BOOL clsPrinter::GetKeepRowsTogether(void)
{
	return mp_bKeepRowsTogether;
}

void clsPrinter::SetKeepRowsTogether(BOOL value)
{
	mp_bKeepRowsTogether = value;
	mp_Calculate();
}

BOOL clsPrinter::GetKeepColumnsTogether(void)
{
	return mp_bKeepColumnsTogether;
}

void clsPrinter::SetKeepColumnsTogether(BOOL value)
{
	mp_bKeepColumnsTogether = value;
	mp_Calculate();
}

BOOL clsPrinter::GetKeepTimeIntervalsTogether(void)
{
	return mp_bKeepTimeIntervalsTogether;
}

void clsPrinter::SetKeepTimeIntervalsTogether(BOOL value)
{
	mp_bKeepTimeIntervalsTogether = value;
	mp_Calculate();
}

BOOL clsPrinter::GetShowAllColumns(void)
{
	return mp_bShowAllColumns;
}

void clsPrinter::SetShowAllColumns(BOOL value)
{
	mp_bShowAllColumns = value;
	if (mp_bShowAllColumns == TRUE) 
	{
		if (mp_lSplitterPositionBuffer == -1) 
		{
			mp_lSplitterPositionBuffer = mp_oControl->GetSplitter()->GetPosition();
		}
		mp_oControl->GetSplitter()->SetPosition(mp_oControl->GetColumns()->GetWidth() + mp_oControl->GetSplitter()->GetOffset());
	} 
	else if (mp_bShowAllColumns == FALSE) 
	{
		if (mp_lSplitterPositionBuffer != -1) 
		{
			mp_oControl->GetSplitter()->SetPosition(mp_lSplitterPositionBuffer);
			mp_lSplitterPositionBuffer = -1;
		}
	}
	mp_Calculate();
}

LONG clsPrinter::GetInterval(void)
{
	return mp_yInterval;
	mp_Calculate();
}

void clsPrinter::SetInterval(LONG value)
{
	mp_yInterval = value;
	mp_Calculate();
}

LONG clsPrinter::GetFactor(void)
{
	return mp_lFactor;
}

void clsPrinter::SetFactor(LONG value)
{
	if (value <= 0)
	{
		value = 1;
	}
	mp_lFactor = value;
	mp_Calculate();
}

LONG clsPrinter::GetPages(void)
{
	return mp_lPages;
}

LONG clsPrinter::GetXAxisPages(void)
{
	return mp_lXAxisPages;
}

LONG clsPrinter::GetYAxisPages(void)
{
	return mp_lYAxisPages;
}

DATE clsPrinter::GetPrintAreaEndDate(void)
{
	return mp_dtEndDate;
}

DATE clsPrinter::GetPrintAreaStartDate(void)
{
	return mp_dtStartDate;
}

LONG clsPrinter::GetStartRow(void)
{
	return mp_lStartRow;
}

LONG clsPrinter::GetEndRow(void)
{
	return mp_lEndRow;
}

LONG clsPrinter::GetPrintAreaWidth(void)
{
	return mp_oControl->clsG->GetCustomWidth();
}

LONG clsPrinter::GetPrintAreaHeight(void)
{
	return mp_oControl->clsG->GetCustomHeight();
}

// Methods:

void clsPrinter::Clear(void)
{

	mp_dtStartDate = (DATE)0;
	mp_dtEndDate = (DATE)0;
	mp_dtPrintStartDateBuffer = (DATE)0;
	mp_lSplitterPositionBuffer = -1;
	mp_lFirstVisibleRowBuffer = -1;
	mp_oView = NULL;
	mp_yOrientation = OT_PORTRAIT;
	mp_dPageWidth = 0;
	mp_dPageHeight = 0;
	mp_dActualPageWidth = 0;
	mp_dActualPageHeight = 0;
	mp_lPages = 0;
	mp_lPageNumber = 0;
	mp_bInitialized = FALSE;
	if (mp_oControl->GetRows()->GetCount() > 0) 
	{
		mp_lStartRow = 1;
		mp_lEndRow = mp_oControl->GetRows()->GetCount();
	} 
	else 
	{
		mp_lStartRow = -1;
		mp_lEndRow = -1;
	}
	mp_lXAxisPages = 0;
	mp_lYAxisPages = 0;
	mp_fScale = 1;
	mp_bKeepRowsTogether = TRUE;
	mp_bKeepTimeIntervalsTogether = FALSE;
	mp_yInterval = IL_MONTH;
	mp_lFactor = 1;
	mp_bKeepColumnsTogether = TRUE;
	mp_bShowAllColumns = TRUE;
	mp_bHeadingsInEveryPage = FALSE;
	mp_bColumnsInEveryPage = FALSE;
	mp_lColumns = 1;
	mp_aVertical.RemoveAll();
	mp_aHorizontal.RemoveAll();
	mp_yPrintingOperationType = PT_ALL;
	mp_lStartPage = 0;
	mp_lEndPage = 0;

	mp_lMarginLeft = 100;
	mp_lMarginTop = 100;
	mp_lMarginRight = 100;
	mp_lMarginBottom = 100;

	mp_QueryPrinter();

}

void clsPrinter::GetPagePosition(LONG PageNumber, LONG &Column, LONG &Row)
{
	if (PageNumber < 1 || PageNumber > GetPages()) 
	{
		mp_oControl->mp_ErrorReport(PRINTER_INVALID_PAGE, "Invalid PageNumber. PageNumber must be greater than 1 and less than the value of the Pages property.", "ActiveGanttVCCtl.clsPrinter.GetPagePosition");
		return;
	}
	mp_GetPagePosition(PageNumber, Column, Row);
	Column = Column + 1;
	Row = Row + 1;
}

LONG clsPrinter::GetPageNumber(LONG Column, LONG Row)
{
	if (Column < 1 || Column > mp_lXAxisPages) 
	{
		mp_oControl->mp_ErrorReport(PRINTER_INVALID_COLUMN, "Invalid column index. Index must be greater than 1 and less than the value of the XAxisPages property.", "ActiveGanttVCCtl.clsPrinter.GetPageNumber");
		return -1;
	}
	if (Row < 1 || Row > mp_lYAxisPages) 
	{
		mp_oControl->mp_ErrorReport(PRINTER_INVALID_COLUMN, "Invalid row index. Index must be greater than 1 and less than the value of the YAxisPages property.", "ActiveGanttVCCtl.clsPrinter.GetPageNumber");
		return -1;
	}
	if (Row == 1) 
	{
		return Column;
	} 
	else 
	{
		return ((Row - 1) * mp_lXAxisPages) + Column;
	}
}

BOOL clsPrinter::Initialize(DATE StartDate, DATE EndDate, LONG StartRow, LONG EndRow)
{
	mp_oView = mp_oControl->GetCurrentViewObject();
	mp_dtStartDate = StartDate;
	mp_dtEndDate = EndDate;

	mp_lStartRow = StartRow;
	mp_lEndRow = EndRow;

	if (mp_oControl->GetRows()->GetCount() == 0) 
	{
		mp_oControl->mp_ErrorReport(PRINTER_NO_ROWS, "No rows to print.", "ActiveGanttVCCtl.clsPrinter.Initialize");
		return FALSE;
	}

	if ((mp_lStartRow < 0 && mp_lEndRow < 0)) 
	{
		mp_lStartRow = 1;
		mp_lEndRow = mp_oControl->GetRows()->GetCount();
	}

	if (mp_lStartRow <= 0) 
	{
		mp_lStartRow = 1;
	}

	if ((mp_lEndRow > mp_oControl->GetRows()->GetCount()) || EndRow <= -1) 
	{
		mp_lEndRow = mp_oControl->GetRows()->GetCount();
	}

	if (mp_lEndRow < mp_lStartRow) 
	{
		mp_oControl->mp_ErrorReport(PRINTER_INVALID_ROW_RANGE, "Invalid Row Range. EndRow cannot be smaller than StartRow.", "ActiveGanttVCCtl.clsPrinter.Initialize");
		return FALSE;
	}

	if ((StartDate == EndDate) || (EndDate < StartDate)) 
	{
		mp_oControl->mp_ErrorReport(PRINTER_INVALID_DATE_RANGE, "Invalid Date Range. EndDate must be greater than StartDate.", "ActiveGanttVCCtl.clsPrinter.Initialize");
		return FALSE;
	}

	if (mp_CalculatePageDimensions() == FALSE)
    {
        return FALSE;
    }
    if (mp_CheckRows() == FALSE)
    {
        return FALSE;
    }
    if (mp_CheckColumns() == FALSE)
    {
        return FALSE;
    }
    if (mp_CheckTimeIntervals() == FALSE)
    {
        return FALSE;
    }

	if (mp_bShowAllColumns == TRUE) 
	{
		if (mp_lSplitterPositionBuffer == -1) 
		{
			mp_lSplitterPositionBuffer = mp_oControl->GetSplitter()->GetPosition();
		}
		mp_oControl->GetSplitter()->SetPosition(mp_oControl->GetColumns()->GetWidth() + mp_oControl->GetSplitter()->GetOffset());
	}

	if (mp_oControl->f_TimeLineScrollBar()->GetEnabled() == FALSE) 
	{
		mp_dtPrintStartDateBuffer = mp_oView->GetTimeLine()->GetStartDate();
		mp_oView->GetTimeLine()->Setf_StartDate(mp_dtStartDate);
	} 
	else 
	{
		mp_dtPrintStartDateBuffer = mp_oControl->f_TimeLineScrollBar()->GetStartDate();
		mp_oControl->f_TimeLineScrollBar()->SetStartDate(mp_dtStartDate);
	}
	if (mp_lFirstVisibleRowBuffer == -1)
    {
        mp_lFirstVisibleRowBuffer = mp_oView->GetClientArea()->GetFirstVisibleRow();    
    }
	mp_oView->GetClientArea()->SetFirstVisibleRow(1);

	mp_oControl->clsG->SetCustomWidth((mp_oControl->GetMathLib()->DateTimeDiff(mp_oView->GetInterval(), StartDate, EndDate) / mp_oView->GetFactor()) + mp_oControl->GetSplitter()->GetRight());
	mp_oControl->clsG->SetCustomHeight(mp_oControl->GetRows()->Height() + mp_oControl->GetCurrentViewObject()->GetTimeLine()->GetHeight());

	mp_bInitialized = TRUE;
	mp_Calculate();
	return TRUE;
}

void clsPrinter::Terminate(void)
{
	if (mp_bInitialized == TRUE) 
	{
		if (mp_oControl->f_TimeLineScrollBar()->GetEnabled() == FALSE) 
		{
			mp_oView->GetTimeLine()->Setf_StartDate(mp_dtPrintStartDateBuffer);
		}
		else 
		{
			mp_oControl->f_TimeLineScrollBar()->SetStartDate(mp_dtPrintStartDateBuffer);
		}
		if (mp_lSplitterPositionBuffer != -1) 
		{
			mp_oControl->GetSplitter()->SetPosition(mp_lSplitterPositionBuffer);
			mp_lSplitterPositionBuffer = -1;
		}
		if (mp_lFirstVisibleRowBuffer != -1)
		{
			mp_oView->GetClientArea()->SetFirstVisibleRow(mp_lFirstVisibleRowBuffer);
			mp_lFirstVisibleRowBuffer = -1;
		}
		mp_Calculate();
	}
}

void clsPrinter::Calculate(void)
{
	mp_Calculate();
}

void clsPrinter::PrintAll(void)
{
	if (mp_bInitialized == TRUE) 
	{
		mp_lStartPage = 1;
		mp_lEndPage = GetPages();
		mp_yPrintingOperationType = PT_ALL;
		LONG i;
		mp_StartDoc();
		for (i = mp_lStartPage; i <= mp_lEndPage; i++)
		{
			mp_StartPage();
			mp_lPageNumber = i;
			Gdiplus::Graphics oGraphics((HDC)mp_lhdcPrinter);
			mp_PrintPage(&oGraphics);
			mp_EndPage();
		}
		mp_EndDoc();
	}
}

void clsPrinter::PrintRange(LONG StartPage, LONG EndPage)
{
	if (mp_bInitialized == TRUE) 
	{
		if ((StartPage < 1) | (EndPage > GetPages())) 
		{
			mp_oControl->mp_ErrorReport(PRINTER_INVALID_PAGE, "Invalid PageNumber. StartPage and EndPage must be greater than 1 and less than the value of the Pages property.", "ActiveGanttVCCtl.clsPrinter.PrintRange");
			return;
		}
		if (EndPage < StartPage) 
		{
			mp_oControl->mp_ErrorReport(PRINTER_INVALID_PAGE_RANGE, "Invalid Page range. EndPage must be greater than or equal to StartPage.", "ActiveGanttVCCtl.clsPrinter.PrintRange");
			return;
		}
		mp_lStartPage = StartPage;
		mp_lEndPage = EndPage;
		mp_yPrintingOperationType = PT_RANGE;
		LONG i;
		mp_StartDoc();
		for (i = mp_lStartPage; i <= mp_lEndPage; i++)
		{
			mp_StartPage();
			mp_lPageNumber = i;
			Gdiplus::Graphics oGraphics((HDC)mp_lhdcPrinter);
			mp_PrintPage(&oGraphics);
			mp_EndPage();
		}
		mp_EndDoc();
	}
}

void clsPrinter::PrintPage(LONG PageNumber)
{
	if ((PageNumber < 1) | (PageNumber > GetPages())) 
	{
		mp_oControl->mp_ErrorReport(PRINTER_INVALID_PAGE, "Invalid PageNumber. PageNumber must be greater than 1 and less than the value of the Pages property.", "ActiveGanttVCCtl.clsPrinter.PrintRange");
		return;
	}
	if (mp_bInitialized == TRUE) 
	{
		mp_Calculate();
		mp_lPageNumber = PageNumber;
		mp_yPrintingOperationType = PT_SINGLEPAGE;
		mp_StartDoc();
		mp_StartPage();
		Gdiplus::Graphics oGraphics((HDC)mp_lhdcPrinter);
		mp_PrintPage(&oGraphics);
		mp_EndPage();
		mp_EndDoc();
	}
}

void clsPrinter::PreviewPage(OLE_HANDLE hdc, LONG PageNumber, FLOAT Scale, LONG X, LONG Y)
{
	if (Scale <= 0)
	{
		mp_oControl->mp_ErrorReport(PRINTER_INVALID_SCALE, _T("Invalid scale, scale must be greater than 0."), _T("ActiveGanttVCCtl.clsPrinter.PreviewPage"));
		return;
	}
	if ((PageNumber < 1) | (PageNumber > GetPages()))
	{
		mp_oControl->mp_ErrorReport(PRINTER_INVALID_PAGE, _T("Invalid PageNumber. PageNumber must be greater than 1 and less than the value of the Pages property."), _T("ActiveGanttVCCtl.clsPrinter.PreviewPage"));
		return;
	}
	Gdiplus::Graphics oGraphics((HDC)hdc);
	int lMarginLeft = CInt32(GetMarginLeft() * mp_fScaleInvMultp());
	int lMarginTop = CInt32(GetMarginTop() * mp_fScaleInvMultp());
	int X1 = CInt32(X * mp_fScaleInvMultp());
	int Y1 = CInt32(Y * mp_fScaleInvMultp());

	mp_oControl->clsG->SetCustomPrinting(TRUE);
	mp_oControl->mp_PositionScrollBars();
	mp_oControl->clsG->SetCustomDC(&oGraphics);

	oGraphics.ScaleTransform(Scale, Scale);
	oGraphics.TranslateTransform((1 / Scale * X) - X, (1 / Scale * Y) - Y);

	Gdiplus::SolidBrush oBrush(Color::White);
	if (mp_yOrientation == OT_PORTRAIT) 
	{
		oGraphics.FillRectangle(&oBrush, X, Y, CInt32(mp_dPageWidth), CInt32(mp_dPageHeight));
	} 
	else 
	{
		oGraphics.FillRectangle(&oBrush, X, Y, CInt32(mp_dPageHeight), CInt32(mp_dPageWidth));
	}

	mp_Draw(&oGraphics, PageNumber, lMarginLeft, lMarginTop, X1, Y1, FALSE);

	mp_oControl->clsG->SetCustomPrinting(FALSE);
	mp_oControl->mp_PositionScrollBars();
}

void clsPrinter::mp_PrintPage(Graphics* oGraphics)
{
		int lMarginLeft = CInt32(GetMarginLeft() * mp_fScaleInvMultp()) - CInt32(mp_lHardMarginX * mp_fScaleInvMultp());
		int lMarginTop = CInt32(GetMarginTop() * mp_fScaleInvMultp()) - CInt32(mp_lHardMarginY * mp_fScaleInvMultp());

		mp_oControl->clsG->SetCustomPrinting(TRUE);
		mp_oControl->mp_PositionScrollBars();
		mp_oControl->clsG->SetCustomDC(oGraphics);

		mp_Draw(oGraphics, mp_lPageNumber, lMarginLeft, lMarginTop, 0, 0, TRUE);

		mp_oControl->clsG->SetCustomPrinting(FALSE);
		mp_oControl->mp_PositionScrollBars();
}

void clsPrinter::mp_Draw(Graphics* oGraphics, LONG PageNumber, LONG lMarginLeft, LONG lMarginTop, LONG X1, LONG Y1, BOOL bToPrinter)
{
		LONG lColumn = -1;
		LONG lRow = -1;
		Gdiplus::GraphicsState oGraphicsStartState;
		Gdiplus::GraphicsState oGraphicsState;


		lMarginLeft = lMarginLeft + X1;
		lMarginTop = lMarginTop + Y1;

		oGraphicsStartState = oGraphics->Save();
		mp_GetPagePosition(PageNumber, lColumn, lRow);

		oGraphics->ScaleTransform(CReal(mp_fScale), CReal(mp_fScale));
		if (mp_bHeadingsInEveryPage == TRUE || lRow == 0) 
		{
			oGraphicsState = oGraphics->Save();
			oGraphics->TranslateTransform(CReal(-mp_aX(lColumn) + lMarginLeft + mp_lColumnsWidth(lColumn)), CReal(lMarginTop));
			mp_oControl->clsG->mp_lX1 = mp_aX(lColumn);
			mp_oControl->clsG->mp_lY1 = 0;
			mp_oControl->clsG->mp_lX2 = mp_aX(lColumn + 1);
			mp_oControl->clsG->mp_lY2 = mp_lTimeLineHeight();
			mp_oControl->f_Draw();
			oGraphics->Restore(oGraphicsState);
		}
		if (mp_bColumnsInEveryPage == TRUE && mp_lColumns > 0 && lColumn > 0) 
		{
			if (mp_bHeadingsInEveryPage == TRUE || lRow == 0) 
			{
				//// Column Headings
				oGraphicsState = oGraphics->Save();
				oGraphics->TranslateTransform(CReal(lMarginLeft), CReal(lMarginTop));
				mp_oControl->clsG->mp_lX1 = 0;
				mp_oControl->clsG->mp_lY1 = 0;
				mp_oControl->clsG->mp_lX2 = mp_lColumnsWidth(lColumn);
				mp_oControl->clsG->mp_lY2 = mp_lTimeLineHeight();
				//mp_oControl.clsG.ClipRegion(0, 0, mp_lColumnsWidth(lColumn), mp_lTimeLineHeight, False)
				//mp_oControl.clsG.mp_DrawItem(0, 0, mp_oControl.clsG.Width - 1, mp_oControl.clsG.Height - 1, "", "", False, mp_oControl.Image, 0, 0, mp_oControl.Style)
				//mp_oControl.Columns.Position()
				//mp_oControl.Columns.Draw()
				mp_oControl->f_Draw();
				oGraphics->Restore(oGraphicsState);
			}

			//// Columns
			oGraphicsState = oGraphics->Save();
			if (mp_bHeadingsInEveryPage == TRUE || lRow == 0) 
			{
				oGraphics->TranslateTransform(CReal(lMarginLeft), CReal(-mp_aY(lRow) + lMarginTop - mp_lRowsHeight()));
			} 
			else 
			{
				oGraphics->TranslateTransform(CReal(lMarginLeft), CReal(-mp_aY(lRow) + lMarginTop - mp_lTimeLineHeight() - mp_lRowsHeight()));
			}
			mp_oControl->clsG->mp_lX1 = 0;
			mp_oControl->clsG->mp_lY1 = mp_aY(lRow) + mp_lTimeLineHeight() + mp_lRowsHeight();
			mp_oControl->clsG->mp_lX2 = mp_lColumnsWidth(lColumn);
			mp_oControl->clsG->mp_lY2 = mp_aY(lRow + 1) + mp_lTimeLineHeight() + mp_lRowsHeight();
			mp_oControl->f_Draw();
			oGraphics->Restore(oGraphicsState);
		}

		if (mp_bHeadingsInEveryPage == TRUE || lRow == 0) 
		{
			oGraphics->TranslateTransform(CReal(-mp_aX(lColumn) + lMarginLeft + mp_lColumnsWidth(lColumn)), CReal(-mp_aY(lRow) + lMarginTop - mp_lRowsHeight()));
		} 
		else 
		{
			oGraphics->TranslateTransform(CReal(-mp_aX(lColumn) + lMarginLeft + mp_lColumnsWidth(lColumn)), CReal(-mp_aY(lRow) + lMarginTop - mp_lTimeLineHeight() - mp_lRowsHeight()));
		}
		mp_oControl->clsG->mp_lX1 = mp_aX(lColumn);
		mp_oControl->clsG->mp_lY1 = mp_aY(lRow) + mp_lTimeLineHeight() + mp_lRowsHeight();
		mp_oControl->clsG->mp_lX2 = mp_aX(lColumn + 1);
		mp_oControl->clsG->mp_lY2 = mp_aY(lRow + 1) + mp_lTimeLineHeight() + mp_lRowsHeight();
		mp_oControl->f_Draw();


		oGraphics->ResetTransform();


		if (bToPrinter == FALSE) 
		{
			oGraphics->Restore(oGraphicsStartState);
			oGraphics->TranslateTransform(X1 * mp_fScale, Y1 * mp_fScale);
		}
		else 
		{
			oGraphics->TranslateTransform(CReal(-mp_lHardMarginX), CReal(-mp_lHardMarginY));
		}
		mp_oControl->GetPagePrintEventArgs()->Clear();
		mp_oControl->GetPagePrintEventArgs()->SetX1(GetMarginLeft());
		mp_oControl->GetPagePrintEventArgs()->SetY1(GetMarginTop());
		mp_oControl->GetPagePrintEventArgs()->SetX2(GetMarginLeft() + CInt32((mp_aHorizontal[lColumn] + mp_lColumnsWidth(lColumn)) * mp_fScale));
		mp_oControl->GetPagePrintEventArgs()->SetY2(GetMarginTop() + CInt32((mp_aVertical[lRow] + mp_lTimeLineHeight(lRow)) * mp_fScale));

		if (mp_yOrientation == OT_PORTRAIT) 
		{
			mp_oControl->GetPagePrintEventArgs()->SetPageWidth(CInt32(mp_dPageWidth));
			mp_oControl->GetPagePrintEventArgs()->SetPageHeight(CInt32(mp_dPageHeight));
		}
		else if (mp_yOrientation == OT_LANDSCAPE) 
		{
			mp_oControl->GetPagePrintEventArgs()->SetPageWidth(CInt32(mp_dPageHeight));
			mp_oControl->GetPagePrintEventArgs()->SetPageHeight(CInt32(mp_dPageWidth));
		}

		mp_oControl->GetPagePrintEventArgs()->SetActualPageHeight(mp_aVertical[lRow]);
		mp_oControl->GetPagePrintEventArgs()->SetActualPageWidth(mp_aHorizontal[lColumn]);

		mp_oControl->GetPagePrintEventArgs()->SetPageNumber(PageNumber);

		mp_oControl->clsG->mp_lX1 = 0;
		mp_oControl->clsG->mp_lY1 = 0;
		mp_oControl->clsG->mp_lX2 = mp_oControl->clsG->GetCustomWidth();
		mp_oControl->clsG->mp_lY2 = mp_oControl->clsG->GetCustomHeight();

		oGraphics->ResetClip();

		mp_oControl->FirePagePrint();
}

LONG clsPrinter::mp_lTimeLineHeight(LONG lRow)
{
	if (mp_bHeadingsInEveryPage == TRUE || lRow == 0) 
	{
		return mp_oControl->GetCurrentViewObject()->GetTimeLine()->GetHeight() + mp_oControl->Getmt_BorderThickness();
	}
	else 
	{
		return 0;
	}
}

LONG clsPrinter::mp_lRowsHeight(void)
{
	return mp_oControl->GetRows()->Height(mp_lStartRow - 1);
}

LONG clsPrinter::mp_lTimeLineHeight(void)
{
	return mp_oControl->GetCurrentViewObject()->GetTimeLine()->GetHeight() + mp_oControl->Getmt_BorderThickness();
}

LONG clsPrinter::mp_aX(LONG lIndex)
{
	if (lIndex == 0) 
	{
		return 0;
	} 
	else 
	{
		int lReturn = 0;
		int i = 0;
		for (i = 0; i <= lIndex - 1; i++) 
		{
			lReturn = lReturn + mp_aHorizontal[i];
		}
		return lReturn;
	}
}

LONG clsPrinter::mp_aY(LONG lIndex)
{
	if (lIndex == 0) 
	{
		return 0;
	} 
	else 
	{
		int lReturn = 0;
		int i = 0;
		for (i = 0; i <= lIndex - 1; i++) 
		{
			lReturn = lReturn + mp_aVertical[i];
		}
		return lReturn;
	}
}

void clsPrinter::mp_QueryPrinter(void)
{
	if (mp_sPrinterName.GetLength() == 0)
	{
		mp_sPrinterName = mp_sGetDefaultPrinterName();
	}

	LPDEVMODE oDEVMODE;
	HANDLE hPrinter;
	DWORD dwNeeded, dwRet;
	LPWSTR sPrinterName = mp_sPrinterName.GetBuffer();

	OpenPrinter(sPrinterName, &hPrinter, NULL);
	dwNeeded = DocumentProperties(NULL, hPrinter, sPrinterName, NULL, NULL, 0);
	oDEVMODE = (LPDEVMODE)malloc(dwNeeded);
	dwRet = DocumentProperties(NULL, hPrinter, sPrinterName, oDEVMODE, NULL, DM_OUT_BUFFER);

	//Paper Type
	mp_yPaperType = oDEVMODE->dmPaperSize;

	// Orientation
	if (oDEVMODE->dmOrientation == DMORIENT_PORTRAIT)
	{
		mp_yOrientation = OT_PORTRAIT;
	}
	else if (oDEVMODE->dmOrientation == DMORIENT_LANDSCAPE)
	{
		mp_yOrientation = OT_LANDSCAPE;
	}

	// PrinterResolution (HorizontalDPI/VerticalDPI)
	if (oDEVMODE->dmPrintQuality >= -1)
	{
		mp_yPrinterResolution = PR_CUSTOM;
		mp_lHorizontalDPI = oDEVMODE->dmPrintQuality;
		mp_lVerticalDPI = oDEVMODE->dmYResolution;
	}
	else
	{
		switch (oDEVMODE->dmPrintQuality)
		{
		case DMRES_HIGH:
			mp_yPrinterResolution = PR_HIGH;
			break;
		case DMRES_MEDIUM:
			mp_yPrinterResolution = PR_MEDIUM;
			break;
		case DMRES_LOW:
			mp_yPrinterResolution = PR_LOW;
			break;
		case DMRES_DRAFT:
			mp_yPrinterResolution = PR_DRAFT;
			break;
		}
	}

	mp_dPageWidth = (oDEVMODE->dmPaperWidth / 2.54);
	mp_dPageHeight = (oDEVMODE->dmPaperLength / 2.54);


	ClosePrinter(hPrinter);
	free(oDEVMODE);
	mp_sPrinterName.ReleaseBuffer();

	//Hard Margins:
	HDC lhdc = CreateDC(_T("WINSPOOL"), mp_sPrinterName, NULL, NULL);
	mp_lHardMarginX = (float)GetDeviceCaps(lhdc, PHYSICALOFFSETX) / (float)GetDeviceCaps(lhdc, LOGPIXELSX) * 100;
	mp_lHardMarginY = (float)GetDeviceCaps(lhdc, PHYSICALOFFSETY) / (float)GetDeviceCaps(lhdc, LOGPIXELSY) * 100;
	DeleteDC(lhdc);

	mp_Calculate();

}

void clsPrinter::mp_Calculate(void)
{
	mp_CalculatePageDimensions();
	if (mp_bInitialized == TRUE) 
	{
		mp_CalculateColumns();
		mp_CalculateRows();
		mp_lPages = mp_lXAxisPages * mp_lYAxisPages;
	}
	else 
	{
		mp_oControl->clsG->SetCustomWidth(0);
		mp_oControl->clsG->SetCustomHeight(0);
		mp_dtStartDate = (DATE)0;
		mp_dtEndDate = (DATE)0;
		mp_lXAxisPages = 0;
		mp_lYAxisPages = 0;
		mp_lPages = 0;
	}
}

LONG clsPrinter::mp_lActualPageHeight(LONG lRow)
{
	int lTimeLineHeight = mp_oControl->GetCurrentViewObject()->GetTimeLine()->GetHeight() + mp_oControl->Getmt_BorderThickness();
	if (mp_bHeadingsInEveryPage == TRUE || lRow == 0) 
	{
		return CInt32((mp_dActualPageHeight - lTimeLineHeight) * mp_fScaleInvMultp());
	}
	else 
	{
		return CInt32(mp_dActualPageHeight * mp_fScaleInvMultp());
	}
}

void clsPrinter::mp_CalculateRows(void)
{
	mp_aVertical.RemoveAll();
	if (mp_bKeepRowsTogether == TRUE) 
	{
		int lRow = 0;
		int YBuff = 0;
		lRow = mp_lStartRow;
		while (lRow <= mp_lEndRow) 
		{
			if ((YBuff + mp_oControl->GetRows()->Item(CStr(lRow))->GetHeight() + 1) < mp_lActualPageHeight(mp_aVertical.GetCount())) 
			{
				YBuff = YBuff + mp_oControl->GetRows()->Item(CStr(lRow))->GetHeight() + 1;
				lRow = lRow + 1;
			}
			else 
			{
				mp_aVertical.Add(YBuff);
				YBuff = 0;
			}
		}
		if (YBuff > 0) 
		{
			mp_aVertical.Add(YBuff);
			YBuff = 0;
		}
		mp_lYAxisPages = mp_aVertical.GetCount();
	}
	else if (mp_bKeepRowsTogether == FALSE) 
	{
		int lRowsHeight = mp_oControl->GetRows()->Height(mp_lEndRow) - mp_oControl->GetRows()->Height(mp_lStartRow);
		int yBuff = 0;
		while (yBuff <= lRowsHeight) 
		{
			int lActualPageHeight = mp_lActualPageHeight(mp_aVertical.GetCount());
			yBuff = yBuff + lActualPageHeight;
			if (lActualPageHeight + mp_aY(mp_aVertical.GetCount()) < lRowsHeight) 
			{
				mp_aVertical.Add(lActualPageHeight);
			}
			else 
			{
				int lRowsHeight1 = 0;
				if (mp_lEndRow < mp_oControl->GetRows()->GetCount()) 
				{
					lRowsHeight1 = mp_oControl->GetRows()->Height(mp_lEndRow + 1) - mp_oControl->GetRows()->Height(mp_lStartRow);
				}
				else 
				{
					lRowsHeight1 = mp_oControl->GetRows()->Height() - mp_oControl->GetRows()->Height(mp_lStartRow);
				}
				int lAdd = lRowsHeight1 - mp_aY(mp_aVertical.GetCount());
				mp_aVertical.Add(lAdd);
			}

		}
		mp_lYAxisPages = mp_aVertical.GetCount();
	}
}

void clsPrinter::mp_CalculateColumns(void)
{
	int lColumn = 0;
	mp_aHorizontal.RemoveAll();
	if (mp_bKeepColumnsTogether == TRUE) 
	{
		int XBuff = 0;
		lColumn = 1;
		if (mp_bShowAllColumns == TRUE) 
		{
			while (lColumn <= mp_oControl->GetColumns()->GetCount()) 
			{
				int lXBuffColWidth1 = XBuff + mp_oControl->GetColumns()->Item(CStr(lColumn))->GetWidth() + 1;
				int lAPW = mp_lActualPageWidth(mp_aHorizontal.GetCount());
				if (lAPW < 0)
				{
					lAPW = mp_lActualPageWidth(mp_aHorizontal.GetCount());
				}

				if ((XBuff + mp_oControl->GetColumns()->Item(CStr(lColumn))->GetWidth() + 1) < mp_lActualPageWidth(mp_aHorizontal.GetCount())) 
				{
					XBuff = XBuff + mp_oControl->GetColumns()->Item(CStr(lColumn))->GetWidth();
					lColumn = lColumn + 1;
				}
				else 
				{
					int lAdd = XBuff - mp_lColumnsWidth(mp_aHorizontal.GetCount());
					mp_aHorizontal.Add(lAdd);
					XBuff = 0;
				}
			}
		} 
		else if (mp_bShowAllColumns == FALSE) 
		{
			XBuff = mp_oControl->GetSplitter()->GetRight();
		}
		if (mp_bKeepTimeIntervalsTogether == TRUE) 
		{
			int lCustomWidth = 0;
			int XDelta = 0;
			DATE dtBuff;
			mp_dtStartDate = mp_oControl->GetMathLib()->RoundDate(mp_yInterval, mp_lFactor, mp_dtStartDate);
			mp_dtEndDate = mp_oControl->GetMathLib()->RoundDate(mp_yInterval, mp_lFactor, mp_dtEndDate);
			mp_oControl->clsG->SetCustomWidth(CInt32((mp_oControl->GetMathLib()->DateTimeDiff(mp_oView->GetInterval(), mp_dtStartDate, mp_dtEndDate) / mp_oView->GetFactor()) + mp_oControl->GetSplitter()->GetRight()));
			lCustomWidth = mp_oControl->clsG->GetCustomWidth();
			dtBuff = mp_dtStartDate;
			while ((dtBuff < mp_dtEndDate)) 
			{
				XDelta = mp_oControl->GetMathLib()->GetXCoordinateFromDate(mp_oControl->GetMathLib()->DateTimeAdd(mp_yInterval, mp_lFactor, dtBuff)) - mp_oControl->GetMathLib()->GetXCoordinateFromDate(dtBuff);
				if (((XBuff + XDelta) < mp_lActualPageWidth(mp_aHorizontal.GetCount()))) 
				{
					XBuff = XBuff + XDelta;
					dtBuff = mp_oControl->GetMathLib()->DateTimeAdd(mp_yInterval, mp_lFactor, dtBuff);
				}
				else 
				{
					mp_aHorizontal.Add(XBuff);
					XBuff = 0;
				}
			}
			if (XBuff > 0) 
			{
				mp_aHorizontal.Add(XBuff);
				XBuff = 0;
			}
		}
		else if (mp_bKeepTimeIntervalsTogether == FALSE) 
		{
			if (mp_bColumnsInEveryPage == TRUE && mp_lColumns > 0) 
			{
				mp_lXAxisPages = CInt32(ceil((double)mp_oControl->clsG->GetCustomWidth() - (double)mp_lColumnsWidth(-1)) / (double)mp_lActualPageWidth(-1));
			}
			else if (mp_bColumnsInEveryPage == FALSE || mp_lColumns == 0) 
			{
				mp_lXAxisPages = CInt32(ceil((double)mp_oControl->clsG->GetCustomWidth() / (double)mp_lActualPageWidth()));
			}
			for (lColumn = 1; lColumn <= mp_lXAxisPages; lColumn++) 
			{
				int lAdd = CInt32((mp_dActualPageWidth * mp_fScaleInvMultp()) - mp_lColumnsWidth(lColumn - 1));
				mp_aHorizontal.Add(lAdd);
			}
		}
		mp_lXAxisPages = mp_aHorizontal.GetCount();
	}
	else if (mp_bKeepColumnsTogether == FALSE) 
	{
		if (mp_bColumnsInEveryPage == TRUE && mp_lColumns > 0) 
		{
			mp_lXAxisPages = CInt32(ceil(((double)mp_oControl->clsG->GetCustomWidth() - (double)mp_lColumnsWidth(-1)) / (double)mp_lActualPageWidth(-1)));
		} 
		else if (mp_bColumnsInEveryPage == FALSE || mp_lColumns == 0) 
		{
			mp_lXAxisPages = CInt32(ceil((double)mp_oControl->clsG->GetCustomWidth() / (double)mp_lActualPageWidth()));
		}
		for (lColumn = 1; lColumn <= mp_lXAxisPages; lColumn++) 
		{
			int lAdd = CInt32((mp_dActualPageWidth * mp_fScaleInvMultp()) - mp_lColumnsWidth(lColumn - 1));
			mp_aHorizontal.Add(lAdd);
		}
	}
}

FLOAT clsPrinter::mp_fScaleInvMultp(void)
{
	return 1 / mp_fScale;
}

LONG clsPrinter::mp_lActualPageWidth(void)
{
	if (mp_bColumnsInEveryPage == FALSE || mp_lColumns == 0) 
	{
		return CInt32(mp_dActualPageWidth * mp_fScaleInvMultp());
	} 
	else 
	{
		ASSERT(false);
		return 0;
	}
}

LONG clsPrinter::mp_lActualPageWidth(LONG lColumn)
{
	if (mp_bColumnsInEveryPage == TRUE && mp_lColumns > 0) 
	{
		return CInt32((mp_dActualPageWidth - mp_lColumnsWidth(lColumn)) * mp_fScaleInvMultp());
	} 
	else if (lColumn == -1) 
	{
		return CInt32((mp_dActualPageWidth - mp_lColumnsWidth(-1)) * mp_fScaleInvMultp());
	} 
	else if (mp_bColumnsInEveryPage == FALSE || mp_lColumns == 0) 
	{
		return CInt32(mp_dActualPageWidth * mp_fScaleInvMultp());
	} 
	else 
	{
		ASSERT(false);
		return 0;
	}
}

LONG clsPrinter::mp_lColumnsWidth(LONG lColumn)
{
	int lReturn = 0;
	if (lColumn == 0) 
	{
		return 0;
	}
	if (mp_bColumnsInEveryPage == TRUE && mp_lColumns > 0 && mp_bShowAllColumns == FALSE) 
	{
		lReturn = mp_oControl->GetSplitter()->GetRight() - mp_oControl->Getmt_BorderThickness();
	}
	else if (mp_bColumnsInEveryPage == TRUE && mp_lColumns > 0 && mp_bShowAllColumns == TRUE) 
	{
		int i = 0;
		for (i = 1; i <= mp_lColumns; i++) 
		{
			if (mp_oControl->GetColumns()->Item(CStr(i))->GetVisible() == TRUE) 
			{
				lReturn = lReturn + mp_oControl->GetColumns()->Item(CStr(i))->GetWidth();
			}
		}
	}
	return lReturn + mp_oControl->Getmt_BorderThickness();
}

void clsPrinter::mp_GetPagePosition(LONG lPageNumber, LONG &lColumn, LONG &lRow)
{
	lRow = CInt32(ceil((double)lPageNumber / (double)GetXAxisPages())) - 1;
	lColumn = lPageNumber - (lRow * GetXAxisPages()) - 1;
}


//Check


BOOL clsPrinter::mp_CalculatePageDimensions(void)
{
    if (mp_yOrientation == OT_PORTRAIT)
    {
        mp_dActualPageWidth = mp_dPageWidth - (GetMarginLeft() + GetMarginRight());
        mp_dActualPageHeight = mp_dPageHeight - (GetMarginTop() + GetMarginBottom());
    }
    else if (mp_yOrientation == OT_LANDSCAPE)
    {
        mp_dActualPageWidth = mp_dPageHeight - (GetMarginLeft() + GetMarginRight());
        mp_dActualPageHeight = mp_dPageWidth - (GetMarginTop() + GetMarginBottom());
    }
    if (mp_dActualPageWidth <= 0 || mp_dActualPageHeight <= 0)
    {
        mp_oControl->mp_ErrorReport(PRINTER_INVALID_SPECS_MARGINS, "Invalid printing specifications, page printing area cannot be 0. Reduce the margins.", "ActiveGanttCSNCtl.clsPrinter.Initialize");
        return FALSE;
    }
    else
    {
        return TRUE;
    }
}

BOOL clsPrinter::mp_CheckRows(void)
{
    int i = 0;
    int lActualPageHeight = CInt32(mp_dActualPageHeight * mp_fScaleInvMultp()) - CInt32(mp_oControl->GetCurrentViewObject()->GetTimeLine()->GetHeight() * mp_fScaleInvMultp());
    if (mp_bKeepRowsTogether == TRUE)
    {
        for (i = 1; i <= mp_oControl->GetRows()->GetCount(); i++)
        {
            if ((mp_oControl->GetRows()->Item(CStr(i))->GetHeight() + 1) > lActualPageHeight)
            {
                mp_oControl->mp_ErrorReport(PRINTER_INVALID_SPECS_HEIGHT, "Invalid printing specifications, reduce the top or bottom margin, set the KeepRowsTogether property to false or reduce the scale.", "ActiveGanttVCCtl.clsPrinter.Initialize");
                return FALSE;
            }
        }
    }
    return TRUE;
}

BOOL clsPrinter::mp_CheckColumns(void)
{
    int i = 0;
    int lActualPageWidth = 0;
    lActualPageWidth = CInt32(mp_dActualPageWidth * mp_fScaleInvMultp()) - CInt32(mp_lColumnsInEveryPageWidth() * mp_fScaleInvMultp());
    if (lActualPageWidth <= 0)
    {
        mp_oControl->mp_ErrorReport(PRINTER_INVALID_SPECS_COLUMNSINEVERYPAGE, "Invalid printing specifications, reduce the number of Columns in every page, set the ColumnsInEveryPage property to false, reduce the left or right margin or reduce the scale.", "ActiveGanttVCCtl.clsPrinter.Initialize");
        return false;
    }
    if (mp_bKeepColumnsTogether == TRUE)
    {
        for (i = 1; i <= mp_oControl->GetColumns()->GetCount(); i++)
        {
            if ((mp_oControl->GetColumns()->Item(CStr(i))->GetWidth() + 1) > lActualPageWidth)
            {
                mp_oControl->mp_ErrorReport(PRINTER_INVALID_SPECS_COLUMNS, "Invalid printing specifications, reduce the left or right margin, set the KeepColumnsTogether property to false or reduce the scale.", "ActiveGanttVCCtl.clsPrinter.Initialize");
                return FALSE;
            }
        }
    }
    return TRUE;
}

BOOL clsPrinter::mp_CheckTimeIntervals(void)
{
    int lActualPageWidth = 0;
    lActualPageWidth = CInt32(mp_dActualPageWidth * mp_fScaleInvMultp()) - CInt32(mp_lColumnsInEveryPageWidth() * mp_fScaleInvMultp());
    if (mp_bKeepTimeIntervalsTogether == TRUE && mp_bKeepColumnsTogether == TRUE)
    {
        DATE dtStartDate = mp_oView->GetTimeLine()->GetStartDate();
        DATE dtEndDate = mp_oControl->GetMathLib()->DateTimeAdd(mp_yInterval, mp_lFactor, dtStartDate);
        int lIntervalLength = CInt32((mp_oControl->GetMathLib()->GetXCoordinateFromDate(dtEndDate) - mp_oControl->GetMathLib()->GetXCoordinateFromDate(dtStartDate)) * mp_fScaleInvMultp());
        if (lIntervalLength > lActualPageWidth)
        {
            mp_oControl->mp_ErrorReport(PRINTER_INVALID_SPECS_INTERVAL, "Invalid printing specifications, reduce the left or right margin, set the KeepTimeIntervalsTogether property to false or reduce the scale.", "ActiveGanttVCCtl.clsPrinter.Initialize");
            return FALSE;
        }
    }
    return TRUE;
}

LONG clsPrinter::mp_lColumnsInEveryPageWidth(void)
{
    if (mp_bColumnsInEveryPage == TRUE && mp_lColumns > 0)
    {
        int i = 0;
        int lReturn = 0;
        for (i = 1; i <= mp_lColumns; i++)
        {
            if (mp_oControl->GetColumns()->Item(CStr(i))->GetVisible() == TRUE)
            {
                lReturn = lReturn + mp_oControl->GetColumns()->Item(CStr(i))->GetWidth();
            }
        }
        return lReturn;
    }
    else
    {
        return 0;
    }
}



//C++ Specific

CString clsPrinter::mp_sGetDefaultPrinterName(void)
{
	CString sReturn;
	TCHAR szPrinterName[MAX_PATH] = _T("");
	DWORD nLen = MAX_PATH ;

	GetDefaultPrinterW(szPrinterName, &nLen);
	sReturn = (TCHAR*)szPrinterName;
	return sReturn;
}

void clsPrinter::mp_StartDoc(void)
{
	mp_lhdcPrinter = CreateDC(_T("WINSPOOL"), mp_sPrinterName, NULL, NULL);
	DOCINFO docInfo;
	ZeroMemory(&docInfo, sizeof(docInfo));
	docInfo.cbSize = sizeof(docInfo);
	docInfo.lpszDocName = L"GdiplusPrint";


	DWORD dwNeeded, dwRet;
	LPWSTR sPrinterName = mp_sPrinterName.GetBuffer();

	OpenPrinter(sPrinterName, &mp_hPrinter, NULL);
	dwNeeded = DocumentProperties(NULL, mp_hPrinter, sPrinterName, NULL, NULL, 0);
	mp_oDEVMODE = (LPDEVMODE)malloc(dwNeeded);
	dwRet = DocumentProperties(NULL, mp_hPrinter, sPrinterName, mp_oDEVMODE, NULL, DM_OUT_BUFFER);

	mp_oDEVMODE->dmPaperSize = (short)mp_yPaperType;

	// Orientation
	if (mp_yOrientation == OT_PORTRAIT)
	{
		mp_oDEVMODE->dmOrientation = DMORIENT_PORTRAIT;
	}
	else if (mp_yOrientation == OT_LANDSCAPE)
	{
		mp_oDEVMODE->dmOrientation = DMORIENT_LANDSCAPE;
	}

	// PrinterResolution (HorizontalDPI/VerticalDPI)
	if (mp_yPrinterResolution == PR_CUSTOM)
	{
		mp_oDEVMODE->dmPrintQuality = (short)mp_lHorizontalDPI;
		mp_oDEVMODE->dmYResolution = (short)mp_lVerticalDPI;
	}
	else
	{
		switch (mp_yPrinterResolution)
		{
		case PR_HIGH:
			mp_oDEVMODE->dmPrintQuality = DMRES_HIGH;
			break;
		case PR_MEDIUM:
			mp_oDEVMODE->dmPrintQuality = DMRES_MEDIUM;
			break;
		case PR_LOW:
			mp_oDEVMODE->dmPrintQuality = DMRES_LOW;
			break;
		case PR_DRAFT:
			mp_oDEVMODE->dmPrintQuality = DMRES_DRAFT;
			break;
		}
	}

	mp_sPrinterName.ReleaseBuffer();


	mp_lhdcPrinter = ResetDC(mp_lhdcPrinter, mp_oDEVMODE);

	::StartDoc(mp_lhdcPrinter, &docInfo);
}

void clsPrinter::mp_EndDoc(void)
{
	::EndDoc(mp_lhdcPrinter);
	DeleteDC(mp_lhdcPrinter);
	mp_lhdcPrinter = NULL;
	ClosePrinter(mp_hPrinter);
	free(mp_oDEVMODE);
}

void clsPrinter::mp_StartPage(void)
{
	::StartPage(mp_lhdcPrinter);
}

void clsPrinter::mp_EndPage(void)
{
	::EndPage(mp_lhdcPrinter);
}