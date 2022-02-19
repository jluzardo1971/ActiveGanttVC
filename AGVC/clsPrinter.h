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

class clsView;
class clsGDIGraphics;

class clsPrinter : public CCmdTarget
{
	DECLARE_DYNCREATE(clsPrinter)
	clsPrinter();

	typedef enum E_PRINTINGOPERATIONTYPE
	{
        PT_SINGLEPAGE = 1,
        PT_RANGE = 2,
        PT_ALL = 3,
	}E_PRINTINGOPERATIONTYPE;

public:
	CActiveGanttVCCtl* mp_oControl;

	clsPrinter(CActiveGanttVCCtl* oControl);
	virtual ~clsPrinter();
	virtual void OnFinalRelease();
	void Finalize(void);

// Member Variables
public:
	CString mp_sPrinterName;
	LONG mp_yPaperType;
	LONG mp_yPrinterResolution;
	LONG mp_lHorizontalDPI;
	LONG mp_lVerticalDPI;
	LONG mp_lMarginLeft;
	LONG mp_lMarginTop;
	LONG mp_lMarginRight;
	LONG mp_lMarginBottom;

	FLOAT mp_lHardMarginX;
	FLOAT mp_lHardMarginY;


	//clsView* mp_oView;

	HDC mp_lhdcPrinter;
	
	DATE mp_dtStartDate;
	DATE mp_dtEndDate;
	DATE mp_dtPrintStartDateBuffer;
	LONG mp_lSplitterPositionBuffer;
	LONG mp_lFirstVisibleRowBuffer;
	clsView* mp_oView;
	LONG mp_yOrientation;
	DOUBLE mp_dPageWidth;
	DOUBLE mp_dPageHeight;
	DOUBLE mp_dActualPageWidth;
	DOUBLE mp_dActualPageHeight;
	LONG mp_lPages;
	LONG mp_lPageNumber;
	BOOL mp_bInitialized;
	LONG mp_lStartRow;
	LONG mp_lEndRow;
	LONG mp_lXAxisPages;
	LONG mp_lYAxisPages;
	FLOAT mp_fScale;
	BOOL mp_bKeepRowsTogether;
	BOOL mp_bKeepTimeIntervalsTogether;
	LONG mp_yInterval;
	LONG mp_lFactor;
	BOOL mp_bKeepColumnsTogether;
	BOOL mp_bShowAllColumns;
	BOOL mp_bHeadingsInEveryPage;
	BOOL mp_bColumnsInEveryPage;

	LONG mp_lColumns;
	CArray<int, int&> mp_aVertical;
	CArray<int, int&> mp_aHorizontal;
	LONG mp_yPrintingOperationType;
	LONG mp_lStartPage;
	LONG mp_lEndPage;



//Internal Properties (Extensions to ODL)
public:
	CString GetPrinterName(void);
	void SetPrinterName(CString value);

	LONG GetPrinterResolution(void);
	void SetPrinterResolution(LONG value);

	LONG GetHorizontalDPI(void);
	void SetHorizontalDPI(LONG value);

	LONG GetVerticalDPI(void);
	void SetVerticalDPI(LONG value);

	LONG GetPaperType(void);
	void SetPaperType(LONG value);

	LONG GetOrientation(void);
	void SetOrientation(LONG value);

	LONG GetMarginLeft(void);
	void SetMarginLeft(LONG value);

	LONG GetMarginTop(void);
	void SetMarginTop(LONG value);

	LONG GetMarginRight(void);
	void SetMarginRight(LONG value);

	LONG GetMarginBottom(void);
	void SetMarginBottom(LONG value);

	FLOAT GetScale(void);
	void SetScale(FLOAT value);

	BOOL GetHeadingsInEveryPage(void);
	void SetHeadingsInEveryPage(BOOL value);

	BOOL GetColumnsInEveryPage(void);
	void SetColumnsInEveryPage(BOOL value);

	LONG GetColumns(void);
	void SetColumns(LONG value);

	BOOL GetKeepRowsTogether(void);
	void SetKeepRowsTogether(BOOL value);

	BOOL GetKeepColumnsTogether(void);
	void SetKeepColumnsTogether(BOOL value);

	BOOL GetKeepTimeIntervalsTogether(void);
	void SetKeepTimeIntervalsTogether(BOOL value);

	BOOL GetShowAllColumns(void);
	void SetShowAllColumns(BOOL value);

	LONG GetInterval(void);
	void SetInterval(LONG value);

	LONG GetFactor(void);
	void SetFactor(LONG value);

	LONG GetPages(void);

	LONG GetXAxisPages(void);

	LONG GetYAxisPages(void);

	DATE GetPrintAreaEndDate(void);

	DATE GetPrintAreaStartDate(void);

	LONG GetStartRow(void);

	LONG GetEndRow(void);

	LONG GetPrintAreaWidth(void);

	LONG GetPrintAreaHeight(void);


//Internal Properties
public:


//Internal Methods
public:
	void Clear(void);

	void GetPagePosition(LONG PageNumber, LONG &Column, LONG &Row);

	LONG GetPageNumber(LONG Column, LONG Row);

	BOOL Initialize(DATE StartDate, DATE EndDate, LONG StartRow, LONG EndRow);

	void Terminate(void);

	void Calculate(void);

	void PrintAll(void);

	void PrintRange(LONG StartPage, LONG EndPage);

	void PrintPage(LONG PageNumber);

	void PreviewPage(OLE_HANDLE hdc, LONG PageNumber, FLOAT Scale, LONG X, LONG Y);



//Private Methods:
	void mp_PrintPage(Graphics* oGraphics);
	void mp_Draw(Graphics* oGraphics, LONG PageNumber, LONG lMarginLeft, LONG lMarginTop, LONG X1, LONG Y1, BOOL bToPrinter);
	LONG mp_lTimeLineHeight(LONG lRow);
	LONG mp_lRowsHeight(void);
	LONG mp_lTimeLineHeight(void);
	LONG mp_aX(LONG lIndex);
	LONG mp_aY(LONG lIndex);
	void mp_QueryPrinter(void);
	void mp_Calculate(void);
	LONG mp_lActualPageHeight(LONG lRow);
	void mp_CalculateRows(void);
	void mp_CalculateColumns(void);
	FLOAT mp_fScaleInvMultp(void);
	LONG mp_lActualPageWidth(void);
	LONG mp_lActualPageWidth(LONG lColumn);
	LONG mp_lColumnsWidth(LONG lColumn);
	void mp_GetPagePosition(LONG lPageNumber, LONG &lColumn, LONG &lRow);

	LONG mp_lColumnsInEveryPageWidth(void);
	BOOL mp_CheckTimeIntervals(void);
	BOOL mp_CheckColumns(void);
	BOOL mp_CheckRows(void);
	BOOL mp_CalculatePageDimensions(void);

	//C++ Specific
	CString mp_sGetDefaultPrinterName(void);

	void clsPrinter::mp_StartDoc(void);
	void clsPrinter::mp_EndDoc(void);
	void clsPrinter::mp_StartPage(void);
	void clsPrinter::mp_EndPage(void);


	LPDEVMODE mp_oDEVMODE;
	HANDLE mp_hPrinter;







protected:
	DECLARE_DISPATCH_MAP()
	DECLARE_MESSAGE_MAP()
	DECLARE_OLECREATE(clsPrinter)
	DECLARE_INTERFACE_MAP()

protected:


// Properties Exposed to ODL

BSTR odl_GetPrinterName(void)
{
	
	return GetPrinterName().AllocSysString();
}

void odl_SetPrinterName(LPCTSTR value)
{
	
	SetPrinterName(value);
}

LONG odl_GetPrinterResolution(void)
{
	
	return GetPrinterResolution();
}

void odl_SetPrinterResolution(LONG value)
{
	
	SetPrinterResolution(value);
}

LONG odl_GetHorizontalDPI(void)
{
	
	return GetHorizontalDPI();
}

void odl_SetHorizontalDPI(LONG value)
{
	
	SetHorizontalDPI(value);
}

LONG odl_GetVerticalDPI(void)
{
	
	return GetVerticalDPI();
}

void odl_SetVerticalDPI(LONG value)
{
	
	SetVerticalDPI(value);
}

LONG odl_GetPaperType(void)
{
	
	return GetPaperType();
}

void odl_SetPaperType(LONG value)
{
	
	SetPaperType(value);
}

LONG odl_GetOrientation(void)
{
	
	return GetOrientation();
}

void odl_SetOrientation(LONG value)
{
	
	SetOrientation(value);
}

LONG odl_GetMarginLeft(void)
{
	
	return GetMarginLeft();
}

void odl_SetMarginLeft(LONG value)
{
	
	SetMarginLeft(value);
}

LONG odl_GetMarginTop(void)
{
	
	return GetMarginTop();
}

void odl_SetMarginTop(LONG value)
{
	
	SetMarginTop(value);
}

LONG odl_GetMarginRight(void)
{
	
	return GetMarginRight();
}

void odl_SetMarginRight(LONG value)
{
	
	SetMarginRight(value);
}

LONG odl_GetMarginBottom(void)
{
	
	return GetMarginBottom();
}

void odl_SetMarginBottom(LONG value)
{
	
	SetMarginBottom(value);
}

FLOAT odl_GetScale(void)
{
	
	return GetScale();
}

void odl_SetScale(FLOAT value)
{
	
	SetScale(value);
}

BOOL odl_GetHeadingsInEveryPage(void)
{
	
	return GetHeadingsInEveryPage();
}

void odl_SetHeadingsInEveryPage(BOOL value)
{
	
	SetHeadingsInEveryPage(value);
}

BOOL odl_GetColumnsInEveryPage(void)
{
	
	return GetColumnsInEveryPage();
}

void odl_SetColumnsInEveryPage(BOOL value)
{
	
	SetColumnsInEveryPage(value);
}

LONG odl_GetColumns(void)
{
	
	return GetColumns();
}

void odl_SetColumns(LONG value)
{
	
	SetColumns(value);
}

BOOL odl_GetKeepRowsTogether(void)
{
	
	return GetKeepRowsTogether();
}

void odl_SetKeepRowsTogether(BOOL value)
{
	
	SetKeepRowsTogether(value);
}

BOOL odl_GetKeepColumnsTogether(void)
{
	
	return GetKeepColumnsTogether();
}

void odl_SetKeepColumnsTogether(BOOL value)
{
	
	SetKeepColumnsTogether(value);
}

BOOL odl_GetKeepTimeIntervalsTogether(void)
{
	
	return GetKeepTimeIntervalsTogether();
}

void odl_SetKeepTimeIntervalsTogether(BOOL value)
{
	
	SetKeepTimeIntervalsTogether(value);
}

BOOL odl_GetShowAllColumns(void)
{
	
	return GetShowAllColumns();
}

void odl_SetShowAllColumns(BOOL value)
{
	
	SetShowAllColumns(value);
}

LONG odl_GetInterval(void)
{
	
	return GetInterval();
}

void odl_SetInterval(LONG value)
{
	
	SetInterval(value);
}

LONG odl_GetFactor(void)
{
	
	return GetFactor();
}

void odl_SetFactor(LONG value)
{
	
	SetFactor(value);
}

LONG odl_GetPages(void)
{
	
	return GetPages();
}

LONG odl_GetXAxisPages(void)
{
	
	return GetXAxisPages();
}

LONG odl_GetYAxisPages(void)
{
	
	return GetYAxisPages();
}

DATE odl_GetPrintAreaEndDate(void)
{
	
	return mp_dtEndDate;
}

DATE odl_GetPrintAreaStartDate(void)
{
	
	return mp_dtStartDate;
}

LONG odl_GetStartRow(void)
{
	
	return GetStartRow();
}

LONG odl_GetEndRow(void)
{
	
	return GetEndRow();
}

LONG odl_GetPrintAreaWidth(void)
{
	
	return GetPrintAreaWidth();
}

LONG odl_GetPrintAreaHeight(void)
{
	
	return GetPrintAreaHeight();
}

void odl_Clear(void)
{
	
	Clear();
}

void odl_GetPagePosition(LONG PageNumber, LONG FAR* Column, LONG FAR* Row)
{
	
	GetPagePosition(PageNumber, *Column, *Row);
}

LONG odl_GetPageNumber(LONG Column, LONG Row)
{
	
	return GetPageNumber(Column, Row);
}

BOOL odl_Initialize(DATE StartDate, DATE EndDate, LONG StartRow, LONG EndRow)
{
	return Initialize(StartDate, StartDate, StartRow, EndRow);
}

void odl_Terminate(void)
{
	
	Terminate();
}

void odl_Calculate(void)
{
	
	Calculate();
}

void odl_PrintAll(void)
{
	
	PrintAll();
}

void odl_PrintRange(LONG StartPage, LONG EndPage)
{
	
	PrintRange(StartPage, EndPage);
}

void odl_PrintPage(LONG PageNumber)
{
	
	PrintPage(PageNumber);
}

void odl_PreviewPage(OLE_HANDLE hdc, LONG PageNumber, FLOAT Scale, LONG X, LONG Y)
{
	
	PreviewPage(hdc, PageNumber, Scale, X, Y);
}






};