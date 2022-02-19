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
#include "ActiveGanttVCWrappers.h"

namespace ActiveGanttVC
{
IMPLEMENT_DYNCREATE(CActiveGanttVCCtl, CWnd)

LONG CclsImage::GetWidth()
{
	LONG result;
	GetProperty(0x2, VT_I4, (void*)&result);
	return result;
}

LONG CclsImage::GetHeight()
{
	LONG result;
	GetProperty(0x3, VT_I4, (void*)&result);
	return result;
}

CString CclsImage::GetFilename()
{
	CString result;
	GetProperty(0x6, VT_BSTR, (void*)&result);
	return result;
}

BOOL CclsImage::GetEmbeddedColorManagement()
{
	BOOL result;
	GetProperty(0x7, VT_BOOL, (void*)&result);
	return result;
}

void CclsImage::FromFile(LPCTSTR Filename, BOOL UseEmbeddedColorManagement)
{
	static BYTE params[] = VTS_BSTR VTS_BOOL;
	InvokeHelper(0x1, DISPATCH_METHOD, VT_EMPTY, NULL, params, Filename, UseEmbeddedColorManagement);
}

void CclsImage::Clear(void)
{
	InvokeHelper(0x4, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
}

LONG CclsImage::hImage(void)
{
	LONG oReturn;
	InvokeHelper(0x5, DISPATCH_METHOD, VT_I4, (void*)&oReturn, NULL);
	return oReturn;
}

void CclsFont::SetFontName(LPCTSTR propval)
{
	SetProperty(0x1, VT_BSTR, propval);
}

CString CclsFont::GetFontName()
{
	CString result;
	GetProperty(0x1, VT_BSTR, (void*)&result);
	return result;
}

void CclsFont::SetSize(FLOAT propval)
{
	SetProperty(0x2, VT_R4, propval);
}

FLOAT CclsFont::GetSize()
{
	FLOAT result;
	GetProperty(0x2, VT_R4, (void*)&result);
	return result;
}

void CclsFont::SetStyle(GRE_FONTSTYLE propval)
{
	LONG lpropval = (LONG)propval;
	SetProperty(0x3, VT_I4, lpropval);
}

GRE_FONTSTYLE CclsFont::GetStyle()
{
	LONG result;
	GetProperty(0x3, VT_I4, (void*)&result);
	return (GRE_FONTSTYLE)result;
}

void CclsFont::SetUnit(GRE_UNIT propval)
{
	LONG lpropval = (LONG)propval;
	SetProperty(0x4, VT_I4, lpropval);
}

GRE_UNIT CclsFont::GetUnit()
{
	LONG result;
	GetProperty(0x4, VT_I4, (void*)&result);
	return (GRE_UNIT)result;
}

LONG CclsGDIGraphics::GetHDC(void)
{
	LONG oReturn;
	InvokeHelper(0x1, DISPATCH_METHOD, VT_HANDLE, (void*)&oReturn, NULL);
	return oReturn;
}

void CclsGDIGraphics::ReleaseHDC(LONG hdc)
{
	static BYTE params[] = VTS_HANDLE;
	InvokeHelper(0x2, DISPATCH_METHOD, VT_EMPTY, NULL, params, hdc);
}

LONG CclsGDIGraphics::StringWidth(LPCTSTR Text, LPDISPATCH Font, LONG hWnd)
{
	LONG oReturn;
	static BYTE params[] = VTS_BSTR VTS_DISPATCH VTS_I4;
	InvokeHelper(0x3, DISPATCH_METHOD, VT_I4, (void*)&oReturn, params, Text, Font, hWnd);
	return oReturn;
}

LONG CclsGDIGraphics::StringHeight(LPCTSTR Text, LPDISPATCH Font, LONG hWnd)
{
	LONG oReturn;
	static BYTE params[] = VTS_BSTR VTS_DISPATCH VTS_I4;
	InvokeHelper(0x4, DISPATCH_METHOD, VT_I4, (void*)&oReturn, params, Text, Font, hWnd);
	return oReturn;
}

void CclsGDIGraphics::SetLayoutRect(LONG X1, LONG Y1, LONG X2, LONG Y2)
{
	static BYTE params[] = VTS_I4 VTS_I4 VTS_I4 VTS_I4;
	InvokeHelper(0x5, DISPATCH_METHOD, VT_EMPTY, NULL, params, X1, Y1, X2, Y2);
}

void CclsPrinter::SetPrinterName(LPCTSTR propval)
{
	SetProperty(0x1, VT_BSTR, propval);
}

CString CclsPrinter::GetPrinterName()
{
	CString result;
	GetProperty(0x1, VT_BSTR, (void*)&result);
	return result;
}

void CclsPrinter::SetPrinterResolution(GRE_PRINTERRESOLUTION propval)
{
	LONG lpropval = (LONG)propval;
	SetProperty(0x2, VT_I4, lpropval);
}

GRE_PRINTERRESOLUTION CclsPrinter::GetPrinterResolution()
{
	LONG result;
	GetProperty(0x2, VT_I4, (void*)&result);
	return (GRE_PRINTERRESOLUTION)result;
}

void CclsPrinter::SetHorizontalDPI(LONG propval)
{
	SetProperty(0x3, VT_I4, propval);
}

LONG CclsPrinter::GetHorizontalDPI()
{
	LONG result;
	GetProperty(0x3, VT_I4, (void*)&result);
	return result;
}

void CclsPrinter::SetVerticalDPI(LONG propval)
{
	SetProperty(0x4, VT_I4, propval);
}

LONG CclsPrinter::GetVerticalDPI()
{
	LONG result;
	GetProperty(0x4, VT_I4, (void*)&result);
	return result;
}

void CclsPrinter::SetPaperType(GRE_PAPERTYPE propval)
{
	LONG lpropval = (LONG)propval;
	SetProperty(0x5, VT_I4, lpropval);
}

GRE_PAPERTYPE CclsPrinter::GetPaperType()
{
	LONG result;
	GetProperty(0x5, VT_I4, (void*)&result);
	return (GRE_PAPERTYPE)result;
}

void CclsPrinter::SetOrientation(GRE_ORIENTATION propval)
{
	LONG lpropval = (LONG)propval;
	SetProperty(0x6, VT_I4, lpropval);
}

GRE_ORIENTATION CclsPrinter::GetOrientation()
{
	LONG result;
	GetProperty(0x6, VT_I4, (void*)&result);
	return (GRE_ORIENTATION)result;
}

void CclsPrinter::SetMarginLeft(LONG propval)
{
	SetProperty(0x7, VT_I4, propval);
}

LONG CclsPrinter::GetMarginLeft()
{
	LONG result;
	GetProperty(0x7, VT_I4, (void*)&result);
	return result;
}

void CclsPrinter::SetMarginTop(LONG propval)
{
	SetProperty(0x8, VT_I4, propval);
}

LONG CclsPrinter::GetMarginTop()
{
	LONG result;
	GetProperty(0x8, VT_I4, (void*)&result);
	return result;
}

void CclsPrinter::SetMarginRight(LONG propval)
{
	SetProperty(0x9, VT_I4, propval);
}

LONG CclsPrinter::GetMarginRight()
{
	LONG result;
	GetProperty(0x9, VT_I4, (void*)&result);
	return result;
}

void CclsPrinter::SetMarginBottom(LONG propval)
{
	SetProperty(0xa, VT_I4, propval);
}

LONG CclsPrinter::GetMarginBottom()
{
	LONG result;
	GetProperty(0xa, VT_I4, (void*)&result);
	return result;
}

void CclsPrinter::SetScale(FLOAT propval)
{
	SetProperty(0xb, VT_R4, propval);
}

FLOAT CclsPrinter::GetScale()
{
	FLOAT result;
	GetProperty(0xb, VT_R4, (void*)&result);
	return result;
}

void CclsPrinter::SetHeadingsInEveryPage(BOOL propval)
{
	SetProperty(0xc, VT_BOOL, propval);
}

BOOL CclsPrinter::GetHeadingsInEveryPage()
{
	BOOL result;
	GetProperty(0xc, VT_BOOL, (void*)&result);
	return result;
}

void CclsPrinter::SetColumnsInEveryPage(BOOL propval)
{
	SetProperty(0xd, VT_BOOL, propval);
}

BOOL CclsPrinter::GetColumnsInEveryPage()
{
	BOOL result;
	GetProperty(0xd, VT_BOOL, (void*)&result);
	return result;
}

void CclsPrinter::SetColumns(LONG propval)
{
	SetProperty(0xe, VT_I4, propval);
}

LONG CclsPrinter::GetColumns()
{
	LONG result;
	GetProperty(0xe, VT_I4, (void*)&result);
	return result;
}

void CclsPrinter::SetKeepRowsTogether(BOOL propval)
{
	SetProperty(0xf, VT_BOOL, propval);
}

BOOL CclsPrinter::GetKeepRowsTogether()
{
	BOOL result;
	GetProperty(0xf, VT_BOOL, (void*)&result);
	return result;
}

void CclsPrinter::SetKeepColumnsTogether(BOOL propval)
{
	SetProperty(0x10, VT_BOOL, propval);
}

BOOL CclsPrinter::GetKeepColumnsTogether()
{
	BOOL result;
	GetProperty(0x10, VT_BOOL, (void*)&result);
	return result;
}

void CclsPrinter::SetKeepTimeIntervalsTogether(BOOL propval)
{
	SetProperty(0x11, VT_BOOL, propval);
}

BOOL CclsPrinter::GetKeepTimeIntervalsTogether()
{
	BOOL result;
	GetProperty(0x11, VT_BOOL, (void*)&result);
	return result;
}

void CclsPrinter::SetShowAllColumns(BOOL propval)
{
	SetProperty(0x12, VT_BOOL, propval);
}

BOOL CclsPrinter::GetShowAllColumns()
{
	BOOL result;
	GetProperty(0x12, VT_BOOL, (void*)&result);
	return result;
}

void CclsPrinter::SetInterval(E_INTERVAL propval)
{
	LONG lpropval = (LONG)propval;
	SetProperty(0x13, VT_I4, lpropval);
}

E_INTERVAL CclsPrinter::GetInterval()
{
	LONG result;
	GetProperty(0x13, VT_I4, (void*)&result);
	return (E_INTERVAL)result;
}

void CclsPrinter::SetFactor(LONG propval)
{
	SetProperty(0x14, VT_I4, propval);
}

LONG CclsPrinter::GetFactor()
{
	LONG result;
	GetProperty(0x14, VT_I4, (void*)&result);
	return result;
}

LONG CclsPrinter::GetPages()
{
	LONG result;
	GetProperty(0x15, VT_I4, (void*)&result);
	return result;
}

LONG CclsPrinter::GetXAxisPages()
{
	LONG result;
	GetProperty(0x16, VT_I4, (void*)&result);
	return result;
}

LONG CclsPrinter::GetYAxisPages()
{
	LONG result;
	GetProperty(0x17, VT_I4, (void*)&result);
	return result;
}

DATE CclsPrinter::GetPrintAreaEndDate()
{
	DATE result;
	GetProperty(0x18, VT_DATE, (void*)&result);
	return result;
}

DATE CclsPrinter::GetPrintAreaStartDate()
{
	DATE result;
	GetProperty(0x19, VT_DATE, (void*)&result);
	return result;
}

LONG CclsPrinter::GetStartRow()
{
	LONG result;
	GetProperty(0x1a, VT_I4, (void*)&result);
	return result;
}

LONG CclsPrinter::GetEndRow()
{
	LONG result;
	GetProperty(0x1b, VT_I4, (void*)&result);
	return result;
}

LONG CclsPrinter::GetPrintAreaWidth()
{
	LONG result;
	GetProperty(0x1c, VT_I4, (void*)&result);
	return result;
}

LONG CclsPrinter::GetPrintAreaHeight()
{
	LONG result;
	GetProperty(0x1d, VT_I4, (void*)&result);
	return result;
}

void CclsPrinter::SetPScale(FLOAT propval)
{
	SetProperty(0x28, VT_R4, propval);
}

FLOAT CclsPrinter::GetPScale()
{
	FLOAT result;
	GetProperty(0x28, VT_R4, (void*)&result);
	return result;
}

void CclsPrinter::Clear(void)
{
	InvokeHelper(0x1e, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
}

void CclsPrinter::GetPagePosition(LONG PageNumber, LONG* Column, LONG* Row)
{
	static BYTE params[] = VTS_I4 VTS_PI4 VTS_PI4;
	InvokeHelper(0x1f, DISPATCH_METHOD, VT_EMPTY, NULL, params, PageNumber, Column, Row);
}

LONG CclsPrinter::GetPageNumber(LONG Column, LONG Row)
{
	LONG oReturn;
	static BYTE params[] = VTS_I4 VTS_I4;
	InvokeHelper(0x20, DISPATCH_METHOD, VT_I4, (void*)&oReturn, params, Column, Row);
	return oReturn;
}

BOOL CclsPrinter::Initialize(DATE StartDate, DATE EndDate, LONG StartRow, LONG EndRow)
{
	BOOL oReturn;
	static BYTE params[] = VTS_DISPATCH VTS_DISPATCH VTS_I4 VTS_I4;
	InvokeHelper(0x21, DISPATCH_METHOD, VT_BOOL, (void*)&oReturn, params, StartDate, EndDate, StartRow, EndRow);
	return oReturn;
}

void CclsPrinter::Terminate(void)
{
	InvokeHelper(0x22, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
}

void CclsPrinter::Calculate(void)
{
	InvokeHelper(0x23, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
}

void CclsPrinter::PrintAll(void)
{
	InvokeHelper(0x24, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
}

void CclsPrinter::PrintRange(LONG StartPage, LONG EndPage)
{
	static BYTE params[] = VTS_I4 VTS_I4;
	InvokeHelper(0x25, DISPATCH_METHOD, VT_EMPTY, NULL, params, StartPage, EndPage);
}

void CclsPrinter::PrintPage(LONG PageNumber)
{
	static BYTE params[] = VTS_I4;
	InvokeHelper(0x26, DISPATCH_METHOD, VT_EMPTY, NULL, params, PageNumber);
}

void CclsPrinter::PreviewPage(LONG hdc, LONG PageNumber, FLOAT Scale, LONG X, LONG Y)
{
	static BYTE params[] = VTS_HANDLE VTS_I4 VTS_R4 VTS_I4 VTS_I4;
	InvokeHelper(0x27, DISPATCH_METHOD, VT_EMPTY, NULL, params, hdc, PageNumber, Scale, X, Y);
}

DATE CclsMath::DateTimeAdd(E_INTERVAL Interval, LONG Number, DATE dtDate)
{
	DATE oReturn;
	static BYTE params[] = VTS_I4 VTS_I4 VTS_DATE;
	InvokeHelper(0x1, DISPATCH_METHOD, VT_DATE, (void*)&oReturn, params, Interval, Number, dtDate);
	return oReturn;
}

LONG CclsMath::DateTimeDiff(E_INTERVAL Interval, DATE dtDate1, DATE dtDate2)
{
	LONG oReturn;
	static BYTE params[] = VTS_I4 VTS_DATE VTS_DATE;
	InvokeHelper(0x2, DISPATCH_METHOD, VT_I4, (void*)&oReturn, params, Interval, dtDate1, dtDate2);
	return oReturn;
}

LONG CclsMath::GetXCoordinateFromDate(DATE dtCoordinate)
{
	LONG oReturn;
	static BYTE params[] = VTS_DATE;
	InvokeHelper(0x3, DISPATCH_METHOD, VT_I4, (void*)&oReturn, params, dtCoordinate);
	return oReturn;
}

DATE CclsMath::GetDateFromXCoordinate(LONG v_lXCoordinate)
{
	DATE oReturn;
	static BYTE params[] = VTS_I4;
	InvokeHelper(0x4, DISPATCH_METHOD, VT_DATE, (void*)&oReturn, params, v_lXCoordinate);
	return oReturn;
}

LONG CclsMath::GetRowIndexByPosition(LONG Y)
{
	LONG oReturn;
	static BYTE params[] = VTS_I4;
	InvokeHelper(0x5, DISPATCH_METHOD, VT_I4, (void*)&oReturn, params, Y);
	return oReturn;
}

LONG CclsMath::GetCellIndexByPosition(LONG X)
{
	LONG oReturn;
	static BYTE params[] = VTS_I4;
	InvokeHelper(0x6, DISPATCH_METHOD, VT_I4, (void*)&oReturn, params, X);
	return oReturn;
}

LONG CclsMath::GetColumnIndexByPosition(LONG X, LONG Y)
{
	LONG oReturn;
	static BYTE params[] = VTS_I4 VTS_I4;
	InvokeHelper(0x7, DISPATCH_METHOD, VT_I4, (void*)&oReturn, params, X, Y);
	return oReturn;
}

LONG CclsMath::GetTaskIndexByPosition(LONG X, LONG Y)
{
	LONG oReturn;
	static BYTE params[] = VTS_I4 VTS_I4;
	InvokeHelper(0x8, DISPATCH_METHOD, VT_I4, (void*)&oReturn, params, X, Y);
	return oReturn;
}

LONG CclsMath::GetPercentageIndexByPosition(LONG X, LONG Y)
{
	LONG oReturn;
	static BYTE params[] = VTS_I4 VTS_I4;
	InvokeHelper(0x9, DISPATCH_METHOD, VT_I4, (void*)&oReturn, params, X, Y);
	return oReturn;
}

LONG CclsMath::GetNodeIndexByCheckBoxPosition(LONG X, LONG Y)
{
	LONG oReturn;
	static BYTE params[] = VTS_I4 VTS_I4;
	InvokeHelper(0xa, DISPATCH_METHOD, VT_I4, (void*)&oReturn, params, X, Y);
	return oReturn;
}

LONG CclsMath::GetNodeIndexBySignPosition(LONG X, LONG Y)
{
	LONG oReturn;
	static BYTE params[] = VTS_I4 VTS_I4;
	InvokeHelper(0xb, DISPATCH_METHOD, VT_I4, (void*)&oReturn, params, X, Y);
	return oReturn;
}

BOOL CclsMath::DetectConflict(DATE StartDate, DATE EndDate, LPCTSTR RowKey, LONG ExcludeIndex, LPCTSTR LayerIndex)
{
	BOOL oReturn;
	static BYTE params[] = VTS_DATE VTS_DATE VTS_BSTR VTS_I4 VTS_BSTR;
	InvokeHelper(0xc, DISPATCH_METHOD, VT_BOOL, (void*)&oReturn, params, StartDate, EndDate, RowKey, ExcludeIndex, LayerIndex);
	return oReturn;
}

DATE CclsMath::RoundDate(E_INTERVAL Interval, LONG Number, DATE dtDate)
{
	DATE oReturn;
	static BYTE params[] = VTS_I4 VTS_I4 VTS_DATE;
	InvokeHelper(0xd, DISPATCH_METHOD, VT_DATE, (void*)&oReturn, params, Interval, Number, dtDate);
	return oReturn;
}

LONG CclsMath::RoundDouble(DOUBLE dParam)
{
	LONG oReturn;
	static BYTE params[] = VTS_R8;
	InvokeHelper(0xe, DISPATCH_METHOD, VT_I4, (void*)&oReturn, params, dParam);
	return oReturn;
}

LONG CclsMath::GetPredecessorIndexByPosition(LONG X, LONG Y)
{
	LONG oReturn;
	static BYTE params[] = VTS_I4 VTS_I4;
	InvokeHelper(0xf, DISPATCH_METHOD, VT_I4, (void*)&oReturn, params, X, Y);
	return oReturn;
}

LONG CclsMath::CalculateDuration(DATE dtStartDate, DATE dtEndDate, E_INTERVAL DurationInterval)
{
	LONG oReturn;
	static BYTE params[] = VTS_DATE VTS_DATE VTS_I4;
	InvokeHelper(0x10, DISPATCH_METHOD, VT_I4, (void*)&oReturn, params, dtStartDate, dtEndDate, DurationInterval);
	return oReturn;
}

DATE CclsMath::GetEndDate(DATE StartDate, E_INTERVAL DurationInterval, LONG DurationFactor)
{
	DATE oReturn;
	static BYTE params[] = VTS_DATE VTS_I4 VTS_I4;
	InvokeHelper(0x11, DISPATCH_METHOD, VT_DATE, (void*)&oReturn, params, StartDate, DurationInterval, DurationFactor);
	return oReturn;
}

void CclsCustomBorderStyle::SetLeft(BOOL propval)
{
	SetProperty(0x1, VT_BOOL, propval);
}

BOOL CclsCustomBorderStyle::GetLeft()
{
	BOOL result;
	GetProperty(0x1, VT_BOOL, (void*)&result);
	return result;
}

void CclsCustomBorderStyle::SetTop(BOOL propval)
{
	SetProperty(0x2, VT_BOOL, propval);
}

BOOL CclsCustomBorderStyle::GetTop()
{
	BOOL result;
	GetProperty(0x2, VT_BOOL, (void*)&result);
	return result;
}

void CclsCustomBorderStyle::SetRight(BOOL propval)
{
	SetProperty(0x3, VT_BOOL, propval);
}

BOOL CclsCustomBorderStyle::GetRight()
{
	BOOL result;
	GetProperty(0x3, VT_BOOL, (void*)&result);
	return result;
}

void CclsCustomBorderStyle::SetBottom(BOOL propval)
{
	SetProperty(0x4, VT_BOOL, propval);
}

BOOL CclsCustomBorderStyle::GetBottom()
{
	BOOL result;
	GetProperty(0x4, VT_BOOL, (void*)&result);
	return result;
}

CString CclsCustomBorderStyle::GetXML(void)
{
	CString oReturn;
	InvokeHelper(0x5, DISPATCH_METHOD, VT_BSTR, (void*)&oReturn, NULL);
	return oReturn;
}

void CclsCustomBorderStyle::SetXML(LPCTSTR sXML)
{
	static BYTE params[] = VTS_BSTR;
	InvokeHelper(0x6, DISPATCH_METHOD, VT_EMPTY, NULL, params, sXML);
}

void CclsTaskStyle::SetEndBorderColor(unsigned long propval)
{
	SetProperty(0x1, VT_COLOR, propval);
}

unsigned long CclsTaskStyle::GetEndBorderColor()
{
	unsigned long result;
	GetProperty(0x1, VT_COLOR, (void*)&result);
	return result;
}

void CclsTaskStyle::SetEndFillColor(unsigned long propval)
{
	SetProperty(0x2, VT_COLOR, propval);
}

unsigned long CclsTaskStyle::GetEndFillColor()
{
	unsigned long result;
	GetProperty(0x2, VT_COLOR, (void*)&result);
	return result;
}

void CclsTaskStyle::SetStartBorderColor(unsigned long propval)
{
	SetProperty(0x3, VT_COLOR, propval);
}

unsigned long CclsTaskStyle::GetStartBorderColor()
{
	unsigned long result;
	GetProperty(0x3, VT_COLOR, (void*)&result);
	return result;
}

void CclsTaskStyle::SetStartFillColor(unsigned long propval)
{
	SetProperty(0x4, VT_COLOR, propval);
}

unsigned long CclsTaskStyle::GetStartFillColor()
{
	unsigned long result;
	GetProperty(0x4, VT_COLOR, (void*)&result);
	return result;
}

CclsImage CclsTaskStyle::GetStartImage()
{
	LPDISPATCH pDispatch;
	GetProperty(0x5, VT_DISPATCH, (void*)&pDispatch);
	return CclsImage(pDispatch);
}

CclsImage CclsTaskStyle::GetMiddleImage()
{
	LPDISPATCH pDispatch;
	GetProperty(0x6, VT_DISPATCH, (void*)&pDispatch);
	return CclsImage(pDispatch);
}

CclsImage CclsTaskStyle::GetEndImage()
{
	LPDISPATCH pDispatch;
	GetProperty(0x7, VT_DISPATCH, (void*)&pDispatch);
	return CclsImage(pDispatch);
}

void CclsTaskStyle::SetStartShapeIndex(GRE_FIGURETYPE propval)
{
	LONG lpropval = (LONG)propval;
	SetProperty(0x8, VT_I4, lpropval);
}

GRE_FIGURETYPE CclsTaskStyle::GetStartShapeIndex()
{
	LONG result;
	GetProperty(0x8, VT_I4, (void*)&result);
	return (GRE_FIGURETYPE)result;
}

void CclsTaskStyle::SetEndShapeIndex(GRE_FIGURETYPE propval)
{
	LONG lpropval = (LONG)propval;
	SetProperty(0x9, VT_I4, lpropval);
}

GRE_FIGURETYPE CclsTaskStyle::GetEndShapeIndex()
{
	LONG result;
	GetProperty(0x9, VT_I4, (void*)&result);
	return (GRE_FIGURETYPE)result;
}

void CclsTaskStyle::SetStartImageTag(LPCTSTR propval)
{
	SetProperty(0xc, VT_BSTR, propval);
}

CString CclsTaskStyle::GetStartImageTag()
{
	CString result;
	GetProperty(0xc, VT_BSTR, (void*)&result);
	return result;
}

void CclsTaskStyle::SetMiddleImageTag(LPCTSTR propval)
{
	SetProperty(0xd, VT_BSTR, propval);
}

CString CclsTaskStyle::GetMiddleImageTag()
{
	CString result;
	GetProperty(0xd, VT_BSTR, (void*)&result);
	return result;
}

void CclsTaskStyle::SetEndImageTag(LPCTSTR propval)
{
	SetProperty(0xe, VT_BSTR, propval);
}

CString CclsTaskStyle::GetEndImageTag()
{
	CString result;
	GetProperty(0xe, VT_BSTR, (void*)&result);
	return result;
}

CString CclsTaskStyle::GetXML(void)
{
	CString oReturn;
	InvokeHelper(0xa, DISPATCH_METHOD, VT_BSTR, (void*)&oReturn, NULL);
	return oReturn;
}

void CclsTaskStyle::SetXML(LPCTSTR sXML)
{
	static BYTE params[] = VTS_BSTR;
	InvokeHelper(0xb, DISPATCH_METHOD, VT_EMPTY, NULL, params, sXML);
}

void CclsMilestoneStyle::SetBorderColor(unsigned long propval)
{
	SetProperty(0x1, VT_COLOR, propval);
}

unsigned long CclsMilestoneStyle::GetBorderColor()
{
	unsigned long result;
	GetProperty(0x1, VT_COLOR, (void*)&result);
	return result;
}

void CclsMilestoneStyle::SetFillColor(unsigned long propval)
{
	SetProperty(0x2, VT_COLOR, propval);
}

unsigned long CclsMilestoneStyle::GetFillColor()
{
	unsigned long result;
	GetProperty(0x2, VT_COLOR, (void*)&result);
	return result;
}

void CclsMilestoneStyle::SetShapeIndex(GRE_FIGURETYPE propval)
{
	LONG lpropval = (LONG)propval;
	SetProperty(0x3, VT_I4, lpropval);
}

GRE_FIGURETYPE CclsMilestoneStyle::GetShapeIndex()
{
	LONG result;
	GetProperty(0x3, VT_I4, (void*)&result);
	return (GRE_FIGURETYPE)result;
}

CclsImage CclsMilestoneStyle::GetImage()
{
	LPDISPATCH pDispatch;
	GetProperty(0x6, VT_DISPATCH, (void*)&pDispatch);
	return CclsImage(pDispatch);
}

void CclsMilestoneStyle::SetImageTag(LPCTSTR propval)
{
	SetProperty(0x7, VT_BSTR, propval);
}

CString CclsMilestoneStyle::GetImageTag()
{
	CString result;
	GetProperty(0x7, VT_BSTR, (void*)&result);
	return result;
}

CString CclsMilestoneStyle::GetXML(void)
{
	CString oReturn;
	InvokeHelper(0x4, DISPATCH_METHOD, VT_BSTR, (void*)&oReturn, NULL);
	return oReturn;
}

void CclsMilestoneStyle::SetXML(LPCTSTR sXML)
{
	static BYTE params[] = VTS_BSTR;
	InvokeHelper(0x5, DISPATCH_METHOD, VT_EMPTY, NULL, params, sXML);
}

void CclsPredecessorStyle::SetLineColor(unsigned long propval)
{
	SetProperty(0x1, VT_COLOR, propval);
}

unsigned long CclsPredecessorStyle::GetLineColor()
{
	unsigned long result;
	GetProperty(0x1, VT_COLOR, (void*)&result);
	return result;
}

void CclsPredecessorStyle::SetXOffset(LONG propval)
{
	SetProperty(0x2, VT_I4, propval);
}

LONG CclsPredecessorStyle::GetXOffset()
{
	LONG result;
	GetProperty(0x2, VT_I4, (void*)&result);
	return result;
}

void CclsPredecessorStyle::SetYOffset(LONG propval)
{
	SetProperty(0x3, VT_I4, propval);
}

LONG CclsPredecessorStyle::GetYOffset()
{
	LONG result;
	GetProperty(0x3, VT_I4, (void*)&result);
	return result;
}

void CclsPredecessorStyle::SetLineWidth(LONG propval)
{
	SetProperty(0x4, VT_I4, propval);
}

LONG CclsPredecessorStyle::GetLineWidth()
{
	LONG result;
	GetProperty(0x4, VT_I4, (void*)&result);
	return result;
}

void CclsPredecessorStyle::SetLineStyle(GRE_LINEDRAWSTYLE propval)
{
	LONG lpropval = (LONG)propval;
	SetProperty(0x5, VT_I4, lpropval);
}

GRE_LINEDRAWSTYLE CclsPredecessorStyle::GetLineStyle()
{
	LONG result;
	GetProperty(0x5, VT_I4, (void*)&result);
	return (GRE_LINEDRAWSTYLE)result;
}

void CclsPredecessorStyle::SetArrowSize(LONG propval)
{
	SetProperty(0x6, VT_I4, propval);
}

LONG CclsPredecessorStyle::GetArrowSize()
{
	LONG result;
	GetProperty(0x6, VT_I4, (void*)&result);
	return result;
}

CString CclsPredecessorStyle::GetXML(void)
{
	CString oReturn;
	InvokeHelper(0x7, DISPATCH_METHOD, VT_BSTR, (void*)&oReturn, NULL);
	return oReturn;
}

void CclsPredecessorStyle::SetXML(LPCTSTR sXML)
{
	static BYTE params[] = VTS_BSTR;
	InvokeHelper(0x8, DISPATCH_METHOD, VT_EMPTY, NULL, params, sXML);
}

void CclsTextFlags::SetVerticalAlignment(GRE_VERTICALALIGNMENT propval)
{
	LONG lpropval = (LONG)propval;
	SetProperty(0x1, VT_I4, lpropval);
}

GRE_VERTICALALIGNMENT CclsTextFlags::GetVerticalAlignment()
{
	LONG result;
	GetProperty(0x1, VT_I4, (void*)&result);
	return (GRE_VERTICALALIGNMENT)result;
}

void CclsTextFlags::SetHorizontalAlignment(GRE_HORIZONTALALIGNMENT propval)
{
	LONG lpropval = (LONG)propval;
	SetProperty(0x2, VT_I4, lpropval);
}

GRE_HORIZONTALALIGNMENT CclsTextFlags::GetHorizontalAlignment()
{
	LONG result;
	GetProperty(0x2, VT_I4, (void*)&result);
	return (GRE_HORIZONTALALIGNMENT)result;
}

void CclsTextFlags::SetWordWrap(BOOL propval)
{
	SetProperty(0x3, VT_BOOL, propval);
}

BOOL CclsTextFlags::GetWordWrap()
{
	BOOL result;
	GetProperty(0x3, VT_BOOL, (void*)&result);
	return result;
}

void CclsTextFlags::SetRightToLeft(BOOL propval)
{
	SetProperty(0x4, VT_BOOL, propval);
}

BOOL CclsTextFlags::GetRightToLeft()
{
	BOOL result;
	GetProperty(0x4, VT_BOOL, (void*)&result);
	return result;
}

void CclsTextFlags::SetOffsetBottom(LONG propval)
{
	SetProperty(0x5, VT_I4, propval);
}

LONG CclsTextFlags::GetOffsetBottom()
{
	LONG result;
	GetProperty(0x5, VT_I4, (void*)&result);
	return result;
}

void CclsTextFlags::SetOffsetLeft(LONG propval)
{
	SetProperty(0x6, VT_I4, propval);
}

LONG CclsTextFlags::GetOffsetLeft()
{
	LONG result;
	GetProperty(0x6, VT_I4, (void*)&result);
	return result;
}

void CclsTextFlags::SetOffsetRight(LONG propval)
{
	SetProperty(0x7, VT_I4, propval);
}

LONG CclsTextFlags::GetOffsetRight()
{
	LONG result;
	GetProperty(0x7, VT_I4, (void*)&result);
	return result;
}

void CclsTextFlags::SetOffsetTop(LONG propval)
{
	SetProperty(0x8, VT_I4, propval);
}

LONG CclsTextFlags::GetOffsetTop()
{
	LONG result;
	GetProperty(0x8, VT_I4, (void*)&result);
	return result;
}

CString CclsTextFlags::GetXML(void)
{
	CString oReturn;
	InvokeHelper(0x9, DISPATCH_METHOD, VT_BSTR, (void*)&oReturn, NULL);
	return oReturn;
}

void CclsTextFlags::SetXML(LPCTSTR sXML)
{
	static BYTE params[] = VTS_BSTR;
	InvokeHelper(0xa, DISPATCH_METHOD, VT_EMPTY, NULL, params, sXML);
}

void CclsButtonBorderStyle::SetRaisedExteriorLeftTopColor(unsigned long propval)
{
	SetProperty(0x1, VT_COLOR, propval);
}

unsigned long CclsButtonBorderStyle::GetRaisedExteriorLeftTopColor()
{
	unsigned long result;
	GetProperty(0x1, VT_COLOR, (void*)&result);
	return result;
}

void CclsButtonBorderStyle::SetRaisedInteriorLeftTopColor(unsigned long propval)
{
	SetProperty(0x2, VT_COLOR, propval);
}

unsigned long CclsButtonBorderStyle::GetRaisedInteriorLeftTopColor()
{
	unsigned long result;
	GetProperty(0x2, VT_COLOR, (void*)&result);
	return result;
}

void CclsButtonBorderStyle::SetRaisedExteriorRightBottomColor(unsigned long propval)
{
	SetProperty(0x3, VT_COLOR, propval);
}

unsigned long CclsButtonBorderStyle::GetRaisedExteriorRightBottomColor()
{
	unsigned long result;
	GetProperty(0x3, VT_COLOR, (void*)&result);
	return result;
}

void CclsButtonBorderStyle::SetRaisedInteriorRightBottomColor(unsigned long propval)
{
	SetProperty(0x4, VT_COLOR, propval);
}

unsigned long CclsButtonBorderStyle::GetRaisedInteriorRightBottomColor()
{
	unsigned long result;
	GetProperty(0x4, VT_COLOR, (void*)&result);
	return result;
}

void CclsButtonBorderStyle::SetSunkenExteriorLeftTopColor(unsigned long propval)
{
	SetProperty(0x5, VT_COLOR, propval);
}

unsigned long CclsButtonBorderStyle::GetSunkenExteriorLeftTopColor()
{
	unsigned long result;
	GetProperty(0x5, VT_COLOR, (void*)&result);
	return result;
}

void CclsButtonBorderStyle::SetSunkenInteriorLeftTopColor(unsigned long propval)
{
	SetProperty(0x6, VT_COLOR, propval);
}

unsigned long CclsButtonBorderStyle::GetSunkenInteriorLeftTopColor()
{
	unsigned long result;
	GetProperty(0x6, VT_COLOR, (void*)&result);
	return result;
}

void CclsButtonBorderStyle::SetSunkenExteriorRightBottomColor(unsigned long propval)
{
	SetProperty(0x7, VT_COLOR, propval);
}

unsigned long CclsButtonBorderStyle::GetSunkenExteriorRightBottomColor()
{
	unsigned long result;
	GetProperty(0x7, VT_COLOR, (void*)&result);
	return result;
}

void CclsButtonBorderStyle::SetSunkenInteriorRightBottomColor(unsigned long propval)
{
	SetProperty(0x8, VT_COLOR, propval);
}

unsigned long CclsButtonBorderStyle::GetSunkenInteriorRightBottomColor()
{
	unsigned long result;
	GetProperty(0x8, VT_COLOR, (void*)&result);
	return result;
}

CString CclsButtonBorderStyle::GetXML(void)
{
	CString oReturn;
	InvokeHelper(0x9, DISPATCH_METHOD, VT_BSTR, (void*)&oReturn, NULL);
	return oReturn;
}

void CclsButtonBorderStyle::SetXML(LPCTSTR sXML)
{
	static BYTE params[] = VTS_BSTR;
	InvokeHelper(0xa, DISPATCH_METHOD, VT_EMPTY, NULL, params, sXML);
}

void CclsScrollBarStyle::SetArrowColor(unsigned long propval)
{
	SetProperty(0x1, VT_COLOR, propval);
}

unsigned long CclsScrollBarStyle::GetArrowColor()
{
	unsigned long result;
	GetProperty(0x1, VT_COLOR, (void*)&result);
	return result;
}

void CclsScrollBarStyle::SetDropShadowArrowColor(unsigned long propval)
{
	SetProperty(0x2, VT_COLOR, propval);
}

unsigned long CclsScrollBarStyle::GetDropShadowArrowColor()
{
	unsigned long result;
	GetProperty(0x2, VT_COLOR, (void*)&result);
	return result;
}

void CclsScrollBarStyle::SetDropShadow(BOOL propval)
{
	SetProperty(0x3, VT_BOOL, propval);
}

BOOL CclsScrollBarStyle::GetDropShadow()
{
	BOOL result;
	GetProperty(0x3, VT_BOOL, (void*)&result);
	return result;
}

void CclsScrollBarStyle::SetLeftX(LONG propval)
{
	SetProperty(0x4, VT_I4, propval);
}

LONG CclsScrollBarStyle::GetLeftX()
{
	LONG result;
	GetProperty(0x4, VT_I4, (void*)&result);
	return result;
}

void CclsScrollBarStyle::SetLeftY(LONG propval)
{
	SetProperty(0x5, VT_I4, propval);
}

LONG CclsScrollBarStyle::GetLeftY()
{
	LONG result;
	GetProperty(0x5, VT_I4, (void*)&result);
	return result;
}

void CclsScrollBarStyle::SetUpX(LONG propval)
{
	SetProperty(0x6, VT_I4, propval);
}

LONG CclsScrollBarStyle::GetUpX()
{
	LONG result;
	GetProperty(0x6, VT_I4, (void*)&result);
	return result;
}

void CclsScrollBarStyle::SetUpY(LONG propval)
{
	SetProperty(0x7, VT_I4, propval);
}

LONG CclsScrollBarStyle::GetUpY()
{
	LONG result;
	GetProperty(0x7, VT_I4, (void*)&result);
	return result;
}

void CclsScrollBarStyle::SetRightX(LONG propval)
{
	SetProperty(0x8, VT_I4, propval);
}

LONG CclsScrollBarStyle::GetRightX()
{
	LONG result;
	GetProperty(0x8, VT_I4, (void*)&result);
	return result;
}

void CclsScrollBarStyle::SetRightY(LONG propval)
{
	SetProperty(0x9, VT_I4, propval);
}

LONG CclsScrollBarStyle::GetRightY()
{
	LONG result;
	GetProperty(0x9, VT_I4, (void*)&result);
	return result;
}

void CclsScrollBarStyle::SetDownX(LONG propval)
{
	SetProperty(0xa, VT_I4, propval);
}

LONG CclsScrollBarStyle::GetDownX()
{
	LONG result;
	GetProperty(0xa, VT_I4, (void*)&result);
	return result;
}

void CclsScrollBarStyle::SetDownY(LONG propval)
{
	SetProperty(0xb, VT_I4, propval);
}

LONG CclsScrollBarStyle::GetDownY()
{
	LONG result;
	GetProperty(0xb, VT_I4, (void*)&result);
	return result;
}

void CclsScrollBarStyle::SetDropShadowLeftX(LONG propval)
{
	SetProperty(0xc, VT_I4, propval);
}

LONG CclsScrollBarStyle::GetDropShadowLeftX()
{
	LONG result;
	GetProperty(0xc, VT_I4, (void*)&result);
	return result;
}

void CclsScrollBarStyle::SetDropShadowLeftY(LONG propval)
{
	SetProperty(0xd, VT_I4, propval);
}

LONG CclsScrollBarStyle::GetDropShadowLeftY()
{
	LONG result;
	GetProperty(0xd, VT_I4, (void*)&result);
	return result;
}

void CclsScrollBarStyle::SetDropShadowUpX(LONG propval)
{
	SetProperty(0xe, VT_I4, propval);
}

LONG CclsScrollBarStyle::GetDropShadowUpX()
{
	LONG result;
	GetProperty(0xe, VT_I4, (void*)&result);
	return result;
}

void CclsScrollBarStyle::SetDropShadowUpY(LONG propval)
{
	SetProperty(0xf, VT_I4, propval);
}

LONG CclsScrollBarStyle::GetDropShadowUpY()
{
	LONG result;
	GetProperty(0xf, VT_I4, (void*)&result);
	return result;
}

void CclsScrollBarStyle::SetDropShadowRightX(LONG propval)
{
	SetProperty(0x10, VT_I4, propval);
}

LONG CclsScrollBarStyle::GetDropShadowRightX()
{
	LONG result;
	GetProperty(0x10, VT_I4, (void*)&result);
	return result;
}

void CclsScrollBarStyle::SetDropShadowRightY(LONG propval)
{
	SetProperty(0x11, VT_I4, propval);
}

LONG CclsScrollBarStyle::GetDropShadowRightY()
{
	LONG result;
	GetProperty(0x11, VT_I4, (void*)&result);
	return result;
}

void CclsScrollBarStyle::SetDropShadowDownX(LONG propval)
{
	SetProperty(0x12, VT_I4, propval);
}

LONG CclsScrollBarStyle::GetDropShadowDownX()
{
	LONG result;
	GetProperty(0x12, VT_I4, (void*)&result);
	return result;
}

void CclsScrollBarStyle::SetDropShadowDownY(LONG propval)
{
	SetProperty(0x13, VT_I4, propval);
}

LONG CclsScrollBarStyle::GetDropShadowDownY()
{
	LONG result;
	GetProperty(0x13, VT_I4, (void*)&result);
	return result;
}

void CclsScrollBarStyle::SetArrowSize(LONG propval)
{
	SetProperty(0x14, VT_I4, propval);
}

LONG CclsScrollBarStyle::GetArrowSize()
{
	LONG result;
	GetProperty(0x14, VT_I4, (void*)&result);
	return result;
}

CString CclsScrollBarStyle::GetXML(void)
{
	CString oReturn;
	InvokeHelper(0x15, DISPATCH_METHOD, VT_BSTR, (void*)&oReturn, NULL);
	return oReturn;
}

void CclsScrollBarStyle::SetXML(LPCTSTR sXML)
{
	static BYTE params[] = VTS_BSTR;
	InvokeHelper(0x16, DISPATCH_METHOD, VT_EMPTY, NULL, params, sXML);
}

void CclsSelectionRectangleStyle::SetOffsetBottom(LONG propval)
{
	SetProperty(0x1, VT_I4, propval);
}

LONG CclsSelectionRectangleStyle::GetOffsetBottom()
{
	LONG result;
	GetProperty(0x1, VT_I4, (void*)&result);
	return result;
}

void CclsSelectionRectangleStyle::SetOffsetLeft(LONG propval)
{
	SetProperty(0x2, VT_I4, propval);
}

LONG CclsSelectionRectangleStyle::GetOffsetLeft()
{
	LONG result;
	GetProperty(0x2, VT_I4, (void*)&result);
	return result;
}

void CclsSelectionRectangleStyle::SetOffsetRight(LONG propval)
{
	SetProperty(0x3, VT_I4, propval);
}

LONG CclsSelectionRectangleStyle::GetOffsetRight()
{
	LONG result;
	GetProperty(0x3, VT_I4, (void*)&result);
	return result;
}

void CclsSelectionRectangleStyle::SetOffsetTop(LONG propval)
{
	SetProperty(0x4, VT_I4, propval);
}

LONG CclsSelectionRectangleStyle::GetOffsetTop()
{
	LONG result;
	GetProperty(0x4, VT_I4, (void*)&result);
	return result;
}

void CclsSelectionRectangleStyle::SetVisible(BOOL propval)
{
	SetProperty(0x5, VT_BOOL, propval);
}

BOOL CclsSelectionRectangleStyle::GetVisible()
{
	BOOL result;
	GetProperty(0x5, VT_BOOL, (void*)&result);
	return result;
}

void CclsSelectionRectangleStyle::SetMode(E_SELECTIONRECTANGLEMODE propval)
{
	LONG lpropval = (LONG)propval;
	SetProperty(0x6, VT_I4, lpropval);
}

E_SELECTIONRECTANGLEMODE CclsSelectionRectangleStyle::GetMode()
{
	LONG result;
	GetProperty(0x6, VT_I4, (void*)&result);
	return (E_SELECTIONRECTANGLEMODE)result;
}

void CclsSelectionRectangleStyle::SetBorderWidth(LONG propval)
{
	SetProperty(0x7, VT_I4, propval);
}

LONG CclsSelectionRectangleStyle::GetBorderWidth()
{
	LONG result;
	GetProperty(0x7, VT_I4, (void*)&result);
	return result;
}

void CclsSelectionRectangleStyle::SetColor(unsigned long propval)
{
	SetProperty(0x8, VT_COLOR, propval);
}

unsigned long CclsSelectionRectangleStyle::GetColor()
{
	unsigned long result;
	GetProperty(0x8, VT_COLOR, (void*)&result);
	return result;
}

CString CclsSelectionRectangleStyle::GetXML(void)
{
	CString oReturn;
	InvokeHelper(0x9, DISPATCH_METHOD, VT_BSTR, (void*)&oReturn, NULL);
	return oReturn;
}

void CclsSelectionRectangleStyle::SetXML(LPCTSTR sXML)
{
	static BYTE params[] = VTS_BSTR;
	InvokeHelper(0xa, DISPATCH_METHOD, VT_EMPTY, NULL, params, sXML);
}

void CclsStyle::SetHatchBackColor(unsigned long propval)
{
	SetProperty(0x1, VT_COLOR, propval);
}

unsigned long CclsStyle::GetHatchBackColor()
{
	unsigned long result;
	GetProperty(0x1, VT_COLOR, (void*)&result);
	return result;
}

void CclsStyle::SetHatchForeColor(unsigned long propval)
{
	SetProperty(0x2, VT_COLOR, propval);
}

unsigned long CclsStyle::GetHatchForeColor()
{
	unsigned long result;
	GetProperty(0x2, VT_COLOR, (void*)&result);
	return result;
}

void CclsStyle::SetKey(LPCTSTR propval)
{
	SetProperty(0x3, VT_BSTR, propval);
}

CString CclsStyle::GetKey()
{
	CString result;
	GetProperty(0x3, VT_BSTR, (void*)&result);
	return result;
}

void CclsStyle::SetAppearance(E_STYLEAPPEARANCE propval)
{
	LONG lpropval = (LONG)propval;
	SetProperty(0x4, VT_I4, lpropval);
}

E_STYLEAPPEARANCE CclsStyle::GetAppearance()
{
	LONG result;
	GetProperty(0x4, VT_I4, (void*)&result);
	return (E_STYLEAPPEARANCE)result;
}

void CclsStyle::SetBackColor(unsigned long propval)
{
	SetProperty(0x5, VT_COLOR, propval);
}

unsigned long CclsStyle::GetBackColor()
{
	unsigned long result;
	GetProperty(0x5, VT_COLOR, (void*)&result);
	return result;
}

void CclsStyle::SetPattern(GRE_PATTERN propval)
{
	LONG lpropval = (LONG)propval;
	SetProperty(0x6, VT_I4, lpropval);
}

GRE_PATTERN CclsStyle::GetPattern()
{
	LONG result;
	GetProperty(0x6, VT_I4, (void*)&result);
	return (GRE_PATTERN)result;
}

void CclsStyle::SetBorderColor(unsigned long propval)
{
	SetProperty(0x7, VT_COLOR, propval);
}

unsigned long CclsStyle::GetBorderColor()
{
	unsigned long result;
	GetProperty(0x7, VT_COLOR, (void*)&result);
	return result;
}

void CclsStyle::SetBorderStyle(GRE_BORDERSTYLE propval)
{
	LONG lpropval = (LONG)propval;
	SetProperty(0x8, VT_I4, lpropval);
}

GRE_BORDERSTYLE CclsStyle::GetBorderStyle()
{
	LONG result;
	GetProperty(0x8, VT_I4, (void*)&result);
	return (GRE_BORDERSTYLE)result;
}

void CclsStyle::SetButtonStyle(GRE_BUTTONSTYLE propval)
{
	LONG lpropval = (LONG)propval;
	SetProperty(0x9, VT_I4, lpropval);
}

GRE_BUTTONSTYLE CclsStyle::GetButtonStyle()
{
	LONG result;
	GetProperty(0x9, VT_I4, (void*)&result);
	return (GRE_BUTTONSTYLE)result;
}

void CclsStyle::SetTextAlignmentHorizontal(GRE_HORIZONTALALIGNMENT propval)
{
	LONG lpropval = (LONG)propval;
	SetProperty(0xa, VT_I4, lpropval);
}

GRE_HORIZONTALALIGNMENT CclsStyle::GetTextAlignmentHorizontal()
{
	LONG result;
	GetProperty(0xa, VT_I4, (void*)&result);
	return (GRE_HORIZONTALALIGNMENT)result;
}

void CclsStyle::SetTextAlignmentVertical(GRE_VERTICALALIGNMENT propval)
{
	LONG lpropval = (LONG)propval;
	SetProperty(0xb, VT_I4, lpropval);
}

GRE_VERTICALALIGNMENT CclsStyle::GetTextAlignmentVertical()
{
	LONG result;
	GetProperty(0xb, VT_I4, (void*)&result);
	return (GRE_VERTICALALIGNMENT)result;
}

void CclsStyle::SetTextVisible(BOOL propval)
{
	SetProperty(0xc, VT_BOOL, propval);
}

BOOL CclsStyle::GetTextVisible()
{
	BOOL result;
	GetProperty(0xc, VT_BOOL, (void*)&result);
	return result;
}

void CclsStyle::SetTextXMargin(LONG propval)
{
	SetProperty(0xd, VT_I4, propval);
}

LONG CclsStyle::GetTextXMargin()
{
	LONG result;
	GetProperty(0xd, VT_I4, (void*)&result);
	return result;
}

void CclsStyle::SetTextYMargin(LONG propval)
{
	SetProperty(0xe, VT_I4, propval);
}

LONG CclsStyle::GetTextYMargin()
{
	LONG result;
	GetProperty(0xe, VT_I4, (void*)&result);
	return result;
}

void CclsStyle::SetClipText(BOOL propval)
{
	SetProperty(0xf, VT_BOOL, propval);
}

BOOL CclsStyle::GetClipText()
{
	BOOL result;
	GetProperty(0xf, VT_BOOL, (void*)&result);
	return result;
}

CclsFont CclsStyle::GetFont()
{
	LPDISPATCH pDispatch;
	GetProperty(0x10, VT_DISPATCH, (void*)&pDispatch);
	return CclsFont(pDispatch);
}

void CclsStyle::SetForeColor(unsigned long propval)
{
	SetProperty(0x11, VT_COLOR, propval);
}

unsigned long CclsStyle::GetForeColor()
{
	unsigned long result;
	GetProperty(0x11, VT_COLOR, (void*)&result);
	return result;
}

void CclsStyle::SetOffsetBottom(LONG propval)
{
	SetProperty(0x12, VT_I4, propval);
}

LONG CclsStyle::GetOffsetBottom()
{
	LONG result;
	GetProperty(0x12, VT_I4, (void*)&result);
	return result;
}

void CclsStyle::SetOffsetTop(LONG propval)
{
	SetProperty(0x13, VT_I4, propval);
}

LONG CclsStyle::GetOffsetTop()
{
	LONG result;
	GetProperty(0x13, VT_I4, (void*)&result);
	return result;
}

void CclsStyle::SetImageAlignmentHorizontal(GRE_HORIZONTALALIGNMENT propval)
{
	LONG lpropval = (LONG)propval;
	SetProperty(0x14, VT_I4, lpropval);
}

GRE_HORIZONTALALIGNMENT CclsStyle::GetImageAlignmentHorizontal()
{
	LONG result;
	GetProperty(0x14, VT_I4, (void*)&result);
	return (GRE_HORIZONTALALIGNMENT)result;
}

void CclsStyle::SetImageAlignmentVertical(GRE_VERTICALALIGNMENT propval)
{
	LONG lpropval = (LONG)propval;
	SetProperty(0x15, VT_I4, lpropval);
}

GRE_VERTICALALIGNMENT CclsStyle::GetImageAlignmentVertical()
{
	LONG result;
	GetProperty(0x15, VT_I4, (void*)&result);
	return (GRE_VERTICALALIGNMENT)result;
}

void CclsStyle::SetImageXMargin(LONG propval)
{
	SetProperty(0x16, VT_I4, propval);
}

LONG CclsStyle::GetImageXMargin()
{
	LONG result;
	GetProperty(0x16, VT_I4, (void*)&result);
	return result;
}

void CclsStyle::SetImageYMargin(LONG propval)
{
	SetProperty(0x17, VT_I4, propval);
}

LONG CclsStyle::GetImageYMargin()
{
	LONG result;
	GetProperty(0x17, VT_I4, (void*)&result);
	return result;
}

void CclsStyle::SetPlacement(E_PLACEMENT propval)
{
	LONG lpropval = (LONG)propval;
	SetProperty(0x18, VT_I4, lpropval);
}

E_PLACEMENT CclsStyle::GetPlacement()
{
	LONG result;
	GetProperty(0x18, VT_I4, (void*)&result);
	return (E_PLACEMENT)result;
}

void CclsStyle::SetUseMask(BOOL propval)
{
	SetProperty(0x19, VT_BOOL, propval);
}

BOOL CclsStyle::GetUseMask()
{
	BOOL result;
	GetProperty(0x19, VT_BOOL, (void*)&result);
	return result;
}

void CclsStyle::SetTextPlacement(E_TEXTPLACEMENT propval)
{
	LONG lpropval = (LONG)propval;
	SetProperty(0x1a, VT_I4, lpropval);
}

E_TEXTPLACEMENT CclsStyle::GetTextPlacement()
{
	LONG result;
	GetProperty(0x1a, VT_I4, (void*)&result);
	return (E_TEXTPLACEMENT)result;
}

void CclsStyle::SetPatternFactor(LONG propval)
{
	SetProperty(0x1b, VT_I4, propval);
}

LONG CclsStyle::GetPatternFactor()
{
	LONG result;
	GetProperty(0x1b, VT_I4, (void*)&result);
	return result;
}

void CclsStyle::SetHatchStyle(GRE_HATCHSTYLE propval)
{
	LONG lpropval = (LONG)propval;
	SetProperty(0x1c, VT_I4, lpropval);
}

GRE_HATCHSTYLE CclsStyle::GetHatchStyle()
{
	LONG result;
	GetProperty(0x1c, VT_I4, (void*)&result);
	return (GRE_HATCHSTYLE)result;
}

void CclsStyle::SetStartGradientColor(unsigned long propval)
{
	SetProperty(0x1d, VT_COLOR, propval);
}

unsigned long CclsStyle::GetStartGradientColor()
{
	unsigned long result;
	GetProperty(0x1d, VT_COLOR, (void*)&result);
	return result;
}

void CclsStyle::SetEndGradientColor(unsigned long propval)
{
	SetProperty(0x1e, VT_COLOR, propval);
}

unsigned long CclsStyle::GetEndGradientColor()
{
	unsigned long result;
	GetProperty(0x1e, VT_COLOR, (void*)&result);
	return result;
}

void CclsStyle::SetGradientFillMode(GRE_GRADIENTFILLMODE propval)
{
	LONG lpropval = (LONG)propval;
	SetProperty(0x1f, VT_I4, lpropval);
}

GRE_GRADIENTFILLMODE CclsStyle::GetGradientFillMode()
{
	LONG result;
	GetProperty(0x1f, VT_I4, (void*)&result);
	return (GRE_GRADIENTFILLMODE)result;
}

void CclsStyle::SetFillMode(GRE_FILLMODE propval)
{
	LONG lpropval = (LONG)propval;
	SetProperty(0x20, VT_I4, lpropval);
}

GRE_FILLMODE CclsStyle::GetFillMode()
{
	LONG result;
	GetProperty(0x20, VT_I4, (void*)&result);
	return (GRE_FILLMODE)result;
}

void CclsStyle::SetBackgroundMode(GRE_BACKGROUNDMODE propval)
{
	LONG lpropval = (LONG)propval;
	SetProperty(0x21, VT_I4, lpropval);
}

GRE_BACKGROUNDMODE CclsStyle::GetBackgroundMode()
{
	LONG result;
	GetProperty(0x21, VT_I4, (void*)&result);
	return (GRE_BACKGROUNDMODE)result;
}

void CclsStyle::SetDrawTextInVisibleArea(BOOL propval)
{
	SetProperty(0x22, VT_BOOL, propval);
}

BOOL CclsStyle::GetDrawTextInVisibleArea()
{
	BOOL result;
	GetProperty(0x22, VT_BOOL, (void*)&result);
	return result;
}

void CclsStyle::SetTag(LPCTSTR propval)
{
	SetProperty(0x23, VT_BSTR, propval);
}

CString CclsStyle::GetTag()
{
	CString result;
	GetProperty(0x23, VT_BSTR, (void*)&result);
	return result;
}

CclsTaskStyle CclsStyle::GetTaskStyle()
{
	LPDISPATCH pDispatch;
	GetProperty(0x24, VT_DISPATCH, (void*)&pDispatch);
	return CclsTaskStyle(pDispatch);
}

CclsMilestoneStyle CclsStyle::GetMilestoneStyle()
{
	LPDISPATCH pDispatch;
	GetProperty(0x25, VT_DISPATCH, (void*)&pDispatch);
	return CclsMilestoneStyle(pDispatch);
}

CclsPredecessorStyle CclsStyle::GetPredecessorStyle()
{
	LPDISPATCH pDispatch;
	GetProperty(0x26, VT_DISPATCH, (void*)&pDispatch);
	return CclsPredecessorStyle(pDispatch);
}

CclsTextFlags CclsStyle::GetTextFlags()
{
	LPDISPATCH pDispatch;
	GetProperty(0x27, VT_DISPATCH, (void*)&pDispatch);
	return CclsTextFlags(pDispatch);
}

CclsCustomBorderStyle CclsStyle::GetCustomBorderStyle()
{
	LPDISPATCH pDispatch;
	GetProperty(0x28, VT_DISPATCH, (void*)&pDispatch);
	return CclsCustomBorderStyle(pDispatch);
}

LONG CclsStyle::GetIndex()
{
	LONG result;
	GetProperty(0x2b, VT_I4, (void*)&result);
	return result;
}

void CclsStyle::SetBorderWidth(LONG propval)
{
	SetProperty(0x2c, VT_I4, propval);
}

LONG CclsStyle::GetBorderWidth()
{
	LONG result;
	GetProperty(0x2c, VT_I4, (void*)&result);
	return result;
}

CclsScrollBarStyle CclsStyle::GetScrollBarStyle()
{
	LPDISPATCH pDispatch;
	GetProperty(0x2d, VT_DISPATCH, (void*)&pDispatch);
	return CclsScrollBarStyle(pDispatch);
}

CclsSelectionRectangleStyle CclsStyle::GetSelectionRectangleStyle()
{
	LPDISPATCH pDispatch;
	GetProperty(0x2e, VT_DISPATCH, (void*)&pDispatch);
	return CclsSelectionRectangleStyle(pDispatch);
}

CclsButtonBorderStyle CclsStyle::GetButtonBorderStyle()
{
	LPDISPATCH pDispatch;
	GetProperty(0x2f, VT_DISPATCH, (void*)&pDispatch);
	return CclsButtonBorderStyle(pDispatch);
}

void CclsStyle::SetTextEditBackColor(unsigned long propval)
{
	SetProperty(0x30, VT_COLOR, propval);
}

unsigned long CclsStyle::GetTextEditBackColor()
{
	unsigned long result;
	GetProperty(0x30, VT_COLOR, (void*)&result);
	return result;
}

void CclsStyle::SetTextEditForeColor(unsigned long propval)
{
	SetProperty(0x31, VT_COLOR, propval);
}

unsigned long CclsStyle::GetTextEditForeColor()
{
	unsigned long result;
	GetProperty(0x31, VT_COLOR, (void*)&result);
	return result;
}

CString CclsStyle::GetXML(void)
{
	CString oReturn;
	InvokeHelper(0x29, DISPATCH_METHOD, VT_BSTR, (void*)&oReturn, NULL);
	return oReturn;
}

void CclsStyle::SetXML(LPCTSTR sXML)
{
	static BYTE params[] = VTS_BSTR;
	InvokeHelper(0x2a, DISPATCH_METHOD, VT_EMPTY, NULL, params, sXML);
}

void CclsStyle::Clear(void)
{
	InvokeHelper(0x32, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
}

CclsStyle CclsStyle::Clone(LPCTSTR Key)
{
	CclsStyle oReturn;
	static BYTE params[] = VTS_BSTR;
	InvokeHelper(0x33, DISPATCH_METHOD, VT_DISPATCH, (void*)&oReturn, params, Key);
	return oReturn;
}

LONG CclsStyles::GetCount()
{
	LONG result;
	GetProperty(0x1, VT_I4, (void*)&result);
	return result;
}

CclsStyle CclsStyles::Item(LPCTSTR Index)
{
	CclsStyle oReturn;
	static BYTE params[] = VTS_BSTR;
	InvokeHelper(0x2, DISPATCH_METHOD, VT_DISPATCH, (void*)&oReturn, params, Index);
	return oReturn;
}

CclsStyle CclsStyles::Add(LPCTSTR Key)
{
	CclsStyle oReturn;
	static BYTE params[] = VTS_BSTR;
	InvokeHelper(0x3, DISPATCH_METHOD, VT_DISPATCH, (void*)&oReturn, params, Key);
	return oReturn;
}

void CclsStyles::Clear(void)
{
	InvokeHelper(0x4, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
}

void CclsStyles::Remove(LPCTSTR Index)
{
	static BYTE params[] = VTS_BSTR;
	InvokeHelper(0x5, DISPATCH_METHOD, VT_EMPTY, NULL, params, Index);
}

CString CclsStyles::GetXML(void)
{
	CString oReturn;
	InvokeHelper(0x6, DISPATCH_METHOD, VT_BSTR, (void*)&oReturn, NULL);
	return oReturn;
}

void CclsStyles::SetXML(LPCTSTR sXML)
{
	static BYTE params[] = VTS_BSTR;
	InvokeHelper(0x7, DISPATCH_METHOD, VT_EMPTY, NULL, params, sXML);
}

BOOL CclsStyles::ContainsKey(LPCTSTR Key)
{
	BOOL oReturn;
	static BYTE params[] = VTS_BSTR;
	InvokeHelper(0x8, DISPATCH_METHOD, VT_BOOL, (void*)&oReturn, params, Key);
	return oReturn;
}

void CclsButtonState::SetNormalStyleIndex(LPCTSTR propval)
{
	SetProperty(0x1, VT_BSTR, propval);
}

CString CclsButtonState::GetNormalStyleIndex()
{
	CString result;
	GetProperty(0x1, VT_BSTR, (void*)&result);
	return result;
}

CclsStyle CclsButtonState::GetNormalStyle()
{
	LPDISPATCH pDispatch;
	GetProperty(0x2, VT_DISPATCH, (void*)&pDispatch);
	return CclsStyle(pDispatch);
}

void CclsButtonState::SetPressedStyleIndex(LPCTSTR propval)
{
	SetProperty(0x3, VT_BSTR, propval);
}

CString CclsButtonState::GetPressedStyleIndex()
{
	CString result;
	GetProperty(0x3, VT_BSTR, (void*)&result);
	return result;
}

CclsStyle CclsButtonState::GetPressedStyle()
{
	LPDISPATCH pDispatch;
	GetProperty(0x4, VT_DISPATCH, (void*)&pDispatch);
	return CclsStyle(pDispatch);
}

void CclsButtonState::SetDisabledStyleIndex(LPCTSTR propval)
{
	SetProperty(0x5, VT_BSTR, propval);
}

CString CclsButtonState::GetDisabledStyleIndex()
{
	CString result;
	GetProperty(0x5, VT_BSTR, (void*)&result);
	return result;
}

CclsStyle CclsButtonState::GetDisabledStyle()
{
	LPDISPATCH pDispatch;
	GetProperty(0x6, VT_DISPATCH, (void*)&pDispatch);
	return CclsStyle(pDispatch);
}

CString CclsButtonState::GetXML(void)
{
	CString oReturn;
	InvokeHelper(0x7, DISPATCH_METHOD, VT_BSTR, (void*)&oReturn, NULL);
	return oReturn;
}

void CclsButtonState::SetXML(LPCTSTR sXML)
{
	static BYTE params[] = VTS_BSTR;
	InvokeHelper(0x8, DISPATCH_METHOD, VT_EMPTY, NULL, params, sXML);
}

void CclsVScrollBarTemplate::SetTimerInterval(LONG propval)
{
	SetProperty(0x1, VT_I4, propval);
}

LONG CclsVScrollBarTemplate::GetTimerInterval()
{
	LONG result;
	GetProperty(0x1, VT_I4, (void*)&result);
	return result;
}

void CclsVScrollBarTemplate::SetStyleIndex(LPCTSTR propval)
{
	SetProperty(0x2, VT_BSTR, propval);
}

CString CclsVScrollBarTemplate::GetStyleIndex()
{
	CString result;
	GetProperty(0x2, VT_BSTR, (void*)&result);
	return result;
}

CclsStyle CclsVScrollBarTemplate::GetStyle()
{
	LPDISPATCH pDispatch;
	GetProperty(0x3, VT_DISPATCH, (void*)&pDispatch);
	return CclsStyle(pDispatch);
}

CclsButtonState CclsVScrollBarTemplate::GetArrowButtons()
{
	LPDISPATCH pDispatch;
	GetProperty(0x4, VT_DISPATCH, (void*)&pDispatch);
	return CclsButtonState(pDispatch);
}

CclsButtonState CclsVScrollBarTemplate::GetThumbButton()
{
	LPDISPATCH pDispatch;
	GetProperty(0x5, VT_DISPATCH, (void*)&pDispatch);
	return CclsButtonState(pDispatch);
}

CString CclsVScrollBarTemplate::GetXML(void)
{
	CString oReturn;
	InvokeHelper(0x6, DISPATCH_METHOD, VT_BSTR, (void*)&oReturn, NULL);
	return oReturn;
}

void CclsVScrollBarTemplate::SetXML(LPCTSTR sXML)
{
	static BYTE params[] = VTS_BSTR;
	InvokeHelper(0x7, DISPATCH_METHOD, VT_EMPTY, NULL, params, sXML);
}

void CclsHScrollBarTemplate::SetTimerInterval(LONG propval)
{
	SetProperty(0x1, VT_I4, propval);
}

LONG CclsHScrollBarTemplate::GetTimerInterval()
{
	LONG result;
	GetProperty(0x1, VT_I4, (void*)&result);
	return result;
}

void CclsHScrollBarTemplate::SetStyleIndex(LPCTSTR propval)
{
	SetProperty(0x2, VT_BSTR, propval);
}

CString CclsHScrollBarTemplate::GetStyleIndex()
{
	CString result;
	GetProperty(0x2, VT_BSTR, (void*)&result);
	return result;
}

CclsStyle CclsHScrollBarTemplate::GetStyle()
{
	LPDISPATCH pDispatch;
	GetProperty(0x3, VT_DISPATCH, (void*)&pDispatch);
	return CclsStyle(pDispatch);
}

CclsButtonState CclsHScrollBarTemplate::GetArrowButtons()
{
	LPDISPATCH pDispatch;
	GetProperty(0x4, VT_DISPATCH, (void*)&pDispatch);
	return CclsButtonState(pDispatch);
}

CclsButtonState CclsHScrollBarTemplate::GetThumbButton()
{
	LPDISPATCH pDispatch;
	GetProperty(0x5, VT_DISPATCH, (void*)&pDispatch);
	return CclsButtonState(pDispatch);
}

CString CclsHScrollBarTemplate::GetXML(void)
{
	CString oReturn;
	InvokeHelper(0x6, DISPATCH_METHOD, VT_BSTR, (void*)&oReturn, NULL);
	return oReturn;
}

void CclsHScrollBarTemplate::SetXML(LPCTSTR sXML)
{
	static BYTE params[] = VTS_BSTR;
	InvokeHelper(0x7, DISPATCH_METHOD, VT_EMPTY, NULL, params, sXML);
}

void CclsScrollBarSeparator::SetStyleIndex(LPCTSTR propval)
{
	SetProperty(0x1, VT_BSTR, propval);
}

CString CclsScrollBarSeparator::GetStyleIndex()
{
	CString result;
	GetProperty(0x1, VT_BSTR, (void*)&result);
	return result;
}

CclsStyle CclsScrollBarSeparator::GetStyle()
{
	LPDISPATCH pDispatch;
	GetProperty(0x2, VT_DISPATCH, (void*)&pDispatch);
	return CclsStyle(pDispatch);
}

CString CclsScrollBarSeparator::GetXML(void)
{
	CString oReturn;
	InvokeHelper(0x3, DISPATCH_METHOD, VT_BSTR, (void*)&oReturn, NULL);
	return oReturn;
}

void CclsScrollBarSeparator::SetXML(LPCTSTR sXML)
{
	static BYTE params[] = VTS_BSTR;
	InvokeHelper(0x4, DISPATCH_METHOD, VT_EMPTY, NULL, params, sXML);
}

LONG CclsVerticalScrollBar::GetMin()
{
	LONG result;
	GetProperty(0x1, VT_I4, (void*)&result);
	return result;
}

LONG CclsVerticalScrollBar::GetMax()
{
	LONG result;
	GetProperty(0x2, VT_I4, (void*)&result);
	return result;
}

void CclsVerticalScrollBar::SetValue(LONG propval)
{
	SetProperty(0x3, VT_I4, propval);
}

LONG CclsVerticalScrollBar::GetValue()
{
	LONG result;
	GetProperty(0x3, VT_I4, (void*)&result);
	return result;
}

void CclsVerticalScrollBar::SetVisible(BOOL propval)
{
	SetProperty(0x4, VT_BOOL, propval);
}

BOOL CclsVerticalScrollBar::GetVisible()
{
	BOOL result;
	GetProperty(0x4, VT_BOOL, (void*)&result);
	return result;
}

void CclsVerticalScrollBar::SetSmallChange(LONG propval)
{
	SetProperty(0x5, VT_I4, propval);
}

LONG CclsVerticalScrollBar::GetSmallChange()
{
	LONG result;
	GetProperty(0x5, VT_I4, (void*)&result);
	return result;
}

void CclsVerticalScrollBar::SetLargeChange(LONG propval)
{
	SetProperty(0x6, VT_I4, propval);
}

LONG CclsVerticalScrollBar::GetLargeChange()
{
	LONG result;
	GetProperty(0x6, VT_I4, (void*)&result);
	return result;
}

CclsVScrollBarTemplate CclsVerticalScrollBar::GetScrollBar()
{
	LPDISPATCH pDispatch;
	GetProperty(0x7, VT_DISPATCH, (void*)&pDispatch);
	return CclsVScrollBarTemplate(pDispatch);
}

CString CclsVerticalScrollBar::GetXML(void)
{
	CString oReturn;
	InvokeHelper(0x8, DISPATCH_METHOD, VT_BSTR, (void*)&oReturn, NULL);
	return oReturn;
}

void CclsVerticalScrollBar::SetXML(LPCTSTR sXML)
{
	static BYTE params[] = VTS_BSTR;
	InvokeHelper(0x9, DISPATCH_METHOD, VT_EMPTY, NULL, params, sXML);
}

LONG CclsHorizontalScrollBar::GetMin()
{
	LONG result;
	GetProperty(0x1, VT_I4, (void*)&result);
	return result;
}

LONG CclsHorizontalScrollBar::GetMax()
{
	LONG result;
	GetProperty(0x2, VT_I4, (void*)&result);
	return result;
}

void CclsHorizontalScrollBar::SetValue(LONG propval)
{
	SetProperty(0x3, VT_I4, propval);
}

LONG CclsHorizontalScrollBar::GetValue()
{
	LONG result;
	GetProperty(0x3, VT_I4, (void*)&result);
	return result;
}

void CclsHorizontalScrollBar::SetVisible(BOOL propval)
{
	SetProperty(0x4, VT_BOOL, propval);
}

BOOL CclsHorizontalScrollBar::GetVisible()
{
	BOOL result;
	GetProperty(0x4, VT_BOOL, (void*)&result);
	return result;
}

void CclsHorizontalScrollBar::SetSmallChange(LONG propval)
{
	SetProperty(0x5, VT_I4, propval);
}

LONG CclsHorizontalScrollBar::GetSmallChange()
{
	LONG result;
	GetProperty(0x5, VT_I4, (void*)&result);
	return result;
}

void CclsHorizontalScrollBar::SetLargeChange(LONG propval)
{
	SetProperty(0x6, VT_I4, propval);
}

LONG CclsHorizontalScrollBar::GetLargeChange()
{
	LONG result;
	GetProperty(0x6, VT_I4, (void*)&result);
	return result;
}

CclsHScrollBarTemplate CclsHorizontalScrollBar::GetScrollBar()
{
	LPDISPATCH pDispatch;
	GetProperty(0x7, VT_DISPATCH, (void*)&pDispatch);
	return CclsHScrollBarTemplate(pDispatch);
}

CString CclsHorizontalScrollBar::GetXML(void)
{
	CString oReturn;
	InvokeHelper(0x8, DISPATCH_METHOD, VT_BSTR, (void*)&oReturn, NULL);
	return oReturn;
}

void CclsHorizontalScrollBar::SetXML(LPCTSTR sXML)
{
	static BYTE params[] = VTS_BSTR;
	InvokeHelper(0x9, DISPATCH_METHOD, VT_EMPTY, NULL, params, sXML);
}

void CclsSplitter::SetAppearance(E_BORDERSTYLE propval)
{
	LONG lpropval = (LONG)propval;
	SetProperty(0x1, VT_I4, lpropval);
}

E_BORDERSTYLE CclsSplitter::GetAppearance()
{
	LONG result;
	GetProperty(0x1, VT_I4, (void*)&result);
	return (E_BORDERSTYLE)result;
}

void CclsSplitter::SetPosition(LONG propval)
{
	SetProperty(0x2, VT_I4, propval);
}

LONG CclsSplitter::GetPosition()
{
	LONG result;
	GetProperty(0x2, VT_I4, (void*)&result);
	return result;
}

void CclsSplitter::SetType(E_SPLITTERTYPE propval)
{
	LONG lpropval = (LONG)propval;
	SetProperty(0x7, VT_I4, lpropval);
}

E_SPLITTERTYPE CclsSplitter::GetType()
{
	LONG result;
	GetProperty(0x7, VT_I4, (void*)&result);
	return (E_SPLITTERTYPE)result;
}

void CclsSplitter::SetWidth(LONG propval)
{
	SetProperty(0x8, VT_I4, propval);
}

LONG CclsSplitter::GetWidth()
{
	LONG result;
	GetProperty(0x8, VT_I4, (void*)&result);
	return result;
}

void CclsSplitter::SetStyleIndex(LPCTSTR propval)
{
	SetProperty(0x9, VT_BSTR, propval);
}

CString CclsSplitter::GetStyleIndex()
{
	CString result;
	GetProperty(0x9, VT_BSTR, (void*)&result);
	return result;
}

CclsStyle CclsSplitter::GetStyle()
{
	LPDISPATCH pDispatch;
	GetProperty(0xa, VT_DISPATCH, (void*)&pDispatch);
	return CclsStyle(pDispatch);
}

void CclsSplitter::SetOffset(LONG propval)
{
	SetProperty(0xb, VT_I4, propval);
}

LONG CclsSplitter::GetOffset()
{
	LONG result;
	GetProperty(0xb, VT_I4, (void*)&result);
	return result;
}

CString CclsSplitter::GetXML(void)
{
	CString oReturn;
	InvokeHelper(0x3, DISPATCH_METHOD, VT_BSTR, (void*)&oReturn, NULL);
	return oReturn;
}

void CclsSplitter::SetXML(LPCTSTR sXML)
{
	static BYTE params[] = VTS_BSTR;
	InvokeHelper(0x4, DISPATCH_METHOD, VT_EMPTY, NULL, params, sXML);
}

void CclsSplitter::SetColor(LONG Index, unsigned long oColor)
{
	static BYTE params[] = VTS_I4 VTS_COLOR;
	InvokeHelper(0x5, DISPATCH_METHOD, VT_EMPTY, NULL, params, Index, oColor);
}

unsigned long CclsSplitter::GetColor(LONG Index)
{
	unsigned long oReturn;
	static BYTE params[] = VTS_I4;
	InvokeHelper(0x6, DISPATCH_METHOD, VT_COLOR, (void*)&oReturn, params, Index);
	return oReturn;
}

void CclsTreeview::SetFirstVisibleNode(LONG propval)
{
	SetProperty(0x1, VT_I4, propval);
}

LONG CclsTreeview::GetFirstVisibleNode()
{
	LONG result;
	GetProperty(0x1, VT_I4, (void*)&result);
	return result;
}

LONG CclsTreeview::GetLastVisibleNode()
{
	LONG result;
	GetProperty(0x2, VT_I4, (void*)&result);
	return result;
}

void CclsTreeview::SetIndentation(LONG propval)
{
	SetProperty(0x3, VT_I4, propval);
}

LONG CclsTreeview::GetIndentation()
{
	LONG result;
	GetProperty(0x3, VT_I4, (void*)&result);
	return result;
}

void CclsTreeview::SetCheckBoxBorderColor(unsigned long propval)
{
	SetProperty(0x4, VT_COLOR, propval);
}

unsigned long CclsTreeview::GetCheckBoxBorderColor()
{
	unsigned long result;
	GetProperty(0x4, VT_COLOR, (void*)&result);
	return result;
}

void CclsTreeview::SetCheckBoxBackColor(unsigned long propval)
{
	SetProperty(0x5, VT_COLOR, propval);
}

unsigned long CclsTreeview::GetCheckBoxBackColor()
{
	unsigned long result;
	GetProperty(0x5, VT_COLOR, (void*)&result);
	return result;
}

void CclsTreeview::SetCheckBoxMarkColor(unsigned long propval)
{
	SetProperty(0x6, VT_COLOR, propval);
}

unsigned long CclsTreeview::GetCheckBoxMarkColor()
{
	unsigned long result;
	GetProperty(0x6, VT_COLOR, (void*)&result);
	return result;
}

void CclsTreeview::SetBackColor(unsigned long propval)
{
	SetProperty(0x7, VT_COLOR, propval);
}

unsigned long CclsTreeview::GetBackColor()
{
	unsigned long result;
	GetProperty(0x7, VT_COLOR, (void*)&result);
	return result;
}

void CclsTreeview::SetPathSeparator(LPCTSTR propval)
{
	SetProperty(0x8, VT_BSTR, propval);
}

CString CclsTreeview::GetPathSeparator()
{
	CString result;
	GetProperty(0x8, VT_BSTR, (void*)&result);
	return result;
}

void CclsTreeview::SetTreeLines(BOOL propval)
{
	SetProperty(0x9, VT_BOOL, propval);
}

BOOL CclsTreeview::GetTreeLines()
{
	BOOL result;
	GetProperty(0x9, VT_BOOL, (void*)&result);
	return result;
}

void CclsTreeview::SetTreeviewMode(E_TREEVIEWMODE propval)
{
	LONG lpropval = (LONG)propval;
	SetProperty(0xa, VT_I4, lpropval);
}

E_TREEVIEWMODE CclsTreeview::GetTreeviewMode()
{
	LONG result;
	GetProperty(0xa, VT_I4, (void*)&result);
	return (E_TREEVIEWMODE)result;
}

void CclsTreeview::SetImages(BOOL propval)
{
	SetProperty(0xb, VT_BOOL, propval);
}

BOOL CclsTreeview::GetImages()
{
	BOOL result;
	GetProperty(0xb, VT_BOOL, (void*)&result);
	return result;
}

void CclsTreeview::SetCheckBoxes(BOOL propval)
{
	SetProperty(0xc, VT_BOOL, propval);
}

BOOL CclsTreeview::GetCheckBoxes()
{
	BOOL result;
	GetProperty(0xc, VT_BOOL, (void*)&result);
	return result;
}

void CclsTreeview::SetFullColumnSelect(BOOL propval)
{
	SetProperty(0xd, VT_BOOL, propval);
}

BOOL CclsTreeview::GetFullColumnSelect()
{
	BOOL result;
	GetProperty(0xd, VT_BOOL, (void*)&result);
	return result;
}

void CclsTreeview::SetExpansionOnSelection(BOOL propval)
{
	SetProperty(0xe, VT_BOOL, propval);
}

BOOL CclsTreeview::GetExpansionOnSelection()
{
	BOOL result;
	GetProperty(0xe, VT_BOOL, (void*)&result);
	return result;
}

void CclsTreeview::SetSelectedBackColor(unsigned long propval)
{
	SetProperty(0xf, VT_COLOR, propval);
}

unsigned long CclsTreeview::GetSelectedBackColor()
{
	unsigned long result;
	GetProperty(0xf, VT_COLOR, (void*)&result);
	return result;
}

void CclsTreeview::SetSelectedForeColor(unsigned long propval)
{
	SetProperty(0x10, VT_COLOR, propval);
}

unsigned long CclsTreeview::GetSelectedForeColor()
{
	unsigned long result;
	GetProperty(0x10, VT_COLOR, (void*)&result);
	return result;
}

void CclsTreeview::SetTreeLineColor(unsigned long propval)
{
	SetProperty(0x11, VT_COLOR, propval);
}

unsigned long CclsTreeview::GetTreeLineColor()
{
	unsigned long result;
	GetProperty(0x11, VT_COLOR, (void*)&result);
	return result;
}

void CclsTreeview::SetPlusMinusBorderColor(unsigned long propval)
{
	SetProperty(0x12, VT_COLOR, propval);
}

unsigned long CclsTreeview::GetPlusMinusBorderColor()
{
	unsigned long result;
	GetProperty(0x12, VT_COLOR, (void*)&result);
	return result;
}

void CclsTreeview::SetPlusMinusSignColor(unsigned long propval)
{
	SetProperty(0x13, VT_COLOR, propval);
}

unsigned long CclsTreeview::GetPlusMinusSignColor()
{
	unsigned long result;
	GetProperty(0x13, VT_COLOR, (void*)&result);
	return result;
}

void CclsTreeview::SetCheckBoxBackgroundMode(GRE_BACKGROUNDMODE propval)
{
	LONG lpropval = (LONG)propval;
	SetProperty(0x17, VT_I4, lpropval);
}

GRE_BACKGROUNDMODE CclsTreeview::GetCheckBoxBackgroundMode()
{
	LONG result;
	GetProperty(0x17, VT_I4, (void*)&result);
	return (GRE_BACKGROUNDMODE)result;
}

void CclsTreeview::SetPlusMinusBackColor(unsigned long propval)
{
	SetProperty(0x18, VT_COLOR, propval);
}

unsigned long CclsTreeview::GetPlusMinusBackColor()
{
	unsigned long result;
	GetProperty(0x18, VT_COLOR, (void*)&result);
	return result;
}

void CclsTreeview::SetPlusMinusBackgroundMode(GRE_BACKGROUNDMODE propval)
{
	LONG lpropval = (LONG)propval;
	SetProperty(0x19, VT_I4, lpropval);
}

GRE_BACKGROUNDMODE CclsTreeview::GetPlusMinusBackgroundMode()
{
	LONG result;
	GetProperty(0x19, VT_I4, (void*)&result);
	return (GRE_BACKGROUNDMODE)result;
}

void CclsTreeview::SetExpandIconForeColor(unsigned long propval)
{
	SetProperty(0x1a, VT_COLOR, propval);
}

unsigned long CclsTreeview::GetExpandIconForeColor()
{
	unsigned long result;
	GetProperty(0x1a, VT_COLOR, (void*)&result);
	return result;
}

void CclsTreeview::SetExpandIconBackColor(unsigned long propval)
{
	SetProperty(0x1b, VT_COLOR, propval);
}

unsigned long CclsTreeview::GetExpandIconBackColor()
{
	unsigned long result;
	GetProperty(0x1b, VT_COLOR, (void*)&result);
	return result;
}

void CclsTreeview::SetExpandIconDropShadowColor(unsigned long propval)
{
	SetProperty(0x1c, VT_COLOR, propval);
}

unsigned long CclsTreeview::GetExpandIconDropShadowColor()
{
	unsigned long result;
	GetProperty(0x1c, VT_COLOR, (void*)&result);
	return result;
}

void CclsTreeview::SetCollapseIconForeColor(unsigned long propval)
{
	SetProperty(0x1d, VT_COLOR, propval);
}

unsigned long CclsTreeview::GetCollapseIconForeColor()
{
	unsigned long result;
	GetProperty(0x1d, VT_COLOR, (void*)&result);
	return result;
}

void CclsTreeview::SetCollapseIconDropShadowColor(unsigned long propval)
{
	SetProperty(0x1e, VT_COLOR, propval);
}

unsigned long CclsTreeview::GetCollapseIconDropShadowColor()
{
	unsigned long result;
	GetProperty(0x1e, VT_COLOR, (void*)&result);
	return result;
}

void CclsTreeview::ClearSelections(void)
{
	InvokeHelper(0x14, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
}

void CclsTreeview::SetXML(LPCTSTR sXML)
{
	static BYTE params[] = VTS_BSTR;
	InvokeHelper(0x15, DISPATCH_METHOD, VT_EMPTY, NULL, params, sXML);
}

CString CclsTreeview::GetXML(void)
{
	CString oReturn;
	InvokeHelper(0x16, DISPATCH_METHOD, VT_BSTR, (void*)&oReturn, NULL);
	return oReturn;
}

CclsFont CclsToolTip::GetFont()
{
	LPDISPATCH pDispatch;
	GetProperty(0x1, VT_DISPATCH, (void*)&pDispatch);
	return CclsFont(pDispatch);
}

void CclsToolTip::SetBackColor(unsigned long propval)
{
	SetProperty(0x2, VT_COLOR, propval);
}

unsigned long CclsToolTip::GetBackColor()
{
	unsigned long result;
	GetProperty(0x2, VT_COLOR, (void*)&result);
	return result;
}

void CclsToolTip::SetForeColor(unsigned long propval)
{
	SetProperty(0x3, VT_COLOR, propval);
}

unsigned long CclsToolTip::GetForeColor()
{
	unsigned long result;
	GetProperty(0x3, VT_COLOR, (void*)&result);
	return result;
}

void CclsToolTip::SetBorderColor(unsigned long propval)
{
	SetProperty(0x4, VT_COLOR, propval);
}

unsigned long CclsToolTip::GetBorderColor()
{
	unsigned long result;
	GetProperty(0x4, VT_COLOR, (void*)&result);
	return result;
}

void CclsToolTip::SetText(LPCTSTR propval)
{
	SetProperty(0x5, VT_BSTR, propval);
}

CString CclsToolTip::GetText()
{
	CString result;
	GetProperty(0x5, VT_BSTR, (void*)&result);
	return result;
}

void CclsToolTip::SetAutomaticSizing(BOOL propval)
{
	SetProperty(0x6, VT_BOOL, propval);
}

BOOL CclsToolTip::GetAutomaticSizing()
{
	BOOL result;
	GetProperty(0x6, VT_BOOL, (void*)&result);
	return result;
}

void CclsToolTip::SetLeft(LONG propval)
{
	SetProperty(0x7, VT_I4, propval);
}

LONG CclsToolTip::GetLeft()
{
	LONG result;
	GetProperty(0x7, VT_I4, (void*)&result);
	return result;
}

LONG CclsToolTip::GetRight()
{
	LONG result;
	GetProperty(0x8, VT_I4, (void*)&result);
	return result;
}

void CclsToolTip::SetTop(LONG propval)
{
	SetProperty(0x9, VT_I4, propval);
}

LONG CclsToolTip::GetTop()
{
	LONG result;
	GetProperty(0x9, VT_I4, (void*)&result);
	return result;
}

LONG CclsToolTip::GetBottom()
{
	LONG result;
	GetProperty(0xa, VT_I4, (void*)&result);
	return result;
}

void CclsToolTip::SetWidth(LONG propval)
{
	SetProperty(0xb, VT_I4, propval);
}

LONG CclsToolTip::GetWidth()
{
	LONG result;
	GetProperty(0xb, VT_I4, (void*)&result);
	return result;
}

void CclsToolTip::SetHeight(LONG propval)
{
	SetProperty(0xc, VT_I4, propval);
}

LONG CclsToolTip::GetHeight()
{
	LONG result;
	GetProperty(0xc, VT_I4, (void*)&result);
	return result;
}

void CclsToolTip::SetVisible(BOOL propval)
{
	SetProperty(0xd, VT_BOOL, propval);
}

BOOL CclsToolTip::GetVisible()
{
	BOOL result;
	GetProperty(0xd, VT_BOOL, (void*)&result);
	return result;
}

void CclsLayer::SetKey(LPCTSTR propval)
{
	SetProperty(0x1, VT_BSTR, propval);
}

CString CclsLayer::GetKey()
{
	CString result;
	GetProperty(0x1, VT_BSTR, (void*)&result);
	return result;
}

void CclsLayer::SetVisible(BOOL propval)
{
	SetProperty(0x2, VT_BOOL, propval);
}

BOOL CclsLayer::GetVisible()
{
	BOOL result;
	GetProperty(0x2, VT_BOOL, (void*)&result);
	return result;
}

void CclsLayer::SetTag(LPCTSTR propval)
{
	SetProperty(0x3, VT_BSTR, propval);
}

CString CclsLayer::GetTag()
{
	CString result;
	GetProperty(0x3, VT_BSTR, (void*)&result);
	return result;
}

LONG CclsLayer::GetIndex()
{
	LONG result;
	GetProperty(0x6, VT_I4, (void*)&result);
	return result;
}

CString CclsLayer::GetXML(void)
{
	CString oReturn;
	InvokeHelper(0x4, DISPATCH_METHOD, VT_BSTR, (void*)&oReturn, NULL);
	return oReturn;
}

void CclsLayer::SetXML(LPCTSTR sXML)
{
	static BYTE params[] = VTS_BSTR;
	InvokeHelper(0x5, DISPATCH_METHOD, VT_EMPTY, NULL, params, sXML);
}

LONG CclsLayers::GetCount()
{
	LONG result;
	GetProperty(0x1, VT_I4, (void*)&result);
	return result;
}

CclsLayer CclsLayers::Item(LPCTSTR Index)
{
	CclsLayer oReturn;
	static BYTE params[] = VTS_BSTR;
	InvokeHelper(0x2, DISPATCH_METHOD, VT_DISPATCH, (void*)&oReturn, params, Index);
	return oReturn;
}

CclsLayer CclsLayers::Add(LPCTSTR Key, BOOL Visible)
{
	CclsLayer oReturn;
	static BYTE params[] = VTS_BSTR VTS_BOOL;
	InvokeHelper(0x3, DISPATCH_METHOD, VT_DISPATCH, (void*)&oReturn, params, Key, Visible);
	return oReturn;
}

void CclsLayers::Clear(void)
{
	InvokeHelper(0x4, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
}

void CclsLayers::Remove(LPCTSTR Index)
{
	static BYTE params[] = VTS_BSTR;
	InvokeHelper(0x5, DISPATCH_METHOD, VT_EMPTY, NULL, params, Index);
}

CString CclsLayers::GetXML(void)
{
	CString oReturn;
	InvokeHelper(0x6, DISPATCH_METHOD, VT_BSTR, (void*)&oReturn, NULL);
	return oReturn;
}

void CclsLayers::SetXML(LPCTSTR sXML)
{
	static BYTE params[] = VTS_BSTR;
	InvokeHelper(0x7, DISPATCH_METHOD, VT_EMPTY, NULL, params, sXML);
}

LPDISPATCH CclsNode::GetRow()
{
	LPDISPATCH result;
	GetProperty(0x1, VT_DISPATCH, (void*)&result);
	return result;
}

void CclsNode::SetCheckBoxVisible(BOOL propval)
{
	SetProperty(0x2, VT_BOOL, propval);
}

BOOL CclsNode::GetCheckBoxVisible()
{
	BOOL result;
	GetProperty(0x2, VT_BOOL, (void*)&result);
	return result;
}

void CclsNode::SetImageVisible(BOOL propval)
{
	SetProperty(0x3, VT_BOOL, propval);
}

BOOL CclsNode::GetImageVisible()
{
	BOOL result;
	GetProperty(0x3, VT_BOOL, (void*)&result);
	return result;
}

LONG CclsNode::GetLeft()
{
	LONG result;
	GetProperty(0x4, VT_I4, (void*)&result);
	return result;
}

LONG CclsNode::GetTop()
{
	LONG result;
	GetProperty(0x5, VT_I4, (void*)&result);
	return result;
}

LONG CclsNode::GetBottom()
{
	LONG result;
	GetProperty(0x6, VT_I4, (void*)&result);
	return result;
}

void CclsNode::SetDepth(LONG propval)
{
	SetProperty(0x7, VT_I4, propval);
}

LONG CclsNode::GetDepth()
{
	LONG result;
	GetProperty(0x7, VT_I4, (void*)&result);
	return result;
}

void CclsNode::SetExpanded(BOOL propval)
{
	SetProperty(0x8, VT_BOOL, propval);
}

BOOL CclsNode::GetExpanded()
{
	BOOL result;
	GetProperty(0x8, VT_BOOL, (void*)&result);
	return result;
}

void CclsNode::SetSelected(BOOL propval)
{
	SetProperty(0x9, VT_BOOL, propval);
}

BOOL CclsNode::GetSelected()
{
	BOOL result;
	GetProperty(0x9, VT_BOOL, (void*)&result);
	return result;
}

CclsImage CclsNode::GetExpandedImage()
{
	LPDISPATCH pDispatch;
	GetProperty(0xa, VT_DISPATCH, (void*)&pDispatch);
	return CclsImage(pDispatch);
}

CclsImage CclsNode::GetSelectedImage()
{
	LPDISPATCH pDispatch;
	GetProperty(0xb, VT_DISPATCH, (void*)&pDispatch);
	return CclsImage(pDispatch);
}

CclsImage CclsNode::GetImage()
{
	LPDISPATCH pDispatch;
	GetProperty(0xc, VT_DISPATCH, (void*)&pDispatch);
	return CclsImage(pDispatch);
}

void CclsNode::SetTag(LPCTSTR propval)
{
	SetProperty(0xd, VT_BSTR, propval);
}

CString CclsNode::GetTag()
{
	CString result;
	GetProperty(0xd, VT_BSTR, (void*)&result);
	return result;
}

void CclsNode::SetText(LPCTSTR propval)
{
	SetProperty(0xe, VT_BSTR, propval);
}

CString CclsNode::GetText()
{
	CString result;
	GetProperty(0xe, VT_BSTR, (void*)&result);
	return result;
}

void CclsNode::SetChecked(BOOL propval)
{
	SetProperty(0xf, VT_BOOL, propval);
}

BOOL CclsNode::GetChecked()
{
	BOOL result;
	GetProperty(0xf, VT_BOOL, (void*)&result);
	return result;
}

void CclsNode::SetHeight(LONG propval)
{
	SetProperty(0x10, VT_I4, propval);
}

LONG CclsNode::GetHeight()
{
	LONG result;
	GetProperty(0x10, VT_I4, (void*)&result);
	return result;
}

BOOL CclsNode::GetHidden()
{
	BOOL result;
	GetProperty(0x11, VT_BOOL, (void*)&result);
	return result;
}

void CclsNode::SetExpandedImageTag(LPCTSTR propval)
{
	SetProperty(0x1f, VT_BSTR, propval);
}

CString CclsNode::GetExpandedImageTag()
{
	CString result;
	GetProperty(0x1f, VT_BSTR, (void*)&result);
	return result;
}

void CclsNode::SetSelectedImageTag(LPCTSTR propval)
{
	SetProperty(0x20, VT_BSTR, propval);
}

CString CclsNode::GetSelectedImageTag()
{
	CString result;
	GetProperty(0x20, VT_BSTR, (void*)&result);
	return result;
}

void CclsNode::SetImageTag(LPCTSTR propval)
{
	SetProperty(0x21, VT_BSTR, propval);
}

CString CclsNode::GetImageTag()
{
	CString result;
	GetProperty(0x21, VT_BSTR, (void*)&result);
	return result;
}

void CclsNode::SetStyleIndex(LPCTSTR propval)
{
	SetProperty(0x22, VT_BSTR, propval);
}

CString CclsNode::GetStyleIndex()
{
	CString result;
	GetProperty(0x22, VT_BSTR, (void*)&result);
	return result;
}

CclsStyle CclsNode::GetStyle()
{
	LPDISPATCH pDispatch;
	GetProperty(0x23, VT_DISPATCH, (void*)&pDispatch);
	return CclsStyle(pDispatch);
}

void CclsNode::SetAllowTextEdit(BOOL propval)
{
	SetProperty(0x24, VT_BOOL, propval);
}

BOOL CclsNode::GetAllowTextEdit()
{
	BOOL result;
	GetProperty(0x24, VT_BOOL, (void*)&result);
	return result;
}

CclsNode CclsNode::NextSibling(void)
{
	CclsNode oReturn;
	InvokeHelper(0x12, DISPATCH_METHOD, VT_DISPATCH, (void*)&oReturn, NULL);
	return oReturn;
}

CclsNode CclsNode::PreviousSibling(void)
{
	CclsNode oReturn;
	InvokeHelper(0x13, DISPATCH_METHOD, VT_DISPATCH, (void*)&oReturn, NULL);
	return oReturn;
}

BOOL CclsNode::IsFirstSibling(void)
{
	BOOL oReturn;
	InvokeHelper(0x14, DISPATCH_METHOD, VT_BOOL, (void*)&oReturn, NULL);
	return oReturn;
}

CclsNode CclsNode::FirstSibling(void)
{
	CclsNode oReturn;
	InvokeHelper(0x15, DISPATCH_METHOD, VT_DISPATCH, (void*)&oReturn, NULL);
	return oReturn;
}

BOOL CclsNode::IsLastSibling(void)
{
	BOOL oReturn;
	InvokeHelper(0x16, DISPATCH_METHOD, VT_BOOL, (void*)&oReturn, NULL);
	return oReturn;
}

CclsNode CclsNode::LastSibling(void)
{
	CclsNode oReturn;
	InvokeHelper(0x17, DISPATCH_METHOD, VT_DISPATCH, (void*)&oReturn, NULL);
	return oReturn;
}

CclsNode CclsNode::Child(void)
{
	CclsNode oReturn;
	InvokeHelper(0x18, DISPATCH_METHOD, VT_DISPATCH, (void*)&oReturn, NULL);
	return oReturn;
}

CclsNode CclsNode::Parent(void)
{
	CclsNode oReturn;
	InvokeHelper(0x19, DISPATCH_METHOD, VT_DISPATCH, (void*)&oReturn, NULL);
	return oReturn;
}

BOOL CclsNode::IsRoot(void)
{
	BOOL oReturn;
	InvokeHelper(0x1a, DISPATCH_METHOD, VT_BOOL, (void*)&oReturn, NULL);
	return oReturn;
}

CString CclsNode::FullPath(void)
{
	CString oReturn;
	InvokeHelper(0x1b, DISPATCH_METHOD, VT_BSTR, (void*)&oReturn, NULL);
	return oReturn;
}

LONG CclsNode::Children(void)
{
	LONG oReturn;
	InvokeHelper(0x1c, DISPATCH_METHOD, VT_I4, (void*)&oReturn, NULL);
	return oReturn;
}

CString CclsNode::GetXML(void)
{
	CString oReturn;
	InvokeHelper(0x1d, DISPATCH_METHOD, VT_BSTR, (void*)&oReturn, NULL);
	return oReturn;
}

void CclsNode::SetXML(LPCTSTR sXML)
{
	static BYTE params[] = VTS_BSTR;
	InvokeHelper(0x1e, DISPATCH_METHOD, VT_EMPTY, NULL, params, sXML);
}

CclsRow CclsCell::GetRow()
{
	LPDISPATCH pDispatch;
	GetProperty(0x1, VT_DISPATCH, (void*)&pDispatch);
	return CclsRow(pDispatch);
}

void CclsCell::SetKey(LPCTSTR propval)
{
	SetProperty(0x2, VT_BSTR, propval);
}

CString CclsCell::GetKey()
{
	CString result;
	GetProperty(0x2, VT_BSTR, (void*)&result);
	return result;
}

void CclsCell::SetText(LPCTSTR propval)
{
	SetProperty(0x3, VT_BSTR, propval);
}

CString CclsCell::GetText()
{
	CString result;
	GetProperty(0x3, VT_BSTR, (void*)&result);
	return result;
}

CclsImage CclsCell::GetImage()
{
	LPDISPATCH pDispatch;
	GetProperty(0x4, VT_DISPATCH, (void*)&pDispatch);
	return CclsImage(pDispatch);
}

void CclsCell::SetStyleIndex(LPCTSTR propval)
{
	SetProperty(0x5, VT_BSTR, propval);
}

CString CclsCell::GetStyleIndex()
{
	CString result;
	GetProperty(0x5, VT_BSTR, (void*)&result);
	return result;
}

CclsStyle CclsCell::GetStyle()
{
	LPDISPATCH pDispatch;
	GetProperty(0x6, VT_DISPATCH, (void*)&pDispatch);
	return CclsStyle(pDispatch);
}

void CclsCell::SetTag(LPCTSTR propval)
{
	SetProperty(0x7, VT_BSTR, propval);
}

CString CclsCell::GetTag()
{
	CString result;
	GetProperty(0x7, VT_BSTR, (void*)&result);
	return result;
}

CString CclsCell::GetRowKey()
{
	CString result;
	GetProperty(0x8, VT_BSTR, (void*)&result);
	return result;
}

LONG CclsCell::GetLeft()
{
	LONG result;
	GetProperty(0x9, VT_I4, (void*)&result);
	return result;
}

LONG CclsCell::GetTop()
{
	LONG result;
	GetProperty(0xa, VT_I4, (void*)&result);
	return result;
}

LONG CclsCell::GetRight()
{
	LONG result;
	GetProperty(0xb, VT_I4, (void*)&result);
	return result;
}

LONG CclsCell::GetBottom()
{
	LONG result;
	GetProperty(0xc, VT_I4, (void*)&result);
	return result;
}

LONG CclsCell::GetLeftTrim()
{
	LONG result;
	GetProperty(0xd, VT_I4, (void*)&result);
	return result;
}

LONG CclsCell::GetRightTrim()
{
	LONG result;
	GetProperty(0xe, VT_I4, (void*)&result);
	return result;
}

LONG CclsCell::GetIndex()
{
	LONG result;
	GetProperty(0x11, VT_I4, (void*)&result);
	return result;
}

void CclsCell::SetImageTag(LPCTSTR propval)
{
	SetProperty(0x12, VT_BSTR, propval);
}

CString CclsCell::GetImageTag()
{
	CString result;
	GetProperty(0x12, VT_BSTR, (void*)&result);
	return result;
}

void CclsCell::SetAllowTextEdit(BOOL propval)
{
	SetProperty(0x13, VT_BOOL, propval);
}

BOOL CclsCell::GetAllowTextEdit()
{
	BOOL result;
	GetProperty(0x13, VT_BOOL, (void*)&result);
	return result;
}

CString CclsCell::GetXML(void)
{
	CString oReturn;
	InvokeHelper(0xf, DISPATCH_METHOD, VT_BSTR, (void*)&oReturn, NULL);
	return oReturn;
}

void CclsCell::SetXML(LPCTSTR sXML)
{
	static BYTE params[] = VTS_BSTR;
	InvokeHelper(0x10, DISPATCH_METHOD, VT_EMPTY, NULL, params, sXML);
}

LONG CclsCells::GetCount()
{
	LONG result;
	GetProperty(0x1, VT_I4, (void*)&result);
	return result;
}

CclsCell CclsCells::Item(LPCTSTR Index)
{
	CclsCell oReturn;
	static BYTE params[] = VTS_BSTR;
	InvokeHelper(0x2, DISPATCH_METHOD, VT_DISPATCH, (void*)&oReturn, params, Index);
	return oReturn;
}

CString CclsCells::GetXML(void)
{
	CString oReturn;
	InvokeHelper(0x3, DISPATCH_METHOD, VT_BSTR, (void*)&oReturn, NULL);
	return oReturn;
}

void CclsCells::SetXML(LPCTSTR sXML)
{
	static BYTE params[] = VTS_BSTR;
	InvokeHelper(0x4, DISPATCH_METHOD, VT_EMPTY, NULL, params, sXML);
}

void CclsRow::SetAllowMove(BOOL propval)
{
	SetProperty(0x1, VT_BOOL, propval);
}

BOOL CclsRow::GetAllowMove()
{
	BOOL result;
	GetProperty(0x1, VT_BOOL, (void*)&result);
	return result;
}

void CclsRow::SetAllowSize(BOOL propval)
{
	SetProperty(0x2, VT_BOOL, propval);
}

BOOL CclsRow::GetAllowSize()
{
	BOOL result;
	GetProperty(0x2, VT_BOOL, (void*)&result);
	return result;
}

void CclsRow::SetKey(LPCTSTR propval)
{
	SetProperty(0x3, VT_BSTR, propval);
}

CString CclsRow::GetKey()
{
	CString result;
	GetProperty(0x3, VT_BSTR, (void*)&result);
	return result;
}

void CclsRow::SetContainer(BOOL propval)
{
	SetProperty(0x4, VT_BOOL, propval);
}

BOOL CclsRow::GetContainer()
{
	BOOL result;
	GetProperty(0x4, VT_BOOL, (void*)&result);
	return result;
}

void CclsRow::SetUseNodeImages(BOOL propval)
{
	SetProperty(0x5, VT_BOOL, propval);
}

BOOL CclsRow::GetUseNodeImages()
{
	BOOL result;
	GetProperty(0x5, VT_BOOL, (void*)&result);
	return result;
}

void CclsRow::SetMergeCells(BOOL propval)
{
	SetProperty(0x6, VT_BOOL, propval);
}

BOOL CclsRow::GetMergeCells()
{
	BOOL result;
	GetProperty(0x6, VT_BOOL, (void*)&result);
	return result;
}

void CclsRow::SetHeight(LONG propval)
{
	SetProperty(0x7, VT_I4, propval);
}

LONG CclsRow::GetHeight()
{
	LONG result;
	GetProperty(0x7, VT_I4, (void*)&result);
	return result;
}

void CclsRow::SetText(LPCTSTR propval)
{
	SetProperty(0x8, VT_BSTR, propval);
}

CString CclsRow::GetText()
{
	CString result;
	GetProperty(0x8, VT_BSTR, (void*)&result);
	return result;
}

CclsImage CclsRow::GetImage()
{
	LPDISPATCH pDispatch;
	GetProperty(0x9, VT_DISPATCH, (void*)&pDispatch);
	return CclsImage(pDispatch);
}

void CclsRow::SetStyleIndex(LPCTSTR propval)
{
	SetProperty(0xa, VT_BSTR, propval);
}

CString CclsRow::GetStyleIndex()
{
	CString result;
	GetProperty(0xa, VT_BSTR, (void*)&result);
	return result;
}

CclsStyle CclsRow::GetStyle()
{
	LPDISPATCH pDispatch;
	GetProperty(0xb, VT_DISPATCH, (void*)&pDispatch);
	return CclsStyle(pDispatch);
}

void CclsRow::SetTag(LPCTSTR propval)
{
	SetProperty(0xc, VT_BSTR, propval);
}

CString CclsRow::GetTag()
{
	CString result;
	GetProperty(0xc, VT_BSTR, (void*)&result);
	return result;
}

void CclsRow::SetClientAreaStyleIndex(LPCTSTR propval)
{
	SetProperty(0xd, VT_BSTR, propval);
}

CString CclsRow::GetClientAreaStyleIndex()
{
	CString result;
	GetProperty(0xd, VT_BSTR, (void*)&result);
	return result;
}

CclsStyle CclsRow::GetClientAreaStyle()
{
	LPDISPATCH pDispatch;
	GetProperty(0xe, VT_DISPATCH, (void*)&pDispatch);
	return CclsStyle(pDispatch);
}

LONG CclsRow::GetLeft()
{
	LONG result;
	GetProperty(0xf, VT_I4, (void*)&result);
	return result;
}

LONG CclsRow::GetTop()
{
	LONG result;
	GetProperty(0x10, VT_I4, (void*)&result);
	return result;
}

LONG CclsRow::GetRight()
{
	LONG result;
	GetProperty(0x11, VT_I4, (void*)&result);
	return result;
}

LONG CclsRow::GetBottom()
{
	LONG result;
	GetProperty(0x12, VT_I4, (void*)&result);
	return result;
}

BOOL CclsRow::GetVisible()
{
	BOOL result;
	GetProperty(0x13, VT_BOOL, (void*)&result);
	return result;
}

CclsNode CclsRow::GetNode()
{
	LPDISPATCH pDispatch;
	GetProperty(0x14, VT_DISPATCH, (void*)&pDispatch);
	return CclsNode(pDispatch);
}

CclsCells CclsRow::GetCells()
{
	LPDISPATCH pDispatch;
	GetProperty(0x15, VT_DISPATCH, (void*)&pDispatch);
	return CclsCells(pDispatch);
}

LONG CclsRow::GetIndex()
{
	LONG result;
	GetProperty(0x18, VT_I4, (void*)&result);
	return result;
}

void CclsRow::SetImageTag(LPCTSTR propval)
{
	SetProperty(0x19, VT_BSTR, propval);
}

CString CclsRow::GetImageTag()
{
	CString result;
	GetProperty(0x19, VT_BSTR, (void*)&result);
	return result;
}

void CclsRow::SetAllowTextEdit(BOOL propval)
{
	SetProperty(0x1a, VT_BOOL, propval);
}

BOOL CclsRow::GetAllowTextEdit()
{
	BOOL result;
	GetProperty(0x1a, VT_BOOL, (void*)&result);
	return result;
}

void CclsRow::SetHighlighted(BOOL propval)
{
	SetProperty(0x1b, VT_BOOL, propval);
}

BOOL CclsRow::GetHighlighted()
{
	BOOL result;
	GetProperty(0x1b, VT_BOOL, (void*)&result);
	return result;
}

CString CclsRow::GetXML(void)
{
	CString oReturn;
	InvokeHelper(0x16, DISPATCH_METHOD, VT_BSTR, (void*)&oReturn, NULL);
	return oReturn;
}

void CclsRow::SetXML(LPCTSTR sXML)
{
	static BYTE params[] = VTS_BSTR;
	InvokeHelper(0x17, DISPATCH_METHOD, VT_EMPTY, NULL, params, sXML);
}

LONG CclsRows::GetCount()
{
	LONG result;
	GetProperty(0x1, VT_I4, (void*)&result);
	return result;
}

CclsRow CclsRows::Item(LPCTSTR Index)
{
	CclsRow oReturn;
	static BYTE params[] = VTS_BSTR;
	InvokeHelper(0x2, DISPATCH_METHOD, VT_DISPATCH, (void*)&oReturn, params, Index);
	return oReturn;
}

void CclsRows::Clear(void)
{
	InvokeHelper(0x3, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
}

void CclsRows::Remove(LPCTSTR Index)
{
	static BYTE params[] = VTS_BSTR;
	InvokeHelper(0x4, DISPATCH_METHOD, VT_EMPTY, NULL, params, Index);
}

void CclsRows::MoveRow(LONG OriginRowIndex, LONG DestRowIndex)
{
	static BYTE params[] = VTS_I4 VTS_I4;
	InvokeHelper(0x5, DISPATCH_METHOD, VT_EMPTY, NULL, params, OriginRowIndex, DestRowIndex);
}

CString CclsRows::GetXML(void)
{
	CString oReturn;
	InvokeHelper(0x6, DISPATCH_METHOD, VT_BSTR, (void*)&oReturn, NULL);
	return oReturn;
}

void CclsRows::SetXML(LPCTSTR sXML)
{
	static BYTE params[] = VTS_BSTR;
	InvokeHelper(0x7, DISPATCH_METHOD, VT_EMPTY, NULL, params, sXML);
}

CclsRow CclsRows::Add(LPCTSTR Key, LPCTSTR Text, BOOL MergeCells, BOOL Container, LPCTSTR StyleIndex)
{
	CclsRow oReturn;
	static BYTE params[] = VTS_BSTR VTS_BSTR VTS_BOOL VTS_BOOL VTS_BSTR;
	InvokeHelper(0x8, DISPATCH_METHOD, VT_DISPATCH, (void*)&oReturn, params, Key, Text, MergeCells, Container, StyleIndex);
	return oReturn;
}

void CclsRows::SortRows(LPCTSTR PropertyName, BOOL Descending, E_SORTTYPE SortType, LONG StartIndex, LONG EndIndex)
{
	static BYTE params[] = VTS_BSTR VTS_BOOL VTS_I4 VTS_I4 VTS_I4;
	InvokeHelper(0x9, DISPATCH_METHOD, VT_EMPTY, NULL, params, PropertyName, Descending, SortType, StartIndex, EndIndex);
}

void CclsRows::SortCells(LONG CellIndex, LPCTSTR PropertyName, BOOL Descending, E_SORTTYPE SortType, LONG StartIndex, LONG EndIndex)
{
	static BYTE params[] = VTS_I4 VTS_BSTR VTS_BOOL VTS_I4 VTS_I4 VTS_I4;
	InvokeHelper(0xa, DISPATCH_METHOD, VT_EMPTY, NULL, params, CellIndex, PropertyName, Descending, SortType, StartIndex, EndIndex);
}

void CclsRows::UpdateTree(void)
{
	InvokeHelper(0xb, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
}

void CclsRows::BeginLoad(BOOL Preserve)
{
	static BYTE params[] = VTS_BOOL;
	InvokeHelper(0xc, DISPATCH_METHOD, VT_EMPTY, NULL, params, Preserve);
}

CclsRow CclsRows::Load(LPCTSTR sKey)
{
	CclsRow oReturn;
	static BYTE params[] = VTS_BSTR;
	InvokeHelper(0xd, DISPATCH_METHOD, VT_DISPATCH, (void*)&oReturn, params, sKey);
	return oReturn;
}

void CclsRows::EndLoad(void)
{
	InvokeHelper(0xe, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
}

BOOL CclsRows::VerifyTree(void)
{
	BOOL oReturn;
	InvokeHelper(0xf, DISPATCH_METHOD, VT_BOOL, (void*)&oReturn, NULL);
	return oReturn;
}

void CclsPredecessor::SetKey(LPCTSTR propval)
{
	SetProperty(0x1, VT_BSTR, propval);
}

CString CclsPredecessor::GetKey()
{
	CString result;
	GetProperty(0x1, VT_BSTR, (void*)&result);
	return result;
}

void CclsPredecessor::SetVisible(BOOL propval)
{
	SetProperty(0x2, VT_BOOL, propval);
}

BOOL CclsPredecessor::GetVisible()
{
	BOOL result;
	GetProperty(0x2, VT_BOOL, (void*)&result);
	return result;
}

void CclsPredecessor::SetPredecessorKey(LPCTSTR propval)
{
	SetProperty(0x3, VT_BSTR, propval);
}

CString CclsPredecessor::GetPredecessorKey()
{
	CString result;
	GetProperty(0x3, VT_BSTR, (void*)&result);
	return result;
}

CclsTask CclsPredecessor::GetPredecessorTask()
{
	LPDISPATCH pDispatch;
	GetProperty(0x4, VT_DISPATCH, (void*)&pDispatch);
	return CclsTask(pDispatch);
}

void CclsPredecessor::SetPredecessorType(E_CONSTRAINTTYPE propval)
{
	LONG lpropval = (LONG)propval;
	SetProperty(0x5, VT_I4, lpropval);
}

E_CONSTRAINTTYPE CclsPredecessor::GetPredecessorType()
{
	LONG result;
	GetProperty(0x5, VT_I4, (void*)&result);
	return (E_CONSTRAINTTYPE)result;
}

void CclsPredecessor::SetStyleIndex(LPCTSTR propval)
{
	SetProperty(0x6, VT_BSTR, propval);
}

CString CclsPredecessor::GetStyleIndex()
{
	CString result;
	GetProperty(0x6, VT_BSTR, (void*)&result);
	return result;
}

CclsStyle CclsPredecessor::GetStyle()
{
	LPDISPATCH pDispatch;
	GetProperty(0x7, VT_DISPATCH, (void*)&pDispatch);
	return CclsStyle(pDispatch);
}

void CclsPredecessor::SetTag(LPCTSTR propval)
{
	SetProperty(0x8, VT_BSTR, propval);
}

CString CclsPredecessor::GetTag()
{
	CString result;
	GetProperty(0x8, VT_BSTR, (void*)&result);
	return result;
}

LONG CclsPredecessor::GetIndex()
{
	LONG result;
	GetProperty(0xb, VT_I4, (void*)&result);
	return result;
}

void CclsPredecessor::SetSuccessorKey(LPCTSTR propval)
{
	SetProperty(0xc, VT_BSTR, propval);
}

CString CclsPredecessor::GetSuccessorKey()
{
	CString result;
	GetProperty(0xc, VT_BSTR, (void*)&result);
	return result;
}

CclsTask CclsPredecessor::GetSuccessorTask()
{
	LPDISPATCH pDispatch;
	GetProperty(0xd, VT_DISPATCH, (void*)&pDispatch);
	return CclsTask(pDispatch);
}

void CclsPredecessor::SetLagInterval(E_INTERVAL propval)
{
	LONG lpropval = (LONG)propval;
	SetProperty(0xe, VT_I4, lpropval);
}

E_INTERVAL CclsPredecessor::GetLagInterval()
{
	LONG result;
	GetProperty(0xe, VT_I4, (void*)&result);
	return (E_INTERVAL)result;
}

void CclsPredecessor::SetLagFactor(LONG propval)
{
	SetProperty(0xf, VT_I4, propval);
}

LONG CclsPredecessor::GetLagFactor()
{
	LONG result;
	GetProperty(0xf, VT_I4, (void*)&result);
	return result;
}

BOOL CclsPredecessor::GetWarning()
{
	BOOL result;
	GetProperty(0x10, VT_BOOL, (void*)&result);
	return result;
}

void CclsPredecessor::SetWarningStyleIndex(LPCTSTR propval)
{
	SetProperty(0x11, VT_BSTR, propval);
}

CString CclsPredecessor::GetWarningStyleIndex()
{
	CString result;
	GetProperty(0x11, VT_BSTR, (void*)&result);
	return result;
}

CclsStyle CclsPredecessor::GetWarningStyle()
{
	LPDISPATCH pDispatch;
	GetProperty(0x12, VT_DISPATCH, (void*)&pDispatch);
	return CclsStyle(pDispatch);
}

void CclsPredecessor::SetSelectedStyleIndex(LPCTSTR propval)
{
	SetProperty(0x13, VT_BSTR, propval);
}

CString CclsPredecessor::GetSelectedStyleIndex()
{
	CString result;
	GetProperty(0x13, VT_BSTR, (void*)&result);
	return result;
}

CclsStyle CclsPredecessor::GetSelectedStyle()
{
	LPDISPATCH pDispatch;
	GetProperty(0x14, VT_DISPATCH, (void*)&pDispatch);
	return CclsStyle(pDispatch);
}

CString CclsPredecessor::GetXML(void)
{
	CString oReturn;
	InvokeHelper(0x9, DISPATCH_METHOD, VT_BSTR, (void*)&oReturn, NULL);
	return oReturn;
}

void CclsPredecessor::SetXML(LPCTSTR sXML)
{
	static BYTE params[] = VTS_BSTR;
	InvokeHelper(0xa, DISPATCH_METHOD, VT_EMPTY, NULL, params, sXML);
}

LONG CclsPredecessors::GetCount()
{
	LONG result;
	GetProperty(0x1, VT_I4, (void*)&result);
	return result;
}

CclsPredecessor CclsPredecessors::Item(LPCTSTR Index)
{
	CclsPredecessor oReturn;
	static BYTE params[] = VTS_BSTR;
	InvokeHelper(0x2, DISPATCH_METHOD, VT_DISPATCH, (void*)&oReturn, params, Index);
	return oReturn;
}

CclsPredecessor CclsPredecessors::Add(LPCTSTR SuccessorKey, LPCTSTR PredecessorKey, E_CONSTRAINTTYPE PredecessorType, LPCTSTR Key, LPCTSTR StyleIndex)
{
	CclsPredecessor oReturn;
	static BYTE params[] = VTS_BSTR VTS_BSTR VTS_I4 VTS_BSTR VTS_BSTR;
	InvokeHelper(0x3, DISPATCH_METHOD, VT_DISPATCH, (void*)&oReturn, params, SuccessorKey, PredecessorKey, PredecessorType, Key, StyleIndex);
	return oReturn;
}

void CclsPredecessors::Clear(void)
{
	InvokeHelper(0x4, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
}

void CclsPredecessors::Remove(LPCTSTR Index)
{
	static BYTE params[] = VTS_BSTR;
	InvokeHelper(0x5, DISPATCH_METHOD, VT_EMPTY, NULL, params, Index);
}

CString CclsPredecessors::GetXML(void)
{
	CString oReturn;
	InvokeHelper(0x6, DISPATCH_METHOD, VT_BSTR, (void*)&oReturn, NULL);
	return oReturn;
}

void CclsPredecessors::SetXML(LPCTSTR sXML)
{
	static BYTE params[] = VTS_BSTR;
	InvokeHelper(0x7, DISPATCH_METHOD, VT_EMPTY, NULL, params, sXML);
}

void CclsTask::SetKey(LPCTSTR propval)
{
	SetProperty(0x1, VT_BSTR, propval);
}

CString CclsTask::GetKey()
{
	CString result;
	GetProperty(0x1, VT_BSTR, (void*)&result);
	return result;
}

void CclsTask::SetIncomingPredecessors(BOOL propval)
{
	SetProperty(0x2, VT_BOOL, propval);
}

BOOL CclsTask::GetIncomingPredecessors()
{
	BOOL result;
	GetProperty(0x2, VT_BOOL, (void*)&result);
	return result;
}

void CclsTask::SetOutgoingPredecessors(BOOL propval)
{
	SetProperty(0x3, VT_BOOL, propval);
}

BOOL CclsTask::GetOutgoingPredecessors()
{
	BOOL result;
	GetProperty(0x3, VT_BOOL, (void*)&result);
	return result;
}

void CclsTask::SetAllowStretchLeft(BOOL propval)
{
	SetProperty(0x4, VT_BOOL, propval);
}

BOOL CclsTask::GetAllowStretchLeft()
{
	BOOL result;
	GetProperty(0x4, VT_BOOL, (void*)&result);
	return result;
}

void CclsTask::SetAllowStretchRight(BOOL propval)
{
	SetProperty(0x5, VT_BOOL, propval);
}

BOOL CclsTask::GetAllowStretchRight()
{
	BOOL result;
	GetProperty(0x5, VT_BOOL, (void*)&result);
	return result;
}

void CclsTask::SetText(LPCTSTR propval)
{
	SetProperty(0x6, VT_BSTR, propval);
}

CString CclsTask::GetText()
{
	CString result;
	GetProperty(0x6, VT_BSTR, (void*)&result);
	return result;
}

void CclsTask::SetLayerIndex(LPCTSTR propval)
{
	SetProperty(0x7, VT_BSTR, propval);
}

CString CclsTask::GetLayerIndex()
{
	CString result;
	GetProperty(0x7, VT_BSTR, (void*)&result);
	return result;
}

CclsLayer CclsTask::GetLayer()
{
	LPDISPATCH pDispatch;
	GetProperty(0x8, VT_DISPATCH, (void*)&pDispatch);
	return CclsLayer(pDispatch);
}

CclsImage CclsTask::GetImage()
{
	LPDISPATCH pDispatch;
	GetProperty(0x9, VT_DISPATCH, (void*)&pDispatch);
	return CclsImage(pDispatch);
}

void CclsTask::SetRowKey(LPCTSTR propval)
{
	SetProperty(0xa, VT_BSTR, propval);
}

CString CclsTask::GetRowKey()
{
	CString result;
	GetProperty(0xa, VT_BSTR, (void*)&result);
	return result;
}

CclsRow CclsTask::GetRow()
{
	LPDISPATCH pDispatch;
	GetProperty(0xb, VT_DISPATCH, (void*)&pDispatch);
	return CclsRow(pDispatch);
}

void CclsTask::SetStyleIndex(LPCTSTR propval)
{
	SetProperty(0xc, VT_BSTR, propval);
}

CString CclsTask::GetStyleIndex()
{
	CString result;
	GetProperty(0xc, VT_BSTR, (void*)&result);
	return result;
}

CclsStyle CclsTask::GetStyle()
{
	LPDISPATCH pDispatch;
	GetProperty(0xd, VT_DISPATCH, (void*)&pDispatch);
	return CclsStyle(pDispatch);
}

void CclsTask::SetTag(LPCTSTR propval)
{
	SetProperty(0xe, VT_BSTR, propval);
}

CString CclsTask::GetTag()
{
	CString result;
	GetProperty(0xe, VT_BSTR, (void*)&result);
	return result;
}

void CclsTask::SetAllowedMovement(E_MOVEMENTTYPE propval)
{
	LONG lpropval = (LONG)propval;
	SetProperty(0xf, VT_I4, lpropval);
}

E_MOVEMENTTYPE CclsTask::GetAllowedMovement()
{
	LONG result;
	GetProperty(0xf, VT_I4, (void*)&result);
	return (E_MOVEMENTTYPE)result;
}

LONG CclsTask::GetLeftTrim()
{
	LONG result;
	GetProperty(0x10, VT_I4, (void*)&result);
	return result;
}

LONG CclsTask::GetRightTrim()
{
	LONG result;
	GetProperty(0x11, VT_I4, (void*)&result);
	return result;
}

void CclsTask::SetStartDate(DATE propval)
{
	SetProperty(0x12, VT_DATE, propval);
}

DATE CclsTask::GetStartDate()
{
	DATE result;
	GetProperty(0x12, VT_DATE, (void*)&result);
	return result;
}

LONG CclsTask::GetLeft()
{
	LONG result;
	GetProperty(0x13, VT_I4, (void*)&result);
	return result;
}

LONG CclsTask::GetRight()
{
	LONG result;
	GetProperty(0x14, VT_I4, (void*)&result);
	return result;
}

void CclsTask::SetEndDate(DATE propval)
{
	SetProperty(0x15, VT_DATE, propval);
}

DATE CclsTask::GetEndDate()
{
	DATE result;
	GetProperty(0x15, VT_DATE, (void*)&result);
	return result;
}

LONG CclsTask::GetTop()
{
	LONG result;
	GetProperty(0x16, VT_I4, (void*)&result);
	return result;
}

LONG CclsTask::GetBottom()
{
	LONG result;
	GetProperty(0x17, VT_I4, (void*)&result);
	return result;
}

void CclsTask::SetVisible(BOOL propval)
{
	SetProperty(0x18, VT_BOOL, propval);
}

BOOL CclsTask::GetVisible()
{
	BOOL result;
	GetProperty(0x18, VT_BOOL, (void*)&result);
	return result;
}

E_OBJECTTYPE CclsTask::GetType()
{
	LONG result;
	GetProperty(0x19, VT_I4, (void*)&result);
	return (E_OBJECTTYPE)result;
}

LONG CclsTask::GetIndex()
{
	LONG result;
	GetProperty(0x1a, VT_I4, (void*)&result);
	return result;
}

void CclsTask::SetImageTag(LPCTSTR propval)
{
	SetProperty(0x1b, VT_BSTR, propval);
}

CString CclsTask::GetImageTag()
{
	CString result;
	GetProperty(0x1b, VT_BSTR, (void*)&result);
	return result;
}

void CclsTask::SetAllowTextEdit(BOOL propval)
{
	SetProperty(0x1c, VT_BOOL, propval);
}

BOOL CclsTask::GetAllowTextEdit()
{
	BOOL result;
	GetProperty(0x1c, VT_BOOL, (void*)&result);
	return result;
}

BOOL CclsTask::GetWarning()
{
	BOOL result;
	GetProperty(0x1d, VT_BOOL, (void*)&result);
	return result;
}

void CclsTask::SetWarningStyleIndex(LPCTSTR propval)
{
	SetProperty(0x1e, VT_BSTR, propval);
}

CString CclsTask::GetWarningStyleIndex()
{
	CString result;
	GetProperty(0x1e, VT_BSTR, (void*)&result);
	return result;
}

CclsStyle CclsTask::GetWarningStyle()
{
	LPDISPATCH pDispatch;
	GetProperty(0x1f, VT_DISPATCH, (void*)&pDispatch);
	return CclsStyle(pDispatch);
}

void CclsTask::SetTaskType(E_TASKTYPE propval)
{
	LONG lpropval = (LONG)propval;
	SetProperty(0x20, VT_I4, lpropval);
}

E_TASKTYPE CclsTask::GetTaskType()
{
	LONG result;
	GetProperty(0x20, VT_I4, (void*)&result);
	return (E_TASKTYPE)result;
}

void CclsTask::SetDurationInterval(E_INTERVAL propval)
{
	LONG lpropval = (LONG)propval;
	SetProperty(0x21, VT_I4, lpropval);
}

E_INTERVAL CclsTask::GetDurationInterval()
{
	LONG result;
	GetProperty(0x21, VT_I4, (void*)&result);
	return (E_INTERVAL)result;
}

void CclsTask::SetDurationFactor(LONG propval)
{
	SetProperty(0x22, VT_I4, propval);
}

LONG CclsTask::GetDurationFactor()
{
	LONG result;
	GetProperty(0x22, VT_I4, (void*)&result);
	return result;
}

BOOL CclsTask::InConflict(void)
{
	BOOL oReturn;
	InvokeHelper(0x23, DISPATCH_METHOD, VT_BOOL, (void*)&oReturn, NULL);
	return oReturn;
}

CString CclsTask::GetXML(void)
{
	CString oReturn;
	InvokeHelper(0x24, DISPATCH_METHOD, VT_BSTR, (void*)&oReturn, NULL);
	return oReturn;
}

void CclsTask::SetXML(LPCTSTR sXML)
{
	static BYTE params[] = VTS_BSTR;
	InvokeHelper(0x25, DISPATCH_METHOD, VT_EMPTY, NULL, params, sXML);
}

LONG CclsTasks::GetCount()
{
	LONG result;
	GetProperty(0x1, VT_I4, (void*)&result);
	return result;
}

CclsTask CclsTasks::Item(LPCTSTR Index)
{
	CclsTask oReturn;
	static BYTE params[] = VTS_BSTR;
	InvokeHelper(0x2, DISPATCH_METHOD, VT_DISPATCH, (void*)&oReturn, params, Index);
	return oReturn;
}

CclsTask CclsTasks::Add(LPCTSTR Text, LPCTSTR RowKey, DATE StartDate, DATE EndDate, LPCTSTR Key, LPCTSTR StyleIndex, LPCTSTR LayerIndex)
{
	CclsTask oReturn;
	static BYTE params[] = VTS_BSTR VTS_BSTR VTS_DATE VTS_DATE VTS_BSTR VTS_BSTR VTS_BSTR;
	InvokeHelper(0x3, DISPATCH_METHOD, VT_DISPATCH, (void*)&oReturn, params, Text, RowKey, StartDate, EndDate, Key, StyleIndex, LayerIndex);
	return oReturn;
}

void CclsTasks::Clear(void)
{
	InvokeHelper(0x4, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
}

void CclsTasks::Remove(LPCTSTR Index)
{
	static BYTE params[] = VTS_BSTR;
	InvokeHelper(0x5, DISPATCH_METHOD, VT_EMPTY, NULL, params, Index);
}

void CclsTasks::Sort(LPCTSTR PropertyName, BOOL Descending, E_SORTTYPE SortType, LONG StartIndex, LONG EndIndex)
{
	static BYTE params[] = VTS_BSTR VTS_BOOL VTS_I4 VTS_I4 VTS_I4;
	InvokeHelper(0x6, DISPATCH_METHOD, VT_EMPTY, NULL, params, PropertyName, Descending, SortType, StartIndex, EndIndex);
}

CString CclsTasks::GetXML(void)
{
	CString oReturn;
	InvokeHelper(0x7, DISPATCH_METHOD, VT_BSTR, (void*)&oReturn, NULL);
	return oReturn;
}

void CclsTasks::SetXML(LPCTSTR sXML)
{
	static BYTE params[] = VTS_BSTR;
	InvokeHelper(0x8, DISPATCH_METHOD, VT_EMPTY, NULL, params, sXML);
}

void CclsTasks::BeginLoad(BOOL Preserve)
{
	static BYTE params[] = VTS_BOOL;
	InvokeHelper(0x9, DISPATCH_METHOD, VT_EMPTY, NULL, params, Preserve);
}

CclsTask CclsTasks::Load(LPCTSTR sRowKey, LPCTSTR sKey)
{
	CclsTask oReturn;
	static BYTE params[] = VTS_BSTR VTS_BSTR;
	InvokeHelper(0xa, DISPATCH_METHOD, VT_DISPATCH, (void*)&oReturn, params, sRowKey, sKey);
	return oReturn;
}

void CclsTasks::EndLoad(void)
{
	InvokeHelper(0xb, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
}

CclsTask CclsTasks::DAdd(LPCTSTR RowKey, DATE StartDate, E_INTERVAL DurationInterval, LONG DurationFactor, LPCTSTR Text, LPCTSTR Key, LPCTSTR StyleIndex, LPCTSTR LayerIndex)
{
	CclsTask oReturn;
	static BYTE params[] = VTS_BSTR VTS_DATE VTS_I4 VTS_I4 VTS_BSTR VTS_BSTR VTS_BSTR VTS_BSTR;
	InvokeHelper(0xc, DISPATCH_METHOD, VT_DISPATCH, (void*)&oReturn, params, RowKey, StartDate, DurationInterval, DurationFactor, Text, Key, StyleIndex, LayerIndex);
	return oReturn;
}

void CclsTimeBlock::SetKey(LPCTSTR propval)
{
	SetProperty(0x1, VT_BSTR, propval);
}

CString CclsTimeBlock::GetKey()
{
	CString result;
	GetProperty(0x1, VT_BSTR, (void*)&result);
	return result;
}

void CclsTimeBlock::SetTimeBlockType(E_TIMEBLOCKTYPE propval)
{
	LONG lpropval = (LONG)propval;
	SetProperty(0x2, VT_I4, lpropval);
}

E_TIMEBLOCKTYPE CclsTimeBlock::GetTimeBlockType()
{
	LONG result;
	GetProperty(0x2, VT_I4, (void*)&result);
	return (E_TIMEBLOCKTYPE)result;
}

void CclsTimeBlock::SetRecurringType(E_RECURRINGTYPE propval)
{
	LONG lpropval = (LONG)propval;
	SetProperty(0x3, VT_I4, lpropval);
}

E_RECURRINGTYPE CclsTimeBlock::GetRecurringType()
{
	LONG result;
	GetProperty(0x3, VT_I4, (void*)&result);
	return (E_RECURRINGTYPE)result;
}

void CclsTimeBlock::SetStyleIndex(LPCTSTR propval)
{
	SetProperty(0x4, VT_BSTR, propval);
}

CString CclsTimeBlock::GetStyleIndex()
{
	CString result;
	GetProperty(0x4, VT_BSTR, (void*)&result);
	return result;
}

CclsStyle CclsTimeBlock::GetStyle()
{
	LPDISPATCH pDispatch;
	GetProperty(0x5, VT_DISPATCH, (void*)&pDispatch);
	return CclsStyle(pDispatch);
}

void CclsTimeBlock::SetTag(LPCTSTR propval)
{
	SetProperty(0x6, VT_BSTR, propval);
}

CString CclsTimeBlock::GetTag()
{
	CString result;
	GetProperty(0x6, VT_BSTR, (void*)&result);
	return result;
}

void CclsTimeBlock::SetGenerateConflict(BOOL propval)
{
	SetProperty(0x7, VT_BOOL, propval);
}

BOOL CclsTimeBlock::GetGenerateConflict()
{
	BOOL result;
	GetProperty(0x7, VT_BOOL, (void*)&result);
	return result;
}

LONG CclsTimeBlock::GetLeftTrim()
{
	LONG result;
	GetProperty(0x8, VT_I4, (void*)&result);
	return result;
}

LONG CclsTimeBlock::GetRightTrim()
{
	LONG result;
	GetProperty(0x9, VT_I4, (void*)&result);
	return result;
}

LONG CclsTimeBlock::GetLeft()
{
	LONG result;
	GetProperty(0xa, VT_I4, (void*)&result);
	return result;
}

LONG CclsTimeBlock::GetTop()
{
	LONG result;
	GetProperty(0xb, VT_I4, (void*)&result);
	return result;
}

LONG CclsTimeBlock::GetRight()
{
	LONG result;
	GetProperty(0xc, VT_I4, (void*)&result);
	return result;
}

LONG CclsTimeBlock::GetBottom()
{
	LONG result;
	GetProperty(0xd, VT_I4, (void*)&result);
	return result;
}

void CclsTimeBlock::SetVisible(BOOL propval)
{
	SetProperty(0xe, VT_BOOL, propval);
}

BOOL CclsTimeBlock::GetVisible()
{
	BOOL result;
	GetProperty(0xe, VT_BOOL, (void*)&result);
	return result;
}

LONG CclsTimeBlock::GetIndex()
{
	LONG result;
	GetProperty(0xf, VT_I4, (void*)&result);
	return result;
}

void CclsTimeBlock::SetBaseDate(DATE propval)
{
	SetProperty(0x10, VT_DATE, propval);
}

DATE CclsTimeBlock::GetBaseDate()
{
	DATE result;
	GetProperty(0x10, VT_DATE, (void*)&result);
	return result;
}

void CclsTimeBlock::SetBaseWeekDay(E_WEEKDAY propval)
{
	LONG lpropval = (LONG)propval;
	SetProperty(0x11, VT_I4, lpropval);
}

E_WEEKDAY CclsTimeBlock::GetBaseWeekDay()
{
	LONG result;
	GetProperty(0x11, VT_I4, (void*)&result);
	return (E_WEEKDAY)result;
}

void CclsTimeBlock::SetDurationInterval(E_INTERVAL propval)
{
	LONG lpropval = (LONG)propval;
	SetProperty(0x12, VT_I4, lpropval);
}

E_INTERVAL CclsTimeBlock::GetDurationInterval()
{
	LONG result;
	GetProperty(0x12, VT_I4, (void*)&result);
	return (E_INTERVAL)result;
}

void CclsTimeBlock::SetDurationFactor(LONG propval)
{
	SetProperty(0x13, VT_I4, propval);
}

LONG CclsTimeBlock::GetDurationFactor()
{
	LONG result;
	GetProperty(0x13, VT_I4, (void*)&result);
	return result;
}

void CclsTimeBlock::SetNonWorking(BOOL propval)
{
	SetProperty(0x14, VT_BOOL, propval);
}

BOOL CclsTimeBlock::GetNonWorking()
{
	BOOL result;
	GetProperty(0x14, VT_BOOL, (void*)&result);
	return result;
}

CString CclsTimeBlock::GetXML(void)
{
	CString oReturn;
	InvokeHelper(0x15, DISPATCH_METHOD, VT_BSTR, (void*)&oReturn, NULL);
	return oReturn;
}

void CclsTimeBlock::SetXML(LPCTSTR sXML)
{
	static BYTE params[] = VTS_BSTR;
	InvokeHelper(0x16, DISPATCH_METHOD, VT_EMPTY, NULL, params, sXML);
}

LONG CclsTimeBlocks::GetCount()
{
	LONG result;
	GetProperty(0x1, VT_I4, (void*)&result);
	return result;
}

void CclsTimeBlocks::SetIntervalStart(DATE propval)
{
	SetProperty(0x8, VT_DATE, propval);
}

DATE CclsTimeBlocks::GetIntervalStart()
{
	DATE result;
	GetProperty(0x8, VT_DATE, (void*)&result);
	return result;
}

void CclsTimeBlocks::SetIntervalEnd(DATE propval)
{
	SetProperty(0x9, VT_DATE, propval);
}

DATE CclsTimeBlocks::GetIntervalEnd()
{
	DATE result;
	GetProperty(0x9, VT_DATE, (void*)&result);
	return result;
}

void CclsTimeBlocks::SetIntervalType(E_TBINTERVALTYPE propval)
{
	LONG lpropval = (LONG)propval;
	SetProperty(0xa, VT_I4, lpropval);
}

E_TBINTERVALTYPE CclsTimeBlocks::GetIntervalType()
{
	LONG result;
	GetProperty(0xa, VT_I4, (void*)&result);
	return (E_TBINTERVALTYPE)result;
}

CclsTimeBlock CclsTimeBlocks::Item(LPCTSTR Index)
{
	CclsTimeBlock oReturn;
	static BYTE params[] = VTS_BSTR;
	InvokeHelper(0x2, DISPATCH_METHOD, VT_DISPATCH, (void*)&oReturn, params, Index);
	return oReturn;
}

void CclsTimeBlocks::Clear(void)
{
	InvokeHelper(0x3, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
}

void CclsTimeBlocks::Remove(LPCTSTR Index)
{
	static BYTE params[] = VTS_BSTR;
	InvokeHelper(0x4, DISPATCH_METHOD, VT_EMPTY, NULL, params, Index);
}

CString CclsTimeBlocks::GetXML(void)
{
	CString oReturn;
	InvokeHelper(0x5, DISPATCH_METHOD, VT_BSTR, (void*)&oReturn, NULL);
	return oReturn;
}

void CclsTimeBlocks::SetXML(LPCTSTR sXML)
{
	static BYTE params[] = VTS_BSTR;
	InvokeHelper(0x6, DISPATCH_METHOD, VT_EMPTY, NULL, params, sXML);
}

CclsTimeBlock CclsTimeBlocks::Add(LPCTSTR Key)
{
	CclsTimeBlock oReturn;
	static BYTE params[] = VTS_BSTR;
	InvokeHelper(0x7, DISPATCH_METHOD, VT_DISPATCH, (void*)&oReturn, params, Key);
	return oReturn;
}

void CclsTimeBlocks::CalculateInterval(void)
{
	InvokeHelper(0xb, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
}

CString CclsTimeBlocks::CP_GetXML(void)
{
	CString oReturn;
	InvokeHelper(0xc, DISPATCH_METHOD, VT_BSTR, (void*)&oReturn, NULL);
	return oReturn;
}

void CclsTimeBlocks::CP_SetXML(LPCTSTR sXML)
{
	static BYTE params[] = VTS_BSTR;
	InvokeHelper(0xd, DISPATCH_METHOD, VT_EMPTY, NULL, params, sXML);
}

void CclsColumn::SetAllowMove(BOOL propval)
{
	SetProperty(0x1, VT_BOOL, propval);
}

BOOL CclsColumn::GetAllowMove()
{
	BOOL result;
	GetProperty(0x1, VT_BOOL, (void*)&result);
	return result;
}

void CclsColumn::SetAllowSize(BOOL propval)
{
	SetProperty(0x2, VT_BOOL, propval);
}

BOOL CclsColumn::GetAllowSize()
{
	BOOL result;
	GetProperty(0x2, VT_BOOL, (void*)&result);
	return result;
}

void CclsColumn::SetKey(LPCTSTR propval)
{
	SetProperty(0x3, VT_BSTR, propval);
}

CString CclsColumn::GetKey()
{
	CString result;
	GetProperty(0x3, VT_BSTR, (void*)&result);
	return result;
}

void CclsColumn::SetWidth(LONG propval)
{
	SetProperty(0x4, VT_I4, propval);
}

LONG CclsColumn::GetWidth()
{
	LONG result;
	GetProperty(0x4, VT_I4, (void*)&result);
	return result;
}

void CclsColumn::SetText(LPCTSTR propval)
{
	SetProperty(0x5, VT_BSTR, propval);
}

CString CclsColumn::GetText()
{
	CString result;
	GetProperty(0x5, VT_BSTR, (void*)&result);
	return result;
}

CclsImage CclsColumn::GetImage()
{
	LPDISPATCH pDispatch;
	GetProperty(0x6, VT_DISPATCH, (void*)&pDispatch);
	return CclsImage(pDispatch);
}

void CclsColumn::SetStyleIndex(LPCTSTR propval)
{
	SetProperty(0x7, VT_BSTR, propval);
}

CString CclsColumn::GetStyleIndex()
{
	CString result;
	GetProperty(0x7, VT_BSTR, (void*)&result);
	return result;
}

CclsStyle CclsColumn::GetStyle()
{
	LPDISPATCH pDispatch;
	GetProperty(0x8, VT_DISPATCH, (void*)&pDispatch);
	return CclsStyle(pDispatch);
}

void CclsColumn::SetTag(LPCTSTR propval)
{
	SetProperty(0x9, VT_BSTR, propval);
}

CString CclsColumn::GetTag()
{
	CString result;
	GetProperty(0x9, VT_BSTR, (void*)&result);
	return result;
}

LONG CclsColumn::GetLeftTrim()
{
	LONG result;
	GetProperty(0xa, VT_I4, (void*)&result);
	return result;
}

LONG CclsColumn::GetRightTrim()
{
	LONG result;
	GetProperty(0xb, VT_I4, (void*)&result);
	return result;
}

LONG CclsColumn::GetLeft()
{
	LONG result;
	GetProperty(0xc, VT_I4, (void*)&result);
	return result;
}

LONG CclsColumn::GetTop()
{
	LONG result;
	GetProperty(0xd, VT_I4, (void*)&result);
	return result;
}

LONG CclsColumn::GetRight()
{
	LONG result;
	GetProperty(0xe, VT_I4, (void*)&result);
	return result;
}

LONG CclsColumn::GetBottom()
{
	LONG result;
	GetProperty(0xf, VT_I4, (void*)&result);
	return result;
}

BOOL CclsColumn::GetVisible()
{
	BOOL result;
	GetProperty(0x10, VT_BOOL, (void*)&result);
	return result;
}

LONG CclsColumn::GetIndex()
{
	LONG result;
	GetProperty(0x13, VT_I4, (void*)&result);
	return result;
}

void CclsColumn::SetImageTag(LPCTSTR propval)
{
	SetProperty(0x14, VT_BSTR, propval);
}

CString CclsColumn::GetImageTag()
{
	CString result;
	GetProperty(0x14, VT_BSTR, (void*)&result);
	return result;
}

void CclsColumn::SetAllowTextEdit(BOOL propval)
{
	SetProperty(0x15, VT_BOOL, propval);
}

BOOL CclsColumn::GetAllowTextEdit()
{
	BOOL result;
	GetProperty(0x15, VT_BOOL, (void*)&result);
	return result;
}

CString CclsColumn::GetXML(void)
{
	CString oReturn;
	InvokeHelper(0x11, DISPATCH_METHOD, VT_BSTR, (void*)&oReturn, NULL);
	return oReturn;
}

void CclsColumn::SetXML(LPCTSTR sXML)
{
	static BYTE params[] = VTS_BSTR;
	InvokeHelper(0x12, DISPATCH_METHOD, VT_EMPTY, NULL, params, sXML);
}

LONG CclsColumns::GetCount()
{
	LONG result;
	GetProperty(0x1, VT_I4, (void*)&result);
	return result;
}

CclsColumn CclsColumns::Item(LPCTSTR Index)
{
	CclsColumn oReturn;
	static BYTE params[] = VTS_BSTR;
	InvokeHelper(0x2, DISPATCH_METHOD, VT_DISPATCH, (void*)&oReturn, params, Index);
	return oReturn;
}

CclsColumn CclsColumns::Add(LPCTSTR Text, LPCTSTR Key, LONG Width, LPCTSTR StyleIndex)
{
	CclsColumn oReturn;
	static BYTE params[] = VTS_BSTR VTS_BSTR VTS_I4 VTS_BSTR;
	InvokeHelper(0x3, DISPATCH_METHOD, VT_DISPATCH, (void*)&oReturn, params, Text, Key, Width, StyleIndex);
	return oReturn;
}

void CclsColumns::Clear(void)
{
	InvokeHelper(0x4, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
}

void CclsColumns::Remove(LPCTSTR Index)
{
	static BYTE params[] = VTS_BSTR;
	InvokeHelper(0x5, DISPATCH_METHOD, VT_EMPTY, NULL, params, Index);
}

void CclsColumns::MoveColumn(LONG OriginColumnIndex, LONG DestColumnIndex)
{
	static BYTE params[] = VTS_I4 VTS_I4;
	InvokeHelper(0x6, DISPATCH_METHOD, VT_EMPTY, NULL, params, OriginColumnIndex, DestColumnIndex);
}

CString CclsColumns::GetXML(void)
{
	CString oReturn;
	InvokeHelper(0x7, DISPATCH_METHOD, VT_BSTR, (void*)&oReturn, NULL);
	return oReturn;
}

void CclsColumns::SetXML(LPCTSTR sXML)
{
	static BYTE params[] = VTS_BSTR;
	InvokeHelper(0x8, DISPATCH_METHOD, VT_EMPTY, NULL, params, sXML);
}

CclsTextFlags CclsDrawing::GetTextFlags()
{
	LPDISPATCH pDispatch;
	GetProperty(0x8, VT_DISPATCH, (void*)&pDispatch);
	return CclsTextFlags(pDispatch);
}

CclsImage CclsDrawing::GetImage()
{
	LPDISPATCH pDispatch;
	GetProperty(0x9, VT_DISPATCH, (void*)&pDispatch);
	return CclsImage(pDispatch);
}

CclsFont CclsDrawing::GetFont()
{
	LPDISPATCH pDispatch;
	GetProperty(0xa, VT_DISPATCH, (void*)&pDispatch);
	return CclsFont(pDispatch);
}

CclsGDIGraphics CclsDrawing::GraphicsInfo(void)
{
	CclsGDIGraphics oReturn;
	InvokeHelper(0x1, DISPATCH_METHOD, VT_DISPATCH, (void*)&oReturn, NULL);
	return oReturn;
}

void CclsDrawing::DrawLine(LONG X1, LONG Y1, LONG X2, LONG Y2, unsigned long LineColor, GRE_LINEDRAWSTYLE LineStyle, LONG LineWidth)
{
	static BYTE params[] = VTS_I4 VTS_I4 VTS_I4 VTS_I4 VTS_COLOR VTS_I4 VTS_I4;
	InvokeHelper(0x2, DISPATCH_METHOD, VT_EMPTY, NULL, params, X1, Y1, X2, Y2, LineColor, LineStyle, LineWidth);
}

void CclsDrawing::DrawBorder(LONG X1, LONG Y1, LONG X2, LONG Y2, unsigned long LineColor, GRE_LINEDRAWSTYLE LineStyle, LONG LineWidth)
{
	static BYTE params[] = VTS_I4 VTS_I4 VTS_I4 VTS_I4 VTS_COLOR VTS_I4 VTS_I4;
	InvokeHelper(0x3, DISPATCH_METHOD, VT_EMPTY, NULL, params, X1, Y1, X2, Y2, LineColor, LineStyle, LineWidth);
}

void CclsDrawing::DrawRectangle(LONG X1, LONG Y1, LONG X2, LONG Y2, unsigned long LineColor, GRE_LINEDRAWSTYLE LineStyle, LONG LineWidth)
{
	static BYTE params[] = VTS_I4 VTS_I4 VTS_I4 VTS_I4 VTS_COLOR VTS_I4 VTS_I4;
	InvokeHelper(0x4, DISPATCH_METHOD, VT_EMPTY, NULL, params, X1, Y1, X2, Y2, LineColor, LineStyle, LineWidth);
}

void CclsDrawing::DrawText(LONG X1, LONG Y1, LONG X2, LONG Y2, LPCTSTR Text, unsigned long TextColor, LPDISPATCH TextFont)
{
	static BYTE params[] = VTS_I4 VTS_I4 VTS_I4 VTS_I4 VTS_BSTR VTS_COLOR VTS_DISPATCH;
	InvokeHelper(0x5, DISPATCH_METHOD, VT_EMPTY, NULL, params, X1, Y1, X2, Y2, Text, TextColor, TextFont);
}

void CclsDrawing::DrawAlignedText(LONG X1, LONG Y1, LONG X2, LONG Y2, LPCTSTR Text, GRE_HORIZONTALALIGNMENT HorizontalAlignment, GRE_VERTICALALIGNMENT VerticalAlignment, unsigned long TextColor, LPDISPATCH TextFont)
{
	static BYTE params[] = VTS_I4 VTS_I4 VTS_I4 VTS_I4 VTS_BSTR VTS_I4 VTS_I4 VTS_COLOR VTS_DISPATCH;
	InvokeHelper(0x6, DISPATCH_METHOD, VT_EMPTY, NULL, params, X1, Y1, X2, Y2, Text, HorizontalAlignment, VerticalAlignment, TextColor, TextFont);
}

void CclsDrawing::PaintImage(LPDISPATCH Image, LONG X1, LONG Y1, LONG X2, LONG Y2)
{
	static BYTE params[] = VTS_DISPATCH VTS_I4 VTS_I4 VTS_I4 VTS_I4;
	InvokeHelper(0x7, DISPATCH_METHOD, VT_EMPTY, NULL, params, Image, X1, Y1, X2, Y2);
}

void CclsDrawing::DrawObject(LONG X1, LONG Y1, LONG X2, LONG Y2, LPCTSTR StyleIndex, LPCTSTR Text, BOOL Selected, LPDISPATCH Image, GRE_DRAWINGOBJECT ObjectType)
{
	static BYTE params[] = VTS_I4 VTS_I4 VTS_I4 VTS_I4 VTS_BSTR VTS_BSTR VTS_BOOL VTS_DISPATCH VTS_I4;
	InvokeHelper(0xb, DISPATCH_METHOD, VT_EMPTY, NULL, params, X1, Y1, X2, Y2, StyleIndex, Text, Selected, Image, ObjectType);
}

void CclsPercentage::SetAllowSize(BOOL propval)
{
	SetProperty(0x1, VT_BOOL, propval);
}

BOOL CclsPercentage::GetAllowSize()
{
	BOOL result;
	GetProperty(0x1, VT_BOOL, (void*)&result);
	return result;
}

void CclsPercentage::SetKey(LPCTSTR propval)
{
	SetProperty(0x2, VT_BSTR, propval);
}

CString CclsPercentage::GetKey()
{
	CString result;
	GetProperty(0x2, VT_BSTR, (void*)&result);
	return result;
}

void CclsPercentage::SetPercent(FLOAT propval)
{
	SetProperty(0x3, VT_R4, propval);
}

FLOAT CclsPercentage::GetPercent()
{
	FLOAT result;
	GetProperty(0x3, VT_R4, (void*)&result);
	return result;
}

void CclsPercentage::SetTaskKey(LPCTSTR propval)
{
	SetProperty(0x4, VT_BSTR, propval);
}

CString CclsPercentage::GetTaskKey()
{
	CString result;
	GetProperty(0x4, VT_BSTR, (void*)&result);
	return result;
}

CclsTask CclsPercentage::GetTask()
{
	LPDISPATCH pDispatch;
	GetProperty(0x5, VT_DISPATCH, (void*)&pDispatch);
	return CclsTask(pDispatch);
}

CclsLayer CclsPercentage::GetLayer()
{
	LPDISPATCH pDispatch;
	GetProperty(0x6, VT_DISPATCH, (void*)&pDispatch);
	return CclsLayer(pDispatch);
}

void CclsPercentage::SetFormat(LPCTSTR propval)
{
	SetProperty(0x7, VT_BSTR, propval);
}

CString CclsPercentage::GetFormat()
{
	CString result;
	GetProperty(0x7, VT_BSTR, (void*)&result);
	return result;
}

void CclsPercentage::SetStyleIndex(LPCTSTR propval)
{
	SetProperty(0x8, VT_BSTR, propval);
}

CString CclsPercentage::GetStyleIndex()
{
	CString result;
	GetProperty(0x8, VT_BSTR, (void*)&result);
	return result;
}

CclsStyle CclsPercentage::GetStyle()
{
	LPDISPATCH pDispatch;
	GetProperty(0x9, VT_DISPATCH, (void*)&pDispatch);
	return CclsStyle(pDispatch);
}

void CclsPercentage::SetTag(LPCTSTR propval)
{
	SetProperty(0xa, VT_BSTR, propval);
}

CString CclsPercentage::GetTag()
{
	CString result;
	GetProperty(0xa, VT_BSTR, (void*)&result);
	return result;
}

LONG CclsPercentage::GetLeftTrim()
{
	LONG result;
	GetProperty(0xb, VT_I4, (void*)&result);
	return result;
}

LONG CclsPercentage::GetRightTrim()
{
	LONG result;
	GetProperty(0xc, VT_I4, (void*)&result);
	return result;
}

LONG CclsPercentage::GetLeft()
{
	LONG result;
	GetProperty(0xd, VT_I4, (void*)&result);
	return result;
}

LONG CclsPercentage::GetTop()
{
	LONG result;
	GetProperty(0xe, VT_I4, (void*)&result);
	return result;
}

LONG CclsPercentage::GetBottom()
{
	LONG result;
	GetProperty(0xf, VT_I4, (void*)&result);
	return result;
}

LONG CclsPercentage::GetRight()
{
	LONG result;
	GetProperty(0x10, VT_I4, (void*)&result);
	return result;
}

void CclsPercentage::SetVisible(BOOL propval)
{
	SetProperty(0x11, VT_BOOL, propval);
}

BOOL CclsPercentage::GetVisible()
{
	BOOL result;
	GetProperty(0x11, VT_BOOL, (void*)&result);
	return result;
}

LONG CclsPercentage::GetIndex()
{
	LONG result;
	GetProperty(0x14, VT_I4, (void*)&result);
	return result;
}

CString CclsPercentage::GetXML(void)
{
	CString oReturn;
	InvokeHelper(0x12, DISPATCH_METHOD, VT_BSTR, (void*)&oReturn, NULL);
	return oReturn;
}

void CclsPercentage::SetXML(LPCTSTR sXML)
{
	static BYTE params[] = VTS_BSTR;
	InvokeHelper(0x13, DISPATCH_METHOD, VT_EMPTY, NULL, params, sXML);
}

CString CclsPercentage::ToString(void)
{
	CString oReturn;
	InvokeHelper(0x15, DISPATCH_METHOD, VT_BSTR, (void*)&oReturn, NULL);
	return oReturn;
}

LONG CclsPercentages::GetCount()
{
	LONG result;
	GetProperty(0x1, VT_I4, (void*)&result);
	return result;
}

CclsPercentage CclsPercentages::Item(LPCTSTR Index)
{
	CclsPercentage oReturn;
	static BYTE params[] = VTS_BSTR;
	InvokeHelper(0x2, DISPATCH_METHOD, VT_DISPATCH, (void*)&oReturn, params, Index);
	return oReturn;
}

CclsPercentage CclsPercentages::Add(LPCTSTR TaskKey, LPCTSTR StyleIndex, FLOAT Percent, LPCTSTR Key)
{
	CclsPercentage oReturn;
	static BYTE params[] = VTS_BSTR VTS_BSTR VTS_R4 VTS_BSTR;
	InvokeHelper(0x3, DISPATCH_METHOD, VT_DISPATCH, (void*)&oReturn, params, TaskKey, StyleIndex, Percent, Key);
	return oReturn;
}

void CclsPercentages::Clear(void)
{
	InvokeHelper(0x4, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
}

void CclsPercentages::Remove(LPCTSTR Index)
{
	static BYTE params[] = VTS_BSTR;
	InvokeHelper(0x5, DISPATCH_METHOD, VT_EMPTY, NULL, params, Index);
}

CString CclsPercentages::GetXML(void)
{
	CString oReturn;
	InvokeHelper(0x6, DISPATCH_METHOD, VT_BSTR, (void*)&oReturn, NULL);
	return oReturn;
}

void CclsPercentages::SetXML(LPCTSTR sXML)
{
	static BYTE params[] = VTS_BSTR;
	InvokeHelper(0x7, DISPATCH_METHOD, VT_EMPTY, NULL, params, sXML);
}

void CclsGrid::SetVerticalLines(BOOL propval)
{
	SetProperty(0x1, VT_BOOL, propval);
}

BOOL CclsGrid::GetVerticalLines()
{
	BOOL result;
	GetProperty(0x1, VT_BOOL, (void*)&result);
	return result;
}

void CclsGrid::SetSnapToGrid(BOOL propval)
{
	SetProperty(0x2, VT_BOOL, propval);
}

BOOL CclsGrid::GetSnapToGrid()
{
	BOOL result;
	GetProperty(0x2, VT_BOOL, (void*)&result);
	return result;
}

void CclsGrid::SetSnapToGridOnSelection(BOOL propval)
{
	SetProperty(0x3, VT_BOOL, propval);
}

BOOL CclsGrid::GetSnapToGridOnSelection()
{
	BOOL result;
	GetProperty(0x3, VT_BOOL, (void*)&result);
	return result;
}

void CclsGrid::SetColor(unsigned long propval)
{
	SetProperty(0x4, VT_COLOR, propval);
}

unsigned long CclsGrid::GetColor()
{
	unsigned long result;
	GetProperty(0x4, VT_COLOR, (void*)&result);
	return result;
}

void CclsGrid::SetInterval(E_INTERVAL propval)
{
	LONG lpropval = (LONG)propval;
	SetProperty(0x5, VT_I4, lpropval);
}

E_INTERVAL CclsGrid::GetInterval()
{
	LONG result;
	GetProperty(0x5, VT_I4, (void*)&result);
	return (E_INTERVAL)result;
}

void CclsGrid::SetFactor(LONG propval)
{
	SetProperty(0x6, VT_I4, propval);
}

LONG CclsGrid::GetFactor()
{
	LONG result;
	GetProperty(0x6, VT_I4, (void*)&result);
	return result;
}

CString CclsGrid::GetXML(void)
{
	CString oReturn;
	InvokeHelper(0x7, DISPATCH_METHOD, VT_BSTR, (void*)&oReturn, NULL);
	return oReturn;
}

void CclsGrid::SetXML(LPCTSTR sXML)
{
	static BYTE params[] = VTS_BSTR;
	InvokeHelper(0x8, DISPATCH_METHOD, VT_EMPTY, NULL, params, sXML);
}

void CclsClientArea::SetDetectConflicts(BOOL propval)
{
	SetProperty(0x1, VT_BOOL, propval);
}

BOOL CclsClientArea::GetDetectConflicts()
{
	BOOL result;
	GetProperty(0x1, VT_BOOL, (void*)&result);
	return result;
}

void CclsClientArea::SetMilestoneSelectionOffset(LONG propval)
{
	SetProperty(0x2, VT_I4, propval);
}

LONG CclsClientArea::GetMilestoneSelectionOffset()
{
	LONG result;
	GetProperty(0x2, VT_I4, (void*)&result);
	return result;
}

void CclsClientArea::SetFirstVisibleRow(LONG propval)
{
	SetProperty(0x3, VT_I4, propval);
}

LONG CclsClientArea::GetFirstVisibleRow()
{
	LONG result;
	GetProperty(0x3, VT_I4, (void*)&result);
	return result;
}

LONG CclsClientArea::GetLastVisibleRow()
{
	LONG result;
	GetProperty(0x4, VT_I4, (void*)&result);
	return result;
}

void CclsClientArea::SetToolTipFormat(LPCTSTR propval)
{
	SetProperty(0x5, VT_BSTR, propval);
}

CString CclsClientArea::GetToolTipFormat()
{
	CString result;
	GetProperty(0x5, VT_BSTR, (void*)&result);
	return result;
}

void CclsClientArea::SetToolTipsVisible(BOOL propval)
{
	SetProperty(0x6, VT_BOOL, propval);
}

BOOL CclsClientArea::GetToolTipsVisible()
{
	BOOL result;
	GetProperty(0x6, VT_BOOL, (void*)&result);
	return result;
}

LONG CclsClientArea::GetTop()
{
	LONG result;
	GetProperty(0x7, VT_I4, (void*)&result);
	return result;
}

LONG CclsClientArea::GetBottom()
{
	LONG result;
	GetProperty(0x8, VT_I4, (void*)&result);
	return result;
}

LONG CclsClientArea::GetLeft()
{
	LONG result;
	GetProperty(0x9, VT_I4, (void*)&result);
	return result;
}

LONG CclsClientArea::GetRight()
{
	LONG result;
	GetProperty(0xa, VT_I4, (void*)&result);
	return result;
}

LONG CclsClientArea::GetWidth()
{
	LONG result;
	GetProperty(0xb, VT_I4, (void*)&result);
	return result;
}

LONG CclsClientArea::GetHeight()
{
	LONG result;
	GetProperty(0xc, VT_I4, (void*)&result);
	return result;
}

CclsGrid CclsClientArea::GetGrid()
{
	LPDISPATCH pDispatch;
	GetProperty(0xd, VT_DISPATCH, (void*)&pDispatch);
	return CclsGrid(pDispatch);
}

void CclsClientArea::SetPredecessorSelectionOffset(LONG propval)
{
	SetProperty(0x10, VT_I4, propval);
}

LONG CclsClientArea::GetPredecessorSelectionOffset()
{
	LONG result;
	GetProperty(0x10, VT_I4, (void*)&result);
	return result;
}

void CclsClientArea::SetTaskBorderSelectionOffset(LONG propval)
{
	SetProperty(0x11, VT_I4, propval);
}

LONG CclsClientArea::GetTaskBorderSelectionOffset()
{
	LONG result;
	GetProperty(0x11, VT_I4, (void*)&result);
	return result;
}

CString CclsClientArea::GetXML(void)
{
	CString oReturn;
	InvokeHelper(0xe, DISPATCH_METHOD, VT_BSTR, (void*)&oReturn, NULL);
	return oReturn;
}

void CclsClientArea::SetXML(LPCTSTR sXML)
{
	static BYTE params[] = VTS_BSTR;
	InvokeHelper(0xf, DISPATCH_METHOD, VT_EMPTY, NULL, params, sXML);
}

void CclsTimeLineScrollBar::SetInterval(E_INTERVAL propval)
{
	LONG lpropval = (LONG)propval;
	SetProperty(0x1, VT_I4, lpropval);
}

E_INTERVAL CclsTimeLineScrollBar::GetInterval()
{
	LONG result;
	GetProperty(0x1, VT_I4, (void*)&result);
	return (E_INTERVAL)result;
}

void CclsTimeLineScrollBar::SetValue(LONG propval)
{
	SetProperty(0x2, VT_I4, propval);
}

LONG CclsTimeLineScrollBar::GetValue()
{
	LONG result;
	GetProperty(0x2, VT_I4, (void*)&result);
	return result;
}

void CclsTimeLineScrollBar::SetEnabled(BOOL propval)
{
	SetProperty(0x3, VT_BOOL, propval);
}

BOOL CclsTimeLineScrollBar::GetEnabled()
{
	BOOL result;
	GetProperty(0x3, VT_BOOL, (void*)&result);
	return result;
}

void CclsTimeLineScrollBar::SetVisible(BOOL propval)
{
	SetProperty(0x4, VT_BOOL, propval);
}

BOOL CclsTimeLineScrollBar::GetVisible()
{
	BOOL result;
	GetProperty(0x4, VT_BOOL, (void*)&result);
	return result;
}

void CclsTimeLineScrollBar::SetLargeChange(LONG propval)
{
	SetProperty(0x5, VT_I4, propval);
}

LONG CclsTimeLineScrollBar::GetLargeChange()
{
	LONG result;
	GetProperty(0x5, VT_I4, (void*)&result);
	return result;
}

void CclsTimeLineScrollBar::SetMax(LONG propval)
{
	SetProperty(0x6, VT_I4, propval);
}

LONG CclsTimeLineScrollBar::GetMax()
{
	LONG result;
	GetProperty(0x6, VT_I4, (void*)&result);
	return result;
}

void CclsTimeLineScrollBar::SetSmallChange(LONG propval)
{
	SetProperty(0x7, VT_I4, propval);
}

LONG CclsTimeLineScrollBar::GetSmallChange()
{
	LONG result;
	GetProperty(0x7, VT_I4, (void*)&result);
	return result;
}

void CclsTimeLineScrollBar::SetStartDate(DATE propval)
{
	SetProperty(0x8, VT_DATE, propval);
}

DATE CclsTimeLineScrollBar::GetStartDate()
{
	DATE result;
	GetProperty(0x8, VT_DATE, (void*)&result);
	return result;
}

CclsHScrollBarTemplate CclsTimeLineScrollBar::GetScrollBar()
{
	LPDISPATCH pDispatch;
	GetProperty(0x9, VT_DISPATCH, (void*)&pDispatch);
	return CclsHScrollBarTemplate(pDispatch);
}

void CclsTimeLineScrollBar::SetFactor(LONG propval)
{
	SetProperty(0xc, VT_I4, propval);
}

LONG CclsTimeLineScrollBar::GetFactor()
{
	LONG result;
	GetProperty(0xc, VT_I4, (void*)&result);
	return result;
}

CString CclsTimeLineScrollBar::GetXML(void)
{
	CString oReturn;
	InvokeHelper(0xa, DISPATCH_METHOD, VT_BSTR, (void*)&oReturn, NULL);
	return oReturn;
}

void CclsTimeLineScrollBar::SetXML(LPCTSTR sXML)
{
	static BYTE params[] = VTS_BSTR;
	InvokeHelper(0xb, DISPATCH_METHOD, VT_EMPTY, NULL, params, sXML);
}

void CclsTier::SetVisible(BOOL propval)
{
	SetProperty(0x1, VT_BOOL, propval);
}

BOOL CclsTier::GetVisible()
{
	BOOL result;
	GetProperty(0x1, VT_BOOL, (void*)&result);
	return result;
}

void CclsTier::SetTag(LPCTSTR propval)
{
	SetProperty(0x2, VT_BSTR, propval);
}

CString CclsTier::GetTag()
{
	CString result;
	GetProperty(0x2, VT_BSTR, (void*)&result);
	return result;
}

void CclsTier::SetInterval(E_INTERVAL propval)
{
	LONG lpropval = (LONG)propval;
	SetProperty(0x3, VT_I4, lpropval);
}

E_INTERVAL CclsTier::GetInterval()
{
	LONG result;
	GetProperty(0x3, VT_I4, (void*)&result);
	return (E_INTERVAL)result;
}

void CclsTier::SetTierType(E_TIERTYPE propval)
{
	LONG lpropval = (LONG)propval;
	SetProperty(0x4, VT_I4, lpropval);
}

E_TIERTYPE CclsTier::GetTierType()
{
	LONG result;
	GetProperty(0x4, VT_I4, (void*)&result);
	return (E_TIERTYPE)result;
}

void CclsTier::SetHeight(LONG propval)
{
	SetProperty(0x5, VT_I4, propval);
}

LONG CclsTier::GetHeight()
{
	LONG result;
	GetProperty(0x5, VT_I4, (void*)&result);
	return result;
}

void CclsTier::SetFactor(LONG propval)
{
	SetProperty(0x8, VT_I4, propval);
}

LONG CclsTier::GetFactor()
{
	LONG result;
	GetProperty(0x8, VT_I4, (void*)&result);
	return result;
}

void CclsTier::SetStyleIndex(LPCTSTR propval)
{
	SetProperty(0x9, VT_BSTR, propval);
}

CString CclsTier::GetStyleIndex()
{
	CString result;
	GetProperty(0x9, VT_BSTR, (void*)&result);
	return result;
}

CclsStyle CclsTier::GetStyle()
{
	LPDISPATCH pDispatch;
	GetProperty(0xa, VT_DISPATCH, (void*)&pDispatch);
	return CclsStyle(pDispatch);
}

void CclsTier::SetBackgroundMode(E_TIERBACKGROUNDMODE propval)
{
	LONG lpropval = (LONG)propval;
	SetProperty(0xb, VT_I4, lpropval);
}

E_TIERBACKGROUNDMODE CclsTier::GetBackgroundMode()
{
	LONG result;
	GetProperty(0xb, VT_I4, (void*)&result);
	return (E_TIERBACKGROUNDMODE)result;
}

CString CclsTier::GetXML(void)
{
	CString oReturn;
	InvokeHelper(0x6, DISPATCH_METHOD, VT_BSTR, (void*)&oReturn, NULL);
	return oReturn;
}

void CclsTier::SetXML(LPCTSTR sXML)
{
	static BYTE params[] = VTS_BSTR;
	InvokeHelper(0x7, DISPATCH_METHOD, VT_EMPTY, NULL, params, sXML);
}

void CclsTierFormat::SetMinuteIntervalFormat(LPCTSTR propval)
{
	SetProperty(0x1, VT_BSTR, propval);
}

CString CclsTierFormat::GetMinuteIntervalFormat()
{
	CString result;
	GetProperty(0x1, VT_BSTR, (void*)&result);
	return result;
}

void CclsTierFormat::SetHourIntervalFormat(LPCTSTR propval)
{
	SetProperty(0x2, VT_BSTR, propval);
}

CString CclsTierFormat::GetHourIntervalFormat()
{
	CString result;
	GetProperty(0x2, VT_BSTR, (void*)&result);
	return result;
}

void CclsTierFormat::SetDayIntervalFormat(LPCTSTR propval)
{
	SetProperty(0x3, VT_BSTR, propval);
}

CString CclsTierFormat::GetDayIntervalFormat()
{
	CString result;
	GetProperty(0x3, VT_BSTR, (void*)&result);
	return result;
}

void CclsTierFormat::SetDayOfWeekIntervalFormat(LPCTSTR propval)
{
	SetProperty(0x4, VT_BSTR, propval);
}

CString CclsTierFormat::GetDayOfWeekIntervalFormat()
{
	CString result;
	GetProperty(0x4, VT_BSTR, (void*)&result);
	return result;
}

void CclsTierFormat::SetDayOfYearIntervalFormat(LPCTSTR propval)
{
	SetProperty(0x5, VT_BSTR, propval);
}

CString CclsTierFormat::GetDayOfYearIntervalFormat()
{
	CString result;
	GetProperty(0x5, VT_BSTR, (void*)&result);
	return result;
}

void CclsTierFormat::SetWeekIntervalFormat(LPCTSTR propval)
{
	SetProperty(0x6, VT_BSTR, propval);
}

CString CclsTierFormat::GetWeekIntervalFormat()
{
	CString result;
	GetProperty(0x6, VT_BSTR, (void*)&result);
	return result;
}

void CclsTierFormat::SetMonthIntervalFormat(LPCTSTR propval)
{
	SetProperty(0x7, VT_BSTR, propval);
}

CString CclsTierFormat::GetMonthIntervalFormat()
{
	CString result;
	GetProperty(0x7, VT_BSTR, (void*)&result);
	return result;
}

void CclsTierFormat::SetQuarterIntervalFormat(LPCTSTR propval)
{
	SetProperty(0x8, VT_BSTR, propval);
}

CString CclsTierFormat::GetQuarterIntervalFormat()
{
	CString result;
	GetProperty(0x8, VT_BSTR, (void*)&result);
	return result;
}

void CclsTierFormat::SetYearIntervalFormat(LPCTSTR propval)
{
	SetProperty(0x9, VT_BSTR, propval);
}

CString CclsTierFormat::GetYearIntervalFormat()
{
	CString result;
	GetProperty(0x9, VT_BSTR, (void*)&result);
	return result;
}

void CclsTierFormat::SetSecondIntervalFormat(LPCTSTR propval)
{
	SetProperty(0xc, VT_BSTR, propval);
}

CString CclsTierFormat::GetSecondIntervalFormat()
{
	CString result;
	GetProperty(0xc, VT_BSTR, (void*)&result);
	return result;
}

void CclsTierFormat::SetMillisecondIntervalFormat(LPCTSTR propval)
{
	SetProperty(0xd, VT_BSTR, propval);
}

CString CclsTierFormat::GetMillisecondIntervalFormat()
{
	CString result;
	GetProperty(0xd, VT_BSTR, (void*)&result);
	return result;
}

void CclsTierFormat::SetMicrosecondIntervalFormat(LPCTSTR propval)
{
	SetProperty(0xe, VT_BSTR, propval);
}

CString CclsTierFormat::GetMicrosecondIntervalFormat()
{
	CString result;
	GetProperty(0xe, VT_BSTR, (void*)&result);
	return result;
}

CString CclsTierFormat::GetXML(void)
{
	CString oReturn;
	InvokeHelper(0xa, DISPATCH_METHOD, VT_BSTR, (void*)&oReturn, NULL);
	return oReturn;
}

void CclsTierFormat::SetXML(LPCTSTR sXML)
{
	static BYTE params[] = VTS_BSTR;
	InvokeHelper(0xb, DISPATCH_METHOD, VT_EMPTY, NULL, params, sXML);
}

void CclsTierColor::SetKey(LPCTSTR propval)
{
	SetProperty(0x1, VT_BSTR, propval);
}

CString CclsTierColor::GetKey()
{
	CString result;
	GetProperty(0x1, VT_BSTR, (void*)&result);
	return result;
}

void CclsTierColor::SetForeColor(unsigned long propval)
{
	SetProperty(0x2, VT_COLOR, propval);
}

unsigned long CclsTierColor::GetForeColor()
{
	unsigned long result;
	GetProperty(0x2, VT_COLOR, (void*)&result);
	return result;
}

void CclsTierColor::SetBackColor(unsigned long propval)
{
	SetProperty(0x3, VT_COLOR, propval);
}

unsigned long CclsTierColor::GetBackColor()
{
	unsigned long result;
	GetProperty(0x3, VT_COLOR, (void*)&result);
	return result;
}

LONG CclsTierColor::GetIndex()
{
	LONG result;
	GetProperty(0x6, VT_I4, (void*)&result);
	return result;
}

void CclsTierColor::SetStartGradientColor(unsigned long propval)
{
	SetProperty(0x7, VT_COLOR, propval);
}

unsigned long CclsTierColor::GetStartGradientColor()
{
	unsigned long result;
	GetProperty(0x7, VT_COLOR, (void*)&result);
	return result;
}

void CclsTierColor::SetEndGradientColor(unsigned long propval)
{
	SetProperty(0x8, VT_COLOR, propval);
}

unsigned long CclsTierColor::GetEndGradientColor()
{
	unsigned long result;
	GetProperty(0x8, VT_COLOR, (void*)&result);
	return result;
}

void CclsTierColor::SetHatchBackColor(unsigned long propval)
{
	SetProperty(0x9, VT_COLOR, propval);
}

unsigned long CclsTierColor::GetHatchBackColor()
{
	unsigned long result;
	GetProperty(0x9, VT_COLOR, (void*)&result);
	return result;
}

void CclsTierColor::SetHatchForeColor(unsigned long propval)
{
	SetProperty(0xa, VT_COLOR, propval);
}

unsigned long CclsTierColor::GetHatchForeColor()
{
	unsigned long result;
	GetProperty(0xa, VT_COLOR, (void*)&result);
	return result;
}

CString CclsTierColor::GetXML(void)
{
	CString oReturn;
	InvokeHelper(0x4, DISPATCH_METHOD, VT_BSTR, (void*)&oReturn, NULL);
	return oReturn;
}

void CclsTierColor::SetXML(LPCTSTR sXML)
{
	static BYTE params[] = VTS_BSTR;
	InvokeHelper(0x5, DISPATCH_METHOD, VT_EMPTY, NULL, params, sXML);
}

LONG CclsTierColors::GetCount()
{
	LONG result;
	GetProperty(0x1, VT_I4, (void*)&result);
	return result;
}

CclsTierColor CclsTierColors::Item(LPCTSTR Index)
{
	CclsTierColor oReturn;
	static BYTE params[] = VTS_BSTR;
	InvokeHelper(0x2, DISPATCH_METHOD, VT_DISPATCH, (void*)&oReturn, params, Index);
	return oReturn;
}

CString CclsTierColors::GetXML(void)
{
	CString oReturn;
	InvokeHelper(0x3, DISPATCH_METHOD, VT_BSTR, (void*)&oReturn, NULL);
	return oReturn;
}

void CclsTierColors::SetXML(LPCTSTR sXML)
{
	static BYTE params[] = VTS_BSTR;
	InvokeHelper(0x4, DISPATCH_METHOD, VT_EMPTY, NULL, params, sXML);
}

CclsTierColors CclsTierAppearance::GetMinuteColors()
{
	LPDISPATCH pDispatch;
	GetProperty(0x1, VT_DISPATCH, (void*)&pDispatch);
	return CclsTierColors(pDispatch);
}

CclsTierColors CclsTierAppearance::GetHourColors()
{
	LPDISPATCH pDispatch;
	GetProperty(0x2, VT_DISPATCH, (void*)&pDispatch);
	return CclsTierColors(pDispatch);
}

CclsTierColors CclsTierAppearance::GetDayColors()
{
	LPDISPATCH pDispatch;
	GetProperty(0x3, VT_DISPATCH, (void*)&pDispatch);
	return CclsTierColors(pDispatch);
}

CclsTierColors CclsTierAppearance::GetDayOfWeekColors()
{
	LPDISPATCH pDispatch;
	GetProperty(0x4, VT_DISPATCH, (void*)&pDispatch);
	return CclsTierColors(pDispatch);
}

CclsTierColors CclsTierAppearance::GetDayOfYearColors()
{
	LPDISPATCH pDispatch;
	GetProperty(0x5, VT_DISPATCH, (void*)&pDispatch);
	return CclsTierColors(pDispatch);
}

CclsTierColors CclsTierAppearance::GetWeekColors()
{
	LPDISPATCH pDispatch;
	GetProperty(0x6, VT_DISPATCH, (void*)&pDispatch);
	return CclsTierColors(pDispatch);
}

CclsTierColors CclsTierAppearance::GetMonthColors()
{
	LPDISPATCH pDispatch;
	GetProperty(0x7, VT_DISPATCH, (void*)&pDispatch);
	return CclsTierColors(pDispatch);
}

CclsTierColors CclsTierAppearance::GetQuarterColors()
{
	LPDISPATCH pDispatch;
	GetProperty(0x8, VT_DISPATCH, (void*)&pDispatch);
	return CclsTierColors(pDispatch);
}

CclsTierColors CclsTierAppearance::GetYearColors()
{
	LPDISPATCH pDispatch;
	GetProperty(0x9, VT_DISPATCH, (void*)&pDispatch);
	return CclsTierColors(pDispatch);
}

CclsTierColors CclsTierAppearance::GetSecondColors()
{
	LPDISPATCH pDispatch;
	GetProperty(0xc, VT_DISPATCH, (void*)&pDispatch);
	return CclsTierColors(pDispatch);
}

CclsTierColors CclsTierAppearance::GetMillisecondColors()
{
	LPDISPATCH pDispatch;
	GetProperty(0xd, VT_DISPATCH, (void*)&pDispatch);
	return CclsTierColors(pDispatch);
}

CclsTierColors CclsTierAppearance::GetMicrosecondColors()
{
	LPDISPATCH pDispatch;
	GetProperty(0xe, VT_DISPATCH, (void*)&pDispatch);
	return CclsTierColors(pDispatch);
}

CString CclsTierAppearance::GetXML(void)
{
	CString oReturn;
	InvokeHelper(0xa, DISPATCH_METHOD, VT_BSTR, (void*)&oReturn, NULL);
	return oReturn;
}

void CclsTierAppearance::SetXML(LPCTSTR sXML)
{
	static BYTE params[] = VTS_BSTR;
	InvokeHelper(0xb, DISPATCH_METHOD, VT_EMPTY, NULL, params, sXML);
}

CclsTier CclsTierArea::GetUpperTier()
{
	LPDISPATCH pDispatch;
	GetProperty(0x1, VT_DISPATCH, (void*)&pDispatch);
	return CclsTier(pDispatch);
}

CclsTier CclsTierArea::GetMiddleTier()
{
	LPDISPATCH pDispatch;
	GetProperty(0x2, VT_DISPATCH, (void*)&pDispatch);
	return CclsTier(pDispatch);
}

CclsTier CclsTierArea::GetLowerTier()
{
	LPDISPATCH pDispatch;
	GetProperty(0x3, VT_DISPATCH, (void*)&pDispatch);
	return CclsTier(pDispatch);
}

CclsTierFormat CclsTierArea::GetTierFormat()
{
	LPDISPATCH pDispatch;
	GetProperty(0x4, VT_DISPATCH, (void*)&pDispatch);
	return CclsTierFormat(pDispatch);
}

CclsTierAppearance CclsTierArea::GetTierAppearance()
{
	LPDISPATCH pDispatch;
	GetProperty(0x5, VT_DISPATCH, (void*)&pDispatch);
	return CclsTierAppearance(pDispatch);
}

CString CclsTierArea::GetXML(void)
{
	CString oReturn;
	InvokeHelper(0x6, DISPATCH_METHOD, VT_BSTR, (void*)&oReturn, NULL);
	return oReturn;
}

void CclsTierArea::SetXML(LPCTSTR sXML)
{
	static BYTE params[] = VTS_BSTR;
	InvokeHelper(0x7, DISPATCH_METHOD, VT_EMPTY, NULL, params, sXML);
}

void CclsTickMark::SetKey(LPCTSTR propval)
{
	SetProperty(0x1, VT_BSTR, propval);
}

CString CclsTickMark::GetKey()
{
	CString result;
	GetProperty(0x1, VT_BSTR, (void*)&result);
	return result;
}

void CclsTickMark::SetDisplayText(BOOL propval)
{
	SetProperty(0x2, VT_BOOL, propval);
}

BOOL CclsTickMark::GetDisplayText()
{
	BOOL result;
	GetProperty(0x2, VT_BOOL, (void*)&result);
	return result;
}

void CclsTickMark::SetTextFormat(LPCTSTR propval)
{
	SetProperty(0x3, VT_BSTR, propval);
}

CString CclsTickMark::GetTextFormat()
{
	CString result;
	GetProperty(0x3, VT_BSTR, (void*)&result);
	return result;
}

void CclsTickMark::SetTag(LPCTSTR propval)
{
	SetProperty(0x4, VT_BSTR, propval);
}

CString CclsTickMark::GetTag()
{
	CString result;
	GetProperty(0x4, VT_BSTR, (void*)&result);
	return result;
}

void CclsTickMark::SetInterval(E_INTERVAL propval)
{
	LONG lpropval = (LONG)propval;
	SetProperty(0x5, VT_I4, lpropval);
}

E_INTERVAL CclsTickMark::GetInterval()
{
	LONG result;
	GetProperty(0x5, VT_I4, (void*)&result);
	return (E_INTERVAL)result;
}

void CclsTickMark::SetTickMarkType(E_TICKMARKTYPES propval)
{
	LONG lpropval = (LONG)propval;
	SetProperty(0x6, VT_I4, lpropval);
}

E_TICKMARKTYPES CclsTickMark::GetTickMarkType()
{
	LONG result;
	GetProperty(0x6, VT_I4, (void*)&result);
	return (E_TICKMARKTYPES)result;
}

LONG CclsTickMark::GetIndex()
{
	LONG result;
	GetProperty(0x9, VT_I4, (void*)&result);
	return result;
}

void CclsTickMark::SetFactor(LONG propval)
{
	SetProperty(0xa, VT_I4, propval);
}

LONG CclsTickMark::GetFactor()
{
	LONG result;
	GetProperty(0xa, VT_I4, (void*)&result);
	return result;
}

CString CclsTickMark::GetXML(void)
{
	CString oReturn;
	InvokeHelper(0x7, DISPATCH_METHOD, VT_BSTR, (void*)&oReturn, NULL);
	return oReturn;
}

void CclsTickMark::SetXML(LPCTSTR sXML)
{
	static BYTE params[] = VTS_BSTR;
	InvokeHelper(0x8, DISPATCH_METHOD, VT_EMPTY, NULL, params, sXML);
}

LONG CclsTickMarks::GetCount()
{
	LONG result;
	GetProperty(0x1, VT_I4, (void*)&result);
	return result;
}

CclsTickMark CclsTickMarks::Item(LPCTSTR Index)
{
	CclsTickMark oReturn;
	static BYTE params[] = VTS_BSTR;
	InvokeHelper(0x2, DISPATCH_METHOD, VT_DISPATCH, (void*)&oReturn, params, Index);
	return oReturn;
}

void CclsTickMarks::Clear(void)
{
	InvokeHelper(0x3, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
}

void CclsTickMarks::Remove(LPCTSTR Index)
{
	static BYTE params[] = VTS_BSTR;
	InvokeHelper(0x4, DISPATCH_METHOD, VT_EMPTY, NULL, params, Index);
}

CString CclsTickMarks::GetXML(void)
{
	CString oReturn;
	InvokeHelper(0x5, DISPATCH_METHOD, VT_BSTR, (void*)&oReturn, NULL);
	return oReturn;
}

void CclsTickMarks::SetXML(LPCTSTR sXML)
{
	static BYTE params[] = VTS_BSTR;
	InvokeHelper(0x6, DISPATCH_METHOD, VT_EMPTY, NULL, params, sXML);
}

CclsTickMark CclsTickMarks::Add(E_INTERVAL Interval, LONG Factor, E_TICKMARKTYPES TickMarkType, BOOL DisplayText, LPCTSTR TextFormat, LPCTSTR Key)
{
	CclsTickMark oReturn;
	static BYTE params[] = VTS_I4 VTS_I4 VTS_I4 VTS_BOOL VTS_BSTR VTS_BSTR;
	InvokeHelper(0x7, DISPATCH_METHOD, VT_DISPATCH, (void*)&oReturn, params, Interval, Factor, TickMarkType, DisplayText, TextFormat, Key);
	return oReturn;
}

void CclsTickMarkArea::SetHeight(LONG propval)
{
	SetProperty(0x1, VT_I4, propval);
}

LONG CclsTickMarkArea::GetHeight()
{
	LONG result;
	GetProperty(0x1, VT_I4, (void*)&result);
	return result;
}

void CclsTickMarkArea::SetBigTickMarkHeight(LONG propval)
{
	SetProperty(0x2, VT_I4, propval);
}

LONG CclsTickMarkArea::GetBigTickMarkHeight()
{
	LONG result;
	GetProperty(0x2, VT_I4, (void*)&result);
	return result;
}

void CclsTickMarkArea::SetMediumTickMarkHeight(LONG propval)
{
	SetProperty(0x3, VT_I4, propval);
}

LONG CclsTickMarkArea::GetMediumTickMarkHeight()
{
	LONG result;
	GetProperty(0x3, VT_I4, (void*)&result);
	return result;
}

void CclsTickMarkArea::SetSmallTickMarkHeight(LONG propval)
{
	SetProperty(0x4, VT_I4, propval);
}

LONG CclsTickMarkArea::GetSmallTickMarkHeight()
{
	LONG result;
	GetProperty(0x4, VT_I4, (void*)&result);
	return result;
}

void CclsTickMarkArea::SetVisible(BOOL propval)
{
	SetProperty(0x5, VT_BOOL, propval);
}

BOOL CclsTickMarkArea::GetVisible()
{
	BOOL result;
	GetProperty(0x5, VT_BOOL, (void*)&result);
	return result;
}

void CclsTickMarkArea::SetStyleIndex(LPCTSTR propval)
{
	SetProperty(0x6, VT_BSTR, propval);
}

CString CclsTickMarkArea::GetStyleIndex()
{
	CString result;
	GetProperty(0x6, VT_BSTR, (void*)&result);
	return result;
}

void CclsTickMarkArea::SetTextOffset(LONG propval)
{
	SetProperty(0x7, VT_I4, propval);
}

LONG CclsTickMarkArea::GetTextOffset()
{
	LONG result;
	GetProperty(0x7, VT_I4, (void*)&result);
	return result;
}

CclsTickMarks CclsTickMarkArea::GetTickMarks()
{
	LPDISPATCH pDispatch;
	GetProperty(0x8, VT_DISPATCH, (void*)&pDispatch);
	return CclsTickMarks(pDispatch);
}

CclsStyle CclsTickMarkArea::GetStyle()
{
	LPDISPATCH pDispatch;
	GetProperty(0xb, VT_DISPATCH, (void*)&pDispatch);
	return CclsStyle(pDispatch);
}

CString CclsTickMarkArea::GetXML(void)
{
	CString oReturn;
	InvokeHelper(0x9, DISPATCH_METHOD, VT_BSTR, (void*)&oReturn, NULL);
	return oReturn;
}

void CclsTickMarkArea::SetXML(LPCTSTR sXML)
{
	static BYTE params[] = VTS_BSTR;
	InvokeHelper(0xa, DISPATCH_METHOD, VT_EMPTY, NULL, params, sXML);
}

void CclsProgressLine::SetPosition(DATE propval)
{
	SetProperty(0x1, VT_DATE, propval);
}

DATE CclsProgressLine::GetPosition()
{
	DATE result;
	GetProperty(0x1, VT_DATE, (void*)&result);
	return result;
}

void CclsProgressLine::SetForeColor(unsigned long propval)
{
	SetProperty(0x2, VT_COLOR, propval);
}

unsigned long CclsProgressLine::GetForeColor()
{
	unsigned long result;
	GetProperty(0x2, VT_COLOR, (void*)&result);
	return result;
}

void CclsProgressLine::SetLength(E_PROGRESSLINELENGTH propval)
{
	LONG lpropval = (LONG)propval;
	SetProperty(0x3, VT_I4, lpropval);
}

E_PROGRESSLINELENGTH CclsProgressLine::GetLength()
{
	LONG result;
	GetProperty(0x3, VT_I4, (void*)&result);
	return (E_PROGRESSLINELENGTH)result;
}

void CclsProgressLine::SetLineType(E_PROGRESSLINETYPE propval)
{
	LONG lpropval = (LONG)propval;
	SetProperty(0x4, VT_I4, lpropval);
}

E_PROGRESSLINETYPE CclsProgressLine::GetLineType()
{
	LONG result;
	GetProperty(0x4, VT_I4, (void*)&result);
	return (E_PROGRESSLINETYPE)result;
}

void CclsProgressLine::SetWidth(LONG propval)
{
	SetProperty(0x7, VT_I4, propval);
}

LONG CclsProgressLine::GetWidth()
{
	LONG result;
	GetProperty(0x7, VT_I4, (void*)&result);
	return result;
}

void CclsProgressLine::SetLineStyle(E_PROGRESSLINESTYLE propval)
{
	LONG lpropval = (LONG)propval;
	SetProperty(0x8, VT_I4, lpropval);
}

E_PROGRESSLINESTYLE CclsProgressLine::GetLineStyle()
{
	LONG result;
	GetProperty(0x8, VT_I4, (void*)&result);
	return (E_PROGRESSLINESTYLE)result;
}

void CclsProgressLine::SetStyleIndex(LPCTSTR propval)
{
	SetProperty(0xb, VT_BSTR, propval);
}

CString CclsProgressLine::GetStyleIndex()
{
	CString result;
	GetProperty(0xb, VT_BSTR, (void*)&result);
	return result;
}

CclsStyle CclsProgressLine::GetStyle()
{
	LPDISPATCH pDispatch;
	GetProperty(0xc, VT_DISPATCH, (void*)&pDispatch);
	return CclsStyle(pDispatch);
}

CString CclsProgressLine::GetXML(void)
{
	CString oReturn;
	InvokeHelper(0x5, DISPATCH_METHOD, VT_BSTR, (void*)&oReturn, NULL);
	return oReturn;
}

void CclsProgressLine::SetXML(LPCTSTR sXML)
{
	static BYTE params[] = VTS_BSTR;
	InvokeHelper(0x6, DISPATCH_METHOD, VT_EMPTY, NULL, params, sXML);
}

void CclsProgressLine::SetColor(LONG Index, unsigned long oColor)
{
	static BYTE params[] = VTS_I4 VTS_COLOR;
	InvokeHelper(0x9, DISPATCH_METHOD, VT_EMPTY, NULL, params, Index, oColor);
}

unsigned long CclsProgressLine::GetColor(LONG Index)
{
	unsigned long oReturn;
	static BYTE params[] = VTS_I4;
	InvokeHelper(0xa, DISPATCH_METHOD, VT_COLOR, (void*)&oReturn, params, Index);
	return oReturn;
}

void CclsTimeLine::SetStyleIndex(LPCTSTR propval)
{
	SetProperty(0x1, VT_BSTR, propval);
}

CString CclsTimeLine::GetStyleIndex()
{
	CString result;
	GetProperty(0x1, VT_BSTR, (void*)&result);
	return result;
}

CclsStyle CclsTimeLine::GetStyle()
{
	LPDISPATCH pDispatch;
	GetProperty(0x2, VT_DISPATCH, (void*)&pDispatch);
	return CclsStyle(pDispatch);
}

void CclsTimeLine::SetForeColor(unsigned long propval)
{
	SetProperty(0x3, VT_COLOR, propval);
}

unsigned long CclsTimeLine::GetForeColor()
{
	unsigned long result;
	GetProperty(0x3, VT_COLOR, (void*)&result);
	return result;
}

DATE CclsTimeLine::GetEndDate()
{
	DATE result;
	GetProperty(0x4, VT_DATE, (void*)&result);
	return result;
}

DATE CclsTimeLine::GetStartDate()
{
	DATE result;
	GetProperty(0x5, VT_DATE, (void*)&result);
	return result;
}

LONG CclsTimeLine::GetHeight()
{
	LONG result;
	GetProperty(0x6, VT_I4, (void*)&result);
	return result;
}

LONG CclsTimeLine::GetTop()
{
	LONG result;
	GetProperty(0x7, VT_I4, (void*)&result);
	return result;
}

LONG CclsTimeLine::GetBottom()
{
	LONG result;
	GetProperty(0x8, VT_I4, (void*)&result);
	return result;
}

CclsTimeLineScrollBar CclsTimeLine::GetTimeLineScrollBar()
{
	LPDISPATCH pDispatch;
	GetProperty(0x9, VT_DISPATCH, (void*)&pDispatch);
	return CclsTimeLineScrollBar(pDispatch);
}

CclsTierArea CclsTimeLine::GetTierArea()
{
	LPDISPATCH pDispatch;
	GetProperty(0xa, VT_DISPATCH, (void*)&pDispatch);
	return CclsTierArea(pDispatch);
}

CclsTickMarkArea CclsTimeLine::GetTickMarkArea()
{
	LPDISPATCH pDispatch;
	GetProperty(0xb, VT_DISPATCH, (void*)&pDispatch);
	return CclsTickMarkArea(pDispatch);
}

CclsProgressLine CclsTimeLine::GetProgressLine()
{
	LPDISPATCH pDispatch;
	GetProperty(0xc, VT_DISPATCH, (void*)&pDispatch);
	return CclsProgressLine(pDispatch);
}

void CclsTimeLine::Move(E_INTERVAL Interval, LONG Factor)
{
	static BYTE params[] = VTS_I4 VTS_I4;
	InvokeHelper(0xd, DISPATCH_METHOD, VT_EMPTY, NULL, params, Interval, Factor);
}

void CclsTimeLine::Position(DATE TimeLineStartDate)
{
	static BYTE params[] = VTS_DATE;
	InvokeHelper(0xe, DISPATCH_METHOD, VT_EMPTY, NULL, params, TimeLineStartDate);
}

CString CclsTimeLine::GetXML(void)
{
	CString oReturn;
	InvokeHelper(0xf, DISPATCH_METHOD, VT_BSTR, (void*)&oReturn, NULL);
	return oReturn;
}

void CclsTimeLine::SetXML(LPCTSTR sXML)
{
	static BYTE params[] = VTS_BSTR;
	InvokeHelper(0x10, DISPATCH_METHOD, VT_EMPTY, NULL, params, sXML);
}

void CclsView::SetKey(LPCTSTR propval)
{
	SetProperty(0x1, VT_BSTR, propval);
}

CString CclsView::GetKey()
{
	CString result;
	GetProperty(0x1, VT_BSTR, (void*)&result);
	return result;
}

CclsTimeLine CclsView::GetTimeLine()
{
	LPDISPATCH pDispatch;
	GetProperty(0x2, VT_DISPATCH, (void*)&pDispatch);
	return CclsTimeLine(pDispatch);
}

CclsClientArea CclsView::GetClientArea()
{
	LPDISPATCH pDispatch;
	GetProperty(0x3, VT_DISPATCH, (void*)&pDispatch);
	return CclsClientArea(pDispatch);
}

void CclsView::SetTag(LPCTSTR propval)
{
	SetProperty(0x4, VT_BSTR, propval);
}

CString CclsView::GetTag()
{
	CString result;
	GetProperty(0x4, VT_BSTR, (void*)&result);
	return result;
}

void CclsView::SetInterval(E_INTERVAL propval)
{
	LONG lpropval = (LONG)propval;
	SetProperty(0x5, VT_I4, lpropval);
}

E_INTERVAL CclsView::GetInterval()
{
	LONG result;
	GetProperty(0x5, VT_I4, (void*)&result);
	return (E_INTERVAL)result;
}

LONG CclsView::GetIndex()
{
	LONG result;
	GetProperty(0x8, VT_I4, (void*)&result);
	return result;
}

void CclsView::SetFactor(LONG propval)
{
	SetProperty(0x9, VT_I4, propval);
}

LONG CclsView::GetFactor()
{
	LONG result;
	GetProperty(0x9, VT_I4, (void*)&result);
	return result;
}

CString CclsView::GetXML(void)
{
	CString oReturn;
	InvokeHelper(0x6, DISPATCH_METHOD, VT_BSTR, (void*)&oReturn, NULL);
	return oReturn;
}

void CclsView::SetXML(LPCTSTR sXML)
{
	static BYTE params[] = VTS_BSTR;
	InvokeHelper(0x7, DISPATCH_METHOD, VT_EMPTY, NULL, params, sXML);
}

void CclsView::Clear(void)
{
	InvokeHelper(0xa, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
}

CclsView CclsView::Clone(LPCTSTR Key)
{
	CclsView oReturn;
	static BYTE params[] = VTS_BSTR;
	InvokeHelper(0xb, DISPATCH_METHOD, VT_DISPATCH, (void*)&oReturn, params, Key);
	return oReturn;
}

LONG CclsViews::GetCount()
{
	LONG result;
	GetProperty(0x1, VT_I4, (void*)&result);
	return result;
}

CclsView CclsViews::Item(LPCTSTR Index)
{
	CclsView oReturn;
	static BYTE params[] = VTS_BSTR;
	InvokeHelper(0x2, DISPATCH_METHOD, VT_DISPATCH, (void*)&oReturn, params, Index);
	return oReturn;
}

CclsView CclsViews::Add(E_INTERVAL Interval, LONG Factor, E_TIERTYPE UpperTierType, E_TIERTYPE MiddleTierType, E_TIERTYPE LowerTierType, LPCTSTR Key)
{
	CclsView oReturn;
	static BYTE params[] = VTS_I4 VTS_I4 VTS_I4 VTS_I4 VTS_I4 VTS_BSTR;
	InvokeHelper(0x3, DISPATCH_METHOD, VT_DISPATCH, (void*)&oReturn, params, Interval, Factor, UpperTierType, MiddleTierType, LowerTierType, Key);
	return oReturn;
}

void CclsViews::Clear(void)
{
	InvokeHelper(0x4, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
}

void CclsViews::Remove(LPCTSTR Index)
{
	static BYTE params[] = VTS_BSTR;
	InvokeHelper(0x5, DISPATCH_METHOD, VT_EMPTY, NULL, params, Index);
}

CString CclsViews::GetXML(void)
{
	CString oReturn;
	InvokeHelper(0x6, DISPATCH_METHOD, VT_BSTR, (void*)&oReturn, NULL);
	return oReturn;
}

void CclsViews::SetXML(LPCTSTR sXML)
{
	static BYTE params[] = VTS_BSTR;
	InvokeHelper(0x7, DISPATCH_METHOD, VT_EMPTY, NULL, params, sXML);
}

void CCustomTierDrawEventArgs::SetText(LPCTSTR propval)
{
	SetProperty(0x1, VT_BSTR, propval);
}

CString CCustomTierDrawEventArgs::GetText()
{
	CString result;
	GetProperty(0x1, VT_BSTR, (void*)&result);
	return result;
}

void CCustomTierDrawEventArgs::SetCustomDraw(BOOL propval)
{
	SetProperty(0x2, VT_BOOL, propval);
}

BOOL CCustomTierDrawEventArgs::GetCustomDraw()
{
	BOOL result;
	GetProperty(0x2, VT_BOOL, (void*)&result);
	return result;
}

void CCustomTierDrawEventArgs::SetStyleIndex(LPCTSTR propval)
{
	SetProperty(0x3, VT_BSTR, propval);
}

CString CCustomTierDrawEventArgs::GetStyleIndex()
{
	CString result;
	GetProperty(0x3, VT_BSTR, (void*)&result);
	return result;
}

void CCustomTierDrawEventArgs::SetTierPosition(E_TIERPOSITION propval)
{
	LONG lpropval = (LONG)propval;
	SetProperty(0x4, VT_I4, lpropval);
}

E_TIERPOSITION CCustomTierDrawEventArgs::GetTierPosition()
{
	LONG result;
	GetProperty(0x4, VT_I4, (void*)&result);
	return (E_TIERPOSITION)result;
}

void CCustomTierDrawEventArgs::SetStartDate(DATE propval)
{
	SetProperty(0x5, VT_DATE, propval);
}

DATE CCustomTierDrawEventArgs::GetStartDate()
{
	DATE result;
	GetProperty(0x5, VT_DATE, (void*)&result);
	return result;
}

void CCustomTierDrawEventArgs::SetEndDate(DATE propval)
{
	SetProperty(0x6, VT_DATE, propval);
}

DATE CCustomTierDrawEventArgs::GetEndDate()
{
	DATE result;
	GetProperty(0x6, VT_DATE, (void*)&result);
	return result;
}

void CCustomTierDrawEventArgs::SetLeft(LONG propval)
{
	SetProperty(0x7, VT_I4, propval);
}

LONG CCustomTierDrawEventArgs::GetLeft()
{
	LONG result;
	GetProperty(0x7, VT_I4, (void*)&result);
	return result;
}

void CCustomTierDrawEventArgs::SetTop(LONG propval)
{
	SetProperty(0x8, VT_I4, propval);
}

LONG CCustomTierDrawEventArgs::GetTop()
{
	LONG result;
	GetProperty(0x8, VT_I4, (void*)&result);
	return result;
}

void CCustomTierDrawEventArgs::SetRight(LONG propval)
{
	SetProperty(0x9, VT_I4, propval);
}

LONG CCustomTierDrawEventArgs::GetRight()
{
	LONG result;
	GetProperty(0x9, VT_I4, (void*)&result);
	return result;
}

void CCustomTierDrawEventArgs::SetBottom(LONG propval)
{
	SetProperty(0xa, VT_I4, propval);
}

LONG CCustomTierDrawEventArgs::GetBottom()
{
	LONG result;
	GetProperty(0xa, VT_I4, (void*)&result);
	return result;
}

void CCustomTierDrawEventArgs::SetLeftTrim(LONG propval)
{
	SetProperty(0xb, VT_I4, propval);
}

LONG CCustomTierDrawEventArgs::GetLeftTrim()
{
	LONG result;
	GetProperty(0xb, VT_I4, (void*)&result);
	return result;
}

void CCustomTierDrawEventArgs::SetRightTrim(LONG propval)
{
	SetProperty(0xc, VT_I4, propval);
}

LONG CCustomTierDrawEventArgs::GetRightTrim()
{
	LONG result;
	GetProperty(0xc, VT_I4, (void*)&result);
	return result;
}

CclsGDIGraphics CCustomTierDrawEventArgs::GetGraphics()
{
	LPDISPATCH pDispatch;
	GetProperty(0xd, VT_DISPATCH, (void*)&pDispatch);
	return CclsGDIGraphics(pDispatch);
}

void CCustomTierDrawEventArgs::SetInterval(E_INTERVAL propval)
{
	LONG lpropval = (LONG)propval;
	SetProperty(0xe, VT_I4, lpropval);
}

E_INTERVAL CCustomTierDrawEventArgs::GetInterval()
{
	LONG result;
	GetProperty(0xe, VT_I4, (void*)&result);
	return (E_INTERVAL)result;
}

void CCustomTierDrawEventArgs::SetFactor(LONG propval)
{
	SetProperty(0xf, VT_I4, propval);
}

LONG CCustomTierDrawEventArgs::GetFactor()
{
	LONG result;
	GetProperty(0xf, VT_I4, (void*)&result);
	return result;
}

void CDrawEventArgs::SetEventTarget(E_EVENTTARGET propval)
{
	LONG lpropval = (LONG)propval;
	SetProperty(0x1, VT_I4, lpropval);
}

E_EVENTTARGET CDrawEventArgs::GetEventTarget()
{
	LONG result;
	GetProperty(0x1, VT_I4, (void*)&result);
	return (E_EVENTTARGET)result;
}

void CDrawEventArgs::SetCustomDraw(BOOL propval)
{
	SetProperty(0x2, VT_BOOL, propval);
}

BOOL CDrawEventArgs::GetCustomDraw()
{
	BOOL result;
	GetProperty(0x2, VT_BOOL, (void*)&result);
	return result;
}

void CDrawEventArgs::SetObjectIndex(LONG propval)
{
	SetProperty(0x3, VT_I4, propval);
}

LONG CDrawEventArgs::GetObjectIndex()
{
	LONG result;
	GetProperty(0x3, VT_I4, (void*)&result);
	return result;
}

void CDrawEventArgs::SetParentObjectIndex(LONG propval)
{
	SetProperty(0x4, VT_I4, propval);
}

LONG CDrawEventArgs::GetParentObjectIndex()
{
	LONG result;
	GetProperty(0x4, VT_I4, (void*)&result);
	return result;
}

CclsGDIGraphics CDrawEventArgs::GetGraphics()
{
	LPDISPATCH pDispatch;
	GetProperty(0x5, VT_DISPATCH, (void*)&pDispatch);
	return CclsGDIGraphics(pDispatch);
}

void CErrorEventArgs::SetNumber(LONG propval)
{
	SetProperty(0x1, VT_I4, propval);
}

LONG CErrorEventArgs::GetNumber()
{
	LONG result;
	GetProperty(0x1, VT_I4, (void*)&result);
	return result;
}

void CErrorEventArgs::SetDescription(LPCTSTR propval)
{
	SetProperty(0x2, VT_BSTR, propval);
}

CString CErrorEventArgs::GetDescription()
{
	CString result;
	GetProperty(0x2, VT_BSTR, (void*)&result);
	return result;
}

void CErrorEventArgs::SetSource(LPCTSTR propval)
{
	SetProperty(0x3, VT_BSTR, propval);
}

CString CErrorEventArgs::GetSource()
{
	CString result;
	GetProperty(0x3, VT_BSTR, (void*)&result);
	return result;
}

void CKeyEventArgs::SetKeyCode(E_KEYS propval)
{
	LONG lpropval = (LONG)propval;
	SetProperty(0x1, VT_I2, lpropval);
}

E_KEYS CKeyEventArgs::GetKeyCode()
{
	LONG result;
	GetProperty(0x1, VT_I2, (void*)&result);
	return (E_KEYS)result;
}

void CKeyEventArgs::SetCancel(BOOL propval)
{
	SetProperty(0x2, VT_BOOL, propval);
}

BOOL CKeyEventArgs::GetCancel()
{
	BOOL result;
	GetProperty(0x2, VT_BOOL, (void*)&result);
	return result;
}

void CKeyEventArgs::SetCharacterCode(SHORT propval)
{
	SetProperty(0x3, VT_I2, propval);
}

SHORT CKeyEventArgs::GetCharacterCode()
{
	SHORT result;
	GetProperty(0x3, VT_I2, (void*)&result);
	return result;
}

BOOL CKeyEventArgs::GetAlt()
{
	BOOL result;
	GetProperty(0x4, VT_BOOL, (void*)&result);
	return result;
}

BOOL CKeyEventArgs::GetShift()
{
	BOOL result;
	GetProperty(0x5, VT_BOOL, (void*)&result);
	return result;
}

BOOL CKeyEventArgs::GetControl()
{
	BOOL result;
	GetProperty(0x6, VT_BOOL, (void*)&result);
	return result;
}

void CCustomTickMarkAreaDrawEventArgs::SetCustomDraw(BOOL propval)
{
	SetProperty(0x1, VT_BOOL, propval);
}

BOOL CCustomTickMarkAreaDrawEventArgs::GetCustomDraw()
{
	BOOL result;
	GetProperty(0x1, VT_BOOL, (void*)&result);
	return result;
}

CclsGDIGraphics CCustomTickMarkAreaDrawEventArgs::GetGraphics()
{
	LPDISPATCH pDispatch;
	GetProperty(0x2, VT_DISPATCH, (void*)&pDispatch);
	return CclsGDIGraphics(pDispatch);
}

LONG CCustomTickMarkAreaDrawEventArgs::GetLeft()
{
	LONG result;
	GetProperty(0x3, VT_I4, (void*)&result);
	return result;
}

LONG CCustomTickMarkAreaDrawEventArgs::GetTop()
{
	LONG result;
	GetProperty(0x4, VT_I4, (void*)&result);
	return result;
}

LONG CCustomTickMarkAreaDrawEventArgs::GetRight()
{
	LONG result;
	GetProperty(0x5, VT_I4, (void*)&result);
	return result;
}

LONG CCustomTickMarkAreaDrawEventArgs::GetBottom()
{
	LONG result;
	GetProperty(0x6, VT_I4, (void*)&result);
	return result;
}

DATE CCustomTickMarkAreaDrawEventArgs::GetdtDate()
{
	DATE result;
	GetProperty(0x7, VT_DATE, (void*)&result);
	return result;
}

CclsTickMark CCustomTickMarkAreaDrawEventArgs::GetoTickMark()
{
	LPDISPATCH pDispatch;
	GetProperty(0x8, VT_DISPATCH, (void*)&pDispatch);
	return CclsTickMark(pDispatch);
}

LONG CCustomTickMarkAreaDrawEventArgs::GetX()
{
	LONG result;
	GetProperty(0x9, VT_I4, (void*)&result);
	return result;
}

void CTickMarkTextDrawEventArgs::SetCustomDraw(BOOL propval)
{
	SetProperty(0x1, VT_BOOL, propval);
}

BOOL CTickMarkTextDrawEventArgs::GetCustomDraw()
{
	BOOL result;
	GetProperty(0x1, VT_BOOL, (void*)&result);
	return result;
}

void CTickMarkTextDrawEventArgs::SetText(LPCTSTR propval)
{
	SetProperty(0x2, VT_BSTR, propval);
}

CString CTickMarkTextDrawEventArgs::GetText()
{
	CString result;
	GetProperty(0x2, VT_BSTR, (void*)&result);
	return result;
}

DATE CTickMarkTextDrawEventArgs::GetdtDate()
{
	DATE result;
	GetProperty(0x3, VT_DATE, (void*)&result);
	return result;
}

void CMouseEventArgs::SetX(LONG propval)
{
	SetProperty(0x1, VT_I4, propval);
}

LONG CMouseEventArgs::GetX()
{
	LONG result;
	GetProperty(0x1, VT_I4, (void*)&result);
	return result;
}

void CMouseEventArgs::SetY(LONG propval)
{
	SetProperty(0x2, VT_I4, propval);
}

LONG CMouseEventArgs::GetY()
{
	LONG result;
	GetProperty(0x2, VT_I4, (void*)&result);
	return result;
}

void CMouseEventArgs::SetEventTarget(E_EVENTTARGET propval)
{
	LONG lpropval = (LONG)propval;
	SetProperty(0x3, VT_I4, lpropval);
}

E_EVENTTARGET CMouseEventArgs::GetEventTarget()
{
	LONG result;
	GetProperty(0x3, VT_I4, (void*)&result);
	return (E_EVENTTARGET)result;
}

void CMouseEventArgs::SetOperation(E_OPERATION propval)
{
	LONG lpropval = (LONG)propval;
	SetProperty(0x4, VT_I4, lpropval);
}

E_OPERATION CMouseEventArgs::GetOperation()
{
	LONG result;
	GetProperty(0x4, VT_I4, (void*)&result);
	return (E_OPERATION)result;
}

void CMouseEventArgs::SetButton(E_MOUSEBUTTONS propval)
{
	LONG lpropval = (LONG)propval;
	SetProperty(0x5, VT_I4, lpropval);
}

E_MOUSEBUTTONS CMouseEventArgs::GetButton()
{
	LONG result;
	GetProperty(0x5, VT_I4, (void*)&result);
	return (E_MOUSEBUTTONS)result;
}

void CMouseEventArgs::SetCancel(BOOL propval)
{
	SetProperty(0x6, VT_BOOL, propval);
}

BOOL CMouseEventArgs::GetCancel()
{
	BOOL result;
	GetProperty(0x6, VT_BOOL, (void*)&result);
	return result;
}

void CMouseWheelEventArgs::SetX(LONG propval)
{
	SetProperty(0x1, VT_I4, propval);
}

LONG CMouseWheelEventArgs::GetX()
{
	LONG result;
	GetProperty(0x1, VT_I4, (void*)&result);
	return result;
}

void CMouseWheelEventArgs::SetY(LONG propval)
{
	SetProperty(0x2, VT_I4, propval);
}

LONG CMouseWheelEventArgs::GetY()
{
	LONG result;
	GetProperty(0x2, VT_I4, (void*)&result);
	return result;
}

void CMouseWheelEventArgs::SetButton(E_MOUSEBUTTONS propval)
{
	LONG lpropval = (LONG)propval;
	SetProperty(0x3, VT_I4, lpropval);
}

E_MOUSEBUTTONS CMouseWheelEventArgs::GetButton()
{
	LONG result;
	GetProperty(0x3, VT_I4, (void*)&result);
	return (E_MOUSEBUTTONS)result;
}

LONG CMouseWheelEventArgs::GetDelta()
{
	LONG result;
	GetProperty(0x4, VT_I4, (void*)&result);
	return result;
}

void CNodeEventArgs::SetIndex(LONG propval)
{
	SetProperty(0x1, VT_I4, propval);
}

LONG CNodeEventArgs::GetIndex()
{
	LONG result;
	GetProperty(0x1, VT_I4, (void*)&result);
	return result;
}

void CObjectAddedEventArgs::SetTaskIndex(LONG propval)
{
	SetProperty(0x1, VT_I4, propval);
}

LONG CObjectAddedEventArgs::GetTaskIndex()
{
	LONG result;
	GetProperty(0x1, VT_I4, (void*)&result);
	return result;
}

void CObjectAddedEventArgs::SetPredecessorObjectIndex(LONG propval)
{
	SetProperty(0x2, VT_I4, propval);
}

LONG CObjectAddedEventArgs::GetPredecessorObjectIndex()
{
	LONG result;
	GetProperty(0x2, VT_I4, (void*)&result);
	return result;
}

void CObjectAddedEventArgs::SetPredecessorTaskIndex(LONG propval)
{
	SetProperty(0x3, VT_I4, propval);
}

LONG CObjectAddedEventArgs::GetPredecessorTaskIndex()
{
	LONG result;
	GetProperty(0x3, VT_I4, (void*)&result);
	return result;
}

void CObjectAddedEventArgs::SetPredecessorType(E_CONSTRAINTTYPE propval)
{
	LONG lpropval = (LONG)propval;
	SetProperty(0x4, VT_I4, lpropval);
}

E_CONSTRAINTTYPE CObjectAddedEventArgs::GetPredecessorType()
{
	LONG result;
	GetProperty(0x4, VT_I4, (void*)&result);
	return (E_CONSTRAINTTYPE)result;
}

void CObjectAddedEventArgs::SetTaskKey(LPCTSTR propval)
{
	SetProperty(0x5, VT_BSTR, propval);
}

CString CObjectAddedEventArgs::GetTaskKey()
{
	CString result;
	GetProperty(0x5, VT_BSTR, (void*)&result);
	return result;
}

void CObjectAddedEventArgs::SetPredecessorTaskKey(LPCTSTR propval)
{
	SetProperty(0x6, VT_BSTR, propval);
}

CString CObjectAddedEventArgs::GetPredecessorTaskKey()
{
	CString result;
	GetProperty(0x6, VT_BSTR, (void*)&result);
	return result;
}

void CObjectAddedEventArgs::SetEventTarget(E_EVENTTARGET propval)
{
	LONG lpropval = (LONG)propval;
	SetProperty(0x7, VT_I4, lpropval);
}

E_EVENTTARGET CObjectAddedEventArgs::GetEventTarget()
{
	LONG result;
	GetProperty(0x7, VT_I4, (void*)&result);
	return (E_EVENTTARGET)result;
}

void CObjectSelectedEventArgs::SetEventTarget(E_EVENTTARGET propval)
{
	LONG lpropval = (LONG)propval;
	SetProperty(0x1, VT_I4, lpropval);
}

E_EVENTTARGET CObjectSelectedEventArgs::GetEventTarget()
{
	LONG result;
	GetProperty(0x1, VT_I4, (void*)&result);
	return (E_EVENTTARGET)result;
}

void CObjectSelectedEventArgs::SetObjectIndex(LONG propval)
{
	SetProperty(0x2, VT_I4, propval);
}

LONG CObjectSelectedEventArgs::GetObjectIndex()
{
	LONG result;
	GetProperty(0x2, VT_I4, (void*)&result);
	return result;
}

void CObjectSelectedEventArgs::SetParentObjectIndex(LONG propval)
{
	SetProperty(0x3, VT_I4, propval);
}

LONG CObjectSelectedEventArgs::GetParentObjectIndex()
{
	LONG result;
	GetProperty(0x3, VT_I4, (void*)&result);
	return result;
}

void CObjectStateChangedEventArgs::SetEventTarget(E_EVENTTARGET propval)
{
	LONG lpropval = (LONG)propval;
	SetProperty(0x1, VT_I4, lpropval);
}

E_EVENTTARGET CObjectStateChangedEventArgs::GetEventTarget()
{
	LONG result;
	GetProperty(0x1, VT_I4, (void*)&result);
	return (E_EVENTTARGET)result;
}

void CObjectStateChangedEventArgs::SetIndex(LONG propval)
{
	SetProperty(0x2, VT_I4, propval);
}

LONG CObjectStateChangedEventArgs::GetIndex()
{
	LONG result;
	GetProperty(0x2, VT_I4, (void*)&result);
	return result;
}

void CObjectStateChangedEventArgs::SetCancel(BOOL propval)
{
	SetProperty(0x3, VT_BOOL, propval);
}

BOOL CObjectStateChangedEventArgs::GetCancel()
{
	BOOL result;
	GetProperty(0x3, VT_BOOL, (void*)&result);
	return result;
}

void CObjectStateChangedEventArgs::SetDestinationIndex(LONG propval)
{
	SetProperty(0x4, VT_I4, propval);
}

LONG CObjectStateChangedEventArgs::GetDestinationIndex()
{
	LONG result;
	GetProperty(0x4, VT_I4, (void*)&result);
	return result;
}

void CObjectStateChangedEventArgs::SetInitialRowIndex(LONG propval)
{
	SetProperty(0x5, VT_I4, propval);
}

LONG CObjectStateChangedEventArgs::GetInitialRowIndex()
{
	LONG result;
	GetProperty(0x5, VT_I4, (void*)&result);
	return result;
}

void CObjectStateChangedEventArgs::SetFinalRowIndex(LONG propval)
{
	SetProperty(0x6, VT_I4, propval);
}

LONG CObjectStateChangedEventArgs::GetFinalRowIndex()
{
	LONG result;
	GetProperty(0x6, VT_I4, (void*)&result);
	return result;
}

void CObjectStateChangedEventArgs::SetInitialColumnIndex(LONG propval)
{
	SetProperty(0x7, VT_I4, propval);
}

LONG CObjectStateChangedEventArgs::GetInitialColumnIndex()
{
	LONG result;
	GetProperty(0x7, VT_I4, (void*)&result);
	return result;
}

void CObjectStateChangedEventArgs::SetFinalColumnIndex(LONG propval)
{
	SetProperty(0x8, VT_I4, propval);
}

LONG CObjectStateChangedEventArgs::GetFinalColumnIndex()
{
	LONG result;
	GetProperty(0x8, VT_I4, (void*)&result);
	return result;
}

DATE CObjectStateChangedEventArgs::GetStartDate()
{
	DATE result;
	GetProperty(0x9, VT_DATE, (void*)&result);
	return result;
}

DATE CObjectStateChangedEventArgs::GetEndDate()
{
	DATE result;
	GetProperty(0xa, VT_DATE, (void*)&result);
	return result;
}

E_CHANGETYPE CObjectStateChangedEventArgs::GetChangeType()
{
	LONG result;
	GetProperty(0xb, VT_I4, (void*)&result);
	return (E_CHANGETYPE)result;
}

void CPredecessorExceptionEventArgs::SetPredecessorIndex(LONG propval)
{
	SetProperty(0x1, VT_I4, propval);
}

LONG CPredecessorExceptionEventArgs::GetPredecessorIndex()
{
	LONG result;
	GetProperty(0x1, VT_I4, (void*)&result);
	return result;
}

void CPredecessorExceptionEventArgs::SetPredecessorType(E_CONSTRAINTTYPE propval)
{
	LONG lpropval = (LONG)propval;
	SetProperty(0x2, VT_I4, lpropval);
}

E_CONSTRAINTTYPE CPredecessorExceptionEventArgs::GetPredecessorType()
{
	LONG result;
	GetProperty(0x2, VT_I4, (void*)&result);
	return (E_CONSTRAINTTYPE)result;
}

void CPredecessorDrawEventArgs::SetCustomDraw(BOOL propval)
{
	SetProperty(0x1, VT_BOOL, propval);
}

BOOL CPredecessorDrawEventArgs::GetCustomDraw()
{
	BOOL result;
	GetProperty(0x1, VT_BOOL, (void*)&result);
	return result;
}

CclsGDIGraphics CPredecessorDrawEventArgs::GetGraphics()
{
	LPDISPATCH pDispatch;
	GetProperty(0x2, VT_DISPATCH, (void*)&pDispatch);
	return CclsGDIGraphics(pDispatch);
}

void CPredecessorDrawEventArgs::SetPredecessorIndex(LONG propval)
{
	SetProperty(0x3, VT_I4, propval);
}

LONG CPredecessorDrawEventArgs::GetPredecessorIndex()
{
	LONG result;
	GetProperty(0x3, VT_I4, (void*)&result);
	return result;
}

void CPredecessorDrawEventArgs::SetTaskIndex(LONG propval)
{
	SetProperty(0x4, VT_I4, propval);
}

LONG CPredecessorDrawEventArgs::GetTaskIndex()
{
	LONG result;
	GetProperty(0x4, VT_I4, (void*)&result);
	return result;
}

void CPredecessorDrawEventArgs::SetPredecessorType(E_CONSTRAINTTYPE propval)
{
	LONG lpropval = (LONG)propval;
	SetProperty(0x5, VT_I4, lpropval);
}

E_CONSTRAINTTYPE CPredecessorDrawEventArgs::GetPredecessorType()
{
	LONG result;
	GetProperty(0x5, VT_I4, (void*)&result);
	return (E_CONSTRAINTTYPE)result;
}

void CScrollEventArgs::SetScrollBarType(E_SCROLLBAR propval)
{
	LONG lpropval = (LONG)propval;
	SetProperty(0x1, VT_I4, lpropval);
}

E_SCROLLBAR CScrollEventArgs::GetScrollBarType()
{
	LONG result;
	GetProperty(0x1, VT_I4, (void*)&result);
	return (E_SCROLLBAR)result;
}

void CScrollEventArgs::SetOffset(LONG propval)
{
	SetProperty(0x2, VT_I4, propval);
}

LONG CScrollEventArgs::GetOffset()
{
	LONG result;
	GetProperty(0x2, VT_I4, (void*)&result);
	return result;
}

void CTextEditEventArgs::SetObjectType(E_TEXTOBJECTTYPE propval)
{
	LONG lpropval = (LONG)propval;
	SetProperty(0x1, VT_I4, lpropval);
}

E_TEXTOBJECTTYPE CTextEditEventArgs::GetObjectType()
{
	LONG result;
	GetProperty(0x1, VT_I4, (void*)&result);
	return (E_TEXTOBJECTTYPE)result;
}

void CTextEditEventArgs::SetObjectIndex(LONG propval)
{
	SetProperty(0x2, VT_I4, propval);
}

LONG CTextEditEventArgs::GetObjectIndex()
{
	LONG result;
	GetProperty(0x2, VT_I4, (void*)&result);
	return result;
}

void CTextEditEventArgs::SetParentObjectIndex(LONG propval)
{
	SetProperty(0x3, VT_I4, propval);
}

LONG CTextEditEventArgs::GetParentObjectIndex()
{
	LONG result;
	GetProperty(0x3, VT_I4, (void*)&result);
	return result;
}

void CTextEditEventArgs::SetText(LPCTSTR propval)
{
	SetProperty(0x4, VT_BSTR, propval);
}

CString CTextEditEventArgs::GetText()
{
	CString result;
	GetProperty(0x4, VT_BSTR, (void*)&result);
	return result;
}

void CToolTipEventArgs::SetInitialRowIndex(LONG propval)
{
	SetProperty(0x1, VT_I4, propval);
}

LONG CToolTipEventArgs::GetInitialRowIndex()
{
	LONG result;
	GetProperty(0x1, VT_I4, (void*)&result);
	return result;
}

void CToolTipEventArgs::SetFinalRowIndex(LONG propval)
{
	SetProperty(0x2, VT_I4, propval);
}

LONG CToolTipEventArgs::GetFinalRowIndex()
{
	LONG result;
	GetProperty(0x2, VT_I4, (void*)&result);
	return result;
}

void CToolTipEventArgs::SetTaskIndex(LONG propval)
{
	SetProperty(0x3, VT_I4, propval);
}

LONG CToolTipEventArgs::GetTaskIndex()
{
	LONG result;
	GetProperty(0x3, VT_I4, (void*)&result);
	return result;
}

void CToolTipEventArgs::SetMilestoneIndex(LONG propval)
{
	SetProperty(0x4, VT_I4, propval);
}

LONG CToolTipEventArgs::GetMilestoneIndex()
{
	LONG result;
	GetProperty(0x4, VT_I4, (void*)&result);
	return result;
}

void CToolTipEventArgs::SetPercentageIndex(LONG propval)
{
	SetProperty(0x5, VT_I4, propval);
}

LONG CToolTipEventArgs::GetPercentageIndex()
{
	LONG result;
	GetProperty(0x5, VT_I4, (void*)&result);
	return result;
}

void CToolTipEventArgs::SetRowIndex(LONG propval)
{
	SetProperty(0x6, VT_I4, propval);
}

LONG CToolTipEventArgs::GetRowIndex()
{
	LONG result;
	GetProperty(0x6, VT_I4, (void*)&result);
	return result;
}

void CToolTipEventArgs::SetCellIndex(LONG propval)
{
	SetProperty(0x7, VT_I4, propval);
}

LONG CToolTipEventArgs::GetCellIndex()
{
	LONG result;
	GetProperty(0x7, VT_I4, (void*)&result);
	return result;
}

void CToolTipEventArgs::SetColumnIndex(LONG propval)
{
	SetProperty(0x8, VT_I4, propval);
}

LONG CToolTipEventArgs::GetColumnIndex()
{
	LONG result;
	GetProperty(0x8, VT_I4, (void*)&result);
	return result;
}

void CToolTipEventArgs::SetInitialStartDate(DATE propval)
{
	SetProperty(0x9, VT_DATE, propval);
}

DATE CToolTipEventArgs::GetInitialStartDate()
{
	DATE result;
	GetProperty(0x9, VT_DATE, (void*)&result);
	return result;
}

void CToolTipEventArgs::SetInitialEndDate(DATE propval)
{
	SetProperty(0xa, VT_DATE, propval);
}

DATE CToolTipEventArgs::GetInitialEndDate()
{
	DATE result;
	GetProperty(0xa, VT_DATE, (void*)&result);
	return result;
}

void CToolTipEventArgs::SetStartDate(DATE propval)
{
	SetProperty(0xb, VT_DATE, propval);
}

DATE CToolTipEventArgs::GetStartDate()
{
	DATE result;
	GetProperty(0xb, VT_DATE, (void*)&result);
	return result;
}

void CToolTipEventArgs::SetEndDate(DATE propval)
{
	SetProperty(0xc, VT_DATE, propval);
}

DATE CToolTipEventArgs::GetEndDate()
{
	DATE result;
	GetProperty(0xc, VT_DATE, (void*)&result);
	return result;
}

void CToolTipEventArgs::SetXStart(LONG propval)
{
	SetProperty(0xd, VT_I4, propval);
}

LONG CToolTipEventArgs::GetXStart()
{
	LONG result;
	GetProperty(0xd, VT_I4, (void*)&result);
	return result;
}

void CToolTipEventArgs::SetXEnd(LONG propval)
{
	SetProperty(0xe, VT_I4, propval);
}

LONG CToolTipEventArgs::GetXEnd()
{
	LONG result;
	GetProperty(0xe, VT_I4, (void*)&result);
	return result;
}

void CToolTipEventArgs::SetOperation(E_OPERATION propval)
{
	LONG lpropval = (LONG)propval;
	SetProperty(0xf, VT_I4, lpropval);
}

E_OPERATION CToolTipEventArgs::GetOperation()
{
	LONG result;
	GetProperty(0xf, VT_I4, (void*)&result);
	return (E_OPERATION)result;
}

void CToolTipEventArgs::SetEventTarget(E_EVENTTARGET propval)
{
	LONG lpropval = (LONG)propval;
	SetProperty(0x10, VT_I4, lpropval);
}

E_EVENTTARGET CToolTipEventArgs::GetEventTarget()
{
	LONG result;
	GetProperty(0x10, VT_I4, (void*)&result);
	return (E_EVENTTARGET)result;
}

void CToolTipEventArgs::SetTaskPosition(LPCTSTR propval)
{
	SetProperty(0x11, VT_BSTR, propval);
}

CString CToolTipEventArgs::GetTaskPosition()
{
	CString result;
	GetProperty(0x11, VT_BSTR, (void*)&result);
	return result;
}

void CToolTipEventArgs::SetPredecessorPosition(LPCTSTR propval)
{
	SetProperty(0x12, VT_BSTR, propval);
}

CString CToolTipEventArgs::GetPredecessorPosition()
{
	CString result;
	GetProperty(0x12, VT_BSTR, (void*)&result);
	return result;
}

void CToolTipEventArgs::SetX(LONG propval)
{
	SetProperty(0x13, VT_I4, propval);
}

LONG CToolTipEventArgs::GetX()
{
	LONG result;
	GetProperty(0x13, VT_I4, (void*)&result);
	return result;
}

void CToolTipEventArgs::SetY(LONG propval)
{
	SetProperty(0x14, VT_I4, propval);
}

LONG CToolTipEventArgs::GetY()
{
	LONG result;
	GetProperty(0x14, VT_I4, (void*)&result);
	return result;
}

void CToolTipEventArgs::SetX1(LONG propval)
{
	SetProperty(0x15, VT_I4, propval);
}

LONG CToolTipEventArgs::GetX1()
{
	LONG result;
	GetProperty(0x15, VT_I4, (void*)&result);
	return result;
}

void CToolTipEventArgs::SetY1(LONG propval)
{
	SetProperty(0x16, VT_I4, propval);
}

LONG CToolTipEventArgs::GetY1()
{
	LONG result;
	GetProperty(0x16, VT_I4, (void*)&result);
	return result;
}

void CToolTipEventArgs::SetX2(LONG propval)
{
	SetProperty(0x17, VT_I4, propval);
}

LONG CToolTipEventArgs::GetX2()
{
	LONG result;
	GetProperty(0x17, VT_I4, (void*)&result);
	return result;
}

void CToolTipEventArgs::SetY2(LONG propval)
{
	SetProperty(0x18, VT_I4, propval);
}

LONG CToolTipEventArgs::GetY2()
{
	LONG result;
	GetProperty(0x18, VT_I4, (void*)&result);
	return result;
}

void CToolTipEventArgs::SetCustomDraw(BOOL propval)
{
	SetProperty(0x19, VT_BOOL, propval);
}

BOOL CToolTipEventArgs::GetCustomDraw()
{
	BOOL result;
	GetProperty(0x19, VT_BOOL, (void*)&result);
	return result;
}

CclsGDIGraphics CToolTipEventArgs::GetGraphics()
{
	LPDISPATCH pDispatch;
	GetProperty(0x1a, VT_DISPATCH, (void*)&pDispatch);
	return CclsGDIGraphics(pDispatch);
}

void CToolTipEventArgs::SetToolTipType(E_TOOLTIPTYPE propval)
{
	LONG lpropval = (LONG)propval;
	SetProperty(0x1b, VT_I4, lpropval);
}

E_TOOLTIPTYPE CToolTipEventArgs::GetToolTipType()
{
	LONG result;
	GetProperty(0x1b, VT_I4, (void*)&result);
	return (E_TOOLTIPTYPE)result;
}

LONG CPagePrintEventArgs::GetPageNumber()
{
	LONG result;
	GetProperty(0x1, VT_I4, (void*)&result);
	return result;
}

CclsGDIGraphics CPagePrintEventArgs::GetGraphics()
{
	LPDISPATCH pDispatch;
	GetProperty(0x2, VT_DISPATCH, (void*)&pDispatch);
	return CclsGDIGraphics(pDispatch);
}

LONG CPagePrintEventArgs::GetX1()
{
	LONG result;
	GetProperty(0x3, VT_I4, (void*)&result);
	return result;
}

LONG CPagePrintEventArgs::GetY1()
{
	LONG result;
	GetProperty(0x4, VT_I4, (void*)&result);
	return result;
}

LONG CPagePrintEventArgs::GetX2()
{
	LONG result;
	GetProperty(0x5, VT_I4, (void*)&result);
	return result;
}

LONG CPagePrintEventArgs::GetY2()
{
	LONG result;
	GetProperty(0x6, VT_I4, (void*)&result);
	return result;
}

LONG CPagePrintEventArgs::GetPageWidth()
{
	LONG result;
	GetProperty(0x7, VT_I4, (void*)&result);
	return result;
}

LONG CPagePrintEventArgs::GetPageHeight()
{
	LONG result;
	GetProperty(0x8, VT_I4, (void*)&result);
	return result;
}

LONG CPagePrintEventArgs::GetActualPageWidth()
{
	LONG result;
	GetProperty(0x9, VT_I4, (void*)&result);
	return result;
}

LONG CPagePrintEventArgs::GetActualPageHeight()
{
	LONG result;
	GetProperty(0xa, VT_I4, (void*)&result);
	return result;
}

void CclsConfiguration::SetFirstDayOfWeek(E_WEEKDAY propval)
{
	LONG lpropval = (LONG)propval;
	SetProperty(0x1, VT_I4, lpropval);
}

E_WEEKDAY CclsConfiguration::GetFirstDayOfWeek()
{
	LONG result;
	GetProperty(0x1, VT_I4, (void*)&result);
	return (E_WEEKDAY)result;
}

void CclsConfiguration::SetCalendarWeekRule(E_CALENDARWEEKRULES propval)
{
	LONG lpropval = (LONG)propval;
	SetProperty(0x2, VT_I4, lpropval);
}

E_CALENDARWEEKRULES CclsConfiguration::GetCalendarWeekRule()
{
	LONG result;
	GetProperty(0x2, VT_I4, (void*)&result);
	return (E_CALENDARWEEKRULES)result;
}

void CclsConfiguration::SetISO8601Weeks(BOOL propval)
{
	SetProperty(0x3, VT_BOOL, propval);
}

BOOL CclsConfiguration::GetISO8601Weeks()
{
	BOOL result;
	GetProperty(0x3, VT_BOOL, (void*)&result);
	return result;
}

void CclsConfiguration::SetRowHighlightedBackColor(unsigned long propval)
{
	SetProperty(0x6, VT_COLOR, propval);
}

unsigned long CclsConfiguration::GetRowHighlightedBackColor()
{
	unsigned long result;
	GetProperty(0x6, VT_COLOR, (void*)&result);
	return result;
}

void CclsConfiguration::SetRowEvenBackColor(unsigned long propval)
{
	SetProperty(0x7, VT_COLOR, propval);
}

unsigned long CclsConfiguration::GetRowEvenBackColor()
{
	unsigned long result;
	GetProperty(0x7, VT_COLOR, (void*)&result);
	return result;
}

void CclsConfiguration::SetRowOddBackColor(unsigned long propval)
{
	SetProperty(0x8, VT_COLOR, propval);
}

unsigned long CclsConfiguration::GetRowOddBackColor()
{
	unsigned long result;
	GetProperty(0x8, VT_COLOR, (void*)&result);
	return result;
}

void CclsConfiguration::SetAlternatingRowStyles(BOOL propval)
{
	SetProperty(0x9, VT_BOOL, propval);
}

BOOL CclsConfiguration::GetAlternatingRowStyles()
{
	BOOL result;
	GetProperty(0x9, VT_BOOL, (void*)&result);
	return result;
}

CclsFont CclsConfiguration::GetDefaultFont()
{
	LPDISPATCH pDispatch;
	GetProperty(0xa, VT_DISPATCH, (void*)&pDispatch);
	return CclsFont(pDispatch);
}

CclsStyle CclsConfiguration::GetDefaultControlStyle()
{
	LPDISPATCH pDispatch;
	GetProperty(0xb, VT_DISPATCH, (void*)&pDispatch);
	return CclsStyle(pDispatch);
}

CclsStyle CclsConfiguration::GetDefaultTaskStyle()
{
	LPDISPATCH pDispatch;
	GetProperty(0xc, VT_DISPATCH, (void*)&pDispatch);
	return CclsStyle(pDispatch);
}

CclsStyle CclsConfiguration::GetDefaultRowStyle()
{
	LPDISPATCH pDispatch;
	GetProperty(0xd, VT_DISPATCH, (void*)&pDispatch);
	return CclsStyle(pDispatch);
}

CclsStyle CclsConfiguration::GetDefaultClientAreaStyle()
{
	LPDISPATCH pDispatch;
	GetProperty(0xe, VT_DISPATCH, (void*)&pDispatch);
	return CclsStyle(pDispatch);
}

CclsStyle CclsConfiguration::GetDefaultCellStyle()
{
	LPDISPATCH pDispatch;
	GetProperty(0xf, VT_DISPATCH, (void*)&pDispatch);
	return CclsStyle(pDispatch);
}

CclsStyle CclsConfiguration::GetDefaultColumnStyle()
{
	LPDISPATCH pDispatch;
	GetProperty(0x10, VT_DISPATCH, (void*)&pDispatch);
	return CclsStyle(pDispatch);
}

CclsStyle CclsConfiguration::GetDefaultPercentageStyle()
{
	LPDISPATCH pDispatch;
	GetProperty(0x11, VT_DISPATCH, (void*)&pDispatch);
	return CclsStyle(pDispatch);
}

CclsStyle CclsConfiguration::GetDefaultPredecessorStyle()
{
	LPDISPATCH pDispatch;
	GetProperty(0x12, VT_DISPATCH, (void*)&pDispatch);
	return CclsStyle(pDispatch);
}

CclsStyle CclsConfiguration::GetDefaultTimeLineStyle()
{
	LPDISPATCH pDispatch;
	GetProperty(0x13, VT_DISPATCH, (void*)&pDispatch);
	return CclsStyle(pDispatch);
}

CclsStyle CclsConfiguration::GetDefaultTimeBlockStyle()
{
	LPDISPATCH pDispatch;
	GetProperty(0x14, VT_DISPATCH, (void*)&pDispatch);
	return CclsStyle(pDispatch);
}

CclsStyle CclsConfiguration::GetDefaultTickMarkAreaStyle()
{
	LPDISPATCH pDispatch;
	GetProperty(0x15, VT_DISPATCH, (void*)&pDispatch);
	return CclsStyle(pDispatch);
}

CclsStyle CclsConfiguration::GetDefaultSplitterStyle()
{
	LPDISPATCH pDispatch;
	GetProperty(0x16, VT_DISPATCH, (void*)&pDispatch);
	return CclsStyle(pDispatch);
}

CclsStyle CclsConfiguration::GetDefaultProgressLineStyle()
{
	LPDISPATCH pDispatch;
	GetProperty(0x17, VT_DISPATCH, (void*)&pDispatch);
	return CclsStyle(pDispatch);
}

CclsStyle CclsConfiguration::GetDefaultNodeStyle()
{
	LPDISPATCH pDispatch;
	GetProperty(0x18, VT_DISPATCH, (void*)&pDispatch);
	return CclsStyle(pDispatch);
}

CclsStyle CclsConfiguration::GetDefaultTierStyle()
{
	LPDISPATCH pDispatch;
	GetProperty(0x19, VT_DISPATCH, (void*)&pDispatch);
	return CclsStyle(pDispatch);
}

CclsStyle CclsConfiguration::GetDefaultScrollBarStyle()
{
	LPDISPATCH pDispatch;
	GetProperty(0x1a, VT_DISPATCH, (void*)&pDispatch);
	return CclsStyle(pDispatch);
}

CclsStyle CclsConfiguration::GetDefaultSBSeparatorStyle()
{
	LPDISPATCH pDispatch;
	GetProperty(0x1b, VT_DISPATCH, (void*)&pDispatch);
	return CclsStyle(pDispatch);
}

CclsStyle CclsConfiguration::GetDefaultSBNormalStyle()
{
	LPDISPATCH pDispatch;
	GetProperty(0x1c, VT_DISPATCH, (void*)&pDispatch);
	return CclsStyle(pDispatch);
}

CclsStyle CclsConfiguration::GetDefaultSBPressedStyle()
{
	LPDISPATCH pDispatch;
	GetProperty(0x1d, VT_DISPATCH, (void*)&pDispatch);
	return CclsStyle(pDispatch);
}

CclsStyle CclsConfiguration::GetDefaultSBDisabledStyle()
{
	LPDISPATCH pDispatch;
	GetProperty(0x1e, VT_DISPATCH, (void*)&pDispatch);
	return CclsStyle(pDispatch);
}

CString CclsConfiguration::GetXML(void)
{
	CString oReturn;
	InvokeHelper(0x4, DISPATCH_METHOD, VT_BSTR, (void*)&oReturn, NULL);
	return oReturn;
}

void CclsConfiguration::SetXML(LPCTSTR sXML)
{
	static BYTE params[] = VTS_BSTR;
	InvokeHelper(0x5, DISPATCH_METHOD, VT_EMPTY, NULL, params, sXML);
}

void CControlTemplateGradient::SetControlBorderColor(unsigned long propval)
{
	SetProperty(0x1, VT_COLOR, propval);
}

unsigned long CControlTemplateGradient::GetControlBorderColor()
{
	unsigned long result;
	GetProperty(0x1, VT_COLOR, (void*)&result);
	return result;
}

void CControlTemplateGradient::SetHeadingBorderColor(unsigned long propval)
{
	SetProperty(0x2, VT_COLOR, propval);
}

unsigned long CControlTemplateGradient::GetHeadingBorderColor()
{
	unsigned long result;
	GetProperty(0x2, VT_COLOR, (void*)&result);
	return result;
}

void CControlTemplateGradient::SetHeadingForeColor(unsigned long propval)
{
	SetProperty(0x3, VT_COLOR, propval);
}

unsigned long CControlTemplateGradient::GetHeadingForeColor()
{
	unsigned long result;
	GetProperty(0x3, VT_COLOR, (void*)&result);
	return result;
}

void CControlTemplateGradient::SetRowForeColor(unsigned long propval)
{
	SetProperty(0x4, VT_COLOR, propval);
}

unsigned long CControlTemplateGradient::GetRowForeColor()
{
	unsigned long result;
	GetProperty(0x4, VT_COLOR, (void*)&result);
	return result;
}

void CControlTemplateGradient::SetGradientFillMode(GRE_GRADIENTFILLMODE propval)
{
	LONG lpropval = (LONG)propval;
	SetProperty(0x5, VT_I4, lpropval);
}

GRE_GRADIENTFILLMODE CControlTemplateGradient::GetGradientFillMode()
{
	LONG result;
	GetProperty(0x5, VT_I4, (void*)&result);
	return (GRE_GRADIENTFILLMODE)result;
}

void CControlTemplateGradient::SetStartGradientColor(unsigned long propval)
{
	SetProperty(0x6, VT_COLOR, propval);
}

unsigned long CControlTemplateGradient::GetStartGradientColor()
{
	unsigned long result;
	GetProperty(0x6, VT_COLOR, (void*)&result);
	return result;
}

void CControlTemplateGradient::SetEndGradientColor(unsigned long propval)
{
	SetProperty(0x7, VT_COLOR, propval);
}

unsigned long CControlTemplateGradient::GetEndGradientColor()
{
	unsigned long result;
	GetProperty(0x7, VT_COLOR, (void*)&result);
	return result;
}

void CControlTemplateGradient::SetTreeviewColor(unsigned long propval)
{
	SetProperty(0x8, VT_COLOR, propval);
}

unsigned long CControlTemplateGradient::GetTreeviewColor()
{
	unsigned long result;
	GetProperty(0x8, VT_COLOR, (void*)&result);
	return result;
}

void CControlTemplateGradient::SetRowBorderColor(unsigned long propval)
{
	SetProperty(0x9, VT_COLOR, propval);
}

unsigned long CControlTemplateGradient::GetRowBorderColor()
{
	unsigned long result;
	GetProperty(0x9, VT_COLOR, (void*)&result);
	return result;
}

void CControlTemplateSolid::SetControlBorderColor(unsigned long propval)
{
	SetProperty(0x1, VT_COLOR, propval);
}

unsigned long CControlTemplateSolid::GetControlBorderColor()
{
	unsigned long result;
	GetProperty(0x1, VT_COLOR, (void*)&result);
	return result;
}

void CControlTemplateSolid::SetHeadingBorderColor(unsigned long propval)
{
	SetProperty(0x2, VT_COLOR, propval);
}

unsigned long CControlTemplateSolid::GetHeadingBorderColor()
{
	unsigned long result;
	GetProperty(0x2, VT_COLOR, (void*)&result);
	return result;
}

void CControlTemplateSolid::SetHeadingForeColor(unsigned long propval)
{
	SetProperty(0x3, VT_COLOR, propval);
}

unsigned long CControlTemplateSolid::GetHeadingForeColor()
{
	unsigned long result;
	GetProperty(0x3, VT_COLOR, (void*)&result);
	return result;
}

void CControlTemplateSolid::SetRowForeColor(unsigned long propval)
{
	SetProperty(0x4, VT_COLOR, propval);
}

unsigned long CControlTemplateSolid::GetRowForeColor()
{
	unsigned long result;
	GetProperty(0x4, VT_COLOR, (void*)&result);
	return result;
}

void CControlTemplateSolid::SetHeadingBackColor(unsigned long propval)
{
	SetProperty(0x5, VT_COLOR, propval);
}

unsigned long CControlTemplateSolid::GetHeadingBackColor()
{
	unsigned long result;
	GetProperty(0x5, VT_COLOR, (void*)&result);
	return result;
}

void CControlTemplateSolid::SetTreeviewColor(unsigned long propval)
{
	SetProperty(0x6, VT_COLOR, propval);
}

unsigned long CControlTemplateSolid::GetTreeviewColor()
{
	unsigned long result;
	GetProperty(0x6, VT_COLOR, (void*)&result);
	return result;
}

void CControlTemplateSolid::SetRowBorderColor(unsigned long propval)
{
	SetProperty(0x7, VT_COLOR, propval);
}

unsigned long CControlTemplateSolid::GetRowBorderColor()
{
	unsigned long result;
	GetProperty(0x7, VT_COLOR, (void*)&result);
	return result;
}

void CControlTemplateSolid::SetTimeBlockBackColor(unsigned long propval)
{
	SetProperty(0x8, VT_COLOR, propval);
}

unsigned long CControlTemplateSolid::GetTimeBlockBackColor()
{
	unsigned long result;
	GetProperty(0x8, VT_COLOR, (void*)&result);
	return result;
}

void CActiveGanttVCCtl::SetErrorReports(E_REPORTERRORS propval)
{
	LONG lpropval = (LONG)propval;
	SetProperty(0x1, VT_I4, lpropval);
}

E_REPORTERRORS CActiveGanttVCCtl::GetErrorReports()
{
	LONG result;
	GetProperty(0x1, VT_I4, (void*)&result);
	return (E_REPORTERRORS)result;
}

void CActiveGanttVCCtl::SetCurrentLayer(LPCTSTR propval)
{
	SetProperty(0x2, VT_BSTR, propval);
}

CString CActiveGanttVCCtl::GetCurrentLayer()
{
	CString result;
	GetProperty(0x2, VT_BSTR, (void*)&result);
	return result;
}

void CActiveGanttVCCtl::SetCurrentView(LPCTSTR propval)
{
	SetProperty(0x3, VT_BSTR, propval);
}

CString CActiveGanttVCCtl::GetCurrentView()
{
	CString result;
	GetProperty(0x3, VT_BSTR, (void*)&result);
	return result;
}

CclsView CActiveGanttVCCtl::GetCurrentViewObject()
{
	LPDISPATCH pDispatch;
	GetProperty(0x4, VT_DISPATCH, (void*)&pDispatch);
	return CclsView(pDispatch);
}

CclsToolTip CActiveGanttVCCtl::GetToolTip()
{
	LPDISPATCH pDispatch;
	GetProperty(0x5, VT_DISPATCH, (void*)&pDispatch);
	return CclsToolTip(pDispatch);
}

void CActiveGanttVCCtl::SetScrollBarBehaviour(E_SCROLLBEHAVIOUR propval)
{
	LONG lpropval = (LONG)propval;
	SetProperty(0x6, VT_I4, lpropval);
}

E_SCROLLBEHAVIOUR CActiveGanttVCCtl::GetScrollBarBehaviour()
{
	LONG result;
	GetProperty(0x6, VT_I4, (void*)&result);
	return (E_SCROLLBEHAVIOUR)result;
}

void CActiveGanttVCCtl::SetTimeBlockBehaviour(E_TIMEBLOCKBEHAVIOUR propval)
{
	LONG lpropval = (LONG)propval;
	SetProperty(0x7, VT_I4, lpropval);
}

E_TIMEBLOCKBEHAVIOUR CActiveGanttVCCtl::GetTimeBlockBehaviour()
{
	LONG result;
	GetProperty(0x7, VT_I4, (void*)&result);
	return (E_TIMEBLOCKBEHAVIOUR)result;
}

void CActiveGanttVCCtl::SetSelectedTaskIndex(LONG propval)
{
	SetProperty(0x8, VT_I4, propval);
}

LONG CActiveGanttVCCtl::GetSelectedTaskIndex()
{
	LONG result;
	GetProperty(0x8, VT_I4, (void*)&result);
	return result;
}

void CActiveGanttVCCtl::SetSelectedColumnIndex(LONG propval)
{
	SetProperty(0x9, VT_I4, propval);
}

LONG CActiveGanttVCCtl::GetSelectedColumnIndex()
{
	LONG result;
	GetProperty(0x9, VT_I4, (void*)&result);
	return result;
}

void CActiveGanttVCCtl::SetSelectedRowIndex(LONG propval)
{
	SetProperty(0xa, VT_I4, propval);
}

LONG CActiveGanttVCCtl::GetSelectedRowIndex()
{
	LONG result;
	GetProperty(0xa, VT_I4, (void*)&result);
	return result;
}

void CActiveGanttVCCtl::SetSelectedCellIndex(LONG propval)
{
	SetProperty(0xb, VT_I4, propval);
}

LONG CActiveGanttVCCtl::GetSelectedCellIndex()
{
	LONG result;
	GetProperty(0xb, VT_I4, (void*)&result);
	return result;
}

void CActiveGanttVCCtl::SetSelectedPercentageIndex(LONG propval)
{
	SetProperty(0xc, VT_I4, propval);
}

LONG CActiveGanttVCCtl::GetSelectedPercentageIndex()
{
	LONG result;
	GetProperty(0xc, VT_I4, (void*)&result);
	return result;
}

void CActiveGanttVCCtl::SetTreeviewColumnIndex(LONG propval)
{
	SetProperty(0xd, VT_I4, propval);
}

LONG CActiveGanttVCCtl::GetTreeviewColumnIndex()
{
	LONG result;
	GetProperty(0xd, VT_I4, (void*)&result);
	return result;
}

CclsCultureInfo CActiveGanttVCCtl::GetCulture()
{
	LPDISPATCH pDispatch;
	GetProperty(0xe, VT_DISPATCH, (void*)&pDispatch);
	return CclsCultureInfo(pDispatch);
}

CString CActiveGanttVCCtl::GetModuleCompletePath()
{
	CString result;
	GetProperty(0xf, VT_BSTR, (void*)&result);
	return result;
}

CString CActiveGanttVCCtl::GetVersion()
{
	CString result;
	GetProperty(0x10, VT_BSTR, (void*)&result);
	return result;
}

void CActiveGanttVCCtl::SetAllowSplitterMove(BOOL propval)
{
	SetProperty(0x11, VT_BOOL, propval);
}

BOOL CActiveGanttVCCtl::GetAllowSplitterMove()
{
	BOOL result;
	GetProperty(0x11, VT_BOOL, (void*)&result);
	return result;
}

void CActiveGanttVCCtl::SetAllowPredecessorAdd(BOOL propval)
{
	SetProperty(0x12, VT_BOOL, propval);
}

BOOL CActiveGanttVCCtl::GetAllowPredecessorAdd()
{
	BOOL result;
	GetProperty(0x12, VT_BOOL, (void*)&result);
	return result;
}

void CActiveGanttVCCtl::SetAllowAdd(BOOL propval)
{
	SetProperty(0x13, VT_BOOL, propval);
}

BOOL CActiveGanttVCCtl::GetAllowAdd()
{
	BOOL result;
	GetProperty(0x13, VT_BOOL, (void*)&result);
	return result;
}

void CActiveGanttVCCtl::SetAllowEdit(BOOL propval)
{
	SetProperty(0x14, VT_BOOL, propval);
}

BOOL CActiveGanttVCCtl::GetAllowEdit()
{
	BOOL result;
	GetProperty(0x14, VT_BOOL, (void*)&result);
	return result;
}

void CActiveGanttVCCtl::SetAllowColumnSize(BOOL propval)
{
	SetProperty(0x15, VT_BOOL, propval);
}

BOOL CActiveGanttVCCtl::GetAllowColumnSize()
{
	BOOL result;
	GetProperty(0x15, VT_BOOL, (void*)&result);
	return result;
}

void CActiveGanttVCCtl::SetAllowColumnMove(BOOL propval)
{
	SetProperty(0x16, VT_BOOL, propval);
}

BOOL CActiveGanttVCCtl::GetAllowColumnMove()
{
	BOOL result;
	GetProperty(0x16, VT_BOOL, (void*)&result);
	return result;
}

void CActiveGanttVCCtl::SetAllowRowSize(BOOL propval)
{
	SetProperty(0x17, VT_BOOL, propval);
}

BOOL CActiveGanttVCCtl::GetAllowRowSize()
{
	BOOL result;
	GetProperty(0x17, VT_BOOL, (void*)&result);
	return result;
}

void CActiveGanttVCCtl::SetAllowRowMove(BOOL propval)
{
	SetProperty(0x18, VT_BOOL, propval);
}

BOOL CActiveGanttVCCtl::GetAllowRowMove()
{
	BOOL result;
	GetProperty(0x18, VT_BOOL, (void*)&result);
	return result;
}

void CActiveGanttVCCtl::SetAllowTimeLineScroll(BOOL propval)
{
	SetProperty(0x19, VT_BOOL, propval);
}

BOOL CActiveGanttVCCtl::GetAllowTimeLineScroll()
{
	BOOL result;
	GetProperty(0x19, VT_BOOL, (void*)&result);
	return result;
}

void CActiveGanttVCCtl::SetAddMode(E_ADDMODE propval)
{
	LONG lpropval = (LONG)propval;
	SetProperty(0x1a, VT_I4, lpropval);
}

E_ADDMODE CActiveGanttVCCtl::GetAddMode()
{
	LONG result;
	GetProperty(0x1a, VT_I4, (void*)&result);
	return (E_ADDMODE)result;
}

void CActiveGanttVCCtl::SetLayerEnableObjects(E_LAYEROBJECTENABLE propval)
{
	LONG lpropval = (LONG)propval;
	SetProperty(0x1b, VT_I4, lpropval);
}

E_LAYEROBJECTENABLE CActiveGanttVCCtl::GetLayerEnableObjects()
{
	LONG result;
	GetProperty(0x1b, VT_I4, (void*)&result);
	return (E_LAYEROBJECTENABLE)result;
}

void CActiveGanttVCCtl::SetDoubleBuffering(BOOL propval)
{
	SetProperty(0x1c, VT_BOOL, propval);
}

BOOL CActiveGanttVCCtl::GetDoubleBuffering()
{
	BOOL result;
	GetProperty(0x1c, VT_BOOL, (void*)&result);
	return result;
}

void CActiveGanttVCCtl::SetMinRowHeight(LONG propval)
{
	SetProperty(0x1d, VT_I4, propval);
}

LONG CActiveGanttVCCtl::GetMinRowHeight()
{
	LONG result;
	GetProperty(0x1d, VT_I4, (void*)&result);
	return result;
}

void CActiveGanttVCCtl::SetMinColumnWidth(LONG propval)
{
	SetProperty(0x1e, VT_I4, propval);
}

LONG CActiveGanttVCCtl::GetMinColumnWidth()
{
	LONG result;
	GetProperty(0x1e, VT_I4, (void*)&result);
	return result;
}

CclsRows CActiveGanttVCCtl::GetRows()
{
	LPDISPATCH pDispatch;
	GetProperty(0x1f, VT_DISPATCH, (void*)&pDispatch);
	return CclsRows(pDispatch);
}

CclsTasks CActiveGanttVCCtl::GetTasks()
{
	LPDISPATCH pDispatch;
	GetProperty(0x20, VT_DISPATCH, (void*)&pDispatch);
	return CclsTasks(pDispatch);
}

CclsColumns CActiveGanttVCCtl::GetColumns()
{
	LPDISPATCH pDispatch;
	GetProperty(0x21, VT_DISPATCH, (void*)&pDispatch);
	return CclsColumns(pDispatch);
}

CclsStyles CActiveGanttVCCtl::GetStyles()
{
	LPDISPATCH pDispatch;
	GetProperty(0x22, VT_DISPATCH, (void*)&pDispatch);
	return CclsStyles(pDispatch);
}

CclsLayers CActiveGanttVCCtl::GetLayers()
{
	LPDISPATCH pDispatch;
	GetProperty(0x23, VT_DISPATCH, (void*)&pDispatch);
	return CclsLayers(pDispatch);
}

CclsPercentages CActiveGanttVCCtl::GetPercentages()
{
	LPDISPATCH pDispatch;
	GetProperty(0x24, VT_DISPATCH, (void*)&pDispatch);
	return CclsPercentages(pDispatch);
}

CclsTimeBlocks CActiveGanttVCCtl::GetTimeBlocks()
{
	LPDISPATCH pDispatch;
	GetProperty(0x25, VT_DISPATCH, (void*)&pDispatch);
	return CclsTimeBlocks(pDispatch);
}

CclsViews CActiveGanttVCCtl::GetViews()
{
	LPDISPATCH pDispatch;
	GetProperty(0x26, VT_DISPATCH, (void*)&pDispatch);
	return CclsViews(pDispatch);
}

CclsSplitter CActiveGanttVCCtl::GetSplitter()
{
	LPDISPATCH pDispatch;
	GetProperty(0x27, VT_DISPATCH, (void*)&pDispatch);
	return CclsSplitter(pDispatch);
}

CclsPrinter CActiveGanttVCCtl::GetPrinter()
{
	LPDISPATCH pDispatch;
	GetProperty(0x28, VT_DISPATCH, (void*)&pDispatch);
	return CclsPrinter(pDispatch);
}

CclsTreeview CActiveGanttVCCtl::GetTreeview()
{
	LPDISPATCH pDispatch;
	GetProperty(0x29, VT_DISPATCH, (void*)&pDispatch);
	return CclsTreeview(pDispatch);
}

CclsDrawing CActiveGanttVCCtl::GetDrawing()
{
	LPDISPATCH pDispatch;
	GetProperty(0x2a, VT_DISPATCH, (void*)&pDispatch);
	return CclsDrawing(pDispatch);
}

CclsMath CActiveGanttVCCtl::GetMathLib()
{
	LPDISPATCH pDispatch;
	GetProperty(0x2b, VT_DISPATCH, (void*)&pDispatch);
	return CclsMath(pDispatch);
}

CclsVerticalScrollBar CActiveGanttVCCtl::GetVerticalScrollBar()
{
	LPDISPATCH pDispatch;
	GetProperty(0x2c, VT_DISPATCH, (void*)&pDispatch);
	return CclsVerticalScrollBar(pDispatch);
}

CclsHorizontalScrollBar CActiveGanttVCCtl::GetHorizontalScrollBar()
{
	LPDISPATCH pDispatch;
	GetProperty(0x2d, VT_DISPATCH, (void*)&pDispatch);
	return CclsHorizontalScrollBar(pDispatch);
}

CToolTipEventArgs CActiveGanttVCCtl::GetToolTipEventArgs()
{
	LPDISPATCH pDispatch;
	GetProperty(0x2e, VT_DISPATCH, (void*)&pDispatch);
	return CToolTipEventArgs(pDispatch);
}

CObjectAddedEventArgs CActiveGanttVCCtl::GetObjectAddedEventArgs()
{
	LPDISPATCH pDispatch;
	GetProperty(0x2f, VT_DISPATCH, (void*)&pDispatch);
	return CObjectAddedEventArgs(pDispatch);
}

CCustomTierDrawEventArgs CActiveGanttVCCtl::GetCustomTierDrawEventArgs()
{
	LPDISPATCH pDispatch;
	GetProperty(0x30, VT_DISPATCH, (void*)&pDispatch);
	return CCustomTierDrawEventArgs(pDispatch);
}

CMouseEventArgs CActiveGanttVCCtl::GetMouseEventArgs()
{
	LPDISPATCH pDispatch;
	GetProperty(0x31, VT_DISPATCH, (void*)&pDispatch);
	return CMouseEventArgs(pDispatch);
}

CKeyEventArgs CActiveGanttVCCtl::GetKeyEventArgs()
{
	LPDISPATCH pDispatch;
	GetProperty(0x32, VT_DISPATCH, (void*)&pDispatch);
	return CKeyEventArgs(pDispatch);
}

CScrollEventArgs CActiveGanttVCCtl::GetScrollEventArgs()
{
	LPDISPATCH pDispatch;
	GetProperty(0x33, VT_DISPATCH, (void*)&pDispatch);
	return CScrollEventArgs(pDispatch);
}

CDrawEventArgs CActiveGanttVCCtl::GetDrawEventArgs()
{
	LPDISPATCH pDispatch;
	GetProperty(0x34, VT_DISPATCH, (void*)&pDispatch);
	return CDrawEventArgs(pDispatch);
}

CPredecessorDrawEventArgs CActiveGanttVCCtl::GetPredecessorDrawEventArgs()
{
	LPDISPATCH pDispatch;
	GetProperty(0x35, VT_DISPATCH, (void*)&pDispatch);
	return CPredecessorDrawEventArgs(pDispatch);
}

CObjectSelectedEventArgs CActiveGanttVCCtl::GetObjectSelectedEventArgs()
{
	LPDISPATCH pDispatch;
	GetProperty(0x36, VT_DISPATCH, (void*)&pDispatch);
	return CObjectSelectedEventArgs(pDispatch);
}

CObjectStateChangedEventArgs CActiveGanttVCCtl::GetObjectStateChangedEventArgs()
{
	LPDISPATCH pDispatch;
	GetProperty(0x37, VT_DISPATCH, (void*)&pDispatch);
	return CObjectStateChangedEventArgs(pDispatch);
}

CErrorEventArgs CActiveGanttVCCtl::GetErrorEventArgs()
{
	LPDISPATCH pDispatch;
	GetProperty(0x38, VT_DISPATCH, (void*)&pDispatch);
	return CErrorEventArgs(pDispatch);
}

CNodeEventArgs CActiveGanttVCCtl::GetNodeEventArgs()
{
	LPDISPATCH pDispatch;
	GetProperty(0x39, VT_DISPATCH, (void*)&pDispatch);
	return CNodeEventArgs(pDispatch);
}

void CActiveGanttVCCtl::SetControlTag(LPCTSTR propval)
{
	SetProperty(0x3a, VT_BSTR, propval);
}

CString CActiveGanttVCCtl::GetControlTag()
{
	CString result;
	GetProperty(0x3a, VT_BSTR, (void*)&result);
	return result;
}

LPDISPATCH CActiveGanttVCCtl::GetGetIDispatch()
{
	LPDISPATCH result;
	GetProperty(0x3b, VT_DISPATCH, (void*)&result);
	return result;
}

CMouseWheelEventArgs CActiveGanttVCCtl::GetMouseWheelEventArgs()
{
	LPDISPATCH pDispatch;
	GetProperty(0x3c, VT_DISPATCH, (void*)&pDispatch);
	return CMouseWheelEventArgs(pDispatch);
}

CclsTierAppearance CActiveGanttVCCtl::GetTierAppearance()
{
	LPDISPATCH pDispatch;
	GetProperty(0x3d, VT_DISPATCH, (void*)&pDispatch);
	return CclsTierAppearance(pDispatch);
}

CclsTierFormat CActiveGanttVCCtl::GetTierFormat()
{
	LPDISPATCH pDispatch;
	GetProperty(0x3e, VT_DISPATCH, (void*)&pDispatch);
	return CclsTierFormat(pDispatch);
}

void CActiveGanttVCCtl::SetTierAppearanceScope(E_OBJECTSCOPE propval)
{
	LONG lpropval = (LONG)propval;
	SetProperty(0x3f, VT_I4, lpropval);
}

E_OBJECTSCOPE CActiveGanttVCCtl::GetTierAppearanceScope()
{
	LONG result;
	GetProperty(0x3f, VT_I4, (void*)&result);
	return (E_OBJECTSCOPE)result;
}

void CActiveGanttVCCtl::SetTierFormatScope(E_OBJECTSCOPE propval)
{
	LONG lpropval = (LONG)propval;
	SetProperty(0x40, VT_I4, lpropval);
}

E_OBJECTSCOPE CActiveGanttVCCtl::GetTierFormatScope()
{
	LONG result;
	GetProperty(0x40, VT_I4, (void*)&result);
	return (E_OBJECTSCOPE)result;
}

void CActiveGanttVCCtl::SetStyleIndex(LPCTSTR propval)
{
	SetProperty(0x41, VT_BSTR, propval);
}

CString CActiveGanttVCCtl::GetStyleIndex()
{
	CString result;
	GetProperty(0x41, VT_BSTR, (void*)&result);
	return result;
}

CclsStyle CActiveGanttVCCtl::GetStyle()
{
	LPDISPATCH pDispatch;
	GetProperty(0x42, VT_DISPATCH, (void*)&pDispatch);
	return CclsStyle(pDispatch);
}

CclsImage CActiveGanttVCCtl::GetImage()
{
	LPDISPATCH pDispatch;
	GetProperty(0x43, VT_DISPATCH, (void*)&pDispatch);
	return CclsImage(pDispatch);
}

void CActiveGanttVCCtl::SetImageTag(LPCTSTR propval)
{
	SetProperty(0x44, VT_BSTR, propval);
}

CString CActiveGanttVCCtl::GetImageTag()
{
	CString result;
	GetProperty(0x44, VT_BSTR, (void*)&result);
	return result;
}

CclsScrollBarSeparator CActiveGanttVCCtl::GetScrollBarSeparator()
{
	LPDISPATCH pDispatch;
	GetProperty(0x45, VT_DISPATCH, (void*)&pDispatch);
	return CclsScrollBarSeparator(pDispatch);
}

CTextEditEventArgs CActiveGanttVCCtl::GetTextEditEventArgs()
{
	LPDISPATCH pDispatch;
	GetProperty(0x46, VT_DISPATCH, (void*)&pDispatch);
	return CTextEditEventArgs(pDispatch);
}

CclsPredecessors CActiveGanttVCCtl::GetPredecessors()
{
	LPDISPATCH pDispatch;
	GetProperty(0x47, VT_DISPATCH, (void*)&pDispatch);
	return CclsPredecessors(pDispatch);
}

CPredecessorExceptionEventArgs CActiveGanttVCCtl::GetPredecessorExceptionEventArgs()
{
	LPDISPATCH pDispatch;
	GetProperty(0x48, VT_DISPATCH, (void*)&pDispatch);
	return CPredecessorExceptionEventArgs(pDispatch);
}

void CActiveGanttVCCtl::SetSelectedPredecessorIndex(LONG propval)
{
	SetProperty(0x49, VT_I4, propval);
}

LONG CActiveGanttVCCtl::GetSelectedPredecessorIndex()
{
	LONG result;
	GetProperty(0x49, VT_I4, (void*)&result);
	return result;
}

void CActiveGanttVCCtl::SetEnforcePredecessors(BOOL propval)
{
	SetProperty(0x4a, VT_BOOL, propval);
}

BOOL CActiveGanttVCCtl::GetEnforcePredecessors()
{
	BOOL result;
	GetProperty(0x4a, VT_BOOL, (void*)&result);
	return result;
}

void CActiveGanttVCCtl::SetPredecessorMode(E_PREDECESSORMODE propval)
{
	LONG lpropval = (LONG)propval;
	SetProperty(0x4b, VT_I4, lpropval);
}

E_PREDECESSORMODE CActiveGanttVCCtl::GetPredecessorMode()
{
	LONG result;
	GetProperty(0x4b, VT_I4, (void*)&result);
	return (E_PREDECESSORMODE)result;
}

void CActiveGanttVCCtl::SetAddDurationInterval(E_INTERVAL propval)
{
	LONG lpropval = (LONG)propval;
	SetProperty(0x4c, VT_I4, lpropval);
}

E_INTERVAL CActiveGanttVCCtl::GetAddDurationInterval()
{
	LONG result;
	GetProperty(0x4c, VT_I4, (void*)&result);
	return (E_INTERVAL)result;
}

CclsTimeLineScrollBar CActiveGanttVCCtl::GetTimeLineScrollBar()
{
	LPDISPATCH pDispatch;
	GetProperty(0x58, VT_DISPATCH, (void*)&pDispatch);
	return CclsTimeLineScrollBar(pDispatch);
}

CclsProgressLine CActiveGanttVCCtl::GetProgressLine()
{
	LPDISPATCH pDispatch;
	GetProperty(0x59, VT_DISPATCH, (void*)&pDispatch);
	return CclsProgressLine(pDispatch);
}

void CActiveGanttVCCtl::SetTimeLineScrollBarScope(E_OBJECTSCOPE propval)
{
	LONG lpropval = (LONG)propval;
	SetProperty(0x5a, VT_I4, lpropval);
}

E_OBJECTSCOPE CActiveGanttVCCtl::GetTimeLineScrollBarScope()
{
	LONG result;
	GetProperty(0x5a, VT_I4, (void*)&result);
	return (E_OBJECTSCOPE)result;
}

void CActiveGanttVCCtl::SetProgressLineScope(E_OBJECTSCOPE propval)
{
	LONG lpropval = (LONG)propval;
	SetProperty(0x5b, VT_I4, lpropval);
}

E_OBJECTSCOPE CActiveGanttVCCtl::GetProgressLineScope()
{
	LONG result;
	GetProperty(0x5b, VT_I4, (void*)&result);
	return (E_OBJECTSCOPE)result;
}

CclsConfiguration CActiveGanttVCCtl::GetConfiguration()
{
	LPDISPATCH pDispatch;
	GetProperty(0x5c, VT_DISPATCH, (void*)&pDispatch);
	return CclsConfiguration(pDispatch);
}

E_CONTROLTEMPLATE CActiveGanttVCCtl::GetControlTemplate()
{
	LONG result;
	GetProperty(0x5f, VT_I4, (void*)&result);
	return (E_CONTROLTEMPLATE)result;
}

E_OBJECTTEMPLATE CActiveGanttVCCtl::GetObjectTemplate()
{
	LONG result;
	GetProperty(0x60, VT_I4, (void*)&result);
	return (E_OBJECTTEMPLATE)result;
}

CString CActiveGanttVCCtl::GetPrinterErrorMessage()
{
	CString result;
	GetProperty(0x62, VT_BSTR, (void*)&result);
	return result;
}

void CActiveGanttVCCtl::WriteXML(LPCTSTR url)
{
	static BYTE params[] = VTS_BSTR;
	InvokeHelper(0x4d, DISPATCH_METHOD, VT_EMPTY, NULL, params, url);
}

void CActiveGanttVCCtl::ReadXML(LPCTSTR url)
{
	static BYTE params[] = VTS_BSTR;
	InvokeHelper(0x4e, DISPATCH_METHOD, VT_EMPTY, NULL, params, url);
}

void CActiveGanttVCCtl::SetXML(LPCTSTR sXML)
{
	static BYTE params[] = VTS_BSTR;
	InvokeHelper(0x4f, DISPATCH_METHOD, VT_EMPTY, NULL, params, sXML);
}

CString CActiveGanttVCCtl::GetXML(void)
{
	CString oReturn;
	InvokeHelper(0x50, DISPATCH_METHOD, VT_BSTR, (void*)&oReturn, NULL);
	return oReturn;
}

void CActiveGanttVCCtl::ClearSelections(void)
{
	InvokeHelper(0x51, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
}

void CActiveGanttVCCtl::Clear(void)
{
	InvokeHelper(0x52, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
}

void CActiveGanttVCCtl::SaveToImage(LPCTSTR Path, GRE_IMAGEFORMAT Format)
{
	static BYTE params[] = VTS_BSTR VTS_I4;
	InvokeHelper(0x53, DISPATCH_METHOD, VT_EMPTY, NULL, params, Path, Format);
}

void CActiveGanttVCCtl::AboutBox(void)
{
	InvokeHelper(0x54, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
}

void CActiveGanttVCCtl::Redraw(void)
{
	InvokeHelper(0x55, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
}

void CActiveGanttVCCtl::ForceEndTextEdit(void)
{
	InvokeHelper(0x56, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
}

void CActiveGanttVCCtl::CheckPredecessors(void)
{
	InvokeHelper(0x57, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
}

void CActiveGanttVCCtl::ApplyTemplateSolid(LPDISPATCH ControlTemplate, E_OBJECTTEMPLATE ObjectTemplate)
{
	static BYTE params[] = VTS_DISPATCH VTS_I4;
	InvokeHelper(0x5d, DISPATCH_METHOD, VT_EMPTY, NULL, params, ControlTemplate, ObjectTemplate);
}

void CActiveGanttVCCtl::ApplyTemplateGradient(LPDISPATCH ControlTemplate, E_OBJECTTEMPLATE ObjectTemplate)
{
	static BYTE params[] = VTS_DISPATCH VTS_I4;
	InvokeHelper(0x5e, DISPATCH_METHOD, VT_EMPTY, NULL, params, ControlTemplate, ObjectTemplate);
}

void CActiveGanttVCCtl::ApplyTemplate(E_CONTROLTEMPLATE ControlTemplate, E_OBJECTTEMPLATE ObjectTemplate)
{
	static BYTE params[] = VTS_I4 VTS_I4;
	InvokeHelper(0x61, DISPATCH_METHOD, VT_EMPTY, NULL, params, ControlTemplate, ObjectTemplate);
}
}