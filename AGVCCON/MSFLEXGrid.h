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

class COleFont : public COleDispatchDriver
{
public:
	COleFont() {}		// Calls COleDispatchDriver default constructor
	COleFont(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	COleFont(const COleFont& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attributes
public:
	CString GetName();
	void SetName(LPCTSTR);
	CY GetSize();
	void SetSize(const CY&);
	BOOL GetBold();
	void SetBold(BOOL);
	BOOL GetItalic();
	void SetItalic(BOOL);
	BOOL GetUnderline();
	void SetUnderline(BOOL);
	BOOL GetStrikethrough();
	void SetStrikethrough(BOOL);
	short GetWeight();
	void SetWeight(short);
	short GetCharset();
	void SetCharset(short);

// Operations
public:
};

class CPicture : public COleDispatchDriver
{
public:
	CPicture() {}		// Calls COleDispatchDriver default constructor
	CPicture(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CPicture(const CPicture& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attributes
public:
	long GetHandle();
	long GetHPal();
	void SetHPal(long);
	short GetType();
	long GetWidth();
	long GetHeight();

// Operations
public:
	// method 'Render' not emitted because of invalid return type or parameter type
};

class CRowCursor : public COleDispatchDriver
{
public:
	CRowCursor() {}		// Calls COleDispatchDriver default constructor
	CRowCursor(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CRowCursor(const CRowCursor& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attributes
public:

// Operations
public:
};

class CMSFlexGrid : public CWnd
{
protected:
	DECLARE_DYNCREATE(CMSFlexGrid)
public:
	CLSID const& GetClsid()
	{
		static CLSID const clsid
			= { 0x6262d3a0, 0x531b, 0x11cf, { 0x91, 0xf6, 0xc2, 0x86, 0x3c, 0x38, 0x5e, 0x30 } };
		return clsid;
	}
	virtual BOOL Create(LPCTSTR lpszClassName,
		LPCTSTR lpszWindowName, DWORD dwStyle,
		const RECT& rect,
		CWnd* pParentWnd, UINT nID,
		CCreateContext* pContext = NULL)
	{ return CreateControl(GetClsid(), lpszWindowName, dwStyle, rect, pParentWnd, nID); }

    BOOL Create(LPCTSTR lpszWindowName, DWORD dwStyle,
		const RECT& rect, CWnd* pParentWnd, UINT nID,
		CFile* pPersist = NULL, BOOL bStorage = FALSE,
		BSTR bstrLicKey = NULL)
	{ return CreateControl(GetClsid(), lpszWindowName, dwStyle, rect, pParentWnd, nID,
		pPersist, bStorage, bstrLicKey); }

// Attributes
public:

// Operations
public:
	long GetRows();
	void SetRows(long nNewValue);
	long GetCols();
	void SetCols(long nNewValue);
	long GetFixedRows();
	void SetFixedRows(long nNewValue);
	long GetFixedCols();
	void SetFixedCols(long nNewValue);
	short GetVersion();
	CString GetFormatString();
	void SetFormatString(LPCTSTR lpszNewValue);
	long GetTopRow();
	void SetTopRow(long nNewValue);
	long GetLeftCol();
	void SetLeftCol(long nNewValue);
	long GetRow();
	void SetRow(long nNewValue);
	long GetCol();
	void SetCol(long nNewValue);
	long GetRowSel();
	void SetRowSel(long nNewValue);
	long GetColSel();
	void SetColSel(long nNewValue);
	CString GetText();
	void SetText(LPCTSTR lpszNewValue);
	unsigned long GetBackColor();
	void SetBackColor(unsigned long newValue);
	unsigned long GetForeColor();
	void SetForeColor(unsigned long newValue);
	unsigned long GetBackColorFixed();
	void SetBackColorFixed(unsigned long newValue);
	unsigned long GetForeColorFixed();
	void SetForeColorFixed(unsigned long newValue);
	unsigned long GetBackColorSel();
	void SetBackColorSel(unsigned long newValue);
	unsigned long GetForeColorSel();
	void SetForeColorSel(unsigned long newValue);
	unsigned long GetBackColorBkg();
	void SetBackColorBkg(unsigned long newValue);
	BOOL GetWordWrap();
	void SetWordWrap(BOOL bNewValue);
	COleFont GetFont();
	void SetRefFont(LPDISPATCH newValue);
	float GetFontWidth();
	void SetFontWidth(float newValue);
	CString GetCellFontName();
	void SetCellFontName(LPCTSTR lpszNewValue);
	float GetCellFontSize();
	void SetCellFontSize(float newValue);
	BOOL GetCellFontBold();
	void SetCellFontBold(BOOL bNewValue);
	BOOL GetCellFontItalic();
	void SetCellFontItalic(BOOL bNewValue);
	BOOL GetCellFontUnderline();
	void SetCellFontUnderline(BOOL bNewValue);
	BOOL GetCellFontStrikeThrough();
	void SetCellFontStrikeThrough(BOOL bNewValue);
	float GetCellFontWidth();
	void SetCellFontWidth(float newValue);
	long GetTextStyle();
	void SetTextStyle(long nNewValue);
	long GetTextStyleFixed();
	void SetTextStyleFixed(long nNewValue);
	BOOL GetScrollTrack();
	void SetScrollTrack(BOOL bNewValue);
	long GetFocusRect();
	void SetFocusRect(long nNewValue);
	long GetHighLight();
	void SetHighLight(long nNewValue);
	BOOL GetRedraw();
	void SetRedraw(BOOL bNewValue);
	long GetScrollBars();
	void SetScrollBars(long nNewValue);
	long GetMouseRow();
	long GetMouseCol();
	long GetCellLeft();
	long GetCellTop();
	long GetCellWidth();
	long GetCellHeight();
	long GetRowHeightMin();
	void SetRowHeightMin(long nNewValue);
	long GetFillStyle();
	void SetFillStyle(long nNewValue);
	long GetGridLines();
	void SetGridLines(long nNewValue);
	long GetGridLinesFixed();
	void SetGridLinesFixed(long nNewValue);
	unsigned long GetGridColor();
	void SetGridColor(unsigned long newValue);
	unsigned long GetGridColorFixed();
	void SetGridColorFixed(unsigned long newValue);
	unsigned long GetCellBackColor();
	void SetCellBackColor(unsigned long newValue);
	unsigned long GetCellForeColor();
	void SetCellForeColor(unsigned long newValue);
	short GetCellAlignment();
	void SetCellAlignment(short nNewValue);
	long GetCellTextStyle();
	void SetCellTextStyle(long nNewValue);
	short GetCellPictureAlignment();
	void SetCellPictureAlignment(short nNewValue);
	CString GetClip();
	void SetClip(LPCTSTR lpszNewValue);
	void SetSort(short nNewValue);
	long GetSelectionMode();
	void SetSelectionMode(long nNewValue);
	long GetMergeCells();
	void SetMergeCells(long nNewValue);
	BOOL GetAllowBigSelection();
	void SetAllowBigSelection(BOOL bNewValue);
	long GetAllowUserResizing();
	void SetAllowUserResizing(long nNewValue);
	long GetBorderStyle();
	void SetBorderStyle(long nNewValue);
	long GetHWnd();
	BOOL GetEnabled();
	void SetEnabled(BOOL bNewValue);
	long GetAppearance();
	void SetAppearance(long nNewValue);
	long GetMousePointer();
	void SetMousePointer(long nNewValue);
	CPicture GetMouseIcon();
	void SetRefMouseIcon(LPDISPATCH newValue);
	long GetPictureType();
	void SetPictureType(long nNewValue);
	CPicture GetPicture();
	CPicture GetCellPicture();
	void SetRefCellPicture(LPDISPATCH newValue);
	CString GetTextArray(long index);
	void SetTextArray(long index, LPCTSTR lpszNewValue);
	short GetColAlignment(long index);
	void SetColAlignment(long index, short nNewValue);
	long GetColWidth(long index);
	void SetColWidth(long index, long nNewValue);
	long GetRowHeight(long index);
	void SetRowHeight(long index, long nNewValue);
	BOOL GetMergeRow(long index);
	void SetMergeRow(long index, BOOL bNewValue);
	BOOL GetMergeCol(long index);
	void SetMergeCol(long index, BOOL bNewValue);
	void SetRowPosition(long index, long nNewValue);
	void SetColPosition(long index, long nNewValue);
	long GetRowData(long index);
	void SetRowData(long index, long nNewValue);
	long GetColData(long index);
	void SetColData(long index, long nNewValue);
	CString GetTextMatrix(long Row, long Col);
	void SetTextMatrix(long Row, long Col, LPCTSTR lpszNewValue);
	void AddItem(LPCTSTR Item, const VARIANT& index);
	void RemoveItem(long index);
	void Clear();
	void Refresh();
	CRowCursor GetDataSource();
	void SetDataSource(LPDISPATCH newValue);
	BOOL GetRowIsVisible(long index);
	BOOL GetColIsVisible(long index);
	long GetRowPos(long index);
	long GetColPos(long index);
	short GetGridLineWidth();
	void SetGridLineWidth(short nNewValue);
	short GetFixedAlignment(long index);
	void SetFixedAlignment(long index, short nNewValue);
	BOOL GetRightToLeft();
	void SetRightToLeft(BOOL bNewValue);
	long GetOLEDropMode();
	void SetOLEDropMode(long nNewValue);
	void OLEDrag();
};