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

class clsImage;
class clsStyle;
class clsNode;
class clsCells;

class clsRow : public clsItemBase
{
	DECLARE_DYNCREATE(clsRow)
	clsRow();

public:
	CActiveGanttVCCtl* mp_oControl;

	clsRow(CActiveGanttVCCtl* oControl);
	virtual ~clsRow();
	virtual void OnFinalRelease();

// Member Variables
public:
	BOOL mp_bAllowSize;
	BOOL mp_bAllowMove;
	BOOL mp_bContainer;
	BOOL mp_bMergeCells;
	LONG mp_lHeight;
	CString mp_sText;
	clsImage *mp_oImage;
	CString mp_sStyleIndex;
	CString mp_sTag;
	CString mp_sClientAreaStyleIndex;
	LONG mp_lTop;
	LONG mp_lBottom;
	LONG mp_yClientAreaVisibility;
	clsStyle* mp_oStyle;
	clsStyle* mp_oClientAreaStyle;
	BOOL mp_bUseNodeImages;
	LONG mp_lLeft;				
	LONG mp_lRight;				
	BOOL mp_bVisible;				
	clsNode* mp_oNode;				
	clsCells* mp_oCells;
	CString mp_sImageTag;
	LONG mp_lTextLeft;
    LONG mp_lTextTop;
    LONG mp_lTextRight;
    LONG mp_lTextBottom;
	BOOL mp_bAllowTextEdit;
	BOOL mp_bHighlighted;

//Internal Properties (Extensions to ODL)
public:
	BOOL GetAllowMove(void);
	void SetAllowMove(BOOL Value);
	BOOL GetAllowSize(void);
	void SetAllowSize(BOOL Value);
	CString GetKey(void);
	void SetKey(CString Value);
	BOOL GetContainer(void);
	void SetContainer(BOOL Value);
	BOOL GetUseNodeImages(void);
	void SetUseNodeImages(BOOL Value);
	BOOL GetMergeCells(void);
	void SetMergeCells(BOOL Value);
	LONG GetHeight(void);
	void SetHeight(LONG Value);
	CString GetText(void);
	void SetText(CString Value);
	clsImage* GetImage(void);
	CString GetStyleIndex(void);
	void SetStyleIndex(CString Value);
	clsStyle* GetStyle(void);
	CString GetTag(void);
	void SetTag(CString Value);
	CString GetClientAreaStyleIndex(void);
	void SetClientAreaStyleIndex(CString Value);
	clsStyle* GetClientAreaStyle(void);
	LONG GetLeft(void);
	LONG GetTop(void);
	LONG GetRight(void);
	LONG GetBottom(void);
	BOOL GetVisible(void);
	LONG GetIndex(void);
	void SetIndex(LONG Value);
	clsNode* GetNode(void);
	clsCells* GetCells(void);
	CString GetImageTag(void);
	void SetImageTag(CString Value);
	BOOL GetAllowTextEdit(void);
	void SetAllowTextEdit(BOOL Value);
	BOOL GetHighlighted(void);
	void SetHighlighted(BOOL Value);
	Gdiplus::Color Getf_RowBackColor(void);


//Internal Properties
public:
	void Setf_lTop(LONG Value);
	void Setf_lBottom(LONG Value);
	LONG GetClientAreaVisibility(void);
	void SetClientAreaVisibility(LONG Value);

//Internal Methods
public:
	CString GetXML(void);
	void SetXML(CString sXML);

protected:
	DECLARE_DISPATCH_MAP()
	DECLARE_MESSAGE_MAP()
	DECLARE_OLECREATE(clsRow)
	DECLARE_INTERFACE_MAP()



BOOL odl_GetAllowTextEdit(void)
{
	
	return GetAllowTextEdit();
}

void odl_SetAllowTextEdit(BOOL Value)
{
	
	
	SetAllowTextEdit(Value);
}

LONG odl_GetIndex(void)
{
	
	return GetIndex();
}

BOOL odl_GetAllowMove(void)
{
	
	return GetAllowMove();
}

void odl_SetAllowMove(BOOL Value)
{
	
	
	SetAllowMove(Value);
}

BOOL odl_GetAllowSize(void)
{
	
	return GetAllowSize();
}

void odl_SetAllowSize(BOOL Value)
{
	
	
	SetAllowSize(Value);
}

BSTR odl_GetKey(void)
{
	
	return GetKey().AllocSysString();
}

void odl_SetKey(LPCTSTR Value)
{
	
	SetKey(Value);
}

BOOL odl_GetContainer(void)
{
	
	return GetContainer();
}

void odl_SetContainer(BOOL Value)
{
	
	
	SetContainer(Value);
}

BOOL odl_GetUseNodeImages(void)
{
	
	return GetUseNodeImages();
}

void odl_SetUseNodeImages(BOOL Value)
{
	
	
	SetUseNodeImages(Value);
}

BOOL odl_GetMergeCells(void)
{
	
	return GetMergeCells();
}

void odl_SetMergeCells(BOOL Value)
{
	
	SetMergeCells(Value);
}

LONG odl_GetHeight(void)
{
	
	return GetHeight();
}

void odl_SetHeight(LONG Value)
{
	
	SetHeight(Value);
}

BSTR odl_GetText(void)
{
	
	return GetText().AllocSysString();
}

void odl_SetText(LPCTSTR Value)
{
	
	SetText(Value);
}

IDispatch* odl_GetImage(void)
{
	
	return GetImage()->GetIDispatch(TRUE);
}

BSTR odl_GetStyleIndex(void)
{
	
	return GetStyleIndex().AllocSysString();
}

void odl_SetStyleIndex(LPCTSTR Value)
{
	
	SetStyleIndex(Value);
}

IDispatch* odl_GetStyle(void)
{
	
	return GetStyle()->GetIDispatch(TRUE);
}

BSTR odl_GetTag(void)
{
	
	return GetTag().AllocSysString();
}

void odl_SetTag(LPCTSTR Value)
{
	
	SetTag(Value);
}

BSTR odl_GetClientAreaStyleIndex(void)
{
	
	return GetClientAreaStyleIndex().AllocSysString();
}

void odl_SetClientAreaStyleIndex(LPCTSTR Value)
{
	
	SetClientAreaStyleIndex(Value);
}

IDispatch* odl_GetClientAreaStyle(void)
{
	
	return GetClientAreaStyle()->GetIDispatch(TRUE);
}

LONG odl_GetLeft(void)
{
	
	return GetLeft();
}

LONG odl_GetTop(void)
{
	
	return GetTop();
}

LONG odl_GetRight(void)
{
	
	return GetRight();
}

LONG odl_GetBottom(void)
{
	
	return GetBottom();
}

BOOL odl_GetVisible(void)
{
	
	return GetVisible();
}

IDispatch* odl_GetNode(void)
{
	
	return GetNode()->GetIDispatch(TRUE);
}

IDispatch* odl_GetCells(void)
{
	
	return GetCells()->GetIDispatch(TRUE);
}

BSTR odl_GetImageTag(void)
{
	
	return GetImageTag().AllocSysString();
}

void odl_SetImageTag(LPCTSTR Value)
{
	
	SetImageTag(Value);
}

BSTR odl_GetXML(void)
{
	
	return GetXML().AllocSysString();
}

void odl_SetXML(LPCTSTR sXML)
{
	
	SetXML(sXML);
}

BOOL odl_GetHighlighted(void)
{
	
	return GetHighlighted();
}

void odl_SetHighlighted(BOOL Value)
{
	
	SetHighlighted(Value);
}

};