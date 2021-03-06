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

class clsColumn : public clsItemBase
{
	DECLARE_DYNCREATE(clsColumn)
	clsColumn();

public:
	CActiveGanttVCCtl* mp_oControl;

	clsColumn(CActiveGanttVCCtl* oControl);
	virtual ~clsColumn();
	virtual void OnFinalRelease();

// Member Variables
public:
	BOOL mp_bAllowMove;
	BOOL mp_bAllowSize;
	LONG mp_lWidth;
	CString mp_sText;
	clsImage* mp_oImage;
	CString mp_sStyleIndex;
	CString mp_sTag;
	LONG mp_lLeft;
	LONG mp_lRight;
	BOOL mp_bVisible;
	clsStyle* mp_oStyle;
	LONG mp_lLeftTrim;				
	LONG mp_lRightTrim;				
	LONG mp_lTop;				
	LONG mp_lBottom;
	CString mp_sImageTag;
	LONG mp_lTextLeft;
    LONG mp_lTextTop;
    LONG mp_lTextRight;
    LONG mp_lTextBottom;
	BOOL mp_bAllowTextEdit;
	BOOL mp_bTreeViewColumnIndex;

//Internal Properties (Extensions to ODL)
public:
	BOOL GetAllowMove(void);
	void SetAllowMove(BOOL Value);
	BOOL GetAllowSize(void);
	void SetAllowSize(BOOL Value);
	CString GetKey(void);
	void SetKey(CString Value);
	LONG GetWidth(void);
	void SetWidth(LONG Value);
	CString GetText(void);
	void SetText(CString Value);
	clsImage* GetImage(void);
	CString GetStyleIndex(void);
	void SetStyleIndex(CString Value);
	clsStyle* GetStyle(void);
	CString GetTag(void);
	void SetTag(CString Value);
	LONG GetLeftTrim(void);
	LONG GetRightTrim(void);
	LONG GetLeft(void);
	LONG GetTop(void);
	LONG GetRight(void);
	LONG GetBottom(void);
	BOOL GetVisible(void);
	LONG GetIndex(void);
	void SetIndex(LONG Value);
	CString GetImageTag(void);
	void SetImageTag(CString Value);
	BOOL GetAllowTextEdit(void);
	void SetAllowTextEdit(BOOL Value);

//Internal Properties
public:
	void Setf_lLeft(LONG Value);
	void Setf_lRight(LONG Value);
	void Setf_bVisible(BOOL Value);

//Internal Methods
public:
	CString GetXML(void);
	void SetXML(CString sXML);

protected:
	DECLARE_DISPATCH_MAP()
	DECLARE_MESSAGE_MAP()
	DECLARE_OLECREATE(clsColumn)
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

LONG odl_GetWidth(void)
{
	
	return GetWidth();
}

void odl_SetWidth(LONG Value)
{
	
	SetWidth(Value);
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

LONG odl_GetLeftTrim(void)
{
	
	return GetLeftTrim();
}

LONG odl_GetRightTrim(void)
{
	
	return GetRightTrim();
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




};