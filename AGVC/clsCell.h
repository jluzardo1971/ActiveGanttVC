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

class clsRow;
class clsImage;
class clsStyle;
class clsCells;

class clsCell : public clsItemBase
{
	DECLARE_DYNCREATE(clsCell)
	clsCell();

public:
	clsCell(CActiveGanttVCCtl* oControl, clsCells* oCells);
	virtual ~clsCell();
	virtual void OnFinalRelease();

// Member Variables
public:
	CActiveGanttVCCtl* mp_oControl;
	CString mp_sText;
	clsImage* mp_oImage;
	CString mp_sStyleIndex;
	CString mp_sTag;
	clsCells* mp_oCells;
	clsStyle* mp_oStyle;
	clsRow* mp_oRow;				
	CString mp_sRowKey;				
	LONG mp_lLeft;				
	LONG mp_lTop;				
	LONG mp_lRight;				
	LONG mp_lBottom;				
	LONG mp_lLeftTrim;				
	LONG mp_lRightTrim;
	CString mp_sImageTag;
	LONG mp_lTextLeft;
    LONG mp_lTextTop;
    LONG mp_lTextRight;
    LONG mp_lTextBottom;
	BOOL mp_bAllowTextEdit;

//Internal Properties (Extensions to ODL)
public:
	clsRow* GetRow(void);
	CString GetKey(void);
	void SetKey(CString Value);
	CString GetText(void);
	void SetText(CString Value);
	clsImage* GetImage(void);
	CString GetStyleIndex(void);
	void SetStyleIndex(CString Value);
	clsStyle* GetStyle(void);
	CString GetTag(void);
	void SetTag(CString Value);
	CString GetRowKey(void);
	LONG GetLeft(void);
	LONG GetTop(void);
	LONG GetRight(void);
	LONG GetBottom(void);
	LONG GetLeftTrim(void);
	LONG GetRightTrim(void);
	LONG GetIndex(void);
	void SetIndex(LONG Value);
	CString GetImageTag(void);
	void SetImageTag(CString Value);
	BOOL GetAllowTextEdit(void);
	void SetAllowTextEdit(BOOL Value);

//Internal Properties
public:


//Internal Methods
public:
	CString GetXML(void);
	void SetXML(CString sXML);

protected:
	DECLARE_DISPATCH_MAP()
	DECLARE_MESSAGE_MAP()
	DECLARE_OLECREATE(clsCell)
	DECLARE_INTERFACE_MAP()

IDispatch* odl_GetRow(void); //Circular Reference

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




BSTR odl_GetKey(void)
{
	
	return GetKey().AllocSysString();
}

void odl_SetKey(LPCTSTR Value)
{
	
	SetKey(Value);
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

BSTR odl_GetRowKey(void)
{
	
	return GetRowKey().AllocSysString();
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

LONG odl_GetLeftTrim(void)
{
	
	return GetLeftTrim();
}

LONG odl_GetRightTrim(void)
{
	
	return GetRightTrim();
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