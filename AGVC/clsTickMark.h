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

class clsTickMarks;

class clsTickMark : public clsItemBase
{
	DECLARE_DYNCREATE(clsTickMark)
	clsTickMark();

public:
	CActiveGanttVCCtl* mp_oControl;

	clsTickMark(CActiveGanttVCCtl* oControl, clsTickMarks* oTickMarks);
	virtual ~clsTickMark();
	virtual void OnFinalRelease();

// Member Variables
public:
	BOOL mp_bDisplayText;
	clsTickMarks* mp_clsTickMarks;
	CString mp_sTextFormat;
	CString mp_sTag;
	LONG mp_yTickMarkType;
	LONG mp_yInterval;
	LONG mp_lFactor;

//Internal Properties (Extensions to ODL)
public:
	CString GetKey(void);
	void SetKey(CString Value);
	BOOL GetDisplayText(void);
	void SetDisplayText(BOOL Value);
	CString GetTextFormat(void);
	void SetTextFormat(CString Value);
	CString GetTag(void);
	void SetTag(CString Value);
	LONG GetInterval(void);
	void SetInterval(LONG Value);
	LONG GetFactor(void);
	void SetFactor(LONG Value);
	LONG GetTickMarkType(void);
	void SetTickMarkType(LONG Value);
	LONG GetIndex(void);
	void SetIndex(LONG Value);

//Internal Properties
public:

//Internal Methods
public:
	CString GetXML(void);
	void SetXML(CString sXML);
	void Clear();
	void Clone(clsTickMark* oClone);

protected:
	DECLARE_DISPATCH_MAP()
	DECLARE_MESSAGE_MAP()
	DECLARE_OLECREATE(clsTickMark)
	DECLARE_INTERFACE_MAP()

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

BOOL odl_GetDisplayText(void)
{
	
	return GetDisplayText();
}

void odl_SetDisplayText(BOOL Value)
{
	
	
	SetDisplayText(Value);
}

BSTR odl_GetTextFormat(void)
{
	
	return GetTextFormat().AllocSysString();
}

void odl_SetTextFormat(LPCTSTR Value)
{
	
	SetTextFormat(Value);
}

BSTR odl_GetTag(void)
{
	
	return GetTag().AllocSysString();
}

void odl_SetTag(LPCTSTR Value)
{
	
	SetTag(Value);
}

LONG odl_GetInterval(void)
{
	
	return GetInterval();
}

void odl_SetInterval(LONG Value)
{
	
	SetInterval(Value);
}

LONG odl_GetFactor(void)
{
	
	return GetFactor();
}

void odl_SetFactor(LONG Value)
{
	
	SetFactor(Value);
}

LONG odl_GetTickMarkType(void)
{
	
	return GetTickMarkType();
}

void odl_SetTickMarkType(LONG Value)
{
	
	SetTickMarkType(Value);
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