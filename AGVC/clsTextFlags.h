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

class clsStringFormat;

class clsTextFlags : public CCmdTarget
{
	DECLARE_DYNCREATE(clsTextFlags)
	clsTextFlags();

public:
	CActiveGanttVCCtl* mp_oControl;

	clsTextFlags(CActiveGanttVCCtl* oControl);
	clsTextFlags(const clsTextFlags& oTextFlags);
	virtual ~clsTextFlags();
	virtual void OnFinalRelease();

// Member Variables
public:
	LONG mp_yVerticalAlignment;
	LONG mp_yHorizontalAlignment;
	BOOL mp_bWordWrap;
	BOOL mp_bRightToLeft;
	LONG mp_lOffsetBottom;
	LONG mp_lOffsetLeft;
	LONG mp_lOffsetRight;
	LONG mp_lOffsetTop;

//Internal Properties (Extensions to ODL)
public:
	LONG GetVerticalAlignment(void);
	void SetVerticalAlignment(LONG Value);
	LONG GetHorizontalAlignment(void);
	void SetHorizontalAlignment(LONG Value);
	BOOL GetWordWrap(void);
	void SetWordWrap(BOOL Value);
	BOOL GetRightToLeft(void);
	void SetRightToLeft(BOOL Value);
	LONG GetOffsetBottom(void);
	void SetOffsetBottom(LONG Value);
	LONG GetOffsetLeft(void);
	void SetOffsetLeft(LONG Value);
	LONG GetOffsetRight(void);
	void SetOffsetRight(LONG Value);
	LONG GetOffsetTop(void);
	void SetOffsetTop(LONG Value);

//Internal Properties
public:

//Internal Methods
public:
	Gdiplus::StringFormat* GetFlags(void);
	CString GetXML(void);
	void SetXML(CString sXML);
	void Clear();
	void Clone(clsTextFlags* oClone);

protected:
	DECLARE_DISPATCH_MAP()
	DECLARE_MESSAGE_MAP()
	DECLARE_OLECREATE(clsTextFlags)
	DECLARE_INTERFACE_MAP()


LONG odl_GetVerticalAlignment(void)
{
	
	return GetVerticalAlignment();
}

void odl_SetVerticalAlignment(LONG Value)
{
	
	SetVerticalAlignment(Value);
}

LONG odl_GetHorizontalAlignment(void)
{
	
	return GetHorizontalAlignment();
}

void odl_SetHorizontalAlignment(LONG Value)
{
	
	SetHorizontalAlignment(Value);
}

BOOL odl_GetWordWrap(void)
{
	
	return GetWordWrap();
}

void odl_SetWordWrap(BOOL Value)
{
	
	
	SetWordWrap(Value);
}

BOOL odl_GetRightToLeft(void)
{
	
	return GetRightToLeft();
}

void odl_SetRightToLeft(BOOL Value)
{
	
	
	SetRightToLeft(Value);
}

LONG odl_GetOffsetBottom(void)
{
	
	return GetOffsetBottom();
}

void odl_SetOffsetBottom(LONG Value)
{
	
	SetOffsetBottom(Value);
}

LONG odl_GetOffsetLeft(void)
{
	
	return GetOffsetLeft();
}

void odl_SetOffsetLeft(LONG Value)
{
	
	SetOffsetLeft(Value);
}

LONG odl_GetOffsetRight(void)
{
	
	return GetOffsetRight();
}

void odl_SetOffsetRight(LONG Value)
{
	
	SetOffsetRight(Value);
}

LONG odl_GetOffsetTop(void)
{
	
	return GetOffsetTop();
}

void odl_SetOffsetTop(LONG Value)
{
	
	SetOffsetTop(Value);
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