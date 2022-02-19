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


class clsCustomBorderStyle : public CCmdTarget
{
	DECLARE_DYNCREATE(clsCustomBorderStyle)
	clsCustomBorderStyle();

public:
	CActiveGanttVCCtl* mp_oControl;

	clsCustomBorderStyle(CActiveGanttVCCtl* oControl);
	virtual ~clsCustomBorderStyle();
	virtual void OnFinalRelease();

// Member Variables
public:
	BOOL mp_bTop;
	BOOL mp_bBottom;
	BOOL mp_bLeft;
	BOOL mp_bRight;

//Internal Properties (Extensions to ODL)
public:
	BOOL GetLeft(void);
	void SetLeft(BOOL Value);
	BOOL GetTop(void);
	void SetTop(BOOL Value);
	BOOL GetRight(void);
	void SetRight(BOOL Value);
	BOOL GetBottom(void);
	void SetBottom(BOOL Value);

//Internal Properties
public:

//Internal Methods
public:
	CString GetXML(void);
	void SetXML(CString sXML);
	void Clear();
	void Clone(clsCustomBorderStyle* oClone);

protected:
	DECLARE_DISPATCH_MAP()
	DECLARE_MESSAGE_MAP()
	DECLARE_OLECREATE(clsCustomBorderStyle)
	DECLARE_INTERFACE_MAP()



BOOL odl_GetLeft(void)
{
	
	return GetLeft();
}

void odl_SetLeft(BOOL Value)
{
	
	
	SetLeft(Value);
}

BOOL odl_GetTop(void)
{
	
	return GetTop();
}

void odl_SetTop(BOOL Value)
{
	
	
	SetTop(Value);
}

BOOL odl_GetRight(void)
{
	
	return GetRight();
}

void odl_SetRight(BOOL Value)
{
	
	
	SetRight(Value);
}

BOOL odl_GetBottom(void)
{
	
	return GetBottom();
}

void odl_SetBottom(BOOL Value)
{
	
	
	SetBottom(Value);
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