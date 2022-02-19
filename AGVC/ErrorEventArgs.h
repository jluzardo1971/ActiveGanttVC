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


class ErrorEventArgs : public CCmdTarget
{
	DECLARE_DYNCREATE(ErrorEventArgs)

public:
	ErrorEventArgs(void);
	virtual ~ErrorEventArgs();
	virtual void OnFinalRelease();

// Member Variables
public:
	LONG mp_lNumber;				
	CString mp_sDescription;				
	CString mp_sSource;				

//Internal Properties (Extensions to ODL)
public:
	LONG GetNumber(void);
	void SetNumber(LONG Value);
	CString GetDescription(void);
	void SetDescription(CString Value);
	CString GetSource(void);
	void SetSource(CString Value);

//Internal Properties
public:

//Internal Methods
public:
	void Clear(void);

protected:
	DECLARE_DISPATCH_MAP()
	DECLARE_MESSAGE_MAP()
	DECLARE_OLECREATE(ErrorEventArgs)
	DECLARE_INTERFACE_MAP()

LONG odl_GetNumber(void)
{
	
	return GetNumber();
}

void odl_SetNumber(LONG Value)
{
	
	SetNumber(Value);
}

BSTR odl_GetDescription(void)
{
	
	return GetDescription().AllocSysString();
}

void odl_SetDescription(LPCTSTR Value)
{
	
	SetDescription(Value);
}

BSTR odl_GetSource(void)
{
	
	return GetSource().AllocSysString();
}

void odl_SetSource(LPCTSTR Value)
{
	
	SetSource(Value);
}

};