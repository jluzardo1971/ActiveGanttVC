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


// TextEditEventArgs command target

class TextEditEventArgs : public CCmdTarget
{
	DECLARE_DYNCREATE(TextEditEventArgs)

public:
	TextEditEventArgs();
	virtual ~TextEditEventArgs();

	virtual void OnFinalRelease();


// Member Variables
public:
	LONG mp_yObjectType;				
	LONG mp_lObjectIndex;				
	LONG mp_lParentObjectIndex;
	CString mp_sText;

//Internal Properties (Extensions to ODL)
public:
	LONG GetObjectType(void);
	void SetObjectType(LONG Value);
	LONG GetObjectIndex(void);
	void SetObjectIndex(LONG Value);
	LONG GetParentObjectIndex(void);
	void SetParentObjectIndex(LONG Value);
	CString GetText(void);
	void SetText(CString Value);

//Internal Properties
public:

//Internal Methods
public:
	void Clear(void);

protected:
	DECLARE_MESSAGE_MAP()
	DECLARE_OLECREATE(TextEditEventArgs)
	DECLARE_DISPATCH_MAP()
	DECLARE_INTERFACE_MAP()

LONG odl_GetObjectType(void)
{
	
	return GetObjectType();
}

void odl_SetObjectType(LONG Value)
{
	
	SetObjectType(Value);
}

LONG odl_GetObjectIndex(void)
{
	
	return GetObjectIndex();
}

void odl_SetObjectIndex(LONG Value)
{
	
	SetObjectIndex(Value);
}

LONG odl_GetParentObjectIndex(void)
{
	
	return GetParentObjectIndex();
}

void odl_SetParentObjectIndex(LONG Value)
{
	
	SetParentObjectIndex(Value);
}

BSTR odl_GetText(void)
{
	
	return GetText().AllocSysString();
}

void odl_SetText(LPCTSTR Value)
{
	
	SetText(Value);
}

};