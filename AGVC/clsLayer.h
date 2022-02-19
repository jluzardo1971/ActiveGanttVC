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


class clsLayer : public clsItemBase
{
	DECLARE_DYNCREATE(clsLayer)
	clsLayer();

public:
	CActiveGanttVCCtl* mp_oControl;

	clsLayer(CActiveGanttVCCtl* oControl);
	virtual ~clsLayer();
	virtual void OnFinalRelease();

// Member Variables
public:
	BOOL mp_bVisible;
	CString mp_sTag;

//Internal Properties (Extensions to ODL)
public:
	CString GetKey(void);
	void SetKey(CString Value);
	BOOL GetVisible(void);
	void SetVisible(BOOL Value);
	CString GetTag(void);
	void SetTag(CString Value);
	LONG GetIndex(void);
	void SetIndex(LONG Value);

//Internal Properties
public:

//Internal Methods
public:
	void Finalize(void);
	CString GetXML(void);
	void SetXML(CString sXML);

protected:
	DECLARE_DISPATCH_MAP()
	DECLARE_MESSAGE_MAP()
	DECLARE_OLECREATE(clsLayer)
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

BOOL odl_GetVisible(void)
{
	
	return GetVisible();
}

void odl_SetVisible(BOOL Value)
{
	
	
	SetVisible(Value);
}

BSTR odl_GetTag(void)
{
	
	return GetTag().AllocSysString();
}

void odl_SetTag(LPCTSTR Value)
{
	
	SetTag(Value);
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