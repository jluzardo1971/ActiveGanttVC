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


// clsScrollBarSeparator command target

class clsScrollBarSeparator : public CCmdTarget
{
	DECLARE_DYNCREATE(clsScrollBarSeparator)
	clsScrollBarSeparator();

public:
	CActiveGanttVCCtl* mp_oControl;

	clsScrollBarSeparator(CActiveGanttVCCtl* oControl);
	virtual ~clsScrollBarSeparator();

	virtual void OnFinalRelease();

// Member Variables
public:
	CString mp_sStyleIndex;
	clsStyle* mp_oStyle;

//Internal Properties (Extensions to ODL)
public:

//Internal Properties (Extensions to ODL)
public:
	CString GetStyleIndex(void);
	void SetStyleIndex(CString Value);
	clsStyle* GetStyle(void);

//Internal Methods
public:
	CString GetXML(void);
	void SetXML(CString sXML);

protected:
	DECLARE_MESSAGE_MAP()
	DECLARE_OLECREATE(clsScrollBarSeparator)
	DECLARE_DISPATCH_MAP()
	DECLARE_INTERFACE_MAP()

BSTR odl_GetXML(void)
{
	
	return GetXML().AllocSysString();
}

void odl_SetXML(LPCTSTR sXML)
{
	
	SetXML(sXML);
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

};