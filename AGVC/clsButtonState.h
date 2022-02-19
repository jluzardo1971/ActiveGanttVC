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

class clsStyle;


// clsButtonState command target

class clsButtonState : public CCmdTarget
{
	DECLARE_DYNCREATE(clsButtonState)
	clsButtonState();

public:
	CActiveGanttVCCtl* mp_oControl;

	clsButtonState(CActiveGanttVCCtl* oControl, CString sType);
	
	virtual ~clsButtonState();

	virtual void OnFinalRelease();

// Member Variables
public:
    CString mp_sNormalStyleIndex;
    CString mp_sPressedStyleIndex;
    CString mp_sDisabledStyleIndex;
    clsStyle* mp_oNormalStyle;
    clsStyle* mp_oPressedStyle;
    clsStyle* mp_oDisabledStyle;
	CString mp_sType;

//Internal Properties
public:
	CString GetNormalStyleIndex(void);
	void SetNormalStyleIndex(CString Value);
	clsStyle* GetNormalStyle(void);

	CString GetPressedStyleIndex(void);
	void SetPressedStyleIndex(CString Value);
	clsStyle* GetPressedStyle(void);

	CString GetDisabledStyleIndex(void);
	void SetDisabledStyleIndex(CString Value);
	clsStyle* GetDisabledStyle(void);


//Internal Methods
public:
	CString GetXML(void);
	void SetXML(CString sXML);

	void Clear();
	void Clone(clsButtonState* oClone);

protected:
	DECLARE_MESSAGE_MAP()
	DECLARE_OLECREATE(clsButtonState)
	DECLARE_DISPATCH_MAP()
	DECLARE_INTERFACE_MAP()



BSTR odl_GetNormalStyleIndex(void)
{
	
	return GetNormalStyleIndex().AllocSysString();
}

void odl_SetNormalStyleIndex(LPCTSTR Value)
{
	
	SetNormalStyleIndex(Value);
}

IDispatch* odl_GetNormalStyle(void)
{
	
	return GetNormalStyle()->GetIDispatch(TRUE);
}

BSTR odl_GetPressedStyleIndex(void)
{
	
	return GetPressedStyleIndex().AllocSysString();
}

void odl_SetPressedStyleIndex(LPCTSTR Value)
{
	
	SetPressedStyleIndex(Value);
}

IDispatch* odl_GetPressedStyle(void)
{
	
	return GetPressedStyle()->GetIDispatch(TRUE);
}

BSTR odl_GetDisabledStyleIndex(void)
{
	
	return GetDisabledStyleIndex().AllocSysString();
}

void odl_SetDisabledStyleIndex(LPCTSTR Value)
{
	
	SetDisabledStyleIndex(Value);
}

IDispatch* odl_GetDisabledStyle(void)
{
	
	return GetDisabledStyle()->GetIDispatch(TRUE);
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