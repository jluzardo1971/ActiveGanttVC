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

class clsTask;
class clsStyle;
class clsPredecessors;

class clsPredecessor : public clsItemBase
{
	DECLARE_DYNCREATE(clsPredecessor)
	clsPredecessor();

public:
	CActiveGanttVCCtl* mp_oControl;

	clsPredecessor(CActiveGanttVCCtl* oControl, clsPredecessors* oPredecessors);
	virtual ~clsPredecessor();
	virtual void OnFinalRelease();

// Member Variables
public:
	BOOL mp_bVisible;
	clsPredecessors* mp_clsPredecessors;
    CString mp_sSuccessorKey;
    clsTask* mp_oSuccessorTask; 
    CString mp_sPredecessorKey;
    clsTask* mp_oPredecessorTask;
	CString mp_sStyleIndex;
	CString mp_sTag;
	LONG mp_yPredecessorType;
	clsStyle* mp_oStyle;
    LONG mp_yLagInterval;
    LONG mp_lLagFactor;
	CArray<S_Rectangle, S_Rectangle> mp_oRectangles;
	BOOL mp_bWarning;
	CString mp_sWarningStyleIndex;
	clsStyle* mp_oWarningStyle;
	CString mp_sSelectedStyleIndex;
	clsStyle* mp_oSelectedStyle;


//Internal Properties (Extensions to ODL)
public:
	CString GetKey(void);
	void SetKey(CString Value);
	BOOL GetVisible(void);
	void SetVisible(BOOL Value);
	CString GetPredecessorKey(void);
	void SetPredecessorKey(CString Value);
	clsTask* GetPredecessorTask(void);
	LONG GetPredecessorType(void);
	void SetPredecessorType(LONG Value);
	CString GetStyleIndex(void);
	void SetStyleIndex(CString Value);
	clsStyle* GetStyle(void);
	CString GetTag(void);
	void SetTag(CString Value);
	LONG GetIndex(void);
	void SetIndex(LONG Value);

	CString GetSuccessorKey(void);
	void SetSuccessorKey(CString Value);

	clsTask* GetSuccessorTask(void);

	LONG GetLagInterval(void);
	void SetLagInterval(LONG Value);

	LONG GetLagFactor(void);
	void SetLagFactor(LONG Value);

	BOOL GetWarning(void);

	CString GetWarningStyleIndex(void);
	void SetWarningStyleIndex(CString Value);
	clsStyle* GetWarningStyle(void);

	CString GetSelectedStyleIndex(void);
	void SetSelectedStyleIndex(CString Value);
	clsStyle* GetSelectedStyle(void);


//Internal Properties
public:

//Internal Methods
public:
	CString GetXML(void);
	void SetXML(CString sXML);
	void AddRectangle(S_Rectangle oRectangle);
	void ClearRectangles();
	BOOL HitTest(int X, int Y);
	void Check(LONG Mode);

protected:
	DECLARE_DISPATCH_MAP()
	DECLARE_MESSAGE_MAP()
	DECLARE_OLECREATE(clsPredecessor)
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

BSTR odl_GetPredecessorKey(void)
{
	
	return GetPredecessorKey().AllocSysString();
}

void odl_SetPredecessorKey(LPCTSTR Value)
{
	
	SetPredecessorKey(Value);
}

IDispatch* odl_GetPredecessorTask(void)
{
	
	return GetPredecessorTask()->GetIDispatch(TRUE);
}

LONG odl_GetPredecessorType(void)
{
	
	return GetPredecessorType();
}

void odl_SetPredecessorType(LONG Value)
{
	
	SetPredecessorType(Value);
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

BSTR odl_GetSuccessorKey(void)
{
	
	return GetSuccessorKey().AllocSysString();
}

void odl_SetSuccessorKey(LPCTSTR Value)
{
	
	SetSuccessorKey(Value);
}

IDispatch* odl_GetSuccessorTask(void)
{
	
	return GetSuccessorTask()->GetIDispatch(TRUE);
}

LONG odl_GetLagInterval(void)
{
	
	return GetLagInterval();
}

void odl_SetLagInterval(LONG Value)
{
	
	SetLagInterval(Value);
}

LONG odl_GetLagFactor(void)
{
	
	return GetLagFactor();
}

void odl_SetLagFactor(LONG Value)
{
	
	SetLagFactor(Value);
}

BOOL odl_GetWarning(void)
{
	
	return GetWarning();
}



BSTR odl_GetWarningStyleIndex(void)
{
	
	return GetWarningStyleIndex().AllocSysString();
}

void odl_SetWarningStyleIndex(LPCTSTR Value)
{
	
	SetWarningStyleIndex(Value);
}

IDispatch* odl_GetWarningStyle(void)
{
	
	return GetWarningStyle()->GetIDispatch(TRUE);
}

BSTR odl_GetSelectedStyleIndex(void)
{
	
	return GetSelectedStyleIndex().AllocSysString();
}

void odl_SetSelectedStyleIndex(LPCTSTR Value)
{
	
	SetSelectedStyleIndex(Value);
}

IDispatch* odl_GetSelectedStyle(void)
{
	
	return GetSelectedStyle()->GetIDispatch(TRUE);
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