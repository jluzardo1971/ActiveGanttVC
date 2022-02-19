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

class clsTimeLine;
class clsClientArea;

class clsView : public clsItemBase
{
	DECLARE_DYNCREATE(clsView)
	clsView();

public:
	CActiveGanttVCCtl* mp_oControl;

	clsView(CActiveGanttVCCtl* oControl);
	virtual ~clsView();
	virtual void OnFinalRelease();

// Member Variables
public:
	clsTimeLine* mp_oTimeLine;
	clsClientArea* mp_oClientArea;
	CString mp_sTag;
	LONG mp_yScrollInterval;
	LONG mp_yInterval;
	LONG mp_lFactor;

//Internal Properties (Extensions to ODL)
public:
	CString GetKey(void);
	void SetKey(CString Value);
	clsTimeLine* GetTimeLine(void);
	clsClientArea* GetClientArea(void);
	CString GetTag(void);
	void SetTag(CString Value);
	LONG GetInterval(void);
	void SetInterval(LONG Value);
	LONG GetIndex(void);
	void SetIndex(LONG Value);
	LONG GetFactor(void);
	void SetFactor(LONG Value);

//Internal Properties
public:
	LONG Getf_ScrollInterval(void);
	void Setf_ScrollInterval(LONG value);

//Internal Methods
public:
	void Finalize(void);
	CString GetXML(void);
	void SetXML(CString sXML);
	void Clear(void);
	clsView* clsView::Clone(CString Key);

protected:
	DECLARE_DISPATCH_MAP()
	DECLARE_MESSAGE_MAP()
	DECLARE_OLECREATE(clsView)
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

IDispatch* odl_GetTimeLine(void)
{
	
	return GetTimeLine()->GetIDispatch(TRUE);
}

IDispatch* odl_GetClientArea(void)
{
	
	return GetClientArea()->GetIDispatch(TRUE);
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

BSTR odl_GetXML(void)
{
	
	return GetXML().AllocSysString();
}

void odl_SetXML(LPCTSTR sXML)
{
	
	SetXML(sXML);
}

void odl_Clear(void)
{
	
	Clear();
}

IDispatch* odl_Clone(LPCTSTR Key)
{
	
	return Clone(Key)->GetIDispatch(TRUE);
}


};