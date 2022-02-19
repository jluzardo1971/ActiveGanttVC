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
class clsLayer;
class clsStyle;

class clsPercentage : public clsItemBase
{
	DECLARE_DYNCREATE(clsPercentage)
	clsPercentage();

public:
	CActiveGanttVCCtl* mp_oControl;

	clsPercentage(CActiveGanttVCCtl* oControl);
	virtual ~clsPercentage();
	virtual void OnFinalRelease();

// Member Variables
public:
	FLOAT mp_fPercent;
	CString mp_sTaskKey;
	CString mp_sTag;
	BOOL mp_bVisible;
	BOOL mp_bAllowSize;
	clsTask* mp_oTask;
	CString mp_sFormat;
	CString mp_sStyleIndex;
	clsStyle* mp_oStyle;

//Internal Properties (Extensions to ODL)
public:
	BOOL GetAllowSize(void);
	void SetAllowSize(BOOL Value);
	CString GetKey(void);
	void SetKey(CString Value);
	FLOAT GetPercent(void);
	void SetPercent(FLOAT Value);
	CString GetTaskKey(void);
	void SetTaskKey(CString Value);
	clsTask* GetTask(void);
	clsLayer* GetLayer(void);
	CString GetTag(void);
	void SetTag(CString Value);
	LONG GetLeftTrim(void);
	LONG GetRightTrim(void);
	LONG GetLeft(void);
	LONG GetTop(void);
	LONG GetBottom(void);
	LONG GetRight(void);
	BOOL GetVisible(void);
	void SetVisible(BOOL Value);
	LONG GetIndex(void);
	void SetIndex(LONG Value);
	CString GetFormat(void);
	void SetFormat(CString Value);
	CString GetStyleIndex(void);
	void SetStyleIndex(CString Value);
	clsStyle* GetStyle(void);

//Internal Properties
public:
	BOOL Getf_bLeftVisible(void);
	BOOL Getf_bRightVisible(void);
	LONG GetRightSel(void);

//Internal Methods
public:
	CString GetXML(void);
	void SetXML(CString sXML);
	CString ToString(void);

protected:
	DECLARE_DISPATCH_MAP()
	DECLARE_MESSAGE_MAP()
	DECLARE_OLECREATE(clsPercentage)
	DECLARE_INTERFACE_MAP()

LONG odl_GetIndex(void)
{
	
	return GetIndex();
}

BOOL odl_GetAllowSize(void)
{
	
	return GetAllowSize();
}

void odl_SetAllowSize(BOOL Value)
{
	
	
	SetAllowSize(Value);
}

BSTR odl_GetKey(void)
{
	
	return GetKey().AllocSysString();
}

void odl_SetKey(LPCTSTR Value)
{
	
	SetKey(Value);
}

FLOAT odl_GetPercent(void)
{
	
	return GetPercent();
}

void odl_SetPercent(FLOAT Value)
{
	
	SetPercent(Value);
}

BSTR odl_GetTaskKey(void)
{
	
	return GetTaskKey().AllocSysString();
}

void odl_SetTaskKey(LPCTSTR Value)
{
	
	SetTaskKey(Value);
}

IDispatch* odl_GetTask(void)
{
	
	return GetTask()->GetIDispatch(TRUE);
}

IDispatch* odl_GetLayer(void)
{
	
	return GetLayer()->GetIDispatch(TRUE);
}

BSTR odl_GetTag(void)
{
	
	return GetTag().AllocSysString();
}

void odl_SetTag(LPCTSTR Value)
{
	
	SetTag(Value);
}

LONG odl_GetLeftTrim(void)
{
	
	return GetLeftTrim();
}

LONG odl_GetRightTrim(void)
{
	
	return GetRightTrim();
}

LONG odl_GetLeft(void)
{
	
	return GetLeft();
}

LONG odl_GetTop(void)
{
	
	return GetTop();
}

LONG odl_GetBottom(void)
{
	
	return GetBottom();
}

LONG odl_GetRight(void)
{
	
	return GetRight();
}

BOOL odl_GetVisible(void)
{
	
	return GetVisible();
}

void odl_SetVisible(BOOL Value)
{
	
	
	SetVisible(Value);
}

BSTR odl_GetFormat(void)
{
	
	return GetFormat().AllocSysString();
}

void odl_SetFormat(LPCTSTR Value)
{
	
	SetFormat(Value);
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

BSTR odl_GetXML(void)
{
	
	return GetXML().AllocSysString();
}

void odl_SetXML(LPCTSTR sXML)
{
	
	SetXML(sXML);
}

BSTR odl_ToString(void)
{
	return ToString().AllocSysString();
}



};