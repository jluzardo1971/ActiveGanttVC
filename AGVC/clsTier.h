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

class clsFont;
class clsTierArea;

class clsTier : public CCmdTarget
{
	DECLARE_DYNCREATE(clsTier)

public:
	CActiveGanttVCCtl* mp_oControl;

	clsTier(CActiveGanttVCCtl* oControl, clsTierArea* oTierArea,LONG yTierPosition);
	clsTier();
	virtual ~clsTier();
	virtual void OnFinalRelease();

// Member Variables
public:
	BOOL mp_bVisible;
	LONG mp_lFactor;
	LONG mp_lHeight;
	LONG mp_yInterval;
	CString mp_sTag;
	LONG mp_yTierType;
	LONG mp_yTierPosition;
	CString mp_sTierPosition;
	clsTierArea* mp_oTierArea;
	CString mp_sStyleIndex;
	clsStyle* mp_oStyle;
	LONG mp_yBackgroundMode;

//Internal Properties (Extensions to ODL)
public:
	BOOL GetVisible(void);
	void SetVisible(BOOL Value);
	CString GetTag(void);
	void SetTag(CString Value);
	LONG GetInterval(void);
	void SetInterval(LONG Value);
	LONG GetFactor(void);
	void SetFactor(LONG Value);
	LONG GetTierType(void);
	void SetTierType(LONG Value);
	LONG GetHeight(void);
	void SetHeight(LONG Value);
	CString GetStyleIndex(void);
	void SetStyleIndex(CString Value);
	clsStyle* GetStyle(void);
	LONG GetBackgroundMode(void);
	void SetBackgroundMode(LONG Value);

//Internal Properties
public:

//Internal Methods
public:
	void Position(void);
	void Draw(COleDateTime dtStart,COleDateTime dtEnd,LONG lTop,LONG lBottom);
	CString GetXML(void);
	void SetXML(CString sXML);
	void Clear();
	void Clone(clsTier* oClone);

protected:
	DECLARE_DISPATCH_MAP()
	DECLARE_MESSAGE_MAP()
	DECLARE_OLECREATE(clsTier)
	DECLARE_INTERFACE_MAP()

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

	LONG odl_GetTierType(void)
	{
		
		return GetTierType();
	}

	void odl_SetTierType(LONG Value)
	{
		
		SetTierType(Value);
	}

	LONG odl_GetHeight(void)
	{
		
		return GetHeight();
	}

	void odl_SetHeight(LONG Value)
	{
		
		SetHeight(Value);
	}

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

	LONG odl_GetBackgroundMode(void)
	{
		
		return GetBackgroundMode();
	}

	void odl_SetBackgroundMode(LONG Value)
	{
		
		SetBackgroundMode(Value);
	}


};