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

class clsTickMarks;
class clsTimeLine;
class clsStyle;
class clsTickMark;

class clsTickMarkArea : public CCmdTarget
{
	DECLARE_DYNCREATE(clsTickMarkArea)
	clsTickMarkArea();

public:
	CActiveGanttVCCtl* mp_oControl;

	clsTickMarkArea(CActiveGanttVCCtl* oControl, clsTimeLine* oTimeLine);
	virtual ~clsTickMarkArea();
	virtual void OnFinalRelease();

// Member Variables
public:
	LONG mp_lHeight;
	LONG mp_lBigTickMarkHeight;
	LONG mp_lMediumTickMarkHeight;
	LONG mp_lSmallTickMarkHeight;
	BOOL mp_bVisible;
	LONG mp_lTextOffset;
	clsTimeLine* mp_oTimeLine;
	CString mp_sStyleIndex;
	clsStyle* mp_oStyle;
	clsTickMarks* mp_oTickMarks;				

//Internal Properties (Extensions to ODL)
public:
	LONG GetHeight(void);
	void SetHeight(LONG Value);
	LONG GetBigTickMarkHeight(void);
	void SetBigTickMarkHeight(LONG Value);
	LONG GetMediumTickMarkHeight(void);
	void SetMediumTickMarkHeight(LONG Value);
	LONG GetSmallTickMarkHeight(void);
	void SetSmallTickMarkHeight(LONG Value);
	BOOL GetVisible(void);
	void SetVisible(BOOL Value);
	CString GetStyleIndex(void);
	void SetStyleIndex(CString Value);
	LONG GetTextOffset(void);
	void SetTextOffset(LONG Value);
	clsTickMarks* GetTickMarks(void);
	clsStyle* GetStyle(void);

//Internal Properties
public:

//Internal Methods
public:
	void Draw(void);
	void PaintTickMark(COleDateTime dtDate, clsTickMark* oTickMark);
	void PaintText(LONG fXCoordinate, clsTickMark* oTickMark);
	CString GetXML(void);
	void SetXML(CString sXML);
	void Clear();
	void Clone(clsTickMarkArea* oClone);

protected:
	DECLARE_DISPATCH_MAP()
	DECLARE_MESSAGE_MAP()
	DECLARE_OLECREATE(clsTickMarkArea)
	DECLARE_INTERFACE_MAP()

LONG odl_GetHeight(void)
{
	
	return GetHeight();
}

void odl_SetHeight(LONG Value)
{
	
	SetHeight(Value);
}

LONG odl_GetBigTickMarkHeight(void)
{
	
	return GetBigTickMarkHeight();
}

void odl_SetBigTickMarkHeight(LONG Value)
{
	
	SetBigTickMarkHeight(Value);
}

LONG odl_GetMediumTickMarkHeight(void)
{
	
	return GetMediumTickMarkHeight();
}

void odl_SetMediumTickMarkHeight(LONG Value)
{
	
	SetMediumTickMarkHeight(Value);
}

LONG odl_GetSmallTickMarkHeight(void)
{
	
	return GetSmallTickMarkHeight();
}

void odl_SetSmallTickMarkHeight(LONG Value)
{
	
	SetSmallTickMarkHeight(Value);
}

BOOL odl_GetVisible(void)
{
	
	return GetVisible();
}

void odl_SetVisible(BOOL Value)
{
	
	
	SetVisible(Value);
}

BSTR odl_GetStyleIndex(void)
{
	
	return GetStyleIndex().AllocSysString();
}

void odl_SetStyleIndex(LPCTSTR Value)
{
	
	SetStyleIndex(Value);
}

LONG odl_GetTextOffset(void)
{
	
	return GetTextOffset();
}

void odl_SetTextOffset(LONG Value)
{
	
	SetTextOffset(Value);
}

IDispatch* odl_GetTickMarks(void)
{
	
	return GetTickMarks()->GetIDispatch(TRUE);
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

};