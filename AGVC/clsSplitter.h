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


class clsSplitter : public CCmdTarget
{
	DECLARE_DYNCREATE(clsSplitter)
	clsSplitter();

public:
	CActiveGanttVCCtl* mp_oControl;

	clsSplitter(CActiveGanttVCCtl* oControl);
	virtual ~clsSplitter();
	virtual void OnFinalRelease();

// Member Variables
public:
	LONG mp_lPosition;
	LONG mp_yAppearance;
	CArray<Gdiplus::Color, Gdiplus::Color> mp_aColors;
	LONG mp_yType;
	LONG mp_lWidth;
	CString mp_sStyleIndex;
	clsStyle* mp_oStyle;
	LONG mp_lOffset;



//Internal Properties (Extensions to ODL)
public:
	LONG GetAppearance(void);
	void SetAppearance(LONG Value);
	LONG GetPosition(void);
	void SetPosition(LONG Value);
	void SetColor(LONG Index, Gdiplus::Color oColor);
	Gdiplus::Color GetColor(LONG Index);
	LONG GetType(void);
	void SetType(LONG Value);
	LONG GetWidth(void);
	void SetWidth(LONG Value);
	CString GetStyleIndex(void);
	void SetStyleIndex(CString Value);
	clsStyle* GetStyle(void);
	LONG GetOffset(void);
	void SetOffset(LONG Value);

//Internal Properties
public:
	LONG GetLeft(void);
	LONG GetRight(void);

//Internal Methods
public:
	void Finalize(void);
	void Draw(void);
	void f_AdjustPosition(void);
	CString GetXML(void);
	void SetXML(CString sXML);

protected:
	DECLARE_DISPATCH_MAP()
	DECLARE_MESSAGE_MAP()
	DECLARE_OLECREATE(clsSplitter)
	DECLARE_INTERFACE_MAP()

LONG odl_GetAppearance(void)
{
	
	return GetAppearance();
}

void odl_SetAppearance(LONG Value)
{
	
	SetAppearance(Value);
}

LONG odl_GetPosition(void)
{
	
	return GetPosition();
}

void odl_SetPosition(LONG Value)
{
	
	SetPosition(Value);
}

LONG odl_GetOffset(void)
{
	
	return GetOffset();
}

void odl_SetOffset(LONG Value)
{
	
	SetOffset(Value);
}

void odl_SetColor(LONG Index, OLE_COLOR oColor)
{
	
	Gdiplus::Color oColor1;
	oColor1.SetFromCOLORREF(g_TranslateColor(oColor));
	SetColor(Index, oColor1);
}

OLE_COLOR odl_GetColor(LONG Index)
{
	
	Gdiplus::Color oReturnColor;
	oReturnColor = GetColor(Index);
	return oReturnColor.ToCOLORREF();
}

LONG odl_GetType(void)
{
	
	return GetType();
}

void odl_SetType(LONG Value)
{
	
	SetType(Value);
}

LONG odl_GetWidth(void)
{
	
	return GetWidth();
}

void odl_SetWidth(LONG Value)
{
	
	SetWidth(Value);
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


};