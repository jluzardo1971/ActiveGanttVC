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

class clsProgressLine : public CCmdTarget
{
	DECLARE_DYNCREATE(clsProgressLine)
	clsProgressLine();

public:
	CActiveGanttVCCtl* mp_oControl;

	clsProgressLine(CActiveGanttVCCtl* oControl);
	virtual ~clsProgressLine();
	virtual void OnFinalRelease();

// Member Variables
public:
	Gdiplus::Color mp_clrForeColor;
	COleDateTime mp_dtPosition;
	LONG mp_yLength;
	LONG mp_yLineType;
    LONG mp_lWidth;
    LONG mp_yLineStyle;
    CArray<Gdiplus::Color, Gdiplus::Color> mp_aColors;
    CString mp_sStyleIndex;
    clsStyle* mp_oStyle;

//Internal Properties (Extensions to ODL)
public:
	COleDateTime GetPosition(void);
	void SetPosition(COleDateTime Value);
	Gdiplus::Color GetForeColor(void);
	void SetForeColor(Gdiplus::Color Value);
	LONG GetLength(void);
	void SetLength(LONG Value);
	LONG GetLineType(void);
	void SetLineType(LONG Value);
	void SetColor(LONG Index, Gdiplus::Color oColor);
	Gdiplus::Color GetColor(LONG Index);
	LONG GetLineStyle(void);
	void SetLineStyle(LONG Value);
	LONG GetWidth(void);
	void SetWidth(LONG Value);
	CString GetStyleIndex(void);
	void SetStyleIndex(CString Value);
	clsStyle* GetStyle(void);

	clsTimeLine* mp_oTimeLine();

//Internal Properties
public:

//Internal Methods
public:
	void Finalize(void);
	void Draw(void);
	CString GetXML(void);
	void SetXML(CString sXML);
	void Clear();
	void Clone(clsProgressLine* oClone);

protected:
	DECLARE_DISPATCH_MAP()
	DECLARE_MESSAGE_MAP()
	DECLARE_OLECREATE(clsProgressLine)
	DECLARE_INTERFACE_MAP()

DATE odl_GetPosition(void)
{
	
	return mp_dtPosition;
}

void odl_SetPosition(DATE Value)
{
	SetPosition(Value);
}

OLE_COLOR odl_GetForeColor(void)
{
	
	return (OLE_COLOR) GetForeColor().ToCOLORREF();
}

void odl_SetForeColor(OLE_COLOR Value)
{
	
	Gdiplus::Color oColor;
	oColor.SetFromCOLORREF(g_TranslateColor(Value));
	SetForeColor(oColor);
}

LONG odl_GetLength(void)
{
	
	return GetLength();
}

void odl_SetLength(LONG Value)
{
	
	SetLength(Value);
}

LONG odl_GetLineType(void)
{
	
	return GetLineType();
}

void odl_SetLineType(LONG Value)
{
	
	SetLineType(Value);
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

LONG odl_GetLineStyle(void)
{
	
	return GetLineStyle();
}

void odl_SetLineStyle(LONG Value)
{
	
	SetLineStyle(Value);
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