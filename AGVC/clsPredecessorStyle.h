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


class clsPredecessorStyle : public CCmdTarget
{
	DECLARE_DYNCREATE(clsPredecessorStyle)
	clsPredecessorStyle();

public:
	CActiveGanttVCCtl* mp_oControl;

	clsPredecessorStyle(CActiveGanttVCCtl* oControl);
	virtual ~clsPredecessorStyle();
	virtual void OnFinalRelease();

// Member Variables
public:
	LONG mp_lArrowSize;
	LONG mp_yLineStyle;
	LONG mp_lLineWidth;
	LONG mp_lXOffset;
	LONG mp_lYOffset;
	Gdiplus::Color mp_clrLineColor;

//Internal Properties (Extensions to ODL)
public:
	Gdiplus::Color GetLineColor(void);
	void SetLineColor(Gdiplus::Color Value);
	LONG GetXOffset(void);
	void SetXOffset(LONG Value);
	LONG GetYOffset(void);
	void SetYOffset(LONG Value);
	LONG GetLineWidth(void);
	void SetLineWidth(LONG Value);
	LONG GetLineStyle(void);
	void SetLineStyle(LONG Value);
	LONG GetArrowSize(void);
	void SetArrowSize(LONG Value);

//Internal Properties
public:

//Internal Methods
public:
	CString GetXML(void);
	void SetXML(CString sXML);
	void Clear();
	void Clone(clsPredecessorStyle* oClone);

protected:
	DECLARE_DISPATCH_MAP()
	DECLARE_MESSAGE_MAP()
	DECLARE_OLECREATE(clsPredecessorStyle)
	DECLARE_INTERFACE_MAP()

OLE_COLOR odl_GetLineColor(void)
{
	
	return (OLE_COLOR) GetLineColor().ToCOLORREF();
}

void odl_SetLineColor(OLE_COLOR Value)
{
	
	Gdiplus::Color oColor;
	oColor.SetFromCOLORREF(g_TranslateColor(Value));
	SetLineColor(oColor);
}

LONG odl_GetXOffset(void)
{
	
	return GetXOffset();
}

void odl_SetXOffset(LONG Value)
{
	
	SetXOffset(Value);
}

LONG odl_GetYOffset(void)
{
	
	return GetYOffset();
}

void odl_SetYOffset(LONG Value)
{
	
	SetYOffset(Value);
}

LONG odl_GetLineWidth(void)
{
	
	return GetLineWidth();
}

void odl_SetLineWidth(LONG Value)
{
	
	SetLineWidth(Value);
}

LONG odl_GetLineStyle(void)
{
	
	return GetLineStyle();
}

void odl_SetLineStyle(LONG Value)
{
	
	SetLineStyle(Value);
}

LONG odl_GetArrowSize(void)
{
	
	return GetArrowSize();
}

void odl_SetArrowSize(LONG Value)
{
	
	SetArrowSize(Value);
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