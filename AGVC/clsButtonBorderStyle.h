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


// clsButtonBorderStyle command target

class clsButtonBorderStyle : public CCmdTarget
{
	DECLARE_DYNCREATE(clsButtonBorderStyle)
	clsButtonBorderStyle();

public:
	CActiveGanttVCCtl* mp_oControl;

	clsButtonBorderStyle(CActiveGanttVCCtl* oControl);
	virtual ~clsButtonBorderStyle();

	virtual void OnFinalRelease();

// Member Variables
public:
	Gdiplus::Color mp_clrRaisedExteriorLeftTopColor;
    Gdiplus::Color mp_clrRaisedInteriorLeftTopColor;
    Gdiplus::Color mp_clrRaisedExteriorRightBottomColor;
    Gdiplus::Color mp_clrRaisedInteriorRightBottomColor;
    Gdiplus::Color mp_clrSunkenExteriorLeftTopColor;
    Gdiplus::Color mp_clrSunkenInteriorLeftTopColor;
    Gdiplus::Color mp_clrSunkenExteriorRightBottomColor;
    Gdiplus::Color mp_clrSunkenInteriorRightBottomColor;

//Internal Properties (Extensions to ODL)
public:
	Gdiplus::Color GetRaisedExteriorLeftTopColor(void);
	void SetRaisedExteriorLeftTopColor(Gdiplus::Color Value);
	Gdiplus::Color GetRaisedInteriorLeftTopColor(void);
	void SetRaisedInteriorLeftTopColor(Gdiplus::Color Value);
	Gdiplus::Color GetRaisedExteriorRightBottomColor(void);
	void SetRaisedExteriorRightBottomColor(Gdiplus::Color Value);
	Gdiplus::Color GetRaisedInteriorRightBottomColor(void);
	void SetRaisedInteriorRightBottomColor(Gdiplus::Color Value);
	Gdiplus::Color GetSunkenExteriorLeftTopColor(void);
	void SetSunkenExteriorLeftTopColor(Gdiplus::Color Value);
	Gdiplus::Color GetSunkenInteriorLeftTopColor(void);
	void SetSunkenInteriorLeftTopColor(Gdiplus::Color Value);
	Gdiplus::Color GetSunkenExteriorRightBottomColor(void);
	void SetSunkenExteriorRightBottomColor(Gdiplus::Color Value);
	Gdiplus::Color GetSunkenInteriorRightBottomColor(void);
	void SetSunkenInteriorRightBottomColor(Gdiplus::Color Value);


//Internal Methods
public:
	CString GetXML(void);
	void SetXML(CString sXML);
	void Clear();
	void Clone(clsButtonBorderStyle* oClone);

protected:
	DECLARE_MESSAGE_MAP()
	DECLARE_OLECREATE(clsButtonBorderStyle)
	DECLARE_DISPATCH_MAP()
	DECLARE_INTERFACE_MAP()



OLE_COLOR odl_GetRaisedExteriorLeftTopColor(void)
{
	
	return (OLE_COLOR) GetRaisedExteriorLeftTopColor().ToCOLORREF();
}

void odl_SetRaisedExteriorLeftTopColor(OLE_COLOR Value)
{
	
	Gdiplus::Color oColor;
	oColor.SetFromCOLORREF(g_TranslateColor(Value));
	SetRaisedExteriorLeftTopColor(oColor);
}

OLE_COLOR odl_GetRaisedInteriorLeftTopColor(void)
{
	
	return (OLE_COLOR) GetRaisedInteriorLeftTopColor().ToCOLORREF();
}

void odl_SetRaisedInteriorLeftTopColor(OLE_COLOR Value)
{
	
	Gdiplus::Color oColor;
	oColor.SetFromCOLORREF(g_TranslateColor(Value));
	SetRaisedInteriorLeftTopColor(oColor);
}

OLE_COLOR odl_GetRaisedExteriorRightBottomColor(void)
{
	
	return (OLE_COLOR) GetRaisedExteriorRightBottomColor().ToCOLORREF();
}

void odl_SetRaisedExteriorRightBottomColor(OLE_COLOR Value)
{
	
	Gdiplus::Color oColor;
	oColor.SetFromCOLORREF(g_TranslateColor(Value));
	SetRaisedExteriorRightBottomColor(oColor);
}

OLE_COLOR odl_GetRaisedInteriorRightBottomColor(void)
{
	
	return (OLE_COLOR) GetRaisedInteriorRightBottomColor().ToCOLORREF();
}

void odl_SetRaisedInteriorRightBottomColor(OLE_COLOR Value)
{
	
	Gdiplus::Color oColor;
	oColor.SetFromCOLORREF(g_TranslateColor(Value));
	SetRaisedInteriorRightBottomColor(oColor);
}

OLE_COLOR odl_GetSunkenExteriorLeftTopColor(void)
{
	
	return (OLE_COLOR) GetSunkenExteriorLeftTopColor().ToCOLORREF();
}

void odl_SetSunkenExteriorLeftTopColor(OLE_COLOR Value)
{
	
	Gdiplus::Color oColor;
	oColor.SetFromCOLORREF(g_TranslateColor(Value));
	SetSunkenExteriorLeftTopColor(oColor);
}

OLE_COLOR odl_GetSunkenInteriorLeftTopColor(void)
{
	
	return (OLE_COLOR) GetSunkenInteriorLeftTopColor().ToCOLORREF();
}

void odl_SetSunkenInteriorLeftTopColor(OLE_COLOR Value)
{
	
	Gdiplus::Color oColor;
	oColor.SetFromCOLORREF(g_TranslateColor(Value));
	SetSunkenInteriorLeftTopColor(oColor);
}

OLE_COLOR odl_GetSunkenExteriorRightBottomColor(void)
{
	
	return (OLE_COLOR) GetSunkenExteriorRightBottomColor().ToCOLORREF();
}

void odl_SetSunkenExteriorRightBottomColor(OLE_COLOR Value)
{
	
	Gdiplus::Color oColor;
	oColor.SetFromCOLORREF(g_TranslateColor(Value));
	SetSunkenExteriorRightBottomColor(oColor);
}

OLE_COLOR odl_GetSunkenInteriorRightBottomColor(void)
{
	
	return (OLE_COLOR) GetSunkenInteriorRightBottomColor().ToCOLORREF();
}

void odl_SetSunkenInteriorRightBottomColor(OLE_COLOR Value)
{
	
	Gdiplus::Color oColor;
	oColor.SetFromCOLORREF(g_TranslateColor(Value));
	SetSunkenInteriorRightBottomColor(oColor);
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