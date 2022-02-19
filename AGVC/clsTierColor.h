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

class clsTierColors;

class clsTierColor : public clsItemBase
{
	DECLARE_DYNCREATE(clsTierColor)
	clsTierColor();

public:
	CActiveGanttVCCtl* mp_oControl;

	clsTierColor(CActiveGanttVCCtl* oControl, clsTierColors* oTierColors);
	virtual ~clsTierColor();
	virtual void OnFinalRelease();

// Member Variables
public:
	clsTierColors* mp_clsTierColors;
	Gdiplus::Color mp_clrBackColor;
	Gdiplus::Color mp_clrForeColor;
	Gdiplus::Color mp_clrStartGradientColor;
	Gdiplus::Color mp_clrEndGradientColor;
	Gdiplus::Color mp_clrHatchBackColor;
	Gdiplus::Color mp_clrHatchForeColor;

//Internal Properties (Extensions to ODL)
public:
	CString GetKey(void);
	void SetKey(CString Value);
	Gdiplus::Color GetForeColor(void);
	void SetForeColor(Gdiplus::Color Value);
	Gdiplus::Color GetBackColor(void);
	void SetBackColor(Gdiplus::Color Value);
	LONG GetIndex(void);
	void SetIndex(LONG Value);
	Gdiplus::Color GetStartGradientColor(void);
	void SetStartGradientColor(Gdiplus::Color Value);
	Gdiplus::Color GetEndGradientColor(void);
	void SetEndGradientColor(Gdiplus::Color Value);
	Gdiplus::Color GetHatchBackColor(void);
	void SetHatchBackColor(Gdiplus::Color Value);
	Gdiplus::Color GetHatchForeColor(void);
	void SetHatchForeColor(Gdiplus::Color Value);

//Internal Properties
public:

//Internal Methods
public:
	void Finalize(void);
	CString GetXML(void);
	void SetXML(CString sXML);
	void Clone(clsTierColor* oClone);

protected:
	DECLARE_DISPATCH_MAP()
	DECLARE_MESSAGE_MAP()
	DECLARE_OLECREATE(clsTierColor)
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

OLE_COLOR odl_GetBackColor(void)
{
	
	return (OLE_COLOR) GetBackColor().ToCOLORREF();
}

void odl_SetBackColor(OLE_COLOR Value)
{
	
	Gdiplus::Color oColor;
	oColor.SetFromCOLORREF(g_TranslateColor(Value));
	SetBackColor(oColor);
}

OLE_COLOR odl_GetStartGradientColor(void)
{
	
	return (OLE_COLOR) GetStartGradientColor().ToCOLORREF();
}

void odl_SetStartGradientColor(OLE_COLOR Value)
{
	
	Gdiplus::Color oColor;
	oColor.SetFromCOLORREF(g_TranslateColor(Value));
	SetStartGradientColor(oColor);
}

OLE_COLOR odl_GetEndGradientColor(void)
{
	
	return (OLE_COLOR) GetEndGradientColor().ToCOLORREF();
}

void odl_SetEndGradientColor(OLE_COLOR Value)
{
	
	Gdiplus::Color oColor;
	oColor.SetFromCOLORREF(g_TranslateColor(Value));
	SetEndGradientColor(oColor);
}

OLE_COLOR odl_GetHatchBackColor(void)
{
	
	return (OLE_COLOR) GetHatchBackColor().ToCOLORREF();
}

void odl_SetHatchBackColor(OLE_COLOR Value)
{
	
	Gdiplus::Color oColor;
	oColor.SetFromCOLORREF(g_TranslateColor(Value));
	SetHatchBackColor(oColor);
}

OLE_COLOR odl_GetHatchForeColor(void)
{
	
	return (OLE_COLOR) GetHatchForeColor().ToCOLORREF();
}

void odl_SetHatchForeColor(OLE_COLOR Value)
{
	
	Gdiplus::Color oColor;
	oColor.SetFromCOLORREF(g_TranslateColor(Value));
	SetHatchForeColor(oColor);
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