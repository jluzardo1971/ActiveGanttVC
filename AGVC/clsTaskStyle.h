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

class clsImage;

class clsTaskStyle : public CCmdTarget
{
	DECLARE_DYNCREATE(clsTaskStyle)
	clsTaskStyle();

public:
	CActiveGanttVCCtl* mp_oControl;

	clsTaskStyle(CActiveGanttVCCtl* oControl);
	virtual ~clsTaskStyle();
	virtual void OnFinalRelease();

// Member Variables
public:
	Gdiplus::Color mp_clrEndBorderColor;
	Gdiplus::Color mp_clrEndFillColor;
	Gdiplus::Color mp_clrStartBorderColor;
	Gdiplus::Color mp_clrStartFillColor;
	LONG mp_yEndShapeIndex;
	LONG mp_yStartShapeIndex;
	clsImage* mp_oStartImage;
	clsImage* mp_oMiddleImage;
	clsImage* mp_oEndImage;
	CString mp_sStartImageTag;
	CString mp_sMiddleImageTag;
	CString mp_sEndImageTag;

//Internal Properties (Extensions to ODL)
public:
	Gdiplus::Color GetEndBorderColor(void);
	void SetEndBorderColor(Gdiplus::Color Value);
	Gdiplus::Color GetEndFillColor(void);
	void SetEndFillColor(Gdiplus::Color Value);
	Gdiplus::Color GetStartBorderColor(void);
	void SetStartBorderColor(Gdiplus::Color Value);
	Gdiplus::Color GetStartFillColor(void);
	void SetStartFillColor(Gdiplus::Color Value);
	clsImage* GetStartImage(void);
	clsImage* GetMiddleImage(void);
	clsImage* GetEndImage(void);
	LONG GetStartShapeIndex(void);
	void SetStartShapeIndex(LONG Value);
	LONG GetEndShapeIndex(void);
	void SetEndShapeIndex(LONG Value);
	CString GetStartImageTag(void);
	void SetStartImageTag(CString Value);
	CString GetMiddleImageTag(void);
	void SetMiddleImageTag(CString Value);
	CString GetEndImageTag(void);
	void SetEndImageTag(CString Value);

//Internal Properties
public:

//Internal Methods
public:
	CString GetXML(void);
	void SetXML(CString sXML);
	void Clear();
	void Clone(clsTaskStyle* oClone);

protected:
	DECLARE_DISPATCH_MAP()
	DECLARE_MESSAGE_MAP()
	DECLARE_OLECREATE(clsTaskStyle)
	DECLARE_INTERFACE_MAP()

OLE_COLOR odl_GetEndBorderColor(void)
{
	
	return (OLE_COLOR) GetEndBorderColor().ToCOLORREF();
}

void odl_SetEndBorderColor(OLE_COLOR Value)
{
	
	Gdiplus::Color oColor;
	oColor.SetFromCOLORREF(g_TranslateColor(Value));
	SetEndBorderColor(oColor);
}

OLE_COLOR odl_GetEndFillColor(void)
{
	
	return (OLE_COLOR) GetEndFillColor().ToCOLORREF();
}

void odl_SetEndFillColor(OLE_COLOR Value)
{
	
	Gdiplus::Color oColor;
	oColor.SetFromCOLORREF(g_TranslateColor(Value));
	SetEndFillColor(oColor);
}

OLE_COLOR odl_GetStartBorderColor(void)
{
	
	return (OLE_COLOR) GetStartBorderColor().ToCOLORREF();
}

void odl_SetStartBorderColor(OLE_COLOR Value)
{
	
	Gdiplus::Color oColor;
	oColor.SetFromCOLORREF(g_TranslateColor(Value));
	SetStartBorderColor(oColor);
}

OLE_COLOR odl_GetStartFillColor(void)
{
	
	return (OLE_COLOR) GetStartFillColor().ToCOLORREF();
}

void odl_SetStartFillColor(OLE_COLOR Value)
{
	
	Gdiplus::Color oColor;
	oColor.SetFromCOLORREF(g_TranslateColor(Value));
	SetStartFillColor(oColor);
}

IDispatch* odl_GetStartImage(void)
{
	
	return GetStartImage()->GetIDispatch(TRUE);
}

IDispatch* odl_GetMiddleImage(void)
{
	
	return GetMiddleImage()->GetIDispatch(TRUE);
}

IDispatch* odl_GetEndImage(void)
{
	
	return GetEndImage()->GetIDispatch(TRUE);
}

LONG odl_GetStartShapeIndex(void)
{
	
	return GetStartShapeIndex();
}

void odl_SetStartShapeIndex(LONG Value)
{
	
	SetStartShapeIndex(Value);
}

LONG odl_GetEndShapeIndex(void)
{
	
	return GetEndShapeIndex();
}

void odl_SetEndShapeIndex(LONG Value)
{
	
	SetEndShapeIndex(Value);
}

BSTR odl_GetStartImageTag(void)
{
	
	return GetStartImageTag().AllocSysString();
}

void odl_SetStartImageTag(LPCTSTR Value)
{
	
	SetStartImageTag(Value);
}

BSTR odl_GetMiddleImageTag(void)
{
	
	return GetMiddleImageTag().AllocSysString();
}

void odl_SetMiddleImageTag(LPCTSTR Value)
{
	
	SetMiddleImageTag(Value);
}

BSTR odl_GetEndImageTag(void)
{
	
	return GetEndImageTag().AllocSysString();
}

void odl_SetEndImageTag(LPCTSTR Value)
{
	
	SetEndImageTag(Value);
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