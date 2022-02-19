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


class clsMilestoneStyle : public CCmdTarget
{
	DECLARE_DYNCREATE(clsMilestoneStyle)
	clsMilestoneStyle();

public:
	CActiveGanttVCCtl* mp_oControl;

	clsMilestoneStyle(CActiveGanttVCCtl* oControl);
	virtual ~clsMilestoneStyle();
	virtual void OnFinalRelease();

// Member Variables
public:
	Gdiplus::Color mp_clrBorderColor;
	Gdiplus::Color mp_clrFillColor;
	LONG mp_yShapeIndex;
	clsImage* mp_oImage;
	CString mp_sImageTag;

//Internal Properties (Extensions to ODL)
public:
	Gdiplus::Color GetBorderColor(void);
	void SetBorderColor(Gdiplus::Color Value);
	Gdiplus::Color GetFillColor(void);
	void SetFillColor(Gdiplus::Color Value);
	LONG GetShapeIndex(void);
	void SetShapeIndex(LONG Value);
	clsImage* GetImage(void);
	CString GetImageTag(void);
	void SetImageTag(CString Value);

//Internal Properties
public:

//Internal Methods
public:
	void Finalize(void);
	CString GetXML(void);
	void SetXML(CString sXML);
	void Clear();
	void Clone(clsMilestoneStyle* oClone);

protected:
	DECLARE_DISPATCH_MAP()
	DECLARE_MESSAGE_MAP()
	DECLARE_OLECREATE(clsMilestoneStyle)
	DECLARE_INTERFACE_MAP()

OLE_COLOR odl_GetBorderColor(void)
{
	
	return (OLE_COLOR) GetBorderColor().ToCOLORREF();
}

void odl_SetBorderColor(OLE_COLOR Value)
{
	
	Gdiplus::Color oColor;
	oColor.SetFromCOLORREF(g_TranslateColor(Value));
	SetBorderColor(oColor);
}

OLE_COLOR odl_GetFillColor(void)
{
	
	return (OLE_COLOR) GetFillColor().ToCOLORREF();
}

void odl_SetFillColor(OLE_COLOR Value)
{
	
	Gdiplus::Color oColor;
	oColor.SetFromCOLORREF(g_TranslateColor(Value));
	SetFillColor(oColor);
}

LONG odl_GetShapeIndex(void)
{
	
	return GetShapeIndex();
}

void odl_SetShapeIndex(LONG Value)
{
	
	SetShapeIndex(Value);
}

IDispatch* odl_GetImage(void)
{
	
	return GetImage()->GetIDispatch(TRUE);
}

BSTR odl_GetImageTag(void)
{
	
	return GetImageTag().AllocSysString();
}

void odl_SetImageTag(LPCTSTR Value)
{
	
	SetImageTag(Value);
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