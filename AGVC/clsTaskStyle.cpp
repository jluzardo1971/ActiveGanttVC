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
#include "stdafx.h"
#include "ActiveGanttVC.h"
#include "ActiveGanttVCCtl.h"
#include "clsXML.h"
#include "clsTaskStyle.h"

IMPLEMENT_DYNCREATE(clsTaskStyle, CCmdTarget)

//{94D3203F-EC64-42EC-A418-1527E82AE807}
static const IID IID_IclsTaskStyle = { 0x94D3203F, 0xEC64, 0x42EC, { 0xA4, 0x18, 0x15, 0x27, 0xE8, 0x2A, 0xE8, 0x07} };

//{C07FF05F-B755-42CA-95EA-9BF7FE441B63}
IMPLEMENT_OLECREATE_FLAGS(clsTaskStyle, "AGVC.clsTaskStyle", afxRegApartmentThreading, 0xC07FF05F, 0xB755, 0x42CA, 0x95, 0xEA, 0x9B, 0xF7, 0xFE, 0x44, 0x1B, 0x63)

BEGIN_DISPATCH_MAP(clsTaskStyle, CCmdTarget)
	DISP_PROPERTY_EX_ID(clsTaskStyle, "EndBorderColor", 1, odl_GetEndBorderColor, odl_SetEndBorderColor, VT_COLOR) 
	DISP_PROPERTY_EX_ID(clsTaskStyle, "EndFillColor", 2, odl_GetEndFillColor, odl_SetEndFillColor, VT_COLOR) 
	DISP_PROPERTY_EX_ID(clsTaskStyle, "StartBorderColor", 3, odl_GetStartBorderColor, odl_SetStartBorderColor, VT_COLOR) 
	DISP_PROPERTY_EX_ID(clsTaskStyle, "StartFillColor", 4, odl_GetStartFillColor, odl_SetStartFillColor, VT_COLOR) 
	DISP_PROPERTY_EX_ID(clsTaskStyle, "StartImage", 5, odl_GetStartImage, SetNotSupported, VT_DISPATCH) 
	DISP_PROPERTY_EX_ID(clsTaskStyle, "MiddleImage", 6, odl_GetMiddleImage, SetNotSupported, VT_DISPATCH) 
	DISP_PROPERTY_EX_ID(clsTaskStyle, "EndImage", 7, odl_GetEndImage, SetNotSupported, VT_DISPATCH) 
	DISP_PROPERTY_EX_ID(clsTaskStyle, "StartShapeIndex", 8, odl_GetStartShapeIndex, odl_SetStartShapeIndex, VT_I4) 
	DISP_PROPERTY_EX_ID(clsTaskStyle, "EndShapeIndex", 9, odl_GetEndShapeIndex, odl_SetEndShapeIndex, VT_I4) 
	DISP_FUNCTION_ID(clsTaskStyle, "GetXML", 10, odl_GetXML, VT_BSTR, VTS_NONE) 
	DISP_FUNCTION_ID(clsTaskStyle, "SetXML", 11, odl_SetXML, VT_EMPTY, VTS_BSTR)
	DISP_PROPERTY_EX_ID(clsTaskStyle, "StartImageTag", 12, odl_GetStartImageTag, odl_SetStartImageTag, VT_BSTR)
	DISP_PROPERTY_EX_ID(clsTaskStyle, "MiddleImageTag", 13, odl_GetMiddleImageTag, odl_SetMiddleImageTag, VT_BSTR)
	DISP_PROPERTY_EX_ID(clsTaskStyle, "EndImageTag", 14, odl_GetEndImageTag, odl_SetEndImageTag, VT_BSTR)
END_DISPATCH_MAP()

BEGIN_INTERFACE_MAP(clsTaskStyle, CCmdTarget)
	INTERFACE_PART(clsTaskStyle, IID_IclsTaskStyle, Dispatch)
END_INTERFACE_MAP()

BEGIN_MESSAGE_MAP(clsTaskStyle, CCmdTarget)
END_MESSAGE_MAP()

clsTaskStyle::clsTaskStyle()
{
	AfxMessageBox(_T("Invalid constructor has been invoked for clsTaskStyle. Please inform support@sourcecodestore.com of this error"), MB_OK | MB_ICONSTOP | MB_APPLMODAL);
}

clsTaskStyle::clsTaskStyle(CActiveGanttVCCtl* oControl)
{
	EnableAutomation();
	AfxOleLockApp();
	mp_oControl = oControl;
	mp_oEndImage = new clsImage();
	mp_oMiddleImage = new clsImage();
	mp_oStartImage = new clsImage();
	Clear();
}

clsTaskStyle::~clsTaskStyle()
{
	delete mp_oEndImage;
	mp_oEndImage = NULL;
	delete mp_oMiddleImage;
	mp_oMiddleImage = NULL;
	delete mp_oStartImage;
	mp_oStartImage = NULL;
	AfxOleUnlockApp();
}

void clsTaskStyle::OnFinalRelease()
{
	CCmdTarget::OnFinalRelease();
}

Gdiplus::Color clsTaskStyle::GetEndBorderColor(void)
{
	return mp_clrEndBorderColor;
}

void clsTaskStyle::SetEndBorderColor(Gdiplus::Color Value)
{
	mp_clrEndBorderColor = Value;
}

Gdiplus::Color clsTaskStyle::GetEndFillColor(void)
{
	return mp_clrEndFillColor;
}

void clsTaskStyle::SetEndFillColor(Gdiplus::Color Value)
{
	mp_clrEndFillColor = Value;
}

Gdiplus::Color clsTaskStyle::GetStartBorderColor(void)
{
	return mp_clrStartBorderColor;
}

void clsTaskStyle::SetStartBorderColor(Gdiplus::Color Value)
{
	mp_clrStartBorderColor = Value;
}

Gdiplus::Color clsTaskStyle::GetStartFillColor(void)
{
	return mp_clrStartFillColor;
}

void clsTaskStyle::SetStartFillColor(Gdiplus::Color Value)
{
	mp_clrStartFillColor = Value;
}

clsImage* clsTaskStyle::GetStartImage(void)
{
	return mp_oStartImage;
}

clsImage* clsTaskStyle::GetMiddleImage(void)
{
	return mp_oMiddleImage;
}

clsImage* clsTaskStyle::GetEndImage(void)
{
	return mp_oEndImage;
}

LONG clsTaskStyle::GetStartShapeIndex(void)
{
	return mp_yStartShapeIndex;
}

void clsTaskStyle::SetStartShapeIndex(LONG Value)
{
	mp_yStartShapeIndex = Value;
}

LONG clsTaskStyle::GetEndShapeIndex(void)
{
	return mp_yEndShapeIndex;
}

void clsTaskStyle::SetEndShapeIndex(LONG Value)
{
	mp_yEndShapeIndex = Value;
}

CString clsTaskStyle::GetStartImageTag(void)
{
	return mp_sStartImageTag;
}

void clsTaskStyle::SetStartImageTag(CString Value)
{
	mp_sStartImageTag = Value;
}

CString clsTaskStyle::GetMiddleImageTag(void)
{
	return mp_sMiddleImageTag;
}

void clsTaskStyle::SetMiddleImageTag(CString Value)
{
	mp_sMiddleImageTag = Value;
}

CString clsTaskStyle::GetEndImageTag(void)
{
	return mp_sEndImageTag;
}

void clsTaskStyle::SetEndImageTag(CString Value)
{
	mp_sEndImageTag = Value;
}

CString clsTaskStyle::GetXML(void)
{
	clsXML oXML(mp_oControl, "TaskStyle");
	oXML.InitializeWriter();
	oXML.WriteProperty("EndBorderColor", mp_clrEndBorderColor);
	oXML.WriteProperty("EndFillColor", mp_clrEndFillColor);
	oXML.WriteProperty("EndImage", *mp_oEndImage);
	oXML.WriteProperty("EndShapeIndex", mp_yEndShapeIndex);
	oXML.WriteProperty("MiddleImage", *mp_oMiddleImage);
	oXML.WriteProperty("StartBorderColor", mp_clrStartBorderColor);
	oXML.WriteProperty("StartFillColor", mp_clrStartFillColor);
	oXML.WriteProperty("StartImage", *mp_oStartImage);
	oXML.WriteProperty("StartShapeIndex", mp_yStartShapeIndex);
	oXML.WriteProperty("StartImageTag", mp_sStartImageTag);
	oXML.WriteProperty("MiddleImageTag", mp_sMiddleImageTag);
	oXML.WriteProperty("EndImageTag", mp_sEndImageTag);
	return oXML.GetXML();
}

void clsTaskStyle::SetXML(CString sXML)
{
	clsXML oXML(mp_oControl, "TaskStyle");
	oXML.SetXML(sXML);
	oXML.InitializeReader();
	oXML.ReadProperty("EndBorderColor", mp_clrEndBorderColor);
	oXML.ReadProperty("EndFillColor", mp_clrEndFillColor);
	oXML.ReadProperty("EndImage", *mp_oEndImage);
	oXML.ReadProperty("EndShapeIndex", mp_yEndShapeIndex);
	oXML.ReadProperty("MiddleImage", *mp_oMiddleImage);
	oXML.ReadProperty("StartBorderColor", mp_clrStartBorderColor);
	oXML.ReadProperty("StartFillColor", mp_clrStartFillColor);
	oXML.ReadProperty("StartImage", *mp_oStartImage);
	oXML.ReadProperty("StartShapeIndex", mp_yStartShapeIndex);
	oXML.ReadProperty("StartImageTag", mp_sStartImageTag);
	oXML.ReadProperty("MiddleImageTag", mp_sMiddleImageTag);
	oXML.ReadProperty("EndImageTag", mp_sEndImageTag);
}

void clsTaskStyle::Clear()
{
	mp_clrEndBorderColor = Color::Black;
	mp_clrEndFillColor = Color::Black;
	mp_clrStartBorderColor = Color::Black;
	mp_clrStartFillColor = Color::Black;
	mp_yEndShapeIndex = FT_NONE;
	mp_yStartShapeIndex = FT_NONE;
	mp_sStartImageTag = _T("");
	mp_sMiddleImageTag = _T("");
	mp_sEndImageTag = _T("");
}

void clsTaskStyle::Clone(clsTaskStyle* oClone)
{
    oClone->SetEndBorderColor(mp_clrEndBorderColor);
    oClone->SetEndFillColor(mp_clrEndFillColor);
    mp_oEndImage->Clone(oClone->mp_oEndImage);
    oClone->SetEndShapeIndex(mp_yEndShapeIndex);
    mp_oMiddleImage->Clone(oClone->mp_oMiddleImage);
    oClone->SetStartBorderColor(mp_clrStartBorderColor);
    oClone->SetStartFillColor(mp_clrStartFillColor);
    mp_oStartImage->Clone(oClone->mp_oStartImage);
    oClone->SetStartShapeIndex(mp_yStartShapeIndex);
    oClone->SetStartImageTag(mp_sStartImageTag);
    oClone->SetMiddleImageTag(mp_sMiddleImageTag);
    oClone->SetEndImageTag(mp_sEndImageTag);
}