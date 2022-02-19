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
#include "clsMilestoneStyle.h"

IMPLEMENT_DYNCREATE(clsMilestoneStyle, CCmdTarget)

//{00A53440-9950-4F4B-9385-ED168260CC41}
static const IID IID_IclsMilestoneStyle = { 0x00A53440, 0x9950, 0x4F4B, { 0x93, 0x85, 0xED, 0x16, 0x82, 0x60, 0xCC, 0x41} };

//{FB4EB87B-10CA-4683-AF38-F975FCFBD3A3}
IMPLEMENT_OLECREATE_FLAGS(clsMilestoneStyle, "AGVC.clsMilestoneStyle", afxRegApartmentThreading, 0xFB4EB87B, 0x10CA, 0x4683, 0xAF, 0x38, 0xF9, 0x75, 0xFC, 0xFB, 0xD3, 0xA3)

BEGIN_DISPATCH_MAP(clsMilestoneStyle, CCmdTarget)
	DISP_PROPERTY_EX_ID(clsMilestoneStyle, "BorderColor", 1, odl_GetBorderColor, odl_SetBorderColor, VT_COLOR) 
	DISP_PROPERTY_EX_ID(clsMilestoneStyle, "FillColor", 2, odl_GetFillColor, odl_SetFillColor, VT_COLOR) 
	DISP_PROPERTY_EX_ID(clsMilestoneStyle, "ShapeIndex", 3, odl_GetShapeIndex, odl_SetShapeIndex, VT_I4) 
	DISP_FUNCTION_ID(clsMilestoneStyle, "GetXML", 4, odl_GetXML, VT_BSTR, VTS_NONE) 
	DISP_FUNCTION_ID(clsMilestoneStyle, "SetXML", 5, odl_SetXML, VT_EMPTY, VTS_BSTR)
	DISP_PROPERTY_EX_ID(clsMilestoneStyle, "Image", 6, odl_GetImage, SetNotSupported, VT_DISPATCH)
	DISP_PROPERTY_EX_ID(clsMilestoneStyle, "ImageTag", 7, odl_GetImageTag, odl_SetImageTag, VT_BSTR)
END_DISPATCH_MAP()

BEGIN_INTERFACE_MAP(clsMilestoneStyle, CCmdTarget)
	INTERFACE_PART(clsMilestoneStyle, IID_IclsMilestoneStyle, Dispatch)
END_INTERFACE_MAP()

BEGIN_MESSAGE_MAP(clsMilestoneStyle, CCmdTarget)
END_MESSAGE_MAP()

clsMilestoneStyle::clsMilestoneStyle()
{
	AfxMessageBox(_T("Invalid constructor has been invoked for clsMilestoneStyle. Please inform support@sourcecodestore.com of this error"), MB_OK | MB_ICONSTOP | MB_APPLMODAL);
}

clsMilestoneStyle::clsMilestoneStyle(CActiveGanttVCCtl* oControl)
{
	EnableAutomation();
	AfxOleLockApp();
	mp_oImage = new clsImage();
	Clear();
}

clsMilestoneStyle::~clsMilestoneStyle()
{
	delete mp_oImage;
	mp_oImage = NULL;
	AfxOleUnlockApp();
}

void clsMilestoneStyle::OnFinalRelease()
{
	CCmdTarget::OnFinalRelease();
}

Gdiplus::Color clsMilestoneStyle::GetBorderColor(void)
{
	return mp_clrBorderColor;
}

void clsMilestoneStyle::SetBorderColor(Gdiplus::Color Value)
{
	mp_clrBorderColor = Value;
}

Gdiplus::Color clsMilestoneStyle::GetFillColor(void)
{
	return mp_clrFillColor;
}

void clsMilestoneStyle::SetFillColor(Gdiplus::Color Value)
{
	mp_clrFillColor = Value;
}

LONG clsMilestoneStyle::GetShapeIndex(void)
{
	return mp_yShapeIndex;
}

void clsMilestoneStyle::SetShapeIndex(LONG Value)
{
	mp_yShapeIndex = Value;
}

clsImage* clsMilestoneStyle::GetImage(void)
{
	return mp_oImage;
}

CString clsMilestoneStyle::GetImageTag(void)
{
	return mp_sImageTag;
}

void clsMilestoneStyle::SetImageTag(CString Value)
{
	mp_sImageTag = Value;
}

void clsMilestoneStyle::Finalize(void)
{
}

CString clsMilestoneStyle::GetXML(void)
{
	clsXML oXML(mp_oControl, "MilestoneStyle");
	oXML.InitializeWriter();
	oXML.WriteProperty("BorderColor", mp_clrBorderColor);
	oXML.WriteProperty("FillColor", mp_clrFillColor);
	oXML.WriteProperty("ShapeIndex", mp_yShapeIndex);
	oXML.WriteProperty("Image", *mp_oImage);
	oXML.WriteProperty("ImageTag", mp_sImageTag);
	return oXML.GetXML();
}

void clsMilestoneStyle::SetXML(CString sXML)
{
	clsXML oXML(mp_oControl, "MilestoneStyle");
	oXML.SetXML(sXML);
	oXML.InitializeReader();
	oXML.ReadProperty("BorderColor", mp_clrBorderColor);
	oXML.ReadProperty("FillColor", mp_clrFillColor);
	oXML.ReadProperty("ShapeIndex", mp_yShapeIndex);
	oXML.ReadProperty("Image", *mp_oImage);
	oXML.ReadProperty("ImageTag", mp_sImageTag);
}

void clsMilestoneStyle::Clear()
{
	mp_clrBorderColor = Color::Black;
	mp_clrFillColor = Color::Black;
	mp_yShapeIndex = FT_NONE;
	
	mp_sImageTag = _T("");
}

void clsMilestoneStyle::Clone(clsMilestoneStyle* oClone)
{
    oClone->SetBorderColor(mp_clrBorderColor);
    oClone->SetFillColor(mp_clrFillColor);
    oClone->SetShapeIndex(mp_yShapeIndex);
    mp_oImage->Clone(oClone->mp_oImage);
    oClone->SetImageTag(mp_sImageTag);
}