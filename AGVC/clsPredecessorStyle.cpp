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
#include "clsPredecessorStyle.h"

IMPLEMENT_DYNCREATE(clsPredecessorStyle, CCmdTarget)

//{4680FBAC-FBC4-4592-B65F-1B3FD8FEC0AD}
static const IID IID_IclsPredecessorStyle = { 0x4680FBAC, 0xFBC4, 0x4592, { 0xB6, 0x5F, 0x1B, 0x3F, 0xD8, 0xFE, 0xC0, 0xAD} };

//{21D83EC2-2378-460F-9BF1-39BD57A0374F}
IMPLEMENT_OLECREATE_FLAGS(clsPredecessorStyle, "AGVC.clsPredecessorStyle", afxRegApartmentThreading, 0x21D83EC2, 0x2378, 0x460F, 0x9B, 0xF1, 0x39, 0xBD, 0x57, 0xA0, 0x37, 0x4F)

BEGIN_DISPATCH_MAP(clsPredecessorStyle, CCmdTarget)
	DISP_PROPERTY_EX_ID(clsPredecessorStyle, "LineColor", 1, odl_GetLineColor, odl_SetLineColor, VT_COLOR) 
	DISP_PROPERTY_EX_ID(clsPredecessorStyle, "XOffset", 2, odl_GetXOffset, odl_SetXOffset, VT_I4) 
	DISP_PROPERTY_EX_ID(clsPredecessorStyle, "YOffset", 3, odl_GetYOffset, odl_SetYOffset, VT_I4) 
	DISP_PROPERTY_EX_ID(clsPredecessorStyle, "LineWidth", 4, odl_GetLineWidth, odl_SetLineWidth, VT_I4) 
	DISP_PROPERTY_EX_ID(clsPredecessorStyle, "LineStyle", 5, odl_GetLineStyle, odl_SetLineStyle, VT_I4) 
	DISP_PROPERTY_EX_ID(clsPredecessorStyle, "ArrowSize", 6, odl_GetArrowSize, odl_SetArrowSize, VT_I4) 
	DISP_FUNCTION_ID(clsPredecessorStyle, "GetXML", 7, odl_GetXML, VT_BSTR, VTS_NONE) 
	DISP_FUNCTION_ID(clsPredecessorStyle, "SetXML", 8, odl_SetXML, VT_EMPTY, VTS_BSTR) 
END_DISPATCH_MAP()

BEGIN_INTERFACE_MAP(clsPredecessorStyle, CCmdTarget)
	INTERFACE_PART(clsPredecessorStyle, IID_IclsPredecessorStyle, Dispatch)
END_INTERFACE_MAP()

BEGIN_MESSAGE_MAP(clsPredecessorStyle, CCmdTarget)
END_MESSAGE_MAP()

clsPredecessorStyle::clsPredecessorStyle()
{
	AfxMessageBox(_T("Invalid constructor has been invoked for clsPredecessorStyle. Please inform support@sourcecodestore.com of this error"), MB_OK | MB_ICONSTOP | MB_APPLMODAL);
}

clsPredecessorStyle::clsPredecessorStyle(CActiveGanttVCCtl* oControl)
{
	EnableAutomation();
	AfxOleLockApp();
	mp_oControl = oControl;
	Clear();
}

clsPredecessorStyle::~clsPredecessorStyle()
{
	AfxOleUnlockApp();
}

void clsPredecessorStyle::OnFinalRelease()
{
	CCmdTarget::OnFinalRelease();
}

Gdiplus::Color clsPredecessorStyle::GetLineColor(void)
{
	return mp_clrLineColor;
}

void clsPredecessorStyle::SetLineColor(Gdiplus::Color Value)
{
	mp_clrLineColor = Value;
}

LONG clsPredecessorStyle::GetXOffset(void)
{
	return mp_lXOffset;
}

void clsPredecessorStyle::SetXOffset(LONG Value)
{
	mp_lXOffset = Value;
}

LONG clsPredecessorStyle::GetYOffset(void)
{
	return mp_lYOffset;
}

void clsPredecessorStyle::SetYOffset(LONG Value)
{
	mp_lYOffset = Value;
}

LONG clsPredecessorStyle::GetLineWidth(void)
{
	return mp_lLineWidth;
}

void clsPredecessorStyle::SetLineWidth(LONG Value)
{
	mp_lLineWidth = Value;
}

LONG clsPredecessorStyle::GetLineStyle(void)
{
	return mp_yLineStyle;
}

void clsPredecessorStyle::SetLineStyle(LONG Value)
{
	mp_yLineStyle = Value;
}

LONG clsPredecessorStyle::GetArrowSize(void)
{
	return mp_lArrowSize;
}

void clsPredecessorStyle::SetArrowSize(LONG Value)
{
	if ((Value < 1)) 
	{
		Value = 1;
	}
	mp_lArrowSize = Value;
}

CString clsPredecessorStyle::GetXML(void)
{
	clsXML oXML(mp_oControl, "PredecessorStyle");
	oXML.InitializeWriter();
	oXML.WriteProperty("ArrowSize", mp_lArrowSize);
	oXML.WriteProperty("LineColor", mp_clrLineColor);
	oXML.WriteProperty("LineStyle", mp_yLineStyle);
	oXML.WriteProperty("LineWidth", mp_lLineWidth);
	oXML.WriteProperty("XOffset", mp_lXOffset);
	oXML.WriteProperty("YOffset", mp_lYOffset);
	return oXML.GetXML();
}

void clsPredecessorStyle::SetXML(CString sXML)
{
	clsXML oXML(mp_oControl, "PredecessorStyle");
	oXML.SetXML(sXML);
	oXML.InitializeReader();
	oXML.ReadProperty("ArrowSize", mp_lArrowSize);
	oXML.ReadProperty("LineColor", mp_clrLineColor);
	oXML.ReadProperty("LineStyle", mp_yLineStyle);
	oXML.ReadProperty("LineWidth", mp_lLineWidth);
	oXML.ReadProperty("XOffset", mp_lXOffset);
	oXML.ReadProperty("YOffset", mp_lYOffset);
}

void clsPredecessorStyle::Clear()
{
	mp_lArrowSize = 3;
	mp_yLineStyle = LDS_SOLID;
	mp_lLineWidth = 1;
	mp_lXOffset = 10;
	mp_lYOffset = 10;
	mp_clrLineColor = Color::Black;
}

void clsPredecessorStyle::Clone(clsPredecessorStyle* oClone)
{
    oClone->SetArrowSize(mp_lArrowSize);
    oClone->SetLineColor(mp_clrLineColor);
    oClone->SetLineStyle(mp_yLineStyle);
    oClone->SetLineWidth(mp_lLineWidth);
    oClone->SetXOffset(mp_lXOffset);
    oClone->SetYOffset(mp_lYOffset);
}