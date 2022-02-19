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
#include "ControlTemplateGradient.h"

IMPLEMENT_DYNCREATE(ControlTemplateGradient, CCmdTarget)

//{C6257C4E-4DE3-406D-AECB-11D8183BD5DF}
static const IID IID_IControlTemplateGradient = { 0xC6257C4E, 0x4DE3, 0x406D, { 0xAE, 0xCB, 0x11, 0xD8, 0x18, 0x3B, 0xD5, 0xDF} };

//{CAFA3808-FDAA-4AF4-8AD8-63A17DA79566}
IMPLEMENT_OLECREATE_FLAGS(ControlTemplateGradient, "AGVC.ControlTemplateGradient", afxRegApartmentThreading, 0xCAFA3808, 0xFDAA, 0x4AF4, 0x8A, 0xD8, 0x63, 0xA1, 0x7D, 0xA7, 0x95, 0x66)

BEGIN_INTERFACE_MAP(ControlTemplateGradient, CCmdTarget)
	INTERFACE_PART(ControlTemplateGradient, IID_IControlTemplateGradient, Dispatch)
END_INTERFACE_MAP()

BEGIN_MESSAGE_MAP(ControlTemplateGradient, CCmdTarget)
END_MESSAGE_MAP()


BEGIN_DISPATCH_MAP(ControlTemplateGradient, CCmdTarget)
	DISP_PROPERTY_EX_ID(ControlTemplateGradient, "ControlBorderColor", 1, odl_GetControlBorderColor, odl_SetControlBorderColor, VT_COLOR)
	DISP_PROPERTY_EX_ID(ControlTemplateGradient, "HeadingBorderColor", 2, odl_GetHeadingBorderColor, odl_SetHeadingBorderColor, VT_COLOR)
	DISP_PROPERTY_EX_ID(ControlTemplateGradient, "HeadingForeColor", 3, odl_GetHeadingForeColor, odl_SetHeadingForeColor, VT_COLOR)
	DISP_PROPERTY_EX_ID(ControlTemplateGradient, "RowForeColor", 4, odl_GetRowForeColor, odl_SetRowForeColor, VT_COLOR)
	DISP_PROPERTY_EX_ID(ControlTemplateGradient, "GradientFillMode", 5, odl_GetGradientFillMode, odl_SetGradientFillMode, VT_I4)
	DISP_PROPERTY_EX_ID(ControlTemplateGradient, "StartGradientColor", 6, odl_GetStartGradientColor, odl_SetStartGradientColor, VT_COLOR)
	DISP_PROPERTY_EX_ID(ControlTemplateGradient, "EndGradientColor", 7, odl_GetEndGradientColor, odl_SetEndGradientColor, VT_COLOR)
	DISP_PROPERTY_EX_ID(ControlTemplateGradient, "TreeviewColor", 8, odl_GetTreeviewColor, odl_SetTreeviewColor, VT_COLOR)
	DISP_PROPERTY_EX_ID(ControlTemplateGradient, "RowBorderColor", 9, odl_GetRowBorderColor, odl_SetRowBorderColor, VT_COLOR)
END_DISPATCH_MAP()

ControlTemplateGradient::ControlTemplateGradient()
{
	EnableAutomation();	
	AfxOleLockApp();
    ControlBorderColor = Color::MakeARGB(255, 100, 145, 204);
    HeadingBorderColor = Color::MakeARGB(255, 197, 206, 216);
    HeadingForeColor = Color::Black;
    RowForeColor = Color::Black;
    GradientFillMode = GDT_VERTICAL;
    StartGradientColor = Color::MakeARGB(255, 179, 206, 235);
    EndGradientColor = Color::MakeARGB(255, 161, 193, 232);
    TreeviewColor = Color::MakeARGB(255, 100, 145, 204);
    RowBorderColor = Color::MakeARGB(255, 192, 192, 192);
}

ControlTemplateGradient::~ControlTemplateGradient()
{
	AfxOleUnlockApp();
}

void ControlTemplateGradient::OnFinalRelease()
{
	CCmdTarget::OnFinalRelease();
}

