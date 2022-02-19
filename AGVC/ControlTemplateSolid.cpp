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
#include "ControlTemplateSolid.h"

IMPLEMENT_DYNCREATE(ControlTemplateSolid, CCmdTarget)

//{76B6BFCB-3922-451C-ABDB-9E9DC26CD8C4}
static const IID IID_IControlTemplateSolid = { 0x76B6BFCB, 0x3922, 0x451C, { 0xAB, 0xDB, 0x9E, 0x9D, 0xC2, 0x6C, 0xD8, 0xC4} };

//{2678DCAD-0753-40E7-8207-852D1B0AAB35}
IMPLEMENT_OLECREATE_FLAGS(ControlTemplateSolid, "AGVC.ControlTemplateSolid", afxRegApartmentThreading, 0x2678DCAD, 0x0753, 0x40E7, 0x82, 0x07, 0x85, 0x2D, 0x1B, 0x0A, 0xAB, 0x35)

BEGIN_INTERFACE_MAP(ControlTemplateSolid, CCmdTarget)
	INTERFACE_PART(ControlTemplateSolid, IID_IControlTemplateSolid, Dispatch)
END_INTERFACE_MAP()

BEGIN_MESSAGE_MAP(ControlTemplateSolid, CCmdTarget)
END_MESSAGE_MAP()


BEGIN_DISPATCH_MAP(ControlTemplateSolid, CCmdTarget)
	DISP_PROPERTY_EX_ID(ControlTemplateSolid, "ControlBorderColor", 1, odl_GetControlBorderColor, odl_SetControlBorderColor, VT_COLOR)
	DISP_PROPERTY_EX_ID(ControlTemplateSolid, "HeadingBorderColor", 2, odl_GetHeadingBorderColor, odl_SetHeadingBorderColor, VT_COLOR) 
	DISP_PROPERTY_EX_ID(ControlTemplateSolid, "HeadingForeColor", 3, odl_GetHeadingForeColor, odl_SetHeadingForeColor, VT_COLOR) 
	DISP_PROPERTY_EX_ID(ControlTemplateSolid, "RowForeColor", 4, odl_GetRowForeColor, odl_SetRowForeColor, VT_COLOR) 
	DISP_PROPERTY_EX_ID(ControlTemplateSolid, "HeadingBackColor", 5, odl_GetHeadingBackColor, odl_SetHeadingBackColor, VT_COLOR)
	DISP_PROPERTY_EX_ID(ControlTemplateSolid, "TreeviewColor", 6, odl_GetTreeviewColor, odl_SetTreeviewColor, VT_COLOR) 
	DISP_PROPERTY_EX_ID(ControlTemplateSolid, "RowBorderColor", 7, odl_GetRowBorderColor, odl_SetRowBorderColor, VT_COLOR) 
	DISP_PROPERTY_EX_ID(ControlTemplateSolid, "TimeBlockBackColor", 8, odl_GetTimeBlockBackColor, odl_SetTimeBlockBackColor, VT_COLOR) 
END_DISPATCH_MAP()

ControlTemplateSolid::ControlTemplateSolid()
{
	EnableAutomation();	
	AfxOleLockApp();
    ControlBorderColor = Color::MakeARGB(255, 100, 145, 204);
    HeadingBorderColor = Color::MakeARGB(255, 197, 206, 216);
    HeadingForeColor = Color::Black;
    RowForeColor = Color::Black;
    HeadingBackColor = Color::White;
    TreeviewColor = Color::MakeARGB(255, 100, 145, 204);
    RowBorderColor = Color::MakeARGB(255, 192, 192, 192);
    TimeBlockBackColor = Color::MakeARGB(255, 192, 192, 192);
}

ControlTemplateSolid::~ControlTemplateSolid()
{
	AfxOleUnlockApp();
}


void ControlTemplateSolid::OnFinalRelease()
{
	CCmdTarget::OnFinalRelease();
}