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
#include "clsGrid.h"



IMPLEMENT_DYNCREATE(clsGrid, CCmdTarget)

//{A7C516C7-44C6-4DC2-BF5F-A50E9CF27982}
static const IID IID_IclsGrid = { 0xA7C516C7, 0x44C6, 0x4DC2, { 0xBF, 0x5F, 0xA5, 0x0E, 0x9C, 0xF2, 0x79, 0x82} };

//{A0CA2CF9-1217-47CB-ABA8-CB7B60E56802}
IMPLEMENT_OLECREATE_FLAGS(clsGrid, "AGVC.clsGrid", afxRegApartmentThreading, 0xA0CA2CF9, 0x1217, 0x47CB, 0xAB, 0xA8, 0xCB, 0x7B, 0x60, 0xE5, 0x68, 0x02)

BEGIN_DISPATCH_MAP(clsGrid, CCmdTarget)
	DISP_PROPERTY_EX_ID(clsGrid, "VerticalLines", 1, odl_GetVerticalLines, odl_SetVerticalLines, VT_BOOL) 
	DISP_PROPERTY_EX_ID(clsGrid, "SnapToGrid", 2, odl_GetSnapToGrid, odl_SetSnapToGrid, VT_BOOL) 
	DISP_PROPERTY_EX_ID(clsGrid, "SnapToGridOnSelection", 3, odl_GetSnapToGridOnSelection, odl_SetSnapToGridOnSelection, VT_BOOL) 
	DISP_PROPERTY_EX_ID(clsGrid, "Color", 4, odl_GetColor, odl_SetColor, VT_COLOR) 
	DISP_PROPERTY_EX_ID(clsGrid, "Interval", 5, odl_GetInterval, odl_SetInterval, VT_I4)
	DISP_PROPERTY_EX_ID(clsGrid, "Factor", 6, odl_GetFactor, odl_SetFactor, VT_I4)
	DISP_FUNCTION_ID(clsGrid, "GetXML", 7, odl_GetXML, VT_BSTR, VTS_NONE) 
	DISP_FUNCTION_ID(clsGrid, "SetXML", 8, odl_SetXML, VT_EMPTY, VTS_BSTR)
END_DISPATCH_MAP()

BEGIN_INTERFACE_MAP(clsGrid, CCmdTarget)
	INTERFACE_PART(clsGrid, IID_IclsGrid, Dispatch)
END_INTERFACE_MAP()

BEGIN_MESSAGE_MAP(clsGrid, CCmdTarget)
END_MESSAGE_MAP()

clsGrid::clsGrid(CActiveGanttVCCtl* oControl, clsTimeLine* oTimeLine)
{
	EnableAutomation();
	AfxOleLockApp();
	mp_oControl = oControl;
	mp_oTimeLine = oTimeLine;
	Clear();
}

clsGrid::clsGrid()
{
	AfxMessageBox(_T("Invalid constructor has been invoked for clsGrid. Please inform support@sourcecodestore.com of this error"), MB_OK | MB_ICONSTOP | MB_APPLMODAL);
}

clsGrid::~clsGrid()
{
	AfxOleUnlockApp();
}

void clsGrid::OnFinalRelease()
{
	CCmdTarget::OnFinalRelease();
}

BOOL clsGrid::GetVerticalLines(void)
{
	return mp_bVerticalLines;
}

void clsGrid::SetVerticalLines(BOOL Value)
{
	mp_bVerticalLines = Value;
}

BOOL clsGrid::GetSnapToGrid(void)
{
	return mp_bSnapToGrid;
}

void clsGrid::SetSnapToGrid(BOOL Value)
{
	mp_bSnapToGrid = Value;
}

BOOL clsGrid::GetSnapToGridOnSelection(void)
{
	return mp_bSnapToGridOnSelection;
}

void clsGrid::SetSnapToGridOnSelection(BOOL Value)
{
	mp_bSnapToGridOnSelection = Value;
}

Gdiplus::Color clsGrid::GetColor(void)
{
	return mp_clrColor;
}

void clsGrid::SetColor(Gdiplus::Color Value)
{
	mp_clrColor = Value;
}

LONG clsGrid::GetInterval(void)
{
	return mp_yInterval;
}

void clsGrid::SetInterval(LONG Value)
{
	mp_yInterval = Value;
}

LONG clsGrid::GetFactor(void)
{
	return mp_lFactor;
}

void clsGrid::SetFactor(LONG Value)
{
	mp_lFactor = Value;
}

void clsGrid::Finalize(void)
{
}

void clsGrid::Draw(void)
{
	COleDateTime dtBuff;
	if (mp_bVerticalLines == FALSE)
	{
		return;
	}
	if (mp_oControl->GetMathLib()->GetXCoordinateFromDate(mp_oControl->GetMathLib()->DateTimeAdd(mp_yInterval, mp_lFactor, mp_oTimeLine->GetStartDate())) - mp_oControl->GetMathLib()->GetXCoordinateFromDate(mp_oTimeLine->GetStartDate()) < 5)
	{
		return;
	}
	mp_oControl->clsG->mp_ClipRegion(mp_oTimeLine->Getf_lStart(), mp_oControl->GetCurrentViewObject()->GetClientArea()->GetTop(), mp_oTimeLine->Getf_lEnd(), mp_oControl->GetCurrentViewObject()->GetClientArea()->GetBottom(), TRUE);
	dtBuff = mp_oControl->GetMathLib()->RoundDate(mp_yInterval, mp_lFactor, mp_oTimeLine->GetStartDate());
	if (mp_oControl->GetMathLib()->GetXCoordinateFromDate(dtBuff) >= mp_oTimeLine->Getf_lStart())
	{
		mp_PaintVerticalGridLine(mp_oControl->GetMathLib()->GetXCoordinateFromDate(dtBuff), LDS_SOLID);
	}
	while (dtBuff < mp_oTimeLine->GetEndDate()) 
	{
		dtBuff = mp_oControl->GetMathLib()->DateTimeAdd(mp_yInterval, mp_lFactor, dtBuff);
		mp_PaintVerticalGridLine(mp_oControl->GetMathLib()->GetXCoordinateFromDate(dtBuff), LDS_SOLID);
	}
}

void clsGrid::mp_PaintVerticalGridLine(LONG fXCoordinate,LONG v_lDrawStyle)
{
	if (mp_bVerticalLines == TRUE)
	{
		mp_oControl->clsG->mp_DrawLine(fXCoordinate, mp_oControl->GetCurrentViewObject()->GetClientArea()->GetTop(), fXCoordinate, mp_oControl->GetRows()->GetTopOffset(), LT_NORMAL, mp_clrColor, (GRE_LINEDRAWSTYLE)v_lDrawStyle, 1, TRUE);
	}
}

CString clsGrid::GetXML(void)
{
	clsXML oXML(mp_oControl, "Grid");
	oXML.InitializeWriter();
	oXML.WriteProperty("Color", mp_clrColor);
	oXML.WriteProperty("Interval", mp_yInterval);
	oXML.WriteProperty("Factor", mp_lFactor);
	oXML.WriteProperty("SnapToGrid", mp_bSnapToGrid);
	oXML.WriteProperty("SnapToGridOnSelection", mp_bSnapToGridOnSelection);
	oXML.WriteProperty("VerticalLines", mp_bVerticalLines);
	return oXML.GetXML();
}

void clsGrid::SetXML(CString sXML)
{
	clsXML oXML(mp_oControl, "Grid");
	oXML.SetXML(sXML);
	oXML.InitializeReader();
	oXML.ReadProperty("Color", mp_clrColor);
	oXML.ReadProperty("Interval", mp_yInterval);
	oXML.ReadProperty("Factor", mp_lFactor);
	oXML.ReadProperty("SnapToGrid", mp_bSnapToGrid);
	oXML.ReadProperty("SnapToGridOnSelection", mp_bSnapToGridOnSelection);
	oXML.ReadProperty("VerticalLines", mp_bVerticalLines);
}

void clsGrid::Clear()
{
	mp_bVerticalLines = FALSE;
	mp_bSnapToGrid = FALSE;
	mp_bSnapToGridOnSelection = TRUE;
	mp_clrColor = Color::MakeARGB(255, 192, 192, 192);
	mp_yInterval = IL_MINUTE;
	mp_lFactor = 15;
}

void clsGrid::Clone(clsGrid* oClone)
{
    oClone->SetColor(mp_clrColor);
    oClone->SetInterval(mp_yInterval);
    oClone->SetFactor(mp_lFactor);
    oClone->SetSnapToGrid(mp_bSnapToGrid);
    oClone->SetSnapToGridOnSelection(mp_bSnapToGridOnSelection);
    oClone->SetVerticalLines(mp_bVerticalLines);
}