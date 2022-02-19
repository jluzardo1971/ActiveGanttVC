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
#include "clsTimeLineScrollBar.h"



IMPLEMENT_DYNCREATE(clsTimeLineScrollBar, CCmdTarget)

//{FE03E8A1-4EA7-4CAC-90ED-20A53EAA714F}
static const IID IID_IclsTimeLineScrollBar = { 0xFE03E8A1, 0x4EA7, 0x4CAC, { 0x90, 0xED, 0x20, 0xA5, 0x3E, 0xAA, 0x71, 0x4F} };

//{D805318C-5FF6-4A77-B988-E3A31541E80D}
IMPLEMENT_OLECREATE_FLAGS(clsTimeLineScrollBar, "AGVC.clsTimeLineScrollBar", afxRegApartmentThreading, 0xD805318C, 0x5FF6, 0x4A77, 0xB9, 0x88, 0xE3, 0xA3, 0x15, 0x41, 0xE8, 0x0D)

BEGIN_DISPATCH_MAP(clsTimeLineScrollBar, CCmdTarget)
	DISP_PROPERTY_EX_ID(clsTimeLineScrollBar, "Interval", 1, odl_GetInterval, odl_SetInterval, VT_I4) 
	DISP_PROPERTY_EX_ID(clsTimeLineScrollBar, "Value", 2, odl_GetValue, odl_SetValue, VT_I4) 
	DISP_PROPERTY_EX_ID(clsTimeLineScrollBar, "Enabled", 3, odl_GetEnabled, odl_SetEnabled, VT_BOOL) 
	DISP_PROPERTY_EX_ID(clsTimeLineScrollBar, "Visible", 4, odl_GetVisible, odl_SetVisible, VT_BOOL) 
	DISP_PROPERTY_EX_ID(clsTimeLineScrollBar, "LargeChange", 5, odl_GetLargeChange, odl_SetLargeChange, VT_I4) 
	DISP_PROPERTY_EX_ID(clsTimeLineScrollBar, "Max", 6, odl_GetMax, odl_SetMax, VT_I4) 
	DISP_PROPERTY_EX_ID(clsTimeLineScrollBar, "SmallChange", 7, odl_GetSmallChange, odl_SetSmallChange, VT_I4) 
	DISP_PROPERTY_EX_ID(clsTimeLineScrollBar, "StartDate", 8, odl_GetStartDate, odl_SetStartDate, VT_DATE) 
	DISP_PROPERTY_EX_ID(clsTimeLineScrollBar, "ScrollBar", 9, odl_GetScrollBar, SetNotSupported, VT_DISPATCH) 
	DISP_FUNCTION_ID(clsTimeLineScrollBar, "GetXML", 10, odl_GetXML, VT_BSTR, VTS_NONE) 
	DISP_FUNCTION_ID(clsTimeLineScrollBar, "SetXML", 11, odl_SetXML, VT_EMPTY, VTS_BSTR)
	DISP_PROPERTY_EX_ID(clsTimeLineScrollBar, "Factor", 12, odl_GetFactor, odl_SetFactor, VT_I4) 
END_DISPATCH_MAP()

BEGIN_INTERFACE_MAP(clsTimeLineScrollBar, CCmdTarget)
	INTERFACE_PART(clsTimeLineScrollBar, IID_IclsTimeLineScrollBar, Dispatch)
END_INTERFACE_MAP()

BEGIN_MESSAGE_MAP(clsTimeLineScrollBar, CCmdTarget)
END_MESSAGE_MAP()

clsTimeLineScrollBar::clsTimeLineScrollBar()
{
	AfxMessageBox(_T("Invalid constructor has been invoked for clsTimeLineScrollBar. Please inform support@sourcecodestore.com of this error"), MB_OK | MB_ICONSTOP | MB_APPLMODAL);
}

clsTimeLineScrollBar::clsTimeLineScrollBar(CActiveGanttVCCtl* oControl)
{
	EnableAutomation();
	AfxOleLockApp();
	mp_oControl = oControl;
	mp_oScrollBar = new clsHScrollBarTemplate(mp_oControl);
	Clear();
}

clsTimeLineScrollBar::~clsTimeLineScrollBar()
{
	delete mp_oScrollBar;
	mp_oScrollBar = NULL;
	AfxOleUnlockApp();
}

void clsTimeLineScrollBar::OnFinalRelease()
{
	CCmdTarget::OnFinalRelease();
}

LONG clsTimeLineScrollBar::GetInterval(void)
{
	return mp_yInterval;
}

void clsTimeLineScrollBar::SetInterval(LONG Value)
{
	mp_yInterval = Value;
}

LONG clsTimeLineScrollBar::GetFactor(void)
{
	return mp_lFactor;
}

void clsTimeLineScrollBar::SetFactor(LONG Value)
{
	mp_lFactor = Value;
}

LONG clsTimeLineScrollBar::GetValue(void)
{
	return mp_oScrollBar->GetValue();
}

void clsTimeLineScrollBar::SetValue(LONG Value)
{
	if (Value < 0)
	{
		Value = 0;
	}
	if (Value > mp_oScrollBar->GetMax())
	{
		Value = mp_oScrollBar->GetMax();
	}
	mp_oScrollBar->SetValue(Value);
}

BOOL clsTimeLineScrollBar::GetEnabled(void)
{
	return mp_oScrollBar->GetEnabled();
}

void clsTimeLineScrollBar::SetEnabled(BOOL Value)
{
	mp_oScrollBar->SetEnabled(Value);
}

BOOL clsTimeLineScrollBar::mf_Visible(void)
{
	return mp_bVisible;
}

BOOL clsTimeLineScrollBar::GetVisible(void)
{
	if (GetState() != SS_SHOWN)
	{
		return FALSE;
	}
	else
	{
		return mp_bVisible;
	}
}

void clsTimeLineScrollBar::SetVisible(BOOL Value)
{
	mp_bVisible = Value;
}

LONG clsTimeLineScrollBar::GetLargeChange(void)
{
	return mp_oScrollBar->GetLargeChange();
}

void clsTimeLineScrollBar::SetLargeChange(LONG Value)
{
	mp_oScrollBar->SetLargeChange(Value);
}

LONG clsTimeLineScrollBar::GetMax(void)
{
	return mp_oScrollBar->GetMax();
}

void clsTimeLineScrollBar::SetMax(LONG Value)
{
	mp_oScrollBar->SetMax(Value);
}

LONG clsTimeLineScrollBar::GetSmallChange(void)
{
	return mp_oScrollBar->GetSmallChange();
}

void clsTimeLineScrollBar::SetSmallChange(LONG Value)
{
	mp_oScrollBar->SetSmallChange(Value);
}

COleDateTime clsTimeLineScrollBar::GetStartDate(void)
{
	return mp_dtStartDate;
}

void clsTimeLineScrollBar::SetStartDate(COleDateTime Value)
{
	mp_dtStartDate = Value;
}

clsHScrollBarTemplate* clsTimeLineScrollBar::GetScrollBar(void)
{
	return mp_oScrollBar;
}

LONG clsTimeLineScrollBar::GetState(void)
{
	return mp_oScrollBar->GetState();
}

void clsTimeLineScrollBar::SetState(LONG Value)
{
	mp_oScrollBar->SetState(Value);
}

LONG clsTimeLineScrollBar::GetWidth(void)
{
	return mp_oScrollBar->Width;
}

void clsTimeLineScrollBar::SetWidth(LONG Value)
{
	mp_oScrollBar->Width = Value;
}

LONG clsTimeLineScrollBar::GetHeight(void)
{
	return mp_oScrollBar->Height;
}

void clsTimeLineScrollBar::SetHeight(LONG Value)
{
	mp_oScrollBar->Height = Value;
}

LONG clsTimeLineScrollBar::GetLeft(void)
{
	return mp_oScrollBar->Left;
}

void clsTimeLineScrollBar::SetLeft(LONG Value)
{
	mp_oScrollBar->Left = Value;
}

LONG clsTimeLineScrollBar::GetTop(void)
{
	return mp_oScrollBar->Top;
}

void clsTimeLineScrollBar::SetTop(LONG Value)
{
	mp_oScrollBar->Top = Value;
}

void clsTimeLineScrollBar::Finalize(void)
{
}

CString clsTimeLineScrollBar::GetXML(void)
{
	clsXML oXML(mp_oControl, "TimeLineScrollBar");
	oXML.InitializeWriter();
	oXML.WriteProperty("Enabled", mp_oScrollBar->mp_bEnabled);
	oXML.WriteProperty("Interval", mp_yInterval);
	oXML.WriteProperty("Factor", mp_lFactor);
	oXML.WriteProperty("LargeChange", mp_oScrollBar->mp_lLargeChange);
	oXML.WriteProperty("Max", mp_oScrollBar->mp_lMax);
	oXML.WriteProperty("SmallChange", mp_oScrollBar->mp_lSmallChange);
	oXML.WriteProperty("StartDate", mp_dtStartDate);
	oXML.WriteProperty("Value", mp_oScrollBar->mp_lValue);
	oXML.WriteProperty("Visible", mp_bVisible);
	oXML.WriteObject(mp_oScrollBar->GetXML());
	return oXML.GetXML();
}

void clsTimeLineScrollBar::SetXML(CString sXML)
{
	clsXML oXML(mp_oControl, "TimeLineScrollBar");
	oXML.SetXML(sXML);
	oXML.InitializeReader();
	oXML.ReadProperty("Enabled", mp_oScrollBar->mp_bEnabled);
	oXML.ReadProperty("Interval", mp_yInterval);
	oXML.ReadProperty("Factor", mp_lFactor);
	oXML.ReadProperty("LargeChange", mp_oScrollBar->mp_lLargeChange);
	oXML.ReadProperty("Max", mp_oScrollBar->mp_lMax);
	oXML.ReadProperty("SmallChange", mp_oScrollBar->mp_lSmallChange);
	oXML.ReadProperty("StartDate", mp_dtStartDate);
	oXML.ReadProperty("Value", mp_oScrollBar->mp_lValue);
	oXML.ReadProperty("Visible", mp_bVisible);
	mp_oScrollBar->SetXML(oXML.ReadObject("ScrollBar"));
}

void clsTimeLineScrollBar::oHScrollBar_ValueChanged(LONG Offset)
{
}

void clsTimeLineScrollBar::Position(void)
{
	SetLeft(mp_oControl->GetSplitter()->GetRight());
	SetTop(mp_oControl->clsG->Height() - GetHeight() - mp_oControl->Getmt_BorderThickness());
	SetWidth(mp_oControl->clsG->Width() - mp_oControl->Getmt_BorderThickness() - mp_oControl->GetSplitter()->GetRight() - mp_oControl->GetVerticalScrollBar()->GetWidth());
}

void clsTimeLineScrollBar::Clear()
{
	mp_oScrollBar->Initialize(SCR_HORIZONTAL2);
	mp_dtStartDate = (DATE)0;
	mp_oScrollBar->SetMin(0);
	mp_yInterval = IL_MINUTE;
	mp_lFactor = 1;
	mp_oScrollBar->SetEnabled(FALSE);
	mp_bVisible = FALSE;
}

void clsTimeLineScrollBar::Clone(clsTimeLineScrollBar* oClone)
{
	oClone->SetEnabled(mp_oScrollBar->GetEnabled());
    oClone->SetInterval(mp_yInterval);
    oClone->SetFactor(mp_lFactor);
    oClone->SetStartDate(mp_dtStartDate);
    oClone->SetVisible(mp_bVisible);
    mp_oScrollBar->Clone(oClone->mp_oScrollBar);
}