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
#include "clsHorizontalScrollBar.h"



IMPLEMENT_DYNCREATE(clsHorizontalScrollBar, CCmdTarget)

//{C98BEB5B-261C-4183-A179-0F3E15CBBF89}
static const IID IID_IclsHorizontalScrollBar = { 0xC98BEB5B, 0x261C, 0x4183, { 0xA1, 0x79, 0x0F, 0x3E, 0x15, 0xCB, 0xBF, 0x89} };

//{E36C2CEB-770D-46E7-99BE-945BA476BC6F}
IMPLEMENT_OLECREATE_FLAGS(clsHorizontalScrollBar, "AGVC.clsHorizontalScrollBar", afxRegApartmentThreading, 0xE36C2CEB, 0x770D, 0x46E7, 0x99, 0xBE, 0x94, 0x5B, 0xA4, 0x76, 0xBC, 0x6F)

BEGIN_DISPATCH_MAP(clsHorizontalScrollBar, CCmdTarget)
	DISP_PROPERTY_EX_ID(clsHorizontalScrollBar, "Min", 1, odl_GetMin, SetNotSupported, VT_I4) 
	DISP_PROPERTY_EX_ID(clsHorizontalScrollBar, "Max", 2, odl_GetMax, SetNotSupported, VT_I4) 
	DISP_PROPERTY_EX_ID(clsHorizontalScrollBar, "Value", 3, odl_GetValue, odl_SetValue, VT_I4) 
	DISP_PROPERTY_EX_ID(clsHorizontalScrollBar, "Visible", 4, odl_GetVisible, odl_SetVisible, VT_BOOL) 
	DISP_PROPERTY_EX_ID(clsHorizontalScrollBar, "SmallChange", 5, odl_GetSmallChange, odl_SetSmallChange, VT_I4) 
	DISP_PROPERTY_EX_ID(clsHorizontalScrollBar, "LargeChange", 6, odl_GetLargeChange, odl_SetLargeChange, VT_I4) 
	DISP_PROPERTY_EX_ID(clsHorizontalScrollBar, "ScrollBar", 7, odl_GetScrollBar, SetNotSupported, VT_DISPATCH)
	DISP_FUNCTION_ID(clsHorizontalScrollBar, "GetXML", 8, odl_GetXML, VT_BSTR, VTS_NONE) 
	DISP_FUNCTION_ID(clsHorizontalScrollBar, "SetXML", 9, odl_SetXML, VT_EMPTY, VTS_BSTR)
END_DISPATCH_MAP()

BEGIN_INTERFACE_MAP(clsHorizontalScrollBar, CCmdTarget)
	INTERFACE_PART(clsHorizontalScrollBar, IID_IclsHorizontalScrollBar, Dispatch)
END_INTERFACE_MAP()

BEGIN_MESSAGE_MAP(clsHorizontalScrollBar, CCmdTarget)
END_MESSAGE_MAP()

clsHorizontalScrollBar::clsHorizontalScrollBar()
{
	AfxMessageBox(_T("Invalid constructor has been invoked for clsHorizontalScrollBar. Please inform support@sourcecodestore.com of this error"), MB_OK | MB_ICONSTOP | MB_APPLMODAL);
}

clsHorizontalScrollBar::clsHorizontalScrollBar(CActiveGanttVCCtl* oControl)
{
	EnableAutomation();
	AfxOleLockApp();
	mp_oControl = oControl;
	mp_bVisible = TRUE;
	mp_oScrollBar = new clsHScrollBarTemplate(mp_oControl);
	mp_oScrollBar->Initialize(SCR_HORIZONTAL1);
	mp_oScrollBar->SetLargeChange(1);
	mp_oScrollBar->SetSmallChange(1);
	mp_oScrollBar->SetMin(0);
	mp_oScrollBar->SetMax(0);
	mp_oScrollBar->SetValue(0);
}

clsHorizontalScrollBar::~clsHorizontalScrollBar()
{
	delete mp_oScrollBar;
	mp_oScrollBar = NULL;
	AfxOleUnlockApp();
}

void clsHorizontalScrollBar::OnFinalRelease()
{
	CCmdTarget::OnFinalRelease();
}



LONG clsHorizontalScrollBar::GetMin(void)
{
	return 0;
}

LONG clsHorizontalScrollBar::GetMax(void)
{
	return mp_oScrollBar->GetMax();
}

LONG clsHorizontalScrollBar::GetValue(void)
{
	return mp_oScrollBar->GetValue();
}

void clsHorizontalScrollBar::SetValue(LONG Value)
{
	mp_oScrollBar->SetValue(Value);
}

BOOL clsHorizontalScrollBar::mf_Visible(void)
{
	return mp_bVisible;
}

BOOL clsHorizontalScrollBar::GetVisible(void)
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

void clsHorizontalScrollBar::SetVisible(BOOL Value)
{
	mp_bVisible = Value;
}

LONG clsHorizontalScrollBar::GetSmallChange(void)
{
	return mp_oScrollBar->GetSmallChange();
}

void clsHorizontalScrollBar::SetSmallChange(LONG Value)
{
	mp_oScrollBar->SetSmallChange(Value);
}

LONG clsHorizontalScrollBar::GetLargeChange(void)
{
	return mp_oScrollBar->GetLargeChange();
}

void clsHorizontalScrollBar::SetLargeChange(LONG Value)
{
	mp_oScrollBar->SetLargeChange(Value);
}

clsHScrollBarTemplate* clsHorizontalScrollBar::GetScrollBar(void)
{
	return mp_oScrollBar;
}

LONG clsHorizontalScrollBar::GetState(void)
{
	return mp_oScrollBar->GetState();
}

void clsHorizontalScrollBar::SetState(LONG Value)
{
	mp_oScrollBar->SetState(Value);
}

LONG clsHorizontalScrollBar::GetWidth(void)
{
	return mp_oScrollBar->Width;
}

void clsHorizontalScrollBar::SetWidth(LONG Value)
{
	mp_oScrollBar->Width = Value;
}

LONG clsHorizontalScrollBar::GetHeight(void)
{
	return mp_oScrollBar->Height;
}

void clsHorizontalScrollBar::SetHeight(LONG Value)
{
	mp_oScrollBar->Height = Value;
}

LONG clsHorizontalScrollBar::GetLeft(void)
{
	return mp_oScrollBar->Left;
}

void clsHorizontalScrollBar::SetLeft(LONG Value)
{
	mp_oScrollBar->Left = Value;
}

LONG clsHorizontalScrollBar::GetTop(void)
{
	return mp_oScrollBar->Top;
}

void clsHorizontalScrollBar::SetTop(LONG Value)
{
	mp_oScrollBar->Top = Value;
}

void clsHorizontalScrollBar::Update(void)
{
}

void clsHorizontalScrollBar::Reset(void)
{
	mp_oScrollBar->SetMin(0);
	mp_oScrollBar->SetMax(0);
	mp_oScrollBar->SetValue(0);
}

void clsHorizontalScrollBar::Position(void)
{
	SetLeft(mp_oControl->Getmt_BorderThickness());
	SetTop(mp_oControl->clsG->Height() - GetHeight() - mp_oControl->Getmt_BorderThickness());
	if (mp_oControl->GetSplitter()->GetLeft() > 0)
	{
		SetWidth(mp_oControl->GetSplitter()->GetLeft());
	}
	mp_oScrollBar->SetMax(mp_oControl->GetColumns()->GetWidth() - mp_oControl->GetSplitter()->GetPosition());
}

void clsHorizontalScrollBar::oHScrollBar_ValueChanged(LONG Offset)
{
}

CString clsHorizontalScrollBar::GetXML(void)
{
	clsXML oXML(mp_oControl, "HorizontalScrollBar");
	oXML.InitializeWriter();
	oXML.WriteObject(mp_oScrollBar->GetXML());
	return oXML.GetXML();
}

void clsHorizontalScrollBar::SetXML(CString sXML)
{
	clsXML oXML(mp_oControl, "HorizontalScrollBar");
	oXML.SetXML(sXML);
	oXML.InitializeReader();
	mp_oScrollBar->SetXML(oXML.ReadObject("ScrollBar"));
}