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
#include "clsVerticalScrollBar.h"



IMPLEMENT_DYNCREATE(clsVerticalScrollBar, CCmdTarget)

//{340B331A-EBFE-4D4B-A860-7D8EECE8EF45}
static const IID IID_IclsVerticalScrollBar = { 0x340B331A, 0xEBFE, 0x4D4B, { 0xA8, 0x60, 0x7D, 0x8E, 0xEC, 0xE8, 0xEF, 0x45} };

//{A36496EB-178E-426C-BD1A-7B987873A2D4}
IMPLEMENT_OLECREATE_FLAGS(clsVerticalScrollBar, "AGVC.clsVerticalScrollBar", afxRegApartmentThreading, 0xA36496EB, 0x178E, 0x426C, 0xBD, 0x1A, 0x7B, 0x98, 0x78, 0x73, 0xA2, 0xD4)

BEGIN_DISPATCH_MAP(clsVerticalScrollBar, CCmdTarget)
	DISP_PROPERTY_EX_ID(clsVerticalScrollBar, "Min", 1, odl_GetMin, SetNotSupported, VT_I4) 
	DISP_PROPERTY_EX_ID(clsVerticalScrollBar, "Max", 2, odl_GetMax, SetNotSupported, VT_I4) 
	DISP_PROPERTY_EX_ID(clsVerticalScrollBar, "Value", 3, odl_GetValue, odl_SetValue, VT_I4) 
	DISP_PROPERTY_EX_ID(clsVerticalScrollBar, "Visible", 4, odl_GetVisible, odl_SetVisible, VT_BOOL) 
	DISP_PROPERTY_EX_ID(clsVerticalScrollBar, "SmallChange", 5, odl_GetSmallChange, odl_SetSmallChange, VT_I4) 
	DISP_PROPERTY_EX_ID(clsVerticalScrollBar, "LargeChange", 6, odl_GetLargeChange, odl_SetLargeChange, VT_I4) 
	DISP_PROPERTY_EX_ID(clsVerticalScrollBar, "ScrollBar", 7, odl_GetScrollBar, SetNotSupported, VT_DISPATCH)
	DISP_FUNCTION_ID(clsVerticalScrollBar, "GetXML", 8, odl_GetXML, VT_BSTR, VTS_NONE) 
	DISP_FUNCTION_ID(clsVerticalScrollBar, "SetXML", 9, odl_SetXML, VT_EMPTY, VTS_BSTR)
END_DISPATCH_MAP()

BEGIN_INTERFACE_MAP(clsVerticalScrollBar, CCmdTarget)
	INTERFACE_PART(clsVerticalScrollBar, IID_IclsVerticalScrollBar, Dispatch)
END_INTERFACE_MAP()

BEGIN_MESSAGE_MAP(clsVerticalScrollBar, CCmdTarget)
END_MESSAGE_MAP()

clsVerticalScrollBar::clsVerticalScrollBar()
{
	AfxMessageBox(_T("Invalid constructor has been invoked for clsVerticalScrollBar. Please inform support@sourcecodestore.com of this error"), MB_OK | MB_ICONSTOP | MB_APPLMODAL);
}

clsVerticalScrollBar::clsVerticalScrollBar(CActiveGanttVCCtl* oControl)
{
	EnableAutomation();
	AfxOleLockApp();
	mp_oControl = oControl;
	mp_bVisible = TRUE;
	mp_oScrollBar = new clsVScrollBarTemplate(mp_oControl);
	mp_oScrollBar->Initialize(SCR_VERTICAL);
	mp_oScrollBar->SetLargeChange(1);
	mp_oScrollBar->SetSmallChange(1);
	mp_oScrollBar->SetMin(1);
	mp_oScrollBar->SetMax(1);
	mp_oScrollBar->SetValue(1);
}

clsVerticalScrollBar::~clsVerticalScrollBar()
{
	delete mp_oScrollBar;
	mp_oScrollBar = NULL;
	AfxOleUnlockApp();
}

void clsVerticalScrollBar::OnFinalRelease()
{
	CCmdTarget::OnFinalRelease();
}



LONG clsVerticalScrollBar::GetMin(void)
{
	if (mp_oControl->GetRows()->GetCount() == 0)
	{
		return 0;
	}
	else
	{
		return mp_oScrollBar->GetMin();
	}
}

LONG clsVerticalScrollBar::GetMax(void)
{
	if (mp_oControl->GetRows()->GetCount() == 0)
	{
		return 0;
	}
	else
	{
		mp_oScrollBar->SetMax(mp_oControl->GetRows()->GetCount() - mp_oControl->GetRows()->HiddenRows());
		return mp_oControl->GetRows()->GetCount() - mp_oControl->GetRows()->HiddenRows();
	}
}

LONG clsVerticalScrollBar::GetValue(void)
{
	if (mp_oControl->GetRows()->GetCount() == 0)
	{
		return 0;
	}
	else
	{
		return mp_oScrollBar->GetValue();
	}
}

void clsVerticalScrollBar::SetValue(LONG Value)
{
	if (mp_oControl->GetRows()->GetCount() > 0)
	{
		if (Value < 1)
		{
			Value = 1;
		}
		if (Value > mp_oControl->GetRows()->GetCount())
		{
			Value = mp_oControl->GetRows()->GetCount();
		}
		mp_oScrollBar->SetValue(Value);
	}
}

BOOL clsVerticalScrollBar::mf_Visible(void)
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

BOOL clsVerticalScrollBar::GetVisible(void)
{
	return mp_bVisible;
}

void clsVerticalScrollBar::SetVisible(BOOL Value)
{
	mp_bVisible = Value;
}

LONG clsVerticalScrollBar::GetSmallChange(void)
{
	return mp_oScrollBar->GetSmallChange();
}

void clsVerticalScrollBar::SetSmallChange(LONG Value)
{
	mp_oScrollBar->SetSmallChange(Value);
}

LONG clsVerticalScrollBar::GetLargeChange(void)
{
	return mp_oScrollBar->GetLargeChange();
}

void clsVerticalScrollBar::SetLargeChange(LONG Value)
{
	mp_oScrollBar->SetLargeChange(Value);
}

clsVScrollBarTemplate* clsVerticalScrollBar::GetScrollBar(void)
{
	return mp_oScrollBar;
}

LONG clsVerticalScrollBar::GetState(void)
{
	return mp_oScrollBar->GetState();
}

void clsVerticalScrollBar::SetState(LONG Value)
{
	mp_oScrollBar->SetState(Value);
}

LONG clsVerticalScrollBar::GetWidth(void)
{
	return mp_oScrollBar->Width;
}

void clsVerticalScrollBar::SetWidth(LONG Value)
{
	mp_oScrollBar->Width = Value;
}

LONG clsVerticalScrollBar::GetHeight(void)
{
	return mp_oScrollBar->Height;
}

void clsVerticalScrollBar::SetHeight(LONG Value)
{
	mp_oScrollBar->Height = Value;
}

LONG clsVerticalScrollBar::GetLeft(void)
{
	return mp_oScrollBar->Left;
}

void clsVerticalScrollBar::SetLeft(LONG Value)
{
	mp_oScrollBar->Left = Value;
}

LONG clsVerticalScrollBar::GetTop(void)
{
	return mp_oScrollBar->Top;
}

void clsVerticalScrollBar::SetTop(LONG Value)
{
	mp_oScrollBar->Top = Value;
}

void clsVerticalScrollBar::Update(void)
{
	LONG lHiddenRows = mp_oControl->GetRows()->HiddenRows();
	if (mp_oControl->GetRows()->GetCount() > 0)
	{
		if (mp_oScrollBar->GetValue() > (mp_oControl->GetRows()->GetCount() - lHiddenRows))
		{
			mp_oScrollBar->SetValue(mp_oControl->GetRows()->GetCount() - lHiddenRows);
		}
		mp_oScrollBar->SetMax(mp_oControl->GetRows()->GetCount() - lHiddenRows);
	}
	else
	{
		Reset();
	}
}

void clsVerticalScrollBar::Reset(void)
{
	mp_oScrollBar->SetMin(1);
	mp_oScrollBar->SetMax(1);
	mp_oScrollBar->SetValue(1);
}

void clsVerticalScrollBar::Position(void)
{
	SetLeft(mp_oControl->clsG->Width() - GetWidth() - mp_oControl->Getmt_BorderThickness());
	SetTop(mp_oControl->Getmt_TopMargin());
	SetHeight(mp_oControl->clsG->Height() - (mp_oControl->Getmt_BorderThickness() * 2) - mp_oControl->GetHorizontalScrollBar()->GetHeight());
	SetSmallChange(1);
}

void clsVerticalScrollBar::oVScrollBar_ValueChanged(LONG Offset)
{
}

CString clsVerticalScrollBar::GetXML(void)
{
	clsXML oXML(mp_oControl, "VerticalScrollBar");
	oXML.InitializeWriter();
	oXML.WriteObject(mp_oScrollBar->GetXML());
	return oXML.GetXML();
}

void clsVerticalScrollBar::SetXML(CString sXML)
{
	clsXML oXML(mp_oControl, "VerticalScrollBar");
	oXML.SetXML(sXML);
	oXML.InitializeReader();
	mp_oScrollBar->SetXML(oXML.ReadObject("ScrollBar"));
}