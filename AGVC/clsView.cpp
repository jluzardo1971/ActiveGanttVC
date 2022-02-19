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
#include "clsView.h"



IMPLEMENT_DYNCREATE(clsView, clsItemBase)

//{4C3F2458-7CED-4729-9B42-99AD68923E07}
static const IID IID_IclsView = { 0x4C3F2458, 0x7CED, 0x4729, { 0x9B, 0x42, 0x99, 0xAD, 0x68, 0x92, 0x3E, 0x07} };

//{3EE51508-D007-4D56-8409-BAFB60B3AAA3}
IMPLEMENT_OLECREATE_FLAGS(clsView, "AGVC.clsView", afxRegApartmentThreading, 0x3EE51508, 0xD007, 0x4D56, 0x84, 0x09, 0xBA, 0xFB, 0x60, 0xB3, 0xAA, 0xA3)

BEGIN_DISPATCH_MAP(clsView, clsItemBase)
	DISP_PROPERTY_EX_ID(clsView, "Key", 1, odl_GetKey, odl_SetKey, VT_BSTR) 
	DISP_PROPERTY_EX_ID(clsView, "TimeLine", 2, odl_GetTimeLine, SetNotSupported, VT_DISPATCH) 
	DISP_PROPERTY_EX_ID(clsView, "ClientArea", 3, odl_GetClientArea, SetNotSupported, VT_DISPATCH) 
	DISP_PROPERTY_EX_ID(clsView, "Tag", 4, odl_GetTag, odl_SetTag, VT_BSTR) 
	DISP_PROPERTY_EX_ID(clsView, "Interval", 5, odl_GetInterval, odl_SetInterval, VT_I4) 
	DISP_FUNCTION_ID(clsView, "GetXML", 6, odl_GetXML, VT_BSTR, VTS_NONE) 
	DISP_FUNCTION_ID(clsView, "SetXML", 7, odl_SetXML, VT_EMPTY, VTS_BSTR) 
	DISP_PROPERTY_EX_ID(clsView, "Index", 8, odl_GetIndex, SetNotSupported, VT_I4)
	DISP_PROPERTY_EX_ID(clsView, "Factor", 9, odl_GetFactor, odl_SetFactor, VT_I4)
	DISP_FUNCTION_ID(clsView, "Clear", 10, odl_Clear, VT_EMPTY, VTS_NONE) 
	DISP_FUNCTION_ID(clsView, "Clone", 11, odl_Clone, VT_DISPATCH, VTS_BSTR) 
END_DISPATCH_MAP()

BEGIN_INTERFACE_MAP(clsView, clsItemBase)
	INTERFACE_PART(clsView, IID_IclsView, Dispatch)
END_INTERFACE_MAP()

BEGIN_MESSAGE_MAP(clsView, clsItemBase)
END_MESSAGE_MAP()

clsView::clsView()
{
	AfxMessageBox(_T("Invalid constructor has been invoked for clsView. Please inform support@sourcecodestore.com of this error"), MB_OK | MB_ICONSTOP | MB_APPLMODAL);
}

clsView::clsView(CActiveGanttVCCtl* oControl)
{
	EnableAutomation();
	AfxOleLockApp();
	mp_oControl = oControl;
	mp_oTimeLine = new clsTimeLine(mp_oControl, this);
	mp_oClientArea = new clsClientArea(mp_oControl, mp_oTimeLine);
	Clear();
}

clsView::~clsView()
{
	delete mp_oTimeLine;
	mp_oTimeLine = NULL;
	delete mp_oClientArea;
	mp_oClientArea = NULL;
	AfxOleUnlockApp();
}

void clsView::OnFinalRelease()
{
	clsItemBase::OnFinalRelease();
}

CString clsView::GetKey(void)
{
	return mp_sKey;
}

void clsView::SetKey(CString Value)
{
	mp_oControl->GetViews()->mp_oCollection->mp_SetKey(&mp_sKey, Value, VIEWS_SET_KEY);
}

clsTimeLine* clsView::GetTimeLine(void)
{
	return mp_oTimeLine;
}

clsClientArea* clsView::GetClientArea(void)
{
	return mp_oClientArea;
}

CString clsView::GetTag(void)
{
	return mp_sTag;
}

void clsView::SetTag(CString Value)
{
	mp_sTag = Value;
}

LONG clsView::GetInterval(void)
{
	return mp_yInterval;
}

void clsView::SetInterval(LONG Value)
{
	mp_yInterval = Value;
	if (mp_yInterval == IL_YEAR)
	{
		mp_yScrollInterval = IL_YEAR;
	}
	else
	{
		mp_yScrollInterval = mp_yInterval + 1;
	}
}

LONG clsView::GetFactor(void)
{
	return mp_lFactor;
}

void clsView::SetFactor(LONG Value)
{
	mp_lFactor = Value;
}

LONG clsView::GetIndex(void)
{
	return mp_lIndex;
}

void clsView::SetIndex(LONG Value)
{
	mp_lIndex = Value;
}

LONG clsView::Getf_ScrollInterval(void)
{
	return mp_yScrollInterval;
}

void clsView::Setf_ScrollInterval(LONG value)
{
	mp_yScrollInterval = value;
}

void clsView::Finalize(void)
{
}

CString clsView::GetXML(void)
{
	clsXML oXML(mp_oControl, "View");
	oXML.InitializeWriter();
	oXML.WriteProperty("Interval", mp_yInterval);
	oXML.WriteProperty("Factor", mp_lFactor);
	oXML.WriteProperty("Key", mp_sKey);
	oXML.WriteProperty("Tag", mp_sTag);
	oXML.WriteObject(mp_oClientArea->GetXML());
	oXML.WriteObject(mp_oTimeLine->GetXML());
	return oXML.GetXML();
}

void clsView::SetXML(CString sXML)
{
	clsXML oXML(mp_oControl, "View");
	oXML.SetXML(sXML);
	oXML.InitializeReader();
    oXML.ReadProperty("Interval", mp_yInterval);
    oXML.ReadProperty("Factor", mp_lFactor);
    oXML.ReadProperty("Key", mp_sKey);
    oXML.ReadProperty("Tag", mp_sTag);
    mp_oClientArea->SetXML(oXML.ReadObject("ClientArea"));
    mp_oTimeLine->SetXML(oXML.ReadObject("TimeLine"));
}

void clsView::Clear(void)
{
	mp_yInterval = IL_MINUTE;
	mp_lFactor = 1;
	mp_yScrollInterval = IL_HOUR;
	mp_sTag = "";
	GetTimeLine()->Clear();
	GetClientArea()->Clear();
}

clsView* clsView::Clone(CString Key)
{
	clsView* oClone = mp_oControl->GetViews()->Add(IL_SECOND, 1, ST_DAYOFWEEK, ST_DAYOFWEEK, ST_DAYOFWEEK, Key);
    oClone->SetInterval(mp_yInterval);
    oClone->SetFactor(mp_lFactor);
    oClone->Setf_ScrollInterval(mp_yScrollInterval);
    oClone->SetTag(mp_sTag);
    GetTimeLine()->Clone(oClone->GetTimeLine());
    GetClientArea()->Clone(oClone->GetClientArea());
    return oClone;
}