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
#include "clsViews.h"

IMPLEMENT_DYNCREATE(clsViews, CCmdTarget)

//{72ED06E0-4FF4-4851-BF45-E65708C520CF}
static const IID IID_IclsViews = { 0x72ED06E0, 0x4FF4, 0x4851, { 0xBF, 0x45, 0xE6, 0x57, 0x08, 0xC5, 0x20, 0xCF} };

//{830F6BCD-4925-417D-A94A-131410CB8D81}
IMPLEMENT_OLECREATE_FLAGS(clsViews, "AGVC.clsViews", afxRegApartmentThreading, 0x830F6BCD, 0x4925, 0x417D, 0xA9, 0x4A, 0x13, 0x14, 0x10, 0xCB, 0x8D, 0x81)

BEGIN_DISPATCH_MAP(clsViews, CCmdTarget)
	DISP_PROPERTY_EX_ID(clsViews, "Count", 1, odl_GetCount, SetNotSupported, VT_I4) 
	DISP_PROPERTY_PARAM_ID(clsViews, "Item", 2, odl_Item, SetNotSupported, VT_DISPATCH, VTS_BSTR) 
	DISP_FUNCTION_ID(clsViews, "Add", 3, odl_Add, VT_DISPATCH, VTS_I4 VTS_I4 VTS_I4 VTS_I4 VTS_I4 VTS_BSTR) 
	DISP_FUNCTION_ID(clsViews, "Clear", 4, odl_Clear, VT_EMPTY, VTS_NONE) 
	DISP_FUNCTION_ID(clsViews, "Remove", 5, odl_Remove, VT_EMPTY, VTS_BSTR) 
	DISP_FUNCTION_ID(clsViews, "GetXML", 6, odl_GetXML, VT_BSTR, VTS_NONE) 
	DISP_FUNCTION_ID(clsViews, "SetXML", 7, odl_SetXML, VT_EMPTY, VTS_BSTR)
	DISP_DEFVALUE(clsViews, "Item")
END_DISPATCH_MAP()

BEGIN_INTERFACE_MAP(clsViews, CCmdTarget)
	INTERFACE_PART(clsViews, IID_IclsViews, Dispatch)
END_INTERFACE_MAP()

BEGIN_MESSAGE_MAP(clsViews, CCmdTarget)
END_MESSAGE_MAP()

clsViews::clsViews()
{
	AfxMessageBox(_T("Invalid constructor has been invoked for clsViews. Please inform support@sourcecodestore.com of this error"), MB_OK | MB_ICONSTOP | MB_APPLMODAL);
}

clsViews::clsViews(CActiveGanttVCCtl* oControl)
{
	EnableAutomation();
	AfxOleLockApp();

	mp_oControl = oControl;
	
	mp_oCollection = new clsCollectionBase(mp_oControl, "View");
	mp_oDefaultView = new clsView(mp_oControl);
	mp_oDefaultView->SetInterval(IL_MINUTE);
	mp_oDefaultView->SetFactor(1);
	mp_oDefaultView->GetTimeLine()->GetTierArea()->GetUpperTier()->SetTierType((LONG)ST_MONTH);
	mp_oDefaultView->GetTimeLine()->GetTierArea()->GetMiddleTier()->SetVisible(FALSE);
	mp_oDefaultView->GetTimeLine()->GetTierArea()->GetLowerTier()->SetTierType((LONG)ST_DAYOFWEEK);
	mp_oDefaultView->GetClientArea()->SetToolTipFormat("hh:mmtt");
	mp_oDefaultView->GetTimeLine()->GetTickMarkArea()->GetTickMarks()->Add(IL_HOUR, 1, TLT_BIG, TRUE, "hh:mmtt", "");
	mp_oDefaultView->GetTimeLine()->GetTickMarkArea()->GetTickMarks()->Add(IL_MINUTE, 30, TLT_BIG, FALSE, "", "");
	mp_oDefaultView->GetTimeLine()->GetTickMarkArea()->GetTickMarks()->Add(IL_MINUTE, 15, TLT_MEDIUM, FALSE, "", "");
	mp_oDefaultView->GetTimeLine()->GetTickMarkArea()->GetTickMarks()->Add(IL_MINUTE, 5, TLT_SMALL, FALSE, "", "");
	mp_oDefaultView->GetTimeLine()->GetTickMarkArea()->SetHeight(30);
}

clsViews::~clsViews()
{
	delete mp_oCollection;
	mp_oCollection = NULL;
	delete mp_oDefaultView;
	mp_oDefaultView = NULL;
	AfxOleUnlockApp();
}

void clsViews::OnFinalRelease()
{
	CCmdTarget::OnFinalRelease();
}

LONG clsViews::GetCount(void)
{
	return mp_oCollection->m_lCount();
}

clsCollectionBase* clsViews::GetCollection(void)
{
	return mp_oCollection;
}

void clsViews::Finalize(void)
{
}

clsView* clsViews::Item(CString Index)
{
	clsView *oView;
	oView = (clsView*)mp_oCollection->m_oItem(Index, VIEWS_ITEM_1, VIEWS_ITEM_2, VIEWS_ITEM_3, VIEWS_ITEM_4);
	return oView;
}

clsView* clsViews::FItem(CString Index)
{
	if (Index == "0")
	{
		return mp_oDefaultView;
	}
	else
	{
		return (clsView*) mp_oCollection->m_oItem(Index, VIEWS_ITEM_1, VIEWS_ITEM_2, VIEWS_ITEM_3, VIEWS_ITEM_4);
	}
}

clsView* clsViews::Add(LONG Interval,LONG Factor,LONG UpperTierType,LONG MiddleTierType,LONG LowerTierType,CString Key)
{
	mp_oCollection->SetAddMode(TRUE);
	clsView* oView = new clsView(mp_oControl);
	oView->SetInterval(Interval);
	oView->SetFactor(Factor);
	oView->GetTimeLine()->GetTierArea()->GetUpperTier()->SetTierType(UpperTierType);
	oView->GetTimeLine()->GetTierArea()->GetMiddleTier()->SetTierType(MiddleTierType);
	oView->GetTimeLine()->GetTierArea()->GetLowerTier()->SetTierType(LowerTierType);
	oView->SetKey(Key);
	mp_oCollection->m_Add(oView, Key, VIEWS_ADD_1, VIEWS_ADD_2, FALSE, VIEWS_ADD_3);
	return oView;
}

void clsViews::Clear(void)
{
	mp_oCollection->m_Clear();
	mp_oControl->SetCurrentView("");
}

void clsViews::Remove(CString Index)
{
	mp_oCollection->m_Remove(Index, VIEWS_REMOVE_1, VIEWS_REMOVE_2, VIEWS_REMOVE_3, VIEWS_REMOVE_4);
	mp_oControl->SetCurrentView("");
}

CString clsViews::GetXML(void)
{
	LONG lIndex;
	clsView* oView; 
	clsXML oXML(mp_oControl, "Views");
	oXML.InitializeWriter();
	for (lIndex = 1; lIndex <= GetCount(); lIndex++)
	{
		oView = (clsView*) mp_oCollection->m_oReturnArrayElement(lIndex);
		oXML.WriteObject(oView->GetXML());
	}
	return oXML.GetXML();
}

void clsViews::SetXML(CString sXML)
{
	LONG lIndex;
	clsXML oXML(mp_oControl, "Views");
	oXML.SetXML(sXML);
	oXML.InitializeReader();
	mp_oCollection->m_Clear();
	for (lIndex = 1;lIndex <= oXML.ReadCollectionCount();lIndex++)
	{
		clsView* oView = new clsView(mp_oControl);
		oView->SetXML(oXML.ReadCollectionObject(lIndex));
		mp_oCollection->SetAddMode(TRUE);
		mp_oCollection->m_Add(oView, oView->mp_sKey, VIEWS_ADD_1, VIEWS_ADD_2, FALSE, VIEWS_ADD_3);
	}
}