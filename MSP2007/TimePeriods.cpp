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
#include "clsXML.h"
#include "TimePeriods.h"

IMPLEMENT_DYNCREATE(TimePeriods, CCmdTarget)

//{07DEA964-964B-4693-B5C5-1C678CED29C9}
static const IID IID_ITimePeriods = { 0x07DEA964, 0x964B, 0x4693, { 0xB5, 0xC5, 0x1C, 0x67, 0x8C, 0xED, 0x29, 0xC9} };

//{8065A0BB-DCC8-4613-951E-2387EA5C80E8}
IMPLEMENT_OLECREATE_FLAGS(TimePeriods, "MSP2007.TimePeriods", afxRegApartmentThreading, 0x8065A0BB, 0xDCC8, 0x4613, 0x95, 0x1E, 0x23, 0x87, 0xEA, 0x5C, 0x80, 0xE8)

BEGIN_DISPATCH_MAP(TimePeriods, CCmdTarget)
DISP_PROPERTY_EX_ID(TimePeriods, "Count", 1, odl_GetCount, SetNotSupported, VT_I4)
DISP_PROPERTY_PARAM_ID(TimePeriods, "Item", 2, odl_Item, SetNotSupported, VT_DISPATCH, VTS_BSTR)
DISP_FUNCTION_ID(TimePeriods, "Add", 3, odl_Add, VT_DISPATCH, VTS_NONE)
DISP_FUNCTION_ID(TimePeriods, "Clear", 4, odl_Clear, VT_EMPTY, VTS_NONE)
DISP_FUNCTION_ID(TimePeriods, "Remove", 5, odl_Remove, VT_EMPTY, VTS_BSTR)
DISP_FUNCTION_ID(TimePeriods, "IsNull", 6, IsNull, VT_BOOL, VTS_NONE)
DISP_FUNCTION_ID(TimePeriods, "Initialize", 7, Initialize, VT_EMPTY, VTS_NONE)
DISP_FUNCTION_ID(TimePeriods, "SetXML", 8, odl_SetXML, VT_EMPTY, VTS_BSTR)
DISP_FUNCTION_ID(TimePeriods, "GetXML", 9, odl_GetXML, VT_BSTR, VTS_NONE)
DISP_DEFVALUE(TimePeriods, "Item")
END_DISPATCH_MAP()

BEGIN_INTERFACE_MAP(TimePeriods, CCmdTarget)
	INTERFACE_PART(TimePeriods, IID_ITimePeriods, Dispatch)
END_INTERFACE_MAP()

BEGIN_MESSAGE_MAP(TimePeriods, CCmdTarget)
END_MESSAGE_MAP()


TimePeriods::TimePeriods()
{
	EnableAutomation();
	AfxOleLockApp();
	InitVars();
}

void TimePeriods::Initialize(void)
{
	InitVars();
}

void TimePeriods::InitVars(void)
{
	mp_oCollection = new clsCollectionBase("TimePeriod");
}

TimePeriods::~TimePeriods()
{
	delete mp_oCollection;
	AfxOleUnlockApp();
}

void TimePeriods::OnFinalRelease()
{
	CCmdTarget::OnFinalRelease();
}

LONG TimePeriods::odl_GetCount(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetCount();
}

LONG TimePeriods::GetCount(void)
{
	return mp_oCollection->m_lCount();
}

IDispatch* TimePeriods::odl_Item(LPCTSTR Index)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return Item(Index)->GetIDispatch(TRUE);
}

IDispatch* TimePeriods::odl_Add(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return Add()->GetIDispatch(TRUE);
}

void TimePeriods::odl_Clear(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	Clear();
}

void TimePeriods::odl_Remove(LPCTSTR Index)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	Remove(Index);
}

BOOL TimePeriods::IsNull(void)
{
	BOOL bReturn = TRUE;
	if (GetCount() > 0)
	{
		bReturn = FALSE;
	}
	return bReturn;
}

TimePeriod* TimePeriods::Item(CString Index)
{
	TimePeriod *oTimePeriod;
	oTimePeriod = (TimePeriod*)mp_oCollection->m_oItem(Index, MP_ITEM_1, MP_ITEM_2, MP_ITEM_3, MP_ITEM_4);
	return oTimePeriod;
}

TimePeriod* TimePeriods::Add()
{
	mp_oCollection->SetAddMode(TRUE);
	TimePeriod* oTimePeriod = new TimePeriod();
	oTimePeriod->mp_oCollection = mp_oCollection;
	mp_oCollection->m_Add(oTimePeriod, _T(""), MP_ADD_1, MP_ADD_2, FALSE, MP_ADD_3);
	return oTimePeriod;
}

void TimePeriods::Clear(void)
{
	mp_oCollection->m_Clear();
}

void TimePeriods::Remove(CString Index)
{
	mp_oCollection->m_Remove(Index, MP_REMOVE_1, MP_REMOVE_2, MP_REMOVE_3, MP_REMOVE_4);
}

BSTR TimePeriods::odl_GetXML(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetXML().AllocSysString();
}

void TimePeriods::odl_SetXML(LPCTSTR sXML)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetXML(sXML);
}

CString TimePeriods::GetXML(void)
{
	if (IsNull() == TRUE)
	{
		return "<TimePeriods/>";
	}
	LONG lIndex;
	TimePeriod* oTimePeriod;
	clsXML oXML("TimePeriods");
	oXML.InitializeWriter();
	for (lIndex = 1; lIndex <= GetCount(); lIndex++)
	{
		oTimePeriod = (TimePeriod*) mp_oCollection->m_oReturnArrayElement(lIndex);
		oXML.WriteObject(oTimePeriod->GetXML());
	}
	return oXML.GetXML();
}

void TimePeriods::SetXML(CString sXML)
{
	LONG lIndex;
	clsXML oXML("TimePeriods");
	oXML.SetSupportOptional(TRUE);
	oXML.SetXML(sXML);
	oXML.InitializeReader();
	mp_oCollection->m_Clear();
	for (lIndex = 1;lIndex <= oXML.ReadCollectionCount();lIndex++)
	{
		TimePeriod* oTimePeriod = new TimePeriod();
		oTimePeriod->SetXML(oXML.ReadCollectionObject(lIndex));
		mp_oCollection->SetAddMode(TRUE);
		CString sKey = _T("");
		oTimePeriod->mp_oCollection = mp_oCollection;
		mp_oCollection->m_Add(oTimePeriod, sKey, MP_ADD_1, MP_ADD_2, FALSE, MP_ADD_3);
	}
}