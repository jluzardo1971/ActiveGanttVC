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
#include "TimephasedData_C.h"

IMPLEMENT_DYNCREATE(TimephasedData_C, CCmdTarget)

//{A43ECACF-8DB9-4062-955C-924794B75258}
static const IID IID_ITimephasedData_C = { 0xA43ECACF, 0x8DB9, 0x4062, { 0x95, 0x5C, 0x92, 0x47, 0x94, 0xB7, 0x52, 0x58} };

//{074A1713-031B-45BB-8236-7901684E8D62}
IMPLEMENT_OLECREATE_FLAGS(TimephasedData_C, "MSP2003.TimephasedData_C", afxRegApartmentThreading, 0x074A1713, 0x031B, 0x45BB, 0x82, 0x36, 0x79, 0x01, 0x68, 0x4E, 0x8D, 0x62)

BEGIN_DISPATCH_MAP(TimephasedData_C, CCmdTarget)
DISP_PROPERTY_EX_ID(TimephasedData_C, "Count", 1, odl_GetCount, SetNotSupported, VT_I4)
DISP_PROPERTY_PARAM_ID(TimephasedData_C, "Item", 2, odl_Item, SetNotSupported, VT_DISPATCH, VTS_BSTR)
DISP_FUNCTION_ID(TimephasedData_C, "Add", 3, odl_Add, VT_DISPATCH, VTS_NONE)
DISP_FUNCTION_ID(TimephasedData_C, "Clear", 4, odl_Clear, VT_EMPTY, VTS_NONE)
DISP_FUNCTION_ID(TimephasedData_C, "Remove", 5, odl_Remove, VT_EMPTY, VTS_BSTR)
DISP_FUNCTION_ID(TimephasedData_C, "IsNull", 6, IsNull, VT_BOOL, VTS_NONE)
DISP_FUNCTION_ID(TimephasedData_C, "Initialize", 7, Initialize, VT_EMPTY, VTS_NONE)
DISP_DEFVALUE(TimephasedData_C, "Item")
END_DISPATCH_MAP()

BEGIN_INTERFACE_MAP(TimephasedData_C, CCmdTarget)
	INTERFACE_PART(TimephasedData_C, IID_ITimephasedData_C, Dispatch)
END_INTERFACE_MAP()

BEGIN_MESSAGE_MAP(TimephasedData_C, CCmdTarget)
END_MESSAGE_MAP()


TimephasedData_C::TimephasedData_C()
{
	EnableAutomation();
	AfxOleLockApp();
	InitVars();
}

void TimephasedData_C::Initialize(void)
{
	InitVars();
}

void TimephasedData_C::InitVars(void)
{
	mp_oCollection = new clsCollectionBase("TimephasedData");
}

TimephasedData_C::~TimephasedData_C()
{
	delete mp_oCollection;
	AfxOleUnlockApp();
}

void TimephasedData_C::OnFinalRelease()
{
	CCmdTarget::OnFinalRelease();
}

LONG TimephasedData_C::odl_GetCount(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetCount();
}

LONG TimephasedData_C::GetCount(void)
{
	return mp_oCollection->m_lCount();
}

IDispatch* TimephasedData_C::odl_Item(LPCTSTR Index)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return Item(Index)->GetIDispatch(TRUE);
}

IDispatch* TimephasedData_C::odl_Add(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return Add()->GetIDispatch(TRUE);
}

void TimephasedData_C::odl_Clear(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	Clear();
}

void TimephasedData_C::odl_Remove(LPCTSTR Index)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	Remove(Index);
}

BOOL TimephasedData_C::IsNull(void)
{
	BOOL bReturn = TRUE;
	if (GetCount() > 0)
	{
		bReturn = FALSE;
	}
	return bReturn;
}

TimephasedData* TimephasedData_C::Item(CString Index)
{
	TimephasedData *oTimephasedData;
	oTimephasedData = (TimephasedData*)mp_oCollection->m_oItem(Index, MP_ITEM_1, MP_ITEM_2, MP_ITEM_3, MP_ITEM_4);
	return oTimephasedData;
}

TimephasedData* TimephasedData_C::Add()
{
	mp_oCollection->SetAddMode(TRUE);
	TimephasedData* oTimephasedData = new TimephasedData();
	oTimephasedData->mp_oCollection = mp_oCollection;
	mp_oCollection->m_Add(oTimephasedData, _T(""), MP_ADD_1, MP_ADD_2, FALSE, MP_ADD_3);
	return oTimephasedData;
}

void TimephasedData_C::Clear(void)
{
	mp_oCollection->m_Clear();
}

void TimephasedData_C::Remove(CString Index)
{
	mp_oCollection->m_Remove(Index, MP_REMOVE_1, MP_REMOVE_2, MP_REMOVE_3, MP_REMOVE_4);
}

void TimephasedData_C::ReadObjectProtected(clsXML &oXML)
{
	LONG lIndex;
	for (lIndex = 1; lIndex <= oXML.ReadCollectionCount(); lIndex++)
	{
		if (oXML.GetCollectionObjectName(lIndex) == "TimephasedData")
		{
			TimephasedData* oTimephasedData = new TimephasedData();
			oTimephasedData->SetXML(oXML.ReadCollectionObject(lIndex));
			mp_oCollection->SetAddMode(TRUE);
			CString sKey = _T("");
			oTimephasedData->mp_oCollection = mp_oCollection;
			mp_oCollection->m_Add(oTimephasedData, sKey, MP_ADD_1, MP_ADD_2, FALSE, MP_ADD_3);
		}
	}
}

void TimephasedData_C::WriteObjectProtected(clsXML &oXML)
{
	LONG lIndex;
	TimephasedData* oTimephasedData;
	for (lIndex = 1; lIndex <= GetCount(); lIndex++)
	{
		oTimephasedData = (TimephasedData*) mp_oCollection->m_oReturnArrayElement(lIndex);
		oXML.WriteObject(oTimephasedData->GetXML());
	}
}