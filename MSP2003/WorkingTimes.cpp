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
#include "WorkingTimes.h"

IMPLEMENT_DYNCREATE(WorkingTimes, CCmdTarget)

//{6B91250E-48D5-4FC5-AC78-01AF9D147968}
static const IID IID_IWorkingTimes = { 0x6B91250E, 0x48D5, 0x4FC5, { 0xAC, 0x78, 0x01, 0xAF, 0x9D, 0x14, 0x79, 0x68} };

//{E4B056B3-0CE4-4E1C-91F8-B15CEB3684CE}
IMPLEMENT_OLECREATE_FLAGS(WorkingTimes, "MSP2003.WorkingTimes", afxRegApartmentThreading, 0xE4B056B3, 0x0CE4, 0x4E1C, 0x91, 0xF8, 0xB1, 0x5C, 0xEB, 0x36, 0x84, 0xCE)

BEGIN_DISPATCH_MAP(WorkingTimes, CCmdTarget)
DISP_PROPERTY_EX_ID(WorkingTimes, "Count", 1, odl_GetCount, SetNotSupported, VT_I4)
DISP_PROPERTY_PARAM_ID(WorkingTimes, "Item", 2, odl_Item, SetNotSupported, VT_DISPATCH, VTS_BSTR)
DISP_FUNCTION_ID(WorkingTimes, "Add", 3, odl_Add, VT_DISPATCH, VTS_NONE)
DISP_FUNCTION_ID(WorkingTimes, "Clear", 4, odl_Clear, VT_EMPTY, VTS_NONE)
DISP_FUNCTION_ID(WorkingTimes, "Remove", 5, odl_Remove, VT_EMPTY, VTS_BSTR)
DISP_FUNCTION_ID(WorkingTimes, "IsNull", 6, IsNull, VT_BOOL, VTS_NONE)
DISP_FUNCTION_ID(WorkingTimes, "Initialize", 7, Initialize, VT_EMPTY, VTS_NONE)
DISP_FUNCTION_ID(WorkingTimes, "SetXML", 8, odl_SetXML, VT_EMPTY, VTS_BSTR)
DISP_FUNCTION_ID(WorkingTimes, "GetXML", 9, odl_GetXML, VT_BSTR, VTS_NONE)
DISP_DEFVALUE(WorkingTimes, "Item")
END_DISPATCH_MAP()

BEGIN_INTERFACE_MAP(WorkingTimes, CCmdTarget)
	INTERFACE_PART(WorkingTimes, IID_IWorkingTimes, Dispatch)
END_INTERFACE_MAP()

BEGIN_MESSAGE_MAP(WorkingTimes, CCmdTarget)
END_MESSAGE_MAP()


WorkingTimes::WorkingTimes()
{
	EnableAutomation();
	AfxOleLockApp();
	InitVars();
}

void WorkingTimes::Initialize(void)
{
	InitVars();
}

void WorkingTimes::InitVars(void)
{
	mp_oCollection = new clsCollectionBase("WorkingTime");
}

WorkingTimes::~WorkingTimes()
{
	delete mp_oCollection;
	AfxOleUnlockApp();
}

void WorkingTimes::OnFinalRelease()
{
	CCmdTarget::OnFinalRelease();
}

LONG WorkingTimes::odl_GetCount(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetCount();
}

LONG WorkingTimes::GetCount(void)
{
	return mp_oCollection->m_lCount();
}

IDispatch* WorkingTimes::odl_Item(LPCTSTR Index)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return Item(Index)->GetIDispatch(TRUE);
}

IDispatch* WorkingTimes::odl_Add(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return Add()->GetIDispatch(TRUE);
}

void WorkingTimes::odl_Clear(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	Clear();
}

void WorkingTimes::odl_Remove(LPCTSTR Index)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	Remove(Index);
}

BOOL WorkingTimes::IsNull(void)
{
	BOOL bReturn = TRUE;
	if (GetCount() > 0)
	{
		bReturn = FALSE;
	}
	return bReturn;
}

WorkingTime* WorkingTimes::Item(CString Index)
{
	WorkingTime *oWorkingTime;
	oWorkingTime = (WorkingTime*)mp_oCollection->m_oItem(Index, MP_ITEM_1, MP_ITEM_2, MP_ITEM_3, MP_ITEM_4);
	return oWorkingTime;
}

WorkingTime* WorkingTimes::Add()
{
	mp_oCollection->SetAddMode(TRUE);
	WorkingTime* oWorkingTime = new WorkingTime();
	oWorkingTime->mp_oCollection = mp_oCollection;
	mp_oCollection->m_Add(oWorkingTime, _T(""), MP_ADD_1, MP_ADD_2, FALSE, MP_ADD_3);
	return oWorkingTime;
}

void WorkingTimes::Clear(void)
{
	mp_oCollection->m_Clear();
}

void WorkingTimes::Remove(CString Index)
{
	mp_oCollection->m_Remove(Index, MP_REMOVE_1, MP_REMOVE_2, MP_REMOVE_3, MP_REMOVE_4);
}

BSTR WorkingTimes::odl_GetXML(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetXML().AllocSysString();
}

void WorkingTimes::odl_SetXML(LPCTSTR sXML)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetXML(sXML);
}

CString WorkingTimes::GetXML(void)
{
	if (IsNull() == TRUE)
	{
		return "<WorkingTimes/>";
	}
	LONG lIndex;
	WorkingTime* oWorkingTime;
	clsXML oXML("WorkingTimes");
	oXML.InitializeWriter();
	for (lIndex = 1; lIndex <= GetCount(); lIndex++)
	{
		oWorkingTime = (WorkingTime*) mp_oCollection->m_oReturnArrayElement(lIndex);
		oXML.WriteObject(oWorkingTime->GetXML());
	}
	return oXML.GetXML();
}

void WorkingTimes::SetXML(CString sXML)
{
	LONG lIndex;
	clsXML oXML("WorkingTimes");
	oXML.SetSupportOptional(TRUE);
	oXML.SetXML(sXML);
	oXML.InitializeReader();
	mp_oCollection->m_Clear();
	for (lIndex = 1;lIndex <= oXML.ReadCollectionCount();lIndex++)
	{
		WorkingTime* oWorkingTime = new WorkingTime();
		oWorkingTime->SetXML(oXML.ReadCollectionObject(lIndex));
		mp_oCollection->SetAddMode(TRUE);
		CString sKey = _T("");
		oWorkingTime->mp_oCollection = mp_oCollection;
		mp_oCollection->m_Add(oWorkingTime, sKey, MP_ADD_1, MP_ADD_2, FALSE, MP_ADD_3);
	}
}