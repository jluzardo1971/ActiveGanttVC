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
#include "OutlineCodes.h"

IMPLEMENT_DYNCREATE(OutlineCodes, CCmdTarget)

//{F64748F5-7C89-41D7-9B05-8A0ACAC85D88}
static const IID IID_IOutlineCodes = { 0xF64748F5, 0x7C89, 0x41D7, { 0x9B, 0x05, 0x8A, 0x0A, 0xCA, 0xC8, 0x5D, 0x88} };

//{2377E9AF-3DE0-4D84-90FB-27C2A68931EC}
IMPLEMENT_OLECREATE_FLAGS(OutlineCodes, "MSP2003.OutlineCodes", afxRegApartmentThreading, 0x2377E9AF, 0x3DE0, 0x4D84, 0x90, 0xFB, 0x27, 0xC2, 0xA6, 0x89, 0x31, 0xEC)

BEGIN_DISPATCH_MAP(OutlineCodes, CCmdTarget)
DISP_PROPERTY_EX_ID(OutlineCodes, "Count", 1, odl_GetCount, SetNotSupported, VT_I4)
DISP_PROPERTY_PARAM_ID(OutlineCodes, "Item", 2, odl_Item, SetNotSupported, VT_DISPATCH, VTS_BSTR)
DISP_FUNCTION_ID(OutlineCodes, "Add", 3, odl_Add, VT_DISPATCH, VTS_NONE)
DISP_FUNCTION_ID(OutlineCodes, "Clear", 4, odl_Clear, VT_EMPTY, VTS_NONE)
DISP_FUNCTION_ID(OutlineCodes, "Remove", 5, odl_Remove, VT_EMPTY, VTS_BSTR)
DISP_FUNCTION_ID(OutlineCodes, "IsNull", 6, IsNull, VT_BOOL, VTS_NONE)
DISP_FUNCTION_ID(OutlineCodes, "Initialize", 7, Initialize, VT_EMPTY, VTS_NONE)
DISP_FUNCTION_ID(OutlineCodes, "SetXML", 8, odl_SetXML, VT_EMPTY, VTS_BSTR)
DISP_FUNCTION_ID(OutlineCodes, "GetXML", 9, odl_GetXML, VT_BSTR, VTS_NONE)
DISP_DEFVALUE(OutlineCodes, "Item")
END_DISPATCH_MAP()

BEGIN_INTERFACE_MAP(OutlineCodes, CCmdTarget)
	INTERFACE_PART(OutlineCodes, IID_IOutlineCodes, Dispatch)
END_INTERFACE_MAP()

BEGIN_MESSAGE_MAP(OutlineCodes, CCmdTarget)
END_MESSAGE_MAP()


OutlineCodes::OutlineCodes()
{
	EnableAutomation();
	AfxOleLockApp();
	InitVars();
}

void OutlineCodes::Initialize(void)
{
	InitVars();
}

void OutlineCodes::InitVars(void)
{
	mp_oCollection = new clsCollectionBase("OutlineCode");
}

OutlineCodes::~OutlineCodes()
{
	delete mp_oCollection;
	AfxOleUnlockApp();
}

void OutlineCodes::OnFinalRelease()
{
	CCmdTarget::OnFinalRelease();
}

LONG OutlineCodes::odl_GetCount(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetCount();
}

LONG OutlineCodes::GetCount(void)
{
	return mp_oCollection->m_lCount();
}

IDispatch* OutlineCodes::odl_Item(LPCTSTR Index)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return Item(Index)->GetIDispatch(TRUE);
}

IDispatch* OutlineCodes::odl_Add(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return Add()->GetIDispatch(TRUE);
}

void OutlineCodes::odl_Clear(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	Clear();
}

void OutlineCodes::odl_Remove(LPCTSTR Index)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	Remove(Index);
}

BOOL OutlineCodes::IsNull(void)
{
	BOOL bReturn = TRUE;
	if (GetCount() > 0)
	{
		bReturn = FALSE;
	}
	return bReturn;
}

OutlineCode* OutlineCodes::Item(CString Index)
{
	OutlineCode *oOutlineCode;
	oOutlineCode = (OutlineCode*)mp_oCollection->m_oItem(Index, MP_ITEM_1, MP_ITEM_2, MP_ITEM_3, MP_ITEM_4);
	return oOutlineCode;
}

OutlineCode* OutlineCodes::Add()
{
	mp_oCollection->SetAddMode(TRUE);
	OutlineCode* oOutlineCode = new OutlineCode();
	oOutlineCode->mp_oCollection = mp_oCollection;
	mp_oCollection->m_Add(oOutlineCode, _T(""), MP_ADD_1, MP_ADD_2, FALSE, MP_ADD_3);
	return oOutlineCode;
}

void OutlineCodes::Clear(void)
{
	mp_oCollection->m_Clear();
}

void OutlineCodes::Remove(CString Index)
{
	mp_oCollection->m_Remove(Index, MP_REMOVE_1, MP_REMOVE_2, MP_REMOVE_3, MP_REMOVE_4);
}

BSTR OutlineCodes::odl_GetXML(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetXML().AllocSysString();
}

void OutlineCodes::odl_SetXML(LPCTSTR sXML)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetXML(sXML);
}

CString OutlineCodes::GetXML(void)
{
	if (IsNull() == TRUE)
	{
		return "<OutlineCodes/>";
	}
	LONG lIndex;
	OutlineCode* oOutlineCode;
	clsXML oXML("OutlineCodes");
	oXML.InitializeWriter();
	for (lIndex = 1; lIndex <= GetCount(); lIndex++)
	{
		oOutlineCode = (OutlineCode*) mp_oCollection->m_oReturnArrayElement(lIndex);
		oXML.WriteObject(oOutlineCode->GetXML());
	}
	return oXML.GetXML();
}

void OutlineCodes::SetXML(CString sXML)
{
	LONG lIndex;
	clsXML oXML("OutlineCodes");
	oXML.SetSupportOptional(TRUE);
	oXML.SetXML(sXML);
	oXML.InitializeReader();
	mp_oCollection->m_Clear();
	for (lIndex = 1;lIndex <= oXML.ReadCollectionCount();lIndex++)
	{
		OutlineCode* oOutlineCode = new OutlineCode();
		oOutlineCode->SetXML(oXML.ReadCollectionObject(lIndex));
		mp_oCollection->SetAddMode(TRUE);
		CString sKey = _T("");
		oOutlineCode->mp_oCollection = mp_oCollection;
		mp_oCollection->m_Add(oOutlineCode, sKey, MP_ADD_1, MP_ADD_2, FALSE, MP_ADD_3);
	}
}