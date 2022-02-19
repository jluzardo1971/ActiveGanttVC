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
#include "OutlineCodeMasks.h"

IMPLEMENT_DYNCREATE(OutlineCodeMasks, CCmdTarget)

//{DADFAE4C-F26C-458D-85D3-92038A5BB6C7}
static const IID IID_IOutlineCodeMasks = { 0xDADFAE4C, 0xF26C, 0x458D, { 0x85, 0xD3, 0x92, 0x03, 0x8A, 0x5B, 0xB6, 0xC7} };

//{69771163-ACC7-4D5C-9F72-7E1E7B3CFBE0}
IMPLEMENT_OLECREATE_FLAGS(OutlineCodeMasks, "MSP2010.OutlineCodeMasks", afxRegApartmentThreading, 0x69771163, 0xACC7, 0x4D5C, 0x9F, 0x72, 0x7E, 0x1E, 0x7B, 0x3C, 0xFB, 0xE0)

BEGIN_DISPATCH_MAP(OutlineCodeMasks, CCmdTarget)
DISP_PROPERTY_EX_ID(OutlineCodeMasks, "Count", 1, odl_GetCount, SetNotSupported, VT_I4)
DISP_PROPERTY_PARAM_ID(OutlineCodeMasks, "Item", 2, odl_Item, SetNotSupported, VT_DISPATCH, VTS_BSTR)
DISP_FUNCTION_ID(OutlineCodeMasks, "Add", 3, odl_Add, VT_DISPATCH, VTS_NONE)
DISP_FUNCTION_ID(OutlineCodeMasks, "Clear", 4, odl_Clear, VT_EMPTY, VTS_NONE)
DISP_FUNCTION_ID(OutlineCodeMasks, "Remove", 5, odl_Remove, VT_EMPTY, VTS_BSTR)
DISP_FUNCTION_ID(OutlineCodeMasks, "IsNull", 6, IsNull, VT_BOOL, VTS_NONE)
DISP_FUNCTION_ID(OutlineCodeMasks, "Initialize", 7, Initialize, VT_EMPTY, VTS_NONE)
DISP_FUNCTION_ID(OutlineCodeMasks, "SetXML", 8, odl_SetXML, VT_EMPTY, VTS_BSTR)
DISP_FUNCTION_ID(OutlineCodeMasks, "GetXML", 9, odl_GetXML, VT_BSTR, VTS_NONE)
DISP_DEFVALUE(OutlineCodeMasks, "Item")
END_DISPATCH_MAP()

BEGIN_INTERFACE_MAP(OutlineCodeMasks, CCmdTarget)
	INTERFACE_PART(OutlineCodeMasks, IID_IOutlineCodeMasks, Dispatch)
END_INTERFACE_MAP()

BEGIN_MESSAGE_MAP(OutlineCodeMasks, CCmdTarget)
END_MESSAGE_MAP()


OutlineCodeMasks::OutlineCodeMasks()
{
	EnableAutomation();
	AfxOleLockApp();
	InitVars();
}

void OutlineCodeMasks::Initialize(void)
{
	InitVars();
}

void OutlineCodeMasks::InitVars(void)
{
	mp_oCollection = new clsCollectionBase("OutlineCodeMask");
}

OutlineCodeMasks::~OutlineCodeMasks()
{
	delete mp_oCollection;
	AfxOleUnlockApp();
}

void OutlineCodeMasks::OnFinalRelease()
{
	CCmdTarget::OnFinalRelease();
}

LONG OutlineCodeMasks::odl_GetCount(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetCount();
}

LONG OutlineCodeMasks::GetCount(void)
{
	return mp_oCollection->m_lCount();
}

IDispatch* OutlineCodeMasks::odl_Item(LPCTSTR Index)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return Item(Index)->GetIDispatch(TRUE);
}

IDispatch* OutlineCodeMasks::odl_Add(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return Add()->GetIDispatch(TRUE);
}

void OutlineCodeMasks::odl_Clear(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	Clear();
}

void OutlineCodeMasks::odl_Remove(LPCTSTR Index)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	Remove(Index);
}

BOOL OutlineCodeMasks::IsNull(void)
{
	BOOL bReturn = TRUE;
	if (GetCount() > 0)
	{
		bReturn = FALSE;
	}
	return bReturn;
}

OutlineCodeMask* OutlineCodeMasks::Item(CString Index)
{
	OutlineCodeMask *oOutlineCodeMask;
	oOutlineCodeMask = (OutlineCodeMask*)mp_oCollection->m_oItem(Index, MP_ITEM_1, MP_ITEM_2, MP_ITEM_3, MP_ITEM_4);
	return oOutlineCodeMask;
}

OutlineCodeMask* OutlineCodeMasks::Add()
{
	mp_oCollection->SetAddMode(TRUE);
	OutlineCodeMask* oOutlineCodeMask = new OutlineCodeMask();
	oOutlineCodeMask->mp_oCollection = mp_oCollection;
	mp_oCollection->m_Add(oOutlineCodeMask, _T(""), MP_ADD_1, MP_ADD_2, FALSE, MP_ADD_3);
	return oOutlineCodeMask;
}

void OutlineCodeMasks::Clear(void)
{
	mp_oCollection->m_Clear();
}

void OutlineCodeMasks::Remove(CString Index)
{
	mp_oCollection->m_Remove(Index, MP_REMOVE_1, MP_REMOVE_2, MP_REMOVE_3, MP_REMOVE_4);
}

BSTR OutlineCodeMasks::odl_GetXML(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetXML().AllocSysString();
}

void OutlineCodeMasks::odl_SetXML(LPCTSTR sXML)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetXML(sXML);
}

CString OutlineCodeMasks::GetXML(void)
{
	if (IsNull() == TRUE)
	{
		return "<Masks/>";
	}
	LONG lIndex;
	OutlineCodeMask* oOutlineCodeMask;
	clsXML oXML("Masks");
	oXML.InitializeWriter();
	for (lIndex = 1; lIndex <= GetCount(); lIndex++)
	{
		oOutlineCodeMask = (OutlineCodeMask*) mp_oCollection->m_oReturnArrayElement(lIndex);
		oXML.WriteObject(oOutlineCodeMask->GetXML());
	}
	return oXML.GetXML();
}

void OutlineCodeMasks::SetXML(CString sXML)
{
	LONG lIndex;
	clsXML oXML("Masks");
	oXML.SetSupportOptional(TRUE);
	oXML.SetXML(sXML);
	oXML.InitializeReader();
	mp_oCollection->m_Clear();
	for (lIndex = 1;lIndex <= oXML.ReadCollectionCount();lIndex++)
	{
		OutlineCodeMask* oOutlineCodeMask = new OutlineCodeMask();
		oOutlineCodeMask->SetXML(oXML.ReadCollectionObject(lIndex));
		mp_oCollection->SetAddMode(TRUE);
		CString sKey = _T("");
		oOutlineCodeMask->mp_oCollection = mp_oCollection;
		mp_oCollection->m_Add(oOutlineCodeMask, sKey, MP_ADD_1, MP_ADD_2, FALSE, MP_ADD_3);
	}
}