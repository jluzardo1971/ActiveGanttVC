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
#include "OutlineCodeValues.h"

IMPLEMENT_DYNCREATE(OutlineCodeValues, CCmdTarget)

//{BB6FF4C3-0627-4AF9-95AE-5FEFD8897205}
static const IID IID_IOutlineCodeValues = { 0xBB6FF4C3, 0x0627, 0x4AF9, { 0x95, 0xAE, 0x5F, 0xEF, 0xD8, 0x89, 0x72, 0x05} };

//{68FB69F0-72CC-48A8-9CA7-C73C80CDC634}
IMPLEMENT_OLECREATE_FLAGS(OutlineCodeValues, "MSP2003.OutlineCodeValues", afxRegApartmentThreading, 0x68FB69F0, 0x72CC, 0x48A8, 0x9C, 0xA7, 0xC7, 0x3C, 0x80, 0xCD, 0xC6, 0x34)

BEGIN_DISPATCH_MAP(OutlineCodeValues, CCmdTarget)
DISP_PROPERTY_EX_ID(OutlineCodeValues, "Count", 1, odl_GetCount, SetNotSupported, VT_I4)
DISP_PROPERTY_PARAM_ID(OutlineCodeValues, "Item", 2, odl_Item, SetNotSupported, VT_DISPATCH, VTS_BSTR)
DISP_FUNCTION_ID(OutlineCodeValues, "Add", 3, odl_Add, VT_DISPATCH, VTS_NONE)
DISP_FUNCTION_ID(OutlineCodeValues, "Clear", 4, odl_Clear, VT_EMPTY, VTS_NONE)
DISP_FUNCTION_ID(OutlineCodeValues, "Remove", 5, odl_Remove, VT_EMPTY, VTS_BSTR)
DISP_FUNCTION_ID(OutlineCodeValues, "IsNull", 6, IsNull, VT_BOOL, VTS_NONE)
DISP_FUNCTION_ID(OutlineCodeValues, "Initialize", 7, Initialize, VT_EMPTY, VTS_NONE)
DISP_FUNCTION_ID(OutlineCodeValues, "SetXML", 8, odl_SetXML, VT_EMPTY, VTS_BSTR)
DISP_FUNCTION_ID(OutlineCodeValues, "GetXML", 9, odl_GetXML, VT_BSTR, VTS_NONE)
DISP_DEFVALUE(OutlineCodeValues, "Item")
END_DISPATCH_MAP()

BEGIN_INTERFACE_MAP(OutlineCodeValues, CCmdTarget)
	INTERFACE_PART(OutlineCodeValues, IID_IOutlineCodeValues, Dispatch)
END_INTERFACE_MAP()

BEGIN_MESSAGE_MAP(OutlineCodeValues, CCmdTarget)
END_MESSAGE_MAP()


OutlineCodeValues::OutlineCodeValues()
{
	EnableAutomation();
	AfxOleLockApp();
	InitVars();
}

void OutlineCodeValues::Initialize(void)
{
	InitVars();
}

void OutlineCodeValues::InitVars(void)
{
	mp_oCollection = new clsCollectionBase("OutlineCodeValue");
}

OutlineCodeValues::~OutlineCodeValues()
{
	delete mp_oCollection;
	AfxOleUnlockApp();
}

void OutlineCodeValues::OnFinalRelease()
{
	CCmdTarget::OnFinalRelease();
}

LONG OutlineCodeValues::odl_GetCount(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetCount();
}

LONG OutlineCodeValues::GetCount(void)
{
	return mp_oCollection->m_lCount();
}

IDispatch* OutlineCodeValues::odl_Item(LPCTSTR Index)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return Item(Index)->GetIDispatch(TRUE);
}

IDispatch* OutlineCodeValues::odl_Add(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return Add()->GetIDispatch(TRUE);
}

void OutlineCodeValues::odl_Clear(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	Clear();
}

void OutlineCodeValues::odl_Remove(LPCTSTR Index)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	Remove(Index);
}

BOOL OutlineCodeValues::IsNull(void)
{
	BOOL bReturn = TRUE;
	if (GetCount() > 0)
	{
		bReturn = FALSE;
	}
	return bReturn;
}

OutlineCodeValue* OutlineCodeValues::Item(CString Index)
{
	OutlineCodeValue *oOutlineCodeValue;
	oOutlineCodeValue = (OutlineCodeValue*)mp_oCollection->m_oItem(Index, MP_ITEM_1, MP_ITEM_2, MP_ITEM_3, MP_ITEM_4);
	return oOutlineCodeValue;
}

OutlineCodeValue* OutlineCodeValues::Add()
{
	mp_oCollection->SetAddMode(TRUE);
	OutlineCodeValue* oOutlineCodeValue = new OutlineCodeValue();
	oOutlineCodeValue->mp_oCollection = mp_oCollection;
	mp_oCollection->m_Add(oOutlineCodeValue, _T(""), MP_ADD_1, MP_ADD_2, FALSE, MP_ADD_3);
	return oOutlineCodeValue;
}

void OutlineCodeValues::Clear(void)
{
	mp_oCollection->m_Clear();
}

void OutlineCodeValues::Remove(CString Index)
{
	mp_oCollection->m_Remove(Index, MP_REMOVE_1, MP_REMOVE_2, MP_REMOVE_3, MP_REMOVE_4);
}

BSTR OutlineCodeValues::odl_GetXML(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetXML().AllocSysString();
}

void OutlineCodeValues::odl_SetXML(LPCTSTR sXML)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetXML(sXML);
}

CString OutlineCodeValues::GetXML(void)
{
	if (IsNull() == TRUE)
	{
		return "<Values/>";
	}
	LONG lIndex;
	OutlineCodeValue* oOutlineCodeValue;
	clsXML oXML("Values");
	oXML.InitializeWriter();
	for (lIndex = 1; lIndex <= GetCount(); lIndex++)
	{
		oOutlineCodeValue = (OutlineCodeValue*) mp_oCollection->m_oReturnArrayElement(lIndex);
		oXML.WriteObject(oOutlineCodeValue->GetXML());
	}
	return oXML.GetXML();
}

void OutlineCodeValues::SetXML(CString sXML)
{
	LONG lIndex;
	clsXML oXML("Values");
	oXML.SetSupportOptional(TRUE);
	oXML.SetXML(sXML);
	oXML.InitializeReader();
	mp_oCollection->m_Clear();
	for (lIndex = 1;lIndex <= oXML.ReadCollectionCount();lIndex++)
	{
		OutlineCodeValue* oOutlineCodeValue = new OutlineCodeValue();
		oOutlineCodeValue->SetXML(oXML.ReadCollectionObject(lIndex));
		mp_oCollection->SetAddMode(TRUE);
		CString sKey = _T("");
		oOutlineCodeValue->mp_oCollection = mp_oCollection;
		mp_oCollection->m_Add(oOutlineCodeValue, sKey, MP_ADD_1, MP_ADD_2, FALSE, MP_ADD_3);
	}
}