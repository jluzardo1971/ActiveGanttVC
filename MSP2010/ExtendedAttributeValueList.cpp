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
#include "ExtendedAttributeValueList.h"

IMPLEMENT_DYNCREATE(ExtendedAttributeValueList, CCmdTarget)

//{F231EC1C-E43D-4169-B8D8-1184ADBAF7E7}
static const IID IID_IExtendedAttributeValueList = { 0xF231EC1C, 0xE43D, 0x4169, { 0xB8, 0xD8, 0x11, 0x84, 0xAD, 0xBA, 0xF7, 0xE7} };

//{C5B5EA42-B84E-4435-800A-D9EFA340EAAF}
IMPLEMENT_OLECREATE_FLAGS(ExtendedAttributeValueList, "MSP2010.ExtendedAttributeValueList", afxRegApartmentThreading, 0xC5B5EA42, 0xB84E, 0x4435, 0x80, 0x0A, 0xD9, 0xEF, 0xA3, 0x40, 0xEA, 0xAF)

BEGIN_DISPATCH_MAP(ExtendedAttributeValueList, CCmdTarget)
DISP_PROPERTY_EX_ID(ExtendedAttributeValueList, "Count", 1, odl_GetCount, SetNotSupported, VT_I4)
DISP_PROPERTY_PARAM_ID(ExtendedAttributeValueList, "Item", 2, odl_Item, SetNotSupported, VT_DISPATCH, VTS_BSTR)
DISP_FUNCTION_ID(ExtendedAttributeValueList, "Add", 3, odl_Add, VT_DISPATCH, VTS_NONE)
DISP_FUNCTION_ID(ExtendedAttributeValueList, "Clear", 4, odl_Clear, VT_EMPTY, VTS_NONE)
DISP_FUNCTION_ID(ExtendedAttributeValueList, "Remove", 5, odl_Remove, VT_EMPTY, VTS_BSTR)
DISP_FUNCTION_ID(ExtendedAttributeValueList, "IsNull", 6, IsNull, VT_BOOL, VTS_NONE)
DISP_FUNCTION_ID(ExtendedAttributeValueList, "Initialize", 7, Initialize, VT_EMPTY, VTS_NONE)
DISP_FUNCTION_ID(ExtendedAttributeValueList, "SetXML", 8, odl_SetXML, VT_EMPTY, VTS_BSTR)
DISP_FUNCTION_ID(ExtendedAttributeValueList, "GetXML", 9, odl_GetXML, VT_BSTR, VTS_NONE)
DISP_DEFVALUE(ExtendedAttributeValueList, "Item")
END_DISPATCH_MAP()

BEGIN_INTERFACE_MAP(ExtendedAttributeValueList, CCmdTarget)
	INTERFACE_PART(ExtendedAttributeValueList, IID_IExtendedAttributeValueList, Dispatch)
END_INTERFACE_MAP()

BEGIN_MESSAGE_MAP(ExtendedAttributeValueList, CCmdTarget)
END_MESSAGE_MAP()


ExtendedAttributeValueList::ExtendedAttributeValueList()
{
	EnableAutomation();
	AfxOleLockApp();
	InitVars();
}

void ExtendedAttributeValueList::Initialize(void)
{
	InitVars();
}

void ExtendedAttributeValueList::InitVars(void)
{
	mp_oCollection = new clsCollectionBase("ExtendedAttributeValue");
}

ExtendedAttributeValueList::~ExtendedAttributeValueList()
{
	delete mp_oCollection;
	AfxOleUnlockApp();
}

void ExtendedAttributeValueList::OnFinalRelease()
{
	CCmdTarget::OnFinalRelease();
}

LONG ExtendedAttributeValueList::odl_GetCount(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetCount();
}

LONG ExtendedAttributeValueList::GetCount(void)
{
	return mp_oCollection->m_lCount();
}

IDispatch* ExtendedAttributeValueList::odl_Item(LPCTSTR Index)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return Item(Index)->GetIDispatch(TRUE);
}

IDispatch* ExtendedAttributeValueList::odl_Add(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return Add()->GetIDispatch(TRUE);
}

void ExtendedAttributeValueList::odl_Clear(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	Clear();
}

void ExtendedAttributeValueList::odl_Remove(LPCTSTR Index)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	Remove(Index);
}

BOOL ExtendedAttributeValueList::IsNull(void)
{
	BOOL bReturn = TRUE;
	if (GetCount() > 0)
	{
		bReturn = FALSE;
	}
	return bReturn;
}

ExtendedAttributeValue* ExtendedAttributeValueList::Item(CString Index)
{
	ExtendedAttributeValue *oExtendedAttributeValue;
	oExtendedAttributeValue = (ExtendedAttributeValue*)mp_oCollection->m_oItem(Index, MP_ITEM_1, MP_ITEM_2, MP_ITEM_3, MP_ITEM_4);
	return oExtendedAttributeValue;
}

ExtendedAttributeValue* ExtendedAttributeValueList::Add()
{
	mp_oCollection->SetAddMode(TRUE);
	ExtendedAttributeValue* oExtendedAttributeValue = new ExtendedAttributeValue();
	oExtendedAttributeValue->mp_oCollection = mp_oCollection;
	mp_oCollection->m_Add(oExtendedAttributeValue, _T(""), MP_ADD_1, MP_ADD_2, FALSE, MP_ADD_3);
	return oExtendedAttributeValue;
}

void ExtendedAttributeValueList::Clear(void)
{
	mp_oCollection->m_Clear();
}

void ExtendedAttributeValueList::Remove(CString Index)
{
	mp_oCollection->m_Remove(Index, MP_REMOVE_1, MP_REMOVE_2, MP_REMOVE_3, MP_REMOVE_4);
}

BSTR ExtendedAttributeValueList::odl_GetXML(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetXML().AllocSysString();
}

void ExtendedAttributeValueList::odl_SetXML(LPCTSTR sXML)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetXML(sXML);
}

CString ExtendedAttributeValueList::GetXML(void)
{
	if (IsNull() == TRUE)
	{
		return "<ValueList/>";
	}
	LONG lIndex;
	ExtendedAttributeValue* oExtendedAttributeValue;
	clsXML oXML("ValueList");
	oXML.InitializeWriter();
	for (lIndex = 1; lIndex <= GetCount(); lIndex++)
	{
		oExtendedAttributeValue = (ExtendedAttributeValue*) mp_oCollection->m_oReturnArrayElement(lIndex);
		oXML.WriteObject(oExtendedAttributeValue->GetXML());
	}
	return oXML.GetXML();
}

void ExtendedAttributeValueList::SetXML(CString sXML)
{
	LONG lIndex;
	clsXML oXML("ValueList");
	oXML.SetSupportOptional(TRUE);
	oXML.SetXML(sXML);
	oXML.InitializeReader();
	mp_oCollection->m_Clear();
	for (lIndex = 1;lIndex <= oXML.ReadCollectionCount();lIndex++)
	{
		ExtendedAttributeValue* oExtendedAttributeValue = new ExtendedAttributeValue();
		oExtendedAttributeValue->SetXML(oXML.ReadCollectionObject(lIndex));
		mp_oCollection->SetAddMode(TRUE);
		CString sKey = _T("");
		oExtendedAttributeValue->mp_oCollection = mp_oCollection;
		mp_oCollection->m_Add(oExtendedAttributeValue, sKey, MP_ADD_1, MP_ADD_2, FALSE, MP_ADD_3);
	}
}