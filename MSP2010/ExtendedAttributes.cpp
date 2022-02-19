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
#include "ExtendedAttributes.h"

IMPLEMENT_DYNCREATE(ExtendedAttributes, CCmdTarget)

//{7A5ADB3A-DD52-44E5-AD3F-12F95EA7155F}
static const IID IID_IExtendedAttributes = { 0x7A5ADB3A, 0xDD52, 0x44E5, { 0xAD, 0x3F, 0x12, 0xF9, 0x5E, 0xA7, 0x15, 0x5F} };

//{4E4B0593-C7EB-46F8-942D-6ADE2DE348EB}
IMPLEMENT_OLECREATE_FLAGS(ExtendedAttributes, "MSP2010.ExtendedAttributes", afxRegApartmentThreading, 0x4E4B0593, 0xC7EB, 0x46F8, 0x94, 0x2D, 0x6A, 0xDE, 0x2D, 0xE3, 0x48, 0xEB)

BEGIN_DISPATCH_MAP(ExtendedAttributes, CCmdTarget)
DISP_PROPERTY_EX_ID(ExtendedAttributes, "Count", 1, odl_GetCount, SetNotSupported, VT_I4)
DISP_PROPERTY_PARAM_ID(ExtendedAttributes, "Item", 2, odl_Item, SetNotSupported, VT_DISPATCH, VTS_BSTR)
DISP_FUNCTION_ID(ExtendedAttributes, "Add", 3, odl_Add, VT_DISPATCH, VTS_NONE)
DISP_FUNCTION_ID(ExtendedAttributes, "Clear", 4, odl_Clear, VT_EMPTY, VTS_NONE)
DISP_FUNCTION_ID(ExtendedAttributes, "Remove", 5, odl_Remove, VT_EMPTY, VTS_BSTR)
DISP_FUNCTION_ID(ExtendedAttributes, "IsNull", 6, IsNull, VT_BOOL, VTS_NONE)
DISP_FUNCTION_ID(ExtendedAttributes, "Initialize", 7, Initialize, VT_EMPTY, VTS_NONE)
DISP_FUNCTION_ID(ExtendedAttributes, "SetXML", 8, odl_SetXML, VT_EMPTY, VTS_BSTR)
DISP_FUNCTION_ID(ExtendedAttributes, "GetXML", 9, odl_GetXML, VT_BSTR, VTS_NONE)
DISP_DEFVALUE(ExtendedAttributes, "Item")
END_DISPATCH_MAP()

BEGIN_INTERFACE_MAP(ExtendedAttributes, CCmdTarget)
	INTERFACE_PART(ExtendedAttributes, IID_IExtendedAttributes, Dispatch)
END_INTERFACE_MAP()

BEGIN_MESSAGE_MAP(ExtendedAttributes, CCmdTarget)
END_MESSAGE_MAP()


ExtendedAttributes::ExtendedAttributes()
{
	EnableAutomation();
	AfxOleLockApp();
	InitVars();
}

void ExtendedAttributes::Initialize(void)
{
	InitVars();
}

void ExtendedAttributes::InitVars(void)
{
	mp_oCollection = new clsCollectionBase("ExtendedAttribute");
}

ExtendedAttributes::~ExtendedAttributes()
{
	delete mp_oCollection;
	AfxOleUnlockApp();
}

void ExtendedAttributes::OnFinalRelease()
{
	CCmdTarget::OnFinalRelease();
}

LONG ExtendedAttributes::odl_GetCount(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetCount();
}

LONG ExtendedAttributes::GetCount(void)
{
	return mp_oCollection->m_lCount();
}

IDispatch* ExtendedAttributes::odl_Item(LPCTSTR Index)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return Item(Index)->GetIDispatch(TRUE);
}

IDispatch* ExtendedAttributes::odl_Add(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return Add()->GetIDispatch(TRUE);
}

void ExtendedAttributes::odl_Clear(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	Clear();
}

void ExtendedAttributes::odl_Remove(LPCTSTR Index)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	Remove(Index);
}

BOOL ExtendedAttributes::IsNull(void)
{
	BOOL bReturn = TRUE;
	if (GetCount() > 0)
	{
		bReturn = FALSE;
	}
	return bReturn;
}

ExtendedAttribute* ExtendedAttributes::Item(CString Index)
{
	ExtendedAttribute *oExtendedAttribute;
	oExtendedAttribute = (ExtendedAttribute*)mp_oCollection->m_oItem(Index, MP_ITEM_1, MP_ITEM_2, MP_ITEM_3, MP_ITEM_4);
	return oExtendedAttribute;
}

ExtendedAttribute* ExtendedAttributes::Add()
{
	mp_oCollection->SetAddMode(TRUE);
	ExtendedAttribute* oExtendedAttribute = new ExtendedAttribute();
	oExtendedAttribute->mp_oCollection = mp_oCollection;
	mp_oCollection->m_Add(oExtendedAttribute, _T(""), MP_ADD_1, MP_ADD_2, FALSE, MP_ADD_3);
	return oExtendedAttribute;
}

void ExtendedAttributes::Clear(void)
{
	mp_oCollection->m_Clear();
}

void ExtendedAttributes::Remove(CString Index)
{
	mp_oCollection->m_Remove(Index, MP_REMOVE_1, MP_REMOVE_2, MP_REMOVE_3, MP_REMOVE_4);
}

BSTR ExtendedAttributes::odl_GetXML(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetXML().AllocSysString();
}

void ExtendedAttributes::odl_SetXML(LPCTSTR sXML)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetXML(sXML);
}

CString ExtendedAttributes::GetXML(void)
{
	if (IsNull() == TRUE)
	{
		return "<ExtendedAttributes/>";
	}
	LONG lIndex;
	ExtendedAttribute* oExtendedAttribute;
	clsXML oXML("ExtendedAttributes");
	oXML.InitializeWriter();
	for (lIndex = 1; lIndex <= GetCount(); lIndex++)
	{
		oExtendedAttribute = (ExtendedAttribute*) mp_oCollection->m_oReturnArrayElement(lIndex);
		oXML.WriteObject(oExtendedAttribute->GetXML());
	}
	return oXML.GetXML();
}

void ExtendedAttributes::SetXML(CString sXML)
{
	LONG lIndex;
	clsXML oXML("ExtendedAttributes");
	oXML.SetSupportOptional(TRUE);
	oXML.SetXML(sXML);
	oXML.InitializeReader();
	mp_oCollection->m_Clear();
	for (lIndex = 1;lIndex <= oXML.ReadCollectionCount();lIndex++)
	{
		ExtendedAttribute* oExtendedAttribute = new ExtendedAttribute();
		oExtendedAttribute->SetXML(oXML.ReadCollectionObject(lIndex));
		mp_oCollection->SetAddMode(TRUE);
		CString sKey = _T("");
		oExtendedAttribute->mp_oCollection = mp_oCollection;
		mp_oCollection->m_Add(oExtendedAttribute, sKey, MP_ADD_1, MP_ADD_2, FALSE, MP_ADD_3);
	}
}