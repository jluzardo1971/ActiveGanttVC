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
#include "OutlineCodeValue.h"

IMPLEMENT_DYNCREATE(OutlineCodeValue, CCmdTarget)

//{05FE0DBC-E7FC-4BCE-B79A-E44472EC0F47}
static const IID IID_IOutlineCodeValue = { 0x05FE0DBC, 0xE7FC, 0x4BCE, { 0xB7, 0x9A, 0xE4, 0x44, 0x72, 0xEC, 0x0F, 0x47} };

//{47D89C4B-7DC5-4485-888F-B37605182910}
IMPLEMENT_OLECREATE_FLAGS(OutlineCodeValue, "MSP2010.OutlineCodeValue", afxRegApartmentThreading, 0x47D89C4B, 0x7DC5, 0x4485, 0x88, 0x8F, 0xB3, 0x76, 0x05, 0x18, 0x29, 0x10)

BEGIN_DISPATCH_MAP(OutlineCodeValue, CCmdTarget)
DISP_PROPERTY_EX_ID(OutlineCodeValue, "lValueID", 1, odl_GetValueID, odl_SetValueID, VT_I4)
DISP_PROPERTY_EX_ID(OutlineCodeValue, "sFieldGUID", 2, odl_GetFieldGUID, odl_SetFieldGUID, VT_BSTR)
DISP_PROPERTY_EX_ID(OutlineCodeValue, "yType", 3, odl_GetType, odl_SetType, VT_I4)
DISP_PROPERTY_EX_ID(OutlineCodeValue, "bIsCollapsed", 4, odl_GetIsCollapsed, odl_SetIsCollapsed, VT_BOOL)
DISP_PROPERTY_EX_ID(OutlineCodeValue, "lParentValueID", 5, odl_GetParentValueID, odl_SetParentValueID, VT_I4)
DISP_PROPERTY_EX_ID(OutlineCodeValue, "sValue", 6, odl_GetValue, odl_SetValue, VT_BSTR)
DISP_PROPERTY_EX_ID(OutlineCodeValue, "sDescription", 7, odl_GetDescription, odl_SetDescription, VT_BSTR)
DISP_PROPERTY_EX_ID(OutlineCodeValue, "Key", 8, odl_GetKey, odl_SetKey, VT_BSTR)
DISP_FUNCTION_ID(OutlineCodeValue, "GetXML", 9, odl_GetXML, VT_BSTR, VTS_NONE)
DISP_FUNCTION_ID(OutlineCodeValue, "SetXML", 10, odl_SetXML, VT_EMPTY, VTS_BSTR)
DISP_FUNCTION_ID(OutlineCodeValue, "IsNull", 11, IsNull, VT_BOOL, VTS_NONE)
DISP_FUNCTION_ID(OutlineCodeValue, "Initialize", 12, Initialize, VT_EMPTY, VTS_NONE)
END_DISPATCH_MAP()

BEGIN_INTERFACE_MAP(OutlineCodeValue, CCmdTarget)
	INTERFACE_PART(OutlineCodeValue, IID_IOutlineCodeValue, Dispatch)
END_INTERFACE_MAP()

BEGIN_MESSAGE_MAP(OutlineCodeValue, CCmdTarget)
END_MESSAGE_MAP()


void OutlineCodeValue::Initialize(void)
{
	InitVars();
}

OutlineCodeValue::OutlineCodeValue()
{
	EnableAutomation();
	AfxOleLockApp();
	InitVars();
}

void OutlineCodeValue::InitVars(void)
{
	mp_lValueID = 0;
	mp_sFieldGUID = "";
	mp_yType = T_DATE;
	mp_bIsCollapsed = FALSE;
	mp_lParentValueID = 0;
	mp_sValue = "";
	mp_sDescription = "";
}

OutlineCodeValue::~OutlineCodeValue()
{
	AfxOleUnlockApp();
}

void OutlineCodeValue::OnFinalRelease()
{
	clsItemBase::OnFinalRelease();
}

LONG OutlineCodeValue::odl_GetValueID(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetValueID();
}

LONG OutlineCodeValue::GetValueID(void)
{
	return (LONG) mp_lValueID;
}

void OutlineCodeValue::odl_SetValueID(LONG newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetValueID((LONG)newVal);
}

void OutlineCodeValue::SetValueID(LONG newVal)
{
	mp_lValueID = newVal;
}

BSTR OutlineCodeValue::odl_GetFieldGUID(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetFieldGUID().AllocSysString();
}

CString OutlineCodeValue::GetFieldGUID(void)
{
	return (CString) mp_sFieldGUID;
}

void OutlineCodeValue::odl_SetFieldGUID(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetFieldGUID(newVal);
}

void OutlineCodeValue::SetFieldGUID(CString newVal)
{
	mp_sFieldGUID = newVal;
}

LONG OutlineCodeValue::odl_GetType(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetType();
}

E_TYPE OutlineCodeValue::GetType(void)
{
	return (E_TYPE) mp_yType;
}

void OutlineCodeValue::odl_SetType(LONG newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetType((E_TYPE)newVal);
}

void OutlineCodeValue::SetType(E_TYPE newVal)
{
	mp_yType = newVal;
}

BOOL OutlineCodeValue::odl_GetIsCollapsed(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetIsCollapsed();
}

BOOL OutlineCodeValue::GetIsCollapsed(void)
{
	return (BOOL) mp_bIsCollapsed;
}

void OutlineCodeValue::odl_SetIsCollapsed(BOOL newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetIsCollapsed((BOOL)newVal);
}

void OutlineCodeValue::SetIsCollapsed(BOOL newVal)
{
	mp_bIsCollapsed = newVal;
}

LONG OutlineCodeValue::odl_GetParentValueID(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetParentValueID();
}

LONG OutlineCodeValue::GetParentValueID(void)
{
	return (LONG) mp_lParentValueID;
}

void OutlineCodeValue::odl_SetParentValueID(LONG newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetParentValueID((LONG)newVal);
}

void OutlineCodeValue::SetParentValueID(LONG newVal)
{
	mp_lParentValueID = newVal;
}

BSTR OutlineCodeValue::odl_GetValue(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetValue().AllocSysString();
}

CString OutlineCodeValue::GetValue(void)
{
	return (CString) mp_sValue;
}

void OutlineCodeValue::odl_SetValue(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetValue(newVal);
}

void OutlineCodeValue::SetValue(CString newVal)
{
	mp_sValue = newVal;
}

BSTR OutlineCodeValue::odl_GetDescription(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetDescription().AllocSysString();
}

CString OutlineCodeValue::GetDescription(void)
{
	return (CString) mp_sDescription;
}

void OutlineCodeValue::odl_SetDescription(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetDescription(newVal);
}

void OutlineCodeValue::SetDescription(CString newVal)
{
	mp_sDescription = newVal;
}

BSTR OutlineCodeValue::odl_GetKey(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetKey().AllocSysString();
}

void OutlineCodeValue::odl_SetKey(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	USES_CONVERSION;
	SetKey(F_TOUCSTR(T2A(newVal)));
}

CString OutlineCodeValue::GetKey(void)
{
	return mp_sKey;
}

void OutlineCodeValue::SetKey(CString newVal)
{
	mp_oCollection->mp_SetKey(&mp_sKey, newVal, MP_SET_KEY);
}

BSTR OutlineCodeValue::odl_GetXML(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetXML().AllocSysString();
}

void OutlineCodeValue::odl_SetXML(LPCTSTR sXML)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetXML(sXML);
}

BOOL OutlineCodeValue::IsNull(void)
{
	BOOL bReturn = TRUE;
	if (mp_lValueID != 0)
	{
		bReturn = FALSE;
	}
	if (mp_sFieldGUID != "")
	{
		bReturn = FALSE;
	}
	if (mp_yType != T_DATE)
	{
		bReturn = FALSE;
	}
	if (mp_bIsCollapsed != FALSE)
	{
		bReturn = FALSE;
	}
	if (mp_lParentValueID != 0)
	{
		bReturn = FALSE;
	}
	if (mp_sValue != "")
	{
		bReturn = FALSE;
	}
	if (mp_sDescription != "")
	{
		bReturn = FALSE;
	}
	return bReturn;
}

CString OutlineCodeValue::GetXML(void)
{
	if (IsNull() == TRUE)
	{
		return "<Value/>";
	}
	clsXML oXML("Value");
	oXML.InitializeWriter();
	oXML.SetSupportOptional(TRUE);
	oXML.SetBoolsAreNumeric(TRUE);
	oXML.WriteProperty("ValueID", mp_lValueID);
	oXML.WriteProperty("FieldGUID", mp_sFieldGUID);
	oXML.WriteProperty("Type", mp_yType);
	oXML.WriteProperty("IsCollapsed", mp_bIsCollapsed);
	oXML.WriteProperty("ParentValueID", mp_lParentValueID);
	if (mp_sValue != "")
	{
		oXML.WriteProperty("Value", mp_sValue);
	}
	if (mp_sDescription != "")
	{
		oXML.WriteProperty("Description", mp_sDescription);
	}
	return oXML.GetXML();
}

void OutlineCodeValue::SetXML(CString sXML)
{
	clsXML oXML("Value");
	oXML.SetSupportOptional(TRUE);
	oXML.SetXML(sXML);
	oXML.InitializeReader();
	oXML.ReadProperty("ValueID", mp_lValueID);
	oXML.ReadProperty("FieldGUID", mp_sFieldGUID);
	oXML.ReadProperty("Type", mp_yType);
	oXML.ReadProperty("IsCollapsed", mp_bIsCollapsed);
	oXML.ReadProperty("ParentValueID", mp_lParentValueID);
	oXML.ReadProperty("Value", mp_sValue);
	oXML.ReadProperty("Description", mp_sDescription);
}