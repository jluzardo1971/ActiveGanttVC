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

//{D067F180-D6F5-47DC-89E2-E15024229539}
static const IID IID_IOutlineCodeValue = { 0xD067F180, 0xD6F5, 0x47DC, { 0x89, 0xE2, 0xE1, 0x50, 0x24, 0x22, 0x95, 0x39} };

//{C52E7FE9-44FF-42AA-B196-1A69A12FB56A}
IMPLEMENT_OLECREATE_FLAGS(OutlineCodeValue, "MSP2007.OutlineCodeValue", afxRegApartmentThreading, 0xC52E7FE9, 0x44FF, 0x42AA, 0xB1, 0x96, 0x1A, 0x69, 0xA1, 0x2F, 0xB5, 0x6A)

BEGIN_DISPATCH_MAP(OutlineCodeValue, CCmdTarget)
DISP_PROPERTY_EX_ID(OutlineCodeValue, "lValueID", 1, odl_GetValueID, odl_SetValueID, VT_I4)
DISP_PROPERTY_EX_ID(OutlineCodeValue, "sFieldGUID", 2, odl_GetFieldGUID, odl_SetFieldGUID, VT_BSTR)
DISP_PROPERTY_EX_ID(OutlineCodeValue, "yType", 3, odl_GetType, odl_SetType, VT_I4)
DISP_PROPERTY_EX_ID(OutlineCodeValue, "lParentValueID", 4, odl_GetParentValueID, odl_SetParentValueID, VT_I4)
DISP_PROPERTY_EX_ID(OutlineCodeValue, "sValue", 5, odl_GetValue, odl_SetValue, VT_BSTR)
DISP_PROPERTY_EX_ID(OutlineCodeValue, "sDescription", 6, odl_GetDescription, odl_SetDescription, VT_BSTR)
DISP_PROPERTY_EX_ID(OutlineCodeValue, "Key", 7, odl_GetKey, odl_SetKey, VT_BSTR)
DISP_FUNCTION_ID(OutlineCodeValue, "GetXML", 8, odl_GetXML, VT_BSTR, VTS_NONE)
DISP_FUNCTION_ID(OutlineCodeValue, "SetXML", 9, odl_SetXML, VT_EMPTY, VTS_BSTR)
DISP_FUNCTION_ID(OutlineCodeValue, "IsNull", 10, IsNull, VT_BOOL, VTS_NONE)
DISP_FUNCTION_ID(OutlineCodeValue, "Initialize", 11, Initialize, VT_EMPTY, VTS_NONE)
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
	oXML.ReadProperty("ParentValueID", mp_lParentValueID);
	oXML.ReadProperty("Value", mp_sValue);
	oXML.ReadProperty("Description", mp_sDescription);
}