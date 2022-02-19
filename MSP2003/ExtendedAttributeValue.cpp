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
#include "ExtendedAttributeValue.h"

IMPLEMENT_DYNCREATE(ExtendedAttributeValue, CCmdTarget)

//{29C85DB0-FAAF-4F4C-A9DF-97E0D1DEA460}
static const IID IID_IExtendedAttributeValue = { 0x29C85DB0, 0xFAAF, 0x4F4C, { 0xA9, 0xDF, 0x97, 0xE0, 0xD1, 0xDE, 0xA4, 0x60} };

//{46616062-AF7C-4D4B-8BD0-D50DC52F1F13}
IMPLEMENT_OLECREATE_FLAGS(ExtendedAttributeValue, "MSP2003.ExtendedAttributeValue", afxRegApartmentThreading, 0x46616062, 0xAF7C, 0x4D4B, 0x8B, 0xD0, 0xD5, 0x0D, 0xC5, 0x2F, 0x1F, 0x13)

BEGIN_DISPATCH_MAP(ExtendedAttributeValue, CCmdTarget)
DISP_PROPERTY_EX_ID(ExtendedAttributeValue, "lID", 1, odl_GetID, odl_SetID, VT_I4)
DISP_PROPERTY_EX_ID(ExtendedAttributeValue, "sValue", 2, odl_GetValue, odl_SetValue, VT_BSTR)
DISP_PROPERTY_EX_ID(ExtendedAttributeValue, "sDescription", 3, odl_GetDescription, odl_SetDescription, VT_BSTR)
DISP_PROPERTY_EX_ID(ExtendedAttributeValue, "Key", 4, odl_GetKey, odl_SetKey, VT_BSTR)
DISP_FUNCTION_ID(ExtendedAttributeValue, "GetXML", 5, odl_GetXML, VT_BSTR, VTS_NONE)
DISP_FUNCTION_ID(ExtendedAttributeValue, "SetXML", 6, odl_SetXML, VT_EMPTY, VTS_BSTR)
DISP_FUNCTION_ID(ExtendedAttributeValue, "IsNull", 7, IsNull, VT_BOOL, VTS_NONE)
DISP_FUNCTION_ID(ExtendedAttributeValue, "Initialize", 8, Initialize, VT_EMPTY, VTS_NONE)
END_DISPATCH_MAP()

BEGIN_INTERFACE_MAP(ExtendedAttributeValue, CCmdTarget)
	INTERFACE_PART(ExtendedAttributeValue, IID_IExtendedAttributeValue, Dispatch)
END_INTERFACE_MAP()

BEGIN_MESSAGE_MAP(ExtendedAttributeValue, CCmdTarget)
END_MESSAGE_MAP()


void ExtendedAttributeValue::Initialize(void)
{
	InitVars();
}

ExtendedAttributeValue::ExtendedAttributeValue()
{
	EnableAutomation();
	AfxOleLockApp();
	InitVars();
}

void ExtendedAttributeValue::InitVars(void)
{
	mp_lID = 0;
	mp_sValue = "";
	mp_sDescription = "";
}

ExtendedAttributeValue::~ExtendedAttributeValue()
{
	AfxOleUnlockApp();
}

void ExtendedAttributeValue::OnFinalRelease()
{
	clsItemBase::OnFinalRelease();
}

LONG ExtendedAttributeValue::odl_GetID(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetID();
}

LONG ExtendedAttributeValue::GetID(void)
{
	return (LONG) mp_lID;
}

void ExtendedAttributeValue::odl_SetID(LONG newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetID((LONG)newVal);
}

void ExtendedAttributeValue::SetID(LONG newVal)
{
	mp_lID = newVal;
}

BSTR ExtendedAttributeValue::odl_GetValue(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetValue().AllocSysString();
}

CString ExtendedAttributeValue::GetValue(void)
{
	return (CString) mp_sValue;
}

void ExtendedAttributeValue::odl_SetValue(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetValue(newVal);
}

void ExtendedAttributeValue::SetValue(CString newVal)
{
	mp_sValue = newVal;
}

BSTR ExtendedAttributeValue::odl_GetDescription(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetDescription().AllocSysString();
}

CString ExtendedAttributeValue::GetDescription(void)
{
	return (CString) mp_sDescription;
}

void ExtendedAttributeValue::odl_SetDescription(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetDescription(newVal);
}

void ExtendedAttributeValue::SetDescription(CString newVal)
{
	mp_sDescription = newVal;
}

BSTR ExtendedAttributeValue::odl_GetKey(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetKey().AllocSysString();
}

void ExtendedAttributeValue::odl_SetKey(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	USES_CONVERSION;
	SetKey(F_TOUCSTR(T2A(newVal)));
}

CString ExtendedAttributeValue::GetKey(void)
{
	return mp_sKey;
}

void ExtendedAttributeValue::SetKey(CString newVal)
{
	mp_oCollection->mp_SetKey(&mp_sKey, newVal, MP_SET_KEY);
}

BSTR ExtendedAttributeValue::odl_GetXML(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetXML().AllocSysString();
}

void ExtendedAttributeValue::odl_SetXML(LPCTSTR sXML)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetXML(sXML);
}

BOOL ExtendedAttributeValue::IsNull(void)
{
	BOOL bReturn = TRUE;
	if (mp_lID != 0)
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

CString ExtendedAttributeValue::GetXML(void)
{
	if (IsNull() == TRUE)
	{
		return "<Value/>";
	}
	clsXML oXML("Value");
	oXML.InitializeWriter();
	oXML.SetSupportOptional(TRUE);
	oXML.SetBoolsAreNumeric(TRUE);
	oXML.WriteProperty("ID", mp_lID);
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

void ExtendedAttributeValue::SetXML(CString sXML)
{
	clsXML oXML("Value");
	oXML.SetSupportOptional(TRUE);
	oXML.SetXML(sXML);
	oXML.InitializeReader();
	oXML.ReadProperty("ID", mp_lID);
	oXML.ReadProperty("Value", mp_sValue);
	oXML.ReadProperty("Description", mp_sDescription);
}