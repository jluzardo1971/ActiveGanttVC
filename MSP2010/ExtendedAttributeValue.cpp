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

//{1BD09511-294F-4C1F-B7C6-24BA59B44E83}
static const IID IID_IExtendedAttributeValue = { 0x1BD09511, 0x294F, 0x4C1F, { 0xB7, 0xC6, 0x24, 0xBA, 0x59, 0xB4, 0x4E, 0x83} };

//{8EDF1318-985C-4AC8-BC60-3C35769CA434}
IMPLEMENT_OLECREATE_FLAGS(ExtendedAttributeValue, "MSP2010.ExtendedAttributeValue", afxRegApartmentThreading, 0x8EDF1318, 0x985C, 0x4AC8, 0xBC, 0x60, 0x3C, 0x35, 0x76, 0x9C, 0xA4, 0x34)

BEGIN_DISPATCH_MAP(ExtendedAttributeValue, CCmdTarget)
DISP_PROPERTY_EX_ID(ExtendedAttributeValue, "lID", 1, odl_GetID, odl_SetID, VT_I4)
DISP_PROPERTY_EX_ID(ExtendedAttributeValue, "sValue", 2, odl_GetValue, odl_SetValue, VT_BSTR)
DISP_PROPERTY_EX_ID(ExtendedAttributeValue, "sDescription", 3, odl_GetDescription, odl_SetDescription, VT_BSTR)
DISP_PROPERTY_EX_ID(ExtendedAttributeValue, "sPhonetic", 4, odl_GetPhonetic, odl_SetPhonetic, VT_BSTR)
DISP_PROPERTY_EX_ID(ExtendedAttributeValue, "Key", 5, odl_GetKey, odl_SetKey, VT_BSTR)
DISP_FUNCTION_ID(ExtendedAttributeValue, "GetXML", 6, odl_GetXML, VT_BSTR, VTS_NONE)
DISP_FUNCTION_ID(ExtendedAttributeValue, "SetXML", 7, odl_SetXML, VT_EMPTY, VTS_BSTR)
DISP_FUNCTION_ID(ExtendedAttributeValue, "IsNull", 8, IsNull, VT_BOOL, VTS_NONE)
DISP_FUNCTION_ID(ExtendedAttributeValue, "Initialize", 9, Initialize, VT_EMPTY, VTS_NONE)
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
	mp_sPhonetic = "";
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

BSTR ExtendedAttributeValue::odl_GetPhonetic(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetPhonetic().AllocSysString();
}

CString ExtendedAttributeValue::GetPhonetic(void)
{
	return (CString) mp_sPhonetic;
}

void ExtendedAttributeValue::odl_SetPhonetic(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetPhonetic(newVal);
}

void ExtendedAttributeValue::SetPhonetic(CString newVal)
{
	mp_sPhonetic = newVal;
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
	if (mp_sPhonetic != "")
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
	if (mp_sPhonetic != "")
	{
		oXML.WriteProperty("Phonetic", mp_sPhonetic);
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
	oXML.ReadProperty("Phonetic", mp_sPhonetic);
}