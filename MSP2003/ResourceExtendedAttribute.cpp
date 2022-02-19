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
#include "ResourceExtendedAttribute.h"

IMPLEMENT_DYNCREATE(ResourceExtendedAttribute, CCmdTarget)

//{8AE27DC1-0878-4637-8BCD-FF29F82C36A1}
static const IID IID_IResourceExtendedAttribute = { 0x8AE27DC1, 0x0878, 0x4637, { 0x8B, 0xCD, 0xFF, 0x29, 0xF8, 0x2C, 0x36, 0xA1} };

//{D82A4D8F-3490-4CD1-A27A-FC592111E5D1}
IMPLEMENT_OLECREATE_FLAGS(ResourceExtendedAttribute, "MSP2003.ResourceExtendedAttribute", afxRegApartmentThreading, 0xD82A4D8F, 0x3490, 0x4CD1, 0xA2, 0x7A, 0xFC, 0x59, 0x21, 0x11, 0xE5, 0xD1)

BEGIN_DISPATCH_MAP(ResourceExtendedAttribute, CCmdTarget)
DISP_PROPERTY_EX_ID(ResourceExtendedAttribute, "lUID", 1, odl_GetUID, odl_SetUID, VT_I4)
DISP_PROPERTY_EX_ID(ResourceExtendedAttribute, "sFieldID", 2, odl_GetFieldID, odl_SetFieldID, VT_BSTR)
DISP_PROPERTY_EX_ID(ResourceExtendedAttribute, "sValue", 3, odl_GetValue, odl_SetValue, VT_BSTR)
DISP_PROPERTY_EX_ID(ResourceExtendedAttribute, "lValueID", 4, odl_GetValueID, odl_SetValueID, VT_I4)
DISP_PROPERTY_EX_ID(ResourceExtendedAttribute, "yDurationFormat", 5, odl_GetDurationFormat, odl_SetDurationFormat, VT_I4)
DISP_PROPERTY_EX_ID(ResourceExtendedAttribute, "Key", 6, odl_GetKey, odl_SetKey, VT_BSTR)
DISP_FUNCTION_ID(ResourceExtendedAttribute, "GetXML", 7, odl_GetXML, VT_BSTR, VTS_NONE)
DISP_FUNCTION_ID(ResourceExtendedAttribute, "SetXML", 8, odl_SetXML, VT_EMPTY, VTS_BSTR)
DISP_FUNCTION_ID(ResourceExtendedAttribute, "IsNull", 9, IsNull, VT_BOOL, VTS_NONE)
DISP_FUNCTION_ID(ResourceExtendedAttribute, "Initialize", 10, Initialize, VT_EMPTY, VTS_NONE)
END_DISPATCH_MAP()

BEGIN_INTERFACE_MAP(ResourceExtendedAttribute, CCmdTarget)
	INTERFACE_PART(ResourceExtendedAttribute, IID_IResourceExtendedAttribute, Dispatch)
END_INTERFACE_MAP()

BEGIN_MESSAGE_MAP(ResourceExtendedAttribute, CCmdTarget)
END_MESSAGE_MAP()


void ResourceExtendedAttribute::Initialize(void)
{
	InitVars();
}

ResourceExtendedAttribute::ResourceExtendedAttribute()
{
	EnableAutomation();
	AfxOleLockApp();
	InitVars();
}

void ResourceExtendedAttribute::InitVars(void)
{
	mp_lUID = 0;
	mp_sFieldID = "";
	mp_sValue = "";
	mp_lValueID = 0;
	mp_yDurationFormat = DF_M;
}

ResourceExtendedAttribute::~ResourceExtendedAttribute()
{
	AfxOleUnlockApp();
}

void ResourceExtendedAttribute::OnFinalRelease()
{
	clsItemBase::OnFinalRelease();
}

LONG ResourceExtendedAttribute::odl_GetUID(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetUID();
}

LONG ResourceExtendedAttribute::GetUID(void)
{
	return (LONG) mp_lUID;
}

void ResourceExtendedAttribute::odl_SetUID(LONG newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetUID((LONG)newVal);
}

void ResourceExtendedAttribute::SetUID(LONG newVal)
{
	mp_lUID = newVal;
}

BSTR ResourceExtendedAttribute::odl_GetFieldID(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetFieldID().AllocSysString();
}

CString ResourceExtendedAttribute::GetFieldID(void)
{
	return (CString) mp_sFieldID;
}

void ResourceExtendedAttribute::odl_SetFieldID(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetFieldID(newVal);
}

void ResourceExtendedAttribute::SetFieldID(CString newVal)
{
	mp_sFieldID = newVal;
}

BSTR ResourceExtendedAttribute::odl_GetValue(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetValue().AllocSysString();
}

CString ResourceExtendedAttribute::GetValue(void)
{
	return (CString) mp_sValue;
}

void ResourceExtendedAttribute::odl_SetValue(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetValue(newVal);
}

void ResourceExtendedAttribute::SetValue(CString newVal)
{
	mp_sValue = newVal;
}

LONG ResourceExtendedAttribute::odl_GetValueID(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetValueID();
}

LONG ResourceExtendedAttribute::GetValueID(void)
{
	return (LONG) mp_lValueID;
}

void ResourceExtendedAttribute::odl_SetValueID(LONG newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetValueID((LONG)newVal);
}

void ResourceExtendedAttribute::SetValueID(LONG newVal)
{
	mp_lValueID = newVal;
}

LONG ResourceExtendedAttribute::odl_GetDurationFormat(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetDurationFormat();
}

E_DURATIONFORMAT ResourceExtendedAttribute::GetDurationFormat(void)
{
	return (E_DURATIONFORMAT) mp_yDurationFormat;
}

void ResourceExtendedAttribute::odl_SetDurationFormat(LONG newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetDurationFormat((E_DURATIONFORMAT)newVal);
}

void ResourceExtendedAttribute::SetDurationFormat(E_DURATIONFORMAT newVal)
{
	mp_yDurationFormat = newVal;
}

BSTR ResourceExtendedAttribute::odl_GetKey(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetKey().AllocSysString();
}

void ResourceExtendedAttribute::odl_SetKey(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	USES_CONVERSION;
	SetKey(F_TOUCSTR(T2A(newVal)));
}

CString ResourceExtendedAttribute::GetKey(void)
{
	return mp_sKey;
}

void ResourceExtendedAttribute::SetKey(CString newVal)
{
	mp_oCollection->mp_SetKey(&mp_sKey, newVal, MP_SET_KEY);
}

BSTR ResourceExtendedAttribute::odl_GetXML(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetXML().AllocSysString();
}

void ResourceExtendedAttribute::odl_SetXML(LPCTSTR sXML)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetXML(sXML);
}

BOOL ResourceExtendedAttribute::IsNull(void)
{
	BOOL bReturn = TRUE;
	if (mp_lUID != 0)
	{
		bReturn = FALSE;
	}
	if (mp_sFieldID != "")
	{
		bReturn = FALSE;
	}
	if (mp_sValue != "")
	{
		bReturn = FALSE;
	}
	if (mp_lValueID != 0)
	{
		bReturn = FALSE;
	}
	if (mp_yDurationFormat != DF_M)
	{
		bReturn = FALSE;
	}
	return bReturn;
}

CString ResourceExtendedAttribute::GetXML(void)
{
	if (IsNull() == TRUE)
	{
		return "<ExtendedAttribute/>";
	}
	clsXML oXML("ExtendedAttribute");
	oXML.InitializeWriter();
	oXML.SetSupportOptional(TRUE);
	oXML.SetBoolsAreNumeric(TRUE);
	oXML.WriteProperty("UID", mp_lUID);
	if (mp_sFieldID != "")
	{
		oXML.WriteProperty("FieldID", mp_sFieldID);
	}
	if (mp_sValue != "")
	{
		oXML.WriteProperty("Value", mp_sValue);
	}
	oXML.WriteProperty("ValueID", mp_lValueID);
	oXML.WriteProperty("DurationFormat", mp_yDurationFormat);
	return oXML.GetXML();
}

void ResourceExtendedAttribute::SetXML(CString sXML)
{
	clsXML oXML("ExtendedAttribute");
	oXML.SetSupportOptional(TRUE);
	oXML.SetXML(sXML);
	oXML.InitializeReader();
	oXML.ReadProperty("UID", mp_lUID);
	oXML.ReadProperty("FieldID", mp_sFieldID);
	oXML.ReadProperty("Value", mp_sValue);
	oXML.ReadProperty("ValueID", mp_lValueID);
	oXML.ReadProperty("DurationFormat", mp_yDurationFormat);
}