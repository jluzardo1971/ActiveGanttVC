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

//{A20781A8-191A-4858-87F5-DC46CF95ABC4}
static const IID IID_IResourceExtendedAttribute = { 0xA20781A8, 0x191A, 0x4858, { 0x87, 0xF5, 0xDC, 0x46, 0xCF, 0x95, 0xAB, 0xC4} };

//{129AD19B-70FC-44C9-8EBA-C8FFA60FF8F7}
IMPLEMENT_OLECREATE_FLAGS(ResourceExtendedAttribute, "MSP2010.ResourceExtendedAttribute", afxRegApartmentThreading, 0x129AD19B, 0x70FC, 0x44C9, 0x8E, 0xBA, 0xC8, 0xFF, 0xA6, 0x0F, 0xF8, 0xF7)

BEGIN_DISPATCH_MAP(ResourceExtendedAttribute, CCmdTarget)
DISP_PROPERTY_EX_ID(ResourceExtendedAttribute, "sFieldID", 1, odl_GetFieldID, odl_SetFieldID, VT_BSTR)
DISP_PROPERTY_EX_ID(ResourceExtendedAttribute, "sValue", 2, odl_GetValue, odl_SetValue, VT_BSTR)
DISP_PROPERTY_EX_ID(ResourceExtendedAttribute, "lValueGUID", 3, odl_GetValueGUID, odl_SetValueGUID, VT_I4)
DISP_PROPERTY_EX_ID(ResourceExtendedAttribute, "yDurationFormat", 4, odl_GetDurationFormat, odl_SetDurationFormat, VT_I4)
DISP_PROPERTY_EX_ID(ResourceExtendedAttribute, "Key", 5, odl_GetKey, odl_SetKey, VT_BSTR)
DISP_FUNCTION_ID(ResourceExtendedAttribute, "GetXML", 6, odl_GetXML, VT_BSTR, VTS_NONE)
DISP_FUNCTION_ID(ResourceExtendedAttribute, "SetXML", 7, odl_SetXML, VT_EMPTY, VTS_BSTR)
DISP_FUNCTION_ID(ResourceExtendedAttribute, "IsNull", 8, IsNull, VT_BOOL, VTS_NONE)
DISP_FUNCTION_ID(ResourceExtendedAttribute, "Initialize", 9, Initialize, VT_EMPTY, VTS_NONE)
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
	mp_sFieldID = "";
	mp_sValue = "";
	mp_lValueGUID = 0;
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

LONG ResourceExtendedAttribute::odl_GetValueGUID(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetValueGUID();
}

LONG ResourceExtendedAttribute::GetValueGUID(void)
{
	return (LONG) mp_lValueGUID;
}

void ResourceExtendedAttribute::odl_SetValueGUID(LONG newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetValueGUID((LONG)newVal);
}

void ResourceExtendedAttribute::SetValueGUID(LONG newVal)
{
	mp_lValueGUID = newVal;
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
	if (mp_sFieldID != "")
	{
		bReturn = FALSE;
	}
	if (mp_sValue != "")
	{
		bReturn = FALSE;
	}
	if (mp_lValueGUID != 0)
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
	if (mp_sFieldID != "")
	{
		oXML.WriteProperty("FieldID", mp_sFieldID);
	}
	if (mp_sValue != "")
	{
		oXML.WriteProperty("Value", mp_sValue);
	}
	oXML.WriteProperty("ValueGUID", mp_lValueGUID);
	oXML.WriteProperty("DurationFormat", mp_yDurationFormat);
	return oXML.GetXML();
}

void ResourceExtendedAttribute::SetXML(CString sXML)
{
	clsXML oXML("ExtendedAttribute");
	oXML.SetSupportOptional(TRUE);
	oXML.SetXML(sXML);
	oXML.InitializeReader();
	oXML.ReadProperty("FieldID", mp_sFieldID);
	oXML.ReadProperty("Value", mp_sValue);
	oXML.ReadProperty("ValueGUID", mp_lValueGUID);
	oXML.ReadProperty("DurationFormat", mp_yDurationFormat);
}