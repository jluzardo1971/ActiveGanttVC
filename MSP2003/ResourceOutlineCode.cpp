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
#include "ResourceOutlineCode.h"

IMPLEMENT_DYNCREATE(ResourceOutlineCode, CCmdTarget)

//{A42B0139-082D-4154-8184-1FE3A8607FF3}
static const IID IID_IResourceOutlineCode = { 0xA42B0139, 0x082D, 0x4154, { 0x81, 0x84, 0x1F, 0xE3, 0xA8, 0x60, 0x7F, 0xF3} };

//{7AFC2E4C-6972-4129-A82B-94AFEBA6EE3C}
IMPLEMENT_OLECREATE_FLAGS(ResourceOutlineCode, "MSP2003.ResourceOutlineCode", afxRegApartmentThreading, 0x7AFC2E4C, 0x6972, 0x4129, 0xA8, 0x2B, 0x94, 0xAF, 0xEB, 0xA6, 0xEE, 0x3C)

BEGIN_DISPATCH_MAP(ResourceOutlineCode, CCmdTarget)
DISP_PROPERTY_EX_ID(ResourceOutlineCode, "lUID", 1, odl_GetUID, odl_SetUID, VT_I4)
DISP_PROPERTY_EX_ID(ResourceOutlineCode, "sFieldID", 2, odl_GetFieldID, odl_SetFieldID, VT_BSTR)
DISP_PROPERTY_EX_ID(ResourceOutlineCode, "lValueID", 3, odl_GetValueID, odl_SetValueID, VT_I4)
DISP_PROPERTY_EX_ID(ResourceOutlineCode, "Key", 4, odl_GetKey, odl_SetKey, VT_BSTR)
DISP_FUNCTION_ID(ResourceOutlineCode, "GetXML", 5, odl_GetXML, VT_BSTR, VTS_NONE)
DISP_FUNCTION_ID(ResourceOutlineCode, "SetXML", 6, odl_SetXML, VT_EMPTY, VTS_BSTR)
DISP_FUNCTION_ID(ResourceOutlineCode, "IsNull", 7, IsNull, VT_BOOL, VTS_NONE)
DISP_FUNCTION_ID(ResourceOutlineCode, "Initialize", 8, Initialize, VT_EMPTY, VTS_NONE)
END_DISPATCH_MAP()

BEGIN_INTERFACE_MAP(ResourceOutlineCode, CCmdTarget)
	INTERFACE_PART(ResourceOutlineCode, IID_IResourceOutlineCode, Dispatch)
END_INTERFACE_MAP()

BEGIN_MESSAGE_MAP(ResourceOutlineCode, CCmdTarget)
END_MESSAGE_MAP()


void ResourceOutlineCode::Initialize(void)
{
	InitVars();
}

ResourceOutlineCode::ResourceOutlineCode()
{
	EnableAutomation();
	AfxOleLockApp();
	InitVars();
}

void ResourceOutlineCode::InitVars(void)
{
	mp_lUID = 0;
	mp_sFieldID = "";
	mp_lValueID = 0;
}

ResourceOutlineCode::~ResourceOutlineCode()
{
	AfxOleUnlockApp();
}

void ResourceOutlineCode::OnFinalRelease()
{
	clsItemBase::OnFinalRelease();
}

LONG ResourceOutlineCode::odl_GetUID(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetUID();
}

LONG ResourceOutlineCode::GetUID(void)
{
	return (LONG) mp_lUID;
}

void ResourceOutlineCode::odl_SetUID(LONG newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetUID((LONG)newVal);
}

void ResourceOutlineCode::SetUID(LONG newVal)
{
	mp_lUID = newVal;
}

BSTR ResourceOutlineCode::odl_GetFieldID(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetFieldID().AllocSysString();
}

CString ResourceOutlineCode::GetFieldID(void)
{
	return (CString) mp_sFieldID;
}

void ResourceOutlineCode::odl_SetFieldID(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetFieldID(newVal);
}

void ResourceOutlineCode::SetFieldID(CString newVal)
{
	mp_sFieldID = newVal;
}

LONG ResourceOutlineCode::odl_GetValueID(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetValueID();
}

LONG ResourceOutlineCode::GetValueID(void)
{
	return (LONG) mp_lValueID;
}

void ResourceOutlineCode::odl_SetValueID(LONG newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetValueID((LONG)newVal);
}

void ResourceOutlineCode::SetValueID(LONG newVal)
{
	mp_lValueID = newVal;
}

BSTR ResourceOutlineCode::odl_GetKey(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetKey().AllocSysString();
}

void ResourceOutlineCode::odl_SetKey(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	USES_CONVERSION;
	SetKey(F_TOUCSTR(T2A(newVal)));
}

CString ResourceOutlineCode::GetKey(void)
{
	return mp_sKey;
}

void ResourceOutlineCode::SetKey(CString newVal)
{
	mp_oCollection->mp_SetKey(&mp_sKey, newVal, MP_SET_KEY);
}

BSTR ResourceOutlineCode::odl_GetXML(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetXML().AllocSysString();
}

void ResourceOutlineCode::odl_SetXML(LPCTSTR sXML)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetXML(sXML);
}

BOOL ResourceOutlineCode::IsNull(void)
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
	if (mp_lValueID != 0)
	{
		bReturn = FALSE;
	}
	return bReturn;
}

CString ResourceOutlineCode::GetXML(void)
{
	if (IsNull() == TRUE)
	{
		return "<OutlineCode/>";
	}
	clsXML oXML("OutlineCode");
	oXML.InitializeWriter();
	oXML.SetSupportOptional(TRUE);
	oXML.SetBoolsAreNumeric(TRUE);
	oXML.WriteProperty("UID", mp_lUID);
	if (mp_sFieldID != "")
	{
		oXML.WriteProperty("FieldID", mp_sFieldID);
	}
	oXML.WriteProperty("ValueID", mp_lValueID);
	return oXML.GetXML();
}

void ResourceOutlineCode::SetXML(CString sXML)
{
	clsXML oXML("OutlineCode");
	oXML.SetSupportOptional(TRUE);
	oXML.SetXML(sXML);
	oXML.InitializeReader();
	oXML.ReadProperty("UID", mp_lUID);
	oXML.ReadProperty("FieldID", mp_sFieldID);
	oXML.ReadProperty("ValueID", mp_lValueID);
}