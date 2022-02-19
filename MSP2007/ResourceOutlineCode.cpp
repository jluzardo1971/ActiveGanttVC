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

//{BA6AB210-C861-4CDB-B215-E573A9F2121A}
static const IID IID_IResourceOutlineCode = { 0xBA6AB210, 0xC861, 0x4CDB, { 0xB2, 0x15, 0xE5, 0x73, 0xA9, 0xF2, 0x12, 0x1A} };

//{DDE96251-3360-46D1-BC5F-D7EBC919837E}
IMPLEMENT_OLECREATE_FLAGS(ResourceOutlineCode, "MSP2007.ResourceOutlineCode", afxRegApartmentThreading, 0xDDE96251, 0x3360, 0x46D1, 0xBC, 0x5F, 0xD7, 0xEB, 0xC9, 0x19, 0x83, 0x7E)

BEGIN_DISPATCH_MAP(ResourceOutlineCode, CCmdTarget)
DISP_PROPERTY_EX_ID(ResourceOutlineCode, "sFieldID", 1, odl_GetFieldID, odl_SetFieldID, VT_BSTR)
DISP_PROPERTY_EX_ID(ResourceOutlineCode, "lValueID", 2, odl_GetValueID, odl_SetValueID, VT_I4)
DISP_PROPERTY_EX_ID(ResourceOutlineCode, "lValueGUID", 3, odl_GetValueGUID, odl_SetValueGUID, VT_I4)
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
	mp_sFieldID = "";
	mp_lValueID = 0;
	mp_lValueGUID = 0;
}

ResourceOutlineCode::~ResourceOutlineCode()
{
	AfxOleUnlockApp();
}

void ResourceOutlineCode::OnFinalRelease()
{
	clsItemBase::OnFinalRelease();
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

LONG ResourceOutlineCode::odl_GetValueGUID(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetValueGUID();
}

LONG ResourceOutlineCode::GetValueGUID(void)
{
	return (LONG) mp_lValueGUID;
}

void ResourceOutlineCode::odl_SetValueGUID(LONG newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetValueGUID((LONG)newVal);
}

void ResourceOutlineCode::SetValueGUID(LONG newVal)
{
	mp_lValueGUID = newVal;
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
	if (mp_sFieldID != "")
	{
		bReturn = FALSE;
	}
	if (mp_lValueID != 0)
	{
		bReturn = FALSE;
	}
	if (mp_lValueGUID != 0)
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
	if (mp_sFieldID != "")
	{
		oXML.WriteProperty("FieldID", mp_sFieldID);
	}
	oXML.WriteProperty("ValueID", mp_lValueID);
	oXML.WriteProperty("ValueGUID", mp_lValueGUID);
	return oXML.GetXML();
}

void ResourceOutlineCode::SetXML(CString sXML)
{
	clsXML oXML("OutlineCode");
	oXML.SetSupportOptional(TRUE);
	oXML.SetXML(sXML);
	oXML.InitializeReader();
	oXML.ReadProperty("FieldID", mp_sFieldID);
	oXML.ReadProperty("ValueID", mp_lValueID);
	oXML.ReadProperty("ValueGUID", mp_lValueGUID);
}