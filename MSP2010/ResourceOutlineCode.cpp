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

//{96D27B77-DCED-4D3A-AA4E-42B9E204A2A4}
static const IID IID_IResourceOutlineCode = { 0x96D27B77, 0xDCED, 0x4D3A, { 0xAA, 0x4E, 0x42, 0xB9, 0xE2, 0x04, 0xA2, 0xA4} };

//{B1EA143B-94CC-4367-AAF0-7D20F1A06FC2}
IMPLEMENT_OLECREATE_FLAGS(ResourceOutlineCode, "MSP2010.ResourceOutlineCode", afxRegApartmentThreading, 0xB1EA143B, 0x94CC, 0x4367, 0xAA, 0xF0, 0x7D, 0x20, 0xF1, 0xA0, 0x6F, 0xC2)

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