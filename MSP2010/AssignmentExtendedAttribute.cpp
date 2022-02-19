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
#include "AssignmentExtendedAttribute.h"

IMPLEMENT_DYNCREATE(AssignmentExtendedAttribute, CCmdTarget)

//{67B0933A-23C2-4E65-A144-5C7BE8A3F179}
static const IID IID_IAssignmentExtendedAttribute = { 0x67B0933A, 0x23C2, 0x4E65, { 0xA1, 0x44, 0x5C, 0x7B, 0xE8, 0xA3, 0xF1, 0x79} };

//{38C6E149-F2FA-423C-AA6B-425A2FFDD2E9}
IMPLEMENT_OLECREATE_FLAGS(AssignmentExtendedAttribute, "MSP2010.AssignmentExtendedAttribute", afxRegApartmentThreading, 0x38C6E149, 0xF2FA, 0x423C, 0xAA, 0x6B, 0x42, 0x5A, 0x2F, 0xFD, 0xD2, 0xE9)

BEGIN_DISPATCH_MAP(AssignmentExtendedAttribute, CCmdTarget)
DISP_PROPERTY_EX_ID(AssignmentExtendedAttribute, "sFieldID", 1, odl_GetFieldID, odl_SetFieldID, VT_BSTR)
DISP_PROPERTY_EX_ID(AssignmentExtendedAttribute, "sValue", 2, odl_GetValue, odl_SetValue, VT_BSTR)
DISP_PROPERTY_EX_ID(AssignmentExtendedAttribute, "lValueGUID", 3, odl_GetValueGUID, odl_SetValueGUID, VT_I4)
DISP_PROPERTY_EX_ID(AssignmentExtendedAttribute, "yDurationFormat", 4, odl_GetDurationFormat, odl_SetDurationFormat, VT_I4)
DISP_PROPERTY_EX_ID(AssignmentExtendedAttribute, "Key", 5, odl_GetKey, odl_SetKey, VT_BSTR)
DISP_FUNCTION_ID(AssignmentExtendedAttribute, "GetXML", 6, odl_GetXML, VT_BSTR, VTS_NONE)
DISP_FUNCTION_ID(AssignmentExtendedAttribute, "SetXML", 7, odl_SetXML, VT_EMPTY, VTS_BSTR)
DISP_FUNCTION_ID(AssignmentExtendedAttribute, "IsNull", 8, IsNull, VT_BOOL, VTS_NONE)
DISP_FUNCTION_ID(AssignmentExtendedAttribute, "Initialize", 9, Initialize, VT_EMPTY, VTS_NONE)
END_DISPATCH_MAP()

BEGIN_INTERFACE_MAP(AssignmentExtendedAttribute, CCmdTarget)
	INTERFACE_PART(AssignmentExtendedAttribute, IID_IAssignmentExtendedAttribute, Dispatch)
END_INTERFACE_MAP()

BEGIN_MESSAGE_MAP(AssignmentExtendedAttribute, CCmdTarget)
END_MESSAGE_MAP()


void AssignmentExtendedAttribute::Initialize(void)
{
	InitVars();
}

AssignmentExtendedAttribute::AssignmentExtendedAttribute()
{
	EnableAutomation();
	AfxOleLockApp();
	InitVars();
}

void AssignmentExtendedAttribute::InitVars(void)
{
	mp_sFieldID = "";
	mp_sValue = "";
	mp_lValueGUID = 0;
	mp_yDurationFormat = DF_M;
}

AssignmentExtendedAttribute::~AssignmentExtendedAttribute()
{
	AfxOleUnlockApp();
}

void AssignmentExtendedAttribute::OnFinalRelease()
{
	clsItemBase::OnFinalRelease();
}

BSTR AssignmentExtendedAttribute::odl_GetFieldID(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetFieldID().AllocSysString();
}

CString AssignmentExtendedAttribute::GetFieldID(void)
{
	return (CString) mp_sFieldID;
}

void AssignmentExtendedAttribute::odl_SetFieldID(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetFieldID(newVal);
}

void AssignmentExtendedAttribute::SetFieldID(CString newVal)
{
	mp_sFieldID = newVal;
}

BSTR AssignmentExtendedAttribute::odl_GetValue(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetValue().AllocSysString();
}

CString AssignmentExtendedAttribute::GetValue(void)
{
	return (CString) mp_sValue;
}

void AssignmentExtendedAttribute::odl_SetValue(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetValue(newVal);
}

void AssignmentExtendedAttribute::SetValue(CString newVal)
{
	mp_sValue = newVal;
}

LONG AssignmentExtendedAttribute::odl_GetValueGUID(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetValueGUID();
}

LONG AssignmentExtendedAttribute::GetValueGUID(void)
{
	return (LONG) mp_lValueGUID;
}

void AssignmentExtendedAttribute::odl_SetValueGUID(LONG newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetValueGUID((LONG)newVal);
}

void AssignmentExtendedAttribute::SetValueGUID(LONG newVal)
{
	mp_lValueGUID = newVal;
}

LONG AssignmentExtendedAttribute::odl_GetDurationFormat(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetDurationFormat();
}

E_DURATIONFORMAT AssignmentExtendedAttribute::GetDurationFormat(void)
{
	return (E_DURATIONFORMAT) mp_yDurationFormat;
}

void AssignmentExtendedAttribute::odl_SetDurationFormat(LONG newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetDurationFormat((E_DURATIONFORMAT)newVal);
}

void AssignmentExtendedAttribute::SetDurationFormat(E_DURATIONFORMAT newVal)
{
	mp_yDurationFormat = newVal;
}

BSTR AssignmentExtendedAttribute::odl_GetKey(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetKey().AllocSysString();
}

void AssignmentExtendedAttribute::odl_SetKey(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	USES_CONVERSION;
	SetKey(F_TOUCSTR(T2A(newVal)));
}

CString AssignmentExtendedAttribute::GetKey(void)
{
	return mp_sKey;
}

void AssignmentExtendedAttribute::SetKey(CString newVal)
{
	mp_oCollection->mp_SetKey(&mp_sKey, newVal, MP_SET_KEY);
}

BSTR AssignmentExtendedAttribute::odl_GetXML(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetXML().AllocSysString();
}

void AssignmentExtendedAttribute::odl_SetXML(LPCTSTR sXML)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetXML(sXML);
}

BOOL AssignmentExtendedAttribute::IsNull(void)
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

CString AssignmentExtendedAttribute::GetXML(void)
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

void AssignmentExtendedAttribute::SetXML(CString sXML)
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