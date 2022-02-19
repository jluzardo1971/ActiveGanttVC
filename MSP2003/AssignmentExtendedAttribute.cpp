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

//{61ADF839-29BE-4FB3-8359-C102C74506D2}
static const IID IID_IAssignmentExtendedAttribute = { 0x61ADF839, 0x29BE, 0x4FB3, { 0x83, 0x59, 0xC1, 0x02, 0xC7, 0x45, 0x06, 0xD2} };

//{D54DC6D0-EBBA-4760-8A25-D817103F8196}
IMPLEMENT_OLECREATE_FLAGS(AssignmentExtendedAttribute, "MSP2003.AssignmentExtendedAttribute", afxRegApartmentThreading, 0xD54DC6D0, 0xEBBA, 0x4760, 0x8A, 0x25, 0xD8, 0x17, 0x10, 0x3F, 0x81, 0x96)

BEGIN_DISPATCH_MAP(AssignmentExtendedAttribute, CCmdTarget)
DISP_PROPERTY_EX_ID(AssignmentExtendedAttribute, "lUID", 1, odl_GetUID, odl_SetUID, VT_I4)
DISP_PROPERTY_EX_ID(AssignmentExtendedAttribute, "sFieldID", 2, odl_GetFieldID, odl_SetFieldID, VT_BSTR)
DISP_PROPERTY_EX_ID(AssignmentExtendedAttribute, "sValue", 3, odl_GetValue, odl_SetValue, VT_BSTR)
DISP_PROPERTY_EX_ID(AssignmentExtendedAttribute, "lValueID", 4, odl_GetValueID, odl_SetValueID, VT_I4)
DISP_PROPERTY_EX_ID(AssignmentExtendedAttribute, "yDurationFormat", 5, odl_GetDurationFormat, odl_SetDurationFormat, VT_I4)
DISP_PROPERTY_EX_ID(AssignmentExtendedAttribute, "Key", 6, odl_GetKey, odl_SetKey, VT_BSTR)
DISP_FUNCTION_ID(AssignmentExtendedAttribute, "GetXML", 7, odl_GetXML, VT_BSTR, VTS_NONE)
DISP_FUNCTION_ID(AssignmentExtendedAttribute, "SetXML", 8, odl_SetXML, VT_EMPTY, VTS_BSTR)
DISP_FUNCTION_ID(AssignmentExtendedAttribute, "IsNull", 9, IsNull, VT_BOOL, VTS_NONE)
DISP_FUNCTION_ID(AssignmentExtendedAttribute, "Initialize", 10, Initialize, VT_EMPTY, VTS_NONE)
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
	mp_lUID = 0;
	mp_sFieldID = "";
	mp_sValue = "";
	mp_lValueID = 0;
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

LONG AssignmentExtendedAttribute::odl_GetUID(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetUID();
}

LONG AssignmentExtendedAttribute::GetUID(void)
{
	return (LONG) mp_lUID;
}

void AssignmentExtendedAttribute::odl_SetUID(LONG newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetUID((LONG)newVal);
}

void AssignmentExtendedAttribute::SetUID(LONG newVal)
{
	mp_lUID = newVal;
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

LONG AssignmentExtendedAttribute::odl_GetValueID(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetValueID();
}

LONG AssignmentExtendedAttribute::GetValueID(void)
{
	return (LONG) mp_lValueID;
}

void AssignmentExtendedAttribute::odl_SetValueID(LONG newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetValueID((LONG)newVal);
}

void AssignmentExtendedAttribute::SetValueID(LONG newVal)
{
	mp_lValueID = newVal;
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

void AssignmentExtendedAttribute::SetXML(CString sXML)
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