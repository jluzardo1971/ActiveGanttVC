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
#include "TaskOutlineCode.h"

IMPLEMENT_DYNCREATE(TaskOutlineCode, CCmdTarget)

//{475C3036-316E-47B9-AA71-3FA45FB1A9DB}
static const IID IID_ITaskOutlineCode = { 0x475C3036, 0x316E, 0x47B9, { 0xAA, 0x71, 0x3F, 0xA4, 0x5F, 0xB1, 0xA9, 0xDB} };

//{0F4DD562-6E32-4E4C-A716-CA05E3539591}
IMPLEMENT_OLECREATE_FLAGS(TaskOutlineCode, "MSP2003.TaskOutlineCode", afxRegApartmentThreading, 0x0F4DD562, 0x6E32, 0x4E4C, 0xA7, 0x16, 0xCA, 0x05, 0xE3, 0x53, 0x95, 0x91)

BEGIN_DISPATCH_MAP(TaskOutlineCode, CCmdTarget)
DISP_PROPERTY_EX_ID(TaskOutlineCode, "lUID", 1, odl_GetUID, odl_SetUID, VT_I4)
DISP_PROPERTY_EX_ID(TaskOutlineCode, "sFieldID", 2, odl_GetFieldID, odl_SetFieldID, VT_BSTR)
DISP_PROPERTY_EX_ID(TaskOutlineCode, "lValueID", 3, odl_GetValueID, odl_SetValueID, VT_I4)
DISP_PROPERTY_EX_ID(TaskOutlineCode, "Key", 4, odl_GetKey, odl_SetKey, VT_BSTR)
DISP_FUNCTION_ID(TaskOutlineCode, "GetXML", 5, odl_GetXML, VT_BSTR, VTS_NONE)
DISP_FUNCTION_ID(TaskOutlineCode, "SetXML", 6, odl_SetXML, VT_EMPTY, VTS_BSTR)
DISP_FUNCTION_ID(TaskOutlineCode, "IsNull", 7, IsNull, VT_BOOL, VTS_NONE)
DISP_FUNCTION_ID(TaskOutlineCode, "Initialize", 8, Initialize, VT_EMPTY, VTS_NONE)
END_DISPATCH_MAP()

BEGIN_INTERFACE_MAP(TaskOutlineCode, CCmdTarget)
	INTERFACE_PART(TaskOutlineCode, IID_ITaskOutlineCode, Dispatch)
END_INTERFACE_MAP()

BEGIN_MESSAGE_MAP(TaskOutlineCode, CCmdTarget)
END_MESSAGE_MAP()


void TaskOutlineCode::Initialize(void)
{
	InitVars();
}

TaskOutlineCode::TaskOutlineCode()
{
	EnableAutomation();
	AfxOleLockApp();
	InitVars();
}

void TaskOutlineCode::InitVars(void)
{
	mp_lUID = 0;
	mp_sFieldID = "";
	mp_lValueID = 0;
}

TaskOutlineCode::~TaskOutlineCode()
{
	AfxOleUnlockApp();
}

void TaskOutlineCode::OnFinalRelease()
{
	clsItemBase::OnFinalRelease();
}

LONG TaskOutlineCode::odl_GetUID(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetUID();
}

LONG TaskOutlineCode::GetUID(void)
{
	return (LONG) mp_lUID;
}

void TaskOutlineCode::odl_SetUID(LONG newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetUID((LONG)newVal);
}

void TaskOutlineCode::SetUID(LONG newVal)
{
	mp_lUID = newVal;
}

BSTR TaskOutlineCode::odl_GetFieldID(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetFieldID().AllocSysString();
}

CString TaskOutlineCode::GetFieldID(void)
{
	return (CString) mp_sFieldID;
}

void TaskOutlineCode::odl_SetFieldID(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetFieldID(newVal);
}

void TaskOutlineCode::SetFieldID(CString newVal)
{
	mp_sFieldID = newVal;
}

LONG TaskOutlineCode::odl_GetValueID(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetValueID();
}

LONG TaskOutlineCode::GetValueID(void)
{
	return (LONG) mp_lValueID;
}

void TaskOutlineCode::odl_SetValueID(LONG newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetValueID((LONG)newVal);
}

void TaskOutlineCode::SetValueID(LONG newVal)
{
	mp_lValueID = newVal;
}

BSTR TaskOutlineCode::odl_GetKey(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetKey().AllocSysString();
}

void TaskOutlineCode::odl_SetKey(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	USES_CONVERSION;
	SetKey(F_TOUCSTR(T2A(newVal)));
}

CString TaskOutlineCode::GetKey(void)
{
	return mp_sKey;
}

void TaskOutlineCode::SetKey(CString newVal)
{
	mp_oCollection->mp_SetKey(&mp_sKey, newVal, MP_SET_KEY);
}

BSTR TaskOutlineCode::odl_GetXML(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetXML().AllocSysString();
}

void TaskOutlineCode::odl_SetXML(LPCTSTR sXML)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetXML(sXML);
}

BOOL TaskOutlineCode::IsNull(void)
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

CString TaskOutlineCode::GetXML(void)
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

void TaskOutlineCode::SetXML(CString sXML)
{
	clsXML oXML("OutlineCode");
	oXML.SetSupportOptional(TRUE);
	oXML.SetXML(sXML);
	oXML.InitializeReader();
	oXML.ReadProperty("UID", mp_lUID);
	oXML.ReadProperty("FieldID", mp_sFieldID);
	oXML.ReadProperty("ValueID", mp_lValueID);
}