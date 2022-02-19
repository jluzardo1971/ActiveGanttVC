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

//{C7F8832B-2F39-46FE-9450-3405C50DEE13}
static const IID IID_ITaskOutlineCode = { 0xC7F8832B, 0x2F39, 0x46FE, { 0x94, 0x50, 0x34, 0x05, 0xC5, 0x0D, 0xEE, 0x13} };

//{B4E788B6-687F-4364-834C-71EADB6F27F7}
IMPLEMENT_OLECREATE_FLAGS(TaskOutlineCode, "MSP2007.TaskOutlineCode", afxRegApartmentThreading, 0xB4E788B6, 0x687F, 0x4364, 0x83, 0x4C, 0x71, 0xEA, 0xDB, 0x6F, 0x27, 0xF7)

BEGIN_DISPATCH_MAP(TaskOutlineCode, CCmdTarget)
DISP_PROPERTY_EX_ID(TaskOutlineCode, "sFieldID", 1, odl_GetFieldID, odl_SetFieldID, VT_BSTR)
DISP_PROPERTY_EX_ID(TaskOutlineCode, "lValueID", 2, odl_GetValueID, odl_SetValueID, VT_I4)
DISP_PROPERTY_EX_ID(TaskOutlineCode, "lValueGUID", 3, odl_GetValueGUID, odl_SetValueGUID, VT_I4)
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
	mp_sFieldID = "";
	mp_lValueID = 0;
	mp_lValueGUID = 0;
}

TaskOutlineCode::~TaskOutlineCode()
{
	AfxOleUnlockApp();
}

void TaskOutlineCode::OnFinalRelease()
{
	clsItemBase::OnFinalRelease();
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

LONG TaskOutlineCode::odl_GetValueGUID(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetValueGUID();
}

LONG TaskOutlineCode::GetValueGUID(void)
{
	return (LONG) mp_lValueGUID;
}

void TaskOutlineCode::odl_SetValueGUID(LONG newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetValueGUID((LONG)newVal);
}

void TaskOutlineCode::SetValueGUID(LONG newVal)
{
	mp_lValueGUID = newVal;
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
	if (mp_sFieldID != "")
	{
		oXML.WriteProperty("FieldID", mp_sFieldID);
	}
	oXML.WriteProperty("ValueID", mp_lValueID);
	oXML.WriteProperty("ValueGUID", mp_lValueGUID);
	return oXML.GetXML();
}

void TaskOutlineCode::SetXML(CString sXML)
{
	clsXML oXML("OutlineCode");
	oXML.SetSupportOptional(TRUE);
	oXML.SetXML(sXML);
	oXML.InitializeReader();
	oXML.ReadProperty("FieldID", mp_sFieldID);
	oXML.ReadProperty("ValueID", mp_lValueID);
	oXML.ReadProperty("ValueGUID", mp_lValueGUID);
}