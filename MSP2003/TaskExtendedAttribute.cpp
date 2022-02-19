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
#include "TaskExtendedAttribute.h"

IMPLEMENT_DYNCREATE(TaskExtendedAttribute, CCmdTarget)

//{FBF04692-829E-4BC8-ABB2-8BCA750C00AA}
static const IID IID_ITaskExtendedAttribute = { 0xFBF04692, 0x829E, 0x4BC8, { 0xAB, 0xB2, 0x8B, 0xCA, 0x75, 0x0C, 0x00, 0xAA} };

//{13338FBE-7F80-4FD6-BFFD-4FB07E30E124}
IMPLEMENT_OLECREATE_FLAGS(TaskExtendedAttribute, "MSP2003.TaskExtendedAttribute", afxRegApartmentThreading, 0x13338FBE, 0x7F80, 0x4FD6, 0xBF, 0xFD, 0x4F, 0xB0, 0x7E, 0x30, 0xE1, 0x24)

BEGIN_DISPATCH_MAP(TaskExtendedAttribute, CCmdTarget)
DISP_PROPERTY_EX_ID(TaskExtendedAttribute, "lUID", 1, odl_GetUID, odl_SetUID, VT_I4)
DISP_PROPERTY_EX_ID(TaskExtendedAttribute, "sFieldID", 2, odl_GetFieldID, odl_SetFieldID, VT_BSTR)
DISP_PROPERTY_EX_ID(TaskExtendedAttribute, "sValue", 3, odl_GetValue, odl_SetValue, VT_BSTR)
DISP_PROPERTY_EX_ID(TaskExtendedAttribute, "lValueID", 4, odl_GetValueID, odl_SetValueID, VT_I4)
DISP_PROPERTY_EX_ID(TaskExtendedAttribute, "yDurationFormat", 5, odl_GetDurationFormat, odl_SetDurationFormat, VT_I4)
DISP_PROPERTY_EX_ID(TaskExtendedAttribute, "Key", 6, odl_GetKey, odl_SetKey, VT_BSTR)
DISP_FUNCTION_ID(TaskExtendedAttribute, "GetXML", 7, odl_GetXML, VT_BSTR, VTS_NONE)
DISP_FUNCTION_ID(TaskExtendedAttribute, "SetXML", 8, odl_SetXML, VT_EMPTY, VTS_BSTR)
DISP_FUNCTION_ID(TaskExtendedAttribute, "IsNull", 9, IsNull, VT_BOOL, VTS_NONE)
DISP_FUNCTION_ID(TaskExtendedAttribute, "Initialize", 10, Initialize, VT_EMPTY, VTS_NONE)
END_DISPATCH_MAP()

BEGIN_INTERFACE_MAP(TaskExtendedAttribute, CCmdTarget)
	INTERFACE_PART(TaskExtendedAttribute, IID_ITaskExtendedAttribute, Dispatch)
END_INTERFACE_MAP()

BEGIN_MESSAGE_MAP(TaskExtendedAttribute, CCmdTarget)
END_MESSAGE_MAP()


void TaskExtendedAttribute::Initialize(void)
{
	InitVars();
}

TaskExtendedAttribute::TaskExtendedAttribute()
{
	EnableAutomation();
	AfxOleLockApp();
	InitVars();
}

void TaskExtendedAttribute::InitVars(void)
{
	mp_lUID = 0;
	mp_sFieldID = "";
	mp_sValue = "";
	mp_lValueID = 0;
	mp_yDurationFormat = DF_M;
}

TaskExtendedAttribute::~TaskExtendedAttribute()
{
	AfxOleUnlockApp();
}

void TaskExtendedAttribute::OnFinalRelease()
{
	clsItemBase::OnFinalRelease();
}

LONG TaskExtendedAttribute::odl_GetUID(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetUID();
}

LONG TaskExtendedAttribute::GetUID(void)
{
	return (LONG) mp_lUID;
}

void TaskExtendedAttribute::odl_SetUID(LONG newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetUID((LONG)newVal);
}

void TaskExtendedAttribute::SetUID(LONG newVal)
{
	mp_lUID = newVal;
}

BSTR TaskExtendedAttribute::odl_GetFieldID(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetFieldID().AllocSysString();
}

CString TaskExtendedAttribute::GetFieldID(void)
{
	return (CString) mp_sFieldID;
}

void TaskExtendedAttribute::odl_SetFieldID(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetFieldID(newVal);
}

void TaskExtendedAttribute::SetFieldID(CString newVal)
{
	mp_sFieldID = newVal;
}

BSTR TaskExtendedAttribute::odl_GetValue(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetValue().AllocSysString();
}

CString TaskExtendedAttribute::GetValue(void)
{
	return (CString) mp_sValue;
}

void TaskExtendedAttribute::odl_SetValue(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetValue(newVal);
}

void TaskExtendedAttribute::SetValue(CString newVal)
{
	mp_sValue = newVal;
}

LONG TaskExtendedAttribute::odl_GetValueID(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetValueID();
}

LONG TaskExtendedAttribute::GetValueID(void)
{
	return (LONG) mp_lValueID;
}

void TaskExtendedAttribute::odl_SetValueID(LONG newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetValueID((LONG)newVal);
}

void TaskExtendedAttribute::SetValueID(LONG newVal)
{
	mp_lValueID = newVal;
}

LONG TaskExtendedAttribute::odl_GetDurationFormat(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetDurationFormat();
}

E_DURATIONFORMAT TaskExtendedAttribute::GetDurationFormat(void)
{
	return (E_DURATIONFORMAT) mp_yDurationFormat;
}

void TaskExtendedAttribute::odl_SetDurationFormat(LONG newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetDurationFormat((E_DURATIONFORMAT)newVal);
}

void TaskExtendedAttribute::SetDurationFormat(E_DURATIONFORMAT newVal)
{
	mp_yDurationFormat = newVal;
}

BSTR TaskExtendedAttribute::odl_GetKey(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetKey().AllocSysString();
}

void TaskExtendedAttribute::odl_SetKey(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	USES_CONVERSION;
	SetKey(F_TOUCSTR(T2A(newVal)));
}

CString TaskExtendedAttribute::GetKey(void)
{
	return mp_sKey;
}

void TaskExtendedAttribute::SetKey(CString newVal)
{
	mp_oCollection->mp_SetKey(&mp_sKey, newVal, MP_SET_KEY);
}

BSTR TaskExtendedAttribute::odl_GetXML(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetXML().AllocSysString();
}

void TaskExtendedAttribute::odl_SetXML(LPCTSTR sXML)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetXML(sXML);
}

BOOL TaskExtendedAttribute::IsNull(void)
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

CString TaskExtendedAttribute::GetXML(void)
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

void TaskExtendedAttribute::SetXML(CString sXML)
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