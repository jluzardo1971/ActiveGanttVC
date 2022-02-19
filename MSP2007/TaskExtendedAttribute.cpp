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

//{3BF1CF4A-B1BC-435E-A5C2-F164C3E9B73B}
static const IID IID_ITaskExtendedAttribute = { 0x3BF1CF4A, 0xB1BC, 0x435E, { 0xA5, 0xC2, 0xF1, 0x64, 0xC3, 0xE9, 0xB7, 0x3B} };

//{FB38BADA-DC3C-401E-A6F6-984A3AAB0103}
IMPLEMENT_OLECREATE_FLAGS(TaskExtendedAttribute, "MSP2007.TaskExtendedAttribute", afxRegApartmentThreading, 0xFB38BADA, 0xDC3C, 0x401E, 0xA6, 0xF6, 0x98, 0x4A, 0x3A, 0xAB, 0x01, 0x03)

BEGIN_DISPATCH_MAP(TaskExtendedAttribute, CCmdTarget)
DISP_PROPERTY_EX_ID(TaskExtendedAttribute, "sFieldID", 1, odl_GetFieldID, odl_SetFieldID, VT_BSTR)
DISP_PROPERTY_EX_ID(TaskExtendedAttribute, "sValue", 2, odl_GetValue, odl_SetValue, VT_BSTR)
DISP_PROPERTY_EX_ID(TaskExtendedAttribute, "lValueGUID", 3, odl_GetValueGUID, odl_SetValueGUID, VT_I4)
DISP_PROPERTY_EX_ID(TaskExtendedAttribute, "yDurationFormat", 4, odl_GetDurationFormat, odl_SetDurationFormat, VT_I4)
DISP_PROPERTY_EX_ID(TaskExtendedAttribute, "Key", 5, odl_GetKey, odl_SetKey, VT_BSTR)
DISP_FUNCTION_ID(TaskExtendedAttribute, "GetXML", 6, odl_GetXML, VT_BSTR, VTS_NONE)
DISP_FUNCTION_ID(TaskExtendedAttribute, "SetXML", 7, odl_SetXML, VT_EMPTY, VTS_BSTR)
DISP_FUNCTION_ID(TaskExtendedAttribute, "IsNull", 8, IsNull, VT_BOOL, VTS_NONE)
DISP_FUNCTION_ID(TaskExtendedAttribute, "Initialize", 9, Initialize, VT_EMPTY, VTS_NONE)
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
	mp_sFieldID = "";
	mp_sValue = "";
	mp_lValueGUID = 0;
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

LONG TaskExtendedAttribute::odl_GetValueGUID(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetValueGUID();
}

LONG TaskExtendedAttribute::GetValueGUID(void)
{
	return (LONG) mp_lValueGUID;
}

void TaskExtendedAttribute::odl_SetValueGUID(LONG newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetValueGUID((LONG)newVal);
}

void TaskExtendedAttribute::SetValueGUID(LONG newVal)
{
	mp_lValueGUID = newVal;
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

void TaskExtendedAttribute::SetXML(CString sXML)
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