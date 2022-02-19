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
#include "OutlineCodeValue.h"

IMPLEMENT_DYNCREATE(OutlineCodeValue, CCmdTarget)

//{A5F3C819-A3B5-47CA-818F-780C357EEA88}
static const IID IID_IOutlineCodeValue = { 0xA5F3C819, 0xA3B5, 0x47CA, { 0x81, 0x8F, 0x78, 0x0C, 0x35, 0x7E, 0xEA, 0x88} };

//{1C8397F3-93F4-48AA-A28D-CC9262693B74}
IMPLEMENT_OLECREATE_FLAGS(OutlineCodeValue, "MSP2003.OutlineCodeValue", afxRegApartmentThreading, 0x1C8397F3, 0x93F4, 0x48AA, 0xA2, 0x8D, 0xCC, 0x92, 0x62, 0x69, 0x3B, 0x74)

BEGIN_DISPATCH_MAP(OutlineCodeValue, CCmdTarget)
DISP_PROPERTY_EX_ID(OutlineCodeValue, "lValueID", 1, odl_GetValueID, odl_SetValueID, VT_I4)
DISP_PROPERTY_EX_ID(OutlineCodeValue, "lParentValueID", 2, odl_GetParentValueID, odl_SetParentValueID, VT_I4)
DISP_PROPERTY_EX_ID(OutlineCodeValue, "sValue", 3, odl_GetValue, odl_SetValue, VT_BSTR)
DISP_PROPERTY_EX_ID(OutlineCodeValue, "sDescription", 4, odl_GetDescription, odl_SetDescription, VT_BSTR)
DISP_PROPERTY_EX_ID(OutlineCodeValue, "Key", 5, odl_GetKey, odl_SetKey, VT_BSTR)
DISP_FUNCTION_ID(OutlineCodeValue, "GetXML", 6, odl_GetXML, VT_BSTR, VTS_NONE)
DISP_FUNCTION_ID(OutlineCodeValue, "SetXML", 7, odl_SetXML, VT_EMPTY, VTS_BSTR)
DISP_FUNCTION_ID(OutlineCodeValue, "IsNull", 8, IsNull, VT_BOOL, VTS_NONE)
DISP_FUNCTION_ID(OutlineCodeValue, "Initialize", 9, Initialize, VT_EMPTY, VTS_NONE)
END_DISPATCH_MAP()

BEGIN_INTERFACE_MAP(OutlineCodeValue, CCmdTarget)
	INTERFACE_PART(OutlineCodeValue, IID_IOutlineCodeValue, Dispatch)
END_INTERFACE_MAP()

BEGIN_MESSAGE_MAP(OutlineCodeValue, CCmdTarget)
END_MESSAGE_MAP()


void OutlineCodeValue::Initialize(void)
{
	InitVars();
}

OutlineCodeValue::OutlineCodeValue()
{
	EnableAutomation();
	AfxOleLockApp();
	InitVars();
}

void OutlineCodeValue::InitVars(void)
{
	mp_lValueID = 0;
	mp_lParentValueID = 0;
	mp_sValue = "";
	mp_sDescription = "";
}

OutlineCodeValue::~OutlineCodeValue()
{
	AfxOleUnlockApp();
}

void OutlineCodeValue::OnFinalRelease()
{
	clsItemBase::OnFinalRelease();
}

LONG OutlineCodeValue::odl_GetValueID(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetValueID();
}

LONG OutlineCodeValue::GetValueID(void)
{
	return (LONG) mp_lValueID;
}

void OutlineCodeValue::odl_SetValueID(LONG newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetValueID((LONG)newVal);
}

void OutlineCodeValue::SetValueID(LONG newVal)
{
	mp_lValueID = newVal;
}

LONG OutlineCodeValue::odl_GetParentValueID(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetParentValueID();
}

LONG OutlineCodeValue::GetParentValueID(void)
{
	return (LONG) mp_lParentValueID;
}

void OutlineCodeValue::odl_SetParentValueID(LONG newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetParentValueID((LONG)newVal);
}

void OutlineCodeValue::SetParentValueID(LONG newVal)
{
	mp_lParentValueID = newVal;
}

BSTR OutlineCodeValue::odl_GetValue(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetValue().AllocSysString();
}

CString OutlineCodeValue::GetValue(void)
{
	return (CString) mp_sValue;
}

void OutlineCodeValue::odl_SetValue(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetValue(newVal);
}

void OutlineCodeValue::SetValue(CString newVal)
{
	mp_sValue = newVal;
}

BSTR OutlineCodeValue::odl_GetDescription(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetDescription().AllocSysString();
}

CString OutlineCodeValue::GetDescription(void)
{
	return (CString) mp_sDescription;
}

void OutlineCodeValue::odl_SetDescription(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetDescription(newVal);
}

void OutlineCodeValue::SetDescription(CString newVal)
{
	mp_sDescription = newVal;
}

BSTR OutlineCodeValue::odl_GetKey(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetKey().AllocSysString();
}

void OutlineCodeValue::odl_SetKey(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	USES_CONVERSION;
	SetKey(F_TOUCSTR(T2A(newVal)));
}

CString OutlineCodeValue::GetKey(void)
{
	return mp_sKey;
}

void OutlineCodeValue::SetKey(CString newVal)
{
	mp_oCollection->mp_SetKey(&mp_sKey, newVal, MP_SET_KEY);
}

BSTR OutlineCodeValue::odl_GetXML(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetXML().AllocSysString();
}

void OutlineCodeValue::odl_SetXML(LPCTSTR sXML)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetXML(sXML);
}

BOOL OutlineCodeValue::IsNull(void)
{
	BOOL bReturn = TRUE;
	if (mp_lValueID != 0)
	{
		bReturn = FALSE;
	}
	if (mp_lParentValueID != 0)
	{
		bReturn = FALSE;
	}
	if (mp_sValue != "")
	{
		bReturn = FALSE;
	}
	if (mp_sDescription != "")
	{
		bReturn = FALSE;
	}
	return bReturn;
}

CString OutlineCodeValue::GetXML(void)
{
	if (IsNull() == TRUE)
	{
		return "<Value/>";
	}
	clsXML oXML("Value");
	oXML.InitializeWriter();
	oXML.SetSupportOptional(TRUE);
	oXML.SetBoolsAreNumeric(TRUE);
	oXML.WriteProperty("ValueID", mp_lValueID);
	oXML.WriteProperty("ParentValueID", mp_lParentValueID);
	if (mp_sValue != "")
	{
		oXML.WriteProperty("Value", mp_sValue);
	}
	if (mp_sDescription != "")
	{
		oXML.WriteProperty("Description", mp_sDescription);
	}
	return oXML.GetXML();
}

void OutlineCodeValue::SetXML(CString sXML)
{
	clsXML oXML("Value");
	oXML.SetSupportOptional(TRUE);
	oXML.SetXML(sXML);
	oXML.InitializeReader();
	oXML.ReadProperty("ValueID", mp_lValueID);
	oXML.ReadProperty("ParentValueID", mp_lParentValueID);
	oXML.ReadProperty("Value", mp_sValue);
	oXML.ReadProperty("Description", mp_sDescription);
}