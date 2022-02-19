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
#include "WBSMask.h"

IMPLEMENT_DYNCREATE(WBSMask, CCmdTarget)

//{01E1AEBC-2969-4C2E-9931-7DDD21474B45}
static const IID IID_IWBSMask = { 0x01E1AEBC, 0x2969, 0x4C2E, { 0x99, 0x31, 0x7D, 0xDD, 0x21, 0x47, 0x4B, 0x45} };

//{365E1E96-263A-421E-B023-3BE1CE0AB3F6}
IMPLEMENT_OLECREATE_FLAGS(WBSMask, "MSP2007.WBSMask", afxRegApartmentThreading, 0x365E1E96, 0x263A, 0x421E, 0xB0, 0x23, 0x3B, 0xE1, 0xCE, 0x0A, 0xB3, 0xF6)

BEGIN_DISPATCH_MAP(WBSMask, CCmdTarget)
DISP_PROPERTY_EX_ID(WBSMask, "lLevel", 1, odl_GetLevel, odl_SetLevel, VT_I4)
DISP_PROPERTY_EX_ID(WBSMask, "yType", 2, odl_GetType, odl_SetType, VT_I4)
DISP_PROPERTY_EX_ID(WBSMask, "sLength", 3, odl_GetLength, odl_SetLength, VT_BSTR)
DISP_PROPERTY_EX_ID(WBSMask, "sSeparator", 4, odl_GetSeparator, odl_SetSeparator, VT_BSTR)
DISP_PROPERTY_EX_ID(WBSMask, "Key", 5, odl_GetKey, odl_SetKey, VT_BSTR)
DISP_FUNCTION_ID(WBSMask, "GetXML", 6, odl_GetXML, VT_BSTR, VTS_NONE)
DISP_FUNCTION_ID(WBSMask, "SetXML", 7, odl_SetXML, VT_EMPTY, VTS_BSTR)
DISP_FUNCTION_ID(WBSMask, "IsNull", 8, IsNull, VT_BOOL, VTS_NONE)
DISP_FUNCTION_ID(WBSMask, "Initialize", 9, Initialize, VT_EMPTY, VTS_NONE)
END_DISPATCH_MAP()

BEGIN_INTERFACE_MAP(WBSMask, CCmdTarget)
	INTERFACE_PART(WBSMask, IID_IWBSMask, Dispatch)
END_INTERFACE_MAP()

BEGIN_MESSAGE_MAP(WBSMask, CCmdTarget)
END_MESSAGE_MAP()


void WBSMask::Initialize(void)
{
	InitVars();
}

WBSMask::WBSMask()
{
	EnableAutomation();
	AfxOleLockApp();
	InitVars();
}

void WBSMask::InitVars(void)
{
	mp_lLevel = 0;
	mp_yType = T_2_NUMBERS;
	mp_sLength = "";
	mp_sSeparator = "";
}

WBSMask::~WBSMask()
{
	AfxOleUnlockApp();
}

void WBSMask::OnFinalRelease()
{
	clsItemBase::OnFinalRelease();
}

LONG WBSMask::odl_GetLevel(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetLevel();
}

LONG WBSMask::GetLevel(void)
{
	return (LONG) mp_lLevel;
}

void WBSMask::odl_SetLevel(LONG newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetLevel((LONG)newVal);
}

void WBSMask::SetLevel(LONG newVal)
{
	mp_lLevel = newVal;
}

LONG WBSMask::odl_GetType(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetType();
}

E_TYPE_2 WBSMask::GetType(void)
{
	return (E_TYPE_2) mp_yType;
}

void WBSMask::odl_SetType(LONG newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetType((E_TYPE_2)newVal);
}

void WBSMask::SetType(E_TYPE_2 newVal)
{
	mp_yType = newVal;
}

BSTR WBSMask::odl_GetLength(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetLength().AllocSysString();
}

CString WBSMask::GetLength(void)
{
	return (CString) mp_sLength;
}

void WBSMask::odl_SetLength(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetLength(newVal);
}

void WBSMask::SetLength(CString newVal)
{
	mp_sLength = newVal;
}

BSTR WBSMask::odl_GetSeparator(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetSeparator().AllocSysString();
}

CString WBSMask::GetSeparator(void)
{
	return (CString) mp_sSeparator;
}

void WBSMask::odl_SetSeparator(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetSeparator(newVal);
}

void WBSMask::SetSeparator(CString newVal)
{
	mp_sSeparator = newVal;
}

BSTR WBSMask::odl_GetKey(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetKey().AllocSysString();
}

void WBSMask::odl_SetKey(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	USES_CONVERSION;
	SetKey(F_TOUCSTR(T2A(newVal)));
}

CString WBSMask::GetKey(void)
{
	return mp_sKey;
}

void WBSMask::SetKey(CString newVal)
{
	mp_oCollection->mp_SetKey(&mp_sKey, newVal, MP_SET_KEY);
}

BSTR WBSMask::odl_GetXML(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetXML().AllocSysString();
}

void WBSMask::odl_SetXML(LPCTSTR sXML)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetXML(sXML);
}

BOOL WBSMask::IsNull(void)
{
	BOOL bReturn = TRUE;
	if (mp_lLevel != 0)
	{
		bReturn = FALSE;
	}
	if (mp_yType != T_2_NUMBERS)
	{
		bReturn = FALSE;
	}
	if (mp_sLength != "")
	{
		bReturn = FALSE;
	}
	if (mp_sSeparator != "")
	{
		bReturn = FALSE;
	}
	return bReturn;
}

CString WBSMask::GetXML(void)
{
	if (IsNull() == TRUE)
	{
		return "<WBSMask/>";
	}
	clsXML oXML("WBSMask");
	oXML.InitializeWriter();
	oXML.SetSupportOptional(TRUE);
	oXML.SetBoolsAreNumeric(TRUE);
	oXML.WriteProperty("Level", mp_lLevel);
	oXML.WriteProperty("Type", mp_yType);
	oXML.WriteProperty("Length", mp_sLength);
	oXML.WriteProperty("Separator", mp_sSeparator);
	return oXML.GetXML();
}

void WBSMask::SetXML(CString sXML)
{
	clsXML oXML("WBSMask");
	oXML.SetSupportOptional(TRUE);
	oXML.SetXML(sXML);
	oXML.InitializeReader();
	oXML.ReadProperty("Level", mp_lLevel);
	oXML.ReadProperty("Type", mp_yType);
	oXML.ReadProperty("Length", mp_sLength);
	oXML.ReadProperty("Separator", mp_sSeparator);
}