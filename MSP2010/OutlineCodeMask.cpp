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
#include "OutlineCodeMask.h"

IMPLEMENT_DYNCREATE(OutlineCodeMask, CCmdTarget)

//{1A02F843-DDAA-4313-9E25-E3284F9006A0}
static const IID IID_IOutlineCodeMask = { 0x1A02F843, 0xDDAA, 0x4313, { 0x9E, 0x25, 0xE3, 0x28, 0x4F, 0x90, 0x06, 0xA0} };

//{F2389227-EB75-4105-9FAD-BEAB7E4E5BB7}
IMPLEMENT_OLECREATE_FLAGS(OutlineCodeMask, "MSP2010.OutlineCodeMask", afxRegApartmentThreading, 0xF2389227, 0xEB75, 0x4105, 0x9F, 0xAD, 0xBE, 0xAB, 0x7E, 0x4E, 0x5B, 0xB7)

BEGIN_DISPATCH_MAP(OutlineCodeMask, CCmdTarget)
DISP_PROPERTY_EX_ID(OutlineCodeMask, "lLevel", 1, odl_GetLevel, odl_SetLevel, VT_I4)
DISP_PROPERTY_EX_ID(OutlineCodeMask, "yType", 2, odl_GetType, odl_SetType, VT_I4)
DISP_PROPERTY_EX_ID(OutlineCodeMask, "lLength", 3, odl_GetLength, odl_SetLength, VT_I4)
DISP_PROPERTY_EX_ID(OutlineCodeMask, "sSeparator", 4, odl_GetSeparator, odl_SetSeparator, VT_BSTR)
DISP_PROPERTY_EX_ID(OutlineCodeMask, "Key", 5, odl_GetKey, odl_SetKey, VT_BSTR)
DISP_FUNCTION_ID(OutlineCodeMask, "GetXML", 6, odl_GetXML, VT_BSTR, VTS_NONE)
DISP_FUNCTION_ID(OutlineCodeMask, "SetXML", 7, odl_SetXML, VT_EMPTY, VTS_BSTR)
DISP_FUNCTION_ID(OutlineCodeMask, "IsNull", 8, IsNull, VT_BOOL, VTS_NONE)
DISP_FUNCTION_ID(OutlineCodeMask, "Initialize", 9, Initialize, VT_EMPTY, VTS_NONE)
END_DISPATCH_MAP()

BEGIN_INTERFACE_MAP(OutlineCodeMask, CCmdTarget)
	INTERFACE_PART(OutlineCodeMask, IID_IOutlineCodeMask, Dispatch)
END_INTERFACE_MAP()

BEGIN_MESSAGE_MAP(OutlineCodeMask, CCmdTarget)
END_MESSAGE_MAP()


void OutlineCodeMask::Initialize(void)
{
	InitVars();
}

OutlineCodeMask::OutlineCodeMask()
{
	EnableAutomation();
	AfxOleLockApp();
	InitVars();
}

void OutlineCodeMask::InitVars(void)
{
	mp_lLevel = 0;
	mp_yType = T_1_NUMBERS;
	mp_lLength = 0;
	mp_sSeparator = "";
}

OutlineCodeMask::~OutlineCodeMask()
{
	AfxOleUnlockApp();
}

void OutlineCodeMask::OnFinalRelease()
{
	clsItemBase::OnFinalRelease();
}

LONG OutlineCodeMask::odl_GetLevel(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetLevel();
}

LONG OutlineCodeMask::GetLevel(void)
{
	return (LONG) mp_lLevel;
}

void OutlineCodeMask::odl_SetLevel(LONG newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetLevel((LONG)newVal);
}

void OutlineCodeMask::SetLevel(LONG newVal)
{
	mp_lLevel = newVal;
}

LONG OutlineCodeMask::odl_GetType(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetType();
}

E_TYPE_1 OutlineCodeMask::GetType(void)
{
	return (E_TYPE_1) mp_yType;
}

void OutlineCodeMask::odl_SetType(LONG newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetType((E_TYPE_1)newVal);
}

void OutlineCodeMask::SetType(E_TYPE_1 newVal)
{
	mp_yType = newVal;
}

LONG OutlineCodeMask::odl_GetLength(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetLength();
}

LONG OutlineCodeMask::GetLength(void)
{
	return (LONG) mp_lLength;
}

void OutlineCodeMask::odl_SetLength(LONG newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetLength((LONG)newVal);
}

void OutlineCodeMask::SetLength(LONG newVal)
{
	mp_lLength = newVal;
}

BSTR OutlineCodeMask::odl_GetSeparator(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetSeparator().AllocSysString();
}

CString OutlineCodeMask::GetSeparator(void)
{
	return (CString) mp_sSeparator;
}

void OutlineCodeMask::odl_SetSeparator(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetSeparator(newVal);
}

void OutlineCodeMask::SetSeparator(CString newVal)
{
	mp_sSeparator = newVal;
}

BSTR OutlineCodeMask::odl_GetKey(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetKey().AllocSysString();
}

void OutlineCodeMask::odl_SetKey(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	USES_CONVERSION;
	SetKey(F_TOUCSTR(T2A(newVal)));
}

CString OutlineCodeMask::GetKey(void)
{
	return mp_sKey;
}

void OutlineCodeMask::SetKey(CString newVal)
{
	mp_oCollection->mp_SetKey(&mp_sKey, newVal, MP_SET_KEY);
}

BSTR OutlineCodeMask::odl_GetXML(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetXML().AllocSysString();
}

void OutlineCodeMask::odl_SetXML(LPCTSTR sXML)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetXML(sXML);
}

BOOL OutlineCodeMask::IsNull(void)
{
	BOOL bReturn = TRUE;
	if (mp_lLevel != 0)
	{
		bReturn = FALSE;
	}
	if (mp_yType != T_1_NUMBERS)
	{
		bReturn = FALSE;
	}
	if (mp_lLength != 0)
	{
		bReturn = FALSE;
	}
	if (mp_sSeparator != "")
	{
		bReturn = FALSE;
	}
	return bReturn;
}

CString OutlineCodeMask::GetXML(void)
{
	if (IsNull() == TRUE)
	{
		return "<Mask/>";
	}
	clsXML oXML("Mask");
	oXML.InitializeWriter();
	oXML.SetSupportOptional(TRUE);
	oXML.SetBoolsAreNumeric(TRUE);
	oXML.WriteProperty("Level", mp_lLevel);
	oXML.WriteProperty("Type", mp_yType);
	oXML.WriteProperty("Length", mp_lLength);
	if (mp_sSeparator != "")
	{
		oXML.WriteProperty("Separator", mp_sSeparator);
	}
	return oXML.GetXML();
}

void OutlineCodeMask::SetXML(CString sXML)
{
	clsXML oXML("Mask");
	oXML.SetSupportOptional(TRUE);
	oXML.SetXML(sXML);
	oXML.InitializeReader();
	oXML.ReadProperty("Level", mp_lLevel);
	oXML.ReadProperty("Type", mp_yType);
	oXML.ReadProperty("Length", mp_lLength);
	oXML.ReadProperty("Separator", mp_sSeparator);
}