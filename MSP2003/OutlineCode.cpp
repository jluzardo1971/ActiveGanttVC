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
#include "OutlineCode.h"

IMPLEMENT_DYNCREATE(OutlineCode, CCmdTarget)

//{58F926AC-76C9-4064-A3BE-344658862201}
static const IID IID_IOutlineCode = { 0x58F926AC, 0x76C9, 0x4064, { 0xA3, 0xBE, 0x34, 0x46, 0x58, 0x86, 0x22, 0x01} };

//{DEA5D074-702D-4CA3-B550-4A40DACB90E1}
IMPLEMENT_OLECREATE_FLAGS(OutlineCode, "MSP2003.OutlineCode", afxRegApartmentThreading, 0xDEA5D074, 0x702D, 0x4CA3, 0xB5, 0x50, 0x4A, 0x40, 0xDA, 0xCB, 0x90, 0xE1)

BEGIN_DISPATCH_MAP(OutlineCode, CCmdTarget)
DISP_PROPERTY_EX_ID(OutlineCode, "sFieldID", 1, odl_GetFieldID, odl_SetFieldID, VT_BSTR)
DISP_PROPERTY_EX_ID(OutlineCode, "sFieldName", 2, odl_GetFieldName, odl_SetFieldName, VT_BSTR)
DISP_PROPERTY_EX_ID(OutlineCode, "sAlias", 3, odl_GetAlias, odl_SetAlias, VT_BSTR)
DISP_PROPERTY_EX_ID(OutlineCode, "sPhoneticAlias", 4, odl_GetPhoneticAlias, odl_SetPhoneticAlias, VT_BSTR)
DISP_PROPERTY_EX_ID(OutlineCode, "oValues", 5, odl_GetValues, SetNotSupported, VT_DISPATCH)
DISP_PROPERTY_EX_ID(OutlineCode, "bEnterprise", 6, odl_GetEnterprise, odl_SetEnterprise, VT_BOOL)
DISP_PROPERTY_EX_ID(OutlineCode, "lEnterpriseOutlineCodeAlias", 7, odl_GetEnterpriseOutlineCodeAlias, odl_SetEnterpriseOutlineCodeAlias, VT_I4)
DISP_PROPERTY_EX_ID(OutlineCode, "bResourceSubstitutionEnabled", 8, odl_GetResourceSubstitutionEnabled, odl_SetResourceSubstitutionEnabled, VT_BOOL)
DISP_PROPERTY_EX_ID(OutlineCode, "bLeafOnly", 9, odl_GetLeafOnly, odl_SetLeafOnly, VT_BOOL)
DISP_PROPERTY_EX_ID(OutlineCode, "bAllLevelsRequired", 10, odl_GetAllLevelsRequired, odl_SetAllLevelsRequired, VT_BOOL)
DISP_PROPERTY_EX_ID(OutlineCode, "bOnlyTableValuesAllowed", 11, odl_GetOnlyTableValuesAllowed, odl_SetOnlyTableValuesAllowed, VT_BOOL)
DISP_PROPERTY_EX_ID(OutlineCode, "oMasks", 12, odl_GetMasks, SetNotSupported, VT_DISPATCH)
DISP_PROPERTY_EX_ID(OutlineCode, "Key", 13, odl_GetKey, odl_SetKey, VT_BSTR)
DISP_FUNCTION_ID(OutlineCode, "GetXML", 14, odl_GetXML, VT_BSTR, VTS_NONE)
DISP_FUNCTION_ID(OutlineCode, "SetXML", 15, odl_SetXML, VT_EMPTY, VTS_BSTR)
DISP_FUNCTION_ID(OutlineCode, "IsNull", 16, IsNull, VT_BOOL, VTS_NONE)
DISP_FUNCTION_ID(OutlineCode, "Initialize", 17, Initialize, VT_EMPTY, VTS_NONE)
END_DISPATCH_MAP()

BEGIN_INTERFACE_MAP(OutlineCode, CCmdTarget)
	INTERFACE_PART(OutlineCode, IID_IOutlineCode, Dispatch)
END_INTERFACE_MAP()

BEGIN_MESSAGE_MAP(OutlineCode, CCmdTarget)
END_MESSAGE_MAP()


void OutlineCode::Initialize(void)
{
	InitVars();
}

OutlineCode::OutlineCode()
{
	EnableAutomation();
	AfxOleLockApp();
	InitVars();
}

void OutlineCode::InitVars(void)
{
	mp_sFieldID = "";
	mp_sFieldName = "";
	mp_sAlias = "";
	mp_sPhoneticAlias = "";
	mp_oValues = new OutlineCodeValues();
	mp_bEnterprise = FALSE;
	mp_lEnterpriseOutlineCodeAlias = 0;
	mp_bResourceSubstitutionEnabled = FALSE;
	mp_bLeafOnly = FALSE;
	mp_bAllLevelsRequired = FALSE;
	mp_bOnlyTableValuesAllowed = FALSE;
	mp_oMasks = new OutlineCodeMasks();
}

OutlineCode::~OutlineCode()
{
	delete mp_oValues;
	delete mp_oMasks;
	AfxOleUnlockApp();
}

void OutlineCode::OnFinalRelease()
{
	clsItemBase::OnFinalRelease();
}

BSTR OutlineCode::odl_GetFieldID(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetFieldID().AllocSysString();
}

CString OutlineCode::GetFieldID(void)
{
	return (CString) mp_sFieldID;
}

void OutlineCode::odl_SetFieldID(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetFieldID(newVal);
}

void OutlineCode::SetFieldID(CString newVal)
{
	mp_sFieldID = newVal;
}

BSTR OutlineCode::odl_GetFieldName(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetFieldName().AllocSysString();
}

CString OutlineCode::GetFieldName(void)
{
	return (CString) mp_sFieldName;
}

void OutlineCode::odl_SetFieldName(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetFieldName(newVal);
}

void OutlineCode::SetFieldName(CString newVal)
{
	mp_sFieldName = newVal;
}

BSTR OutlineCode::odl_GetAlias(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetAlias().AllocSysString();
}

CString OutlineCode::GetAlias(void)
{
	return (CString) mp_sAlias;
}

void OutlineCode::odl_SetAlias(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetAlias(newVal);
}

void OutlineCode::SetAlias(CString newVal)
{
	mp_sAlias = newVal;
}

BSTR OutlineCode::odl_GetPhoneticAlias(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetPhoneticAlias().AllocSysString();
}

CString OutlineCode::GetPhoneticAlias(void)
{
	return (CString) mp_sPhoneticAlias;
}

void OutlineCode::odl_SetPhoneticAlias(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetPhoneticAlias(newVal);
}

void OutlineCode::SetPhoneticAlias(CString newVal)
{
	mp_sPhoneticAlias = newVal;
}

IDispatch* OutlineCode::odl_GetValues(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetValues()->GetIDispatch(TRUE);
}

OutlineCodeValues* OutlineCode::GetValues(void)
{
	return mp_oValues;
}

BOOL OutlineCode::odl_GetEnterprise(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetEnterprise();
}

BOOL OutlineCode::GetEnterprise(void)
{
	return (BOOL) mp_bEnterprise;
}

void OutlineCode::odl_SetEnterprise(BOOL newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetEnterprise((BOOL)newVal);
}

void OutlineCode::SetEnterprise(BOOL newVal)
{
	mp_bEnterprise = newVal;
}

LONG OutlineCode::odl_GetEnterpriseOutlineCodeAlias(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetEnterpriseOutlineCodeAlias();
}

LONG OutlineCode::GetEnterpriseOutlineCodeAlias(void)
{
	return (LONG) mp_lEnterpriseOutlineCodeAlias;
}

void OutlineCode::odl_SetEnterpriseOutlineCodeAlias(LONG newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetEnterpriseOutlineCodeAlias((LONG)newVal);
}

void OutlineCode::SetEnterpriseOutlineCodeAlias(LONG newVal)
{
	mp_lEnterpriseOutlineCodeAlias = newVal;
}

BOOL OutlineCode::odl_GetResourceSubstitutionEnabled(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetResourceSubstitutionEnabled();
}

BOOL OutlineCode::GetResourceSubstitutionEnabled(void)
{
	return (BOOL) mp_bResourceSubstitutionEnabled;
}

void OutlineCode::odl_SetResourceSubstitutionEnabled(BOOL newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetResourceSubstitutionEnabled((BOOL)newVal);
}

void OutlineCode::SetResourceSubstitutionEnabled(BOOL newVal)
{
	mp_bResourceSubstitutionEnabled = newVal;
}

BOOL OutlineCode::odl_GetLeafOnly(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetLeafOnly();
}

BOOL OutlineCode::GetLeafOnly(void)
{
	return (BOOL) mp_bLeafOnly;
}

void OutlineCode::odl_SetLeafOnly(BOOL newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetLeafOnly((BOOL)newVal);
}

void OutlineCode::SetLeafOnly(BOOL newVal)
{
	mp_bLeafOnly = newVal;
}

BOOL OutlineCode::odl_GetAllLevelsRequired(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetAllLevelsRequired();
}

BOOL OutlineCode::GetAllLevelsRequired(void)
{
	return (BOOL) mp_bAllLevelsRequired;
}

void OutlineCode::odl_SetAllLevelsRequired(BOOL newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetAllLevelsRequired((BOOL)newVal);
}

void OutlineCode::SetAllLevelsRequired(BOOL newVal)
{
	mp_bAllLevelsRequired = newVal;
}

BOOL OutlineCode::odl_GetOnlyTableValuesAllowed(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetOnlyTableValuesAllowed();
}

BOOL OutlineCode::GetOnlyTableValuesAllowed(void)
{
	return (BOOL) mp_bOnlyTableValuesAllowed;
}

void OutlineCode::odl_SetOnlyTableValuesAllowed(BOOL newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetOnlyTableValuesAllowed((BOOL)newVal);
}

void OutlineCode::SetOnlyTableValuesAllowed(BOOL newVal)
{
	mp_bOnlyTableValuesAllowed = newVal;
}

IDispatch* OutlineCode::odl_GetMasks(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetMasks()->GetIDispatch(TRUE);
}

OutlineCodeMasks* OutlineCode::GetMasks(void)
{
	return mp_oMasks;
}

BSTR OutlineCode::odl_GetKey(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetKey().AllocSysString();
}

void OutlineCode::odl_SetKey(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	USES_CONVERSION;
	SetKey(F_TOUCSTR(T2A(newVal)));
}

CString OutlineCode::GetKey(void)
{
	return mp_sKey;
}

void OutlineCode::SetKey(CString newVal)
{
	mp_oCollection->mp_SetKey(&mp_sKey, newVal, MP_SET_KEY);
}

BSTR OutlineCode::odl_GetXML(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetXML().AllocSysString();
}

void OutlineCode::odl_SetXML(LPCTSTR sXML)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetXML(sXML);
}

BOOL OutlineCode::IsNull(void)
{
	BOOL bReturn = TRUE;
	if (mp_sFieldID != "")
	{
		bReturn = FALSE;
	}
	if (mp_sFieldName != "")
	{
		bReturn = FALSE;
	}
	if (mp_sAlias != "")
	{
		bReturn = FALSE;
	}
	if (mp_sPhoneticAlias != "")
	{
		bReturn = FALSE;
	}
	if (mp_oValues->IsNull() == FALSE)
	{
		bReturn = FALSE;
	}
	if (mp_bEnterprise != FALSE)
	{
		bReturn = FALSE;
	}
	if (mp_lEnterpriseOutlineCodeAlias != 0)
	{
		bReturn = FALSE;
	}
	if (mp_bResourceSubstitutionEnabled != FALSE)
	{
		bReturn = FALSE;
	}
	if (mp_bLeafOnly != FALSE)
	{
		bReturn = FALSE;
	}
	if (mp_bAllLevelsRequired != FALSE)
	{
		bReturn = FALSE;
	}
	if (mp_bOnlyTableValuesAllowed != FALSE)
	{
		bReturn = FALSE;
	}
	if (mp_oMasks->IsNull() == FALSE)
	{
		bReturn = FALSE;
	}
	return bReturn;
}

CString OutlineCode::GetXML(void)
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
	if (mp_sFieldName != "")
	{
		oXML.WriteProperty("FieldName", mp_sFieldName);
	}
	if (mp_sAlias != "")
	{
		oXML.WriteProperty("Alias", mp_sAlias);
	}
	if (mp_sPhoneticAlias != "")
	{
		oXML.WriteProperty("PhoneticAlias", mp_sPhoneticAlias);
	}
	if (mp_oValues->IsNull() == FALSE)
	{
		oXML.WriteObject(mp_oValues->GetXML());
	}
	oXML.WriteProperty("Enterprise", mp_bEnterprise);
	oXML.WriteProperty("EnterpriseOutlineCodeAlias", mp_lEnterpriseOutlineCodeAlias);
	oXML.WriteProperty("ResourceSubstitutionEnabled", mp_bResourceSubstitutionEnabled);
	oXML.WriteProperty("LeafOnly", mp_bLeafOnly);
	oXML.WriteProperty("AllLevelsRequired", mp_bAllLevelsRequired);
	oXML.WriteProperty("OnlyTableValuesAllowed", mp_bOnlyTableValuesAllowed);
	if (mp_oMasks->IsNull() == FALSE)
	{
		oXML.WriteObject(mp_oMasks->GetXML());
	}
	return oXML.GetXML();
}

void OutlineCode::SetXML(CString sXML)
{
	clsXML oXML("OutlineCode");
	oXML.SetSupportOptional(TRUE);
	oXML.SetXML(sXML);
	oXML.InitializeReader();
	oXML.ReadProperty("FieldID", mp_sFieldID);
	oXML.ReadProperty("FieldName", mp_sFieldName);
	oXML.ReadProperty("Alias", mp_sAlias);
	oXML.ReadProperty("PhoneticAlias", mp_sPhoneticAlias);
	mp_oValues->SetXML(oXML.ReadObject("Values"));
	oXML.ReadProperty("Enterprise", mp_bEnterprise);
	oXML.ReadProperty("EnterpriseOutlineCodeAlias", mp_lEnterpriseOutlineCodeAlias);
	oXML.ReadProperty("ResourceSubstitutionEnabled", mp_bResourceSubstitutionEnabled);
	oXML.ReadProperty("LeafOnly", mp_bLeafOnly);
	oXML.ReadProperty("AllLevelsRequired", mp_bAllLevelsRequired);
	oXML.ReadProperty("OnlyTableValuesAllowed", mp_bOnlyTableValuesAllowed);
	mp_oMasks->SetXML(oXML.ReadObject("Masks"));
}