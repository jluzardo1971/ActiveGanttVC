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
#include "ExtendedAttribute.h"

IMPLEMENT_DYNCREATE(ExtendedAttribute, CCmdTarget)

//{7C3CD3A9-D622-497A-BE35-0D719A709552}
static const IID IID_IExtendedAttribute = { 0x7C3CD3A9, 0xD622, 0x497A, { 0xBE, 0x35, 0x0D, 0x71, 0x9A, 0x70, 0x95, 0x52} };

//{ADA02125-DCF9-4D15-999B-C166B27D1D21}
IMPLEMENT_OLECREATE_FLAGS(ExtendedAttribute, "MSP2003.ExtendedAttribute", afxRegApartmentThreading, 0xADA02125, 0xDCF9, 0x4D15, 0x99, 0x9B, 0xC1, 0x66, 0xB2, 0x7D, 0x1D, 0x21)

BEGIN_DISPATCH_MAP(ExtendedAttribute, CCmdTarget)
DISP_PROPERTY_EX_ID(ExtendedAttribute, "sFieldID", 1, odl_GetFieldID, odl_SetFieldID, VT_BSTR)
DISP_PROPERTY_EX_ID(ExtendedAttribute, "sFieldName", 2, odl_GetFieldName, odl_SetFieldName, VT_BSTR)
DISP_PROPERTY_EX_ID(ExtendedAttribute, "sAlias", 3, odl_GetAlias, odl_SetAlias, VT_BSTR)
DISP_PROPERTY_EX_ID(ExtendedAttribute, "sPhoneticAlias", 4, odl_GetPhoneticAlias, odl_SetPhoneticAlias, VT_BSTR)
DISP_PROPERTY_EX_ID(ExtendedAttribute, "yRollupType", 5, odl_GetRollupType, odl_SetRollupType, VT_I4)
DISP_PROPERTY_EX_ID(ExtendedAttribute, "yCalculationType", 6, odl_GetCalculationType, odl_SetCalculationType, VT_I4)
DISP_PROPERTY_EX_ID(ExtendedAttribute, "sFormula", 7, odl_GetFormula, odl_SetFormula, VT_BSTR)
DISP_PROPERTY_EX_ID(ExtendedAttribute, "bRestrictValues", 8, odl_GetRestrictValues, odl_SetRestrictValues, VT_BOOL)
DISP_PROPERTY_EX_ID(ExtendedAttribute, "yValuelistSortOrder", 9, odl_GetValuelistSortOrder, odl_SetValuelistSortOrder, VT_I4)
DISP_PROPERTY_EX_ID(ExtendedAttribute, "bAppendNewValues", 10, odl_GetAppendNewValues, odl_SetAppendNewValues, VT_BOOL)
DISP_PROPERTY_EX_ID(ExtendedAttribute, "sDefault", 11, odl_GetDefault, odl_SetDefault, VT_BSTR)
DISP_PROPERTY_EX_ID(ExtendedAttribute, "oValueList", 12, odl_GetValueList, SetNotSupported, VT_DISPATCH)
DISP_PROPERTY_EX_ID(ExtendedAttribute, "Key", 13, odl_GetKey, odl_SetKey, VT_BSTR)
DISP_FUNCTION_ID(ExtendedAttribute, "GetXML", 14, odl_GetXML, VT_BSTR, VTS_NONE)
DISP_FUNCTION_ID(ExtendedAttribute, "SetXML", 15, odl_SetXML, VT_EMPTY, VTS_BSTR)
DISP_FUNCTION_ID(ExtendedAttribute, "IsNull", 16, IsNull, VT_BOOL, VTS_NONE)
DISP_FUNCTION_ID(ExtendedAttribute, "Initialize", 17, Initialize, VT_EMPTY, VTS_NONE)
END_DISPATCH_MAP()

BEGIN_INTERFACE_MAP(ExtendedAttribute, CCmdTarget)
	INTERFACE_PART(ExtendedAttribute, IID_IExtendedAttribute, Dispatch)
END_INTERFACE_MAP()

BEGIN_MESSAGE_MAP(ExtendedAttribute, CCmdTarget)
END_MESSAGE_MAP()


void ExtendedAttribute::Initialize(void)
{
	InitVars();
}

ExtendedAttribute::ExtendedAttribute()
{
	EnableAutomation();
	AfxOleLockApp();
	InitVars();
}

void ExtendedAttribute::InitVars(void)
{
	mp_sFieldID = "";
	mp_sFieldName = "";
	mp_sAlias = "";
	mp_sPhoneticAlias = "";
	mp_yRollupType = RT_MAXIMUM_OR_FOR_FLAG_FIELDS;
	mp_yCalculationType = CT_NONE;
	mp_sFormula = "";
	mp_bRestrictValues = FALSE;
	mp_yValuelistSortOrder = VSO_DESCENDING;
	mp_bAppendNewValues = FALSE;
	mp_sDefault = "";
	mp_oValueList = new ExtendedAttributeValueList();
}

ExtendedAttribute::~ExtendedAttribute()
{
	delete mp_oValueList;
	AfxOleUnlockApp();
}

void ExtendedAttribute::OnFinalRelease()
{
	clsItemBase::OnFinalRelease();
}

BSTR ExtendedAttribute::odl_GetFieldID(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetFieldID().AllocSysString();
}

CString ExtendedAttribute::GetFieldID(void)
{
	return (CString) mp_sFieldID;
}

void ExtendedAttribute::odl_SetFieldID(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetFieldID(newVal);
}

void ExtendedAttribute::SetFieldID(CString newVal)
{
	mp_sFieldID = newVal;
}

BSTR ExtendedAttribute::odl_GetFieldName(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetFieldName().AllocSysString();
}

CString ExtendedAttribute::GetFieldName(void)
{
	return (CString) mp_sFieldName;
}

void ExtendedAttribute::odl_SetFieldName(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetFieldName(newVal);
}

void ExtendedAttribute::SetFieldName(CString newVal)
{
	mp_sFieldName = newVal;
}

BSTR ExtendedAttribute::odl_GetAlias(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetAlias().AllocSysString();
}

CString ExtendedAttribute::GetAlias(void)
{
	return (CString) mp_sAlias;
}

void ExtendedAttribute::odl_SetAlias(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetAlias(newVal);
}

void ExtendedAttribute::SetAlias(CString newVal)
{
	if (newVal.GetLength() > 50)
	{
		newVal = newVal.Mid(0, 50);
	}
	mp_sAlias = newVal;
}

BSTR ExtendedAttribute::odl_GetPhoneticAlias(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetPhoneticAlias().AllocSysString();
}

CString ExtendedAttribute::GetPhoneticAlias(void)
{
	return (CString) mp_sPhoneticAlias;
}

void ExtendedAttribute::odl_SetPhoneticAlias(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetPhoneticAlias(newVal);
}

void ExtendedAttribute::SetPhoneticAlias(CString newVal)
{
	if (newVal.GetLength() > 50)
	{
		newVal = newVal.Mid(0, 50);
	}
	mp_sPhoneticAlias = newVal;
}

LONG ExtendedAttribute::odl_GetRollupType(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetRollupType();
}

E_ROLLUPTYPE ExtendedAttribute::GetRollupType(void)
{
	return (E_ROLLUPTYPE) mp_yRollupType;
}

void ExtendedAttribute::odl_SetRollupType(LONG newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetRollupType((E_ROLLUPTYPE)newVal);
}

void ExtendedAttribute::SetRollupType(E_ROLLUPTYPE newVal)
{
	mp_yRollupType = newVal;
}

LONG ExtendedAttribute::odl_GetCalculationType(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetCalculationType();
}

E_CALCULATIONTYPE ExtendedAttribute::GetCalculationType(void)
{
	return (E_CALCULATIONTYPE) mp_yCalculationType;
}

void ExtendedAttribute::odl_SetCalculationType(LONG newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetCalculationType((E_CALCULATIONTYPE)newVal);
}

void ExtendedAttribute::SetCalculationType(E_CALCULATIONTYPE newVal)
{
	mp_yCalculationType = newVal;
}

BSTR ExtendedAttribute::odl_GetFormula(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetFormula().AllocSysString();
}

CString ExtendedAttribute::GetFormula(void)
{
	return (CString) mp_sFormula;
}

void ExtendedAttribute::odl_SetFormula(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetFormula(newVal);
}

void ExtendedAttribute::SetFormula(CString newVal)
{
	mp_sFormula = newVal;
}

BOOL ExtendedAttribute::odl_GetRestrictValues(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetRestrictValues();
}

BOOL ExtendedAttribute::GetRestrictValues(void)
{
	return (BOOL) mp_bRestrictValues;
}

void ExtendedAttribute::odl_SetRestrictValues(BOOL newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetRestrictValues((BOOL)newVal);
}

void ExtendedAttribute::SetRestrictValues(BOOL newVal)
{
	mp_bRestrictValues = newVal;
}

LONG ExtendedAttribute::odl_GetValuelistSortOrder(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetValuelistSortOrder();
}

E_VALUELISTSORTORDER ExtendedAttribute::GetValuelistSortOrder(void)
{
	return (E_VALUELISTSORTORDER) mp_yValuelistSortOrder;
}

void ExtendedAttribute::odl_SetValuelistSortOrder(LONG newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetValuelistSortOrder((E_VALUELISTSORTORDER)newVal);
}

void ExtendedAttribute::SetValuelistSortOrder(E_VALUELISTSORTORDER newVal)
{
	mp_yValuelistSortOrder = newVal;
}

BOOL ExtendedAttribute::odl_GetAppendNewValues(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetAppendNewValues();
}

BOOL ExtendedAttribute::GetAppendNewValues(void)
{
	return (BOOL) mp_bAppendNewValues;
}

void ExtendedAttribute::odl_SetAppendNewValues(BOOL newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetAppendNewValues((BOOL)newVal);
}

void ExtendedAttribute::SetAppendNewValues(BOOL newVal)
{
	mp_bAppendNewValues = newVal;
}

BSTR ExtendedAttribute::odl_GetDefault(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetDefault().AllocSysString();
}

CString ExtendedAttribute::GetDefault(void)
{
	return (CString) mp_sDefault;
}

void ExtendedAttribute::odl_SetDefault(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetDefault(newVal);
}

void ExtendedAttribute::SetDefault(CString newVal)
{
	mp_sDefault = newVal;
}

IDispatch* ExtendedAttribute::odl_GetValueList(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetValueList()->GetIDispatch(TRUE);
}

ExtendedAttributeValueList* ExtendedAttribute::GetValueList(void)
{
	return mp_oValueList;
}

BSTR ExtendedAttribute::odl_GetKey(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetKey().AllocSysString();
}

void ExtendedAttribute::odl_SetKey(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	USES_CONVERSION;
	SetKey(F_TOUCSTR(T2A(newVal)));
}

CString ExtendedAttribute::GetKey(void)
{
	return mp_sKey;
}

void ExtendedAttribute::SetKey(CString newVal)
{
	mp_oCollection->mp_SetKey(&mp_sKey, newVal, MP_SET_KEY);
}

BSTR ExtendedAttribute::odl_GetXML(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetXML().AllocSysString();
}

void ExtendedAttribute::odl_SetXML(LPCTSTR sXML)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetXML(sXML);
}

BOOL ExtendedAttribute::IsNull(void)
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
	if (mp_yRollupType != RT_MAXIMUM_OR_FOR_FLAG_FIELDS)
	{
		bReturn = FALSE;
	}
	if (mp_yCalculationType != CT_NONE)
	{
		bReturn = FALSE;
	}
	if (mp_sFormula != "")
	{
		bReturn = FALSE;
	}
	if (mp_bRestrictValues != FALSE)
	{
		bReturn = FALSE;
	}
	if (mp_yValuelistSortOrder != VSO_DESCENDING)
	{
		bReturn = FALSE;
	}
	if (mp_bAppendNewValues != FALSE)
	{
		bReturn = FALSE;
	}
	if (mp_sDefault != "")
	{
		bReturn = FALSE;
	}
	if (mp_oValueList->IsNull() == FALSE)
	{
		bReturn = FALSE;
	}
	return bReturn;
}

CString ExtendedAttribute::GetXML(void)
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
	oXML.WriteProperty("RollupType", mp_yRollupType);
	oXML.WriteProperty("CalculationType", mp_yCalculationType);
	if (mp_sFormula != "")
	{
		oXML.WriteProperty("Formula", mp_sFormula);
	}
	oXML.WriteProperty("RestrictValues", mp_bRestrictValues);
	oXML.WriteProperty("ValuelistSortOrder", mp_yValuelistSortOrder);
	oXML.WriteProperty("AppendNewValues", mp_bAppendNewValues);
	if (mp_sDefault != "")
	{
		oXML.WriteProperty("Default", mp_sDefault);
	}
	if (mp_oValueList->IsNull() == FALSE)
	{
		oXML.WriteObject(mp_oValueList->GetXML());
	}
	return oXML.GetXML();
}

void ExtendedAttribute::SetXML(CString sXML)
{
	clsXML oXML("ExtendedAttribute");
	oXML.SetSupportOptional(TRUE);
	oXML.SetXML(sXML);
	oXML.InitializeReader();
	oXML.ReadProperty("FieldID", mp_sFieldID);
	oXML.ReadProperty("FieldName", mp_sFieldName);
	oXML.ReadProperty("Alias", mp_sAlias);
	if (mp_sAlias.GetLength() > 50)
	{
		mp_sAlias = mp_sAlias.Mid(0, 50);
	}
	oXML.ReadProperty("PhoneticAlias", mp_sPhoneticAlias);
	if (mp_sPhoneticAlias.GetLength() > 50)
	{
		mp_sPhoneticAlias = mp_sPhoneticAlias.Mid(0, 50);
	}
	oXML.ReadProperty("RollupType", mp_yRollupType);
	oXML.ReadProperty("CalculationType", mp_yCalculationType);
	oXML.ReadProperty("Formula", mp_sFormula);
	oXML.ReadProperty("RestrictValues", mp_bRestrictValues);
	oXML.ReadProperty("ValuelistSortOrder", mp_yValuelistSortOrder);
	oXML.ReadProperty("AppendNewValues", mp_bAppendNewValues);
	oXML.ReadProperty("Default", mp_sDefault);
	mp_oValueList->SetXML(oXML.ReadObject("ValueList"));
}