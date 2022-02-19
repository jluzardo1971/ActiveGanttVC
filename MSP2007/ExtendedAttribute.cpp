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

//{EAA8EBB4-F014-4940-B097-684FBA887E51}
static const IID IID_IExtendedAttribute = { 0xEAA8EBB4, 0xF014, 0x4940, { 0xB0, 0x97, 0x68, 0x4F, 0xBA, 0x88, 0x7E, 0x51} };

//{4E42608D-DF44-4D0A-B917-5C1977006F1D}
IMPLEMENT_OLECREATE_FLAGS(ExtendedAttribute, "MSP2007.ExtendedAttribute", afxRegApartmentThreading, 0x4E42608D, 0xDF44, 0x4D0A, 0xB9, 0x17, 0x5C, 0x19, 0x77, 0x00, 0x6F, 0x1D)

BEGIN_DISPATCH_MAP(ExtendedAttribute, CCmdTarget)
DISP_PROPERTY_EX_ID(ExtendedAttribute, "sFieldID", 1, odl_GetFieldID, odl_SetFieldID, VT_BSTR)
DISP_PROPERTY_EX_ID(ExtendedAttribute, "sFieldName", 2, odl_GetFieldName, odl_SetFieldName, VT_BSTR)
DISP_PROPERTY_EX_ID(ExtendedAttribute, "yCFType", 3, odl_GetCFType, odl_SetCFType, VT_I4)
DISP_PROPERTY_EX_ID(ExtendedAttribute, "sGuid", 4, odl_GetGuid, odl_SetGuid, VT_BSTR)
DISP_PROPERTY_EX_ID(ExtendedAttribute, "yElemType", 5, odl_GetElemType, odl_SetElemType, VT_I4)
DISP_PROPERTY_EX_ID(ExtendedAttribute, "lMaxMultiValues", 6, odl_GetMaxMultiValues, odl_SetMaxMultiValues, VT_I4)
DISP_PROPERTY_EX_ID(ExtendedAttribute, "bUserDef", 7, odl_GetUserDef, odl_SetUserDef, VT_BOOL)
DISP_PROPERTY_EX_ID(ExtendedAttribute, "sAlias", 8, odl_GetAlias, odl_SetAlias, VT_BSTR)
DISP_PROPERTY_EX_ID(ExtendedAttribute, "sSecondaryPID", 9, odl_GetSecondaryPID, odl_SetSecondaryPID, VT_BSTR)
DISP_PROPERTY_EX_ID(ExtendedAttribute, "bAutoRollDown", 10, odl_GetAutoRollDown, odl_SetAutoRollDown, VT_BOOL)
DISP_PROPERTY_EX_ID(ExtendedAttribute, "sDefaultGuid", 11, odl_GetDefaultGuid, odl_SetDefaultGuid, VT_BSTR)
DISP_PROPERTY_EX_ID(ExtendedAttribute, "sLtuid", 12, odl_GetLtuid, odl_SetLtuid, VT_BSTR)
DISP_PROPERTY_EX_ID(ExtendedAttribute, "sPhoneticAlias", 13, odl_GetPhoneticAlias, odl_SetPhoneticAlias, VT_BSTR)
DISP_PROPERTY_EX_ID(ExtendedAttribute, "yRollupType", 14, odl_GetRollupType, odl_SetRollupType, VT_I4)
DISP_PROPERTY_EX_ID(ExtendedAttribute, "yCalculationType", 15, odl_GetCalculationType, odl_SetCalculationType, VT_I4)
DISP_PROPERTY_EX_ID(ExtendedAttribute, "sFormula", 16, odl_GetFormula, odl_SetFormula, VT_BSTR)
DISP_PROPERTY_EX_ID(ExtendedAttribute, "bRestrictValues", 17, odl_GetRestrictValues, odl_SetRestrictValues, VT_BOOL)
DISP_PROPERTY_EX_ID(ExtendedAttribute, "yValuelistSortOrder", 18, odl_GetValuelistSortOrder, odl_SetValuelistSortOrder, VT_I4)
DISP_PROPERTY_EX_ID(ExtendedAttribute, "bAppendNewValues", 19, odl_GetAppendNewValues, odl_SetAppendNewValues, VT_BOOL)
DISP_PROPERTY_EX_ID(ExtendedAttribute, "sDefault", 20, odl_GetDefault, odl_SetDefault, VT_BSTR)
DISP_PROPERTY_EX_ID(ExtendedAttribute, "oValueList", 21, odl_GetValueList, SetNotSupported, VT_DISPATCH)
DISP_PROPERTY_EX_ID(ExtendedAttribute, "Key", 22, odl_GetKey, odl_SetKey, VT_BSTR)
DISP_FUNCTION_ID(ExtendedAttribute, "GetXML", 23, odl_GetXML, VT_BSTR, VTS_NONE)
DISP_FUNCTION_ID(ExtendedAttribute, "SetXML", 24, odl_SetXML, VT_EMPTY, VTS_BSTR)
DISP_FUNCTION_ID(ExtendedAttribute, "IsNull", 25, IsNull, VT_BOOL, VTS_NONE)
DISP_FUNCTION_ID(ExtendedAttribute, "Initialize", 26, Initialize, VT_EMPTY, VTS_NONE)
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
	mp_yCFType = CFT_COST;
	mp_sGuid = "";
	mp_yElemType = ET_TASK;
	mp_lMaxMultiValues = 0;
	mp_bUserDef = FALSE;
	mp_sAlias = "";
	mp_sSecondaryPID = "";
	mp_bAutoRollDown = FALSE;
	mp_sDefaultGuid = "";
	mp_sLtuid = "";
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

LONG ExtendedAttribute::odl_GetCFType(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetCFType();
}

E_CFTYPE ExtendedAttribute::GetCFType(void)
{
	return (E_CFTYPE) mp_yCFType;
}

void ExtendedAttribute::odl_SetCFType(LONG newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetCFType((E_CFTYPE)newVal);
}

void ExtendedAttribute::SetCFType(E_CFTYPE newVal)
{
	mp_yCFType = newVal;
}

BSTR ExtendedAttribute::odl_GetGuid(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetGuid().AllocSysString();
}

CString ExtendedAttribute::GetGuid(void)
{
	return (CString) mp_sGuid;
}

void ExtendedAttribute::odl_SetGuid(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetGuid(newVal);
}

void ExtendedAttribute::SetGuid(CString newVal)
{
	mp_sGuid = newVal;
}

LONG ExtendedAttribute::odl_GetElemType(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetElemType();
}

E_ELEMTYPE ExtendedAttribute::GetElemType(void)
{
	return (E_ELEMTYPE) mp_yElemType;
}

void ExtendedAttribute::odl_SetElemType(LONG newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetElemType((E_ELEMTYPE)newVal);
}

void ExtendedAttribute::SetElemType(E_ELEMTYPE newVal)
{
	mp_yElemType = newVal;
}

LONG ExtendedAttribute::odl_GetMaxMultiValues(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetMaxMultiValues();
}

LONG ExtendedAttribute::GetMaxMultiValues(void)
{
	return (LONG) mp_lMaxMultiValues;
}

void ExtendedAttribute::odl_SetMaxMultiValues(LONG newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetMaxMultiValues((LONG)newVal);
}

void ExtendedAttribute::SetMaxMultiValues(LONG newVal)
{
	mp_lMaxMultiValues = newVal;
}

BOOL ExtendedAttribute::odl_GetUserDef(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetUserDef();
}

BOOL ExtendedAttribute::GetUserDef(void)
{
	return (BOOL) mp_bUserDef;
}

void ExtendedAttribute::odl_SetUserDef(BOOL newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetUserDef((BOOL)newVal);
}

void ExtendedAttribute::SetUserDef(BOOL newVal)
{
	mp_bUserDef = newVal;
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

BSTR ExtendedAttribute::odl_GetSecondaryPID(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetSecondaryPID().AllocSysString();
}

CString ExtendedAttribute::GetSecondaryPID(void)
{
	return (CString) mp_sSecondaryPID;
}

void ExtendedAttribute::odl_SetSecondaryPID(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetSecondaryPID(newVal);
}

void ExtendedAttribute::SetSecondaryPID(CString newVal)
{
	mp_sSecondaryPID = newVal;
}

BOOL ExtendedAttribute::odl_GetAutoRollDown(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetAutoRollDown();
}

BOOL ExtendedAttribute::GetAutoRollDown(void)
{
	return (BOOL) mp_bAutoRollDown;
}

void ExtendedAttribute::odl_SetAutoRollDown(BOOL newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetAutoRollDown((BOOL)newVal);
}

void ExtendedAttribute::SetAutoRollDown(BOOL newVal)
{
	mp_bAutoRollDown = newVal;
}

BSTR ExtendedAttribute::odl_GetDefaultGuid(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetDefaultGuid().AllocSysString();
}

CString ExtendedAttribute::GetDefaultGuid(void)
{
	return (CString) mp_sDefaultGuid;
}

void ExtendedAttribute::odl_SetDefaultGuid(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetDefaultGuid(newVal);
}

void ExtendedAttribute::SetDefaultGuid(CString newVal)
{
	mp_sDefaultGuid = newVal;
}

BSTR ExtendedAttribute::odl_GetLtuid(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetLtuid().AllocSysString();
}

CString ExtendedAttribute::GetLtuid(void)
{
	return (CString) mp_sLtuid;
}

void ExtendedAttribute::odl_SetLtuid(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetLtuid(newVal);
}

void ExtendedAttribute::SetLtuid(CString newVal)
{
	mp_sLtuid = newVal;
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
	if (mp_yCFType != CFT_COST)
	{
		bReturn = FALSE;
	}
	if (mp_sGuid != "")
	{
		bReturn = FALSE;
	}
	if (mp_yElemType != ET_TASK)
	{
		bReturn = FALSE;
	}
	if (mp_lMaxMultiValues != 0)
	{
		bReturn = FALSE;
	}
	if (mp_bUserDef != FALSE)
	{
		bReturn = FALSE;
	}
	if (mp_sAlias != "")
	{
		bReturn = FALSE;
	}
	if (mp_sSecondaryPID != "")
	{
		bReturn = FALSE;
	}
	if (mp_bAutoRollDown != FALSE)
	{
		bReturn = FALSE;
	}
	if (mp_sDefaultGuid != "")
	{
		bReturn = FALSE;
	}
	if (mp_sLtuid != "")
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
	oXML.WriteProperty("CFType", mp_yCFType);
	if (mp_sGuid != "")
	{
		oXML.WriteProperty("Guid", mp_sGuid);
	}
	oXML.WriteProperty("ElemType", mp_yElemType);
	oXML.WriteProperty("MaxMultiValues", mp_lMaxMultiValues);
	oXML.WriteProperty("UserDef", mp_bUserDef);
	if (mp_sAlias != "")
	{
		oXML.WriteProperty("Alias", mp_sAlias);
	}
	if (mp_sSecondaryPID != "")
	{
		oXML.WriteProperty("SecondaryPID", mp_sSecondaryPID);
	}
	oXML.WriteProperty("AutoRollDown", mp_bAutoRollDown);
	if (mp_sDefaultGuid != "")
	{
		oXML.WriteProperty("DefaultGuid", mp_sDefaultGuid);
	}
	if (mp_sLtuid != "")
	{
		oXML.WriteProperty("Ltuid", mp_sLtuid);
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
	oXML.ReadProperty("CFType", mp_yCFType);
	oXML.ReadProperty("Guid", mp_sGuid);
	oXML.ReadProperty("ElemType", mp_yElemType);
	oXML.ReadProperty("MaxMultiValues", mp_lMaxMultiValues);
	oXML.ReadProperty("UserDef", mp_bUserDef);
	oXML.ReadProperty("Alias", mp_sAlias);
	if (mp_sAlias.GetLength() > 50)
	{
		mp_sAlias = mp_sAlias.Mid(0, 50);
	}
	oXML.ReadProperty("SecondaryPID", mp_sSecondaryPID);
	oXML.ReadProperty("AutoRollDown", mp_bAutoRollDown);
	oXML.ReadProperty("DefaultGuid", mp_sDefaultGuid);
	oXML.ReadProperty("Ltuid", mp_sLtuid);
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