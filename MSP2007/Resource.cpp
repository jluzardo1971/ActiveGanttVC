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
#include "Resource.h"

IMPLEMENT_DYNCREATE(Resource, CCmdTarget)

//{29CDCFAF-3A4D-48C0-BD06-6BC92BD541C0}
static const IID IID_IResource = { 0x29CDCFAF, 0x3A4D, 0x48C0, { 0xBD, 0x06, 0x6B, 0xC9, 0x2B, 0xD5, 0x41, 0xC0} };

//{F0E44B46-73B6-4698-AE2A-5F2E83DEB8B0}
IMPLEMENT_OLECREATE_FLAGS(Resource, "MSP2007.Resource", afxRegApartmentThreading, 0xF0E44B46, 0x73B6, 0x4698, 0xAE, 0x2A, 0x5F, 0x2E, 0x83, 0xDE, 0xB8, 0xB0)

BEGIN_DISPATCH_MAP(Resource, CCmdTarget)
DISP_PROPERTY_EX_ID(Resource, "lUID", 1, odl_GetUID, odl_SetUID, VT_I4)
DISP_PROPERTY_EX_ID(Resource, "lID", 2, odl_GetID, odl_SetID, VT_I4)
DISP_PROPERTY_EX_ID(Resource, "sName", 3, odl_GetName, odl_SetName, VT_BSTR)
DISP_PROPERTY_EX_ID(Resource, "yType", 4, odl_GetType, odl_SetType, VT_I4)
DISP_PROPERTY_EX_ID(Resource, "bIsNull", 5, odl_GetIsNull, odl_SetIsNull, VT_BOOL)
DISP_PROPERTY_EX_ID(Resource, "sInitials", 6, odl_GetInitials, odl_SetInitials, VT_BSTR)
DISP_PROPERTY_EX_ID(Resource, "sPhonetics", 7, odl_GetPhonetics, odl_SetPhonetics, VT_BSTR)
DISP_PROPERTY_EX_ID(Resource, "sNTAccount", 8, odl_GetNTAccount, odl_SetNTAccount, VT_BSTR)
DISP_PROPERTY_EX_ID(Resource, "sMaterialLabel", 9, odl_GetMaterialLabel, odl_SetMaterialLabel, VT_BSTR)
DISP_PROPERTY_EX_ID(Resource, "sCode", 10, odl_GetCode, odl_SetCode, VT_BSTR)
DISP_PROPERTY_EX_ID(Resource, "sGroup", 11, odl_GetGroup, odl_SetGroup, VT_BSTR)
DISP_PROPERTY_EX_ID(Resource, "yWorkGroup", 12, odl_GetWorkGroup, odl_SetWorkGroup, VT_I4)
DISP_PROPERTY_EX_ID(Resource, "sEmailAddress", 13, odl_GetEmailAddress, odl_SetEmailAddress, VT_BSTR)
DISP_PROPERTY_EX_ID(Resource, "sHyperlink", 14, odl_GetHyperlink, odl_SetHyperlink, VT_BSTR)
DISP_PROPERTY_EX_ID(Resource, "sHyperlinkAddress", 15, odl_GetHyperlinkAddress, odl_SetHyperlinkAddress, VT_BSTR)
DISP_PROPERTY_EX_ID(Resource, "sHyperlinkSubAddress", 16, odl_GetHyperlinkSubAddress, odl_SetHyperlinkSubAddress, VT_BSTR)
DISP_PROPERTY_EX_ID(Resource, "fMaxUnits", 17, odl_GetMaxUnits, odl_SetMaxUnits, VT_R4)
DISP_PROPERTY_EX_ID(Resource, "fPeakUnits", 18, odl_GetPeakUnits, odl_SetPeakUnits, VT_R4)
DISP_PROPERTY_EX_ID(Resource, "bOverAllocated", 19, odl_GetOverAllocated, odl_SetOverAllocated, VT_BOOL)
DISP_PROPERTY_EX_ID(Resource, "dtAvailableFrom", 20, odl_GetAvailableFrom, odl_SetAvailableFrom, VT_DATE)
DISP_PROPERTY_EX_ID(Resource, "dtAvailableTo", 21, odl_GetAvailableTo, odl_SetAvailableTo, VT_DATE)
DISP_PROPERTY_EX_ID(Resource, "dtStart", 22, odl_GetStart, odl_SetStart, VT_DATE)
DISP_PROPERTY_EX_ID(Resource, "dtFinish", 23, odl_GetFinish, odl_SetFinish, VT_DATE)
DISP_PROPERTY_EX_ID(Resource, "bCanLevel", 24, odl_GetCanLevel, odl_SetCanLevel, VT_BOOL)
DISP_PROPERTY_EX_ID(Resource, "yAccrueAt", 25, odl_GetAccrueAt, odl_SetAccrueAt, VT_I4)
DISP_PROPERTY_EX_ID(Resource, "oWork", 26, odl_GetWork, SetNotSupported, VT_DISPATCH)
DISP_PROPERTY_EX_ID(Resource, "oRegularWork", 27, odl_GetRegularWork, SetNotSupported, VT_DISPATCH)
DISP_PROPERTY_EX_ID(Resource, "oOvertimeWork", 28, odl_GetOvertimeWork, SetNotSupported, VT_DISPATCH)
DISP_PROPERTY_EX_ID(Resource, "oActualWork", 29, odl_GetActualWork, SetNotSupported, VT_DISPATCH)
DISP_PROPERTY_EX_ID(Resource, "oRemainingWork", 30, odl_GetRemainingWork, SetNotSupported, VT_DISPATCH)
DISP_PROPERTY_EX_ID(Resource, "oActualOvertimeWork", 31, odl_GetActualOvertimeWork, SetNotSupported, VT_DISPATCH)
DISP_PROPERTY_EX_ID(Resource, "oRemainingOvertimeWork", 32, odl_GetRemainingOvertimeWork, SetNotSupported, VT_DISPATCH)
DISP_PROPERTY_EX_ID(Resource, "lPercentWorkComplete", 33, odl_GetPercentWorkComplete, odl_SetPercentWorkComplete, VT_I4)
DISP_PROPERTY_EX_ID(Resource, "cStandardRate", 34, odl_GetStandardRate, odl_SetStandardRate, VT_R8)
DISP_PROPERTY_EX_ID(Resource, "yStandardRateFormat", 35, odl_GetStandardRateFormat, odl_SetStandardRateFormat, VT_I4)
DISP_PROPERTY_EX_ID(Resource, "cCost", 36, odl_GetCost, odl_SetCost, VT_R8)
DISP_PROPERTY_EX_ID(Resource, "cOvertimeRate", 37, odl_GetOvertimeRate, odl_SetOvertimeRate, VT_R8)
DISP_PROPERTY_EX_ID(Resource, "yOvertimeRateFormat", 38, odl_GetOvertimeRateFormat, odl_SetOvertimeRateFormat, VT_I4)
DISP_PROPERTY_EX_ID(Resource, "cOvertimeCost", 39, odl_GetOvertimeCost, odl_SetOvertimeCost, VT_R8)
DISP_PROPERTY_EX_ID(Resource, "cCostPerUse", 40, odl_GetCostPerUse, odl_SetCostPerUse, VT_R8)
DISP_PROPERTY_EX_ID(Resource, "cActualCost", 41, odl_GetActualCost, odl_SetActualCost, VT_R8)
DISP_PROPERTY_EX_ID(Resource, "cActualOvertimeCost", 42, odl_GetActualOvertimeCost, odl_SetActualOvertimeCost, VT_R8)
DISP_PROPERTY_EX_ID(Resource, "cRemainingCost", 43, odl_GetRemainingCost, odl_SetRemainingCost, VT_R8)
DISP_PROPERTY_EX_ID(Resource, "cRemainingOvertimeCost", 44, odl_GetRemainingOvertimeCost, odl_SetRemainingOvertimeCost, VT_R8)
DISP_PROPERTY_EX_ID(Resource, "fWorkVariance", 45, odl_GetWorkVariance, odl_SetWorkVariance, VT_R4)
DISP_PROPERTY_EX_ID(Resource, "fCostVariance", 46, odl_GetCostVariance, odl_SetCostVariance, VT_R4)
DISP_PROPERTY_EX_ID(Resource, "fSV", 47, odl_GetSV, odl_SetSV, VT_R4)
DISP_PROPERTY_EX_ID(Resource, "fCV", 48, odl_GetCV, odl_SetCV, VT_R4)
DISP_PROPERTY_EX_ID(Resource, "fACWP", 49, odl_GetACWP, odl_SetACWP, VT_R4)
DISP_PROPERTY_EX_ID(Resource, "lCalendarUID", 50, odl_GetCalendarUID, odl_SetCalendarUID, VT_I4)
DISP_PROPERTY_EX_ID(Resource, "sNotes", 51, odl_GetNotes, odl_SetNotes, VT_BSTR)
DISP_PROPERTY_EX_ID(Resource, "fBCWS", 52, odl_GetBCWS, odl_SetBCWS, VT_R4)
DISP_PROPERTY_EX_ID(Resource, "fBCWP", 53, odl_GetBCWP, odl_SetBCWP, VT_R4)
DISP_PROPERTY_EX_ID(Resource, "bIsGeneric", 54, odl_GetIsGeneric, odl_SetIsGeneric, VT_BOOL)
DISP_PROPERTY_EX_ID(Resource, "bIsInactive", 55, odl_GetIsInactive, odl_SetIsInactive, VT_BOOL)
DISP_PROPERTY_EX_ID(Resource, "bIsEnterprise", 56, odl_GetIsEnterprise, odl_SetIsEnterprise, VT_BOOL)
DISP_PROPERTY_EX_ID(Resource, "yBookingType", 57, odl_GetBookingType, odl_SetBookingType, VT_I4)
DISP_PROPERTY_EX_ID(Resource, "oActualWorkProtected", 58, odl_GetActualWorkProtected, SetNotSupported, VT_DISPATCH)
DISP_PROPERTY_EX_ID(Resource, "oActualOvertimeWorkProtected", 59, odl_GetActualOvertimeWorkProtected, SetNotSupported, VT_DISPATCH)
DISP_PROPERTY_EX_ID(Resource, "sActiveDirectoryGUID", 60, odl_GetActiveDirectoryGUID, odl_SetActiveDirectoryGUID, VT_BSTR)
DISP_PROPERTY_EX_ID(Resource, "dtCreationDate", 61, odl_GetCreationDate, odl_SetCreationDate, VT_DATE)
DISP_PROPERTY_EX_ID(Resource, "oExtendedAttribute", 62, odl_GetExtendedAttribute, SetNotSupported, VT_DISPATCH)
DISP_PROPERTY_EX_ID(Resource, "oBaseline", 63, odl_GetBaseline, SetNotSupported, VT_DISPATCH)
DISP_PROPERTY_EX_ID(Resource, "oOutlineCode", 64, odl_GetOutlineCode, SetNotSupported, VT_DISPATCH)
DISP_PROPERTY_EX_ID(Resource, "bIsCostResource", 65, odl_GetIsCostResource, odl_SetIsCostResource, VT_BOOL)
DISP_PROPERTY_EX_ID(Resource, "sAssnOwner", 66, odl_GetAssnOwner, odl_SetAssnOwner, VT_BSTR)
DISP_PROPERTY_EX_ID(Resource, "sAssnOwnerGuid", 67, odl_GetAssnOwnerGuid, odl_SetAssnOwnerGuid, VT_BSTR)
DISP_PROPERTY_EX_ID(Resource, "bIsBudget", 68, odl_GetIsBudget, odl_SetIsBudget, VT_BOOL)
DISP_PROPERTY_EX_ID(Resource, "oAvailabilityPeriods", 69, odl_GetAvailabilityPeriods, SetNotSupported, VT_DISPATCH)
DISP_PROPERTY_EX_ID(Resource, "oRates", 70, odl_GetRates, SetNotSupported, VT_DISPATCH)
DISP_PROPERTY_EX_ID(Resource, "oTimephasedData", 71, odl_GetTimephasedData, SetNotSupported, VT_DISPATCH)
DISP_PROPERTY_EX_ID(Resource, "Key", 72, odl_GetKey, odl_SetKey, VT_BSTR)
DISP_FUNCTION_ID(Resource, "GetXML", 73, odl_GetXML, VT_BSTR, VTS_NONE)
DISP_FUNCTION_ID(Resource, "SetXML", 74, odl_SetXML, VT_EMPTY, VTS_BSTR)
DISP_FUNCTION_ID(Resource, "IsNull", 75, IsNull, VT_BOOL, VTS_NONE)
DISP_FUNCTION_ID(Resource, "Initialize", 76, Initialize, VT_EMPTY, VTS_NONE)
END_DISPATCH_MAP()

BEGIN_INTERFACE_MAP(Resource, CCmdTarget)
	INTERFACE_PART(Resource, IID_IResource, Dispatch)
END_INTERFACE_MAP()

BEGIN_MESSAGE_MAP(Resource, CCmdTarget)
END_MESSAGE_MAP()


void Resource::Initialize(void)
{
	InitVars();
}

Resource::Resource()
{
	EnableAutomation();
	AfxOleLockApp();
	InitVars();
}

void Resource::InitVars(void)
{
	mp_lUID = 0;
	mp_lID = 0;
	mp_sName = "";
	mp_yType = T_6_MATERIAL;
	mp_bIsNull = FALSE;
	mp_sInitials = "";
	mp_sPhonetics = "";
	mp_sNTAccount = "";
	mp_sMaterialLabel = "";
	mp_sCode = "";
	mp_sGroup = "";
	mp_yWorkGroup = WG_DEFAULT;
	mp_sEmailAddress = "";
	mp_sHyperlink = "";
	mp_sHyperlinkAddress = "";
	mp_sHyperlinkSubAddress = "";
	mp_fMaxUnits = 1.0;
	mp_fPeakUnits = 0;
	mp_bOverAllocated = FALSE;
	mp_dtAvailableFrom = (DATE) 0;
	mp_dtAvailableTo = (DATE) 0;
	mp_dtStart = (DATE) 0;
	mp_dtFinish = (DATE) 0;
	mp_bCanLevel = FALSE;
	mp_yAccrueAt = AA_START;
	mp_oWork = new Duration();
	mp_oRegularWork = new Duration();
	mp_oOvertimeWork = new Duration();
	mp_oActualWork = new Duration();
	mp_oRemainingWork = new Duration();
	mp_oActualOvertimeWork = new Duration();
	mp_oRemainingOvertimeWork = new Duration();
	mp_lPercentWorkComplete = 0;
	mp_cStandardRate = 0;
	mp_yStandardRateFormat = SRF_M;
	mp_cCost = 0;
	mp_cOvertimeRate = 0;
	mp_yOvertimeRateFormat = ORF_M;
	mp_cOvertimeCost = 0;
	mp_cCostPerUse = 0;
	mp_cActualCost = 0;
	mp_cActualOvertimeCost = 0;
	mp_cRemainingCost = 0;
	mp_cRemainingOvertimeCost = 0;
	mp_fWorkVariance = 0;
	mp_fCostVariance = 0;
	mp_fSV = 0;
	mp_fCV = 0;
	mp_fACWP = 0;
	mp_lCalendarUID = 0;
	mp_sNotes = "";
	mp_fBCWS = 0;
	mp_fBCWP = 0;
	mp_bIsGeneric = FALSE;
	mp_bIsInactive = FALSE;
	mp_bIsEnterprise = FALSE;
	mp_yBookingType = BT_COMMITED;
	mp_oActualWorkProtected = new Duration();
	mp_oActualOvertimeWorkProtected = new Duration();
	mp_sActiveDirectoryGUID = "";
	mp_dtCreationDate = (DATE) 0;
	mp_oExtendedAttribute_C = new ResourceExtendedAttribute_C();
	mp_oBaseline_C = new ResourceBaseline_C();
	mp_oOutlineCode_C = new ResourceOutlineCode_C();
	mp_bIsCostResource = FALSE;
	mp_sAssnOwner = "";
	mp_sAssnOwnerGuid = "";
	mp_bIsBudget = FALSE;
	mp_oAvailabilityPeriods = new ResourceAvailabilityPeriods();
	mp_oRates = new ResourceRates();
	mp_oTimephasedData_C = new TimephasedData_C();
}

Resource::~Resource()
{
	delete mp_oWork;
	delete mp_oRegularWork;
	delete mp_oOvertimeWork;
	delete mp_oActualWork;
	delete mp_oRemainingWork;
	delete mp_oActualOvertimeWork;
	delete mp_oRemainingOvertimeWork;
	delete mp_oActualWorkProtected;
	delete mp_oActualOvertimeWorkProtected;
	delete mp_oExtendedAttribute_C;
	delete mp_oBaseline_C;
	delete mp_oOutlineCode_C;
	delete mp_oAvailabilityPeriods;
	delete mp_oRates;
	delete mp_oTimephasedData_C;
	AfxOleUnlockApp();
}

void Resource::OnFinalRelease()
{
	clsItemBase::OnFinalRelease();
}

LONG Resource::odl_GetUID(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetUID();
}

LONG Resource::GetUID(void)
{
	return (LONG) mp_lUID;
}

void Resource::odl_SetUID(LONG newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetUID((LONG)newVal);
}

void Resource::SetUID(LONG newVal)
{
	mp_lUID = newVal;
}

LONG Resource::odl_GetID(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetID();
}

LONG Resource::GetID(void)
{
	return (LONG) mp_lID;
}

void Resource::odl_SetID(LONG newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetID((LONG)newVal);
}

void Resource::SetID(LONG newVal)
{
	mp_lID = newVal;
}

BSTR Resource::odl_GetName(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetName().AllocSysString();
}

CString Resource::GetName(void)
{
	return (CString) mp_sName;
}

void Resource::odl_SetName(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetName(newVal);
}

void Resource::SetName(CString newVal)
{
	if (newVal.GetLength() > 512)
	{
		newVal = newVal.Mid(0, 512);
	}
	mp_sName = newVal;
}

LONG Resource::odl_GetType(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetType();
}

E_TYPE_6 Resource::GetType(void)
{
	return (E_TYPE_6) mp_yType;
}

void Resource::odl_SetType(LONG newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetType((E_TYPE_6)newVal);
}

void Resource::SetType(E_TYPE_6 newVal)
{
	mp_yType = newVal;
}

BOOL Resource::odl_GetIsNull(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetIsNull();
}

BOOL Resource::GetIsNull(void)
{
	return (BOOL) mp_bIsNull;
}

void Resource::odl_SetIsNull(BOOL newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetIsNull((BOOL)newVal);
}

void Resource::SetIsNull(BOOL newVal)
{
	mp_bIsNull = newVal;
}

BSTR Resource::odl_GetInitials(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetInitials().AllocSysString();
}

CString Resource::GetInitials(void)
{
	return (CString) mp_sInitials;
}

void Resource::odl_SetInitials(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetInitials(newVal);
}

void Resource::SetInitials(CString newVal)
{
	if (newVal.GetLength() > 512)
	{
		newVal = newVal.Mid(0, 512);
	}
	mp_sInitials = newVal;
}

BSTR Resource::odl_GetPhonetics(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetPhonetics().AllocSysString();
}

CString Resource::GetPhonetics(void)
{
	return (CString) mp_sPhonetics;
}

void Resource::odl_SetPhonetics(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetPhonetics(newVal);
}

void Resource::SetPhonetics(CString newVal)
{
	if (newVal.GetLength() > 512)
	{
		newVal = newVal.Mid(0, 512);
	}
	mp_sPhonetics = newVal;
}

BSTR Resource::odl_GetNTAccount(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetNTAccount().AllocSysString();
}

CString Resource::GetNTAccount(void)
{
	return (CString) mp_sNTAccount;
}

void Resource::odl_SetNTAccount(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetNTAccount(newVal);
}

void Resource::SetNTAccount(CString newVal)
{
	if (newVal.GetLength() > 512)
	{
		newVal = newVal.Mid(0, 512);
	}
	mp_sNTAccount = newVal;
}

BSTR Resource::odl_GetMaterialLabel(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetMaterialLabel().AllocSysString();
}

CString Resource::GetMaterialLabel(void)
{
	return (CString) mp_sMaterialLabel;
}

void Resource::odl_SetMaterialLabel(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetMaterialLabel(newVal);
}

void Resource::SetMaterialLabel(CString newVal)
{
	if (newVal.GetLength() > 512)
	{
		newVal = newVal.Mid(0, 512);
	}
	mp_sMaterialLabel = newVal;
}

BSTR Resource::odl_GetCode(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetCode().AllocSysString();
}

CString Resource::GetCode(void)
{
	return (CString) mp_sCode;
}

void Resource::odl_SetCode(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetCode(newVal);
}

void Resource::SetCode(CString newVal)
{
	if (newVal.GetLength() > 512)
	{
		newVal = newVal.Mid(0, 512);
	}
	mp_sCode = newVal;
}

BSTR Resource::odl_GetGroup(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetGroup().AllocSysString();
}

CString Resource::GetGroup(void)
{
	return (CString) mp_sGroup;
}

void Resource::odl_SetGroup(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetGroup(newVal);
}

void Resource::SetGroup(CString newVal)
{
	if (newVal.GetLength() > 512)
	{
		newVal = newVal.Mid(0, 512);
	}
	mp_sGroup = newVal;
}

LONG Resource::odl_GetWorkGroup(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetWorkGroup();
}

E_WORKGROUP Resource::GetWorkGroup(void)
{
	return (E_WORKGROUP) mp_yWorkGroup;
}

void Resource::odl_SetWorkGroup(LONG newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetWorkGroup((E_WORKGROUP)newVal);
}

void Resource::SetWorkGroup(E_WORKGROUP newVal)
{
	mp_yWorkGroup = newVal;
}

BSTR Resource::odl_GetEmailAddress(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetEmailAddress().AllocSysString();
}

CString Resource::GetEmailAddress(void)
{
	return (CString) mp_sEmailAddress;
}

void Resource::odl_SetEmailAddress(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetEmailAddress(newVal);
}

void Resource::SetEmailAddress(CString newVal)
{
	if (newVal.GetLength() > 512)
	{
		newVal = newVal.Mid(0, 512);
	}
	mp_sEmailAddress = newVal;
}

BSTR Resource::odl_GetHyperlink(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetHyperlink().AllocSysString();
}

CString Resource::GetHyperlink(void)
{
	return (CString) mp_sHyperlink;
}

void Resource::odl_SetHyperlink(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetHyperlink(newVal);
}

void Resource::SetHyperlink(CString newVal)
{
	if (newVal.GetLength() > 512)
	{
		newVal = newVal.Mid(0, 512);
	}
	mp_sHyperlink = newVal;
}

BSTR Resource::odl_GetHyperlinkAddress(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetHyperlinkAddress().AllocSysString();
}

CString Resource::GetHyperlinkAddress(void)
{
	return (CString) mp_sHyperlinkAddress;
}

void Resource::odl_SetHyperlinkAddress(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetHyperlinkAddress(newVal);
}

void Resource::SetHyperlinkAddress(CString newVal)
{
	if (newVal.GetLength() > 512)
	{
		newVal = newVal.Mid(0, 512);
	}
	mp_sHyperlinkAddress = newVal;
}

BSTR Resource::odl_GetHyperlinkSubAddress(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetHyperlinkSubAddress().AllocSysString();
}

CString Resource::GetHyperlinkSubAddress(void)
{
	return (CString) mp_sHyperlinkSubAddress;
}

void Resource::odl_SetHyperlinkSubAddress(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetHyperlinkSubAddress(newVal);
}

void Resource::SetHyperlinkSubAddress(CString newVal)
{
	if (newVal.GetLength() > 512)
	{
		newVal = newVal.Mid(0, 512);
	}
	mp_sHyperlinkSubAddress = newVal;
}

FLOAT Resource::odl_GetMaxUnits(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetMaxUnits();
}

FLOAT Resource::GetMaxUnits(void)
{
	return (FLOAT) mp_fMaxUnits;
}

void Resource::odl_SetMaxUnits(FLOAT newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetMaxUnits((FLOAT)newVal);
}

void Resource::SetMaxUnits(FLOAT newVal)
{
	mp_fMaxUnits = newVal;
}

FLOAT Resource::odl_GetPeakUnits(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetPeakUnits();
}

FLOAT Resource::GetPeakUnits(void)
{
	return (FLOAT) mp_fPeakUnits;
}

void Resource::odl_SetPeakUnits(FLOAT newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetPeakUnits((FLOAT)newVal);
}

void Resource::SetPeakUnits(FLOAT newVal)
{
	mp_fPeakUnits = newVal;
}

BOOL Resource::odl_GetOverAllocated(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetOverAllocated();
}

BOOL Resource::GetOverAllocated(void)
{
	return (BOOL) mp_bOverAllocated;
}

void Resource::odl_SetOverAllocated(BOOL newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetOverAllocated((BOOL)newVal);
}

void Resource::SetOverAllocated(BOOL newVal)
{
	mp_bOverAllocated = newVal;
}

DATE Resource::odl_GetAvailableFrom(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetAvailableFrom();
}

COleDateTime Resource::GetAvailableFrom(void)
{
	return (COleDateTime) mp_dtAvailableFrom;
}

void Resource::odl_SetAvailableFrom(DATE newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetAvailableFrom((COleDateTime)newVal);
}

void Resource::SetAvailableFrom(COleDateTime newVal)
{
	mp_dtAvailableFrom = newVal;
}

DATE Resource::odl_GetAvailableTo(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetAvailableTo();
}

COleDateTime Resource::GetAvailableTo(void)
{
	return (COleDateTime) mp_dtAvailableTo;
}

void Resource::odl_SetAvailableTo(DATE newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetAvailableTo((COleDateTime)newVal);
}

void Resource::SetAvailableTo(COleDateTime newVal)
{
	mp_dtAvailableTo = newVal;
}

DATE Resource::odl_GetStart(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetStart();
}

COleDateTime Resource::GetStart(void)
{
	return (COleDateTime) mp_dtStart;
}

void Resource::odl_SetStart(DATE newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetStart((COleDateTime)newVal);
}

void Resource::SetStart(COleDateTime newVal)
{
	mp_dtStart = newVal;
}

DATE Resource::odl_GetFinish(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetFinish();
}

COleDateTime Resource::GetFinish(void)
{
	return (COleDateTime) mp_dtFinish;
}

void Resource::odl_SetFinish(DATE newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetFinish((COleDateTime)newVal);
}

void Resource::SetFinish(COleDateTime newVal)
{
	mp_dtFinish = newVal;
}

BOOL Resource::odl_GetCanLevel(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetCanLevel();
}

BOOL Resource::GetCanLevel(void)
{
	return (BOOL) mp_bCanLevel;
}

void Resource::odl_SetCanLevel(BOOL newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetCanLevel((BOOL)newVal);
}

void Resource::SetCanLevel(BOOL newVal)
{
	mp_bCanLevel = newVal;
}

LONG Resource::odl_GetAccrueAt(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetAccrueAt();
}

E_ACCRUEAT Resource::GetAccrueAt(void)
{
	return (E_ACCRUEAT) mp_yAccrueAt;
}

void Resource::odl_SetAccrueAt(LONG newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetAccrueAt((E_ACCRUEAT)newVal);
}

void Resource::SetAccrueAt(E_ACCRUEAT newVal)
{
	mp_yAccrueAt = newVal;
}

IDispatch* Resource::odl_GetWork(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetWork()->GetIDispatch(TRUE);
}

Duration* Resource::GetWork(void)
{
	return mp_oWork;
}

IDispatch* Resource::odl_GetRegularWork(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetRegularWork()->GetIDispatch(TRUE);
}

Duration* Resource::GetRegularWork(void)
{
	return mp_oRegularWork;
}

IDispatch* Resource::odl_GetOvertimeWork(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetOvertimeWork()->GetIDispatch(TRUE);
}

Duration* Resource::GetOvertimeWork(void)
{
	return mp_oOvertimeWork;
}

IDispatch* Resource::odl_GetActualWork(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetActualWork()->GetIDispatch(TRUE);
}

Duration* Resource::GetActualWork(void)
{
	return mp_oActualWork;
}

IDispatch* Resource::odl_GetRemainingWork(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetRemainingWork()->GetIDispatch(TRUE);
}

Duration* Resource::GetRemainingWork(void)
{
	return mp_oRemainingWork;
}

IDispatch* Resource::odl_GetActualOvertimeWork(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetActualOvertimeWork()->GetIDispatch(TRUE);
}

Duration* Resource::GetActualOvertimeWork(void)
{
	return mp_oActualOvertimeWork;
}

IDispatch* Resource::odl_GetRemainingOvertimeWork(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetRemainingOvertimeWork()->GetIDispatch(TRUE);
}

Duration* Resource::GetRemainingOvertimeWork(void)
{
	return mp_oRemainingOvertimeWork;
}

LONG Resource::odl_GetPercentWorkComplete(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetPercentWorkComplete();
}

LONG Resource::GetPercentWorkComplete(void)
{
	return (LONG) mp_lPercentWorkComplete;
}

void Resource::odl_SetPercentWorkComplete(LONG newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetPercentWorkComplete((LONG)newVal);
}

void Resource::SetPercentWorkComplete(LONG newVal)
{
	mp_lPercentWorkComplete = newVal;
}

DOUBLE Resource::odl_GetStandardRate(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetStandardRate();
}

DOUBLE Resource::GetStandardRate(void)
{
	return (DOUBLE) mp_cStandardRate;
}

void Resource::odl_SetStandardRate(DOUBLE newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetStandardRate((DOUBLE)newVal);
}

void Resource::SetStandardRate(DOUBLE newVal)
{
	mp_cStandardRate = newVal;
}

LONG Resource::odl_GetStandardRateFormat(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetStandardRateFormat();
}

E_STANDARDRATEFORMAT Resource::GetStandardRateFormat(void)
{
	return (E_STANDARDRATEFORMAT) mp_yStandardRateFormat;
}

void Resource::odl_SetStandardRateFormat(LONG newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetStandardRateFormat((E_STANDARDRATEFORMAT)newVal);
}

void Resource::SetStandardRateFormat(E_STANDARDRATEFORMAT newVal)
{
	mp_yStandardRateFormat = newVal;
}

DOUBLE Resource::odl_GetCost(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetCost();
}

DOUBLE Resource::GetCost(void)
{
	return (DOUBLE) mp_cCost;
}

void Resource::odl_SetCost(DOUBLE newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetCost((DOUBLE)newVal);
}

void Resource::SetCost(DOUBLE newVal)
{
	mp_cCost = newVal;
}

DOUBLE Resource::odl_GetOvertimeRate(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetOvertimeRate();
}

DOUBLE Resource::GetOvertimeRate(void)
{
	return (DOUBLE) mp_cOvertimeRate;
}

void Resource::odl_SetOvertimeRate(DOUBLE newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetOvertimeRate((DOUBLE)newVal);
}

void Resource::SetOvertimeRate(DOUBLE newVal)
{
	mp_cOvertimeRate = newVal;
}

LONG Resource::odl_GetOvertimeRateFormat(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetOvertimeRateFormat();
}

E_OVERTIMERATEFORMAT Resource::GetOvertimeRateFormat(void)
{
	return (E_OVERTIMERATEFORMAT) mp_yOvertimeRateFormat;
}

void Resource::odl_SetOvertimeRateFormat(LONG newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetOvertimeRateFormat((E_OVERTIMERATEFORMAT)newVal);
}

void Resource::SetOvertimeRateFormat(E_OVERTIMERATEFORMAT newVal)
{
	mp_yOvertimeRateFormat = newVal;
}

DOUBLE Resource::odl_GetOvertimeCost(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetOvertimeCost();
}

DOUBLE Resource::GetOvertimeCost(void)
{
	return (DOUBLE) mp_cOvertimeCost;
}

void Resource::odl_SetOvertimeCost(DOUBLE newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetOvertimeCost((DOUBLE)newVal);
}

void Resource::SetOvertimeCost(DOUBLE newVal)
{
	mp_cOvertimeCost = newVal;
}

DOUBLE Resource::odl_GetCostPerUse(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetCostPerUse();
}

DOUBLE Resource::GetCostPerUse(void)
{
	return (DOUBLE) mp_cCostPerUse;
}

void Resource::odl_SetCostPerUse(DOUBLE newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetCostPerUse((DOUBLE)newVal);
}

void Resource::SetCostPerUse(DOUBLE newVal)
{
	mp_cCostPerUse = newVal;
}

DOUBLE Resource::odl_GetActualCost(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetActualCost();
}

DOUBLE Resource::GetActualCost(void)
{
	return (DOUBLE) mp_cActualCost;
}

void Resource::odl_SetActualCost(DOUBLE newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetActualCost((DOUBLE)newVal);
}

void Resource::SetActualCost(DOUBLE newVal)
{
	mp_cActualCost = newVal;
}

DOUBLE Resource::odl_GetActualOvertimeCost(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetActualOvertimeCost();
}

DOUBLE Resource::GetActualOvertimeCost(void)
{
	return (DOUBLE) mp_cActualOvertimeCost;
}

void Resource::odl_SetActualOvertimeCost(DOUBLE newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetActualOvertimeCost((DOUBLE)newVal);
}

void Resource::SetActualOvertimeCost(DOUBLE newVal)
{
	mp_cActualOvertimeCost = newVal;
}

DOUBLE Resource::odl_GetRemainingCost(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetRemainingCost();
}

DOUBLE Resource::GetRemainingCost(void)
{
	return (DOUBLE) mp_cRemainingCost;
}

void Resource::odl_SetRemainingCost(DOUBLE newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetRemainingCost((DOUBLE)newVal);
}

void Resource::SetRemainingCost(DOUBLE newVal)
{
	mp_cRemainingCost = newVal;
}

DOUBLE Resource::odl_GetRemainingOvertimeCost(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetRemainingOvertimeCost();
}

DOUBLE Resource::GetRemainingOvertimeCost(void)
{
	return (DOUBLE) mp_cRemainingOvertimeCost;
}

void Resource::odl_SetRemainingOvertimeCost(DOUBLE newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetRemainingOvertimeCost((DOUBLE)newVal);
}

void Resource::SetRemainingOvertimeCost(DOUBLE newVal)
{
	mp_cRemainingOvertimeCost = newVal;
}

FLOAT Resource::odl_GetWorkVariance(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetWorkVariance();
}

FLOAT Resource::GetWorkVariance(void)
{
	return (FLOAT) mp_fWorkVariance;
}

void Resource::odl_SetWorkVariance(FLOAT newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetWorkVariance((FLOAT)newVal);
}

void Resource::SetWorkVariance(FLOAT newVal)
{
	mp_fWorkVariance = newVal;
}

FLOAT Resource::odl_GetCostVariance(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetCostVariance();
}

FLOAT Resource::GetCostVariance(void)
{
	return (FLOAT) mp_fCostVariance;
}

void Resource::odl_SetCostVariance(FLOAT newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetCostVariance((FLOAT)newVal);
}

void Resource::SetCostVariance(FLOAT newVal)
{
	mp_fCostVariance = newVal;
}

FLOAT Resource::odl_GetSV(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetSV();
}

FLOAT Resource::GetSV(void)
{
	return (FLOAT) mp_fSV;
}

void Resource::odl_SetSV(FLOAT newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetSV((FLOAT)newVal);
}

void Resource::SetSV(FLOAT newVal)
{
	mp_fSV = newVal;
}

FLOAT Resource::odl_GetCV(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetCV();
}

FLOAT Resource::GetCV(void)
{
	return (FLOAT) mp_fCV;
}

void Resource::odl_SetCV(FLOAT newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetCV((FLOAT)newVal);
}

void Resource::SetCV(FLOAT newVal)
{
	mp_fCV = newVal;
}

FLOAT Resource::odl_GetACWP(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetACWP();
}

FLOAT Resource::GetACWP(void)
{
	return (FLOAT) mp_fACWP;
}

void Resource::odl_SetACWP(FLOAT newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetACWP((FLOAT)newVal);
}

void Resource::SetACWP(FLOAT newVal)
{
	mp_fACWP = newVal;
}

LONG Resource::odl_GetCalendarUID(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetCalendarUID();
}

LONG Resource::GetCalendarUID(void)
{
	return (LONG) mp_lCalendarUID;
}

void Resource::odl_SetCalendarUID(LONG newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetCalendarUID((LONG)newVal);
}

void Resource::SetCalendarUID(LONG newVal)
{
	mp_lCalendarUID = newVal;
}

BSTR Resource::odl_GetNotes(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetNotes().AllocSysString();
}

CString Resource::GetNotes(void)
{
	return (CString) mp_sNotes;
}

void Resource::odl_SetNotes(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetNotes(newVal);
}

void Resource::SetNotes(CString newVal)
{
	mp_sNotes = newVal;
}

FLOAT Resource::odl_GetBCWS(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetBCWS();
}

FLOAT Resource::GetBCWS(void)
{
	return (FLOAT) mp_fBCWS;
}

void Resource::odl_SetBCWS(FLOAT newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetBCWS((FLOAT)newVal);
}

void Resource::SetBCWS(FLOAT newVal)
{
	mp_fBCWS = newVal;
}

FLOAT Resource::odl_GetBCWP(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetBCWP();
}

FLOAT Resource::GetBCWP(void)
{
	return (FLOAT) mp_fBCWP;
}

void Resource::odl_SetBCWP(FLOAT newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetBCWP((FLOAT)newVal);
}

void Resource::SetBCWP(FLOAT newVal)
{
	mp_fBCWP = newVal;
}

BOOL Resource::odl_GetIsGeneric(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetIsGeneric();
}

BOOL Resource::GetIsGeneric(void)
{
	return (BOOL) mp_bIsGeneric;
}

void Resource::odl_SetIsGeneric(BOOL newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetIsGeneric((BOOL)newVal);
}

void Resource::SetIsGeneric(BOOL newVal)
{
	mp_bIsGeneric = newVal;
}

BOOL Resource::odl_GetIsInactive(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetIsInactive();
}

BOOL Resource::GetIsInactive(void)
{
	return (BOOL) mp_bIsInactive;
}

void Resource::odl_SetIsInactive(BOOL newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetIsInactive((BOOL)newVal);
}

void Resource::SetIsInactive(BOOL newVal)
{
	mp_bIsInactive = newVal;
}

BOOL Resource::odl_GetIsEnterprise(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetIsEnterprise();
}

BOOL Resource::GetIsEnterprise(void)
{
	return (BOOL) mp_bIsEnterprise;
}

void Resource::odl_SetIsEnterprise(BOOL newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetIsEnterprise((BOOL)newVal);
}

void Resource::SetIsEnterprise(BOOL newVal)
{
	mp_bIsEnterprise = newVal;
}

LONG Resource::odl_GetBookingType(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetBookingType();
}

E_BOOKINGTYPE Resource::GetBookingType(void)
{
	return (E_BOOKINGTYPE) mp_yBookingType;
}

void Resource::odl_SetBookingType(LONG newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetBookingType((E_BOOKINGTYPE)newVal);
}

void Resource::SetBookingType(E_BOOKINGTYPE newVal)
{
	mp_yBookingType = newVal;
}

IDispatch* Resource::odl_GetActualWorkProtected(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetActualWorkProtected()->GetIDispatch(TRUE);
}

Duration* Resource::GetActualWorkProtected(void)
{
	return mp_oActualWorkProtected;
}

IDispatch* Resource::odl_GetActualOvertimeWorkProtected(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetActualOvertimeWorkProtected()->GetIDispatch(TRUE);
}

Duration* Resource::GetActualOvertimeWorkProtected(void)
{
	return mp_oActualOvertimeWorkProtected;
}

BSTR Resource::odl_GetActiveDirectoryGUID(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetActiveDirectoryGUID().AllocSysString();
}

CString Resource::GetActiveDirectoryGUID(void)
{
	return (CString) mp_sActiveDirectoryGUID;
}

void Resource::odl_SetActiveDirectoryGUID(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetActiveDirectoryGUID(newVal);
}

void Resource::SetActiveDirectoryGUID(CString newVal)
{
	if (newVal.GetLength() > 16)
	{
		newVal = newVal.Mid(0, 16);
	}
	mp_sActiveDirectoryGUID = newVal;
}

DATE Resource::odl_GetCreationDate(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetCreationDate();
}

COleDateTime Resource::GetCreationDate(void)
{
	return (COleDateTime) mp_dtCreationDate;
}

void Resource::odl_SetCreationDate(DATE newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetCreationDate((COleDateTime)newVal);
}

void Resource::SetCreationDate(COleDateTime newVal)
{
	mp_dtCreationDate = newVal;
}

IDispatch* Resource::odl_GetExtendedAttribute(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetExtendedAttribute()->GetIDispatch(TRUE);
}

ResourceExtendedAttribute_C* Resource::GetExtendedAttribute(void)
{
	return mp_oExtendedAttribute_C;
}

IDispatch* Resource::odl_GetBaseline(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetBaseline()->GetIDispatch(TRUE);
}

ResourceBaseline_C* Resource::GetBaseline(void)
{
	return mp_oBaseline_C;
}

IDispatch* Resource::odl_GetOutlineCode(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetOutlineCode()->GetIDispatch(TRUE);
}

ResourceOutlineCode_C* Resource::GetOutlineCode(void)
{
	return mp_oOutlineCode_C;
}

BOOL Resource::odl_GetIsCostResource(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetIsCostResource();
}

BOOL Resource::GetIsCostResource(void)
{
	return (BOOL) mp_bIsCostResource;
}

void Resource::odl_SetIsCostResource(BOOL newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetIsCostResource((BOOL)newVal);
}

void Resource::SetIsCostResource(BOOL newVal)
{
	mp_bIsCostResource = newVal;
}

BSTR Resource::odl_GetAssnOwner(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetAssnOwner().AllocSysString();
}

CString Resource::GetAssnOwner(void)
{
	return (CString) mp_sAssnOwner;
}

void Resource::odl_SetAssnOwner(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetAssnOwner(newVal);
}

void Resource::SetAssnOwner(CString newVal)
{
	mp_sAssnOwner = newVal;
}

BSTR Resource::odl_GetAssnOwnerGuid(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetAssnOwnerGuid().AllocSysString();
}

CString Resource::GetAssnOwnerGuid(void)
{
	return (CString) mp_sAssnOwnerGuid;
}

void Resource::odl_SetAssnOwnerGuid(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetAssnOwnerGuid(newVal);
}

void Resource::SetAssnOwnerGuid(CString newVal)
{
	mp_sAssnOwnerGuid = newVal;
}

BOOL Resource::odl_GetIsBudget(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetIsBudget();
}

BOOL Resource::GetIsBudget(void)
{
	return (BOOL) mp_bIsBudget;
}

void Resource::odl_SetIsBudget(BOOL newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetIsBudget((BOOL)newVal);
}

void Resource::SetIsBudget(BOOL newVal)
{
	mp_bIsBudget = newVal;
}

IDispatch* Resource::odl_GetAvailabilityPeriods(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetAvailabilityPeriods()->GetIDispatch(TRUE);
}

ResourceAvailabilityPeriods* Resource::GetAvailabilityPeriods(void)
{
	return mp_oAvailabilityPeriods;
}

IDispatch* Resource::odl_GetRates(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetRates()->GetIDispatch(TRUE);
}

ResourceRates* Resource::GetRates(void)
{
	return mp_oRates;
}

IDispatch* Resource::odl_GetTimephasedData(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetTimephasedData()->GetIDispatch(TRUE);
}

TimephasedData_C* Resource::GetTimephasedData(void)
{
	return mp_oTimephasedData_C;
}

BSTR Resource::odl_GetKey(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetKey().AllocSysString();
}

void Resource::odl_SetKey(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	USES_CONVERSION;
	SetKey(F_TOUCSTR(T2A(newVal)));
}

CString Resource::GetKey(void)
{
	return mp_sKey;
}

void Resource::SetKey(CString newVal)
{
	mp_oCollection->mp_SetKey(&mp_sKey, newVal, MP_SET_KEY);
}

BSTR Resource::odl_GetXML(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetXML().AllocSysString();
}

void Resource::odl_SetXML(LPCTSTR sXML)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetXML(sXML);
}

BOOL Resource::IsNull(void)
{
	BOOL bReturn = TRUE;
	if (mp_lUID != 0)
	{
		bReturn = FALSE;
	}
	if (mp_lID != 0)
	{
		bReturn = FALSE;
	}
	if (mp_sName != "")
	{
		bReturn = FALSE;
	}
	if (mp_yType != T_6_MATERIAL)
	{
		bReturn = FALSE;
	}
	if (mp_bIsNull != FALSE)
	{
		bReturn = FALSE;
	}
	if (mp_sInitials != "")
	{
		bReturn = FALSE;
	}
	if (mp_sPhonetics != "")
	{
		bReturn = FALSE;
	}
	if (mp_sNTAccount != "")
	{
		bReturn = FALSE;
	}
	if (mp_sMaterialLabel != "")
	{
		bReturn = FALSE;
	}
	if (mp_sCode != "")
	{
		bReturn = FALSE;
	}
	if (mp_sGroup != "")
	{
		bReturn = FALSE;
	}
	if (mp_yWorkGroup != WG_DEFAULT)
	{
		bReturn = FALSE;
	}
	if (mp_sEmailAddress != "")
	{
		bReturn = FALSE;
	}
	if (mp_sHyperlink != "")
	{
		bReturn = FALSE;
	}
	if (mp_sHyperlinkAddress != "")
	{
		bReturn = FALSE;
	}
	if (mp_sHyperlinkSubAddress != "")
	{
		bReturn = FALSE;
	}
	if (mp_fMaxUnits != 1.0)
	{
		bReturn = FALSE;
	}
	if (mp_fPeakUnits != 0)
	{
		bReturn = FALSE;
	}
	if (mp_bOverAllocated != FALSE)
	{
		bReturn = FALSE;
	}
	if (mp_dtAvailableFrom.m_dt != 36494)
	{
		bReturn = FALSE;
	}
	if (mp_dtAvailableTo.m_dt != 36494)
	{
		bReturn = FALSE;
	}
	if (mp_dtStart.m_dt != 36494)
	{
		bReturn = FALSE;
	}
	if (mp_dtFinish.m_dt != 36494)
	{
		bReturn = FALSE;
	}
	if (mp_bCanLevel != FALSE)
	{
		bReturn = FALSE;
	}
	if (mp_yAccrueAt != AA_START)
	{
		bReturn = FALSE;
	}
	if (mp_oWork->IsNull() == FALSE)
	{
		bReturn = FALSE;
	}
	if (mp_oRegularWork->IsNull() == FALSE)
	{
		bReturn = FALSE;
	}
	if (mp_oOvertimeWork->IsNull() == FALSE)
	{
		bReturn = FALSE;
	}
	if (mp_oActualWork->IsNull() == FALSE)
	{
		bReturn = FALSE;
	}
	if (mp_oRemainingWork->IsNull() == FALSE)
	{
		bReturn = FALSE;
	}
	if (mp_oActualOvertimeWork->IsNull() == FALSE)
	{
		bReturn = FALSE;
	}
	if (mp_oRemainingOvertimeWork->IsNull() == FALSE)
	{
		bReturn = FALSE;
	}
	if (mp_lPercentWorkComplete != 0)
	{
		bReturn = FALSE;
	}
	if (mp_cStandardRate != 0)
	{
		bReturn = FALSE;
	}
	if (mp_yStandardRateFormat != SRF_M)
	{
		bReturn = FALSE;
	}
	if (mp_cCost != 0)
	{
		bReturn = FALSE;
	}
	if (mp_cOvertimeRate != 0)
	{
		bReturn = FALSE;
	}
	if (mp_yOvertimeRateFormat != ORF_M)
	{
		bReturn = FALSE;
	}
	if (mp_cOvertimeCost != 0)
	{
		bReturn = FALSE;
	}
	if (mp_cCostPerUse != 0)
	{
		bReturn = FALSE;
	}
	if (mp_cActualCost != 0)
	{
		bReturn = FALSE;
	}
	if (mp_cActualOvertimeCost != 0)
	{
		bReturn = FALSE;
	}
	if (mp_cRemainingCost != 0)
	{
		bReturn = FALSE;
	}
	if (mp_cRemainingOvertimeCost != 0)
	{
		bReturn = FALSE;
	}
	if (mp_fWorkVariance != 0)
	{
		bReturn = FALSE;
	}
	if (mp_fCostVariance != 0)
	{
		bReturn = FALSE;
	}
	if (mp_fSV != 0)
	{
		bReturn = FALSE;
	}
	if (mp_fCV != 0)
	{
		bReturn = FALSE;
	}
	if (mp_fACWP != 0)
	{
		bReturn = FALSE;
	}
	if (mp_lCalendarUID != 0)
	{
		bReturn = FALSE;
	}
	if (mp_sNotes != "")
	{
		bReturn = FALSE;
	}
	if (mp_fBCWS != 0)
	{
		bReturn = FALSE;
	}
	if (mp_fBCWP != 0)
	{
		bReturn = FALSE;
	}
	if (mp_bIsGeneric != FALSE)
	{
		bReturn = FALSE;
	}
	if (mp_bIsInactive != FALSE)
	{
		bReturn = FALSE;
	}
	if (mp_bIsEnterprise != FALSE)
	{
		bReturn = FALSE;
	}
	if (mp_yBookingType != BT_COMMITED)
	{
		bReturn = FALSE;
	}
	if (mp_oActualWorkProtected->IsNull() == FALSE)
	{
		bReturn = FALSE;
	}
	if (mp_oActualOvertimeWorkProtected->IsNull() == FALSE)
	{
		bReturn = FALSE;
	}
	if (mp_sActiveDirectoryGUID != "")
	{
		bReturn = FALSE;
	}
	if (mp_dtCreationDate.m_dt != 36494)
	{
		bReturn = FALSE;
	}
	if (mp_oExtendedAttribute_C->IsNull() == FALSE)
	{
		bReturn = FALSE;
	}
	if (mp_oBaseline_C->IsNull() == FALSE)
	{
		bReturn = FALSE;
	}
	if (mp_oOutlineCode_C->IsNull() == FALSE)
	{
		bReturn = FALSE;
	}
	if (mp_bIsCostResource != FALSE)
	{
		bReturn = FALSE;
	}
	if (mp_sAssnOwner != "")
	{
		bReturn = FALSE;
	}
	if (mp_sAssnOwnerGuid != "")
	{
		bReturn = FALSE;
	}
	if (mp_bIsBudget != FALSE)
	{
		bReturn = FALSE;
	}
	if (mp_oAvailabilityPeriods->IsNull() == FALSE)
	{
		bReturn = FALSE;
	}
	if (mp_oRates->IsNull() == FALSE)
	{
		bReturn = FALSE;
	}
	if (mp_oTimephasedData_C->IsNull() == FALSE)
	{
		bReturn = FALSE;
	}
	return bReturn;
}

CString Resource::GetXML(void)
{
	if (IsNull() == TRUE)
	{
		return "<Resource/>";
	}
	clsXML oXML("Resource");
	oXML.InitializeWriter();
	oXML.SetSupportOptional(TRUE);
	oXML.SetBoolsAreNumeric(TRUE);
	oXML.WriteProperty("UID", mp_lUID);
	oXML.WriteProperty("ID", mp_lID);
	if (mp_sName != "")
	{
		oXML.WriteProperty("Name", mp_sName);
	}
	oXML.WriteProperty("Type", mp_yType);
	oXML.WriteProperty("IsNull", mp_bIsNull);
	if (mp_sInitials != "")
	{
		oXML.WriteProperty("Initials", mp_sInitials);
	}
	if (mp_sPhonetics != "")
	{
		oXML.WriteProperty("Phonetics", mp_sPhonetics);
	}
	if (mp_sNTAccount != "")
	{
		oXML.WriteProperty("NTAccount", mp_sNTAccount);
	}
	if (mp_sMaterialLabel != "")
	{
		oXML.WriteProperty("MaterialLabel", mp_sMaterialLabel);
	}
	if (mp_sCode != "")
	{
		oXML.WriteProperty("Code", mp_sCode);
	}
	if (mp_sGroup != "")
	{
		oXML.WriteProperty("Group", mp_sGroup);
	}
	oXML.WriteProperty("WorkGroup", mp_yWorkGroup);
	if (mp_sEmailAddress != "")
	{
		oXML.WriteProperty("EmailAddress", mp_sEmailAddress);
	}
	if (mp_sHyperlink != "")
	{
		oXML.WriteProperty("Hyperlink", mp_sHyperlink);
	}
	if (mp_sHyperlinkAddress != "")
	{
		oXML.WriteProperty("HyperlinkAddress", mp_sHyperlinkAddress);
	}
	if (mp_sHyperlinkSubAddress != "")
	{
		oXML.WriteProperty("HyperlinkSubAddress", mp_sHyperlinkSubAddress);
	}
	oXML.WriteProperty("MaxUnits", mp_fMaxUnits);
	oXML.WriteProperty("PeakUnits", mp_fPeakUnits);
	oXML.WriteProperty("OverAllocated", mp_bOverAllocated);
	if (mp_dtAvailableFrom.m_dt != 36494)
	{
		oXML.WriteProperty("AvailableFrom", mp_dtAvailableFrom);
	}
	if (mp_dtAvailableTo.m_dt != 36494)
	{
		oXML.WriteProperty("AvailableTo", mp_dtAvailableTo);
	}
	if (mp_dtStart.m_dt != 36494)
	{
		oXML.WriteProperty("Start", mp_dtStart);
	}
	if (mp_dtFinish.m_dt != 36494)
	{
		oXML.WriteProperty("Finish", mp_dtFinish);
	}
	oXML.WriteProperty("CanLevel", mp_bCanLevel);
	oXML.WriteProperty("AccrueAt", mp_yAccrueAt);
	oXML.WriteProperty("Work", mp_oWork);
	oXML.WriteProperty("RegularWork", mp_oRegularWork);
	oXML.WriteProperty("OvertimeWork", mp_oOvertimeWork);
	oXML.WriteProperty("ActualWork", mp_oActualWork);
	oXML.WriteProperty("RemainingWork", mp_oRemainingWork);
	oXML.WriteProperty("ActualOvertimeWork", mp_oActualOvertimeWork);
	oXML.WriteProperty("RemainingOvertimeWork", mp_oRemainingOvertimeWork);
	oXML.WriteProperty("PercentWorkComplete", mp_lPercentWorkComplete);
	oXML.WriteProperty("StandardRate", mp_cStandardRate);
	oXML.WriteProperty("StandardRateFormat", mp_yStandardRateFormat);
	oXML.WriteProperty("Cost", mp_cCost);
	oXML.WriteProperty("OvertimeRate", mp_cOvertimeRate);
	oXML.WriteProperty("OvertimeRateFormat", mp_yOvertimeRateFormat);
	oXML.WriteProperty("OvertimeCost", mp_cOvertimeCost);
	oXML.WriteProperty("CostPerUse", mp_cCostPerUse);
	oXML.WriteProperty("ActualCost", mp_cActualCost);
	oXML.WriteProperty("ActualOvertimeCost", mp_cActualOvertimeCost);
	oXML.WriteProperty("RemainingCost", mp_cRemainingCost);
	oXML.WriteProperty("RemainingOvertimeCost", mp_cRemainingOvertimeCost);
	oXML.WriteProperty("WorkVariance", mp_fWorkVariance);
	oXML.WriteProperty("CostVariance", mp_fCostVariance);
	oXML.WriteProperty("SV", mp_fSV);
	oXML.WriteProperty("CV", mp_fCV);
	oXML.WriteProperty("ACWP", mp_fACWP);
	oXML.WriteProperty("CalendarUID", mp_lCalendarUID);
	if (mp_sNotes != "")
	{
		oXML.WriteProperty("Notes", mp_sNotes);
	}
	oXML.WriteProperty("BCWS", mp_fBCWS);
	oXML.WriteProperty("BCWP", mp_fBCWP);
	oXML.WriteProperty("IsGeneric", mp_bIsGeneric);
	oXML.WriteProperty("IsInactive", mp_bIsInactive);
	oXML.WriteProperty("IsEnterprise", mp_bIsEnterprise);
	oXML.WriteProperty("BookingType", mp_yBookingType);
	if (mp_oActualWorkProtected->IsNull() == FALSE)
	{
		oXML.WriteProperty("ActualWorkProtected", mp_oActualWorkProtected);
	}
	if (mp_oActualOvertimeWorkProtected->IsNull() == FALSE)
	{
		oXML.WriteProperty("ActualOvertimeWorkProtected", mp_oActualOvertimeWorkProtected);
	}
	if (mp_sActiveDirectoryGUID != "")
	{
		oXML.WriteProperty("ActiveDirectoryGUID", mp_sActiveDirectoryGUID);
	}
	if (mp_dtCreationDate.m_dt != 36494)
	{
		oXML.WriteProperty("CreationDate", mp_dtCreationDate);
	}
	if (mp_oExtendedAttribute_C->IsNull() == FALSE)
	{
		mp_oExtendedAttribute_C->WriteObjectProtected(oXML);
	}
	if (mp_oBaseline_C->IsNull() == FALSE)
	{
		mp_oBaseline_C->WriteObjectProtected(oXML);
	}
	if (mp_oOutlineCode_C->IsNull() == FALSE)
	{
		mp_oOutlineCode_C->WriteObjectProtected(oXML);
	}
	oXML.WriteProperty("IsCostResource", mp_bIsCostResource);
	if (mp_sAssnOwner != "")
	{
		oXML.WriteProperty("AssnOwner", mp_sAssnOwner);
	}
	if (mp_sAssnOwnerGuid != "")
	{
		oXML.WriteProperty("AssnOwnerGuid", mp_sAssnOwnerGuid);
	}
	oXML.WriteProperty("IsBudget", mp_bIsBudget);
	if (mp_oAvailabilityPeriods->IsNull() == FALSE)
	{
		oXML.WriteObject(mp_oAvailabilityPeriods->GetXML());
	}
	if (mp_oRates->IsNull() == FALSE)
	{
		oXML.WriteObject(mp_oRates->GetXML());
	}
	if (mp_oTimephasedData_C->IsNull() == FALSE)
	{
		mp_oTimephasedData_C->WriteObjectProtected(oXML);
	}
	return oXML.GetXML();
}

void Resource::SetXML(CString sXML)
{
	clsXML oXML("Resource");
	oXML.SetSupportOptional(TRUE);
	oXML.SetXML(sXML);
	oXML.InitializeReader();
	oXML.ReadProperty("UID", mp_lUID);
	oXML.ReadProperty("ID", mp_lID);
	oXML.ReadProperty("Name", mp_sName);
	if (mp_sName.GetLength() > 512)
	{
		mp_sName = mp_sName.Mid(0, 512);
	}
	oXML.ReadProperty("Type", mp_yType);
	oXML.ReadProperty("IsNull", mp_bIsNull);
	oXML.ReadProperty("Initials", mp_sInitials);
	if (mp_sInitials.GetLength() > 512)
	{
		mp_sInitials = mp_sInitials.Mid(0, 512);
	}
	oXML.ReadProperty("Phonetics", mp_sPhonetics);
	if (mp_sPhonetics.GetLength() > 512)
	{
		mp_sPhonetics = mp_sPhonetics.Mid(0, 512);
	}
	oXML.ReadProperty("NTAccount", mp_sNTAccount);
	if (mp_sNTAccount.GetLength() > 512)
	{
		mp_sNTAccount = mp_sNTAccount.Mid(0, 512);
	}
	oXML.ReadProperty("MaterialLabel", mp_sMaterialLabel);
	if (mp_sMaterialLabel.GetLength() > 512)
	{
		mp_sMaterialLabel = mp_sMaterialLabel.Mid(0, 512);
	}
	oXML.ReadProperty("Code", mp_sCode);
	if (mp_sCode.GetLength() > 512)
	{
		mp_sCode = mp_sCode.Mid(0, 512);
	}
	oXML.ReadProperty("Group", mp_sGroup);
	if (mp_sGroup.GetLength() > 512)
	{
		mp_sGroup = mp_sGroup.Mid(0, 512);
	}
	oXML.ReadProperty("WorkGroup", mp_yWorkGroup);
	oXML.ReadProperty("EmailAddress", mp_sEmailAddress);
	if (mp_sEmailAddress.GetLength() > 512)
	{
		mp_sEmailAddress = mp_sEmailAddress.Mid(0, 512);
	}
	oXML.ReadProperty("Hyperlink", mp_sHyperlink);
	if (mp_sHyperlink.GetLength() > 512)
	{
		mp_sHyperlink = mp_sHyperlink.Mid(0, 512);
	}
	oXML.ReadProperty("HyperlinkAddress", mp_sHyperlinkAddress);
	if (mp_sHyperlinkAddress.GetLength() > 512)
	{
		mp_sHyperlinkAddress = mp_sHyperlinkAddress.Mid(0, 512);
	}
	oXML.ReadProperty("HyperlinkSubAddress", mp_sHyperlinkSubAddress);
	if (mp_sHyperlinkSubAddress.GetLength() > 512)
	{
		mp_sHyperlinkSubAddress = mp_sHyperlinkSubAddress.Mid(0, 512);
	}
	oXML.ReadProperty("MaxUnits", mp_fMaxUnits);
	oXML.ReadProperty("PeakUnits", mp_fPeakUnits);
	oXML.ReadProperty("OverAllocated", mp_bOverAllocated);
	oXML.ReadProperty("AvailableFrom", mp_dtAvailableFrom);
	oXML.ReadProperty("AvailableTo", mp_dtAvailableTo);
	oXML.ReadProperty("Start", mp_dtStart);
	oXML.ReadProperty("Finish", mp_dtFinish);
	oXML.ReadProperty("CanLevel", mp_bCanLevel);
	oXML.ReadProperty("AccrueAt", mp_yAccrueAt);
	oXML.ReadProperty("Work", mp_oWork);
	oXML.ReadProperty("RegularWork", mp_oRegularWork);
	oXML.ReadProperty("OvertimeWork", mp_oOvertimeWork);
	oXML.ReadProperty("ActualWork", mp_oActualWork);
	oXML.ReadProperty("RemainingWork", mp_oRemainingWork);
	oXML.ReadProperty("ActualOvertimeWork", mp_oActualOvertimeWork);
	oXML.ReadProperty("RemainingOvertimeWork", mp_oRemainingOvertimeWork);
	oXML.ReadProperty("PercentWorkComplete", mp_lPercentWorkComplete);
	oXML.ReadProperty("StandardRate", mp_cStandardRate);
	oXML.ReadProperty("StandardRateFormat", mp_yStandardRateFormat);
	oXML.ReadProperty("Cost", mp_cCost);
	oXML.ReadProperty("OvertimeRate", mp_cOvertimeRate);
	oXML.ReadProperty("OvertimeRateFormat", mp_yOvertimeRateFormat);
	oXML.ReadProperty("OvertimeCost", mp_cOvertimeCost);
	oXML.ReadProperty("CostPerUse", mp_cCostPerUse);
	oXML.ReadProperty("ActualCost", mp_cActualCost);
	oXML.ReadProperty("ActualOvertimeCost", mp_cActualOvertimeCost);
	oXML.ReadProperty("RemainingCost", mp_cRemainingCost);
	oXML.ReadProperty("RemainingOvertimeCost", mp_cRemainingOvertimeCost);
	oXML.ReadProperty("WorkVariance", mp_fWorkVariance);
	oXML.ReadProperty("CostVariance", mp_fCostVariance);
	oXML.ReadProperty("SV", mp_fSV);
	oXML.ReadProperty("CV", mp_fCV);
	oXML.ReadProperty("ACWP", mp_fACWP);
	oXML.ReadProperty("CalendarUID", mp_lCalendarUID);
	oXML.ReadProperty("Notes", mp_sNotes);
	oXML.ReadProperty("BCWS", mp_fBCWS);
	oXML.ReadProperty("BCWP", mp_fBCWP);
	oXML.ReadProperty("IsGeneric", mp_bIsGeneric);
	oXML.ReadProperty("IsInactive", mp_bIsInactive);
	oXML.ReadProperty("IsEnterprise", mp_bIsEnterprise);
	oXML.ReadProperty("BookingType", mp_yBookingType);
	oXML.ReadProperty("ActualWorkProtected", mp_oActualWorkProtected);
	oXML.ReadProperty("ActualOvertimeWorkProtected", mp_oActualOvertimeWorkProtected);
	oXML.ReadProperty("ActiveDirectoryGUID", mp_sActiveDirectoryGUID);
	if (mp_sActiveDirectoryGUID.GetLength() > 16)
	{
		mp_sActiveDirectoryGUID = mp_sActiveDirectoryGUID.Mid(0, 16);
	}
	oXML.ReadProperty("CreationDate", mp_dtCreationDate);
	mp_oExtendedAttribute_C->ReadObjectProtected(oXML);
	mp_oBaseline_C->ReadObjectProtected(oXML);
	mp_oOutlineCode_C->ReadObjectProtected(oXML);
	oXML.ReadProperty("IsCostResource", mp_bIsCostResource);
	oXML.ReadProperty("AssnOwner", mp_sAssnOwner);
	oXML.ReadProperty("AssnOwnerGuid", mp_sAssnOwnerGuid);
	oXML.ReadProperty("IsBudget", mp_bIsBudget);
	mp_oAvailabilityPeriods->SetXML(oXML.ReadObject("AvailabilityPeriods"));
	mp_oRates->SetXML(oXML.ReadObject("Rates"));
	mp_oTimephasedData_C->ReadObjectProtected(oXML);
}