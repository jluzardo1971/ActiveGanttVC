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
#include "ResourceRate.h"

IMPLEMENT_DYNCREATE(ResourceRate, CCmdTarget)

//{8E11B2EB-0AE8-48D6-8338-697F802A0053}
static const IID IID_IResourceRate = { 0x8E11B2EB, 0x0AE8, 0x48D6, { 0x83, 0x38, 0x69, 0x7F, 0x80, 0x2A, 0x00, 0x53} };

//{9759A10D-A745-4DB6-A8BD-6C34C58387F0}
IMPLEMENT_OLECREATE_FLAGS(ResourceRate, "MSP2007.ResourceRate", afxRegApartmentThreading, 0x9759A10D, 0xA745, 0x4DB6, 0xA8, 0xBD, 0x6C, 0x34, 0xC5, 0x83, 0x87, 0xF0)

BEGIN_DISPATCH_MAP(ResourceRate, CCmdTarget)
DISP_PROPERTY_EX_ID(ResourceRate, "dtRatesFrom", 1, odl_GetRatesFrom, odl_SetRatesFrom, VT_DATE)
DISP_PROPERTY_EX_ID(ResourceRate, "dtRatesTo", 2, odl_GetRatesTo, odl_SetRatesTo, VT_DATE)
DISP_PROPERTY_EX_ID(ResourceRate, "yRateTable", 3, odl_GetRateTable, odl_SetRateTable, VT_I4)
DISP_PROPERTY_EX_ID(ResourceRate, "cStandardRate", 4, odl_GetStandardRate, odl_SetStandardRate, VT_R8)
DISP_PROPERTY_EX_ID(ResourceRate, "yStandardRateFormat", 5, odl_GetStandardRateFormat, odl_SetStandardRateFormat, VT_I4)
DISP_PROPERTY_EX_ID(ResourceRate, "cOvertimeRate", 6, odl_GetOvertimeRate, odl_SetOvertimeRate, VT_R8)
DISP_PROPERTY_EX_ID(ResourceRate, "yOvertimeRateFormat", 7, odl_GetOvertimeRateFormat, odl_SetOvertimeRateFormat, VT_I4)
DISP_PROPERTY_EX_ID(ResourceRate, "cCostPerUse", 8, odl_GetCostPerUse, odl_SetCostPerUse, VT_R8)
DISP_PROPERTY_EX_ID(ResourceRate, "Key", 9, odl_GetKey, odl_SetKey, VT_BSTR)
DISP_FUNCTION_ID(ResourceRate, "GetXML", 10, odl_GetXML, VT_BSTR, VTS_NONE)
DISP_FUNCTION_ID(ResourceRate, "SetXML", 11, odl_SetXML, VT_EMPTY, VTS_BSTR)
DISP_FUNCTION_ID(ResourceRate, "IsNull", 12, IsNull, VT_BOOL, VTS_NONE)
DISP_FUNCTION_ID(ResourceRate, "Initialize", 13, Initialize, VT_EMPTY, VTS_NONE)
END_DISPATCH_MAP()

BEGIN_INTERFACE_MAP(ResourceRate, CCmdTarget)
	INTERFACE_PART(ResourceRate, IID_IResourceRate, Dispatch)
END_INTERFACE_MAP()

BEGIN_MESSAGE_MAP(ResourceRate, CCmdTarget)
END_MESSAGE_MAP()


void ResourceRate::Initialize(void)
{
	InitVars();
}

ResourceRate::ResourceRate()
{
	EnableAutomation();
	AfxOleLockApp();
	InitVars();
}

void ResourceRate::InitVars(void)
{
	mp_dtRatesFrom = (DATE) 0;
	mp_dtRatesTo = (DATE) 0;
	mp_yRateTable = RT_A;
	mp_cStandardRate = 0;
	mp_yStandardRateFormat = SRF_1_M;
	mp_cOvertimeRate = 0;
	mp_yOvertimeRateFormat = ORF_M;
	mp_cCostPerUse = 0;
}

ResourceRate::~ResourceRate()
{
	AfxOleUnlockApp();
}

void ResourceRate::OnFinalRelease()
{
	clsItemBase::OnFinalRelease();
}

DATE ResourceRate::odl_GetRatesFrom(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetRatesFrom();
}

COleDateTime ResourceRate::GetRatesFrom(void)
{
	return (COleDateTime) mp_dtRatesFrom;
}

void ResourceRate::odl_SetRatesFrom(DATE newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetRatesFrom((COleDateTime)newVal);
}

void ResourceRate::SetRatesFrom(COleDateTime newVal)
{
	mp_dtRatesFrom = newVal;
}

DATE ResourceRate::odl_GetRatesTo(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetRatesTo();
}

COleDateTime ResourceRate::GetRatesTo(void)
{
	return (COleDateTime) mp_dtRatesTo;
}

void ResourceRate::odl_SetRatesTo(DATE newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetRatesTo((COleDateTime)newVal);
}

void ResourceRate::SetRatesTo(COleDateTime newVal)
{
	mp_dtRatesTo = newVal;
}

LONG ResourceRate::odl_GetRateTable(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetRateTable();
}

E_RATETABLE ResourceRate::GetRateTable(void)
{
	return (E_RATETABLE) mp_yRateTable;
}

void ResourceRate::odl_SetRateTable(LONG newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetRateTable((E_RATETABLE)newVal);
}

void ResourceRate::SetRateTable(E_RATETABLE newVal)
{
	mp_yRateTable = newVal;
}

DOUBLE ResourceRate::odl_GetStandardRate(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetStandardRate();
}

DOUBLE ResourceRate::GetStandardRate(void)
{
	return (DOUBLE) mp_cStandardRate;
}

void ResourceRate::odl_SetStandardRate(DOUBLE newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetStandardRate((DOUBLE)newVal);
}

void ResourceRate::SetStandardRate(DOUBLE newVal)
{
	mp_cStandardRate = newVal;
}

LONG ResourceRate::odl_GetStandardRateFormat(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetStandardRateFormat();
}

E_STANDARDRATEFORMAT_1 ResourceRate::GetStandardRateFormat(void)
{
	return (E_STANDARDRATEFORMAT_1) mp_yStandardRateFormat;
}

void ResourceRate::odl_SetStandardRateFormat(LONG newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetStandardRateFormat((E_STANDARDRATEFORMAT_1)newVal);
}

void ResourceRate::SetStandardRateFormat(E_STANDARDRATEFORMAT_1 newVal)
{
	mp_yStandardRateFormat = newVal;
}

DOUBLE ResourceRate::odl_GetOvertimeRate(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetOvertimeRate();
}

DOUBLE ResourceRate::GetOvertimeRate(void)
{
	return (DOUBLE) mp_cOvertimeRate;
}

void ResourceRate::odl_SetOvertimeRate(DOUBLE newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetOvertimeRate((DOUBLE)newVal);
}

void ResourceRate::SetOvertimeRate(DOUBLE newVal)
{
	mp_cOvertimeRate = newVal;
}

LONG ResourceRate::odl_GetOvertimeRateFormat(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetOvertimeRateFormat();
}

E_OVERTIMERATEFORMAT ResourceRate::GetOvertimeRateFormat(void)
{
	return (E_OVERTIMERATEFORMAT) mp_yOvertimeRateFormat;
}

void ResourceRate::odl_SetOvertimeRateFormat(LONG newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetOvertimeRateFormat((E_OVERTIMERATEFORMAT)newVal);
}

void ResourceRate::SetOvertimeRateFormat(E_OVERTIMERATEFORMAT newVal)
{
	mp_yOvertimeRateFormat = newVal;
}

DOUBLE ResourceRate::odl_GetCostPerUse(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetCostPerUse();
}

DOUBLE ResourceRate::GetCostPerUse(void)
{
	return (DOUBLE) mp_cCostPerUse;
}

void ResourceRate::odl_SetCostPerUse(DOUBLE newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetCostPerUse((DOUBLE)newVal);
}

void ResourceRate::SetCostPerUse(DOUBLE newVal)
{
	mp_cCostPerUse = newVal;
}

BSTR ResourceRate::odl_GetKey(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetKey().AllocSysString();
}

void ResourceRate::odl_SetKey(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	USES_CONVERSION;
	SetKey(F_TOUCSTR(T2A(newVal)));
}

CString ResourceRate::GetKey(void)
{
	return mp_sKey;
}

void ResourceRate::SetKey(CString newVal)
{
	mp_oCollection->mp_SetKey(&mp_sKey, newVal, MP_SET_KEY);
}

BSTR ResourceRate::odl_GetXML(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetXML().AllocSysString();
}

void ResourceRate::odl_SetXML(LPCTSTR sXML)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetXML(sXML);
}

BOOL ResourceRate::IsNull(void)
{
	BOOL bReturn = TRUE;
	if (mp_dtRatesFrom.m_dt != 36494)
	{
		bReturn = FALSE;
	}
	if (mp_dtRatesTo.m_dt != 36494)
	{
		bReturn = FALSE;
	}
	if (mp_yRateTable != RT_A)
	{
		bReturn = FALSE;
	}
	if (mp_cStandardRate != 0)
	{
		bReturn = FALSE;
	}
	if (mp_yStandardRateFormat != SRF_1_M)
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
	if (mp_cCostPerUse != 0)
	{
		bReturn = FALSE;
	}
	return bReturn;
}

CString ResourceRate::GetXML(void)
{
	if (IsNull() == TRUE)
	{
		return "<Rate/>";
	}
	clsXML oXML("Rate");
	oXML.InitializeWriter();
	oXML.SetSupportOptional(TRUE);
	oXML.SetBoolsAreNumeric(TRUE);
	if (mp_dtRatesFrom.m_dt != 36494)
	{
		oXML.WriteProperty("RatesFrom", mp_dtRatesFrom);
	}
	if (mp_dtRatesTo.m_dt != 36494)
	{
		oXML.WriteProperty("RatesTo", mp_dtRatesTo);
	}
	oXML.WriteProperty("RateTable", mp_yRateTable);
	oXML.WriteProperty("StandardRate", mp_cStandardRate);
	oXML.WriteProperty("StandardRateFormat", mp_yStandardRateFormat);
	oXML.WriteProperty("OvertimeRate", mp_cOvertimeRate);
	oXML.WriteProperty("OvertimeRateFormat", mp_yOvertimeRateFormat);
	oXML.WriteProperty("CostPerUse", mp_cCostPerUse);
	return oXML.GetXML();
}

void ResourceRate::SetXML(CString sXML)
{
	clsXML oXML("Rate");
	oXML.SetSupportOptional(TRUE);
	oXML.SetXML(sXML);
	oXML.InitializeReader();
	oXML.ReadProperty("RatesFrom", mp_dtRatesFrom);
	oXML.ReadProperty("RatesTo", mp_dtRatesTo);
	oXML.ReadProperty("RateTable", mp_yRateTable);
	oXML.ReadProperty("StandardRate", mp_cStandardRate);
	oXML.ReadProperty("StandardRateFormat", mp_yStandardRateFormat);
	oXML.ReadProperty("OvertimeRate", mp_cOvertimeRate);
	oXML.ReadProperty("OvertimeRateFormat", mp_yOvertimeRateFormat);
	oXML.ReadProperty("CostPerUse", mp_cCostPerUse);
}