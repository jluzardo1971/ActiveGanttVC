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
#include "ResourceAvailabilityPeriod.h"

IMPLEMENT_DYNCREATE(ResourceAvailabilityPeriod, CCmdTarget)

//{F9B4605B-2904-43F1-9136-E613BBB9C28A}
static const IID IID_IResourceAvailabilityPeriod = { 0xF9B4605B, 0x2904, 0x43F1, { 0x91, 0x36, 0xE6, 0x13, 0xBB, 0xB9, 0xC2, 0x8A} };

//{4CFDFFE2-48B9-45D8-8E67-032BAFB8A4B5}
IMPLEMENT_OLECREATE_FLAGS(ResourceAvailabilityPeriod, "MSP2003.ResourceAvailabilityPeriod", afxRegApartmentThreading, 0x4CFDFFE2, 0x48B9, 0x45D8, 0x8E, 0x67, 0x03, 0x2B, 0xAF, 0xB8, 0xA4, 0xB5)

BEGIN_DISPATCH_MAP(ResourceAvailabilityPeriod, CCmdTarget)
DISP_PROPERTY_EX_ID(ResourceAvailabilityPeriod, "dtAvailableFrom", 1, odl_GetAvailableFrom, odl_SetAvailableFrom, VT_DATE)
DISP_PROPERTY_EX_ID(ResourceAvailabilityPeriod, "dtAvailableTo", 2, odl_GetAvailableTo, odl_SetAvailableTo, VT_DATE)
DISP_PROPERTY_EX_ID(ResourceAvailabilityPeriod, "fAvailableUnits", 3, odl_GetAvailableUnits, odl_SetAvailableUnits, VT_R4)
DISP_PROPERTY_EX_ID(ResourceAvailabilityPeriod, "Key", 4, odl_GetKey, odl_SetKey, VT_BSTR)
DISP_FUNCTION_ID(ResourceAvailabilityPeriod, "GetXML", 5, odl_GetXML, VT_BSTR, VTS_NONE)
DISP_FUNCTION_ID(ResourceAvailabilityPeriod, "SetXML", 6, odl_SetXML, VT_EMPTY, VTS_BSTR)
DISP_FUNCTION_ID(ResourceAvailabilityPeriod, "IsNull", 7, IsNull, VT_BOOL, VTS_NONE)
DISP_FUNCTION_ID(ResourceAvailabilityPeriod, "Initialize", 8, Initialize, VT_EMPTY, VTS_NONE)
END_DISPATCH_MAP()

BEGIN_INTERFACE_MAP(ResourceAvailabilityPeriod, CCmdTarget)
	INTERFACE_PART(ResourceAvailabilityPeriod, IID_IResourceAvailabilityPeriod, Dispatch)
END_INTERFACE_MAP()

BEGIN_MESSAGE_MAP(ResourceAvailabilityPeriod, CCmdTarget)
END_MESSAGE_MAP()


void ResourceAvailabilityPeriod::Initialize(void)
{
	InitVars();
}

ResourceAvailabilityPeriod::ResourceAvailabilityPeriod()
{
	EnableAutomation();
	AfxOleLockApp();
	InitVars();
}

void ResourceAvailabilityPeriod::InitVars(void)
{
	mp_dtAvailableFrom = (DATE) 0;
	mp_dtAvailableTo = (DATE) 0;
	mp_fAvailableUnits = 0;
}

ResourceAvailabilityPeriod::~ResourceAvailabilityPeriod()
{
	AfxOleUnlockApp();
}

void ResourceAvailabilityPeriod::OnFinalRelease()
{
	clsItemBase::OnFinalRelease();
}

DATE ResourceAvailabilityPeriod::odl_GetAvailableFrom(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetAvailableFrom();
}

COleDateTime ResourceAvailabilityPeriod::GetAvailableFrom(void)
{
	return (COleDateTime) mp_dtAvailableFrom;
}

void ResourceAvailabilityPeriod::odl_SetAvailableFrom(DATE newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetAvailableFrom((COleDateTime)newVal);
}

void ResourceAvailabilityPeriod::SetAvailableFrom(COleDateTime newVal)
{
	mp_dtAvailableFrom = newVal;
}

DATE ResourceAvailabilityPeriod::odl_GetAvailableTo(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetAvailableTo();
}

COleDateTime ResourceAvailabilityPeriod::GetAvailableTo(void)
{
	return (COleDateTime) mp_dtAvailableTo;
}

void ResourceAvailabilityPeriod::odl_SetAvailableTo(DATE newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetAvailableTo((COleDateTime)newVal);
}

void ResourceAvailabilityPeriod::SetAvailableTo(COleDateTime newVal)
{
	mp_dtAvailableTo = newVal;
}

FLOAT ResourceAvailabilityPeriod::odl_GetAvailableUnits(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetAvailableUnits();
}

FLOAT ResourceAvailabilityPeriod::GetAvailableUnits(void)
{
	return (FLOAT) mp_fAvailableUnits;
}

void ResourceAvailabilityPeriod::odl_SetAvailableUnits(FLOAT newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetAvailableUnits((FLOAT)newVal);
}

void ResourceAvailabilityPeriod::SetAvailableUnits(FLOAT newVal)
{
	mp_fAvailableUnits = newVal;
}

BSTR ResourceAvailabilityPeriod::odl_GetKey(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetKey().AllocSysString();
}

void ResourceAvailabilityPeriod::odl_SetKey(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	USES_CONVERSION;
	SetKey(F_TOUCSTR(T2A(newVal)));
}

CString ResourceAvailabilityPeriod::GetKey(void)
{
	return mp_sKey;
}

void ResourceAvailabilityPeriod::SetKey(CString newVal)
{
	mp_oCollection->mp_SetKey(&mp_sKey, newVal, MP_SET_KEY);
}

BSTR ResourceAvailabilityPeriod::odl_GetXML(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetXML().AllocSysString();
}

void ResourceAvailabilityPeriod::odl_SetXML(LPCTSTR sXML)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetXML(sXML);
}

BOOL ResourceAvailabilityPeriod::IsNull(void)
{
	BOOL bReturn = TRUE;
	if (mp_dtAvailableFrom.m_dt != 36494)
	{
		bReturn = FALSE;
	}
	if (mp_dtAvailableTo.m_dt != 36494)
	{
		bReturn = FALSE;
	}
	if (mp_fAvailableUnits != 0)
	{
		bReturn = FALSE;
	}
	return bReturn;
}

CString ResourceAvailabilityPeriod::GetXML(void)
{
	if (IsNull() == TRUE)
	{
		return "<AvailabilityPeriod/>";
	}
	clsXML oXML("AvailabilityPeriod");
	oXML.InitializeWriter();
	oXML.SetSupportOptional(TRUE);
	oXML.SetBoolsAreNumeric(TRUE);
	if (mp_dtAvailableFrom.m_dt != 36494)
	{
		oXML.WriteProperty("AvailableFrom", mp_dtAvailableFrom);
	}
	if (mp_dtAvailableTo.m_dt != 36494)
	{
		oXML.WriteProperty("AvailableTo", mp_dtAvailableTo);
	}
	oXML.WriteProperty("AvailableUnits", mp_fAvailableUnits);
	return oXML.GetXML();
}

void ResourceAvailabilityPeriod::SetXML(CString sXML)
{
	clsXML oXML("AvailabilityPeriod");
	oXML.SetSupportOptional(TRUE);
	oXML.SetXML(sXML);
	oXML.InitializeReader();
	oXML.ReadProperty("AvailableFrom", mp_dtAvailableFrom);
	oXML.ReadProperty("AvailableTo", mp_dtAvailableTo);
	oXML.ReadProperty("AvailableUnits", mp_fAvailableUnits);
}