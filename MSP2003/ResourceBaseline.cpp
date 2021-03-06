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
#include "ResourceBaseline.h"

IMPLEMENT_DYNCREATE(ResourceBaseline, CCmdTarget)

//{C7D57346-97E8-45BB-B8A2-5D135EB7FBB5}
static const IID IID_IResourceBaseline = { 0xC7D57346, 0x97E8, 0x45BB, { 0xB8, 0xA2, 0x5D, 0x13, 0x5E, 0xB7, 0xFB, 0xB5} };

//{EB0C3222-55A7-4948-9667-7D6653B8944F}
IMPLEMENT_OLECREATE_FLAGS(ResourceBaseline, "MSP2003.ResourceBaseline", afxRegApartmentThreading, 0xEB0C3222, 0x55A7, 0x4948, 0x96, 0x67, 0x7D, 0x66, 0x53, 0xB8, 0x94, 0x4F)

BEGIN_DISPATCH_MAP(ResourceBaseline, CCmdTarget)
DISP_PROPERTY_EX_ID(ResourceBaseline, "lNumber", 1, odl_GetNumber, odl_SetNumber, VT_I4)
DISP_PROPERTY_EX_ID(ResourceBaseline, "oWork", 2, odl_GetWork, SetNotSupported, VT_DISPATCH)
DISP_PROPERTY_EX_ID(ResourceBaseline, "fCost", 3, odl_GetCost, odl_SetCost, VT_R4)
DISP_PROPERTY_EX_ID(ResourceBaseline, "fBCWS", 4, odl_GetBCWS, odl_SetBCWS, VT_R4)
DISP_PROPERTY_EX_ID(ResourceBaseline, "fBCWP", 5, odl_GetBCWP, odl_SetBCWP, VT_R4)
DISP_PROPERTY_EX_ID(ResourceBaseline, "Key", 6, odl_GetKey, odl_SetKey, VT_BSTR)
DISP_FUNCTION_ID(ResourceBaseline, "GetXML", 7, odl_GetXML, VT_BSTR, VTS_NONE)
DISP_FUNCTION_ID(ResourceBaseline, "SetXML", 8, odl_SetXML, VT_EMPTY, VTS_BSTR)
DISP_FUNCTION_ID(ResourceBaseline, "IsNull", 9, IsNull, VT_BOOL, VTS_NONE)
DISP_FUNCTION_ID(ResourceBaseline, "Initialize", 10, Initialize, VT_EMPTY, VTS_NONE)
END_DISPATCH_MAP()

BEGIN_INTERFACE_MAP(ResourceBaseline, CCmdTarget)
	INTERFACE_PART(ResourceBaseline, IID_IResourceBaseline, Dispatch)
END_INTERFACE_MAP()

BEGIN_MESSAGE_MAP(ResourceBaseline, CCmdTarget)
END_MESSAGE_MAP()


void ResourceBaseline::Initialize(void)
{
	InitVars();
}

ResourceBaseline::ResourceBaseline()
{
	EnableAutomation();
	AfxOleLockApp();
	InitVars();
}

void ResourceBaseline::InitVars(void)
{
	mp_lNumber = 0;
	mp_oWork = new Duration();
	mp_fCost = 0;
	mp_fBCWS = 0;
	mp_fBCWP = 0;
}

ResourceBaseline::~ResourceBaseline()
{
	delete mp_oWork;
	AfxOleUnlockApp();
}

void ResourceBaseline::OnFinalRelease()
{
	clsItemBase::OnFinalRelease();
}

LONG ResourceBaseline::odl_GetNumber(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetNumber();
}

LONG ResourceBaseline::GetNumber(void)
{
	return (LONG) mp_lNumber;
}

void ResourceBaseline::odl_SetNumber(LONG newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetNumber((LONG)newVal);
}

void ResourceBaseline::SetNumber(LONG newVal)
{
	mp_lNumber = newVal;
}

IDispatch* ResourceBaseline::odl_GetWork(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetWork()->GetIDispatch(TRUE);
}

Duration* ResourceBaseline::GetWork(void)
{
	return mp_oWork;
}

FLOAT ResourceBaseline::odl_GetCost(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetCost();
}

FLOAT ResourceBaseline::GetCost(void)
{
	return (FLOAT) mp_fCost;
}

void ResourceBaseline::odl_SetCost(FLOAT newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetCost((FLOAT)newVal);
}

void ResourceBaseline::SetCost(FLOAT newVal)
{
	mp_fCost = newVal;
}

FLOAT ResourceBaseline::odl_GetBCWS(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetBCWS();
}

FLOAT ResourceBaseline::GetBCWS(void)
{
	return (FLOAT) mp_fBCWS;
}

void ResourceBaseline::odl_SetBCWS(FLOAT newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetBCWS((FLOAT)newVal);
}

void ResourceBaseline::SetBCWS(FLOAT newVal)
{
	mp_fBCWS = newVal;
}

FLOAT ResourceBaseline::odl_GetBCWP(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetBCWP();
}

FLOAT ResourceBaseline::GetBCWP(void)
{
	return (FLOAT) mp_fBCWP;
}

void ResourceBaseline::odl_SetBCWP(FLOAT newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetBCWP((FLOAT)newVal);
}

void ResourceBaseline::SetBCWP(FLOAT newVal)
{
	mp_fBCWP = newVal;
}

BSTR ResourceBaseline::odl_GetKey(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetKey().AllocSysString();
}

void ResourceBaseline::odl_SetKey(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	USES_CONVERSION;
	SetKey(F_TOUCSTR(T2A(newVal)));
}

CString ResourceBaseline::GetKey(void)
{
	return mp_sKey;
}

void ResourceBaseline::SetKey(CString newVal)
{
	mp_oCollection->mp_SetKey(&mp_sKey, newVal, MP_SET_KEY);
}

BSTR ResourceBaseline::odl_GetXML(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetXML().AllocSysString();
}

void ResourceBaseline::odl_SetXML(LPCTSTR sXML)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetXML(sXML);
}

BOOL ResourceBaseline::IsNull(void)
{
	BOOL bReturn = TRUE;
	if (mp_lNumber != 0)
	{
		bReturn = FALSE;
	}
	if (mp_oWork->IsNull() == FALSE)
	{
		bReturn = FALSE;
	}
	if (mp_fCost != 0)
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
	return bReturn;
}

CString ResourceBaseline::GetXML(void)
{
	if (IsNull() == TRUE)
	{
		return "<Baseline/>";
	}
	clsXML oXML("Baseline");
	oXML.InitializeWriter();
	oXML.SetSupportOptional(TRUE);
	oXML.SetBoolsAreNumeric(TRUE);
	oXML.WriteProperty("Number", mp_lNumber);
	oXML.WriteProperty("Work", mp_oWork);
	oXML.WriteProperty("Cost", mp_fCost);
	oXML.WriteProperty("BCWS", mp_fBCWS);
	oXML.WriteProperty("BCWP", mp_fBCWP);
	return oXML.GetXML();
}

void ResourceBaseline::SetXML(CString sXML)
{
	clsXML oXML("Baseline");
	oXML.SetSupportOptional(TRUE);
	oXML.SetXML(sXML);
	oXML.InitializeReader();
	oXML.ReadProperty("Number", mp_lNumber);
	oXML.ReadProperty("Work", mp_oWork);
	oXML.ReadProperty("Cost", mp_fCost);
	oXML.ReadProperty("BCWS", mp_fBCWS);
	oXML.ReadProperty("BCWP", mp_fBCWP);
}