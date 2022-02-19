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
#include "TaskPredecessorLink.h"

IMPLEMENT_DYNCREATE(TaskPredecessorLink, CCmdTarget)

//{E4E717B4-2CB2-4922-BEC4-B1FF345B3438}
static const IID IID_ITaskPredecessorLink = { 0xE4E717B4, 0x2CB2, 0x4922, { 0xBE, 0xC4, 0xB1, 0xFF, 0x34, 0x5B, 0x34, 0x38} };

//{D617BA53-435B-43C4-82B9-31AA6A04B54F}
IMPLEMENT_OLECREATE_FLAGS(TaskPredecessorLink, "MSP2010.TaskPredecessorLink", afxRegApartmentThreading, 0xD617BA53, 0x435B, 0x43C4, 0x82, 0xB9, 0x31, 0xAA, 0x6A, 0x04, 0xB5, 0x4F)

BEGIN_DISPATCH_MAP(TaskPredecessorLink, CCmdTarget)
DISP_PROPERTY_EX_ID(TaskPredecessorLink, "lPredecessorUID", 1, odl_GetPredecessorUID, odl_SetPredecessorUID, VT_I4)
DISP_PROPERTY_EX_ID(TaskPredecessorLink, "yType", 2, odl_GetType, odl_SetType, VT_I4)
DISP_PROPERTY_EX_ID(TaskPredecessorLink, "bCrossProject", 3, odl_GetCrossProject, odl_SetCrossProject, VT_BOOL)
DISP_PROPERTY_EX_ID(TaskPredecessorLink, "sCrossProjectName", 4, odl_GetCrossProjectName, odl_SetCrossProjectName, VT_BSTR)
DISP_PROPERTY_EX_ID(TaskPredecessorLink, "lLinkLag", 5, odl_GetLinkLag, odl_SetLinkLag, VT_I4)
DISP_PROPERTY_EX_ID(TaskPredecessorLink, "yLagFormat", 6, odl_GetLagFormat, odl_SetLagFormat, VT_I4)
DISP_PROPERTY_EX_ID(TaskPredecessorLink, "Key", 7, odl_GetKey, odl_SetKey, VT_BSTR)
DISP_FUNCTION_ID(TaskPredecessorLink, "GetXML", 8, odl_GetXML, VT_BSTR, VTS_NONE)
DISP_FUNCTION_ID(TaskPredecessorLink, "SetXML", 9, odl_SetXML, VT_EMPTY, VTS_BSTR)
DISP_FUNCTION_ID(TaskPredecessorLink, "IsNull", 10, IsNull, VT_BOOL, VTS_NONE)
DISP_FUNCTION_ID(TaskPredecessorLink, "Initialize", 11, Initialize, VT_EMPTY, VTS_NONE)
END_DISPATCH_MAP()

BEGIN_INTERFACE_MAP(TaskPredecessorLink, CCmdTarget)
	INTERFACE_PART(TaskPredecessorLink, IID_ITaskPredecessorLink, Dispatch)
END_INTERFACE_MAP()

BEGIN_MESSAGE_MAP(TaskPredecessorLink, CCmdTarget)
END_MESSAGE_MAP()


void TaskPredecessorLink::Initialize(void)
{
	InitVars();
}

TaskPredecessorLink::TaskPredecessorLink()
{
	EnableAutomation();
	AfxOleLockApp();
	InitVars();
}

void TaskPredecessorLink::InitVars(void)
{
	mp_lPredecessorUID = 0;
	mp_yType = T_5_FF;
	mp_bCrossProject = FALSE;
	mp_sCrossProjectName = "";
	mp_lLinkLag = 0;
	mp_yLagFormat = LF_M;
}

TaskPredecessorLink::~TaskPredecessorLink()
{
	AfxOleUnlockApp();
}

void TaskPredecessorLink::OnFinalRelease()
{
	clsItemBase::OnFinalRelease();
}

LONG TaskPredecessorLink::odl_GetPredecessorUID(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetPredecessorUID();
}

LONG TaskPredecessorLink::GetPredecessorUID(void)
{
	return (LONG) mp_lPredecessorUID;
}

void TaskPredecessorLink::odl_SetPredecessorUID(LONG newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetPredecessorUID((LONG)newVal);
}

void TaskPredecessorLink::SetPredecessorUID(LONG newVal)
{
	mp_lPredecessorUID = newVal;
}

LONG TaskPredecessorLink::odl_GetType(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetType();
}

E_TYPE_5 TaskPredecessorLink::GetType(void)
{
	return (E_TYPE_5) mp_yType;
}

void TaskPredecessorLink::odl_SetType(LONG newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetType((E_TYPE_5)newVal);
}

void TaskPredecessorLink::SetType(E_TYPE_5 newVal)
{
	mp_yType = newVal;
}

BOOL TaskPredecessorLink::odl_GetCrossProject(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetCrossProject();
}

BOOL TaskPredecessorLink::GetCrossProject(void)
{
	return (BOOL) mp_bCrossProject;
}

void TaskPredecessorLink::odl_SetCrossProject(BOOL newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetCrossProject((BOOL)newVal);
}

void TaskPredecessorLink::SetCrossProject(BOOL newVal)
{
	mp_bCrossProject = newVal;
}

BSTR TaskPredecessorLink::odl_GetCrossProjectName(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetCrossProjectName().AllocSysString();
}

CString TaskPredecessorLink::GetCrossProjectName(void)
{
	return (CString) mp_sCrossProjectName;
}

void TaskPredecessorLink::odl_SetCrossProjectName(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetCrossProjectName(newVal);
}

void TaskPredecessorLink::SetCrossProjectName(CString newVal)
{
	mp_sCrossProjectName = newVal;
}

LONG TaskPredecessorLink::odl_GetLinkLag(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetLinkLag();
}

LONG TaskPredecessorLink::GetLinkLag(void)
{
	return (LONG) mp_lLinkLag;
}

void TaskPredecessorLink::odl_SetLinkLag(LONG newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetLinkLag((LONG)newVal);
}

void TaskPredecessorLink::SetLinkLag(LONG newVal)
{
	mp_lLinkLag = newVal;
}

LONG TaskPredecessorLink::odl_GetLagFormat(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetLagFormat();
}

E_LAGFORMAT TaskPredecessorLink::GetLagFormat(void)
{
	return (E_LAGFORMAT) mp_yLagFormat;
}

void TaskPredecessorLink::odl_SetLagFormat(LONG newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetLagFormat((E_LAGFORMAT)newVal);
}

void TaskPredecessorLink::SetLagFormat(E_LAGFORMAT newVal)
{
	mp_yLagFormat = newVal;
}

BSTR TaskPredecessorLink::odl_GetKey(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetKey().AllocSysString();
}

void TaskPredecessorLink::odl_SetKey(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	USES_CONVERSION;
	SetKey(F_TOUCSTR(T2A(newVal)));
}

CString TaskPredecessorLink::GetKey(void)
{
	return mp_sKey;
}

void TaskPredecessorLink::SetKey(CString newVal)
{
	mp_oCollection->mp_SetKey(&mp_sKey, newVal, MP_SET_KEY);
}

BSTR TaskPredecessorLink::odl_GetXML(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetXML().AllocSysString();
}

void TaskPredecessorLink::odl_SetXML(LPCTSTR sXML)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetXML(sXML);
}

BOOL TaskPredecessorLink::IsNull(void)
{
	BOOL bReturn = TRUE;
	if (mp_lPredecessorUID != 0)
	{
		bReturn = FALSE;
	}
	if (mp_yType != T_5_FF)
	{
		bReturn = FALSE;
	}
	if (mp_bCrossProject != FALSE)
	{
		bReturn = FALSE;
	}
	if (mp_sCrossProjectName != "")
	{
		bReturn = FALSE;
	}
	if (mp_lLinkLag != 0)
	{
		bReturn = FALSE;
	}
	if (mp_yLagFormat != LF_M)
	{
		bReturn = FALSE;
	}
	return bReturn;
}

CString TaskPredecessorLink::GetXML(void)
{
	if (IsNull() == TRUE)
	{
		return "<PredecessorLink/>";
	}
	clsXML oXML("PredecessorLink");
	oXML.InitializeWriter();
	oXML.SetSupportOptional(TRUE);
	oXML.SetBoolsAreNumeric(TRUE);
	oXML.WriteProperty("PredecessorUID", mp_lPredecessorUID);
	oXML.WriteProperty("Type", mp_yType);
	oXML.WriteProperty("CrossProject", mp_bCrossProject);
	if (mp_sCrossProjectName != "")
	{
		oXML.WriteProperty("CrossProjectName", mp_sCrossProjectName);
	}
	oXML.WriteProperty("LinkLag", mp_lLinkLag);
	oXML.WriteProperty("LagFormat", mp_yLagFormat);
	return oXML.GetXML();
}

void TaskPredecessorLink::SetXML(CString sXML)
{
	clsXML oXML("PredecessorLink");
	oXML.SetSupportOptional(TRUE);
	oXML.SetXML(sXML);
	oXML.InitializeReader();
	oXML.ReadProperty("PredecessorUID", mp_lPredecessorUID);
	oXML.ReadProperty("Type", mp_yType);
	oXML.ReadProperty("CrossProject", mp_bCrossProject);
	oXML.ReadProperty("CrossProjectName", mp_sCrossProjectName);
	oXML.ReadProperty("LinkLag", mp_lLinkLag);
	oXML.ReadProperty("LagFormat", mp_yLagFormat);
}