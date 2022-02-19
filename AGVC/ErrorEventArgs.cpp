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
#include "ActiveGanttVC.h"
#include "ActiveGanttVCCtl.h"
#include "clsXML.h"
#include "ErrorEventArgs.h"

IMPLEMENT_DYNCREATE(ErrorEventArgs, CCmdTarget)

//{E8BAA510-4CF6-4B7E-811C-5912EF5EDF8B}
static const IID IID_IErrorEventArgs = { 0xE8BAA510, 0x4CF6, 0x4B7E, { 0x81, 0x1C, 0x59, 0x12, 0xEF, 0x5E, 0xDF, 0x8B} };

//{B52229FE-2B35-4446-BB40-B572B4953925}
IMPLEMENT_OLECREATE_FLAGS(ErrorEventArgs, "AGVC.ErrorEventArgs", afxRegApartmentThreading, 0xB52229FE, 0x2B35, 0x4446, 0xBB, 0x40, 0xB5, 0x72, 0xB4, 0x95, 0x39, 0x25)

BEGIN_DISPATCH_MAP(ErrorEventArgs, CCmdTarget)
	DISP_PROPERTY_EX_ID(ErrorEventArgs, "Number", 1, odl_GetNumber, odl_SetNumber, VT_I4) 
	DISP_PROPERTY_EX_ID(ErrorEventArgs, "Description", 2, odl_GetDescription, odl_SetDescription, VT_BSTR) 
	DISP_PROPERTY_EX_ID(ErrorEventArgs, "Source", 3, odl_GetSource, odl_SetSource, VT_BSTR) 
END_DISPATCH_MAP()

BEGIN_INTERFACE_MAP(ErrorEventArgs, CCmdTarget)
	INTERFACE_PART(ErrorEventArgs, IID_IErrorEventArgs, Dispatch)
END_INTERFACE_MAP()

BEGIN_MESSAGE_MAP(ErrorEventArgs, CCmdTarget)
END_MESSAGE_MAP()

ErrorEventArgs::ErrorEventArgs(void)
{
	EnableAutomation();
	AfxOleLockApp();
	Clear();
}

ErrorEventArgs::~ErrorEventArgs()
{
	AfxOleUnlockApp();
}

void ErrorEventArgs::OnFinalRelease()
{
	CCmdTarget::OnFinalRelease();
}

LONG ErrorEventArgs::GetNumber(void)
{
	return mp_lNumber;
}

void ErrorEventArgs::SetNumber(LONG Value)
{
	mp_lNumber = Value;
}

CString ErrorEventArgs::GetDescription(void)
{
	return mp_sDescription;
}

void ErrorEventArgs::SetDescription(CString Value)
{
	mp_sDescription = Value;
}

CString ErrorEventArgs::GetSource(void)
{
	return mp_sSource;
}

void ErrorEventArgs::SetSource(CString Value)
{
	mp_sSource = Value;
}

void ErrorEventArgs::Clear(void)
{
	mp_lNumber = 0;
	mp_sDescription = "";
	mp_sSource = "";
}