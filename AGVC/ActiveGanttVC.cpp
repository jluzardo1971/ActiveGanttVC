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
#include "cathelp.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

CActiveGanttVCApp NEAR theApp;

const GUID CDECL BASED_CODE _tlid = {0x7e44f611,0x0e19,0x4830,{0xac,0x10,0xd6,0x2e,0x10,0xc1,0xaf,0x42}};
const WORD _wVerMajor = 1;
const WORD _wVerMinor = 0;
const CATID CATID_SafeForScripting = {0x7dd95801,0x9882,0x11cf,{0x9f,0xa9,0x00,0xaa,0x00,0x6c,0x42,0xc4}};
const CATID CATID_SafeForInitializing = {0x7dd95802,0x9882,0x11cf,{0x9f,0xa9,0x00,0xaa,0x00,0x6c,0x42,0xc4}};
const GUID CDECL BASED_CODE _ctlid = {0x688b95d3,0xc09a,0x4e7d,{0x8f,0x01,0xf4,0xba,0xc5,0x9c,0xca,0x18}};

BOOL CActiveGanttVCApp::InitInstance()
{
	BOOL bInit = COleControlModule::InitInstance();

	if (bInit)
	{
		//TODO
	}

	return bInit;
}

int CActiveGanttVCApp::ExitInstance()
{


	return COleControlModule::ExitInstance();
}

STDAPI DllRegisterServer(void)
{

	AFX_MANAGE_STATE(_afxModuleAddrThis);

	if (!AfxOleRegisterTypeLib(AfxGetInstanceHandle(), _tlid))
	  return ResultFromScode(SELFREG_E_TYPELIB);

	if (!COleObjectFactoryEx::UpdateRegistryAll(TRUE))
	  return ResultFromScode(SELFREG_E_CLASS);

	// Mark the control as safe for initializing.

	if (FAILED( CreateComponentCategory(
		  CATID_SafeForInitializing,
		  L"Controls safely initializable from persistent data") ))
		return ResultFromScode(SELFREG_E_CLASS);

	if (FAILED( RegisterCLSIDInCategory(
		  _ctlid, CATID_SafeForInitializing) ))
		return ResultFromScode(SELFREG_E_CLASS);

	// Mark the control as safe for scripting.

	if (FAILED( CreateComponentCategory(
		  CATID_SafeForScripting,
		  L"Controls that are safely scriptable") ))
		return ResultFromScode(SELFREG_E_CLASS);

	if (FAILED( RegisterCLSIDInCategory(
		  _ctlid, CATID_SafeForScripting) ))
		return ResultFromScode(SELFREG_E_CLASS);



	return NOERROR;
}

STDAPI DllUnregisterServer(void)
{
	AFX_MANAGE_STATE(_afxModuleAddrThis);

	if (!AfxOleUnregisterTypeLib(_tlid, _wVerMajor, _wVerMinor))
		return ResultFromScode(SELFREG_E_TYPELIB);

	if (!COleObjectFactoryEx::UpdateRegistryAll(FALSE))
		return ResultFromScode(SELFREG_E_CLASS);

	return NOERROR;
}