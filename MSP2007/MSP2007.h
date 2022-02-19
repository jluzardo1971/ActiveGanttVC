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
#if !defined(AFX_MSP2007_H__D0278FF2_CDC8_4822_AA96_25D28EA6BE9C__INCLUDED_)
#define AFX_MSP2007_H__D0278FF2_CDC8_4822_AA96_25D28EA6BE9C__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

// MSP2007.h : main header file for MSP2007.DLL

#if !defined( __AFXCTL_H__ )
	#error include 'afxctl.h' before including this file
#endif

#include "MSP2007_resource.h"       // main symbols

/////////////////////////////////////////////////////////////////////////////
// CMSP2007App : See MSP2007.cpp for implementation.

class CMSP2007App : public COleControlModule
{
public:
	BOOL InitInstance();
	int ExitInstance();
};

extern const GUID CDECL _tlid;
extern const WORD _wVerMajor;
extern const WORD _wVerMinor;

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_MSP2007_H__D0278FF2_CDC8_4822_AA96_25D28EA6BE9C__INCLUDED)