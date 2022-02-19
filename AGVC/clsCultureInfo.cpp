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
#include "clsCultureInfo.h"


// clsCultureInfo

IMPLEMENT_DYNCREATE(clsCultureInfo, CCmdTarget)

clsCultureInfo::clsCultureInfo()
{
	EnableAutomation();
	AfxOleLockApp();
}

clsCultureInfo::~clsCultureInfo()
{
	AfxOleUnlockApp();
}


void clsCultureInfo::OnFinalRelease()
{
	CCmdTarget::OnFinalRelease();
}


BEGIN_MESSAGE_MAP(clsCultureInfo, CCmdTarget)
END_MESSAGE_MAP()


BEGIN_DISPATCH_MAP(clsCultureInfo, CCmdTarget)
END_DISPATCH_MAP()

// Note: we add support for IID_IclsCultureInfo to support typesafe binding
//  from VBA.  This IID must match the GUID that is attached to the 
//  dispinterface in the .IDL file.

// {02FA1F8D-628F-4B18-82D0-8CBCD105C754}
static const IID IID_IclsCultureInfo = { 0x02FA1F8D, 0x628F, 0x4B18, { 0x82, 0xD0, 0x8C, 0xBC, 0xD1, 0x05, 0xC7, 0x54} };

BEGIN_INTERFACE_MAP(clsCultureInfo, CCmdTarget)
	INTERFACE_PART(clsCultureInfo, IID_IclsCultureInfo, Dispatch)
END_INTERFACE_MAP()

// {ADCEBA0C-B87A-44F5-BE11-29532A09421A}
IMPLEMENT_OLECREATE_FLAGS(clsCultureInfo, "AGVC.clsCultureInfo", afxRegApartmentThreading, 0xADCEBA0C, 0xB87A, 0x44F5, 0xBE, 0x11, 0x29, 0x53, 0x2A, 0x09, 0x42, 0x1A)


// clsCultureInfo message handlers