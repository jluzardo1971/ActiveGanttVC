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
#include "PredecessorExceptionEventArgs.h"


// PredecessorExceptionEventArgs

IMPLEMENT_DYNCREATE(PredecessorExceptionEventArgs, CCmdTarget)


PredecessorExceptionEventArgs::PredecessorExceptionEventArgs()
{
	EnableAutomation();
	AfxOleLockApp();
	Clear();
}

PredecessorExceptionEventArgs::~PredecessorExceptionEventArgs()
{
	// To terminate the application when all objects created with
	// 	with OLE automation, the destructor calls AfxOleUnlockApp.
	
	AfxOleUnlockApp();
}


void PredecessorExceptionEventArgs::OnFinalRelease()
{
	// When the last reference for an automation object is released
	// OnFinalRelease is called.  The base class will automatically
	// deletes the object.  Add additional cleanup required for your
	// object before calling the base class.

	CCmdTarget::OnFinalRelease();
}


BEGIN_MESSAGE_MAP(PredecessorExceptionEventArgs, CCmdTarget)
END_MESSAGE_MAP()


BEGIN_DISPATCH_MAP(PredecessorExceptionEventArgs, CCmdTarget)
	DISP_PROPERTY_EX_ID(PredecessorExceptionEventArgs, "PredecessorIndex", 1, odl_GetPredecessorIndex, odl_SetPredecessorIndex, VT_I4) 
	DISP_PROPERTY_EX_ID(PredecessorExceptionEventArgs, "PredecessorType", 2, odl_GetPredecessorType, odl_SetPredecessorType, VT_I4) 
END_DISPATCH_MAP()

//{C6D3BD00-2EE8-49D2-9659-5D6E430C4EC5}
static const IID IID_IPredecessorExceptionEventArgs = { 0xC6D3BD00, 0x2EE8, 0x49D2, { 0x96, 0x59, 0x5D, 0x6E, 0x43, 0x0C, 0x4E, 0xC5} };

//{A13DA36E-68CD-4D20-890B-6D70E09A6938}
IMPLEMENT_OLECREATE_FLAGS(PredecessorExceptionEventArgs, "AGVC.PredecessorExceptionEventArgs", afxRegApartmentThreading, 0xA13DA36E, 0x68CD, 0x4D20, 0x89, 0x0B, 0x6D, 0x70, 0xE0, 0x9A, 0x69, 0x38)

BEGIN_INTERFACE_MAP(PredecessorExceptionEventArgs, CCmdTarget)
	INTERFACE_PART(PredecessorExceptionEventArgs, IID_IPredecessorExceptionEventArgs, Dispatch)
END_INTERFACE_MAP()

// PredecessorExceptionEventArgs message handlers



LONG PredecessorExceptionEventArgs::GetPredecessorIndex(void)
{
	return mp_lPredecessorIndex;
}

void PredecessorExceptionEventArgs::SetPredecessorIndex(LONG Value)
{
	mp_lPredecessorIndex = Value;
}

LONG PredecessorExceptionEventArgs::GetPredecessorType(void)
{
	return mp_yPredecessorType;
}

void PredecessorExceptionEventArgs::SetPredecessorType(LONG Value)
{
	mp_yPredecessorType = Value;
}

void PredecessorExceptionEventArgs::Clear(void)
{
	mp_lPredecessorIndex = 0;
	mp_yPredecessorType = 0;
}