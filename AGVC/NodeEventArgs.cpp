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
#include "NodeEventArgs.h"

IMPLEMENT_DYNCREATE(NodeEventArgs, CCmdTarget)

//{ABC03980-6B75-4BAB-B542-01566BA94DDF}
static const IID IID_INodeEventArgs = { 0xABC03980, 0x6B75, 0x4BAB, { 0xB5, 0x42, 0x01, 0x56, 0x6B, 0xA9, 0x4D, 0xDF} };

//{B6B5372D-5328-4A1A-85A7-4378C4F080DF}
IMPLEMENT_OLECREATE_FLAGS(NodeEventArgs, "AGVC.NodeEventArgs", afxRegApartmentThreading, 0xB6B5372D, 0x5328, 0x4A1A, 0x85, 0xA7, 0x43, 0x78, 0xC4, 0xF0, 0x80, 0xDF)

BEGIN_DISPATCH_MAP(NodeEventArgs, CCmdTarget)
	DISP_PROPERTY_EX_ID(NodeEventArgs, "Index", 1, odl_GetIndex, odl_SetIndex, VT_I4) 
END_DISPATCH_MAP()

BEGIN_INTERFACE_MAP(NodeEventArgs, CCmdTarget)
	INTERFACE_PART(NodeEventArgs, IID_INodeEventArgs, Dispatch)
END_INTERFACE_MAP()

BEGIN_MESSAGE_MAP(NodeEventArgs, CCmdTarget)
END_MESSAGE_MAP()

NodeEventArgs::NodeEventArgs(void)
{
	EnableAutomation();
	AfxOleLockApp();
	Clear();
}

NodeEventArgs::~NodeEventArgs()
{
	AfxOleUnlockApp();
}

void NodeEventArgs::OnFinalRelease()
{
	CCmdTarget::OnFinalRelease();
}

LONG NodeEventArgs::GetIndex(void)
{
	return mp_lIndex;
}

void NodeEventArgs::SetIndex(LONG Value)
{
	mp_lIndex = Value;
}

void NodeEventArgs::Clear(void)
{
	mp_lIndex = 0;
}