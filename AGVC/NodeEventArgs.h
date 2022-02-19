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
#pragma once


class NodeEventArgs : public CCmdTarget
{
	DECLARE_DYNCREATE(NodeEventArgs)

public:
	NodeEventArgs(void);
	virtual ~NodeEventArgs();
	virtual void OnFinalRelease();

// Member Variables
public:
	LONG mp_lIndex;				

//Internal Properties (Extensions to ODL)
public:
	LONG GetIndex(void);
	void SetIndex(LONG Value);

//Internal Properties
public:

//Internal Methods
public:
	void Clear(void);

protected:
	DECLARE_DISPATCH_MAP()
	DECLARE_MESSAGE_MAP()
	DECLARE_OLECREATE(NodeEventArgs)
	DECLARE_INTERFACE_MAP()

LONG odl_GetIndex(void)
{
	
	return GetIndex();
}

void odl_SetIndex(LONG Value)
{
	
	SetIndex(Value);
}


};