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


// PredecessorExceptionEventArgs command target

class PredecessorExceptionEventArgs : public CCmdTarget
{
	DECLARE_DYNCREATE(PredecessorExceptionEventArgs)

public:
	PredecessorExceptionEventArgs();
	virtual ~PredecessorExceptionEventArgs();

	virtual void OnFinalRelease();

// Member Variables
public:				
	LONG mp_lPredecessorIndex;						
	LONG mp_yPredecessorType;

//Internal Properties (Extensions to ODL)
public:
	LONG GetPredecessorIndex(void);
	void SetPredecessorIndex(LONG Value);
	LONG GetPredecessorType(void);
	void SetPredecessorType(LONG Value);

//Internal Properties
public:

//Internal Methods
public:
	void Clear(void);

protected:
	DECLARE_MESSAGE_MAP()
	DECLARE_OLECREATE(PredecessorExceptionEventArgs)
	DECLARE_DISPATCH_MAP()
	DECLARE_INTERFACE_MAP()

LONG odl_GetPredecessorIndex(void)
{
	
	return GetPredecessorIndex();
}

void odl_SetPredecessorIndex(LONG Value)
{
	
	SetPredecessorIndex(Value);
}

LONG odl_GetPredecessorType(void)
{
	
	return GetPredecessorType();
}

void odl_SetPredecessorType(LONG Value)
{
	
	SetPredecessorType(Value);
}

};