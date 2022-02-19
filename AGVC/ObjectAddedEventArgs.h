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


class ObjectAddedEventArgs : public CCmdTarget
{
	DECLARE_DYNCREATE(ObjectAddedEventArgs)

public:
	ObjectAddedEventArgs(void);
	virtual ~ObjectAddedEventArgs();
	virtual void OnFinalRelease();

// Member Variables
public:
	LONG mp_lTaskIndex;				
	LONG mp_lPredecessorObjectIndex;	
	LONG mp_lPredecessorTaskIndex;				
	LONG mp_yPredecessorType;				
	CString mp_sTaskKey;				
	CString mp_sPredecessorTaskKey;				
	LONG mp_yEventTarget;				

//Internal Properties (Extensions to ODL)
public:
	LONG GetTaskIndex(void);
	void SetTaskIndex(LONG Value);
	LONG GetPredecessorObjectIndex(void);
	void SetPredecessorObjectIndex(LONG Value);
	LONG GetPredecessorTaskIndex(void);
	void SetPredecessorTaskIndex(LONG Value);
	LONG GetPredecessorType(void);
	void SetPredecessorType(LONG Value);
	CString GetTaskKey(void);
	void SetTaskKey(CString Value);
	CString GetPredecessorTaskKey(void);
	void SetPredecessorTaskKey(CString Value);
	LONG GetEventTarget(void);
	void SetEventTarget(LONG Value);

//Internal Properties
public:

//Internal Methods
public:
	void Clear(void);

protected:
	DECLARE_DISPATCH_MAP()
	DECLARE_MESSAGE_MAP()
	DECLARE_OLECREATE(ObjectAddedEventArgs)
	DECLARE_INTERFACE_MAP()

LONG odl_GetTaskIndex(void)
{
	
	return GetTaskIndex();
}

void odl_SetTaskIndex(LONG Value)
{
	
	SetTaskIndex(Value);
}

LONG odl_GetPredecessorObjectIndex(void)
{
	
	return GetPredecessorObjectIndex();
}

void odl_SetPredecessorObjectIndex(LONG Value)
{
	
	SetPredecessorObjectIndex(Value);
}

LONG odl_GetPredecessorTaskIndex(void)
{
	
	return GetPredecessorTaskIndex();
}

void odl_SetPredecessorTaskIndex(LONG Value)
{
	
	SetPredecessorTaskIndex(Value);
}

LONG odl_GetPredecessorType(void)
{
	
	return GetPredecessorType();
}

void odl_SetPredecessorType(LONG Value)
{
	
	SetPredecessorType(Value);
}

BSTR odl_GetTaskKey(void)
{
	
	return GetTaskKey().AllocSysString();
}

void odl_SetTaskKey(LPCTSTR Value)
{
	
	SetTaskKey(Value);
}

BSTR odl_GetPredecessorTaskKey(void)
{
	
	return GetPredecessorTaskKey().AllocSysString();
}

void odl_SetPredecessorTaskKey(LPCTSTR Value)
{
	
	SetPredecessorTaskKey(Value);
}

LONG odl_GetEventTarget(void)
{
	
	return GetEventTarget();
}

void odl_SetEventTarget(LONG Value)
{
	
	SetEventTarget(Value);
}


};