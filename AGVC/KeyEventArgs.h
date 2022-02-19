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


class KeyEventArgs : public CCmdTarget
{
	DECLARE_DYNCREATE(KeyEventArgs)

public:
	KeyEventArgs(void);
	virtual ~KeyEventArgs();
	virtual void OnFinalRelease();

// Member Variables
public:
	SHORT mp_yKeyCode;				
	BOOL mp_bCancel;				
	SHORT mp_yCharacterCode;
	BOOL mp_bAlt;
	BOOL mp_bControl;
	BOOL mp_bShift;

//Internal Properties (Extensions to ODL)
public:
	SHORT GetKeyCode(void);
	void SetKeyCode(SHORT Value);
	BOOL GetCancel(void);
	void SetCancel(BOOL Value);
	SHORT GetCharacterCode(void);
	void SetCharacterCode(SHORT Value);
	BOOL GetAlt(void);
	void SetAlt(BOOL Value);
	BOOL GetShift(void);
	void SetShift(BOOL Value);
	BOOL GetControl(void);
	void SetControl(BOOL Value);

//Internal Properties
public:

//Internal Methods
public:
	void Clear(void);

protected:
	DECLARE_DISPATCH_MAP()
	DECLARE_MESSAGE_MAP()
	DECLARE_OLECREATE(KeyEventArgs)
	DECLARE_INTERFACE_MAP()

SHORT odl_GetKeyCode(void)
{
	
	return GetKeyCode();
}

void odl_SetKeyCode(SHORT Value)
{
	
	SetKeyCode(Value);
}

BOOL odl_GetCancel(void)
{
	
	return GetCancel();
}

void odl_SetCancel(BOOL Value)
{
	
	
	SetCancel(Value);
}

SHORT odl_GetCharacterCode(void)
{
	
	return GetCharacterCode();
}

void odl_SetCharacterCode(SHORT Value)
{
	
	SetCharacterCode(Value);
}


BOOL odl_GetAlt(void)
{
	
	return GetAlt();
}

BOOL odl_GetShift(void)
{
	
	return GetShift();
}

BOOL odl_GetControl(void)
{
	
	return GetControl();
}

};