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

class clsXML;

class AssignmentExtendedAttribute_C : public CCmdTarget
{
	DECLARE_DYNCREATE(AssignmentExtendedAttribute_C)

public:

	AssignmentExtendedAttribute_C();
	virtual ~AssignmentExtendedAttribute_C();
	virtual void OnFinalRelease();

	clsCollectionBase* mp_oCollection;

	LONG odl_GetCount(void);
	LONG GetCount(void);

	IDispatch* odl_Item(LPCTSTR Index);
	AssignmentExtendedAttribute* Item(CString Index);

	IDispatch* odl_Add(void);
	AssignmentExtendedAttribute* Add(void);

	void odl_Clear(void);
	void Clear(void);

	void odl_Remove(LPCTSTR Index);
	void Remove(CString Index);

	void ReadObjectProtected(clsXML &oXML);
	void WriteObjectProtected(clsXML &oXML);

	BOOL IsNull(void);

	void Initialize(void);

	void InitVars(void);

protected:
	DECLARE_DISPATCH_MAP()
	DECLARE_MESSAGE_MAP()
	DECLARE_OLECREATE(AssignmentExtendedAttribute_C)
	DECLARE_INTERFACE_MAP()

};