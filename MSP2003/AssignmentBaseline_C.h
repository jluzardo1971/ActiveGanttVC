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

class AssignmentBaseline_C : public CCmdTarget
{
	DECLARE_DYNCREATE(AssignmentBaseline_C)

public:

	AssignmentBaseline_C();
	virtual ~AssignmentBaseline_C();
	virtual void OnFinalRelease();

	clsCollectionBase* mp_oCollection;

	LONG odl_GetCount(void);
	LONG GetCount(void);

	IDispatch* odl_Item(LPCTSTR Index);
	AssignmentBaseline* Item(CString Index);

	IDispatch* odl_Add(void);
	AssignmentBaseline* Add(void);

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
	DECLARE_OLECREATE(AssignmentBaseline_C)
	DECLARE_INTERFACE_MAP()

};