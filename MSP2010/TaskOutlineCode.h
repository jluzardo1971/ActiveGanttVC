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



class TaskOutlineCode : public clsItemBase
{
	DECLARE_DYNCREATE(TaskOutlineCode)

public:

	TaskOutlineCode();
	virtual ~TaskOutlineCode();
	virtual void OnFinalRelease();

	clsCollectionBase* mp_oCollection;
	CString mp_sFieldID;
	LONG mp_lValueID;
	LONG mp_lValueGUID;

	BSTR odl_GetFieldID(void);
	CString GetFieldID(void);
	void odl_SetFieldID(LPCTSTR newVal);
	void SetFieldID(CString newVal);

	LONG odl_GetValueID(void);
	LONG GetValueID(void);
	void odl_SetValueID(LONG newVal);
	void SetValueID(LONG newVal);

	LONG odl_GetValueGUID(void);
	LONG GetValueGUID(void);
	void odl_SetValueGUID(LONG newVal);
	void SetValueGUID(LONG newVal);

	BSTR odl_GetKey(void);
	CString GetKey(void);
	void odl_SetKey(LPCTSTR newVal);
	void SetKey(CString newVal);


	BSTR odl_GetXML(void);
	CString GetXML(void);

	void odl_SetXML(LPCTSTR sXML);
	void SetXML(CString sXML);

	BOOL IsNull(void);

	void Initialize(void);

	void InitVars(void);

protected:
	DECLARE_DISPATCH_MAP()
	DECLARE_MESSAGE_MAP()
	DECLARE_OLECREATE(TaskOutlineCode)
	DECLARE_INTERFACE_MAP()

};