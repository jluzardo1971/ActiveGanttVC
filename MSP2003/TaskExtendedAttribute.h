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



class TaskExtendedAttribute : public clsItemBase
{
	DECLARE_DYNCREATE(TaskExtendedAttribute)

public:

	TaskExtendedAttribute();
	virtual ~TaskExtendedAttribute();
	virtual void OnFinalRelease();

	clsCollectionBase* mp_oCollection;
	LONG mp_lUID;
	CString mp_sFieldID;
	CString mp_sValue;
	LONG mp_lValueID;
	LONG mp_yDurationFormat;

	LONG odl_GetUID(void);
	LONG GetUID(void);
	void odl_SetUID(LONG newVal);
	void SetUID(LONG newVal);

	BSTR odl_GetFieldID(void);
	CString GetFieldID(void);
	void odl_SetFieldID(LPCTSTR newVal);
	void SetFieldID(CString newVal);

	BSTR odl_GetValue(void);
	CString GetValue(void);
	void odl_SetValue(LPCTSTR newVal);
	void SetValue(CString newVal);

	LONG odl_GetValueID(void);
	LONG GetValueID(void);
	void odl_SetValueID(LONG newVal);
	void SetValueID(LONG newVal);

	LONG odl_GetDurationFormat(void);
	E_DURATIONFORMAT GetDurationFormat(void);
	void odl_SetDurationFormat(LONG newVal);
	void SetDurationFormat(E_DURATIONFORMAT newVal);

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
	DECLARE_OLECREATE(TaskExtendedAttribute)
	DECLARE_INTERFACE_MAP()

};