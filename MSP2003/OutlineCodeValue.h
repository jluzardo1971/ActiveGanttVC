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



class OutlineCodeValue : public clsItemBase
{
	DECLARE_DYNCREATE(OutlineCodeValue)

public:

	OutlineCodeValue();
	virtual ~OutlineCodeValue();
	virtual void OnFinalRelease();

	clsCollectionBase* mp_oCollection;
	LONG mp_lValueID;
	LONG mp_lParentValueID;
	CString mp_sValue;
	CString mp_sDescription;

	LONG odl_GetValueID(void);
	LONG GetValueID(void);
	void odl_SetValueID(LONG newVal);
	void SetValueID(LONG newVal);

	LONG odl_GetParentValueID(void);
	LONG GetParentValueID(void);
	void odl_SetParentValueID(LONG newVal);
	void SetParentValueID(LONG newVal);

	BSTR odl_GetValue(void);
	CString GetValue(void);
	void odl_SetValue(LPCTSTR newVal);
	void SetValue(CString newVal);

	BSTR odl_GetDescription(void);
	CString GetDescription(void);
	void odl_SetDescription(LPCTSTR newVal);
	void SetDescription(CString newVal);

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
	DECLARE_OLECREATE(OutlineCodeValue)
	DECLARE_INTERFACE_MAP()

};