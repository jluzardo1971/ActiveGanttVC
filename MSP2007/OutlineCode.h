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



class OutlineCode : public clsItemBase
{
	DECLARE_DYNCREATE(OutlineCode)

public:

	OutlineCode();
	virtual ~OutlineCode();
	virtual void OnFinalRelease();

	clsCollectionBase* mp_oCollection;
	CString mp_sGuid;
	CString mp_sFieldID;
	CString mp_sFieldName;
	CString mp_sAlias;
	CString mp_sPhoneticAlias;
	OutlineCodeValues* mp_oValues;
	BOOL mp_bEnterprise;
	LONG mp_lEnterpriseOutlineCodeAlias;
	BOOL mp_bResourceSubstitutionEnabled;
	BOOL mp_bLeafOnly;
	BOOL mp_bAllLevelsRequired;
	BOOL mp_bOnlyTableValuesAllowed;
	OutlineCodeMasks* mp_oMasks;

	BSTR odl_GetGuid(void);
	CString GetGuid(void);
	void odl_SetGuid(LPCTSTR newVal);
	void SetGuid(CString newVal);

	BSTR odl_GetFieldID(void);
	CString GetFieldID(void);
	void odl_SetFieldID(LPCTSTR newVal);
	void SetFieldID(CString newVal);

	BSTR odl_GetFieldName(void);
	CString GetFieldName(void);
	void odl_SetFieldName(LPCTSTR newVal);
	void SetFieldName(CString newVal);

	BSTR odl_GetAlias(void);
	CString GetAlias(void);
	void odl_SetAlias(LPCTSTR newVal);
	void SetAlias(CString newVal);

	BSTR odl_GetPhoneticAlias(void);
	CString GetPhoneticAlias(void);
	void odl_SetPhoneticAlias(LPCTSTR newVal);
	void SetPhoneticAlias(CString newVal);

	IDispatch* odl_GetValues(void);
	OutlineCodeValues* GetValues(void);

	BOOL odl_GetEnterprise(void);
	BOOL GetEnterprise(void);
	void odl_SetEnterprise(BOOL newVal);
	void SetEnterprise(BOOL newVal);

	LONG odl_GetEnterpriseOutlineCodeAlias(void);
	LONG GetEnterpriseOutlineCodeAlias(void);
	void odl_SetEnterpriseOutlineCodeAlias(LONG newVal);
	void SetEnterpriseOutlineCodeAlias(LONG newVal);

	BOOL odl_GetResourceSubstitutionEnabled(void);
	BOOL GetResourceSubstitutionEnabled(void);
	void odl_SetResourceSubstitutionEnabled(BOOL newVal);
	void SetResourceSubstitutionEnabled(BOOL newVal);

	BOOL odl_GetLeafOnly(void);
	BOOL GetLeafOnly(void);
	void odl_SetLeafOnly(BOOL newVal);
	void SetLeafOnly(BOOL newVal);

	BOOL odl_GetAllLevelsRequired(void);
	BOOL GetAllLevelsRequired(void);
	void odl_SetAllLevelsRequired(BOOL newVal);
	void SetAllLevelsRequired(BOOL newVal);

	BOOL odl_GetOnlyTableValuesAllowed(void);
	BOOL GetOnlyTableValuesAllowed(void);
	void odl_SetOnlyTableValuesAllowed(BOOL newVal);
	void SetOnlyTableValuesAllowed(BOOL newVal);

	IDispatch* odl_GetMasks(void);
	OutlineCodeMasks* GetMasks(void);

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
	DECLARE_OLECREATE(OutlineCode)
	DECLARE_INTERFACE_MAP()

};