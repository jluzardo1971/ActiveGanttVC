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



class ExtendedAttribute : public clsItemBase
{
	DECLARE_DYNCREATE(ExtendedAttribute)

public:

	ExtendedAttribute();
	virtual ~ExtendedAttribute();
	virtual void OnFinalRelease();

	clsCollectionBase* mp_oCollection;
	CString mp_sFieldID;
	CString mp_sFieldName;
	LONG mp_yCFType;
	CString mp_sGuid;
	LONG mp_yElemType;
	LONG mp_lMaxMultiValues;
	BOOL mp_bUserDef;
	CString mp_sAlias;
	CString mp_sSecondaryPID;
	BOOL mp_bAutoRollDown;
	CString mp_sDefaultGuid;
	CString mp_sLtuid;
	CString mp_sPhoneticAlias;
	LONG mp_yRollupType;
	LONG mp_yCalculationType;
	CString mp_sFormula;
	BOOL mp_bRestrictValues;
	LONG mp_yValuelistSortOrder;
	BOOL mp_bAppendNewValues;
	CString mp_sDefault;
	ExtendedAttributeValueList* mp_oValueList;

	BSTR odl_GetFieldID(void);
	CString GetFieldID(void);
	void odl_SetFieldID(LPCTSTR newVal);
	void SetFieldID(CString newVal);

	BSTR odl_GetFieldName(void);
	CString GetFieldName(void);
	void odl_SetFieldName(LPCTSTR newVal);
	void SetFieldName(CString newVal);

	LONG odl_GetCFType(void);
	E_CFTYPE GetCFType(void);
	void odl_SetCFType(LONG newVal);
	void SetCFType(E_CFTYPE newVal);

	BSTR odl_GetGuid(void);
	CString GetGuid(void);
	void odl_SetGuid(LPCTSTR newVal);
	void SetGuid(CString newVal);

	LONG odl_GetElemType(void);
	E_ELEMTYPE GetElemType(void);
	void odl_SetElemType(LONG newVal);
	void SetElemType(E_ELEMTYPE newVal);

	LONG odl_GetMaxMultiValues(void);
	LONG GetMaxMultiValues(void);
	void odl_SetMaxMultiValues(LONG newVal);
	void SetMaxMultiValues(LONG newVal);

	BOOL odl_GetUserDef(void);
	BOOL GetUserDef(void);
	void odl_SetUserDef(BOOL newVal);
	void SetUserDef(BOOL newVal);

	BSTR odl_GetAlias(void);
	CString GetAlias(void);
	void odl_SetAlias(LPCTSTR newVal);
	void SetAlias(CString newVal);

	BSTR odl_GetSecondaryPID(void);
	CString GetSecondaryPID(void);
	void odl_SetSecondaryPID(LPCTSTR newVal);
	void SetSecondaryPID(CString newVal);

	BOOL odl_GetAutoRollDown(void);
	BOOL GetAutoRollDown(void);
	void odl_SetAutoRollDown(BOOL newVal);
	void SetAutoRollDown(BOOL newVal);

	BSTR odl_GetDefaultGuid(void);
	CString GetDefaultGuid(void);
	void odl_SetDefaultGuid(LPCTSTR newVal);
	void SetDefaultGuid(CString newVal);

	BSTR odl_GetLtuid(void);
	CString GetLtuid(void);
	void odl_SetLtuid(LPCTSTR newVal);
	void SetLtuid(CString newVal);

	BSTR odl_GetPhoneticAlias(void);
	CString GetPhoneticAlias(void);
	void odl_SetPhoneticAlias(LPCTSTR newVal);
	void SetPhoneticAlias(CString newVal);

	LONG odl_GetRollupType(void);
	E_ROLLUPTYPE GetRollupType(void);
	void odl_SetRollupType(LONG newVal);
	void SetRollupType(E_ROLLUPTYPE newVal);

	LONG odl_GetCalculationType(void);
	E_CALCULATIONTYPE GetCalculationType(void);
	void odl_SetCalculationType(LONG newVal);
	void SetCalculationType(E_CALCULATIONTYPE newVal);

	BSTR odl_GetFormula(void);
	CString GetFormula(void);
	void odl_SetFormula(LPCTSTR newVal);
	void SetFormula(CString newVal);

	BOOL odl_GetRestrictValues(void);
	BOOL GetRestrictValues(void);
	void odl_SetRestrictValues(BOOL newVal);
	void SetRestrictValues(BOOL newVal);

	LONG odl_GetValuelistSortOrder(void);
	E_VALUELISTSORTORDER GetValuelistSortOrder(void);
	void odl_SetValuelistSortOrder(LONG newVal);
	void SetValuelistSortOrder(E_VALUELISTSORTORDER newVal);

	BOOL odl_GetAppendNewValues(void);
	BOOL GetAppendNewValues(void);
	void odl_SetAppendNewValues(BOOL newVal);
	void SetAppendNewValues(BOOL newVal);

	BSTR odl_GetDefault(void);
	CString GetDefault(void);
	void odl_SetDefault(LPCTSTR newVal);
	void SetDefault(CString newVal);

	IDispatch* odl_GetValueList(void);
	ExtendedAttributeValueList* GetValueList(void);

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
	DECLARE_OLECREATE(ExtendedAttribute)
	DECLARE_INTERFACE_MAP()

};