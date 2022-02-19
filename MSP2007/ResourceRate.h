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



class ResourceRate : public clsItemBase
{
	DECLARE_DYNCREATE(ResourceRate)

public:

	ResourceRate();
	virtual ~ResourceRate();
	virtual void OnFinalRelease();

	clsCollectionBase* mp_oCollection;
	COleDateTime mp_dtRatesFrom;
	COleDateTime mp_dtRatesTo;
	LONG mp_yRateTable;
	DOUBLE mp_cStandardRate;
	LONG mp_yStandardRateFormat;
	DOUBLE mp_cOvertimeRate;
	LONG mp_yOvertimeRateFormat;
	DOUBLE mp_cCostPerUse;

	DATE odl_GetRatesFrom(void);
	COleDateTime GetRatesFrom(void);
	void odl_SetRatesFrom(DATE newVal);
	void SetRatesFrom(COleDateTime newVal);

	DATE odl_GetRatesTo(void);
	COleDateTime GetRatesTo(void);
	void odl_SetRatesTo(DATE newVal);
	void SetRatesTo(COleDateTime newVal);

	LONG odl_GetRateTable(void);
	E_RATETABLE GetRateTable(void);
	void odl_SetRateTable(LONG newVal);
	void SetRateTable(E_RATETABLE newVal);

	DOUBLE odl_GetStandardRate(void);
	DOUBLE GetStandardRate(void);
	void odl_SetStandardRate(DOUBLE newVal);
	void SetStandardRate(DOUBLE newVal);

	LONG odl_GetStandardRateFormat(void);
	E_STANDARDRATEFORMAT_1 GetStandardRateFormat(void);
	void odl_SetStandardRateFormat(LONG newVal);
	void SetStandardRateFormat(E_STANDARDRATEFORMAT_1 newVal);

	DOUBLE odl_GetOvertimeRate(void);
	DOUBLE GetOvertimeRate(void);
	void odl_SetOvertimeRate(DOUBLE newVal);
	void SetOvertimeRate(DOUBLE newVal);

	LONG odl_GetOvertimeRateFormat(void);
	E_OVERTIMERATEFORMAT GetOvertimeRateFormat(void);
	void odl_SetOvertimeRateFormat(LONG newVal);
	void SetOvertimeRateFormat(E_OVERTIMERATEFORMAT newVal);

	DOUBLE odl_GetCostPerUse(void);
	DOUBLE GetCostPerUse(void);
	void odl_SetCostPerUse(DOUBLE newVal);
	void SetCostPerUse(DOUBLE newVal);

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
	DECLARE_OLECREATE(ResourceRate)
	DECLARE_INTERFACE_MAP()

};