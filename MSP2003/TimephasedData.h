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



class TimephasedData : public clsItemBase
{
	DECLARE_DYNCREATE(TimephasedData)

public:

	TimephasedData();
	virtual ~TimephasedData();
	virtual void OnFinalRelease();

	clsCollectionBase* mp_oCollection;
	LONG mp_yType;
	LONG mp_lUID;
	COleDateTime mp_dtStart;
	COleDateTime mp_dtFinish;
	LONG mp_yUnit;
	CString mp_sValue;

	LONG odl_GetType(void);
	E_TYPE_5 GetType(void);
	void odl_SetType(LONG newVal);
	void SetType(E_TYPE_5 newVal);

	LONG odl_GetUID(void);
	LONG GetUID(void);
	void odl_SetUID(LONG newVal);
	void SetUID(LONG newVal);

	DATE odl_GetStart(void);
	COleDateTime GetStart(void);
	void odl_SetStart(DATE newVal);
	void SetStart(COleDateTime newVal);

	DATE odl_GetFinish(void);
	COleDateTime GetFinish(void);
	void odl_SetFinish(DATE newVal);
	void SetFinish(COleDateTime newVal);

	LONG odl_GetUnit(void);
	E_UNIT GetUnit(void);
	void odl_SetUnit(LONG newVal);
	void SetUnit(E_UNIT newVal);

	BSTR odl_GetValue(void);
	CString GetValue(void);
	void odl_SetValue(LPCTSTR newVal);
	void SetValue(CString newVal);

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
	DECLARE_OLECREATE(TimephasedData)
	DECLARE_INTERFACE_MAP()

};