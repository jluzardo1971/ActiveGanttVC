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



class TimePeriod : public clsItemBase
{
	DECLARE_DYNCREATE(TimePeriod)

public:

	TimePeriod();
	virtual ~TimePeriod();
	virtual void OnFinalRelease();

	clsCollectionBase* mp_oCollection;
	COleDateTime mp_dtFromDate;
	COleDateTime mp_dtToDate;

	DATE odl_GetFromDate(void);
	COleDateTime GetFromDate(void);
	void odl_SetFromDate(DATE newVal);
	void SetFromDate(COleDateTime newVal);

	DATE odl_GetToDate(void);
	COleDateTime GetToDate(void);
	void odl_SetToDate(DATE newVal);
	void SetToDate(COleDateTime newVal);

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
	DECLARE_OLECREATE(TimePeriod)
	DECLARE_INTERFACE_MAP()

};