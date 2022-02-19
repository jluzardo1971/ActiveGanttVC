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



class ResourceAvailabilityPeriod : public clsItemBase
{
	DECLARE_DYNCREATE(ResourceAvailabilityPeriod)

public:

	ResourceAvailabilityPeriod();
	virtual ~ResourceAvailabilityPeriod();
	virtual void OnFinalRelease();

	clsCollectionBase* mp_oCollection;
	COleDateTime mp_dtAvailableFrom;
	COleDateTime mp_dtAvailableTo;
	FLOAT mp_fAvailableUnits;

	DATE odl_GetAvailableFrom(void);
	COleDateTime GetAvailableFrom(void);
	void odl_SetAvailableFrom(DATE newVal);
	void SetAvailableFrom(COleDateTime newVal);

	DATE odl_GetAvailableTo(void);
	COleDateTime GetAvailableTo(void);
	void odl_SetAvailableTo(DATE newVal);
	void SetAvailableTo(COleDateTime newVal);

	FLOAT odl_GetAvailableUnits(void);
	FLOAT GetAvailableUnits(void);
	void odl_SetAvailableUnits(FLOAT newVal);
	void SetAvailableUnits(FLOAT newVal);

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
	DECLARE_OLECREATE(ResourceAvailabilityPeriod)
	DECLARE_INTERFACE_MAP()

};