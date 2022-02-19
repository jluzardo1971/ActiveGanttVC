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



class WorkingTime : public clsItemBase
{
	DECLARE_DYNCREATE(WorkingTime)

public:

	WorkingTime();
	virtual ~WorkingTime();
	virtual void OnFinalRelease();

	clsCollectionBase* mp_oCollection;
	Time* mp_oFromTime;
	Time* mp_oToTime;

	IDispatch* odl_GetFromTime(void);
	Time* GetFromTime(void);

	IDispatch* odl_GetToTime(void);
	Time* GetToTime(void);

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
	DECLARE_OLECREATE(WorkingTime)
	DECLARE_INTERFACE_MAP()

};