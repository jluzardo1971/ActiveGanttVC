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



class TaskBaseline : public clsItemBase
{
	DECLARE_DYNCREATE(TaskBaseline)

public:

	TaskBaseline();
	virtual ~TaskBaseline();
	virtual void OnFinalRelease();

	clsCollectionBase* mp_oCollection;
	TimephasedData_C* mp_oTimephasedData_C;
	LONG mp_lNumber;
	BOOL mp_bInterim;
	COleDateTime mp_dtStart;
	COleDateTime mp_dtFinish;
	Duration* mp_oDuration;
	LONG mp_yDurationFormat;
	BOOL mp_bEstimatedDuration;
	Duration* mp_oWork;
	DOUBLE mp_cCost;
	FLOAT mp_fBCWS;
	FLOAT mp_fBCWP;
	FLOAT mp_fFixedCost;

	IDispatch* odl_GetTimephasedData(void);
	TimephasedData_C* GetTimephasedData(void);

	LONG odl_GetNumber(void);
	LONG GetNumber(void);
	void odl_SetNumber(LONG newVal);
	void SetNumber(LONG newVal);

	BOOL odl_GetInterim(void);
	BOOL GetInterim(void);
	void odl_SetInterim(BOOL newVal);
	void SetInterim(BOOL newVal);

	DATE odl_GetStart(void);
	COleDateTime GetStart(void);
	void odl_SetStart(DATE newVal);
	void SetStart(COleDateTime newVal);

	DATE odl_GetFinish(void);
	COleDateTime GetFinish(void);
	void odl_SetFinish(DATE newVal);
	void SetFinish(COleDateTime newVal);

	IDispatch* odl_GetDuration(void);
	Duration* GetDuration(void);

	LONG odl_GetDurationFormat(void);
	E_DURATIONFORMAT GetDurationFormat(void);
	void odl_SetDurationFormat(LONG newVal);
	void SetDurationFormat(E_DURATIONFORMAT newVal);

	BOOL odl_GetEstimatedDuration(void);
	BOOL GetEstimatedDuration(void);
	void odl_SetEstimatedDuration(BOOL newVal);
	void SetEstimatedDuration(BOOL newVal);

	IDispatch* odl_GetWork(void);
	Duration* GetWork(void);

	DOUBLE odl_GetCost(void);
	DOUBLE GetCost(void);
	void odl_SetCost(DOUBLE newVal);
	void SetCost(DOUBLE newVal);

	FLOAT odl_GetBCWS(void);
	FLOAT GetBCWS(void);
	void odl_SetBCWS(FLOAT newVal);
	void SetBCWS(FLOAT newVal);

	FLOAT odl_GetBCWP(void);
	FLOAT GetBCWP(void);
	void odl_SetBCWP(FLOAT newVal);
	void SetBCWP(FLOAT newVal);

	FLOAT odl_GetFixedCost(void);
	FLOAT GetFixedCost(void);
	void odl_SetFixedCost(FLOAT newVal);
	void SetFixedCost(FLOAT newVal);

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
	DECLARE_OLECREATE(TaskBaseline)
	DECLARE_INTERFACE_MAP()

};