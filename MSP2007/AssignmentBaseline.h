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



class AssignmentBaseline : public clsItemBase
{
	DECLARE_DYNCREATE(AssignmentBaseline)

public:

	AssignmentBaseline();
	virtual ~AssignmentBaseline();
	virtual void OnFinalRelease();

	clsCollectionBase* mp_oCollection;
	TimephasedData_C* mp_oTimephasedData_C;
	CString mp_sNumber;
	CString mp_sStart;
	CString mp_sFinish;
	Duration* mp_oWork;
	CString mp_sCost;
	FLOAT mp_fBCWS;
	FLOAT mp_fBCWP;

	IDispatch* odl_GetTimephasedData(void);
	TimephasedData_C* GetTimephasedData(void);

	BSTR odl_GetNumber(void);
	CString GetNumber(void);
	void odl_SetNumber(LPCTSTR newVal);
	void SetNumber(CString newVal);

	BSTR odl_GetStart(void);
	CString GetStart(void);
	void odl_SetStart(LPCTSTR newVal);
	void SetStart(CString newVal);

	BSTR odl_GetFinish(void);
	CString GetFinish(void);
	void odl_SetFinish(LPCTSTR newVal);
	void SetFinish(CString newVal);

	IDispatch* odl_GetWork(void);
	Duration* GetWork(void);

	BSTR odl_GetCost(void);
	CString GetCost(void);
	void odl_SetCost(LPCTSTR newVal);
	void SetCost(CString newVal);

	FLOAT odl_GetBCWS(void);
	FLOAT GetBCWS(void);
	void odl_SetBCWS(FLOAT newVal);
	void SetBCWS(FLOAT newVal);

	FLOAT odl_GetBCWP(void);
	FLOAT GetBCWP(void);
	void odl_SetBCWP(FLOAT newVal);
	void SetBCWP(FLOAT newVal);

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
	DECLARE_OLECREATE(AssignmentBaseline)
	DECLARE_INTERFACE_MAP()

};