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



class ResourceBaseline : public clsItemBase
{
	DECLARE_DYNCREATE(ResourceBaseline)

public:

	ResourceBaseline();
	virtual ~ResourceBaseline();
	virtual void OnFinalRelease();

	clsCollectionBase* mp_oCollection;
	LONG mp_lNumber;
	Duration* mp_oWork;
	FLOAT mp_fCost;
	FLOAT mp_fBCWS;
	FLOAT mp_fBCWP;

	LONG odl_GetNumber(void);
	LONG GetNumber(void);
	void odl_SetNumber(LONG newVal);
	void SetNumber(LONG newVal);

	IDispatch* odl_GetWork(void);
	Duration* GetWork(void);

	FLOAT odl_GetCost(void);
	FLOAT GetCost(void);
	void odl_SetCost(FLOAT newVal);
	void SetCost(FLOAT newVal);

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
	DECLARE_OLECREATE(ResourceBaseline)
	DECLARE_INTERFACE_MAP()

};