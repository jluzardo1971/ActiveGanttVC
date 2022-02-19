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



class TaskPredecessorLink : public clsItemBase
{
	DECLARE_DYNCREATE(TaskPredecessorLink)

public:

	TaskPredecessorLink();
	virtual ~TaskPredecessorLink();
	virtual void OnFinalRelease();

	clsCollectionBase* mp_oCollection;
	LONG mp_lPredecessorUID;
	LONG mp_yType;
	BOOL mp_bCrossProject;
	CString mp_sCrossProjectName;
	LONG mp_lLinkLag;
	LONG mp_yLagFormat;

	LONG odl_GetPredecessorUID(void);
	LONG GetPredecessorUID(void);
	void odl_SetPredecessorUID(LONG newVal);
	void SetPredecessorUID(LONG newVal);

	LONG odl_GetType(void);
	E_TYPE_5 GetType(void);
	void odl_SetType(LONG newVal);
	void SetType(E_TYPE_5 newVal);

	BOOL odl_GetCrossProject(void);
	BOOL GetCrossProject(void);
	void odl_SetCrossProject(BOOL newVal);
	void SetCrossProject(BOOL newVal);

	BSTR odl_GetCrossProjectName(void);
	CString GetCrossProjectName(void);
	void odl_SetCrossProjectName(LPCTSTR newVal);
	void SetCrossProjectName(CString newVal);

	LONG odl_GetLinkLag(void);
	LONG GetLinkLag(void);
	void odl_SetLinkLag(LONG newVal);
	void SetLinkLag(LONG newVal);

	LONG odl_GetLagFormat(void);
	E_LAGFORMAT GetLagFormat(void);
	void odl_SetLagFormat(LONG newVal);
	void SetLagFormat(E_LAGFORMAT newVal);

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
	DECLARE_OLECREATE(TaskPredecessorLink)
	DECLARE_INTERFACE_MAP()

};