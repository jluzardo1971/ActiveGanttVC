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



class WBSMask : public clsItemBase
{
	DECLARE_DYNCREATE(WBSMask)

public:

	WBSMask();
	virtual ~WBSMask();
	virtual void OnFinalRelease();

	clsCollectionBase* mp_oCollection;
	LONG mp_lLevel;
	LONG mp_yType;
	CString mp_sLength;
	CString mp_sSeparator;

	LONG odl_GetLevel(void);
	LONG GetLevel(void);
	void odl_SetLevel(LONG newVal);
	void SetLevel(LONG newVal);

	LONG odl_GetType(void);
	E_TYPE_1 GetType(void);
	void odl_SetType(LONG newVal);
	void SetType(E_TYPE_1 newVal);

	BSTR odl_GetLength(void);
	CString GetLength(void);
	void odl_SetLength(LPCTSTR newVal);
	void SetLength(CString newVal);

	BSTR odl_GetSeparator(void);
	CString GetSeparator(void);
	void odl_SetSeparator(LPCTSTR newVal);
	void SetSeparator(CString newVal);

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
	DECLARE_OLECREATE(WBSMask)
	DECLARE_INTERFACE_MAP()

};