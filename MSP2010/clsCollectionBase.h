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

typedef CMap<int, int, clsItemBase*, clsItemBase*> CMapIntToItemBase;

class clsCollectionBase  
{
public:
	clsCollectionBase();
	clsCollectionBase(CString sObjectName);
	virtual ~clsCollectionBase();

public:
	CString m_sObjectName;
	CMapIntToItemBase mp_aoCollection;
	BOOL mp_bAddMode;
	BOOL mp_bIgnoreKeyChecks;
	CMapStringToPtr mp_oKeys;
	CString mp_sPropertyName;

public:
	clsItemBase* m_oReturnArrayElement(LONG r_lIndex);
	void m_ReconstKeys();
	void mp_oKeys_Clear();
	void mp_aoCollection_Clear();
	int m_lFindIndexByKey(CString v_sKey);
	clsItemBase* m_oItem(CString v_lIndex, LONG v_lErr1 = 50000, LONG v_lErr2 = 50000, LONG v_lErr3 = 50000, LONG v_lErr4 = 50000);
	void m_Remove(CString v_lIndex, LONG v_lErr1, LONG v_lErr2, LONG v_lErr3, LONG v_lErr4);
	void m_Clear();
	LONG m_lCount();
	void m_Add(clsItemBase* r_oObject, CString v_sKey, LONG v_lErr1, LONG v_lErr2, BOOL v_bKeyRequired, LONG v_lKeyError);
	BOOL m_bIsKeyUnique(LPCTSTR v_sKey);
	void mp_SetKey(CString *sCurrentKey, CString sNewKey, int ErrNumber);
	void SetAddMode(BOOL Value);
};