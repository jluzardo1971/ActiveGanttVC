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

#include "clsItemBase.h"

typedef CMap<int, int, clsItemBase*, clsItemBase*> CMapIntToItemBase;

class CActiveGanttVCCtl;

class clsCollectionBase  
{
public:
	CActiveGanttVCCtl* mp_oControl;

	clsCollectionBase(CActiveGanttVCCtl* oControl, CString sObjectName);
	virtual ~clsCollectionBase();

public:
	CString m_sObjectName;
	CMapIntToItemBase mp_aoCollection;
	BOOL mp_bAddMode;
	BOOL mp_bDescending;
	BOOL mp_bIgnoreKeyChecks;
	BOOL mp_bSortCells;
	LONG mp_lCellIndex;
	CMapStringToPtr mp_oKeys;
	CString mp_sPropertyName;
	E_SORTTYPE mp_ySortType;

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
	LONG m_lReturnIndex(CString v_sIndex, BOOL bIncludeDefault);
	template <class T> BOOL mp_CompareTypesAux(T* Value1, T* Value2);
	void m_SortCells(int CellIndex, CString sPropertyName, BOOL bDescending, E_SORTTYPE SortType, int StartIndex, int EndIndex);
	BOOL mp_bCompareTypes(clsItemBase* oClass1, clsItemBase* oClass2);
	void mp_Merge(CMapIntToItemBase* aSortMap, CMapIntToItemBase* aTempMap, int first, int mid, int last);
	void mp_Sort(CMapIntToItemBase* aSortMap, CMapIntToItemBase* aTempMap, int first = -1, int last = -1);
	void m_Sort(CString sPropertyName, BOOL bDescending, E_SORTTYPE SortType, int StartIndex, int EndIndex);
	LONG m_ReturnOptLongValue(const VARIANT &vParam, LONG lDefault);
	CString m_ReturnOptStrValue(const VARIANT &vParam, CString sDefault);
	BOOL m_ReturnOptBoolValue(const VARIANT &vParam, BOOL bDefault);
	CString m_ReturnOptIndexValue(const VARIANT FAR& vParam, CString sDefault);
	BOOL m_bDoesKeyExist(LPCTSTR v_sKey);
	clsItemBase* m_oReturnArrayElementKey(CString v_sKey);
	LONG m_lCopyAndMoveItems(LONG v_lOriginIndex, LONG v_lDestinationIndex);
	BOOL m_bIsKeyUnique(LPCTSTR v_sKey);
	void mp_SetKey(CString *sCurrentKey, CString sNewKey, int ErrNumber);
	void SetAddMode(BOOL Value);
	void m_GetKeyAndIndex(CString sIndex, CString *sKey, CString *sReturnIndex);
	void m_CollectionRemoveWhere(CString sPropertyName, CString sKey, CString sIndex);
	void m_CollectionRemoveWhereNot(CString sPropertyName, CString sValue);
	void m_CollectionChange(CString sPropertyName, CString sKey, CString sIndex, CString sNewValue);
	void m_CollectionChangeAll(CString sPropertyName, CString sNewValue);
	CString CallByName(CCmdTarget* oObject, CString sPropertyName);
	void CallByName(CCmdTarget* oObject, CString sPropertyName, CString sNewValue);
	clsItemBase* mp_GetAt(CMapIntToItemBase* oMap, int lIndex);

};