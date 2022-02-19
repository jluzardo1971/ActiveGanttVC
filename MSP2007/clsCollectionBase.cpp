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
#include "stdafx.h"
#include "clsCollectionBase.h"

#ifdef _DEBUG
#undef THIS_FILE
static char THIS_FILE[]=__FILE__;
#define new DEBUG_NEW
#endif

clsCollectionBase::clsCollectionBase()
{

}

clsCollectionBase::clsCollectionBase(CString sObjectName)
{
	m_sObjectName = sObjectName;
	mp_bIgnoreKeyChecks = FALSE;
}

clsCollectionBase::~clsCollectionBase()
{
	mp_aoCollection_Clear();
	mp_oKeys_Clear();
}

void clsCollectionBase::mp_oKeys_Clear()
{
	POSITION pos = mp_oKeys.GetStartPosition();
	while (pos != NULL)
	{
		CString Key;
		int* i;
		mp_oKeys.GetNextAssoc(pos, Key, (void*&) i);
		delete i;
	}
	mp_oKeys.RemoveAll();
}

void clsCollectionBase::mp_aoCollection_Clear()
{
	
	
	POSITION pos = mp_aoCollection.GetStartPosition();
	while (pos != NULL)
	{
		int lIndex;
		clsItemBase* oObject;
		mp_aoCollection.GetNextAssoc(pos, lIndex, (clsItemBase*&) oObject);
		delete oObject;
	}
	mp_aoCollection.RemoveAll();
}


void clsCollectionBase::m_Add(clsItemBase* r_oObject, CString v_sKey, LONG v_lErr1, LONG v_lErr2, BOOL v_bKeyRequired, LONG v_lKeyError)
{
	int lUpperBounds;
	if (mp_bAddMode == FALSE)
	{
		g_ErrorReport(ERR_ADDMODE_G, "AddMode must be set to TRUE before executing oCollection.m_Add", "cls" + m_sObjectName + "s");
	}
	if (g_IsNumeric(v_sKey) == TRUE)
	{
		g_ErrorReport((SYS_ERRORS)v_lErr1, _T("Invalid ") + m_sObjectName + _T(" object key, key cannot be numeric"), _T("ActiveGanttVCCtl.cls") + m_sObjectName + _T("s.Add"));
		return;
	}
	if (!v_sKey.IsEmpty())
	{
		if (m_bIsKeyUnique(v_sKey) == FALSE)
		{
			g_ErrorReport((SYS_ERRORS)v_lErr1, _T("Key is not unique in ") + m_sObjectName + _T("s collection"), _T("ActiveGanttVCCtl.cls") + m_sObjectName + _T("s.Add"));
			return;
		}
	}
	lUpperBounds = mp_aoCollection.GetCount() + 1;
	r_oObject->mp_lIndex = lUpperBounds;
	mp_aoCollection.SetAt(lUpperBounds, r_oObject);
	if (!v_sKey.IsEmpty())
	{
		int* lIndex;
		lIndex = new int(lUpperBounds);
		mp_oKeys.SetAt(v_sKey, lIndex);
	}
	mp_bAddMode = FALSE;
}

LONG clsCollectionBase::m_lCount()
{
	return mp_aoCollection.GetCount();
}

void clsCollectionBase::m_Clear()
{
	mp_aoCollection_Clear();
	m_ReconstKeys();
}

void clsCollectionBase::m_Remove(CString v_lIndex, LONG v_lErr1, LONG v_lErr2, LONG v_lErr3, LONG v_lErr4)
{

	int lIndex;
	clsItemBase *oItem;
	
	if (!(g_IsNumeric(v_lIndex) == TRUE))
	{
		if (v_lIndex.IsEmpty())
		{
			g_ErrorReport((SYS_ERRORS)v_lErr1, _T("Invalid ") + m_sObjectName + _T(" object key, key cannot be an empty string"), _T("ActiveGanttVCCtl.cls") + m_sObjectName + _T("s.Remove"));
			return;
		}
		lIndex = m_lFindIndexByKey(v_lIndex);
		if (lIndex == -1)
		{
			g_ErrorReport((SYS_ERRORS)v_lErr2, m_sObjectName + _T(" object not found, invalid key"), _T("ActiveGanttVCCtl.cls") + m_sObjectName + _T("s.Remove"));
			return;
		} 
		mp_aoCollection.Lookup(lIndex, (clsItemBase*&) oItem);
		delete oItem;
	}
	else
	{
		lIndex = g_CInt(v_lIndex);
		if (lIndex < 1 || lIndex > mp_aoCollection.GetCount())
		{
			g_ErrorReport((SYS_ERRORS)v_lErr4, m_sObjectName + _T(" object not found, index out of bounds"), _T("ActiveGanttVCCtl.cls") + m_sObjectName + _T("s.Remove"));
			return;
		}
		mp_aoCollection.Lookup(lIndex, (clsItemBase*&) oItem);
		delete oItem;
	}
	mp_bIgnoreKeyChecks = TRUE;
	int lCount = mp_aoCollection.GetCount();
	int i;
	for (i = lIndex + 1; i <= mp_aoCollection.GetCount(); i++)
	{
		clsItemBase* oBuffer;
		mp_aoCollection.Lookup(i, (clsItemBase*&) oBuffer);
		mp_aoCollection.SetAt(i - 1, oBuffer);
	}
	mp_aoCollection.RemoveKey(lCount);
	mp_bIgnoreKeyChecks = FALSE;
	m_ReconstKeys();	
}

clsItemBase* clsCollectionBase::m_oItem(CString v_lIndex, LONG v_lErr1, LONG v_lErr2, LONG v_lErr3, LONG v_lErr4)
{
	int lIndex;
	if (!(g_IsNumeric(v_lIndex) == TRUE))
	{
		if (v_lIndex.IsEmpty())
		{
			g_ErrorReport((SYS_ERRORS)v_lErr1, _T("Invalid ") + m_sObjectName + _T(" object key, key cannot be an empty string"), _T("ActiveGanttVCCtl.cls") + m_sObjectName + _T("s.m_oItem"));
			return NULL;	
		}
		lIndex = m_lFindIndexByKey(v_lIndex);
		if (lIndex == -1) 
		{
			g_ErrorReport((SYS_ERRORS)v_lErr2, m_sObjectName + _T(" object not found, invalid key (\"")  + v_lIndex + _T("\")"), _T("ActiveGanttVCCtl.cls") + m_sObjectName + _T("s.m_oItem"));
			return NULL;
		} 
	}
	else
	{
		lIndex = g_CInt(v_lIndex);
		if (lIndex < 1 || lIndex > mp_aoCollection.GetCount())
		{
			g_ErrorReport((SYS_ERRORS)v_lErr4, m_sObjectName + _T(" object not found, Index out of bounds"), _T("ActiveGanttVCCtl.cls") + m_sObjectName + _T("s.m_oItem"));
			return NULL;
		}
	}
	clsItemBase* oItem;
	mp_aoCollection.Lookup(lIndex, (clsItemBase*&) oItem);

	return oItem;
}

int clsCollectionBase::m_lFindIndexByKey(CString v_sKey)
{
	int* lResult;
	if (mp_oKeys.Lookup(v_sKey, (void*&)lResult))
	{
		return *lResult;
	}
	else
	{
		return -1;
	}


}

void clsCollectionBase::m_ReconstKeys()
{
	int lIndex;
	clsItemBase *oItem;
	mp_oKeys_Clear();
	for (lIndex = 1; lIndex <= mp_aoCollection.GetSize(); lIndex++)
	{
		mp_aoCollection.Lookup(lIndex, (clsItemBase*&) oItem);
		oItem->mp_lIndex = lIndex;
		if (oItem->mp_sKey.IsEmpty() == FALSE)
		{
			int *lIndex2;
			lIndex2 = new int(lIndex);
			mp_oKeys.SetAt(oItem->mp_sKey, lIndex2);
		}
	}
}



clsItemBase* clsCollectionBase::m_oReturnArrayElement(LONG r_lIndex)
{
	clsItemBase* oItem;
	mp_aoCollection.Lookup(r_lIndex, (clsItemBase*&) oItem);
	return oItem;
}

BOOL clsCollectionBase::m_bIsKeyUnique(LPCTSTR v_sKey)
{
	int* lResult;
	if (mp_oKeys.Lookup(v_sKey, (void*&)lResult))
	{
		return FALSE;
	}
	else
	{
		return TRUE;
	}
}

void clsCollectionBase::mp_SetKey(CString *sCurrentKey, CString sNewKey, int ErrNumber)
{
	if (mp_bIgnoreKeyChecks == FALSE)
	{
		if (g_IsNumeric(sNewKey) || (sNewKey != *sCurrentKey && m_bIsKeyUnique(sNewKey) == FALSE))
		{
			g_ErrorReport((SYS_ERRORS)ErrNumber, "Numeric or duplicate " + m_sObjectName + " object key", "ActiveGanttVCCtl.cls" + m_sObjectName + ".Let Key");
			return;
		}
	}
	*sCurrentKey = sNewKey;
	if (mp_bAddMode == FALSE)
	{
		m_ReconstKeys();
	}
}

void clsCollectionBase::SetAddMode(BOOL value)
{
	mp_bAddMode = value;
}