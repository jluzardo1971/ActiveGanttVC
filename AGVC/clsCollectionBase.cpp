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
#include "ActiveGanttVC.h"
#include "clsCollectionBase.h"
#include "ActiveGanttVCCtl.h"

#ifdef _DEBUG
#undef THIS_FILE
static char THIS_FILE[]=__FILE__;
#define new DEBUG_NEW
#endif


clsCollectionBase::clsCollectionBase(CActiveGanttVCCtl* oControl, CString sObjectName)
{
	mp_oControl = oControl;
	m_sObjectName = sObjectName;
	mp_bDescending = FALSE;
	mp_bIgnoreKeyChecks = FALSE;
	mp_bSortCells = FALSE;
	mp_lCellIndex = 0;
	mp_sPropertyName = "";
	mp_ySortType = ES_STRING;
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
		i = NULL;
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
		oObject = NULL;
	}
	mp_aoCollection.RemoveAll();
}


void clsCollectionBase::m_Add(clsItemBase* r_oObject, CString v_sKey, LONG v_lErr1, LONG v_lErr2, BOOL v_bKeyRequired, LONG v_lKeyError)
{
	int lUpperBounds;
	if (mp_bAddMode == FALSE)
	{
		mp_oControl->mp_ErrorReport(ERR_ADDMODE_G, "AddMode must be set to TRUE before executing oCollection.m_Add", "cls" + m_sObjectName + "s");
	}
	if (g_IsNumeric(v_sKey))
	{
		mp_oControl->mp_ErrorReport((SYS_ERRORS)v_lErr1, _T("Invalid ") + m_sObjectName + _T(" object key, key cannot be numeric"), _T("ActiveGanttVCCtl.cls") + m_sObjectName + _T("s.Add"));
		return;
	}
	if (!v_sKey.IsEmpty())
	{
		if (m_bIsKeyUnique(v_sKey) == FALSE)
		{
			mp_oControl->mp_ErrorReport((SYS_ERRORS)v_lErr1, _T("Key is not unique in ") + m_sObjectName + _T("s collection"), _T("ActiveGanttVCCtl.cls") + m_sObjectName + _T("s.Add"));
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
	
	if (!g_IsNumeric(v_lIndex))
	{
		if (v_lIndex.IsEmpty())
		{
			mp_oControl->mp_ErrorReport((SYS_ERRORS)v_lErr1, _T("Invalid ") + m_sObjectName + _T(" object key, key cannot be an empty string"), _T("ActiveGanttVCCtl.cls") + m_sObjectName + _T("s.Remove"));
			return;
		}
		lIndex = m_lFindIndexByKey(v_lIndex);
		if (lIndex == -1)
		{
			mp_oControl->mp_ErrorReport((SYS_ERRORS)v_lErr2, m_sObjectName + _T(" object not found, invalid key"), _T("ActiveGanttVCCtl.cls") + m_sObjectName + _T("s.Remove"));
			return;
		} 
		mp_aoCollection.Lookup(lIndex, (clsItemBase*&) oItem);
		delete oItem;
		oItem = NULL;
	}
	else
	{
		lIndex = CInt32(v_lIndex);
		if (lIndex < 1 || lIndex > mp_aoCollection.GetCount())
		{
			mp_oControl->mp_ErrorReport((SYS_ERRORS)v_lErr4, m_sObjectName + _T(" object not found, index out of bounds"), _T("ActiveGanttVCCtl.cls") + m_sObjectName + _T("s.Remove"));
			return;
		}
		mp_aoCollection.Lookup(lIndex, (clsItemBase*&) oItem);
		delete oItem;
		oItem = NULL;
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
	if (!g_IsNumeric(v_lIndex))
	{
		if (v_lIndex.IsEmpty())
		{
			mp_oControl->mp_ErrorReport((SYS_ERRORS)v_lErr1, _T("Invalid ") + m_sObjectName + _T(" object key, key cannot be an empty string"), _T("ActiveGanttVCCtl.cls") + m_sObjectName + _T("s.m_oItem"));
			return NULL;	
		}
		lIndex = m_lFindIndexByKey(v_lIndex);
		if (lIndex == -1) 
		{
			mp_oControl->mp_ErrorReport((SYS_ERRORS)v_lErr2, m_sObjectName + _T(" object not found, invalid key (\"")  + v_lIndex + _T("\")"), _T("ActiveGanttVCCtl.cls") + m_sObjectName + _T("s.m_oItem"));
			return NULL;
		} 
	}
	else
	{
		lIndex = CInt32(v_lIndex);
		if (lIndex < 1 || lIndex > mp_aoCollection.GetCount())
		{
			mp_oControl->mp_ErrorReport((SYS_ERRORS)v_lErr4, m_sObjectName + _T(" object not found, Index out of bounds"), _T("ActiveGanttVCCtl.cls") + m_sObjectName + _T("s.m_oItem"));
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
	clsItemBase* oItem = NULL;
	mp_aoCollection.Lookup(r_lIndex, (clsItemBase*&) oItem);
	if (oItem == NULL)
	{
		mp_oControl->mp_ErrorReport((SYS_ERRORS)ERR_RETARRELEMINDEX_G, _T(" object not found, Index out of bounds. Index: ") + CStr(r_lIndex), _T("ActiveGanttVCCtl.cls") + m_sObjectName + _T("s.m_oReturnArrayElement"));
	}
	return oItem;
}

clsItemBase* clsCollectionBase::m_oReturnArrayElementKey(CString v_sKey)
{
	LONG lIndex;
	if (g_IsNumeric(v_sKey))
	{
		lIndex = CInt32(v_sKey);
		clsItemBase* oItem = NULL;
		mp_aoCollection.Lookup(lIndex, (clsItemBase*&) oItem);
		if (oItem == NULL)
		{
			mp_oControl->mp_ErrorReport((SYS_ERRORS)ERR_RETARRELEMKEY_G, _T("Key not found, Key: ") + v_sKey + _T(". Case 1."), _T("ActiveGanttVCCtl.cls") + m_sObjectName + _T("s.m_oReturnArrayElementKey"));
		}
		return oItem;
	}
	else
	{
		lIndex = m_lFindIndexByKey(v_sKey);
		if (lIndex != -1)
		{
			clsItemBase* oItem = NULL;
			mp_aoCollection.Lookup(lIndex, (clsItemBase*&) oItem);
			if (oItem == NULL)
			{
				mp_oControl->mp_ErrorReport((SYS_ERRORS)ERR_RETARRELEMKEY_G, _T("Key not found, Key: ") + v_sKey + _T(". Case 2."), _T("ActiveGanttVCCtl.cls") + m_sObjectName + _T("s.m_oReturnArrayElementKey"));
			}
			return oItem;
		}
		else
		{
			mp_oControl->mp_ErrorReport((SYS_ERRORS)ERR_RETARRELEMKEY_G, _T("Key not found, Key: -1. Case 3."), _T("ActiveGanttVCCtl.cls") + m_sObjectName + _T("s.m_oReturnArrayElementKey"));
			return NULL;
		}

	}

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

LONG clsCollectionBase::m_lCopyAndMoveItems(LONG v_lOriginIndex, LONG v_lDestinationIndex)
{
	clsItemBase* Buffer;
    int lIndex;
    mp_bIgnoreKeyChecks = TRUE;
	mp_aoCollection.Lookup(v_lOriginIndex, (clsItemBase*&) Buffer);
    if (v_lOriginIndex > v_lDestinationIndex)
	{
		for (lIndex = v_lOriginIndex; lIndex >= v_lDestinationIndex + 1; lIndex--)
		{
			clsItemBase* oBuffer1;
			mp_aoCollection.Lookup(lIndex - 1, oBuffer1);
            mp_aoCollection[lIndex] = oBuffer1;
        }
	}
    else
	{
		for (lIndex = v_lOriginIndex; lIndex <= v_lDestinationIndex - 1; lIndex++)
		{
			clsItemBase* oBuffer1;
			mp_aoCollection.Lookup(lIndex + 1, oBuffer1);
            mp_aoCollection[lIndex] = oBuffer1;
        }
    }
    mp_aoCollection[v_lDestinationIndex] = Buffer;
    mp_bIgnoreKeyChecks = FALSE;
    m_ReconstKeys();
	return lIndex;
}



BOOL clsCollectionBase::m_bDoesKeyExist(LPCTSTR v_sKey)
{
	int* lResult;
	if (mp_oKeys.Lookup(v_sKey, (void*&)lResult))
	{
		return TRUE;
	}
	else
	{
		return FALSE;
	}

}

CString clsCollectionBase::m_ReturnOptIndexValue(const VARIANT &vParam, CString sDefault)
{
	CString sBuff;
	if (vParam.vt == VT_BSTR)
	{
		sBuff = vParam.bstrVal;
		return sBuff;
	}
	else if (vParam.vt == (VT_BYREF | VT_BSTR))
	{
		return *vParam.pbstrVal;
	}
	else if (vParam.vt == VT_UI1)
	{
		sBuff = CStr((LONG)vParam.bVal);
		return sBuff;
	}
	else if (vParam.vt == (VT_BYREF | VT_UI1))
	{
		sBuff = CStr((LONG)*vParam.pbVal);
		return sBuff;
	}
	else if (vParam.vt == VT_I2)
	{
		sBuff = CStr((LONG)vParam.iVal);
		return sBuff;
	}
	else if (vParam.vt == (VT_BYREF | VT_I2))
	{
		sBuff = CStr((LONG)*vParam.piVal);
		return sBuff;
	}
	else if (vParam.vt == VT_I4)
	{
		sBuff = CStr(vParam.lVal);
		return sBuff;
	}
	else if (vParam.vt == (VT_BYREF | VT_I4))
	{
		sBuff = CStr(*vParam.plVal);
		return sBuff;
	}
	else if (vParam.vt == VT_ERROR)
	{
		return sDefault;
	}
	else if (vParam.vt == (VT_VARIANT | VT_BYREF))
	{
		VARIANT* varVal;
		varVal = vParam.pvarVal;
		sBuff = varVal->bstrVal;
		return sBuff;
	}
	else
	{
		mp_oControl->mp_ErrorReport(C_ROIV_UNK_TYPE, "Unknown Type, returning null", "cls" + m_sObjectName + "s.m_ReturnOptIndexValue");
		sBuff = _T("");
		return sBuff;
	}
	

}

BOOL clsCollectionBase::m_ReturnOptBoolValue(const VARIANT &vParam, BOOL bDefault)
{
	if (vParam.vt == VT_BOOL)
	{
		return abs(vParam.boolVal);
	}
	if (vParam.vt == (VT_BYREF | VT_BOOL))
	{
		return abs(*vParam.pboolVal);
	}
	else if (vParam.vt == VT_ERROR)
	{
		return bDefault;
	}
	else
	{
		mp_oControl->mp_ErrorReport(C_ROBV_UNK_TYPE, "Unknown Type, returning null", "cls" + m_sObjectName + "s.m_ReturnOptBoolValue");
		return FALSE;
	}

}

CString clsCollectionBase::m_ReturnOptStrValue(const VARIANT &vParam, CString sDefault)
{
	CString sBuff;
	if (vParam.vt == VT_BSTR)
	{
		sBuff = vParam.bstrVal;
		return sBuff;
	}
	else if (vParam.vt == (VT_BYREF | VT_BSTR))
	{
		return *vParam.pbstrVal;
	}
	else if (vParam.vt == VT_I4)
	{
		sBuff = CStr(vParam.lVal);
		return sBuff;
	}
	else if (vParam.vt == (VT_BYREF | VT_I4))
	{
		sBuff = CStr((LONG)*vParam.plVal);
		return sBuff;
	}
	else if (vParam.vt == VT_I2)
	{
		sBuff = CStr((LONG)vParam.iVal);
		return sBuff;
	}
	else if (vParam.vt == (VT_BYREF | VT_I2))
	{
		sBuff = CStr((LONG)*vParam.piVal);
		return sBuff;
	}
	else if (vParam.vt == VT_UI1)
	{
		sBuff = CStr((LONG)vParam.bVal);
		return sBuff;
	}
	else if (vParam.vt == (VT_BYREF | VT_UI1))
	{
		sBuff = CStr((LONG)*vParam.pbVal);
		return sBuff;
	}
	else if (vParam.vt == VT_ERROR)
	{
		return sDefault;
	}
	else if (vParam.vt == (VT_VARIANT | VT_BYREF))
	{
		VARIANT* varVal;
		varVal = vParam.pvarVal;
		sBuff = varVal->bstrVal;
		return sBuff;
	}
	else if (vParam.vt == VT_DISPATCH)
	{
		mp_oControl->mp_ErrorReport(C_ROSV_DISP_TYPE, "Dispatch Type, returning null", "cls" + m_sObjectName + "s.m_ReturnOptStrValue");
		return _T("");
	}
	else
	{
		mp_oControl->mp_ErrorReport(C_ROSV_UNK_TYPE, "Unknown Type, returning null", "cls" + m_sObjectName + "s.m_ReturnOptStrValue");
		return _T("");
	}

}

LONG clsCollectionBase::m_ReturnOptLongValue(const VARIANT &vParam, LONG lDefault)
{
	if (vParam.vt == VT_I4)
	{
		return vParam.lVal;
	}
	else if (vParam.vt == VT_I2)
	{
		return vParam.iVal;
	}
	else if (vParam.vt == VT_ERROR)
	{
		return lDefault;
	}
	else
	{
		mp_oControl->mp_ErrorReport(C_ROLV_UNK_TYPE, "Unknown Type, returning null", "cls" + m_sObjectName + "s.m_ReturnOptLongValue");
		return 0;
	}

}

void clsCollectionBase::m_Sort(CString sPropertyName, BOOL bDescending, E_SORTTYPE SortType, int StartIndex, int EndIndex)
{

	if (m_lCount() == 0)
	{
		return;
	}
	else
	{
		DISPID oPropertyID = NULL;
		clsItemBase* oItem;
		oItem = (clsItemBase*) m_oItem("1");
		LPDISPATCH lpdItem;
		lpdItem = oItem->GetIDispatch(TRUE);
		CString propName = sPropertyName;
		LPOLESTR sPropertyName1 = propName.AllocSysString();
		lpdItem->GetIDsOfNames(IID_NULL, &sPropertyName1 , 1, LOCALE_USER_DEFAULT, &oPropertyID);
		if (oPropertyID == -1)
		{
			mp_oControl->mp_ErrorReport(C_INVALID_PROP_NAME_1, "Invalid property name", "Sort");
			return;
		}

	}
	CMapIntToItemBase aTempMap;
    mp_sPropertyName = sPropertyName;
    mp_bDescending = bDescending;
    mp_ySortType = SortType;
    mp_bSortCells = FALSE;
	//aTempArray.SetSize(EndIndex);
    mp_Sort(&mp_aoCollection, &aTempMap, StartIndex, EndIndex);
    m_ReconstKeys();
}

void clsCollectionBase::m_SortCells(int CellIndex, CString sPropertyName, BOOL bDescending, E_SORTTYPE SortType, int StartIndex, int EndIndex)
{
	if (m_lCount() == 0)
	{
		return;
	}
	else
	{
		DISPID oPropertyID = NULL;
		clsCell* oItem;
		clsRow* oRow;
		oRow = (clsRow*) m_oItem("1");
		if (oRow->GetCells()->mp_oCollection->m_lCount() == 0)
		{
			return;
		}
		else
		{
			oItem = (clsCell*) oRow->GetCells()->mp_oCollection->m_oItem("1");

			LPDISPATCH lpdItem;
			lpdItem = oItem->GetIDispatch(TRUE);
			CString propName = sPropertyName;
			LPOLESTR sPropertyName1 = propName.AllocSysString();
			lpdItem->GetIDsOfNames(IID_NULL, &sPropertyName1 , 1, LOCALE_USER_DEFAULT, &oPropertyID);
			if (oPropertyID == -1)
			{
				mp_oControl->mp_ErrorReport(C_INVALID_PROP_NAME_2, "Invalid property name", "Sort Cells");
				return;
			}
		}

	}
	CMapIntToItemBase aTempMap;
    mp_sPropertyName = sPropertyName;
    mp_bDescending = bDescending;
    mp_ySortType = SortType;
    mp_bSortCells = TRUE;
	mp_lCellIndex = CellIndex;
	//aTempArray.SetSize(EndIndex);
    mp_Sort(&mp_aoCollection, &aTempMap, StartIndex, EndIndex);
    m_ReconstKeys();

}

void clsCollectionBase::mp_Sort(CMapIntToItemBase* aSortMap, CMapIntToItemBase* aTempMap, int first, int last)
{
	int lArrayMBound;
    int lArrayLBound;
    int lArrayUBound;
    if (first == -1)
	{
        lArrayLBound = 1;
	}
    else
	{
        lArrayLBound = first;
    }
	if (last == -1)
	{
        lArrayUBound = aSortMap->GetSize();
	}
    else
	{
        lArrayUBound = last;
    }
    if (lArrayUBound > lArrayLBound)
	{
        lArrayMBound = (lArrayUBound + lArrayLBound) / 2;
        mp_Sort(aSortMap, aTempMap, lArrayLBound, lArrayMBound);
        mp_Sort(aSortMap, aTempMap, lArrayMBound + 1, lArrayUBound);
        mp_Merge(aSortMap, aTempMap, lArrayLBound, lArrayMBound + 1, lArrayUBound);
    }

}

clsItemBase* clsCollectionBase::mp_GetAt(CMapIntToItemBase* oMap, int lIndex)
{
	clsItemBase* oItemBase;
	if (oMap->Lookup(lIndex, (clsItemBase*&) oItemBase) == TRUE)
	{
		return oItemBase;
	}
	else
	{
		return NULL;
	}
}

void clsCollectionBase::mp_Merge(CMapIntToItemBase* aSortMap, CMapIntToItemBase* aTempMap, int first, int mid, int last)
{
	int i;
    int iLow;
    int nNumElements;
    int iTempPos;
    iLow = mid - 1;
    iTempPos = first;
    nNumElements = last - first + 1;
    while (first <= iLow && mid <= last)
	{
            if (mp_bCompareTypes(mp_GetAt(aSortMap, first), mp_GetAt(aSortMap, mid)))
			{
                aTempMap->SetAt(iTempPos, mp_GetAt(aSortMap, first));
                iTempPos = iTempPos + 1;
                first = first + 1;
			}
            else
			{
                aTempMap->SetAt(iTempPos, mp_GetAt(aSortMap, mid));
                iTempPos = iTempPos + 1;
                mid = mid + 1;
            }
    }
    while (first <= iLow)
	{
        aTempMap->SetAt(iTempPos, mp_GetAt(aSortMap, first));
        first = first + 1;
        iTempPos = iTempPos + 1;
    }
    while (mid <= last)
	{
        aTempMap->SetAt(iTempPos, mp_GetAt(aSortMap, mid));
        mid = mid + 1;
        iTempPos = iTempPos + 1;
    }
	for (i = 0; i <= nNumElements - 1; i++)
	{
        aSortMap->SetAt(last, mp_GetAt(aTempMap, last));
        last = last - 1;
    }

}

BOOL clsCollectionBase::mp_bCompareTypes(clsItemBase *oClass1, clsItemBase *oClass2)
{
	BOOL lResult;
	DISPID oPropertyID = NULL;
	LPDISPATCH lpdClass1;
	LPDISPATCH lpdClass2;
	VARIANT FAR *pVarResult1 = new VARIANT;
	VARIANT FAR *pVarResult2 = new VARIANT;
	DISPPARAMS dispparamsNoArgs = {NULL, NULL, 0, 0};

	ASSERT(oClass1 != NULL);
	ASSERT(oClass2 != NULL);


	lpdClass1 = oClass1->GetIDispatch(TRUE);
	lpdClass2 = oClass2->GetIDispatch(TRUE);

	if (mp_bSortCells == FALSE)
	{
		CString propName = mp_sPropertyName;
		LPOLESTR sPropertyName = propName.AllocSysString();

		lpdClass1->GetIDsOfNames(IID_NULL, &sPropertyName , 1, LOCALE_USER_DEFAULT, &oPropertyID);
		lpdClass1->Invoke(oPropertyID, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYGET, &dispparamsNoArgs, pVarResult1, NULL, NULL);

		lpdClass2->GetIDsOfNames(IID_NULL, &sPropertyName , 1, LOCALE_USER_DEFAULT, &oPropertyID);
		lpdClass2->Invoke(oPropertyID, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYGET, &dispparamsNoArgs, pVarResult2, NULL, NULL);
	}
	else
	{
		LPDISPATCH oCell1;
		LPDISPATCH oCell2;

		CString sIndex;

		sIndex.Format(_T("%d"), mp_lCellIndex);


		clsRow* oRow1;
		clsRow* oRow2;

		clsCell* oCellObj1;
		clsCell* oCellObj2;

		oRow1 = (clsRow*) oClass1;
		oRow2 = (clsRow*) oClass2;

		oCellObj1 = (clsCell*) oRow1->GetCells()->mp_oCollection->m_oItem(sIndex);
		oCellObj2 = (clsCell*) oRow2->GetCells()->mp_oCollection->m_oItem(sIndex);

		oCell1 = oCellObj1->GetIDispatch(TRUE);
		oCell2 = oCellObj2->GetIDispatch(TRUE);

		CString propName = mp_sPropertyName;
		LPOLESTR sPropertyName = propName.AllocSysString();

		oCell1->GetIDsOfNames(IID_NULL, &sPropertyName , 1, LOCALE_USER_DEFAULT, &oPropertyID);
		oCell1->Invoke(oPropertyID, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYGET, &dispparamsNoArgs, pVarResult1, NULL, NULL);

		oCell2->GetIDsOfNames(IID_NULL, &sPropertyName , 1, LOCALE_USER_DEFAULT, &oPropertyID);
		oCell2->Invoke(oPropertyID, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYGET, &dispparamsNoArgs, pVarResult2, NULL, NULL);
	}
	

	CString String1;
	CString String2;
	int Long1;
	int Long2;
	COleDateTime Date1;
	COleDateTime Date2;

	switch (pVarResult1->vt)
	{
	case VT_BSTR:
		switch (mp_ySortType)
		{
		case ES_STRING:
			String1 = pVarResult1->bstrVal;
			String2 = pVarResult2->bstrVal;
			lResult = mp_CompareTypesAux(&String1, &String2);
			break;
		case ES_NUMERIC:
			String1 = pVarResult1->bstrVal;
			String2 = pVarResult2->bstrVal;
			Long1 = CInt32(String1);
			Long2 = CInt32(String2);
			lResult = mp_CompareTypesAux(&Long1, &Long2);
			break;
		case ES_DATE:
			String1 = pVarResult1->bstrVal;
			String2 = pVarResult2->bstrVal;
			Date1.ParseDateTime(String1);
			Date2.ParseDateTime(String2);
			lResult = mp_CompareTypesAux(&Date1, &Date2);
			break;
		}
		break;
	case VT_I4:
		Long1 = pVarResult1->lVal;
		Long2 = pVarResult2->lVal;
		lResult = mp_CompareTypesAux(&Long1, &Long2);
		break;
	case VT_I2:
		Long1 = pVarResult1->iVal;
		Long2 = pVarResult2->iVal;
		lResult = mp_CompareTypesAux(&Long1, &Long2);
		break;
	case VT_DATE:
		Date1 = pVarResult1->date;
		Date2 = pVarResult2->date;
		lResult = mp_CompareTypesAux(&Date1, &Date2);
		break;

	}
	delete(pVarResult1);
	pVarResult1 = NULL;
	delete(pVarResult2);
	pVarResult2 = NULL;
	return lResult;
}


template <class T>
BOOL clsCollectionBase::mp_CompareTypesAux(T *Value1, T *Value2)
{
	BOOL bResult;
	if (mp_bDescending == FALSE)
	{
		if (*Value1 <= *Value2)
		{
			bResult = TRUE;	
		}
		else
		{
			bResult = FALSE;
		}
	}
	else
	{
		if (*Value1 >= *Value2)
		{
			bResult = TRUE;
		}
		else
		{
			bResult = FALSE;
		}
	}
	return bResult;
}

LONG clsCollectionBase::m_lReturnIndex(CString v_sIndex, BOOL bIncludeDefault)
{
	LONG lIndex;
    LONG lReturn;
    if (g_IsNumeric(v_sIndex))
	{
        lIndex = CInt32(v_sIndex);
        if (bIncludeDefault == TRUE)
		{
            if (lIndex >= 0 && lIndex <= m_lCount())
			{
                lReturn = lIndex;
			}
            else
			{
                lReturn = -1;
            }
		}
        else
		{
            if (lIndex >= 1 && lIndex <= m_lCount())
			{
                lReturn = lIndex;
			}
            else
			{
                lReturn = -1;
            }
        }
	}
    else
	{
        lReturn = m_lFindIndexByKey(v_sIndex);
    }
    return lReturn;
}

void clsCollectionBase::mp_SetKey(CString *sCurrentKey, CString sNewKey, int ErrNumber)
{
	if (mp_bIgnoreKeyChecks == FALSE)
	{
		if (g_IsNumeric(sNewKey) || (sNewKey != *sCurrentKey && m_bIsKeyUnique(sNewKey) == FALSE))
		{
			mp_oControl->mp_ErrorReport((SYS_ERRORS)ErrNumber, "Numeric or duplicate " + m_sObjectName + " object key", "ActiveGanttVCCtl.cls" + m_sObjectName + ".Let Key");
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

void clsCollectionBase::m_GetKeyAndIndex(CString sIndex, CString *sKey, CString *sReturnIndex)
{
	clsItemBase* oObject;
	oObject = (clsItemBase*) m_oItem(sIndex, GETINDEXANDKEY_ITEM1, GETINDEXANDKEY_ITEM2, GETINDEXANDKEY_ITEM3, GETINDEXANDKEY_ITEM4);
	if (oObject->mp_sKey != "")
	{
		*sKey = oObject->mp_sKey;
	}
	else
	{
		*sKey = CStr(oObject->mp_lIndex);
	}
	*sReturnIndex = CStr(oObject->mp_lIndex);
}

void clsCollectionBase::m_CollectionRemoveWhere(CString sPropertyName, CString sKey, CString sIndex)
{

	int lIndex;
	CCmdTarget* oObject;
	CString sPropertyValue;
	for (lIndex = m_lCount();lIndex >= 1 ;lIndex--) 
	{
		oObject = (CCmdTarget*) m_oReturnArrayElement(lIndex);
		sPropertyValue = CallByName(oObject, sPropertyName);
		if (sPropertyValue == sKey || sPropertyValue == sIndex)
		{
			m_Remove(CStr((LONG)lIndex), ERR_COLLREMWHERE_1_G, ERR_COLLREMWHERE_2_G, ERR_COLLREMWHERE_3_G, ERR_COLLREMWHERE_4_G);
		}
	}

}

void clsCollectionBase::m_CollectionRemoveWhereNot(CString sPropertyName, CString sValue)
{
	int lIndex;
	CCmdTarget* oObject;
	CString sPropertyValue;
	CString sIndex;
	for (lIndex = m_lCount();lIndex >= 1 ;lIndex--) 
	{
		oObject = (CCmdTarget*) m_oReturnArrayElement(lIndex);
		sPropertyValue = CallByName(oObject, sPropertyName);
		if (sPropertyValue != sValue)
		{
			sIndex = CStr((LONG)lIndex);
			m_Remove(sIndex, ERR_COLLREMWHERENOT_1_G, ERR_COLLREMWHERENOT_2_G, ERR_COLLREMWHERENOT_3_G, ERR_COLLREMWHERENOT_4_G);
		}
	}
}

void clsCollectionBase::m_CollectionChange(CString sPropertyName, CString sKey, CString sIndex, CString sNewValue)
{
	int lIndex;
	CCmdTarget* oObject;
	CString sPropertyValue;
	for (lIndex = 1;lIndex <= m_lCount();lIndex++) 
	{
		oObject = (CCmdTarget*) m_oReturnArrayElement(lIndex);
		sPropertyValue = CallByName(oObject, sPropertyName);
		if (sPropertyValue == sKey || sPropertyValue == sIndex)
		{
			CallByName(oObject, sPropertyName, sNewValue);
		}
	}
}

void clsCollectionBase::m_CollectionChangeAll(CString sPropertyName, CString sNewValue)
{
	int lIndex;
	CCmdTarget* oObject;
	for (lIndex = 1;lIndex <= m_lCount();lIndex++) 
	{
		oObject = (CCmdTarget*) m_oReturnArrayElement(lIndex);
		CallByName(oObject, sPropertyName, sNewValue);
	}
}

CString clsCollectionBase::CallByName(CCmdTarget* oObject, CString sPropertyName)
{
	LPDISPATCH lpdObjectDispatch = oObject->GetIDispatch(TRUE);
	OLECHAR FAR* pPropertyName = sPropertyName.AllocSysString();
	DISPID oPropertyID = NULL;
	VARIANT pVarResult;
	VariantInit(&pVarResult);
	CString sReturn;
	DISPPARAMS dispparamsNoArgs = {NULL, NULL, 0, 0};
	lpdObjectDispatch->GetIDsOfNames(IID_NULL, &pPropertyName, 1, LOCALE_USER_DEFAULT, &oPropertyID);
	::SysFreeString(pPropertyName);
	lpdObjectDispatch->Invoke(oPropertyID, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
	sReturn = pVarResult.bstrVal;
	return sReturn;
}

void clsCollectionBase::CallByName(CCmdTarget* oObject, CString sPropertyName, CString sNewValue)
{
	LPDISPATCH lpdObjectDispatch = oObject->GetIDispatch(TRUE);
	LPOLESTR pPropertyName = sPropertyName.AllocSysString();
	DISPID oPropertyID = NULL;
	VARIANT pVarResult;
	VariantInit(&pVarResult);
	CString sReturn;
	DISPPARAMS dispparamsArgs;
	dispparamsArgs.rgvarg = new VARIANT[1]; 
	dispparamsArgs.rgvarg[0].vt = VT_BSTR;
	dispparamsArgs.rgvarg[0].bstrVal = sNewValue.AllocSysString();
	dispparamsArgs.cArgs = 1;
	dispparamsArgs.cNamedArgs = 0;
	dispparamsArgs.rgdispidNamedArgs = 0;
	lpdObjectDispatch->GetIDsOfNames(IID_NULL, &pPropertyName, 1, LOCALE_USER_DEFAULT, &oPropertyID);
	::SysFreeString(pPropertyName);
	lpdObjectDispatch->Invoke(oPropertyID, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYPUT, &dispparamsArgs, &pVarResult, NULL, NULL);
	delete dispparamsArgs.rgvarg;
	dispparamsArgs.rgvarg = NULL;
}