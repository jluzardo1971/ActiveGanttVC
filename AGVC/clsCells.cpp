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
#include "ActiveGanttVCCtl.h"
#include "clsXML.h"
#include "clsCells.h"

IMPLEMENT_DYNCREATE(clsCells, CCmdTarget)

//{4B08BDBB-4D00-480C-9C70-D962A3F33B6F}
static const IID IID_IclsCells = { 0x4B08BDBB, 0x4D00, 0x480C, { 0x9C, 0x70, 0xD9, 0x62, 0xA3, 0xF3, 0x3B, 0x6F} };

//{5856834D-3C39-4DA9-A5E5-09E0A27D14E8}
IMPLEMENT_OLECREATE_FLAGS(clsCells, "AGVC.clsCells", afxRegApartmentThreading, 0x5856834D, 0x3C39, 0x4DA9, 0xA5, 0xE5, 0x09, 0xE0, 0xA2, 0x7D, 0x14, 0xE8)

BEGIN_DISPATCH_MAP(clsCells, CCmdTarget)
	DISP_PROPERTY_EX_ID(clsCells, "Count", 1, odl_GetCount, SetNotSupported, VT_I4) 
	DISP_PROPERTY_PARAM_ID(clsCells, "Item", 2, odl_Item, SetNotSupported, VT_DISPATCH, VTS_BSTR) 
	DISP_FUNCTION_ID(clsCells, "GetXML", 3, odl_GetXML, VT_BSTR, VTS_NONE) 
	DISP_FUNCTION_ID(clsCells, "SetXML", 4, odl_SetXML, VT_EMPTY, VTS_BSTR)
	DISP_DEFVALUE(clsCells, "Item")
END_DISPATCH_MAP()

BEGIN_INTERFACE_MAP(clsCells, CCmdTarget)
	INTERFACE_PART(clsCells, IID_IclsCells, Dispatch)
END_INTERFACE_MAP()

BEGIN_MESSAGE_MAP(clsCells, CCmdTarget)
END_MESSAGE_MAP()

clsCells::clsCells(CActiveGanttVCCtl* oControl, clsRow* oRow)
{
	EnableAutomation();
	AfxOleLockApp();
	mp_oControl = oControl;
	mp_oCollection = new clsCollectionBase(mp_oControl, "Cell");
	mp_oRow = oRow;
}

clsCells::clsCells()
{
	AfxMessageBox(_T("Invalid constructor has been invoked for clsCells. Please inform support@sourcecodestore.com of this error"), MB_OK | MB_ICONSTOP | MB_APPLMODAL);
}

clsCells::~clsCells()
{
	delete mp_oCollection;
	mp_oCollection = NULL;
	AfxOleUnlockApp();
}

void clsCells::OnFinalRelease()
{
	CCmdTarget::OnFinalRelease();
}



LONG clsCells::GetCount(void)
{
	return mp_oCollection->m_lCount();
}

clsCollectionBase* clsCells::GetCollection(void)
{
	return mp_oCollection;
}



void clsCells::Finalize(void)
{
}

clsCell* clsCells::Item(CString Index)
{
	clsCell *oCell;
	oCell = (clsCell*)mp_oCollection->m_oItem(Index, CELLS_ITEM_1, CELLS_ITEM_2, CELLS_ITEM_3, CELLS_ITEM_4);
	return oCell;
}

void clsCells::Add(void)
{
	mp_oCollection->SetAddMode(TRUE);
	clsCell* oCell = new clsCell(mp_oControl, this);
	mp_oCollection->m_Add(oCell, "", CELLS_ADD_1, CELLS_ADD_2, FALSE, CELLS_ADD_3);
	oCell = NULL;
}

void clsCells::Clear(void)
{
	mp_oCollection->m_Clear();
}

void clsCells::Remove(CString Index)
{
	mp_oCollection->m_Remove(Index, CELLS_REMOVE_1, CELLS_REMOVE_2, CELLS_REMOVE_3, CELLS_REMOVE_4);
}

clsRow* clsCells::Row(void)
{
	return mp_oRow;
}

CString clsCells::GetXML(void)
{
	int lIndex;
	clsCell* oCell; 
	clsXML oXML(mp_oControl, "Cells");
	oXML.InitializeWriter();
	for (lIndex = 1; lIndex <= GetCount(); lIndex++)
	{
		oCell = (clsCell*) mp_oCollection->m_oReturnArrayElement(lIndex);
		oXML.WriteObject(oCell->GetXML());
	}
	return oXML.GetXML();
}

void clsCells::SetXML(CString sXML)
{
	int lIndex;
	clsXML oXML(mp_oControl, "Cells");
	oXML.SetXML(sXML);
	oXML.InitializeReader();
	mp_oCollection->m_Clear();
	for (lIndex = 1;lIndex <= oXML.ReadCollectionCount();lIndex++)
	{
		clsCell* oCell = new clsCell(mp_oControl, this);
		oCell->SetXML(oXML.ReadCollectionObject(lIndex));
		mp_oCollection->SetAddMode(TRUE);
		mp_oCollection->m_Add(oCell, oCell->mp_sKey, CELLS_ADD_1, CELLS_ADD_2, FALSE, CELLS_ADD_3);
	}
}