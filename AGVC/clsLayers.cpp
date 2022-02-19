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
#include "clsLayers.h"



IMPLEMENT_DYNCREATE(clsLayers, CCmdTarget)

//{4EF77010-A13E-476F-9328-C907D8D99DB6}
static const IID IID_IclsLayers = { 0x4EF77010, 0xA13E, 0x476F, { 0x93, 0x28, 0xC9, 0x07, 0xD8, 0xD9, 0x9D, 0xB6} };

//{D7E163D0-DE1E-4F59-8EEA-CBF1F3FE0218}
IMPLEMENT_OLECREATE_FLAGS(clsLayers, "AGVC.clsLayers", afxRegApartmentThreading, 0xD7E163D0, 0xDE1E, 0x4F59, 0x8E, 0xEA, 0xCB, 0xF1, 0xF3, 0xFE, 0x02, 0x18)

BEGIN_DISPATCH_MAP(clsLayers, CCmdTarget)
	DISP_PROPERTY_EX_ID(clsLayers, "Count", 1, odl_GetCount, SetNotSupported, VT_I4) 
	DISP_PROPERTY_PARAM_ID(clsLayers, "Item", 2, odl_Item, SetNotSupported, VT_DISPATCH, VTS_BSTR) 
	DISP_FUNCTION_ID(clsLayers, "Add", 3, odl_Add, VT_DISPATCH, VTS_BSTR VTS_BOOL) 
	DISP_FUNCTION_ID(clsLayers, "Clear", 4, odl_Clear, VT_EMPTY, VTS_NONE) 
	DISP_FUNCTION_ID(clsLayers, "Remove", 5, odl_Remove, VT_EMPTY, VTS_BSTR) 
	DISP_FUNCTION_ID(clsLayers, "GetXML", 6, odl_GetXML, VT_BSTR, VTS_NONE) 
	DISP_FUNCTION_ID(clsLayers, "SetXML", 7, odl_SetXML, VT_EMPTY, VTS_BSTR)
	DISP_DEFVALUE(clsLayers, "Item")
END_DISPATCH_MAP()

BEGIN_INTERFACE_MAP(clsLayers, CCmdTarget)
	INTERFACE_PART(clsLayers, IID_IclsLayers, Dispatch)
END_INTERFACE_MAP()

BEGIN_MESSAGE_MAP(clsLayers, CCmdTarget)
END_MESSAGE_MAP()

clsLayers::clsLayers(CActiveGanttVCCtl* oControl)
{
	EnableAutomation();
	AfxOleLockApp();
	mp_oControl = oControl;
	mp_oCollection = new clsCollectionBase(mp_oControl, "Layer");
	mp_oDefaultLayer = new clsLayer(mp_oControl);	
}

clsLayers::clsLayers()
{
	AfxMessageBox(_T("Invalid constructor has been invoked for clsLayers. Please inform support@sourcecodestore.com of this error"), MB_OK | MB_ICONSTOP | MB_APPLMODAL);
}

clsLayers::~clsLayers()
{
	delete mp_oCollection;
	mp_oCollection = NULL;
	delete mp_oDefaultLayer;
	mp_oDefaultLayer = NULL;
	AfxOleUnlockApp();
}

void clsLayers::OnFinalRelease()
{
	CCmdTarget::OnFinalRelease();
}

LONG clsLayers::GetCount(void)
{
	return mp_oCollection->m_lCount();
}

clsCollectionBase* clsLayers::GetCollection(void)
{
	return mp_oCollection;
}



void clsLayers::Finalize(void)
{
}

clsLayer* clsLayers::Item(CString Index)
{
	clsLayer *oLayer;
	oLayer = (clsLayer*)mp_oCollection->m_oItem(Index, LAYERS_ITEM_1, LAYERS_ITEM_2, LAYERS_ITEM_3, LAYERS_ITEM_4);
	return oLayer;
}

clsLayer* clsLayers::FItem(CString Index)
{
	if ( Index == "0" )
	{
		return mp_oDefaultLayer;
	}
	else
	{
		return (clsLayer*) mp_oCollection->m_oItem(Index, LAYERS_ITEM_1, LAYERS_ITEM_2, LAYERS_ITEM_3, LAYERS_ITEM_4);
	}
}

clsLayer* clsLayers::Add(CString Key,BOOL Visible)
{
	mp_oCollection->SetAddMode(TRUE);
	clsLayer* oLayer = new clsLayer(mp_oControl);
	oLayer->SetKey(Key);
	oLayer->SetVisible(Visible);
	mp_oCollection->m_Add(oLayer, Key, LAYERS_ADD_1, LAYERS_ADD_2, FALSE, LAYERS_ADD_3);
	return oLayer;
}

void clsLayers::Clear(void)
{
	mp_oControl->GetTasks()->mp_oCollection->m_CollectionRemoveWhereNot("LayerIndex", "0");
	mp_oControl->SetCurrentLayer("0");
	mp_oCollection->m_Clear();
}

void clsLayers::Remove(CString Index)
{
	CString sRIndex = "";
	CString sRKey = "";
	mp_oCollection->m_GetKeyAndIndex(Index, &sRKey, &sRIndex);
	mp_oControl->GetTasks()->mp_oCollection->m_CollectionRemoveWhere("LayerIndex", sRKey, sRIndex);
	mp_oControl->SetCurrentLayer("0");
	mp_oCollection->m_Remove(Index, LAYERS_REMOVE_1, LAYERS_REMOVE_2, LAYERS_REMOVE_3, LAYERS_REMOVE_4);
}

CString clsLayers::GetXML(void)
{
	LONG lIndex;
	clsLayer* oLayer; 
	clsXML oXML(mp_oControl, "Layers");
	oXML.InitializeWriter();
	for (lIndex = 1; lIndex <= GetCount(); lIndex++)
	{
		oLayer = (clsLayer*) mp_oCollection->m_oReturnArrayElement(lIndex);
		oXML.WriteObject(oLayer->GetXML());
	}
	return oXML.GetXML();
}

void clsLayers::SetXML(CString sXML)
{
	LONG lIndex;
	clsXML oXML(mp_oControl, "Layers");
	oXML.SetXML(sXML);
	oXML.InitializeReader();
	mp_oCollection->m_Clear();
	for (lIndex = 1;lIndex <= oXML.ReadCollectionCount();lIndex++)
	{
		clsLayer* oLayer = new clsLayer(mp_oControl);
		oLayer->SetXML(oXML.ReadCollectionObject(lIndex));
		mp_oCollection->SetAddMode(TRUE);
		mp_oCollection->m_Add(oLayer, oLayer->mp_sKey, LAYERS_ADD_1, LAYERS_ADD_2, FALSE, LAYERS_ADD_3);
	}
}