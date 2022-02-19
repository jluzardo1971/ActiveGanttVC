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
#include "clsStyles.h"



IMPLEMENT_DYNCREATE(clsStyles, CCmdTarget)

//{B954B2C1-D946-4C4D-8DC6-09911E14C520}
static const IID IID_IclsStyles = { 0xB954B2C1, 0xD946, 0x4C4D, { 0x8D, 0xC6, 0x09, 0x91, 0x1E, 0x14, 0xC5, 0x20} };

//{CB84968A-26E2-4F55-8148-40821DC58D98}
IMPLEMENT_OLECREATE_FLAGS(clsStyles, "AGVC.clsStyles", afxRegApartmentThreading, 0xCB84968A, 0x26E2, 0x4F55, 0x81, 0x48, 0x40, 0x82, 0x1D, 0xC5, 0x8D, 0x98)

BEGIN_DISPATCH_MAP(clsStyles, CCmdTarget)
	DISP_PROPERTY_EX_ID(clsStyles, "Count", 1, odl_GetCount, SetNotSupported, VT_I4) 
	DISP_PROPERTY_PARAM_ID(clsStyles, "Item", 2, odl_Item, SetNotSupported, VT_DISPATCH, VTS_BSTR) 
	DISP_FUNCTION_ID(clsStyles, "Add", 3, odl_Add, VT_DISPATCH, VTS_BSTR) 
	DISP_FUNCTION_ID(clsStyles, "Clear", 4, odl_Clear, VT_EMPTY, VTS_NONE) 
	DISP_FUNCTION_ID(clsStyles, "Remove", 5, odl_Remove, VT_EMPTY, VTS_BSTR) 
	DISP_FUNCTION_ID(clsStyles, "GetXML", 6, odl_GetXML, VT_BSTR, VTS_NONE) 
	DISP_FUNCTION_ID(clsStyles, "SetXML", 7, odl_SetXML, VT_EMPTY, VTS_BSTR)
	DISP_FUNCTION_ID(clsStyles, "ContainsKey", 8, odl_ContainsKey, VT_BOOL, VTS_BSTR)
	DISP_DEFVALUE(clsStyles, "Item")
END_DISPATCH_MAP()

BEGIN_INTERFACE_MAP(clsStyles, CCmdTarget)
	INTERFACE_PART(clsStyles, IID_IclsStyles, Dispatch)
END_INTERFACE_MAP()

BEGIN_MESSAGE_MAP(clsStyles, CCmdTarget)
END_MESSAGE_MAP()

clsStyles::clsStyles()
{
	AfxMessageBox(_T("Invalid constructor has been invoked for clsStyles. Please inform support@sourcecodestore.com of this error"), MB_OK | MB_ICONSTOP | MB_APPLMODAL);
}

clsStyles::clsStyles(CActiveGanttVCCtl* oControl)
{
	EnableAutomation();
	AfxOleLockApp();
	mp_oControl = oControl;
	mp_oCollection = new clsCollectionBase(mp_oControl, "Style");

	mp_oDefaultControlStyle = new clsStyle(mp_oControl);
	mp_oDefaultColumnStyle = new clsStyle(mp_oControl);
	mp_oDefaultRowStyle = new clsStyle(mp_oControl);
	mp_oDefaultCellStyle = new clsStyle(mp_oControl);
	mp_oDefaultNodeStyle = new clsStyle(mp_oControl);
	mp_oDefaultSplitterStyle = new clsStyle(mp_oControl);
	mp_oDefaultTaskStyle = new clsStyle(mp_oControl);
	mp_oDefaultPercentageStyle = new clsStyle(mp_oControl);
	mp_oDefaultPredecessorStyle = new clsStyle(mp_oControl);
	mp_oDefaultTimeBlockStyle = new clsStyle(mp_oControl);
	mp_oDefaultProgressLineStyle = new clsStyle(mp_oControl);
	
	mp_oDefaultClientAreaStyle = new clsStyle(mp_oControl);
	mp_oDefaultTimeLineStyle = new clsStyle(mp_oControl);
	mp_oDefaultTierStyle = new clsStyle(mp_oControl);
	mp_oDefaultTickMarkAreaStyle = new clsStyle(mp_oControl);

	mp_oDefaultScrollBarStyle = new clsStyle(mp_oControl);
	mp_oDefaultSBSeparatorStyle = new clsStyle(mp_oControl);
	mp_oDefaultSBNormalStyle = new clsStyle(mp_oControl);
	mp_oDefaultSBPressedStyle = new clsStyle(mp_oControl);
	mp_oDefaultSBDisabledStyle = new clsStyle(mp_oControl);
    mp_ResetDefaultStyles();

}

void clsStyles::mp_ResetDefaultStyles(void)
{

	// Control
	mp_oDefaultControlStyle->Clear();
	mp_oDefaultControlStyle->SetAppearance(SA_SUNKEN);
	mp_oDefaultControlStyle->SetBackColor(Color::White);

    mp_oDefaultColumnStyle->Clear();
	mp_oDefaultColumnStyle->SetAppearance((LONG)SA_RAISED);
	mp_oDefaultColumnStyle->SetTextAlignmentVertical((LONG)VAL_BOTTOM);
	mp_oDefaultColumnStyle->SetTextYMargin(5);
	mp_oDefaultColumnStyle->GetFont()->SetStyle(FS_FONTSTYLEBOLD);

    mp_oDefaultRowStyle->Clear();

    mp_oDefaultCellStyle->Clear();
	mp_oDefaultCellStyle->SetAppearance((LONG)SA_FLAT);
	mp_oDefaultCellStyle->SetBackColor(Color::White);
	mp_oDefaultCellStyle->SetBorderColor(Color::Gray);
	mp_oDefaultCellStyle->SetBorderStyle((LONG)SBR_CUSTOM);
	mp_oDefaultCellStyle->GetCustomBorderStyle()->SetTop(FALSE);
	mp_oDefaultCellStyle->GetCustomBorderStyle()->SetLeft(FALSE);
	mp_oDefaultCellStyle->SetTextAlignmentHorizontal(HAL_LEFT);
	mp_oDefaultCellStyle->SetTextAlignmentVertical((LONG)VAL_TOP);
	mp_oDefaultCellStyle->SetTextYMargin(5);
	mp_oDefaultCellStyle->SetTextXMargin(5);
	mp_oDefaultCellStyle->GetFont()->SetStyle(FS_FONTSTYLEBOLD);
	mp_oDefaultCellStyle->SetImageAlignmentHorizontal((LONG)HAL_LEFT);
	mp_oDefaultCellStyle->SetImageAlignmentVertical((LONG)VAL_TOP);
	mp_oDefaultCellStyle->SetImageXMargin(0);
	mp_oDefaultCellStyle->SetImageYMargin(0);

    mp_oDefaultNodeStyle->Clear();
	mp_oDefaultNodeStyle->SetAppearance(SA_FLAT);
	mp_oDefaultNodeStyle->SetBackColor(Color::White);
	mp_oDefaultNodeStyle->SetBorderColor(Color::Gray);
	mp_oDefaultNodeStyle->SetBorderStyle(SBR_CUSTOM);
	mp_oDefaultNodeStyle->GetCustomBorderStyle()->SetTop(FALSE);
	mp_oDefaultNodeStyle->GetCustomBorderStyle()->SetLeft(FALSE);

    mp_oDefaultSplitterStyle->Clear();
	mp_oDefaultSplitterStyle->SetAppearance((LONG)SA_FLAT);
	mp_oDefaultSplitterStyle->SetBackColor(Color::Black);

    mp_oDefaultTaskStyle->Clear();
    mp_oDefaultTaskStyle->GetMilestoneStyle()->SetShapeIndex(FT_DIAMOND);

    mp_oDefaultPercentageStyle->Clear();
	mp_oDefaultPercentageStyle->SetAppearance(SA_FLAT);
    mp_oDefaultPercentageStyle->SetBackgroundMode(FP_SOLID);
    mp_oDefaultPercentageStyle->SetPlacement(PLC_OFFSETPLACEMENT);
    mp_oDefaultPercentageStyle->SetOffsetTop(10);
    mp_oDefaultPercentageStyle->SetOffsetBottom(15);
	mp_oDefaultPercentageStyle->SetBackColor(Color::Blue);

    mp_oDefaultPredecessorStyle->Clear();

    mp_oDefaultTimeBlockStyle->Clear();
	mp_oDefaultTimeBlockStyle->SetAppearance((LONG)SA_FLAT);
	mp_oDefaultTimeBlockStyle->SetBackgroundMode((LONG)FP_HATCH);
	mp_oDefaultTimeBlockStyle->SetHatchBackColor(Color::White);
	mp_oDefaultTimeBlockStyle->SetHatchForeColor(Color::Gray);
	mp_oDefaultTimeBlockStyle->SetHatchStyle((LONG)HS_PERCENT50);

    mp_oDefaultProgressLineStyle->Clear();
	mp_oDefaultProgressLineStyle->SetAppearance((LONG)SA_FLAT);
	mp_oDefaultProgressLineStyle->SetBackColor(Color::Black);

    // Views
    mp_oDefaultClientAreaStyle->Clear();
	mp_oDefaultClientAreaStyle->SetAppearance(SA_FLAT);
	mp_oDefaultClientAreaStyle->SetBackgroundMode(FP_TRANSPARENT);
	mp_oDefaultClientAreaStyle->SetBorderColor(Color::Gray);
	mp_oDefaultClientAreaStyle->SetBorderStyle(SBR_CUSTOM);
	mp_oDefaultClientAreaStyle->GetCustomBorderStyle()->SetTop(FALSE);
	mp_oDefaultClientAreaStyle->GetCustomBorderStyle()->SetLeft(FALSE);
	mp_oDefaultClientAreaStyle->GetCustomBorderStyle()->SetRight(FALSE);

    mp_oDefaultTimeLineStyle->Clear();

    mp_oDefaultTierStyle->Clear();
	mp_oDefaultTierStyle->SetAppearance(SA_FLAT);
	mp_oDefaultTierStyle->SetBorderStyle(SBR_NONE);
	mp_oDefaultTierStyle->SetDrawTextInVisibleArea(TRUE);

    mp_oDefaultTickMarkAreaStyle->Clear();

    // Scroll Bars
    mp_oDefaultScrollBarStyle->Clear();
	mp_oDefaultScrollBarStyle->SetAppearance(SA_FLAT);
	mp_oDefaultScrollBarStyle->SetBackgroundMode(FP_HATCH);
	mp_oDefaultScrollBarStyle->SetHatchStyle(HS_PERCENT50);
	mp_oDefaultScrollBarStyle->SetHatchForeColor(Color::MakeARGB(255, 192, 192, 192));
	mp_oDefaultScrollBarStyle->SetHatchBackColor(Color::White);
	mp_oDefaultScrollBarStyle->SetBorderStyle(SBR_SINGLE);
	mp_oDefaultScrollBarStyle->SetBorderColor(Color::MakeARGB(255, 192, 192, 192));

    mp_oDefaultSBSeparatorStyle->Clear();
    mp_oDefaultSBSeparatorStyle->SetAppearance(SA_RAISED);

    mp_oDefaultSBNormalStyle->Clear();
    mp_oDefaultSBNormalStyle->SetAppearance(SA_RAISED);

    mp_oDefaultSBPressedStyle->Clear();
    mp_oDefaultSBPressedStyle->SetAppearance(SA_SUNKEN);

    mp_oDefaultSBDisabledStyle->Clear();
	mp_oDefaultSBDisabledStyle->SetAppearance(SA_RAISED);
	mp_oDefaultSBDisabledStyle->GetScrollBarStyle()->SetArrowColor(Color::MakeARGB(255, 192, 192, 192));
}

clsStyles::~clsStyles()
{
	delete mp_oCollection;
	mp_oCollection = NULL;
	delete mp_oDefaultTaskStyle;
	mp_oDefaultTaskStyle = NULL;
	delete mp_oDefaultRowStyle;
	mp_oDefaultRowStyle = NULL;
	delete mp_oDefaultClientAreaStyle;
	mp_oDefaultClientAreaStyle = NULL;
	delete mp_oDefaultCellStyle;
	mp_oDefaultCellStyle = NULL;
	delete mp_oDefaultColumnStyle;
	mp_oDefaultColumnStyle = NULL;
	delete mp_oDefaultPercentageStyle;
	mp_oDefaultPercentageStyle = NULL;
	delete mp_oDefaultPredecessorStyle;
	mp_oDefaultPredecessorStyle = NULL;
	delete mp_oDefaultTimeLineStyle;
	mp_oDefaultTimeLineStyle = NULL;
	delete mp_oDefaultTimeBlockStyle;
	mp_oDefaultTimeBlockStyle = NULL;
	delete mp_oDefaultTickMarkAreaStyle;
	mp_oDefaultTickMarkAreaStyle = NULL;
	delete mp_oDefaultSplitterStyle;
	mp_oDefaultSplitterStyle = NULL;
	delete mp_oDefaultProgressLineStyle;
	mp_oDefaultProgressLineStyle = NULL;
	delete mp_oDefaultControlStyle;
	mp_oDefaultControlStyle = NULL;
	delete mp_oDefaultNodeStyle;
	mp_oDefaultNodeStyle = NULL;
	delete mp_oDefaultTierStyle;
	mp_oDefaultTierStyle = NULL;
	delete mp_oDefaultScrollBarStyle;
	mp_oDefaultScrollBarStyle = NULL;
	delete mp_oDefaultSBSeparatorStyle;
	mp_oDefaultSBSeparatorStyle = NULL;
	delete mp_oDefaultSBNormalStyle;
	mp_oDefaultSBNormalStyle = NULL;
	delete mp_oDefaultSBPressedStyle;
	mp_oDefaultSBPressedStyle = NULL;
	delete mp_oDefaultSBDisabledStyle;
	mp_oDefaultSBDisabledStyle = NULL;
	AfxOleUnlockApp();
}

void clsStyles::OnFinalRelease()
{
	CCmdTarget::OnFinalRelease();
}

LONG clsStyles::GetCount(void)
{
	return mp_oCollection->m_lCount();
}


BOOL clsStyles::ContainsKey(CString Key)
{
    if (g_IsNumeric(Key) == TRUE)
	{
        mp_oControl->mp_ErrorReport(STYLES_ADD_1, "Invalid Style object key, key cannot be numeric", "ActiveGanttVCCtl.Styles.ContainsKey");
        return FALSE;
	}
    return mp_oCollection->m_bDoesKeyExist(Key);
}

clsCollectionBase* clsStyles::GetCollection(void)
{
	return mp_oCollection;
}

void clsStyles::Finalize(void)
{
}

clsStyle* clsStyles::Item(CString Index)
{
	clsStyle *oStyle;
	oStyle = (clsStyle*)mp_oCollection->m_oItem(Index, STYLES_ITEM_1, STYLES_ITEM_2, STYLES_ITEM_3, STYLES_ITEM_4);
	return oStyle;
}

clsStyle* clsStyles::FItem(CString Index)
{
	if (Index == "DS_TASK") 
	{
		return mp_oDefaultTaskStyle;
	}
	else if (Index == "DS_ROW") 
	{
		return mp_oDefaultRowStyle;
	}
	else if (Index == "DS_CLIENTAREA") 
	{
		return mp_oDefaultClientAreaStyle;
	}
	else if (Index == "DS_CELL") 
	{
		return mp_oDefaultCellStyle;
	}
	else if (Index == "DS_COLUMN") 
	{
		return mp_oDefaultColumnStyle;
	}
	else if (Index == "DS_PERCENTAGE") 
	{
		return mp_oDefaultPercentageStyle;
	}
	else if (Index == "DS_PREDECESSOR") 
	{
		return mp_oDefaultPredecessorStyle;
	}
	else if (Index == "DS_TIMELINE") 
	{
		return mp_oDefaultTimeLineStyle;
	}
	else if (Index == "DS_TIMEBLOCK") 
	{
		return mp_oDefaultTimeBlockStyle;
	}
	else if (Index == "DS_TICKMARKAREA") 
	{
		return mp_oDefaultTickMarkAreaStyle;
	}
	else if (Index == "DS_SPLITTER") 
	{
		return mp_oDefaultSplitterStyle;
	}
	else if (Index == "DS_PROGRESSLINE") 
	{
		return mp_oDefaultProgressLineStyle;
	}
	else if (Index == "DS_CONTROL") 
	{
		return mp_oDefaultControlStyle;
	}
	else if (Index == "DS_NODE") 
	{
		return mp_oDefaultNodeStyle;
	}
	else if (Index == "DS_TIER")
	{
		return mp_oDefaultTierStyle;
	}
    else if (Index == "DS_SCROLLBAR")
    {
        return mp_oDefaultScrollBarStyle;
    }
    else if (Index == "DS_SB_NORMAL")
    {
        return mp_oDefaultSBNormalStyle;
    }
    else if (Index == "DS_SB_PRESSED")
    {
        return mp_oDefaultSBPressedStyle;
    }
    else if (Index == "DS_SB_DISABLED")
    {
        return mp_oDefaultSBDisabledStyle;
    }
    else if (Index == "DS_SB_SEPARATOR")
    {
        return mp_oDefaultSBSeparatorStyle;
    }
	else 
	{
		return (clsStyle*) mp_oCollection->m_oItem(Index, STYLES_ITEM_1, STYLES_ITEM_2, STYLES_ITEM_3, STYLES_ITEM_4);
	}
}

clsStyle* clsStyles::Add(CString Key)
{
	mp_oCollection->SetAddMode(TRUE);
	clsStyle* oStyle = new clsStyle(mp_oControl);
	Key = g_Trim(Key);
	oStyle->SetKey(Key);
	mp_oCollection->m_Add(oStyle, Key, STYLES_ADD_1, STYLES_ADD_2, TRUE, STYLES_ADD_3);
	return oStyle;
}

void clsStyles::Clear(void)
{

    LONG lIndex = 0;
    LONG lIndex2 = 0;
    clsColumn* oColumn;
    clsRow* oRow;
    clsCell* oCell;
    clsTask* oTask;
    clsPredecessor* oPredecessor;
    clsTimeBlock* oTimeBlock;
    clsPercentage* oPercentage;
    clsView* oView;

    mp_oControl->SetStyleIndex("");
    mp_oControl->GetSplitter()->SetStyleIndex(_T(""));

    mp_oControl->GetVerticalScrollBar()->GetScrollBar()->SetStyleIndex(_T(""));
    mp_oControl->GetVerticalScrollBar()->GetScrollBar()->GetArrowButtons()->SetNormalStyleIndex(_T(""));
    mp_oControl->GetVerticalScrollBar()->GetScrollBar()->GetArrowButtons()->SetPressedStyleIndex(_T(""));
    mp_oControl->GetVerticalScrollBar()->GetScrollBar()->GetArrowButtons()->SetDisabledStyleIndex(_T(""));
    mp_oControl->GetVerticalScrollBar()->GetScrollBar()->GetThumbButton()->SetNormalStyleIndex(_T(""));
    mp_oControl->GetVerticalScrollBar()->GetScrollBar()->GetThumbButton()->SetPressedStyleIndex(_T(""));
    mp_oControl->GetVerticalScrollBar()->GetScrollBar()->GetThumbButton()->SetDisabledStyleIndex(_T(""));

    mp_oControl->GetHorizontalScrollBar()->GetScrollBar()->SetStyleIndex(_T(""));
    mp_oControl->GetHorizontalScrollBar()->GetScrollBar()->GetArrowButtons()->SetNormalStyleIndex(_T(""));
    mp_oControl->GetHorizontalScrollBar()->GetScrollBar()->GetArrowButtons()->SetPressedStyleIndex(_T(""));
    mp_oControl->GetHorizontalScrollBar()->GetScrollBar()->GetArrowButtons()->SetDisabledStyleIndex(_T(""));
    mp_oControl->GetHorizontalScrollBar()->GetScrollBar()->GetThumbButton()->SetNormalStyleIndex(_T(""));
    mp_oControl->GetHorizontalScrollBar()->GetScrollBar()->GetThumbButton()->SetPressedStyleIndex(_T(""));
    mp_oControl->GetHorizontalScrollBar()->GetScrollBar()->GetThumbButton()->SetDisabledStyleIndex(_T(""));

    mp_oControl->GetScrollBarSeparator()->SetStyleIndex(_T(""));

    for (lIndex = 1; lIndex <= mp_oControl->GetColumns()->GetCount(); lIndex++)
    {
        oColumn = (clsColumn*) mp_oControl->GetColumns()->mp_oCollection->m_oReturnArrayElement(lIndex);
        oColumn->SetStyleIndex(_T(""));
    }
    for (lIndex = 1; lIndex <= mp_oControl->GetRows()->GetCount(); lIndex++)
    {
        oRow = (clsRow*) mp_oControl->GetRows()->mp_oCollection->m_oReturnArrayElement(lIndex);
        oRow->SetStyleIndex(_T(""));
        oRow->SetClientAreaStyleIndex(_T(""));
        oRow->GetNode()->SetStyleIndex(_T(""));
        for (lIndex2 = 1; lIndex2 <= oRow->GetCells()->GetCount(); lIndex2++)
        {
            oCell = (clsCell*) oRow->GetCells()->mp_oCollection->m_oReturnArrayElement(lIndex2);
            oCell->SetStyleIndex(_T(""));
        }
    }
    for (lIndex = 1; lIndex <= mp_oControl->GetTasks()->GetCount(); lIndex++)
    {
        oTask = (clsTask*) mp_oControl->GetTasks()->mp_oCollection->m_oReturnArrayElement(lIndex);
        oTask->SetStyleIndex(_T(""));
		oTask->SetWarningStyleIndex(_T(""));
    }
    for (lIndex = 1; lIndex <= mp_oControl->GetTimeBlocks()->GetCount(); lIndex++)
    {
        oTimeBlock = (clsTimeBlock*) mp_oControl->GetTimeBlocks()->mp_oCollection->m_oReturnArrayElement(lIndex);
        oTimeBlock->SetStyleIndex(_T(""));
    }
    for (lIndex = 1; lIndex <= mp_oControl->GetPercentages()->GetCount(); lIndex++)
    {
        oPercentage = (clsPercentage*) mp_oControl->GetPercentages()->mp_oCollection->m_oReturnArrayElement(lIndex);
        oPercentage->SetStyleIndex(_T(""));
    }
	for (lIndex = 1; lIndex <= mp_oControl->GetPredecessors()->GetCount(); lIndex++)
	{
		oPredecessor = (clsPredecessor*) mp_oControl->GetPredecessors()->mp_oCollection->m_oReturnArrayElement(lIndex);
		oPredecessor->SetStyleIndex(_T(""));
		oPredecessor->SetWarningStyleIndex(_T(""));
		oPredecessor->SetSelectedStyleIndex(_T(""));
	}
    for (lIndex = 1; lIndex <= mp_oControl->GetViews()->GetCount(); lIndex++)
    {
        oView = (clsView*) mp_oControl->GetViews()->mp_oCollection->m_oReturnArrayElement(lIndex);
        oView->GetTimeLine()->SetStyleIndex(_T(""));
        oView->GetTimeLine()->GetTickMarkArea()->SetStyleIndex(_T(""));
        oView->GetTimeLine()->GetTierArea()->GetUpperTier()->SetStyleIndex(_T(""));
        oView->GetTimeLine()->GetTierArea()->GetMiddleTier()->SetStyleIndex(_T(""));
        oView->GetTimeLine()->GetTierArea()->GetLowerTier()->SetStyleIndex(_T(""));
        oView->GetTimeLine()->GetTimeLineScrollBar()->GetScrollBar()->SetStyleIndex(_T(""));
        oView->GetTimeLine()->GetTimeLineScrollBar()->GetScrollBar()->GetArrowButtons()->SetNormalStyleIndex(_T(""));
        oView->GetTimeLine()->GetTimeLineScrollBar()->GetScrollBar()->GetArrowButtons()->SetPressedStyleIndex(_T(""));
        oView->GetTimeLine()->GetTimeLineScrollBar()->GetScrollBar()->GetArrowButtons()->SetDisabledStyleIndex(_T(""));
        oView->GetTimeLine()->GetTimeLineScrollBar()->GetScrollBar()->GetThumbButton()->SetNormalStyleIndex(_T(""));
        oView->GetTimeLine()->GetTimeLineScrollBar()->GetScrollBar()->GetThumbButton()->SetPressedStyleIndex(_T(""));
        oView->GetTimeLine()->GetTimeLineScrollBar()->GetScrollBar()->GetThumbButton()->SetDisabledStyleIndex(_T(""));
    }

	mp_oControl->mp_oTimeLineScrollBar->GetScrollBar()->SetStyleIndex(_T(""));
    mp_oControl->mp_oTimeLineScrollBar->GetScrollBar()->GetArrowButtons()->SetNormalStyleIndex(_T(""));
    mp_oControl->mp_oTimeLineScrollBar->GetScrollBar()->GetArrowButtons()->SetPressedStyleIndex(_T(""));
    mp_oControl->mp_oTimeLineScrollBar->GetScrollBar()->GetArrowButtons()->SetDisabledStyleIndex(_T(""));
    mp_oControl->mp_oTimeLineScrollBar->GetScrollBar()->GetThumbButton()->SetNormalStyleIndex(_T(""));
	mp_oControl->mp_oTimeLineScrollBar->GetScrollBar()->GetThumbButton()->SetPressedStyleIndex(_T(""));
    mp_oControl->mp_oTimeLineScrollBar->GetScrollBar()->GetThumbButton()->SetDisabledStyleIndex(_T(""));

    oView = mp_oControl->GetViews()->FItem("0");
    oView->GetTimeLine()->SetStyleIndex(_T(""));
    oView->GetTimeLine()->GetTickMarkArea()->SetStyleIndex(_T(""));
    oView->GetTimeLine()->GetTierArea()->GetUpperTier()->SetStyleIndex(_T(""));
    oView->GetTimeLine()->GetTierArea()->GetMiddleTier()->SetStyleIndex(_T(""));
    oView->GetTimeLine()->GetTierArea()->GetLowerTier()->SetStyleIndex(_T(""));
    oView->GetTimeLine()->GetTimeLineScrollBar()->GetScrollBar()->SetStyleIndex(_T(""));
    oView->GetTimeLine()->GetTimeLineScrollBar()->GetScrollBar()->GetArrowButtons()->SetNormalStyleIndex(_T(""));
    oView->GetTimeLine()->GetTimeLineScrollBar()->GetScrollBar()->GetArrowButtons()->SetPressedStyleIndex(_T(""));
    oView->GetTimeLine()->GetTimeLineScrollBar()->GetScrollBar()->GetArrowButtons()->SetDisabledStyleIndex(_T(""));
    oView->GetTimeLine()->GetTimeLineScrollBar()->GetScrollBar()->GetThumbButton()->SetNormalStyleIndex(_T(""));
    oView->GetTimeLine()->GetTimeLineScrollBar()->GetScrollBar()->GetThumbButton()->SetPressedStyleIndex(_T(""));
    oView->GetTimeLine()->GetTimeLineScrollBar()->GetScrollBar()->GetThumbButton()->SetDisabledStyleIndex(_T(""));



	mp_oCollection->m_Clear();
	mp_ResetDefaultStyles();
}

void clsStyles::Remove(CString Index)
{

    CString sRIndex = "";
    CString sRKey = "";
    mp_oCollection->m_GetKeyAndIndex(Index, &sRKey, &sRIndex);
    int lIndex = 0;
    int lIndex2 = 0;
    clsColumn* oColumn;
    clsRow* oRow;
    clsCell* oCell;
    clsTask* oTask;
    clsPredecessor* oPredecessor;
    clsTimeBlock* oTimeBlock;
    clsPercentage* oPercentage;
    clsView* oView;

    mp_oControl->SetStyleIndex(mp_GetNewStyleIndex(sRKey, sRIndex, mp_oControl->GetStyleIndex()));
    mp_oControl->GetSplitter()->SetStyleIndex(mp_GetNewStyleIndex(sRKey, sRIndex, mp_oControl->GetSplitter()->GetStyleIndex()));

    mp_oControl->GetVerticalScrollBar()->GetScrollBar()->SetStyleIndex(mp_GetNewStyleIndex(sRKey, sRIndex, mp_oControl->GetVerticalScrollBar()->GetScrollBar()->GetStyleIndex()));
    mp_oControl->GetVerticalScrollBar()->GetScrollBar()->GetArrowButtons()->SetNormalStyleIndex(mp_GetNewStyleIndex(sRKey, sRIndex, mp_oControl->GetVerticalScrollBar()->GetScrollBar()->GetArrowButtons()->GetNormalStyleIndex()));
    mp_oControl->GetVerticalScrollBar()->GetScrollBar()->GetArrowButtons()->SetPressedStyleIndex(mp_GetNewStyleIndex(sRKey, sRIndex, mp_oControl->GetVerticalScrollBar()->GetScrollBar()->GetArrowButtons()->GetPressedStyleIndex()));
    mp_oControl->GetVerticalScrollBar()->GetScrollBar()->GetArrowButtons()->SetDisabledStyleIndex(mp_GetNewStyleIndex(sRKey, sRIndex, mp_oControl->GetVerticalScrollBar()->GetScrollBar()->GetArrowButtons()->GetDisabledStyleIndex()));
    mp_oControl->GetVerticalScrollBar()->GetScrollBar()->GetThumbButton()->SetNormalStyleIndex(mp_GetNewStyleIndex(sRKey, sRIndex, mp_oControl->GetVerticalScrollBar()->GetScrollBar()->GetThumbButton()->GetNormalStyleIndex()));
    mp_oControl->GetVerticalScrollBar()->GetScrollBar()->GetThumbButton()->SetPressedStyleIndex(mp_GetNewStyleIndex(sRKey, sRIndex, mp_oControl->GetVerticalScrollBar()->GetScrollBar()->GetThumbButton()->GetPressedStyleIndex()));
    mp_oControl->GetVerticalScrollBar()->GetScrollBar()->GetThumbButton()->SetDisabledStyleIndex(mp_GetNewStyleIndex(sRKey, sRIndex, mp_oControl->GetVerticalScrollBar()->GetScrollBar()->GetThumbButton()->GetDisabledStyleIndex()));

    mp_oControl->GetHorizontalScrollBar()->GetScrollBar()->SetStyleIndex(mp_GetNewStyleIndex(sRKey, sRIndex, mp_oControl->GetHorizontalScrollBar()->GetScrollBar()->GetStyleIndex()));
    mp_oControl->GetHorizontalScrollBar()->GetScrollBar()->GetArrowButtons()->SetNormalStyleIndex(mp_GetNewStyleIndex(sRKey, sRIndex, mp_oControl->GetHorizontalScrollBar()->GetScrollBar()->GetArrowButtons()->GetNormalStyleIndex()));
    mp_oControl->GetHorizontalScrollBar()->GetScrollBar()->GetArrowButtons()->SetPressedStyleIndex(mp_GetNewStyleIndex(sRKey, sRIndex, mp_oControl->GetHorizontalScrollBar()->GetScrollBar()->GetArrowButtons()->GetPressedStyleIndex()));
    mp_oControl->GetHorizontalScrollBar()->GetScrollBar()->GetArrowButtons()->SetDisabledStyleIndex(mp_GetNewStyleIndex(sRKey, sRIndex, mp_oControl->GetHorizontalScrollBar()->GetScrollBar()->GetArrowButtons()->GetDisabledStyleIndex()));
    mp_oControl->GetHorizontalScrollBar()->GetScrollBar()->GetThumbButton()->SetNormalStyleIndex(mp_GetNewStyleIndex(sRKey, sRIndex, mp_oControl->GetHorizontalScrollBar()->GetScrollBar()->GetThumbButton()->GetNormalStyleIndex()));
    mp_oControl->GetHorizontalScrollBar()->GetScrollBar()->GetThumbButton()->SetPressedStyleIndex(mp_GetNewStyleIndex(sRKey, sRIndex, mp_oControl->GetHorizontalScrollBar()->GetScrollBar()->GetThumbButton()->GetPressedStyleIndex()));
    mp_oControl->GetHorizontalScrollBar()->GetScrollBar()->GetThumbButton()->SetDisabledStyleIndex(mp_GetNewStyleIndex(sRKey, sRIndex, mp_oControl->GetHorizontalScrollBar()->GetScrollBar()->GetThumbButton()->GetDisabledStyleIndex()));

    mp_oControl->GetScrollBarSeparator()->SetStyleIndex(mp_GetNewStyleIndex(sRKey, sRIndex, mp_oControl->GetScrollBarSeparator()->GetStyleIndex()));

    for (lIndex = 1; lIndex <= mp_oControl->GetColumns()->GetCount(); lIndex++)
    {
        oColumn = (clsColumn*) mp_oControl->GetColumns()->mp_oCollection->m_oReturnArrayElement(lIndex);
        oColumn->SetStyleIndex(mp_GetNewStyleIndex(sRKey, sRIndex, oColumn->GetStyleIndex()));
    }
    for (lIndex = 1; lIndex <= mp_oControl->GetRows()->GetCount(); lIndex++)
    {
        oRow = (clsRow*) mp_oControl->GetRows()->mp_oCollection->m_oReturnArrayElement(lIndex);
        oRow->SetStyleIndex(mp_GetNewStyleIndex(sRKey, sRIndex, oRow->GetStyleIndex()));
        oRow->SetClientAreaStyleIndex(mp_GetNewStyleIndex(sRKey, sRIndex, oRow->GetClientAreaStyleIndex()));
        oRow->GetNode()->SetStyleIndex(mp_GetNewStyleIndex(sRKey, sRIndex, oRow->GetNode()->GetStyleIndex()));
        for (lIndex2 = 1; lIndex2 <= oRow->GetCells()->GetCount(); lIndex2++)
        {
            oCell = (clsCell*) oRow->GetCells()->mp_oCollection->m_oReturnArrayElement(lIndex2);
            oCell->SetStyleIndex(mp_GetNewStyleIndex(sRKey, sRIndex, oCell->GetStyleIndex()));
        }
    }
    for (lIndex = 1; lIndex <= mp_oControl->GetTasks()->GetCount(); lIndex++)
    {
        oTask = (clsTask*) mp_oControl->GetTasks()->mp_oCollection->m_oReturnArrayElement(lIndex);
        oTask->SetStyleIndex(mp_GetNewStyleIndex(sRKey, sRIndex, oTask->GetStyleIndex()));
		oTask->SetWarningStyleIndex(mp_GetNewStyleIndex(sRKey, sRIndex, oTask->GetWarningStyleIndex()));
    }
	for (lIndex = 1; lIndex <= mp_oControl->GetPredecessors()->GetCount(); lIndex++)
	{
		oPredecessor = (clsPredecessor*) mp_oControl->GetPredecessors()->mp_oCollection->m_oReturnArrayElement(lIndex);
		oPredecessor->SetStyleIndex(mp_GetNewStyleIndex(sRKey, sRIndex, oPredecessor->GetStyleIndex()));
		oPredecessor->SetWarningStyleIndex(mp_GetNewStyleIndex(sRKey, sRIndex, oPredecessor->GetWarningStyleIndex()));
		oPredecessor->SetSelectedStyleIndex(mp_GetNewStyleIndex(sRKey, sRIndex, oPredecessor->GetSelectedStyleIndex()));
	}
    for (lIndex = 1; lIndex <= mp_oControl->GetTimeBlocks()->GetCount(); lIndex++)
    {
        oTimeBlock = (clsTimeBlock*) mp_oControl->GetTimeBlocks()->mp_oCollection->m_oReturnArrayElement(lIndex);
        oTimeBlock->SetStyleIndex(mp_GetNewStyleIndex(sRKey, sRIndex, oTimeBlock->GetStyleIndex()));
    }
    for (lIndex = 1; lIndex <= mp_oControl->GetPercentages()->GetCount(); lIndex++)
    {
        oPercentage = (clsPercentage*) mp_oControl->GetPercentages()->mp_oCollection->m_oReturnArrayElement(lIndex);
        oPercentage->SetStyleIndex(mp_GetNewStyleIndex(sRKey, sRIndex, oPercentage->GetStyleIndex()));
    }
    for (lIndex = 1; lIndex <= mp_oControl->GetViews()->GetCount(); lIndex++)
    {
        oView = (clsView*) mp_oControl->GetViews()->mp_oCollection->m_oReturnArrayElement(lIndex);
        oView->GetTimeLine()->SetStyleIndex(mp_GetNewStyleIndex(sRKey, sRIndex, oView->GetTimeLine()->GetStyleIndex()));
        oView->GetTimeLine()->GetTickMarkArea()->SetStyleIndex(mp_GetNewStyleIndex(sRKey, sRIndex, oView->GetTimeLine()->GetTickMarkArea()->GetStyleIndex()));
        oView->GetTimeLine()->GetTierArea()->GetUpperTier()->SetStyleIndex(mp_GetNewStyleIndex(sRKey, sRIndex, oView->GetTimeLine()->GetTierArea()->GetUpperTier()->GetStyleIndex()));
        oView->GetTimeLine()->GetTierArea()->GetMiddleTier()->SetStyleIndex(mp_GetNewStyleIndex(sRKey, sRIndex, oView->GetTimeLine()->GetTierArea()->GetMiddleTier()->GetStyleIndex()));
        oView->GetTimeLine()->GetTierArea()->GetLowerTier()->SetStyleIndex(mp_GetNewStyleIndex(sRKey, sRIndex, oView->GetTimeLine()->GetTierArea()->GetLowerTier()->GetStyleIndex()));
        oView->GetTimeLine()->GetTimeLineScrollBar()->GetScrollBar()->SetStyleIndex(mp_GetNewStyleIndex(sRKey, sRIndex, oView->GetTimeLine()->GetTimeLineScrollBar()->GetScrollBar()->GetStyleIndex()));
        oView->GetTimeLine()->GetTimeLineScrollBar()->GetScrollBar()->GetArrowButtons()->SetNormalStyleIndex(mp_GetNewStyleIndex(sRKey, sRIndex, oView->GetTimeLine()->GetTimeLineScrollBar()->GetScrollBar()->GetArrowButtons()->GetNormalStyleIndex()));
        oView->GetTimeLine()->GetTimeLineScrollBar()->GetScrollBar()->GetArrowButtons()->SetPressedStyleIndex(mp_GetNewStyleIndex(sRKey, sRIndex, oView->GetTimeLine()->GetTimeLineScrollBar()->GetScrollBar()->GetArrowButtons()->GetPressedStyleIndex()));
        oView->GetTimeLine()->GetTimeLineScrollBar()->GetScrollBar()->GetArrowButtons()->SetDisabledStyleIndex(mp_GetNewStyleIndex(sRKey, sRIndex, oView->GetTimeLine()->GetTimeLineScrollBar()->GetScrollBar()->GetArrowButtons()->GetDisabledStyleIndex()));
        oView->GetTimeLine()->GetTimeLineScrollBar()->GetScrollBar()->GetThumbButton()->SetNormalStyleIndex(mp_GetNewStyleIndex(sRKey, sRIndex, oView->GetTimeLine()->GetTimeLineScrollBar()->GetScrollBar()->GetThumbButton()->GetNormalStyleIndex()));
        oView->GetTimeLine()->GetTimeLineScrollBar()->GetScrollBar()->GetThumbButton()->SetPressedStyleIndex(mp_GetNewStyleIndex(sRKey, sRIndex, oView->GetTimeLine()->GetTimeLineScrollBar()->GetScrollBar()->GetThumbButton()->GetPressedStyleIndex()));
        oView->GetTimeLine()->GetTimeLineScrollBar()->GetScrollBar()->GetThumbButton()->SetDisabledStyleIndex(mp_GetNewStyleIndex(sRKey, sRIndex, oView->GetTimeLine()->GetTimeLineScrollBar()->GetScrollBar()->GetThumbButton()->GetDisabledStyleIndex()));
    }

    mp_oControl->mp_oTimeLineScrollBar->GetScrollBar()->SetStyleIndex(mp_GetNewStyleIndex(sRKey, sRIndex, oView->GetTimeLine()->GetTimeLineScrollBar()->GetScrollBar()->GetStyleIndex()));
    mp_oControl->mp_oTimeLineScrollBar->GetScrollBar()->GetArrowButtons()->SetNormalStyleIndex(mp_GetNewStyleIndex(sRKey, sRIndex, oView->GetTimeLine()->GetTimeLineScrollBar()->GetScrollBar()->GetArrowButtons()->GetNormalStyleIndex()));
    mp_oControl->mp_oTimeLineScrollBar->GetScrollBar()->GetArrowButtons()->SetPressedStyleIndex(mp_GetNewStyleIndex(sRKey, sRIndex, oView->GetTimeLine()->GetTimeLineScrollBar()->GetScrollBar()->GetArrowButtons()->GetPressedStyleIndex()));
    mp_oControl->mp_oTimeLineScrollBar->GetScrollBar()->GetArrowButtons()->SetDisabledStyleIndex(mp_GetNewStyleIndex(sRKey, sRIndex, oView->GetTimeLine()->GetTimeLineScrollBar()->GetScrollBar()->GetArrowButtons()->GetDisabledStyleIndex()));
    mp_oControl->mp_oTimeLineScrollBar->GetScrollBar()->GetThumbButton()->SetNormalStyleIndex(mp_GetNewStyleIndex(sRKey, sRIndex, oView->GetTimeLine()->GetTimeLineScrollBar()->GetScrollBar()->GetThumbButton()->GetNormalStyleIndex()));
    mp_oControl->mp_oTimeLineScrollBar->GetScrollBar()->GetThumbButton()->SetPressedStyleIndex(mp_GetNewStyleIndex(sRKey, sRIndex, oView->GetTimeLine()->GetTimeLineScrollBar()->GetScrollBar()->GetThumbButton()->GetPressedStyleIndex()));
    mp_oControl->mp_oTimeLineScrollBar->GetScrollBar()->GetThumbButton()->SetDisabledStyleIndex(mp_GetNewStyleIndex(sRKey, sRIndex, oView->GetTimeLine()->GetTimeLineScrollBar()->GetScrollBar()->GetThumbButton()->GetDisabledStyleIndex()));

    oView = mp_oControl->GetViews()->FItem("0");
    oView->GetTimeLine()->SetStyleIndex(mp_GetNewStyleIndex(sRKey, sRIndex, oView->GetTimeLine()->GetStyleIndex()));
    oView->GetTimeLine()->GetTickMarkArea()->SetStyleIndex(mp_GetNewStyleIndex(sRKey, sRIndex, oView->GetTimeLine()->GetTickMarkArea()->GetStyleIndex()));
    oView->GetTimeLine()->GetTierArea()->GetUpperTier()->SetStyleIndex(mp_GetNewStyleIndex(sRKey, sRIndex, oView->GetTimeLine()->GetTierArea()->GetUpperTier()->GetStyleIndex()));
    oView->GetTimeLine()->GetTierArea()->GetMiddleTier()->SetStyleIndex(mp_GetNewStyleIndex(sRKey, sRIndex, oView->GetTimeLine()->GetTierArea()->GetMiddleTier()->GetStyleIndex()));
    oView->GetTimeLine()->GetTierArea()->GetLowerTier()->SetStyleIndex(mp_GetNewStyleIndex(sRKey, sRIndex, oView->GetTimeLine()->GetTierArea()->GetLowerTier()->GetStyleIndex()));
    oView->GetTimeLine()->GetTimeLineScrollBar()->GetScrollBar()->SetStyleIndex(mp_GetNewStyleIndex(sRKey, sRIndex, oView->GetTimeLine()->GetTimeLineScrollBar()->GetScrollBar()->GetStyleIndex()));
    oView->GetTimeLine()->GetTimeLineScrollBar()->GetScrollBar()->GetArrowButtons()->SetNormalStyleIndex(mp_GetNewStyleIndex(sRKey, sRIndex, oView->GetTimeLine()->GetTimeLineScrollBar()->GetScrollBar()->GetArrowButtons()->GetNormalStyleIndex()));
    oView->GetTimeLine()->GetTimeLineScrollBar()->GetScrollBar()->GetArrowButtons()->SetPressedStyleIndex(mp_GetNewStyleIndex(sRKey, sRIndex, oView->GetTimeLine()->GetTimeLineScrollBar()->GetScrollBar()->GetArrowButtons()->GetPressedStyleIndex()));
    oView->GetTimeLine()->GetTimeLineScrollBar()->GetScrollBar()->GetArrowButtons()->SetDisabledStyleIndex(mp_GetNewStyleIndex(sRKey, sRIndex, oView->GetTimeLine()->GetTimeLineScrollBar()->GetScrollBar()->GetArrowButtons()->GetDisabledStyleIndex()));
    oView->GetTimeLine()->GetTimeLineScrollBar()->GetScrollBar()->GetThumbButton()->SetNormalStyleIndex(mp_GetNewStyleIndex(sRKey, sRIndex, oView->GetTimeLine()->GetTimeLineScrollBar()->GetScrollBar()->GetThumbButton()->GetNormalStyleIndex()));
    oView->GetTimeLine()->GetTimeLineScrollBar()->GetScrollBar()->GetThumbButton()->SetPressedStyleIndex(mp_GetNewStyleIndex(sRKey, sRIndex, oView->GetTimeLine()->GetTimeLineScrollBar()->GetScrollBar()->GetThumbButton()->GetPressedStyleIndex()));
    oView->GetTimeLine()->GetTimeLineScrollBar()->GetScrollBar()->GetThumbButton()->SetDisabledStyleIndex(mp_GetNewStyleIndex(sRKey, sRIndex, oView->GetTimeLine()->GetTimeLineScrollBar()->GetScrollBar()->GetThumbButton()->GetDisabledStyleIndex()));

	mp_oCollection->m_Remove(Index, STYLES_REMOVE_1, STYLES_REMOVE_2, STYLES_REMOVE_3, STYLES_REMOVE_4);
}

CString clsStyles::mp_GetNewStyleIndex(CString sKey, CString sIndex, CString sStyleIndex)
{
    if (sIndex == sStyleIndex)
    {
        return _T("");
    }
    if (sKey == sStyleIndex)
    {
        return _T("");
    }
    return sStyleIndex;
}

CString clsStyles::GetXML(void)
{
	LONG lIndex;
	clsStyle* oStyle; 
	clsXML oXML(mp_oControl, "Styles");
	oXML.InitializeWriter();
	for (lIndex = 1; lIndex <= GetCount(); lIndex++)
	{
		oStyle = (clsStyle*) mp_oCollection->m_oReturnArrayElement(lIndex);
		oXML.WriteObject(oStyle->GetXML());
	}
	return oXML.GetXML();
}

void clsStyles::SetXML(CString sXML)
{
	LONG lIndex;
	clsXML oXML(mp_oControl, "Styles");
	oXML.SetXML(sXML);
	oXML.InitializeReader();
	mp_oCollection->m_Clear();
	for (lIndex = 1;lIndex <= oXML.ReadCollectionCount();lIndex++)
	{
		clsStyle* oStyle = new clsStyle(mp_oControl);
		oStyle->SetXML(oXML.ReadCollectionObject(lIndex));
		mp_oCollection->SetAddMode(TRUE);
		mp_oCollection->m_Add(oStyle, oStyle->mp_sKey, STYLES_ADD_1, STYLES_ADD_2, TRUE, STYLES_ADD_3);
	}
}