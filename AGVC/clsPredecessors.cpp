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
#include "clsPredecessors.h"




IMPLEMENT_DYNCREATE(clsPredecessors, CCmdTarget)

//{8C98A18F-998C-412F-BB7D-33C7D0337921}
static const IID IID_IclsPredecessors = { 0x8C98A18F, 0x998C, 0x412F, { 0xBB, 0x7D, 0x33, 0xC7, 0xD0, 0x33, 0x79, 0x21} };

//{C61C8D21-775F-4930-A081-5ACE027356CA}
IMPLEMENT_OLECREATE_FLAGS(clsPredecessors, "AGVC.clsPredecessors", afxRegApartmentThreading, 0xC61C8D21, 0x775F, 0x4930, 0xA0, 0x81, 0x5A, 0xCE, 0x02, 0x73, 0x56, 0xCA)

BEGIN_DISPATCH_MAP(clsPredecessors, CCmdTarget)
	DISP_PROPERTY_EX_ID(clsPredecessors, "Count", 1, odl_GetCount, SetNotSupported, VT_I4) 
	DISP_PROPERTY_PARAM_ID(clsPredecessors, "Item", 2, odl_Item, SetNotSupported, VT_DISPATCH, VTS_BSTR) 
	DISP_FUNCTION_ID(clsPredecessors, "Add", 3, odl_Add, VT_DISPATCH, VTS_BSTR VTS_BSTR VTS_I4 VTS_BSTR VTS_BSTR) 
	DISP_FUNCTION_ID(clsPredecessors, "Clear", 4, odl_Clear, VT_EMPTY, VTS_NONE) 
	DISP_FUNCTION_ID(clsPredecessors, "Remove", 5, odl_Remove, VT_EMPTY, VTS_BSTR) 
	DISP_FUNCTION_ID(clsPredecessors, "GetXML", 6, odl_GetXML, VT_BSTR, VTS_NONE) 
	DISP_FUNCTION_ID(clsPredecessors, "SetXML", 7, odl_SetXML, VT_EMPTY, VTS_BSTR)
	DISP_DEFVALUE(clsPredecessors, "Item")
END_DISPATCH_MAP()

BEGIN_INTERFACE_MAP(clsPredecessors, CCmdTarget)
	INTERFACE_PART(clsPredecessors, IID_IclsPredecessors, Dispatch)
END_INTERFACE_MAP()

BEGIN_MESSAGE_MAP(clsPredecessors, CCmdTarget)
END_MESSAGE_MAP()


clsPredecessors::clsPredecessors(CActiveGanttVCCtl* oControl)
{
	EnableAutomation();
	AfxOleLockApp();
	mp_oControl = oControl;
	mp_oCollection = new clsCollectionBase(mp_oControl, "Predecessor");
}

clsPredecessors::clsPredecessors()
{
	AfxMessageBox(_T("Invalid constructor has been invoked for clsPredecessors. Please inform support@sourcecodestore.com of this error"), MB_OK | MB_ICONSTOP | MB_APPLMODAL);
}

clsPredecessors::~clsPredecessors()
{
	delete mp_oCollection;
	mp_oCollection = NULL;
	AfxOleUnlockApp();
}

void clsPredecessors::OnFinalRelease()
{
	CCmdTarget::OnFinalRelease();
}

LONG clsPredecessors::GetCount(void)
{
	return mp_oCollection->m_lCount();
}

clsCollectionBase* clsPredecessors::GetCollection(void)
{
	return mp_oCollection;
}

clsPredecessor* clsPredecessors::Item(CString Index)
{
	clsPredecessor *oPredecessor;
	oPredecessor = (clsPredecessor*)mp_oCollection->m_oItem(Index, PREDECESSORS_ITEM_1, PREDECESSORS_ITEM_2, PREDECESSORS_ITEM_3, PREDECESSORS_ITEM_4);
	return oPredecessor;
}

void clsPredecessors::Draw(void)
{
    LONG lIndex = 0;
    clsPredecessor* oPredecessor;
    for (lIndex = 1; lIndex <= GetCount(); lIndex++)
    {
        oPredecessor = (clsPredecessor*)mp_oCollection->m_oReturnArrayElement(lIndex);
		oPredecessor->ClearRectangles();
        if (oPredecessor->GetSuccessorTask()->GetRow()->GetHeight() > -1 && oPredecessor->GetPredecessorTask()->GetRow()->GetHeight() > -1 && oPredecessor->GetVisible() == TRUE)
        {
            switch (oPredecessor->GetPredecessorType())
            {
                case PCT_END_TO_END:
                    mp_DrawEndToEnd(oPredecessor);
                    break;
                case PCT_START_TO_START:
                    mp_DrawStartToStart(oPredecessor);
                    break;
                case PCT_START_TO_END:
                    mp_DrawStartToEnd(oPredecessor);
                    break;
                case PCT_END_TO_START:
                    mp_DrawEndToStart(oPredecessor);
                    break;
            }
        }
    }
}

void clsPredecessors::mp_DrawEndToEnd(clsPredecessor* oPredecessor)
{
    clsTask* oPredecessorTask;
    clsTask* oSuccessorTask;
    int lPredecessorCtr = 0;
    int lSuccessorCtr = 0;
    int lXOffset = 0;
    int lYOffset = 0;
    clsPredecessorStyle* oStyle;         
    oStyle = mp_GetPredecessorStyle(oPredecessor);
    lXOffset = oStyle->GetXOffset();
    lYOffset = oStyle->GetYOffset();
    oPredecessorTask = oPredecessor->GetPredecessorTask();
    oSuccessorTask = oPredecessor->GetSuccessorTask();
    if (mp_DrawPredecessor(oPredecessorTask, oSuccessorTask) == FALSE)
    {
        return;
    }
    lPredecessorCtr = oPredecessorTask->GetTop() + ((oPredecessorTask->GetBottom() - oPredecessorTask->GetTop()) / 2);
    lSuccessorCtr = oSuccessorTask->GetTop() + ((oSuccessorTask->GetBottom() - oSuccessorTask->GetTop()) / 2);
    if (oPredecessorTask->GetRight() >= oSuccessorTask->GetRight())
    {
		S_Point oPoints[4];
        oPoints[0].X = oPredecessorTask->GetRight();
        oPoints[0].Y = lPredecessorCtr;
        oPoints[1].X = oPredecessorTask->GetRight() + lXOffset;
        oPoints[1].Y = lPredecessorCtr;
        oPoints[2].X = oPredecessorTask->GetRight() + lXOffset;
        oPoints[2].Y = lSuccessorCtr;
        oPoints[3].X = oSuccessorTask->GetRight();
        oPoints[3].Y = lSuccessorCtr;
        mp_DrawPredecessorLines(oPoints, 4, oPredecessor);
        mp_oControl->clsG->mp_DrawArrow(oPoints[3].X + 1, oPoints[3].Y, AWD_LEFT, oStyle->GetArrowSize(), oStyle->GetLineColor());
    }
    else
    {
		S_Point oPoints[4];
        oPoints[0].X = oSuccessorTask->GetRight();
        oPoints[0].Y = lSuccessorCtr;
        oPoints[1].X = oSuccessorTask->GetRight() + lXOffset;
        oPoints[1].Y = lSuccessorCtr;
        oPoints[2].X = oSuccessorTask->GetRight() + lXOffset;
        oPoints[2].Y = lPredecessorCtr;
        oPoints[3].X = oPredecessorTask->GetRight();
        oPoints[3].Y = lPredecessorCtr;
        mp_DrawPredecessorLines(oPoints, 4, oPredecessor);
        mp_oControl->clsG->mp_DrawArrow(oPoints[0].X + 1, oPoints[0].Y, AWD_LEFT, oStyle->GetArrowSize(), oStyle->GetLineColor());
    }

}

void clsPredecessors::mp_DrawStartToStart(clsPredecessor* oPredecessor)
{
    clsTask* oPredecessorTask;
    clsTask* oSuccessorTask;
    int lPredecessorCtr = 0;
    int lSuccessorCtr = 0;
    int lXOffset = 0;
    int lYOffset = 0;
    clsPredecessorStyle* oStyle;         
    oStyle = mp_GetPredecessorStyle(oPredecessor);
    lXOffset = oStyle->GetXOffset();
    lYOffset = oStyle->GetYOffset();
    oPredecessorTask = oPredecessor->GetPredecessorTask();
    oSuccessorTask = oPredecessor->GetSuccessorTask();
    if (mp_DrawPredecessor(oPredecessorTask, oSuccessorTask) == FALSE)
    {
        return;
    }
    lPredecessorCtr = oPredecessorTask->GetTop() + ((oPredecessorTask->GetBottom() - oPredecessorTask->GetTop()) / 2);
    lSuccessorCtr = oSuccessorTask->GetTop() + ((oSuccessorTask->GetBottom() - oSuccessorTask->GetTop()) / 2);
    if (oPredecessorTask->GetLeft() <= oSuccessorTask->GetLeft())
    {
		S_Point oPoints[4];
        oPoints[0].X = oPredecessorTask->GetLeft();
        oPoints[0].Y = lPredecessorCtr;
        oPoints[1].X = oPredecessorTask->GetLeft() - lXOffset;
        oPoints[1].Y = lPredecessorCtr;
        oPoints[2].X = oPredecessorTask->GetLeft() - lXOffset;
        oPoints[2].Y = lSuccessorCtr;
        oPoints[3].X = oSuccessorTask->GetLeft();
        oPoints[3].Y = lSuccessorCtr;
        mp_DrawPredecessorLines(oPoints, 4, oPredecessor);
        mp_oControl->clsG->mp_DrawArrow(oPoints[3].X - 1, oPoints[3].Y, AWD_RIGHT, oStyle->GetArrowSize(), oStyle->GetLineColor());
    }
    else
    {
		S_Point oPoints[4];
        oPoints[0].X = oSuccessorTask->GetLeft();
        oPoints[0].Y = lSuccessorCtr;
        oPoints[1].X = oSuccessorTask->GetLeft() - lXOffset;
        oPoints[1].Y = lSuccessorCtr;
        oPoints[2].X = oSuccessorTask->GetLeft() - lXOffset;
        oPoints[2].Y = lPredecessorCtr;
        oPoints[3].X = oPredecessorTask->GetLeft();
        oPoints[3].Y = lPredecessorCtr;
        mp_DrawPredecessorLines(oPoints, 4, oPredecessor);
        mp_oControl->clsG->mp_DrawArrow(oPoints[0].X - 1, oPoints[0].Y, AWD_RIGHT, oStyle->GetArrowSize(), oStyle->GetLineColor());
    }
}

void clsPredecessors::mp_DrawStartToEnd(clsPredecessor* oPredecessor)
{
    clsTask* oPredecessorTask;
    clsTask* oSuccessorTask;
    int lPredecessorCtr = 0;
    int lSuccessorCtr = 0;
    int lXOffset = 0;
    int lYOffset = 0;
    clsPredecessorStyle* oStyle;         
    oStyle = mp_GetPredecessorStyle(oPredecessor);
    lXOffset = oStyle->GetXOffset();
    lYOffset = oStyle->GetYOffset();
    oPredecessorTask = oPredecessor->GetPredecessorTask();
    oSuccessorTask = oPredecessor->GetSuccessorTask();
    if (mp_DrawPredecessor(oPredecessorTask, oSuccessorTask) == FALSE)
    {
        return;
    }
    lPredecessorCtr = oPredecessorTask->GetTop() + ((oPredecessorTask->GetBottom() - oPredecessorTask->GetTop()) / 2);
    lSuccessorCtr = oSuccessorTask->GetTop() + ((oSuccessorTask->GetBottom() - oSuccessorTask->GetTop()) / 2);

    //Down
    if (lPredecessorCtr < lSuccessorCtr)
    {
		S_Point oPoints[6];
        oPoints[0].X = oPredecessorTask->GetLeft();
        oPoints[0].Y = lPredecessorCtr;
        oPoints[1].X = oPredecessorTask->GetLeft() - lXOffset;
        oPoints[1].Y = lPredecessorCtr;
        oPoints[2].X = oPredecessorTask->GetLeft() - lXOffset;
        oPoints[2].Y = oPredecessorTask->GetBottom() + lYOffset;
        oPoints[3].X = oSuccessorTask->GetRight() + lXOffset;
        oPoints[3].Y = oPredecessorTask->GetBottom() + lYOffset;
        oPoints[4].X = oSuccessorTask->GetRight() + lXOffset;
        oPoints[4].Y = lSuccessorCtr;
        oPoints[5].X = oSuccessorTask->GetRight();
        oPoints[5].Y = lSuccessorCtr;
        mp_DrawPredecessorLines(oPoints, 6, oPredecessor);
        mp_oControl->clsG->mp_DrawArrow(oPoints[5].X + 1, oPoints[5].Y, AWD_LEFT, oStyle->GetArrowSize(), oStyle->GetLineColor());
    }
    else
    {
		S_Point oPoints[6];
        oPoints[0].X = oPredecessorTask->GetLeft();
        oPoints[0].Y = lPredecessorCtr;
        oPoints[1].X = oPredecessorTask->GetLeft() - lXOffset;
        oPoints[1].Y = lPredecessorCtr;
        oPoints[2].X = oPredecessorTask->GetLeft() - lXOffset;
        oPoints[2].Y = oPredecessorTask->GetTop() - lYOffset;
        oPoints[3].X = oSuccessorTask->GetRight() + lXOffset;
        oPoints[3].Y = oPredecessorTask->GetTop() - lYOffset;
        oPoints[4].X = oSuccessorTask->GetRight() + lXOffset;
        oPoints[4].Y = lSuccessorCtr;
        oPoints[5].X = oSuccessorTask->GetRight();
        oPoints[5].Y = lSuccessorCtr;
        mp_DrawPredecessorLines(oPoints, 6, oPredecessor);
        mp_oControl->clsG->mp_DrawArrow(oPoints[5].X + 1, oPoints[5].Y, AWD_LEFT, oStyle->GetArrowSize(), oStyle->GetLineColor());
    }
}

void clsPredecessors::mp_DrawEndToStart(clsPredecessor* oPredecessor)
{
    clsTask* oPredecessorTask;
    clsTask* oSuccessorTask;
    int lPredecessorCtr = 0;
    int lSuccessorCtr = 0;
    int lXOffset = 0;
    int lYOffset = 0;
    clsPredecessorStyle* oStyle;         
    oStyle = mp_GetPredecessorStyle(oPredecessor);
    lXOffset = oStyle->GetXOffset();
    lYOffset = oStyle->GetYOffset();
    oPredecessorTask = oPredecessor->GetPredecessorTask();
    oSuccessorTask = oPredecessor->GetSuccessorTask();
    if (mp_DrawPredecessor(oPredecessorTask, oSuccessorTask) == FALSE)
    {
        return;
    }
    lPredecessorCtr = oPredecessorTask->GetTop() + ((oPredecessorTask->GetBottom() - oPredecessorTask->GetTop()) / 2);
    lSuccessorCtr = oSuccessorTask->GetTop() + ((oSuccessorTask->GetBottom() - oSuccessorTask->GetTop()) / 2);
    if (oPredecessor->GetPredecessorTask()->GetRight() <= oPredecessor->GetSuccessorTask()->GetLeft())
    {
        ////With Lag
        ////Down
        if (lPredecessorCtr < lSuccessorCtr)
        {
			S_Point oPoints[3];
			oPoints[0].X = oPredecessorTask->GetRight();
			oPoints[0].Y = lPredecessorCtr;
			if (oSuccessorTask->GetType() == OT_TASK)
			{
				oPoints[1].X = oSuccessorTask->GetLeft() + lXOffset;
				oPoints[1].Y = lPredecessorCtr;
				oPoints[2].X = oSuccessorTask->GetLeft() + lXOffset;
				oPoints[2].Y = oSuccessorTask->GetTop();
			}
			else
			{
				oPoints[1].X = mp_oControl->GetMathLib()->GetXCoordinateFromDate(oSuccessorTask->GetStartDate());
				oPoints[1].Y = lPredecessorCtr;
				oPoints[2].X = mp_oControl->GetMathLib()->GetXCoordinateFromDate(oSuccessorTask->GetStartDate());
				oPoints[2].Y = oSuccessorTask->GetTop();
			}
            mp_DrawPredecessorLines(oPoints, 3, oPredecessor);
            mp_oControl->clsG->mp_DrawArrow(oPoints[2].X, oPoints[2].Y - 1, AWD_DOWN, oStyle->GetArrowSize(), oStyle->GetLineColor());
        }
        else
        {
			S_Point oPoints[3];
            oPoints[0].X = oPredecessorTask->GetRight();
            oPoints[0].Y = lPredecessorCtr;
			if (oSuccessorTask->GetType() == OT_TASK)
			{
				oPoints[1].X = oSuccessorTask->GetLeft() + lXOffset;
				oPoints[1].Y = lPredecessorCtr;
				oPoints[2].X = oSuccessorTask->GetLeft() + lXOffset;
				oPoints[2].Y = oSuccessorTask->GetBottom();
			}
			else
			{
				oPoints[1].X = mp_oControl->GetMathLib()->GetXCoordinateFromDate(oSuccessorTask->GetStartDate());
				oPoints[1].Y = lPredecessorCtr;
				oPoints[2].X = mp_oControl->GetMathLib()->GetXCoordinateFromDate(oSuccessorTask->GetStartDate());
				oPoints[2].Y = oSuccessorTask->GetBottom();
			}
            mp_DrawPredecessorLines(oPoints, 3, oPredecessor);
            mp_oControl->clsG->mp_DrawArrow(oPoints[2].X, oPoints[2].Y + 1, AWD_UP, oStyle->GetArrowSize(), oStyle->GetLineColor());
        }
    }
    else
    {
        ////With Lead
        ////Down
        if (lPredecessorCtr < lSuccessorCtr)
        {
			S_Point oPoints[6];
            oPoints[0].X = oPredecessorTask->GetRight();
            oPoints[0].Y = lPredecessorCtr;
            oPoints[1].X = oPredecessorTask->GetRight() + lXOffset;
            oPoints[1].Y = lPredecessorCtr;
            oPoints[2].X = oPredecessorTask->GetRight() + lXOffset;
            oPoints[2].Y = lSuccessorCtr - lYOffset;
            oPoints[3].X = oSuccessorTask->GetLeft() - lXOffset;
            oPoints[3].Y = lSuccessorCtr - lYOffset;
            oPoints[4].X = oSuccessorTask->GetLeft() - lXOffset;
            oPoints[4].Y = lSuccessorCtr;
            oPoints[5].X = oSuccessorTask->GetLeft();
            oPoints[5].Y = lSuccessorCtr;
            mp_DrawPredecessorLines(oPoints, 6, oPredecessor);
            mp_oControl->clsG->mp_DrawArrow(oPoints[5].X - 1, oPoints[5].Y, AWD_RIGHT, oStyle->GetArrowSize(), oStyle->GetLineColor());
            //// Up
        }
        else
        {
			S_Point oPoints[6];
            oPoints[0].X = oPredecessorTask->GetRight();
            oPoints[0].Y = lPredecessorCtr;
            oPoints[1].X = oPredecessorTask->GetRight() + lXOffset;
            oPoints[1].Y = lPredecessorCtr;
            oPoints[2].X = oPredecessorTask->GetRight() + lXOffset;
            oPoints[2].Y = lSuccessorCtr + lYOffset;
            oPoints[3].X = oSuccessorTask->GetLeft() - lXOffset;
            oPoints[3].Y = lSuccessorCtr + lYOffset;
            oPoints[4].X = oSuccessorTask->GetLeft() - lXOffset;
            oPoints[4].Y = lSuccessorCtr;
            oPoints[5].X = oSuccessorTask->GetLeft();
            oPoints[5].Y = lSuccessorCtr;
            mp_DrawPredecessorLines(oPoints, 6, oPredecessor);
            mp_oControl->clsG->mp_DrawArrow(oPoints[5].X - 1, oPoints[5].Y, AWD_RIGHT, oStyle->GetArrowSize(), oStyle->GetLineColor());
        }
    }
}

void clsPredecessors::mp_DrawPredecessorLines(S_Point oPoints[], int Size, clsPredecessor* oPredecessor)
{
	int i = 0;
	for (i = 0; i <= Size - 2; i++)
	{
		clsPredecessorStyle* oStyle;         
		oStyle = mp_GetPredecessorStyle(oPredecessor);
		mp_oControl->clsG->mp_DrawLine(oPoints[i].X, oPoints[i].Y, oPoints[i + 1].X, oPoints[i + 1].Y, LT_NORMAL, oStyle->GetLineColor(), (GRE_LINEDRAWSTYLE)oStyle->GetLineStyle(), oStyle->GetLineWidth());
		S_Rectangle oRectangle;
		////Vertical Line
		if (oPoints[i].X == oPoints[i + 1].X)
		{
			oRectangle.X1 = oPoints[i].X - mp_oControl->GetCurrentViewObject()->GetClientArea()->GetPredecessorSelectionOffset();
			oRectangle.X2 = oPoints[i + 1].X + mp_oControl->GetCurrentViewObject()->GetClientArea()->GetPredecessorSelectionOffset();
			if (oPoints[i].Y < oPoints[i + 1].Y)
			{
				oRectangle.Y1 = oPoints[i].Y;
				oRectangle.Y2 = oPoints[i + 1].Y;
			}
			else
			{
				oRectangle.Y1 = oPoints[i + 1].Y;
				oRectangle.Y2 = oPoints[i].Y;
			}
			////Horizontal Line
		}
		else if (oPoints[i].Y == oPoints[i + 1].Y)
		{
			oRectangle.Y1 = oPoints[i].Y - mp_oControl->GetCurrentViewObject()->GetClientArea()->GetPredecessorSelectionOffset();
			oRectangle.Y2 = oPoints[i + 1].Y + mp_oControl->GetCurrentViewObject()->GetClientArea()->GetPredecessorSelectionOffset();
			if (oPoints[i].X < oPoints[i + 1].X)
			{
				oRectangle.X1 = oPoints[i].X;
				oRectangle.X2 = oPoints[i + 1].X;
			}
			else
			{
				oRectangle.X1 = oPoints[i + 1].X;
				oRectangle.X2 = oPoints[i].X;
			}
		}
		oPredecessor->AddRectangle(oRectangle);
	}
}

clsPredecessorStyle* clsPredecessors::mp_GetPredecessorStyle(clsPredecessor* oPredecessor)
{
	clsPredecessorStyle* oStyle;
	if (oPredecessor->mp_bWarning == FALSE)
	{
		if (oPredecessor->GetIndex() == mp_oControl->GetSelectedPredecessorIndex())
		{
			oStyle = oPredecessor->GetSelectedStyle()->GetPredecessorStyle();
		}
		else
		{
			oStyle = oPredecessor->GetStyle()->GetPredecessorStyle();
		}
	}
	else
	{
		oStyle = oPredecessor->GetWarningStyle()->GetPredecessorStyle();
	}
	return oStyle;
}

BOOL clsPredecessors::mp_DrawPredecessor(clsTask* oTask1, clsTask* oTask2)
{
    if (oTask1->GetRow()->GetClientAreaVisibility() == VS_ABOVEVISIBLEAREA && oTask2->GetRow()->GetClientAreaVisibility() == VS_ABOVEVISIBLEAREA)
	{
        return FALSE;
	}
    if (oTask1->GetRow()->GetClientAreaVisibility() == VS_BELOWVISIBLEAREA && oTask2->GetRow()->GetClientAreaVisibility() == VS_BELOWVISIBLEAREA)
	{
        return FALSE;
	}
    if (oTask1->GetClientAreaVisiblity() == VS_RIGHTOFVISIBLEAREA && oTask2->GetClientAreaVisiblity() == VS_RIGHTOFVISIBLEAREA)
	{
        return FALSE;
	}
    if (oTask1->GetClientAreaVisiblity() == VS_LEFTOFVISIBLEAREA && oTask2->GetClientAreaVisiblity() == VS_LEFTOFVISIBLEAREA)
	{
        return FALSE;
	}
	return TRUE;
}

clsPredecessor* clsPredecessors::Add(CString SuccessorKey, CString PredecessorKey, LONG PredecessorType, CString Key, CString StyleIndex)
{
	mp_oCollection->SetAddMode(TRUE);
	clsPredecessor* oPredecessor = new clsPredecessor(mp_oControl, this);
	oPredecessor->SetPredecessorType(PredecessorType);
	oPredecessor->SetPredecessorKey(PredecessorKey);
	oPredecessor->SetStyleIndex(StyleIndex);
	oPredecessor->SetKey(Key);
	oPredecessor->SetSuccessorKey(SuccessorKey);
	mp_oCollection->m_Add(oPredecessor, Key, PREDECESSORS_ADD_1, PREDECESSORS_ADD_2, FALSE, PREDECESSORS_ADD_3);
	return oPredecessor;
}

void clsPredecessors::Clear(void)
{
	mp_oCollection->m_Clear();
}

void clsPredecessors::Remove(CString Index)
{
	mp_oCollection->m_Remove(Index, PREDECESSORS_REMOVE_1, PREDECESSORS_REMOVE_2, PREDECESSORS_REMOVE_3, PREDECESSORS_REMOVE_4);
}

CString clsPredecessors::GetXML(void)
{
	LONG lIndex;
	clsPredecessor* oPredecessor; 
	clsXML oXML(mp_oControl, "Predecessors");
	oXML.InitializeWriter();
	for (lIndex = 1; lIndex <= GetCount(); lIndex++)
	{
		oPredecessor = (clsPredecessor*) mp_oCollection->m_oReturnArrayElement(lIndex);
		oXML.WriteObject(oPredecessor->GetXML());
	}
	return oXML.GetXML();
}

void clsPredecessors::SetXML(CString sXML)
{
	LONG lIndex;
	clsXML oXML(mp_oControl, "Predecessors");
	oXML.SetXML(sXML);
	oXML.InitializeReader();
	mp_oCollection->m_Clear();
	for (lIndex = 1;lIndex <= oXML.ReadCollectionCount();lIndex++)
	{
		clsPredecessor* oPredecessor = new clsPredecessor(mp_oControl, this);
		oPredecessor->SetXML(oXML.ReadCollectionObject(lIndex));
		mp_oCollection->SetAddMode(TRUE);
		mp_oCollection->m_Add(oPredecessor, oPredecessor->mp_sKey, PREDECESSORS_ADD_1, PREDECESSORS_ADD_2, FALSE, PREDECESSORS_ADD_3);
	}
}