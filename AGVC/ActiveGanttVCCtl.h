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

class clsView;
class clsToolTip;
class clsFont;
class clsCultureInfo;
class clsRows;
class clsTasks;
class clsColumns;
class clsStyles;
class clsLayers;
class clsPercentages;
class clsTimeBlocks;
class clsViews;
class clsSplitter;
class clsPrinter;
class clsTreeview;
class clsDrawing;
class clsMath;
class clsString;
class clsVerticalScrollBar;
class clsHorizontalScrollBar;
class clsMouseKeyboardEvents;
class clsGraphics;
class clsGDIGraphics;
class clsImage;
class ToolTipEventArgs;
class ObjectAddedEventArgs;
class CustomTierDrawEventArgs;
class MouseEventArgs;
class KeyEventArgs;
class ScrollEventArgs;
class DrawEventArgs;
class PredecessorDrawEventArgs;
class ObjectSelectedEventArgs;
class ObjectStateChangedEventArgs;
class ErrorEventArgs;
class NodeEventArgs;
class MouseWheelEventArgs;
class clsXML;

#include "winspool.h"


using namespace Gdiplus;




class CActiveGanttVCCtl : public COleControl
{
	DECLARE_DYNCREATE(CActiveGanttVCCtl)

	ULONG_PTR m_gdiplusToken; //GDI +

	HWND f_Hwnd();


public:
	CActiveGanttVCCtl();
public:
	virtual void OnDraw(CDC* pdc, const CRect& rcBounds, const CRect& rcInvalid);
	virtual void DoPropExchange(CPropExchange* pPX);
	virtual void OnResetState();

	//Key Capture
	afx_msg void OnKeyDown(UINT nChar, UINT nRepCnt, UINT nFlags);
	afx_msg void OnSysKeyDown(UINT nChar, UINT nRepCnt, UINT nFlags);
	afx_msg void OnKeyUp(UINT nChar, UINT nRepCnt, UINT nFlags);
	afx_msg void OnSysKeyUp(UINT nChar, UINT nRepCnt, UINT nFlags);
	afx_msg void OnSysDeadChar(UINT nChar, UINT nRepCnt, UINT nFlags);
	afx_msg void OnDeadChar(UINT nChar, UINT nRepCnt, UINT nFlags);
	afx_msg void OnChar(UINT nChar, UINT nRepCnt, UINT nFlags);
	//Mouse Capture
	afx_msg void OnLButtonDblClk(UINT nFlags, CPoint point);
	afx_msg void OnLButtonDown(UINT nFlags, CPoint point);
	afx_msg void OnLButtonUp(UINT nFlags, CPoint point);
	afx_msg void OnMButtonDblClk(UINT nFlags, CPoint point);
	afx_msg void OnMButtonDown(UINT nFlags, CPoint point);
	afx_msg void OnMButtonUp(UINT nFlags, CPoint point);
	afx_msg void OnRButtonDblClk(UINT nFlags, CPoint point);
	afx_msg void OnRButtonDown(UINT nFlags, CPoint point);
	afx_msg void OnRButtonUp(UINT nFlags, CPoint point);
	afx_msg BOOL OnMouseWheel(UINT nFlags, short zDelta, CPoint pt);
	afx_msg void OnMouseMove(UINT nFlags, CPoint point);
	//Other Events
	afx_msg void OnSize(UINT nType, int cx, int cy);
	afx_msg BOOL OnSetCursor(CWnd* pWnd, UINT nHitTest, UINT message);
	afx_msg int OnCreate(LPCREATESTRUCT lpCreateStruct);

	//Message translators
	void ActiveGanttVCCtl_DoubleClick(void);
	void ActiveGanttVCCtl_Click(void);
	void ActiveGanttVCCtl_MouseDown(LONG X, LONG Y, LONG Button);
	void ActiveGanttVCCtl_MouseMove(LONG X, LONG Y);
	void ActiveGanttVCCtl_MouseUp(LONG X, LONG Y);

	void ActiveGanttVCCtl_KeyDown(void);
	void ActiveGanttVCCtl_KeyPress(void);
	void ActiveGanttVCCtl_KeyUp(void);

	void ActiveGanttVCCtl_OnMouseWheel(void);

	//Timers:
	void SetTimerEx(UINT_PTR nIDEvent, UINT nElapse);
	void KillTimerEx();
	afx_msg void OnTimer(UINT nIDEvent);
	LONG mp_oTimer;

protected:
	~CActiveGanttVCCtl();
	

// Member Variables
public:

	LONG mp_lAboutShown;

	
	CString mp_sPrinterErrorMessage;
	BOOL mp_bAllowAdd;
	BOOL mp_bAllowEdit;
	BOOL mp_bAllowSplitterMove;
	BOOL mp_bAllowRowSize;
	BOOL mp_bAllowRowMove;
	BOOL mp_bAllowColumnSize;
	BOOL mp_bAllowColumnMove;
	BOOL mp_bAllowTimeLineScroll;
	BOOL mp_bAllowPredecessorAdd;
	BOOL mp_bDoubleBuffering;
	BOOL mp_bPropertiesRead;
	BOOL mp_bEnforcePredecessors;
	COleDateTime mp_dtTimeLineStartBuffer;
	COleDateTime mp_dtTimeLineEndBuffer;
	LONG mp_lMinColumnWidth;
	LONG mp_lMinRowHeight;
	LONG mp_lSelectedTaskIndex;
	LONG mp_lSelectedColumnIndex;
	LONG mp_lSelectedRowIndex;
	LONG mp_lSelectedCellIndex;
	LONG mp_lSelectedPercentageIndex;
	LONG mp_lSelectedPredecessorIndex;
	LONG mp_lTreeviewColumnIndex;
	CString mp_sCurrentLayer;
	CString mp_sCurrentView;
	LONG mp_yAddMode;
	LONG mp_yAddDurationInterval;
	LONG mp_yScrollBarBehaviour;
	LONG mp_yTimeBlockBehaviour;
	LONG mp_yLayerEnableObjects;
	LONG mp_yErrorReports;
	LONG mp_yTierAppearanceScope;
	LONG mp_yTierFormatScope;
	LONG mp_yTimeLineScrollBarScope;
	LONG mp_yProgressLineScope;
	LONG mp_yPredecessorMode;
	LONG mp_yControlTemplate;
    LONG mp_yObjectTemplate;
	CString mp_sControlTag;
	CString mp_sStyleIndex;

	CString mp_sImageTag;
	

	LONG mp_lObjectCount;

	//Pointers
	clsTimeBlocks* tmpTimeBlocks;
	clsMouseKeyboardEvents* MouseKeyboardEvents;
	clsView* mp_oCurrentView;
	clsGraphics* clsG;
	clsStyle* mp_oStyle;
	clsImage* mp_oImage;
	clsTextBox* mp_oTextBox;
	Gdiplus::Graphics* mp_oGraphics;
	Gdiplus::Bitmap* mp_oBitmap;
	clsCultureInfo* mp_oCulture;
	clsToolTip* mp_oToolTip;				
	CString mp_sModuleCompletePath;				
	CString mp_sVersion;				
	clsRows* mp_oRows;				
	clsTasks* mp_oTasks;				
	clsColumns* mp_oColumns;				
	clsStyles* mp_oStyles;				
	clsLayers* mp_oLayers;				
	clsPercentages* mp_oPercentages;							
	clsTimeBlocks* mp_oTimeBlocks;
	clsPredecessors* mp_oPredecessors;
	clsViews* mp_oViews;				
	clsSplitter* mp_oSplitter;				
	clsPrinter* mp_oPrinter;				
	clsTreeview* mp_oTreeview;				
	clsDrawing* mp_oDrawing;				
	clsMath* mp_oMathLib;							
	clsVerticalScrollBar* mp_oVerticalScrollBar;				
	clsHorizontalScrollBar* mp_oHorizontalScrollBar;
	clsTierAppearance* mp_oTierAppearance;
	clsTierFormat* mp_oTierFormat;
	clsScrollBarSeparator* mp_oScrollBarSeparator;
	clsTimeLineScrollBar* mp_oTimeLineScrollBar;
    clsProgressLine* mp_oProgressLine;
	clsConfiguration* mp_oConfiguration;


	ToolTipEventArgs* mp_oToolTipEventArgs;				
	ObjectAddedEventArgs* mp_oObjectAddedEventArgs;				
	CustomTierDrawEventArgs* mp_oCustomTierDrawEventArgs;				
	MouseEventArgs* mp_oMouseEventArgs;				
	KeyEventArgs* mp_oKeyEventArgs;				
	ScrollEventArgs* mp_oScrollEventArgs;				
	DrawEventArgs* mp_oDrawEventArgs;				
	PredecessorDrawEventArgs* mp_oPredecessorDrawEventArgs;				
	ObjectSelectedEventArgs* mp_oObjectSelectedEventArgs;				
	ObjectStateChangedEventArgs* mp_oObjectStateChangedEventArgs;				
	ErrorEventArgs* mp_oErrorEventArgs;				
	NodeEventArgs* mp_oNodeEventArgs;				
	MouseWheelEventArgs* mp_oMouseWheelEventArgs;
	TextEditEventArgs* mp_oTextEditEventArgs;
	PredecessorExceptionEventArgs* mp_oPredecessorExceptionEventArgs;

	CustomTickMarkAreaDrawEventArgs* mp_oCustomTickMarkAreaDrawEventArgs;
    TickMarkTextDrawEventArgs* mp_oTickMarkTextDrawEventArgs;
	PagePrintEventArgs* mp_oPagePrintEventArgs;

	CBrush mp_oBrush;

//Internal Properties (Extensions to ODL)
public:
	LONG GetErrorReports(void);
	void SetErrorReports(LONG Value);
	CString GetCurrentLayer(void);
	void SetCurrentLayer(CString Value);
	CString GetCurrentView(void);
	void SetCurrentView(CString Value);
	clsView* GetCurrentViewObject(void);
	clsToolTip* GetToolTip(void);
	LONG GetScrollBarBehaviour(void);
	void SetScrollBarBehaviour(LONG Value);
	LONG GetTimeBlockBehaviour(void);
	void SetTimeBlockBehaviour(LONG Value);
	LONG GetSelectedTaskIndex(void);
	void SetSelectedTaskIndex(LONG Value);
	LONG GetSelectedColumnIndex(void);
	void SetSelectedColumnIndex(LONG Value);
	LONG GetSelectedRowIndex(void);
	void SetSelectedRowIndex(LONG Value);
	LONG GetSelectedCellIndex(void);
	void SetSelectedCellIndex(LONG Value);
	LONG GetSelectedPercentageIndex(void);
	void SetSelectedPercentageIndex(LONG Value);
	LONG GetTreeviewColumnIndex(void);
	void SetTreeviewColumnIndex(LONG Value);
	LONG GetSelectedPredecessorIndex(void);
	void SetSelectedPredecessorIndex(LONG Value);
	clsCultureInfo* GetCulture(void);
	CString GetModuleCompletePath(void);
	CString GetVersion(void);
	BOOL GetAllowSplitterMove(void);
	void SetAllowSplitterMove(BOOL Value);
	BOOL GetAllowPredecessorAdd(void);
	void SetAllowPredecessorAdd(BOOL Value);
	BOOL GetAllowAdd(void);
	void SetAllowAdd(BOOL Value);
	BOOL GetAllowEdit(void);
	void SetAllowEdit(BOOL Value);
	BOOL GetAllowColumnSize(void);
	void SetAllowColumnSize(BOOL Value);
	BOOL GetAllowColumnMove(void);
	void SetAllowColumnMove(BOOL Value);
	BOOL GetAllowRowSize(void);
	void SetAllowRowSize(BOOL Value);
	BOOL GetAllowRowMove(void);
	void SetAllowRowMove(BOOL Value);
	BOOL GetAllowTimeLineScroll(void);
	void SetAllowTimeLineScroll(BOOL Value);
	BOOL GetEnforcePredecessors(void);
	void SetEnforcePredecessors(BOOL Value);
	LONG GetAddMode(void);
	void SetAddMode(LONG Value);
	LONG GetPredecessorMode(void);
	void SetPredecessorMode(LONG Value);
	LONG GetLayerEnableObjects(void);
	void SetLayerEnableObjects(LONG Value);
	BOOL GetDoubleBuffering(void);
	void SetDoubleBuffering(BOOL Value);
	LONG GetMinRowHeight(void);
	void SetMinRowHeight(LONG Value);
	LONG GetMinColumnWidth(void);
	void SetMinColumnWidth(LONG Value);
	CString GetControlTag(void);
	void SetControlTag(CString Value);
	LONG GetTierAppearanceScope(void);
	void SetTierAppearanceScope(LONG Value);
	LONG GetTierFormatScope(void);
	void SetTierFormatScope(LONG Value);
	CString GetStyleIndex(void);
	void SetStyleIndex(CString Value);
	clsStyle* GetStyle(void);
	clsImage* GetImage(void);
	CString GetImageTag(void);
	void SetImageTag(CString Value);
	LONG GetAddDurationInterval(void);
	void SetAddDurationInterval(LONG Value);
	LONG GetTimeLineScrollBarScope(void);
	void SetTimeLineScrollBarScope(LONG Value);
	LONG GetProgressLineScope(void);
	void SetProgressLineScope(LONG Value);


	clsRows* GetRows(void);
	clsTasks* GetTasks(void);
	clsColumns* GetColumns(void);
	clsStyles* GetStyles(void);
	clsLayers* GetLayers(void);
	clsPercentages* GetPercentages(void);
	clsTimeBlocks* GetTimeBlocks(void);
	clsPredecessors* GetPredecessors(void);
	clsViews* GetViews(void);
	clsSplitter* GetSplitter(void);
	clsPrinter* GetPrinter(void);
	clsTreeview* GetTreeview(void);
	clsDrawing* GetDrawing(void);
	clsMath* GetMathLib(void);
	clsVerticalScrollBar* GetVerticalScrollBar(void);
	clsHorizontalScrollBar* GetHorizontalScrollBar(void);
	clsTierAppearance* GetTierAppearance(void);
	clsTierFormat* GetTierFormat(void);
	clsScrollBarSeparator* GetScrollBarSeparator(void);
	clsConfiguration* GetConfiguration(void);
	clsTimeLineScrollBar* GetTimeLineScrollBar(void);

	TextEditEventArgs* GetTextEditEventArgs(void);
	ToolTipEventArgs* GetToolTipEventArgs(void);
	ObjectAddedEventArgs* GetObjectAddedEventArgs(void);
	CustomTierDrawEventArgs* GetCustomTierDrawEventArgs(void);
	MouseEventArgs* GetMouseEventArgs(void);
	KeyEventArgs* GetKeyEventArgs(void);
	ScrollEventArgs* GetScrollEventArgs(void);
	DrawEventArgs* GetDrawEventArgs(void);
	PredecessorDrawEventArgs* GetPredecessorDrawEventArgs(void);
	ObjectSelectedEventArgs* GetObjectSelectedEventArgs(void);
	ObjectStateChangedEventArgs* GetObjectStateChangedEventArgs(void);
	ErrorEventArgs* GetErrorEventArgs(void);
	NodeEventArgs* GetNodeEventArgs(void);
	MouseWheelEventArgs* GetMouseWheelEventArgs(void);
	PredecessorExceptionEventArgs* GetPredecessorExceptionEventArgs(void);
	CustomTickMarkAreaDrawEventArgs* GetCustomTickMarkAreaDrawEventArgs(void);
	TickMarkTextDrawEventArgs* GetTickMarkTextDrawEventArgs(void);
	PagePrintEventArgs* GetPagePrintEventArgs(void);

//Internal Properties
public:
	BOOL Getf_UserMode(void);
	LONG Getmt_BorderThickness(void);
	LONG Getmt_TableBottom(void);
	LONG Getmt_TopMargin(void);
	LONG Getmt_LeftMargin(void);
	LONG Getmt_RightMargin(void);
	LONG Getmt_BottomMargin(void);
	clsProgressLine* f_ProgressLine(void);
	clsTimeLineScrollBar* f_TimeLineScrollBar(void);


//Internal Methods
public:
	Gdiplus::Color FromOLEColor(LONG clrOLEColor);

	void OnUpdateEdit();

	void VerticalScrollBar_ValueChanged(LONG Offset);
	void HorizontalScrollBar_ValueChanged(LONG Offset);
	void TimeLineScrollBar_ValueChanged(LONG Offset);
	void mp_PositionScrollBars(void);
	void WriteXML(CString url);
	void ReadXML(CString url);
	void SetXML(CString sXML);
	CString GetXML(void);
	void mp_WriteXML(clsXML* oXML);
	void mp_ReadXML(clsXML* oXML);
	void mp_ErrorReport(LONG ErrNumber,CString ErrDescription,CString ErrSource);
	void ClearSelections(void);
	void Clear(void);
	void CheckPredecessors(void);
	void ForceEndTextEdit(void);
	void SaveToImage(CString Path, LONG Format);
	Gdiplus::Bitmap* GetBitmap(void);
	void AboutBox(void);
	void Redraw(void);

	void ApplyTemplate(LONG ControlTemplate, LONG ObjectTemplate);
	void ApplyTemplateSolid(ControlTemplateSolid &ControlTemplate, LONG ObjectTemplate);
	void ApplyTemplateGradient(ControlTemplateGradient &ControlTemplate, LONG ObjectTemplate);
	LONG GetControlTemplate(void);
	LONG GetObjectTemplate(void);

	DWORD mp_GetPrinterCount(void);



	void FireControlClick(void);
	void FireControlDblClick(void);
	void FireControlMouseDown(void);
	void FireControlMouseMove(void);
	void FireControlMouseUp(void);
	void FireControlMouseHover(void);
	void FireControlKeyDown(void);
	void FireControlKeyUp(void);
	void FireControlKeyPress(void);
	void FireControlMouseWheel(void);
	void FireBeginTextEdit(void);
	void FireEndTextEdit(void);
	void FireDraw(void);
	void FirePredecessorDraw(void);
	void FireCustomTierDraw(void);
	void FireTierTextDraw(void);
	void FireObjectAdded(void);
	void FireObjectSelected(void);
	void FireBeginObjectChange(void);
	void FireObjectChange(void);
	void FireEndObjectChange(void);
	void FireCompleteObjectChange(void);
	void FireActiveGanttError(void);
	void FireControlScroll(void);
	void FireNodeChecked(void);
	void FireNodeExpanded(void);
	void FireControlDraw(void);
	void FireControlRedrawn(void);
	void FireTimeLineChanged(void);
	void FireToolTipOnMouseHover(LONG EventTarget);
	void FireToolTipOnMouseMove(LONG Operation);
	void FireOnMouseHoverToolTipDraw(LONG EventTarget);
	void FireOnMouseMoveToolTipDraw(LONG Operation);
	void FirePredecessorException(void);
	void FireCustomTickMarkAreaDraw(void);
	void FireCustomTickMarkDraw(void);
	void FireTickMarkTextDraw(void);
	void FirePagePrint(void);

	clsTimeBlocks* TempTimeBlocks(void);
	void mp_Draw(void);
	void mp_DrawDesignMode(void);
	void mp_DrawDebugMetrics(void);

	clsGDIGraphics* GetGraphicsObject(void);
	void ReleaseGraphicsObject(void);
	void f_Draw(void);
	


// Event Inline Functions
protected:
	void odl_ControlClick(IDispatch* e)
	{
		FireEvent(1, EVENT_PARAM(VTS_DISPATCH),e);
	}

	void odl_ControlDblClick(IDispatch* e)
	{
		FireEvent(2, EVENT_PARAM(VTS_DISPATCH),e);
	}

	void odl_ControlMouseDown(IDispatch* e)
	{
		FireEvent(3, EVENT_PARAM(VTS_DISPATCH),e);
	}

	void odl_ControlMouseMove(IDispatch* e)
	{
		FireEvent(4, EVENT_PARAM(VTS_DISPATCH),e);
	}

	void odl_ControlMouseUp(IDispatch* e)
	{
		FireEvent(5, EVENT_PARAM(VTS_DISPATCH),e);
	}

	void odl_ControlMouseHover(IDispatch* e)
	{
		FireEvent(6, EVENT_PARAM(VTS_DISPATCH),e);
	}

	void odl_ControlKeyDown(IDispatch* e)
	{
		FireEvent(7, EVENT_PARAM(VTS_DISPATCH),e);
	}

	void odl_ControlKeyPress(IDispatch* e)
	{
		FireEvent(8, EVENT_PARAM(VTS_DISPATCH),e);
	}

	void odl_ControlKeyUp(IDispatch* e)
	{
		FireEvent(9, EVENT_PARAM(VTS_DISPATCH),e);
	}

	void odl_ControlScroll(IDispatch* e)
	{
		FireEvent(10, EVENT_PARAM(VTS_DISPATCH),e);
	}

	void odl_Draw(IDispatch* e)
	{
		FireEvent(11, EVENT_PARAM(VTS_DISPATCH),e);
	}

	void odl_PredecessorDraw(IDispatch* e)
	{
		FireEvent(12, EVENT_PARAM(VTS_DISPATCH),e);
	}

	void odl_CustomTierDraw(IDispatch* e)
	{
		FireEvent(13, EVENT_PARAM(VTS_DISPATCH),e);
	}

	void odl_TierTextDraw(IDispatch* e)
	{
		FireEvent(14, EVENT_PARAM(VTS_DISPATCH),e);
	}

	void odl_ObjectAdded(IDispatch* e)
	{
		FireEvent(15, EVENT_PARAM(VTS_DISPATCH),e);
	}

	void odl_ObjectSelected(IDispatch* e)
	{
		FireEvent(16, EVENT_PARAM(VTS_DISPATCH),e);
	}

	void odl_BeginObjectChange(IDispatch* e)
	{
		FireEvent(17, EVENT_PARAM(VTS_DISPATCH),e);
	}

	void odl_ObjectChange(IDispatch* e)
	{
		FireEvent(18, EVENT_PARAM(VTS_DISPATCH),e);
	}

	void odl_EndObjectChange(IDispatch* e)
	{
		FireEvent(19, EVENT_PARAM(VTS_DISPATCH),e);
	}

	void odl_CompleteObjectChange(IDispatch* e)
	{
		FireEvent(20, EVENT_PARAM(VTS_DISPATCH),e);
	}

	void odl_ActiveGanttError(IDispatch* e)
	{
		FireEvent(21, EVENT_PARAM(VTS_DISPATCH),e);
	}

	void odl_NodeChecked(IDispatch* e)
	{
		FireEvent(22, EVENT_PARAM(VTS_DISPATCH),e);
	}

	void odl_NodeExpanded(IDispatch* e)
	{
		FireEvent(23, EVENT_PARAM(VTS_DISPATCH),e);
	}

	void odl_TimeLineChanged(void)
	{
		FireEvent(24, EVENT_PARAM(VTS_NONE));
	}

	void odl_ControlRedrawn(void)
	{
		FireEvent(25, EVENT_PARAM(VTS_NONE));
	}

	void odl_ControlDraw(void)
	{
		FireEvent(26, EVENT_PARAM(VTS_NONE));
	}

	void odl_ToolTipOnMouseHover(IDispatch* e)
	{
		FireEvent(27, EVENT_PARAM(VTS_DISPATCH),e);
	}

	void odl_ToolTipOnMouseMove(IDispatch* e)
	{
		FireEvent(28, EVENT_PARAM(VTS_DISPATCH),e);
	}

	void odl_OnMouseHoverToolTipDraw(IDispatch* e)
	{
		FireEvent(29, EVENT_PARAM(VTS_DISPATCH),e);
	}

	void odl_OnMouseMoveToolTipDraw(IDispatch* e)
	{
		FireEvent(30, EVENT_PARAM(VTS_DISPATCH),e);
	}

	void odl_ControlMouseWheel(IDispatch* e)
	{
		FireEvent(31, EVENT_PARAM(VTS_DISPATCH),e);
	}

	void odl_BeginTextEdit(IDispatch* e)
	{
		FireEvent(32, EVENT_PARAM(VTS_DISPATCH),e);
	}

	void odl_EndTextEdit(IDispatch* e)
	{
		FireEvent(33, EVENT_PARAM(VTS_DISPATCH),e);
	}

	void odl_PredecessorException(IDispatch* e)
	{
		FireEvent(34, EVENT_PARAM(VTS_DISPATCH),e);
	}

	void odl_CustomTickMarkAreaDraw(IDispatch* e)
	{
		FireEvent(35, EVENT_PARAM(VTS_DISPATCH),e);
	}

	void odl_CustomTickMarkDraw(IDispatch* e)
	{
		FireEvent(36, EVENT_PARAM(VTS_DISPATCH),e);
	}

	void odl_TickMarkTextDraw(IDispatch* e)
	{
		FireEvent(37, EVENT_PARAM(VTS_DISPATCH),e);
	}

	void odl_PagePrint(IDispatch* e)
	{
		FireEvent(38, EVENT_PARAM(VTS_DISPATCH),e);
	}

public:
	void mp_CHKPXPStart(Color clrBackground);
	void mp_CHKPXPScrollButtons(void);
	void mp_CHKPXPLines(void);
	void mp_CHKPXPButtons(void);
	void mp_CHKPXPArrows(void);
	void mp_CHKPXPFigures(void);
	void mp_CHKPXPGradients(void);
	void mp_CHKPXPText(void);
	void mp_CHKPXPHatch(void);


protected:
	DECLARE_OLECREATE_EX(CActiveGanttVCCtl)    // Class factory and guid
	//BEGIN_OLEFACTORY(CActiveGanttVCCtl)
	//	virtual BOOL VerifyUserLicense();
	//	virtual BOOL GetLicenseKey(DWORD, BSTR FAR*);
	//	virtual BOOL VerifyLicenseKey( BSTR bstrKey );
	//END_OLEFACTORY(CActiveGanttVCCtl)

	DECLARE_OLETYPELIB(CActiveGanttVCCtl)
	DECLARE_PROPPAGEIDS(CActiveGanttVCCtl)
	DECLARE_OLECTLTYPE(CActiveGanttVCCtl)

protected:
	DECLARE_DISPATCH_MAP()
	DECLARE_MESSAGE_MAP()
	DECLARE_EVENT_MAP()
public:
	afx_msg HBRUSH OnCtlColor(CDC* pDC, CWnd* pWnd, UINT nCtlColor);


	LONG odl_GetErrorReports(void)
	{
	
		return GetErrorReports();
	}

	void odl_SetErrorReports(LONG Value)
	{
	
		SetErrorReports(Value);
	}

	BSTR odl_GetCurrentLayer(void)
	{
	
		return GetCurrentLayer().AllocSysString();
	}

	void odl_SetCurrentLayer(LPCTSTR Value)
	{
	
		SetCurrentLayer(Value);
	}

	BSTR odl_GetCurrentView(void)
	{
	
		return GetCurrentView().AllocSysString();
	}

	void odl_SetCurrentView(LPCTSTR Value)
	{
	
		SetCurrentView(Value);
	}

	IDispatch* odl_GetCurrentViewObject(void)
	{
	
		return GetCurrentViewObject()->GetIDispatch(TRUE);
	}

	IDispatch* odl_GetToolTip(void)
	{
	
		return GetToolTip()->GetIDispatch(TRUE);
	}

	LONG odl_GetScrollBarBehaviour(void)
	{
	
		return GetScrollBarBehaviour();
	}

	void odl_SetScrollBarBehaviour(LONG Value)
	{
	
		SetScrollBarBehaviour(Value);
	}

	LONG odl_GetTimeBlockBehaviour(void)
	{
	
		return GetTimeBlockBehaviour();
	}

	void odl_SetTimeBlockBehaviour(LONG Value)
	{
	
		SetTimeBlockBehaviour(Value);
	}

	LONG odl_GetSelectedTaskIndex(void)
	{
	
		return GetSelectedTaskIndex();
	}

	void odl_SetSelectedTaskIndex(LONG Value)
	{
	
		SetSelectedTaskIndex(Value);
	}

	LONG odl_GetSelectedColumnIndex(void)
	{
	
		return GetSelectedColumnIndex();
	}

	void odl_SetSelectedColumnIndex(LONG Value)
	{
	
		SetSelectedColumnIndex(Value);
	}

	LONG odl_GetSelectedRowIndex(void)
	{
	
		return GetSelectedRowIndex();
	}

	void odl_SetSelectedRowIndex(LONG Value)
	{
	
		SetSelectedRowIndex(Value);
	}

	LONG odl_GetSelectedCellIndex(void)
	{
	
		return GetSelectedCellIndex();
	}

	void odl_SetSelectedCellIndex(LONG Value)
	{
	
		SetSelectedCellIndex(Value);
	}

	LONG odl_GetSelectedPercentageIndex(void)
	{
	
		return GetSelectedPercentageIndex();
	}

	void odl_SetSelectedPercentageIndex(LONG Value)
	{
	
		SetSelectedPercentageIndex(Value);
	}

	LONG odl_GetTreeviewColumnIndex(void)
	{
	
		return GetTreeviewColumnIndex();
	}

	void odl_SetTreeviewColumnIndex(LONG Value)
	{
	
		SetTreeviewColumnIndex(Value);
	}

	LONG odl_GetSelectedPredecessorIndex(void)
	{
	
		return GetSelectedPredecessorIndex();
	}

	void odl_SetSelectedPredecessorIndex(LONG Value)
	{
	
		SetSelectedPredecessorIndex(Value);
	}

	IDispatch* odl_GetCulture(void)
	{
	
		return GetCulture()->GetIDispatch(TRUE);
	}

	BSTR odl_GetModuleCompletePath(void)
	{
	
		return GetModuleCompletePath().AllocSysString();
	}

	BSTR odl_GetVersion(void)
	{
	
		return GetVersion().AllocSysString();
	}

	BOOL odl_GetAllowSplitterMove(void)
	{
	
		return GetAllowSplitterMove();
	}

	void odl_SetAllowSplitterMove(BOOL Value)
	{
	
	
		SetAllowSplitterMove(Value);
	}

	BOOL odl_GetAllowPredecessorAdd(void)
	{
	
		return GetAllowPredecessorAdd();
	}

	void odl_SetAllowPredecessorAdd(BOOL Value)
	{
	
	
		SetAllowPredecessorAdd(Value);
	}

	BOOL odl_GetAllowAdd(void)
	{
	
		return GetAllowAdd();
	}

	void odl_SetAllowAdd(BOOL Value)
	{
	
	
		SetAllowAdd(Value);
	}

	BOOL odl_GetAllowEdit(void)
	{
	
		return GetAllowEdit();
	}

	void odl_SetAllowEdit(BOOL Value)
	{
	
	
		SetAllowEdit(Value);
	}

	BOOL odl_GetAllowColumnSize(void)
	{
	
		return GetAllowColumnSize();
	}

	void odl_SetAllowColumnSize(BOOL Value)
	{
	
	
		SetAllowColumnSize(Value);
	}

	BOOL odl_GetAllowColumnMove(void)
	{
	
		return GetAllowColumnMove();
	}

	void odl_SetAllowColumnMove(BOOL Value)
	{
	
	
		SetAllowColumnMove(Value);
	}

	BOOL odl_GetAllowRowSize(void)
	{
	
		return GetAllowRowSize();
	}

	void odl_SetAllowRowSize(BOOL Value)
	{
	
	
		SetAllowRowSize(Value);
	}

	BOOL odl_GetAllowRowMove(void)
	{
	
		return GetAllowRowMove();
	}

	void odl_SetAllowRowMove(BOOL Value)
	{
	
	
		SetAllowRowMove(Value);
	}

	BOOL odl_GetAllowTimeLineScroll(void)
	{
	
		return GetAllowTimeLineScroll();
	}

	void odl_SetAllowTimeLineScroll(BOOL Value)
	{
	
		SetAllowTimeLineScroll(Value);
	}

	BOOL odl_GetEnforcePredecessors(void)
	{
	
		return GetEnforcePredecessors();
	}

	void odl_SetEnforcePredecessors(BOOL Value)
	{
	
		SetEnforcePredecessors(Value);
	}

	LONG odl_GetAddMode(void)
	{
	
		return GetAddMode();
	}

	void odl_SetAddMode(LONG Value)
	{
	
		SetAddMode(Value);
	}

	LONG odl_GetPredecessorMode(void)
	{
	
		return GetPredecessorMode();
	}

	void odl_SetPredecessorMode(LONG Value)
	{
	
		SetPredecessorMode(Value);
	}

	LONG odl_GetLayerEnableObjects(void)
	{
	
		return GetLayerEnableObjects();
	}

	void odl_SetLayerEnableObjects(LONG Value)
	{
	
		SetLayerEnableObjects(Value);
	}

	BOOL odl_GetDoubleBuffering(void)
	{
	
		return GetDoubleBuffering();
	}

	void odl_SetDoubleBuffering(BOOL Value)
	{
	
	
		SetDoubleBuffering(Value);
	}

	LONG odl_GetMinRowHeight(void)
	{
	
		return GetMinRowHeight();
	}

	void odl_SetMinRowHeight(LONG Value)
	{
	
		SetMinRowHeight(Value);
	}

	LONG odl_GetMinColumnWidth(void)
	{
	
		return GetMinColumnWidth();
	}

	void odl_SetMinColumnWidth(LONG Value)
	{
	
		SetMinColumnWidth(Value);
	}

	BSTR odl_GetPrinterErrorMessage(void)
	{
		return mp_sPrinterErrorMessage.AllocSysString();
	}

	BSTR odl_GetControlTag(void)
	{
		return GetControlTag().AllocSysString();
	}

	void odl_SetControlTag(LPCTSTR Value)
	{
	
		SetControlTag(Value);
	}

	LONG odl_GetTierAppearanceScope(void)
	{
	
		return GetTierAppearanceScope();
	}

	void odl_SetTierAppearanceScope(LONG Value)
	{
	
		SetTierAppearanceScope(Value);
	}

	LONG odl_GetTierFormatScope(void)
	{
	
		return GetTierFormatScope();
	}

	void odl_SetTierFormatScope(LONG Value)
	{
	
		SetTierFormatScope(Value);
	}

	LONG odl_GetTimeLineScrollBarScope(void)
	{
	
		return GetTimeLineScrollBarScope();
	}

	void odl_SetTimeLineScrollBarScope(LONG Value)
	{
	
		SetTimeLineScrollBarScope(Value);
	}

	LONG odl_GetProgressLineScope(void)
	{
	
		return GetProgressLineScope();
	}

	void odl_SetProgressLineScope(LONG Value)
	{
	
		SetProgressLineScope(Value);
	}

	BSTR odl_GetStyleIndex(void)
	{
	
		return GetStyleIndex().AllocSysString();
	}

	void odl_SetStyleIndex(LPCTSTR Value)
	{
	
		SetStyleIndex(Value);
	}

	IDispatch* odl_GetStyle(void)
	{
	
		return GetStyle()->GetIDispatch(TRUE);
	}

	IDispatch* odl_GetImage(void)
	{
	
		return GetImage()->GetIDispatch(TRUE);
	}

	BSTR odl_GetImageTag(void)
	{
	
		return GetImageTag().AllocSysString();
	}

	void odl_SetImageTag(LPCTSTR Value)
	{
	
		SetImageTag(Value);
	}

	IDispatch* odl_GetRows(void)
	{
	
		return GetRows()->GetIDispatch(TRUE);
	}

	IDispatch* odl_GetTasks(void)
	{
	
		return GetTasks()->GetIDispatch(TRUE);
	}

	IDispatch* odl_GetColumns(void)
	{
	
		return GetColumns()->GetIDispatch(TRUE);
	}

	IDispatch* odl_GetStyles(void)
	{
	
		return GetStyles()->GetIDispatch(TRUE);
	}

	IDispatch* odl_GetLayers(void)
	{
	
		return GetLayers()->GetIDispatch(TRUE);
	}

	IDispatch* odl_GetPercentages(void)
	{
	
		return GetPercentages()->GetIDispatch(TRUE);
	}

	IDispatch* odl_GetTimeBlocks(void)
	{
	
		return GetTimeBlocks()->GetIDispatch(TRUE);
	}

	IDispatch* odl_GetPredecessors(void)
	{
	
		return GetPredecessors()->GetIDispatch(TRUE);
	}

	IDispatch* odl_GetViews(void)
	{
	
		return GetViews()->GetIDispatch(TRUE);
	}

	IDispatch* odl_GetSplitter(void)
	{
	
		return GetSplitter()->GetIDispatch(TRUE);
	}

	IDispatch* odl_GetPrinter(void)
	{
	
		return GetPrinter()->GetIDispatch(TRUE);
	}

	IDispatch* odl_GetTreeview(void)
	{
	
		return GetTreeview()->GetIDispatch(TRUE);
	}

	IDispatch* odl_GetDrawing(void)
	{
	
		return GetDrawing()->GetIDispatch(TRUE);
	}

	IDispatch* odl_GetMathLib(void)
	{
	
		return GetMathLib()->GetIDispatch(TRUE);
	}

	IDispatch* odl_GetVerticalScrollBar(void)
	{
	
		return GetVerticalScrollBar()->GetIDispatch(TRUE);
	}

	IDispatch* odl_GetHorizontalScrollBar(void)
	{
	
		return GetHorizontalScrollBar()->GetIDispatch(TRUE);
	}

	IDispatch* odl_GetTierAppearance(void)
	{
	
		return GetTierAppearance()->GetIDispatch(TRUE);
	}

	IDispatch* odl_GetTierFormat(void)
	{
	
		return GetTierFormat()->GetIDispatch(TRUE);
	}

	IDispatch* odl_GetScrollBarSeparator(void)
	{
	
		return GetScrollBarSeparator()->GetIDispatch(TRUE);
	}

	IDispatch* odl_GetTextEditEventArgs(void)
	{
	
		return GetTextEditEventArgs()->GetIDispatch(TRUE);
	}

	IDispatch* odl_GetToolTipEventArgs(void)
	{
	
		return GetToolTipEventArgs()->GetIDispatch(TRUE);
	}

	IDispatch* odl_GetObjectAddedEventArgs(void)
	{
	
		return GetObjectAddedEventArgs()->GetIDispatch(TRUE);
	}

	IDispatch* odl_GetCustomTierDrawEventArgs(void)
	{
	
		return GetCustomTierDrawEventArgs()->GetIDispatch(TRUE);
	}

	IDispatch* odl_GetMouseEventArgs(void)
	{
	
		return GetMouseEventArgs()->GetIDispatch(TRUE);
	}

	IDispatch* odl_GetKeyEventArgs(void)
	{
	
		return GetKeyEventArgs()->GetIDispatch(TRUE);
	}

	IDispatch* odl_GetScrollEventArgs(void)
	{
	
		return GetScrollEventArgs()->GetIDispatch(TRUE);
	}

	IDispatch* odl_GetDrawEventArgs(void)
	{
	
		return GetDrawEventArgs()->GetIDispatch(TRUE);
	}

	IDispatch* odl_GetPredecessorExceptionEventArgs(void)
	{
	
		return GetPredecessorExceptionEventArgs()->GetIDispatch(TRUE);
	}

	IDispatch* odl_GetPredecessorDrawEventArgs(void)
	{
	
		return GetPredecessorDrawEventArgs()->GetIDispatch(TRUE);
	}

	IDispatch* odl_GetObjectSelectedEventArgs(void)
	{
	
		return GetObjectSelectedEventArgs()->GetIDispatch(TRUE);
	}

	IDispatch* odl_GetObjectStateChangedEventArgs(void)
	{
	
		return GetObjectStateChangedEventArgs()->GetIDispatch(TRUE);
	}

	IDispatch* odl_GetErrorEventArgs(void)
	{
	
		return GetErrorEventArgs()->GetIDispatch(TRUE);
	}

	IDispatch* odl_GetNodeEventArgs(void)
	{
	
		return GetNodeEventArgs()->GetIDispatch(TRUE);
	}

	IDispatch* odl_GetMouseWheelEventArgs(void)
	{
	
		return GetMouseWheelEventArgs()->GetIDispatch(TRUE);
	}

	IDispatch* odl_GetTimeLineScrollBar(void)
	{
	
		return mp_oTimeLineScrollBar->GetIDispatch(TRUE);
	}

	IDispatch* odl_GetProgressLine(void)
	{
	
		return mp_oProgressLine->GetIDispatch(TRUE);
	}

	IDispatch* odl_GetConfiguration(void)
	{
	
		return mp_oConfiguration->GetIDispatch(TRUE);
	}

	IDispatch* odl_GetIDispatch(void)
	{
	
		return this->GetIDispatch(TRUE);
	}

	void odl_WriteXML(LPCTSTR url)
	{
	
		WriteXML(url);
	}

	void odl_ReadXML(LPCTSTR url)
	{
	
		ReadXML(url);
	}

	void odl_SetXML(LPCTSTR sXML)
	{
	
		SetXML(sXML);
	}

	BSTR odl_GetXML(void)
	{
	
		return GetXML().AllocSysString();
	}

	void odl_ClearSelections(void)
	{
	
		ClearSelections();
	}

	void odl_Clear(void)
	{
	
		Clear();
	}

	void odl_ForceEndTextEdit(void)
	{
	
		ForceEndTextEdit();
	}

	void odl_CheckPredecessors(void)
	{
	
		CheckPredecessors();
	}

	void odl_SaveToImage(LPCTSTR Path, LONG Format)
	{
	
		SaveToImage(Path, Format);
	}

	void odl_AboutBox(void)
	{
	
		AboutBox();
	}

	void odl_Redraw(void)
	{
	
		Redraw();
	}

	LONG odl_GetAddDurationInterval(void)
	{
	
		return GetAddDurationInterval();
	}

	void odl_SetAddDurationInterval(LONG Value)
	{
	
		SetAddDurationInterval(Value);
	}

	void odl_ApplyTemplate(LONG ControlTemplate, LONG ObjectTemplate)
	{
	
		ApplyTemplate(ControlTemplate, ObjectTemplate);
	}

	void odl_ApplyTemplate_Solid(IDispatch* ControlTemplate, LONG ObjectTemplate)
	{
	
		ControlTemplateSolid* oControlTemplateSolid = (ControlTemplateSolid*)CCmdTarget::FromIDispatch(ControlTemplate);
		ApplyTemplateSolid(*oControlTemplateSolid, ObjectTemplate);
	}

	void odl_ApplyTemplate_Gradient(IDispatch* ControlTemplate, LONG ObjectTemplate)
	{
	
		ControlTemplateGradient* oControlTemplateGradient = (ControlTemplateGradient*)CCmdTarget::FromIDispatch(ControlTemplate);
		ApplyTemplateGradient(*oControlTemplateGradient, ObjectTemplate);
		//delete oControlTemplateGradient;
		//oControlTemplateGradient = NULL;
	}

	LONG odl_GetControlTemplate(void)
	{
		return GetControlTemplate();
	}

	LONG odl_GetObjectTemplate(void)
	{
		return GetObjectTemplate();
	}

};