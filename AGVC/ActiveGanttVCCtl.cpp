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
#include "ActiveGanttVCPpg.h"
#include "clsXML.h"
//#include "TimeCounter.h"

IMPLEMENT_DYNCREATE(CActiveGanttVCCtl, COleControl)
static xAuxInitGDIPlus GDI_Plus_Controler;
#include "VersionInfo.h"

//{3F906ECE-6A26-4AD1-85D4-BC7B436FBA8A}
const IID BASED_CODE IID_DActiveGanttVC = { 0x3F906ECE, 0x6A26, 0x4AD1, { 0x85, 0xD4, 0xBC, 0x7B, 0x43, 0x6F, 0xBA, 0x8A} };

//{B63F7141-31C8-4BC2-A81D-D90188AB9DC9}
const IID BASED_CODE IID_DActiveGanttVCEvents = { 0xB63F7141, 0x31C8, 0x4BC2, { 0xA8, 0x1D, 0xD9, 0x01, 0x88, 0xAB, 0x9D, 0xC9} };

//{688B95D3-C09A-4E7D-8F01-F4BAC59CCA18}
IMPLEMENT_OLECREATE_EX(CActiveGanttVCCtl, "ActiveGanttVC.ActiveGanttVCCtl", 0x688B95D3, 0xC09A, 0x4E7D, 0x8F, 0x01, 0xF4, 0xBA, 0xC5, 0x9C, 0xCA, 0x18)

BEGIN_DISPATCH_MAP(CActiveGanttVCCtl, COleControl)
	DISP_PROPERTY_EX_ID(CActiveGanttVCCtl, "ErrorReports", 1, odl_GetErrorReports, odl_SetErrorReports, VT_I4) 
	DISP_PROPERTY_EX_ID(CActiveGanttVCCtl, "CurrentLayer", 2, odl_GetCurrentLayer, odl_SetCurrentLayer, VT_BSTR) 
	DISP_PROPERTY_EX_ID(CActiveGanttVCCtl, "CurrentView", 3, odl_GetCurrentView, odl_SetCurrentView, VT_BSTR) 
	DISP_PROPERTY_EX_ID(CActiveGanttVCCtl, "CurrentViewObject", 4, odl_GetCurrentViewObject, SetNotSupported, VT_DISPATCH) 
	DISP_PROPERTY_EX_ID(CActiveGanttVCCtl, "ToolTip", 5, odl_GetToolTip, SetNotSupported, VT_DISPATCH) 
	DISP_PROPERTY_EX_ID(CActiveGanttVCCtl, "ScrollBarBehaviour", 6, odl_GetScrollBarBehaviour, odl_SetScrollBarBehaviour, VT_I4) 
	DISP_PROPERTY_EX_ID(CActiveGanttVCCtl, "TimeBlockBehaviour", 7, odl_GetTimeBlockBehaviour, odl_SetTimeBlockBehaviour, VT_I4) 
	DISP_PROPERTY_EX_ID(CActiveGanttVCCtl, "SelectedTaskIndex", 8, odl_GetSelectedTaskIndex, odl_SetSelectedTaskIndex, VT_I4) 
	DISP_PROPERTY_EX_ID(CActiveGanttVCCtl, "SelectedColumnIndex", 9, odl_GetSelectedColumnIndex, odl_SetSelectedColumnIndex, VT_I4) 
	DISP_PROPERTY_EX_ID(CActiveGanttVCCtl, "SelectedRowIndex", 10, odl_GetSelectedRowIndex, odl_SetSelectedRowIndex, VT_I4) 
	DISP_PROPERTY_EX_ID(CActiveGanttVCCtl, "SelectedCellIndex", 11, odl_GetSelectedCellIndex, odl_SetSelectedCellIndex, VT_I4) 
	DISP_PROPERTY_EX_ID(CActiveGanttVCCtl, "SelectedPercentageIndex", 12, odl_GetSelectedPercentageIndex, odl_SetSelectedPercentageIndex, VT_I4) 
	DISP_PROPERTY_EX_ID(CActiveGanttVCCtl, "TreeviewColumnIndex", 13, odl_GetTreeviewColumnIndex, odl_SetTreeviewColumnIndex, VT_I4) 
	DISP_PROPERTY_EX_ID(CActiveGanttVCCtl, "Culture", 14, odl_GetCulture, SetNotSupported, VT_DISPATCH) 
	DISP_PROPERTY_EX_ID(CActiveGanttVCCtl, "ModuleCompletePath", 15, odl_GetModuleCompletePath, SetNotSupported, VT_BSTR) 
	DISP_PROPERTY_EX_ID(CActiveGanttVCCtl, "Version", 16, odl_GetVersion, SetNotSupported, VT_BSTR) 
	DISP_PROPERTY_EX_ID(CActiveGanttVCCtl, "AllowSplitterMove", 17, odl_GetAllowSplitterMove, odl_SetAllowSplitterMove, VT_BOOL) 
	DISP_PROPERTY_EX_ID(CActiveGanttVCCtl, "AllowPredecessorAdd", 18, odl_GetAllowPredecessorAdd, odl_SetAllowPredecessorAdd, VT_BOOL) 
	DISP_PROPERTY_EX_ID(CActiveGanttVCCtl, "AllowAdd", 19, odl_GetAllowAdd, odl_SetAllowAdd, VT_BOOL) 
	DISP_PROPERTY_EX_ID(CActiveGanttVCCtl, "AllowEdit", 20, odl_GetAllowEdit, odl_SetAllowEdit, VT_BOOL) 
	DISP_PROPERTY_EX_ID(CActiveGanttVCCtl, "AllowColumnSize", 21, odl_GetAllowColumnSize, odl_SetAllowColumnSize, VT_BOOL) 
	DISP_PROPERTY_EX_ID(CActiveGanttVCCtl, "AllowColumnMove", 22, odl_GetAllowColumnMove, odl_SetAllowColumnMove, VT_BOOL) 
	DISP_PROPERTY_EX_ID(CActiveGanttVCCtl, "AllowRowSize", 23, odl_GetAllowRowSize, odl_SetAllowRowSize, VT_BOOL) 
	DISP_PROPERTY_EX_ID(CActiveGanttVCCtl, "AllowRowMove", 24, odl_GetAllowRowMove, odl_SetAllowRowMove, VT_BOOL) 
	DISP_PROPERTY_EX_ID(CActiveGanttVCCtl, "AllowTimeLineScroll", 25, odl_GetAllowTimeLineScroll, odl_SetAllowTimeLineScroll, VT_BOOL) 
	DISP_PROPERTY_EX_ID(CActiveGanttVCCtl, "AddMode", 26, odl_GetAddMode, odl_SetAddMode, VT_I4) 
	DISP_PROPERTY_EX_ID(CActiveGanttVCCtl, "LayerEnableObjects", 27, odl_GetLayerEnableObjects, odl_SetLayerEnableObjects, VT_I4) 
	DISP_PROPERTY_EX_ID(CActiveGanttVCCtl, "DoubleBuffering", 28, odl_GetDoubleBuffering, odl_SetDoubleBuffering, VT_BOOL) 
	DISP_PROPERTY_EX_ID(CActiveGanttVCCtl, "MinRowHeight", 29, odl_GetMinRowHeight, odl_SetMinRowHeight, VT_I4) 
	DISP_PROPERTY_EX_ID(CActiveGanttVCCtl, "MinColumnWidth", 30, odl_GetMinColumnWidth, odl_SetMinColumnWidth, VT_I4)
	DISP_PROPERTY_EX_ID(CActiveGanttVCCtl, "Rows", 31, odl_GetRows, SetNotSupported, VT_DISPATCH) 
	DISP_PROPERTY_EX_ID(CActiveGanttVCCtl, "Tasks", 32, odl_GetTasks, SetNotSupported, VT_DISPATCH) 
	DISP_PROPERTY_EX_ID(CActiveGanttVCCtl, "Columns", 33, odl_GetColumns, SetNotSupported, VT_DISPATCH) 
	DISP_PROPERTY_EX_ID(CActiveGanttVCCtl, "Styles", 34, odl_GetStyles, SetNotSupported, VT_DISPATCH) 
	DISP_PROPERTY_EX_ID(CActiveGanttVCCtl, "Layers", 35, odl_GetLayers, SetNotSupported, VT_DISPATCH) 
	DISP_PROPERTY_EX_ID(CActiveGanttVCCtl, "Percentages", 36, odl_GetPercentages, SetNotSupported, VT_DISPATCH) 
	DISP_PROPERTY_EX_ID(CActiveGanttVCCtl, "TimeBlocks", 37, odl_GetTimeBlocks, SetNotSupported, VT_DISPATCH) 
	DISP_PROPERTY_EX_ID(CActiveGanttVCCtl, "Views", 38, odl_GetViews, SetNotSupported, VT_DISPATCH) 
	DISP_PROPERTY_EX_ID(CActiveGanttVCCtl, "Splitter", 39, odl_GetSplitter, SetNotSupported, VT_DISPATCH) 
	DISP_PROPERTY_EX_ID(CActiveGanttVCCtl, "Printer", 40, odl_GetPrinter, SetNotSupported, VT_DISPATCH) 
	DISP_PROPERTY_EX_ID(CActiveGanttVCCtl, "Treeview", 41, odl_GetTreeview, SetNotSupported, VT_DISPATCH) 
	DISP_PROPERTY_EX_ID(CActiveGanttVCCtl, "Drawing", 42, odl_GetDrawing, SetNotSupported, VT_DISPATCH) 
	DISP_PROPERTY_EX_ID(CActiveGanttVCCtl, "MathLib", 43, odl_GetMathLib, SetNotSupported, VT_DISPATCH) 
	DISP_PROPERTY_EX_ID(CActiveGanttVCCtl, "VerticalScrollBar", 44, odl_GetVerticalScrollBar, SetNotSupported, VT_DISPATCH) 
	DISP_PROPERTY_EX_ID(CActiveGanttVCCtl, "HorizontalScrollBar", 45, odl_GetHorizontalScrollBar, SetNotSupported, VT_DISPATCH) 
	DISP_PROPERTY_EX_ID(CActiveGanttVCCtl, "ToolTipEventArgs", 46, odl_GetToolTipEventArgs, SetNotSupported, VT_DISPATCH) 
	DISP_PROPERTY_EX_ID(CActiveGanttVCCtl, "ObjectAddedEventArgs", 47, odl_GetObjectAddedEventArgs, SetNotSupported, VT_DISPATCH) 
	DISP_PROPERTY_EX_ID(CActiveGanttVCCtl, "CustomTierDrawEventArgs", 48, odl_GetCustomTierDrawEventArgs, SetNotSupported, VT_DISPATCH) 
	DISP_PROPERTY_EX_ID(CActiveGanttVCCtl, "MouseEventArgs", 49, odl_GetMouseEventArgs, SetNotSupported, VT_DISPATCH) 
	DISP_PROPERTY_EX_ID(CActiveGanttVCCtl, "KeyEventArgs", 50, odl_GetKeyEventArgs, SetNotSupported, VT_DISPATCH) 
	DISP_PROPERTY_EX_ID(CActiveGanttVCCtl, "ScrollEventArgs", 51, odl_GetScrollEventArgs, SetNotSupported, VT_DISPATCH) 
	DISP_PROPERTY_EX_ID(CActiveGanttVCCtl, "DrawEventArgs", 52, odl_GetDrawEventArgs, SetNotSupported, VT_DISPATCH) 
	DISP_PROPERTY_EX_ID(CActiveGanttVCCtl, "PredecessorDrawEventArgs", 53, odl_GetPredecessorDrawEventArgs, SetNotSupported, VT_DISPATCH) 
	DISP_PROPERTY_EX_ID(CActiveGanttVCCtl, "ObjectSelectedEventArgs", 54, odl_GetObjectSelectedEventArgs, SetNotSupported, VT_DISPATCH) 
	DISP_PROPERTY_EX_ID(CActiveGanttVCCtl, "ObjectStateChangedEventArgs", 55, odl_GetObjectStateChangedEventArgs, SetNotSupported, VT_DISPATCH) 
	DISP_PROPERTY_EX_ID(CActiveGanttVCCtl, "ErrorEventArgs", 56, odl_GetErrorEventArgs, SetNotSupported, VT_DISPATCH) 
	DISP_PROPERTY_EX_ID(CActiveGanttVCCtl, "NodeEventArgs", 57, odl_GetNodeEventArgs, SetNotSupported, VT_DISPATCH)
	DISP_PROPERTY_EX_ID(CActiveGanttVCCtl, "ControlTag", 58, odl_GetControlTag, odl_SetControlTag, VT_BSTR)
	DISP_PROPERTY_EX_ID(CActiveGanttVCCtl, "GetIDispatch", 59, odl_GetIDispatch, SetNotSupported, VT_DISPATCH) 
	DISP_PROPERTY_EX_ID(CActiveGanttVCCtl, "MouseWheelEventArgs", 60, odl_GetMouseWheelEventArgs, SetNotSupported, VT_DISPATCH)
	DISP_PROPERTY_EX_ID(CActiveGanttVCCtl, "TierAppearance", 61, odl_GetTierAppearance, SetNotSupported, VT_DISPATCH)
	DISP_PROPERTY_EX_ID(CActiveGanttVCCtl, "TierFormat", 62, odl_GetTierFormat, SetNotSupported, VT_DISPATCH)
	DISP_PROPERTY_EX_ID(CActiveGanttVCCtl, "TierAppearanceScope", 63, odl_GetTierAppearanceScope, odl_SetTierAppearanceScope, VT_I4) 
	DISP_PROPERTY_EX_ID(CActiveGanttVCCtl, "TierFormatScope", 64, odl_GetTierFormatScope, odl_SetTierFormatScope, VT_I4)
	DISP_PROPERTY_EX_ID(CActiveGanttVCCtl, "StyleIndex", 65, odl_GetStyleIndex, odl_SetStyleIndex, VT_BSTR) 
	DISP_PROPERTY_EX_ID(CActiveGanttVCCtl, "Style", 66, odl_GetStyle, SetNotSupported, VT_DISPATCH) 
	DISP_PROPERTY_EX_ID(CActiveGanttVCCtl, "Image", 67, odl_GetImage, SetNotSupported, VT_DISPATCH) 
	DISP_PROPERTY_EX_ID(CActiveGanttVCCtl, "ImageTag", 68, odl_GetImageTag, odl_SetImageTag, VT_BSTR)
	DISP_PROPERTY_EX_ID(CActiveGanttVCCtl, "ScrollBarSeparator", 69, odl_GetScrollBarSeparator, SetNotSupported, VT_DISPATCH)
	DISP_PROPERTY_EX_ID(CActiveGanttVCCtl, "TextEditEventArgs", 70, odl_GetTextEditEventArgs, SetNotSupported, VT_DISPATCH)
	DISP_PROPERTY_EX_ID(CActiveGanttVCCtl, "Predecessors", 71, odl_GetPredecessors, SetNotSupported, VT_DISPATCH)
	DISP_PROPERTY_EX_ID(CActiveGanttVCCtl, "PredecessorExceptionEventArgs", 72, odl_GetPredecessorExceptionEventArgs, SetNotSupported, VT_DISPATCH)
	DISP_PROPERTY_EX_ID(CActiveGanttVCCtl, "SelectedPredecessorIndex", 73, odl_GetSelectedPredecessorIndex, odl_SetSelectedPredecessorIndex, VT_I4)
	DISP_PROPERTY_EX_ID(CActiveGanttVCCtl, "EnforcePredecessors", 74, odl_GetEnforcePredecessors, odl_SetEnforcePredecessors, VT_BOOL)
	DISP_PROPERTY_EX_ID(CActiveGanttVCCtl, "PredecessorMode", 75, odl_GetPredecessorMode, odl_SetPredecessorMode, VT_I4)
	DISP_PROPERTY_EX_ID(CActiveGanttVCCtl, "AddDurationInterval", 76, odl_GetAddDurationInterval, odl_SetAddDurationInterval, VT_I4)
	DISP_FUNCTION_ID(CActiveGanttVCCtl, "WriteXML", 77, odl_WriteXML, VT_EMPTY, VTS_BSTR)
	DISP_FUNCTION_ID(CActiveGanttVCCtl, "ReadXML", 78, odl_ReadXML, VT_EMPTY, VTS_BSTR)
	DISP_FUNCTION_ID(CActiveGanttVCCtl, "SetXML", 79, odl_SetXML, VT_EMPTY, VTS_BSTR) 
	DISP_FUNCTION_ID(CActiveGanttVCCtl, "GetXML", 80, odl_GetXML, VT_BSTR, VTS_NONE)
	DISP_FUNCTION_ID(CActiveGanttVCCtl, "ClearSelections", 81, odl_ClearSelections, VT_EMPTY, VTS_NONE)
	DISP_FUNCTION_ID(CActiveGanttVCCtl, "Clear", 82, odl_Clear, VT_EMPTY, VTS_NONE)
	DISP_FUNCTION_ID(CActiveGanttVCCtl, "SaveToImage", 83, odl_SaveToImage, VT_EMPTY, VTS_BSTR VTS_I4)
	DISP_FUNCTION_ID(CActiveGanttVCCtl, "AboutBox", 84, odl_AboutBox, VT_EMPTY, VTS_NONE)
	DISP_FUNCTION_ID(CActiveGanttVCCtl, "Redraw", 85, odl_Redraw, VT_EMPTY, VTS_NONE)
	DISP_FUNCTION_ID(CActiveGanttVCCtl, "ForceEndTextEdit", 86, odl_ForceEndTextEdit, VT_EMPTY, VTS_NONE)
	DISP_FUNCTION_ID(CActiveGanttVCCtl, "CheckPredecessors", 87, odl_CheckPredecessors, VT_EMPTY, VTS_NONE)
	DISP_PROPERTY_EX_ID(CActiveGanttVCCtl, "TimeLineScrollBar", 88, odl_GetTimeLineScrollBar, SetNotSupported, VT_DISPATCH)
	DISP_PROPERTY_EX_ID(CActiveGanttVCCtl, "ProgressLine", 89, odl_GetProgressLine, SetNotSupported, VT_DISPATCH)
	DISP_PROPERTY_EX_ID(CActiveGanttVCCtl, "TimeLineScrollBarScope", 90, odl_GetTimeLineScrollBarScope, odl_SetTimeLineScrollBarScope, VT_I4)
	DISP_PROPERTY_EX_ID(CActiveGanttVCCtl, "ProgressLineScope", 91, odl_GetProgressLineScope, odl_SetProgressLineScope, VT_I4)
	DISP_PROPERTY_EX_ID(CActiveGanttVCCtl, "Configuration", 92, odl_GetConfiguration, SetNotSupported, VT_DISPATCH)
	DISP_FUNCTION_ID(CActiveGanttVCCtl, "ApplyTemplateSolid", 93, odl_ApplyTemplate_Solid, VT_EMPTY, VTS_DISPATCH VTS_I4)
	DISP_FUNCTION_ID(CActiveGanttVCCtl, "ApplyTemplateGradient", 94, odl_ApplyTemplate_Gradient, VT_EMPTY, VTS_DISPATCH VTS_I4)
	DISP_PROPERTY_EX_ID(CActiveGanttVCCtl, "ControlTemplate", 95, odl_GetControlTemplate, SetNotSupported, VT_I4)
	DISP_PROPERTY_EX_ID(CActiveGanttVCCtl, "ObjectTemplate", 96, odl_GetObjectTemplate, SetNotSupported, VT_I4)
	DISP_FUNCTION_ID(CActiveGanttVCCtl, "ApplyTemplate", 97, odl_ApplyTemplate, VT_EMPTY, VTS_I4 VTS_I4)
	DISP_PROPERTY_EX_ID(CActiveGanttVCCtl, "PrinterErrorMessage", 98, odl_GetPrinterErrorMessage, SetNotSupported, VT_BSTR)
END_DISPATCH_MAP()

BEGIN_EVENT_MAP(CActiveGanttVCCtl, COleControl)
	EVENT_CUSTOM_ID("ControlClick", 1, odl_ControlClick, VTS_DISPATCH) 
	EVENT_CUSTOM_ID("ControlDblClick", 2, odl_ControlDblClick, VTS_DISPATCH) 
	EVENT_CUSTOM_ID("ControlMouseDown", 3, odl_ControlMouseDown, VTS_DISPATCH) 
	EVENT_CUSTOM_ID("ControlMouseMove", 4, odl_ControlMouseMove, VTS_DISPATCH) 
	EVENT_CUSTOM_ID("ControlMouseUp", 5, odl_ControlMouseUp, VTS_DISPATCH) 
	EVENT_CUSTOM_ID("ControlMouseHover", 6, odl_ControlMouseHover, VTS_DISPATCH)
	EVENT_CUSTOM_ID("ControlKeyDown", 7, odl_ControlKeyDown, VTS_DISPATCH) 
	EVENT_CUSTOM_ID("ControlKeyPress", 8, odl_ControlKeyPress, VTS_DISPATCH) 
	EVENT_CUSTOM_ID("ControlKeyUp", 9, odl_ControlKeyUp, VTS_DISPATCH) 
	EVENT_CUSTOM_ID("ControlScroll", 10, odl_ControlScroll, VTS_DISPATCH) 
	EVENT_CUSTOM_ID("Draw", 11, odl_Draw, VTS_DISPATCH) 
	EVENT_CUSTOM_ID("PredecessorDraw", 12, odl_PredecessorDraw, VTS_DISPATCH)
	EVENT_CUSTOM_ID("CustomTierDraw", 13, odl_CustomTierDraw, VTS_DISPATCH)
	EVENT_CUSTOM_ID("TierTextDraw", 14, odl_TierTextDraw, VTS_DISPATCH) 
	EVENT_CUSTOM_ID("ObjectAdded", 15, odl_ObjectAdded, VTS_DISPATCH) 
	EVENT_CUSTOM_ID("ObjectSelected", 16, odl_ObjectSelected, VTS_DISPATCH) 
	EVENT_CUSTOM_ID("BeginObjectChange", 17, odl_BeginObjectChange, VTS_DISPATCH) 
	EVENT_CUSTOM_ID("ObjectChange", 18, odl_ObjectChange, VTS_DISPATCH) 
	EVENT_CUSTOM_ID("EndObjectChange", 19, odl_EndObjectChange, VTS_DISPATCH) 
	EVENT_CUSTOM_ID("CompleteObjectChange", 20, odl_CompleteObjectChange, VTS_DISPATCH) 
	EVENT_CUSTOM_ID("ActiveGanttError", 21, odl_ActiveGanttError, VTS_DISPATCH) 
	EVENT_CUSTOM_ID("NodeChecked", 22, odl_NodeChecked, VTS_DISPATCH) 
	EVENT_CUSTOM_ID("NodeExpanded", 23, odl_NodeExpanded, VTS_DISPATCH) 
	EVENT_CUSTOM_ID("TimeLineChanged", 24, odl_TimeLineChanged, VTS_NONE) 
	EVENT_CUSTOM_ID("ControlRedrawn", 25, odl_ControlRedrawn, VTS_NONE) 
	EVENT_CUSTOM_ID("ControlDraw", 26, odl_ControlDraw, VTS_NONE) 
	EVENT_CUSTOM_ID("ToolTipOnMouseHover", 27, odl_ToolTipOnMouseHover, VTS_DISPATCH) 
	EVENT_CUSTOM_ID("ToolTipOnMouseMove", 28, odl_ToolTipOnMouseMove, VTS_DISPATCH) 
	EVENT_CUSTOM_ID("OnMouseHoverToolTipDraw", 29, odl_OnMouseHoverToolTipDraw, VTS_DISPATCH) 
	EVENT_CUSTOM_ID("OnMouseMoveToolTipDraw", 30, odl_OnMouseMoveToolTipDraw, VTS_DISPATCH)
	EVENT_CUSTOM_ID("ControlMouseWheel", 31, odl_ControlMouseWheel, VTS_DISPATCH)
	EVENT_CUSTOM_ID("BeginTextEdit", 32, odl_BeginTextEdit, VTS_DISPATCH)
	EVENT_CUSTOM_ID("EndTextEdit", 33, odl_EndTextEdit, VTS_DISPATCH)
	EVENT_CUSTOM_ID("PredecessorException", 34, odl_PredecessorException, VTS_DISPATCH)
	EVENT_CUSTOM_ID("CustomTickMarkAreaDraw", 35, odl_CustomTickMarkAreaDraw, VTS_DISPATCH)
	EVENT_CUSTOM_ID("CustomTickMarkDraw", 36, odl_CustomTickMarkDraw, VTS_DISPATCH)
	EVENT_CUSTOM_ID("TickMarkTextDraw", 37, odl_TickMarkTextDraw, VTS_DISPATCH)
	EVENT_CUSTOM_ID("PagePrint", 38, odl_PagePrintEventArgs, VTS_DISPATCH)


END_EVENT_MAP()


BEGIN_MESSAGE_MAP(CActiveGanttVCCtl, COleControl)
	//Key Down
	ON_WM_KEYDOWN()
	ON_WM_SYSKEYDOWN()
	//Key Up
	ON_WM_KEYUP()
	ON_WM_SYSKEYUP()
	//Key Press
	ON_WM_CHAR()
	ON_WM_DEADCHAR()
	ON_WM_SYSDEADCHAR()
	//Mouse Capture
	ON_WM_LBUTTONDBLCLK()
	ON_WM_LBUTTONDOWN()
	ON_WM_LBUTTONUP()
	ON_WM_MBUTTONDBLCLK()
	ON_WM_MBUTTONDOWN()
	ON_WM_MBUTTONUP()
	ON_WM_RBUTTONDBLCLK()
	ON_WM_RBUTTONDOWN()
	ON_WM_RBUTTONUP()
	ON_WM_MOUSEWHEEL()
	ON_WM_MOUSEMOVE()
	//Other Events
	ON_WM_SIZE()
	ON_WM_SETCURSOR()
	ON_WM_CREATE()
	ON_WM_TIMER()
	ON_OLEVERB(AFX_IDS_VERB_PROPERTIES, OnProperties)
	ON_WM_CTLCOLOR()
	ON_EN_UPDATE(15620, OnUpdateEdit)
END_MESSAGE_MAP()

BEGIN_PROPPAGEIDS(CActiveGanttVCCtl, 1)
PROPPAGEID(CActiveGanttVCPropPage::guid)
END_PROPPAGEIDS(CActiveGanttVCCtl)

IMPLEMENT_OLETYPELIB(CActiveGanttVCCtl, _tlid, _wVerMajor, _wVerMinor)

static const DWORD BASED_CODE _dwActiveGanttVCOleMisc =
OLEMISC_SIMPLEFRAME |
OLEMISC_ACTIVATEWHENVISIBLE |
OLEMISC_SETCLIENTSITEFIRST |
OLEMISC_INSIDEOUT |
OLEMISC_CANTLINKINSIDE |
OLEMISC_RECOMPOSEONRESIZE;

IMPLEMENT_OLECTLTYPE(CActiveGanttVCCtl, IDS_ACTIVEGANTTVC, _dwActiveGanttVCOleMisc)

static const WCHAR BASED_CODE _szLicString[] = L"f5d3sw2qahgfsedwesv ";

BOOL CActiveGanttVCCtl::CActiveGanttVCCtlFactory::UpdateRegistry(BOOL bRegister)
{

	if (bRegister)
		return AfxOleRegisterControlClass(
		AfxGetInstanceHandle(),
		m_clsid,
		m_lpszProgID,
		IDS_ACTIVEGANTTVC,
		IDB_ABOUTICON,
		afxRegInsertable | afxRegApartmentThreading,
		_dwActiveGanttVCOleMisc,
		_tlid,
		_wVerMajor,
		_wVerMinor);
	else
		return AfxOleUnregisterClass(m_clsid, m_lpszProgID);
}

//BOOL CActiveGanttVCCtl::CActiveGanttVCCtlFactory::VerifyUserLicense()
//{
//	// Verifies the design-time license for the OLE control.
//	HKEY hKey;
//	TCHAR strValue[255];
//	DWORD dwBufLen;
//	CString sValue;
//	BOOL bReturn = FALSE;
//	if (RegOpenKeyEx(HKEY_CLASSES_ROOT, _T("Licenses\\7E44F611-0E19-4830-AC10-D62E10C1AF42"), 0, KEY_QUERY_VALUE, &hKey) == ERROR_SUCCESS)
//	{
//		RegQueryValueEx(hKey, NULL, NULL, NULL, NULL, &dwBufLen) ;
//		if (RegQueryValueEx(hKey, NULL, NULL, NULL, (LPBYTE) strValue, &dwBufLen) == ERROR_SUCCESS)
//		{
//			sValue = strValue;
//			if (sValue == _T("kefellgeglglmeklmfjegllepktkjlieeemk"))
//			{
//				bReturn = TRUE; //The design time license is valid (TRUE == 1).
//			}
//			else
//			{
//				bReturn = FALSE;
//			}
//		}
//		else
//		{
//			bReturn = FALSE;
//		}
//	}
//	else
//	{
//		bReturn = FALSE;
//	}
//	RegCloseKey(hKey);
//	return bReturn;
//}
//
//BOOL CActiveGanttVCCtl::CActiveGanttVCCtlFactory::GetLicenseKey(DWORD dwReserved, BSTR FAR* pbstrKey)
//{
//	//Requests a unique license key from the control's DLL and stores it in the BSTR pointed to by pbstrKey.
//	if (pbstrKey == NULL)
//	{
//		return FALSE;
//	}
//	*pbstrKey = SysAllocString(_szLicString);
//	return (*pbstrKey != NULL);
//}
//
//BOOL CActiveGanttVCCtl::CActiveGanttVCCtlFactory::VerifyLicenseKey(BSTR bstrKey)
//{
//	//Verifies that the container is licensed to use the OLE control.
//	BOOL bLicensed = FALSE;
//	BSTR bstr = NULL;
//	if ((bstrKey != NULL) && GetLicenseKey(0, &bstr))
//	{
//		ASSERT(bstr != NULL);
//		UINT cch = SysStringByteLen(bstr);
//		if ((cch == SysStringByteLen(bstrKey)) && (memcmp(bstr, bstrKey, cch) == 0))
//		{
//			bLicensed = TRUE;
//		}
//
//		SysFreeString(bstr);
//	}
//	return bLicensed;
//}

CActiveGanttVCCtl::CActiveGanttVCCtl()
{

	InitializeIIDs(&IID_DActiveGanttVC, &IID_DActiveGanttVCEvents);

	mp_lAboutShown = 0;


	mp_lObjectCount = AfxGetStaticModuleState()->m_nObjectCount;

	//Initialize GDI+
	GDI_Plus_Controler.Initialize();

	mp_oBitmap = NULL;
	mp_oGraphics = NULL;

	mp_oTimer = 0;

	mp_sPrinterErrorMessage = _T("");
	mp_bAllowAdd = TRUE;
    mp_bAllowEdit = TRUE;
    mp_bAllowSplitterMove = TRUE;
    mp_bAllowRowSize = TRUE;
    mp_bAllowRowMove = TRUE;
    mp_bAllowColumnSize = TRUE;
    mp_bAllowColumnMove = TRUE;
    mp_bAllowTimeLineScroll = TRUE;
    mp_bAllowPredecessorAdd = TRUE;
    mp_bDoubleBuffering = TRUE;
    mp_bPropertiesRead = FALSE;
    mp_bEnforcePredecessors = FALSE;
    mp_lMinColumnWidth = 5;
    mp_lMinRowHeight = 5;
    mp_lSelectedTaskIndex = 0;
    mp_lSelectedColumnIndex = 0;
    mp_lSelectedRowIndex = 0;
    mp_lSelectedCellIndex = 0;
    mp_lSelectedPercentageIndex = 0;
    mp_lSelectedPredecessorIndex = 0;
    mp_lTreeviewColumnIndex = 0;
    mp_sCurrentLayer = _T("0");
    mp_sCurrentView = _T("");
    mp_yAddMode = AT_TASKADD;
	mp_yAddDurationInterval = IL_SECOND;
    mp_yScrollBarBehaviour  = SB_HIDE;
    mp_yTimeBlockBehaviour = TBB_ROWEXTENTS;
    mp_yLayerEnableObjects = EC_INCURRENTLAYERONLY;
    mp_yErrorReports = RE_MSGBOX;
    mp_yTierAppearanceScope = OS_CONTROL;
    mp_yTierFormatScope = OS_CONTROL;
	mp_yTimeLineScrollBarScope = OS_CONTROL;
	mp_yProgressLineScope = OS_CONTROL;
	mp_yTimeLineScrollBarScope = OS_CONTROL;
    mp_yPredecessorMode = PM_CREATEWARNINGFLAG;
	mp_yControlTemplate = STC_NONE;
    mp_yObjectTemplate = STO_NONE;
    mp_sControlTag = _T("");

	mp_oToolTipEventArgs = new ToolTipEventArgs();
	mp_oObjectAddedEventArgs = new ObjectAddedEventArgs();
	mp_oCustomTierDrawEventArgs = new CustomTierDrawEventArgs();
	mp_oMouseEventArgs = new MouseEventArgs();
	mp_oMouseWheelEventArgs = new MouseWheelEventArgs();
	mp_oKeyEventArgs = new KeyEventArgs();
	mp_oScrollEventArgs = new ScrollEventArgs();
	mp_oDrawEventArgs = new DrawEventArgs();
	mp_oPredecessorDrawEventArgs = new PredecessorDrawEventArgs();
	mp_oObjectSelectedEventArgs = new ObjectSelectedEventArgs();
	mp_oObjectStateChangedEventArgs = new ObjectStateChangedEventArgs();
	mp_oErrorEventArgs = new ErrorEventArgs();
	mp_oNodeEventArgs = new NodeEventArgs();
	mp_oTextEditEventArgs = new TextEditEventArgs();
	mp_oPredecessorExceptionEventArgs = new PredecessorExceptionEventArgs();
	mp_oCustomTickMarkAreaDrawEventArgs = new CustomTickMarkAreaDrawEventArgs();
    mp_oTickMarkTextDrawEventArgs = new TickMarkTextDrawEventArgs();
	mp_oPagePrintEventArgs = new PagePrintEventArgs();

	mp_oConfiguration = new clsConfiguration(this);
	clsG = new clsGraphics(this);
	mp_oMathLib = new clsMath(this);
	mp_oStyles = new clsStyles(this);
	mp_sStyleIndex = "DS_CONTROL";
	mp_oStyle = GetStyles()->FItem("DS_CONTROL");
	mp_oVerticalScrollBar = new clsVerticalScrollBar(this);
	mp_oHorizontalScrollBar = new clsHorizontalScrollBar(this);
	mp_oTimeLineScrollBar = new clsTimeLineScrollBar(this);
	mp_oRows = new clsRows(this);
	mp_oTasks = new clsTasks(this);
	mp_oColumns = new clsColumns(this);
	mp_oLayers = new clsLayers(this);
	mp_oPercentages = new clsPercentages(this);
	mp_oTimeBlocks = new clsTimeBlocks(this);
	mp_oPredecessors = new clsPredecessors(this);
	tmpTimeBlocks = new clsTimeBlocks(this);
	if (mp_GetPrinterCount() == 0)
	{
		mp_oPrinter = NULL;
		mp_sPrinterErrorMessage = _T("No Printers Installed");
	}
	else
	{
		mp_oPrinter = new clsPrinter(this);
	}
	mp_oSplitter = new clsSplitter(this);
	mp_oViews = new clsViews(this);
	mp_oTreeview = new clsTreeview(this);
	mp_oCurrentView = mp_oViews->FItem("0");
	MouseKeyboardEvents = new clsMouseKeyboardEvents(this);
	mp_oDrawing = new clsDrawing(this);
	mp_oCulture = new clsCultureInfo();
	mp_oTierAppearance = new clsTierAppearance(this);
	mp_oTierFormat = new clsTierFormat(this);
	mp_oScrollBarSeparator = new clsScrollBarSeparator(this);
	mp_oTextBox = new clsTextBox(this);
	mp_oProgressLine = new clsProgressLine(this);
	
	mp_oImage = new clsImage();
	mp_sImageTag = _T("");


	int iXExtents = GetSystemMetrics (SM_CXVIRTUALSCREEN);
	int iYExtents = GetSystemMetrics (SM_CYVIRTUALSCREEN);


	mp_oBitmap = new Gdiplus::Bitmap(iXExtents, iYExtents, PixelFormat32bppARGB);

	mp_oBrush.CreateSolidBrush(RGB(255, 255, 255));

	
}

DWORD CActiveGanttVCCtl::mp_GetPrinterCount(void)
{
	PRINTER_INFO_2* oPrinterList;
	DWORD lCount = 0;
	DWORD lSize = 0;
	DWORD lLevel = 2;
	EnumPrinters(PRINTER_ENUM_LOCAL, NULL, lLevel, NULL, 0, &lSize, &lCount);
	if ((oPrinterList = (PRINTER_INFO_2*)malloc(lSize)) == 0) 
	{
		return 0;
	}
	else
	{
		EnumPrinters(PRINTER_ENUM_LOCAL, NULL, lLevel, (LPBYTE)oPrinterList, lSize, &lSize, &lCount);
		free(oPrinterList);
		return lCount;
	}
}

CActiveGanttVCCtl::~CActiveGanttVCCtl()
{
	

	delete mp_oConfiguration;
	mp_oConfiguration = NULL;

	delete mp_oProgressLine;
	mp_oProgressLine = NULL;

	delete mp_oTextBox;
	mp_oTextBox = NULL;

	delete mp_oScrollBarSeparator;
	mp_oScrollBarSeparator = NULL;

	delete mp_oTierFormat;
	mp_oTierFormat = NULL;

	delete mp_oTierAppearance;
	mp_oTierAppearance = NULL;

	delete mp_oCulture;
	mp_oCulture = NULL;

	delete mp_oDrawing;
	mp_oDrawing = NULL;

	delete MouseKeyboardEvents;
	MouseKeyboardEvents = NULL;

	delete mp_oTreeview;
	mp_oTreeview = NULL;

	delete mp_oViews;
	mp_oViews = NULL;

	delete mp_oSplitter;
	mp_oSplitter = NULL;

	delete mp_oPrinter;
	mp_oPrinter = NULL;

	delete tmpTimeBlocks;
	tmpTimeBlocks = NULL;

	delete mp_oPredecessors;
	mp_oPredecessors = NULL;

	delete mp_oTimeBlocks;
	mp_oTimeBlocks = NULL;

	delete mp_oPercentages;
	mp_oPercentages = NULL;

	delete mp_oLayers;
	mp_oLayers = NULL;

	delete mp_oColumns;
	mp_oColumns = NULL;

	delete mp_oTasks;
	mp_oTasks = NULL;

	delete mp_oRows;
	mp_oRows = NULL;

	delete mp_oStyles;
	mp_oStyles = NULL;

	delete mp_oMathLib;
	mp_oMathLib = NULL;

	delete clsG;
	clsG = NULL;

	delete mp_oTimeLineScrollBar;
	mp_oTimeLineScrollBar = NULL;

	delete mp_oHorizontalScrollBar;
	mp_oHorizontalScrollBar = NULL;

	delete mp_oVerticalScrollBar;
	mp_oVerticalScrollBar = NULL;

	delete mp_oPagePrintEventArgs;
	mp_oPagePrintEventArgs = NULL;

	delete mp_oCustomTickMarkAreaDrawEventArgs;
	mp_oCustomTickMarkAreaDrawEventArgs = NULL;

    delete mp_oTickMarkTextDrawEventArgs;
	mp_oTickMarkTextDrawEventArgs = NULL;

	delete mp_oPredecessorExceptionEventArgs;
	mp_oPredecessorExceptionEventArgs = NULL;

	delete mp_oTextEditEventArgs;
	mp_oTextEditEventArgs = NULL;

	delete mp_oToolTipEventArgs;
	mp_oToolTipEventArgs = NULL;

	delete mp_oObjectAddedEventArgs;
	mp_oObjectAddedEventArgs = NULL;

	delete mp_oCustomTierDrawEventArgs;
	mp_oCustomTierDrawEventArgs = NULL;

	delete mp_oMouseEventArgs;
	mp_oMouseEventArgs = NULL;

	delete mp_oMouseWheelEventArgs;
	mp_oMouseWheelEventArgs = NULL;

	delete mp_oKeyEventArgs;
	mp_oKeyEventArgs = NULL;

	delete mp_oScrollEventArgs;
	mp_oScrollEventArgs = NULL;

	delete mp_oDrawEventArgs;
	mp_oDrawEventArgs = NULL;

	delete mp_oPredecessorDrawEventArgs;
	mp_oPredecessorDrawEventArgs = NULL;

	delete mp_oObjectSelectedEventArgs;
	mp_oObjectSelectedEventArgs = NULL;

	delete mp_oObjectStateChangedEventArgs;
	mp_oObjectStateChangedEventArgs = NULL;

	delete mp_oErrorEventArgs;
	mp_oErrorEventArgs = NULL;

	delete mp_oNodeEventArgs;
	mp_oNodeEventArgs = NULL;

	delete mp_oImage;
	mp_oImage = NULL;

	delete mp_oBitmap;
	mp_oBitmap = NULL;
	
	delete mp_oGraphics;
	mp_oGraphics = NULL;

	if (Getf_UserMode() == FALSE)
	{
		if (AfxGetStaticModuleState()->m_nObjectCount != (mp_lObjectCount))
		{
			int lNewObjectCount = AfxGetStaticModuleState()->m_nObjectCount;
			int lDiff = lNewObjectCount - mp_lObjectCount;
			mp_ErrorReport(OBJECT_COUNT_ERROR, "ActiveGantt objects (" + CStr(lDiff) + " in total) you have created have not been property deinitialized. In Visual C++ MFC use the ReleaseDispatch method. In Visual Basic set the object to Nothing.", "ActiveGanttVCCtl.Destructor");
		}
	}

	//Terminate GDI+
	GDI_Plus_Controler.Deinitialize();
}

void CActiveGanttVCCtl::OnDraw(CDC* pdc, const CRect& rcBounds, const CRect& rcInvalid)
{

	Gdiplus::Graphics* oDestGraphics = NULL; // used for double buffering

	if (Getf_UserMode() == FALSE)
	{
		// Design mode - No double buffering
		mp_oGraphics = new Gdiplus::Graphics(pdc->m_hDC);
	}
	else
	{
		// User mode - use double buffering
		oDestGraphics = new Gdiplus::Graphics(pdc->m_hDC);
		mp_oGraphics = Graphics::FromImage(mp_oBitmap);
	}

	mp_PositionScrollBars();
	if (Getf_UserMode() == FALSE)
	{
		// Design mode
		mp_DrawDesignMode();
	}
	else
	{
		// User mode
		mp_Draw();
	}
	if (GetCurrentViewObject()->GetTimeLine()->GetStartDate() != mp_dtTimeLineStartBuffer || GetCurrentViewObject()->GetTimeLine()->GetEndDate() != mp_dtTimeLineEndBuffer)
	{
		mp_dtTimeLineStartBuffer = GetCurrentViewObject()->GetTimeLine()->GetStartDate();
		mp_dtTimeLineEndBuffer = GetCurrentViewObject()->GetTimeLine()->GetEndDate();
		FireTimeLineChanged();
	}

	if (Getf_UserMode() == FALSE)
	{
	}
	else
	{
		// User mode - use double buffering
		oDestGraphics->DrawImage(mp_oBitmap,0,0);
	}
	delete mp_oGraphics;
	mp_oGraphics = NULL;
	delete oDestGraphics;
	oDestGraphics = NULL;
}

void CActiveGanttVCCtl::OnSize(UINT nType, int cx, int cy) 
{
	COleControl::OnSize(nType, cx, cy);
	if (clsG->Width() <= 0 || clsG->Height() <= 0)
	{
		return;
	}


	InvalidateControl();
}

void CActiveGanttVCCtl::DoPropExchange(CPropExchange* pPX)
{
	ExchangeVersion(pPX, MAKELONG(_wVerMinor, _wVerMajor));
	COleControl::DoPropExchange(pPX);

	if (pPX->IsLoading() == TRUE)
	{
		mp_bPropertiesRead = TRUE;
	}
	else
	{
	}
}

void CActiveGanttVCCtl::OnResetState()
{
	COleControl::OnResetState();  // Resets defaults found in DoPropExchange
}

void CActiveGanttVCCtl::OnKeyDown(UINT nChar, UINT nRepCnt, UINT nFlags) 
{
	if (nChar == 17)
	{
		GetKeyEventArgs()->SetControl(TRUE);
		return;
	}
	if (nChar == 16)
	{
		GetKeyEventArgs()->SetShift(TRUE);
		return;
	}
	GetKeyEventArgs()->Clear();
	GetKeyEventArgs()->SetKeyCode(nChar);
	ActiveGanttVCCtl_KeyDown();
	COleControl::OnKeyDown(nChar, nRepCnt, nFlags);
}

void CActiveGanttVCCtl::OnSysKeyDown(UINT nChar, UINT nRepCnt, UINT nFlags) 
{
	if (nChar == 18)
	{
		GetKeyEventArgs()->SetAlt(TRUE);
		return;
	}
	GetKeyEventArgs()->Clear();
	GetKeyEventArgs()->SetKeyCode(nChar);
	ActiveGanttVCCtl_KeyDown();
	COleControl::OnSysKeyDown(nChar, nRepCnt, nFlags);
}

void CActiveGanttVCCtl::ActiveGanttVCCtl_KeyDown(void)
{
	MouseKeyboardEvents->KeyDown();
}

void CActiveGanttVCCtl::OnKeyUp(UINT nChar, UINT nRepCnt, UINT nFlags) 
{
	if (nChar == 17)
	{
		GetKeyEventArgs()->SetControl(FALSE);
		return;
	}
	if (nChar == 16)
	{
		GetKeyEventArgs()->SetShift(FALSE);
		return;
	}
	GetKeyEventArgs()->Clear();
	GetKeyEventArgs()->SetKeyCode(nChar);
	ActiveGanttVCCtl_KeyUp();
	COleControl::OnKeyUp(nChar, nRepCnt, nFlags);
}

void CActiveGanttVCCtl::OnSysKeyUp(UINT nChar, UINT nRepCnt, UINT nFlags) 
{

	if (nChar == 18)
	{
		GetKeyEventArgs()->SetAlt(FALSE);
		return;
	}
	GetKeyEventArgs()->Clear();
	GetKeyEventArgs()->SetKeyCode(nChar);
	ActiveGanttVCCtl_KeyUp();
	COleControl::OnSysKeyUp(nChar, nRepCnt, nFlags);
}

void CActiveGanttVCCtl::ActiveGanttVCCtl_KeyUp(void)
{
	MouseKeyboardEvents->KeyUp();
}

void CActiveGanttVCCtl::OnSysDeadChar(UINT nChar, UINT nRepCnt, UINT nFlags)
{
	GetKeyEventArgs()->Clear();
	GetKeyEventArgs()->SetCharacterCode(nChar);
	ActiveGanttVCCtl_KeyPress();
	COleControl::OnSysDeadChar(nChar, nRepCnt, nFlags);
}

void CActiveGanttVCCtl::OnDeadChar(UINT nChar, UINT nRepCnt, UINT nFlags)
{
	GetKeyEventArgs()->Clear();
	GetKeyEventArgs()->SetCharacterCode(nChar);
	ActiveGanttVCCtl_KeyPress();
	COleControl::OnDeadChar(nChar, nRepCnt, nFlags);
}

void CActiveGanttVCCtl::OnChar(UINT nChar, UINT nRepCnt, UINT nFlags)
{
	GetKeyEventArgs()->Clear();
	GetKeyEventArgs()->SetCharacterCode(nChar);
	ActiveGanttVCCtl_KeyPress();
	COleControl::OnChar(nChar, nRepCnt, nFlags);
}

void CActiveGanttVCCtl::ActiveGanttVCCtl_KeyPress(void)
{
	MouseKeyboardEvents->KeyPress();
}

void CActiveGanttVCCtl::OnLButtonDblClk(UINT nFlags, CPoint point) 
{
	ActiveGanttVCCtl_DoubleClick();
	COleControl::OnLButtonDblClk(nFlags, point);
}

void CActiveGanttVCCtl::OnLButtonDown(UINT nFlags, CPoint point) 
{
	ActiveGanttVCCtl_MouseDown(point.x, point.y, BTN_LEFT);
	COleControl::OnLButtonDown(nFlags, point);
}

void CActiveGanttVCCtl::OnLButtonUp(UINT nFlags, CPoint point) 
{
	ActiveGanttVCCtl_MouseUp(point.x, point.y);
	ActiveGanttVCCtl_Click();
	COleControl::OnLButtonUp(nFlags, point);
}

void CActiveGanttVCCtl::OnMButtonDblClk(UINT nFlags, CPoint point) 
{
	ActiveGanttVCCtl_DoubleClick();
	COleControl::OnMButtonDblClk(nFlags, point);
}

void CActiveGanttVCCtl::OnMButtonDown(UINT nFlags, CPoint point) 
{
	ActiveGanttVCCtl_MouseDown(point.x, point.y, BTN_MIDDLE);
	COleControl::OnMButtonDown(nFlags, point);
}

void CActiveGanttVCCtl::OnMButtonUp(UINT nFlags, CPoint point) 
{
	ActiveGanttVCCtl_MouseUp(point.x, point.y);
	ActiveGanttVCCtl_Click();
	COleControl::OnMButtonUp(nFlags, point);
}

void CActiveGanttVCCtl::OnRButtonDblClk(UINT nFlags, CPoint point) 
{
	ActiveGanttVCCtl_DoubleClick();
	COleControl::OnRButtonDblClk(nFlags, point);
}

void CActiveGanttVCCtl::OnRButtonDown(UINT nFlags, CPoint point) 
{
	ActiveGanttVCCtl_MouseDown(point.x, point.y, BTN_RIGHT);
	COleControl::OnRButtonDown(nFlags, point);
}

void CActiveGanttVCCtl::OnRButtonUp(UINT nFlags, CPoint point) 
{
	ActiveGanttVCCtl_MouseUp(point.x, point.y);
	ActiveGanttVCCtl_Click();
	COleControl::OnRButtonUp(nFlags, point);
}

BOOL CActiveGanttVCCtl::OnMouseWheel(UINT nFlags, short zDelta, CPoint pt) 
{
	GetMouseWheelEventArgs()->Clear();
	GetMouseWheelEventArgs()->SetDelta(zDelta);
	GetMouseWheelEventArgs()->SetX(pt.x);
	GetMouseWheelEventArgs()->SetY(pt.y);
	if (nFlags == MK_CONTROL) 
	{
		GetMouseWheelEventArgs()->SetButton(BTN_XBUTTON1);
	}
	else if (nFlags == MK_LBUTTON)
	{
		GetMouseWheelEventArgs()->SetButton(BTN_LEFT);
	}
	else if (nFlags == MK_MBUTTON)
	{
		GetMouseWheelEventArgs()->SetButton(BTN_MIDDLE);
	}
	else if (nFlags == MK_RBUTTON)
	{
		GetMouseWheelEventArgs()->SetButton(BTN_RIGHT);
	}
	else if (nFlags == MK_SHIFT)
	{
		GetMouseWheelEventArgs()->SetButton(BTN_XBUTTON2);
	}
	FireControlMouseWheel();

	return COleControl::OnMouseWheel(nFlags, zDelta, pt);
}

void CActiveGanttVCCtl::OnMouseMove(UINT nFlags, CPoint point) 
{
	ActiveGanttVCCtl_MouseMove(point.x, point.y);
	COleControl::OnMouseMove(nFlags, point);
}

void CActiveGanttVCCtl::ActiveGanttVCCtl_DoubleClick(void)
{
	MouseKeyboardEvents->OnMouseDblClick();
}

void CActiveGanttVCCtl::ActiveGanttVCCtl_Click(void)
{
	MouseKeyboardEvents->OnMouseClick();
}

void CActiveGanttVCCtl::ActiveGanttVCCtl_MouseDown(LONG X, LONG Y, LONG Button)
{
	MouseKeyboardEvents->OnMouseDown(X, Y, Button);
}

void CActiveGanttVCCtl::ActiveGanttVCCtl_MouseMove(LONG X, LONG Y)
{
	MouseKeyboardEvents->OnMouseMoveGeneral(X, Y);
}

void CActiveGanttVCCtl::ActiveGanttVCCtl_MouseUp(LONG X, LONG Y)
{
	MouseKeyboardEvents->OnMouseUp(X, Y);
}

BOOL CActiveGanttVCCtl::OnSetCursor(CWnd* pWnd, UINT nHitTest, UINT message) 
{
	//return COleControl::OnSetCursor(pWnd, nHitTest, message);
	return TRUE;
}

void CActiveGanttVCCtl::mp_CHKPXPStart(Color clrBackground)
{
	//clsG->ClipRegion(0, 0, clsG->Width(), clsG->Height(), FALSE);
	clsG->mp_DrawLine(0, 0, clsG->Width(), clsG->Height(), LT_FILLED, clrBackground, LDS_SOLID);
}

void CActiveGanttVCCtl::mp_CHKPXPScrollButtons(void)
{
	mp_CHKPXPStart(Color::Red);
	clsG->mp_DrawScrollButton(20, 20, 17, 17, SBC_UP, BS_NORMAL);
	clsG->mp_DrawScrollButton(40, 20, 17, 17, SBC_UP, BS_PUSHED);
	clsG->mp_DrawScrollButton(60, 20, 17, 17, SBC_UP, BS_INACTIVE);

	clsG->mp_DrawScrollButton(20, 40, 17, 17, SBC_DOWN, BS_NORMAL);
	clsG->mp_DrawScrollButton(40, 40, 17, 17, SBC_DOWN, BS_PUSHED);
	clsG->mp_DrawScrollButton(60, 40, 17, 17, SBC_DOWN, BS_INACTIVE);

	clsG->mp_DrawScrollButton(20, 60, 17, 17, SBC_LEFT, BS_NORMAL);
	clsG->mp_DrawScrollButton(40, 60, 17, 17, SBC_LEFT, BS_PUSHED);
	clsG->mp_DrawScrollButton(60, 60, 17, 17, SBC_LEFT, BS_INACTIVE);

	clsG->mp_DrawScrollButton(20, 80, 17, 17, SBC_RIGHT, BS_NORMAL);
	clsG->mp_DrawScrollButton(40, 80, 17, 17, SBC_RIGHT, BS_PUSHED);
	clsG->mp_DrawScrollButton(60, 80, 17, 17, SBC_RIGHT, BS_INACTIVE);
}

void CActiveGanttVCCtl::mp_CHKPXPLines(void)
{
	mp_CHKPXPStart(Color::White);
	clsG->mp_DrawLine(20, 10, 60, 10, LT_NORMAL, Color::Black, LDS_SOLID);
	clsG->mp_DrawLine(30, 20, 30, 70, LT_NORMAL, Color::Black, LDS_SOLID);
	clsG->mp_DrawLine(40, 40, 60, 60, LT_BORDER, Color::Black, LDS_SOLID);
	clsG->mp_DrawLine(70, 70, 90, 90, LT_FILLED, Color::Black, LDS_SOLID);
}

void CActiveGanttVCCtl::mp_CHKPXPButtons(void)
{
	mp_CHKPXPStart(Color::Red);
	Rect oRect(20, 20, 100, 50);
	clsG->mp_DrawButton(oRect, BS_NORMAL);
	clsG->mp_DrawItem(50, 50, 100, 100, "", "hello", FALSE, NULL, 100, 100, mp_oStyles->FItem("DS_TICKMARKAREA"));
}

void CActiveGanttVCCtl::mp_CHKPXPArrows(void)
{
	mp_CHKPXPStart(Color::White);
	clsG->mp_DrawArrow(100, 120, AWD_RIGHT, 20, Color::Black);
	clsG->mp_DrawArrow(50, 120, AWD_LEFT, 20, Color::Black);
	clsG->mp_DrawArrow(100, 150, AWD_UP, 20, Color::Black);
	clsG->mp_DrawArrow(50, 170, AWD_DOWN, 20, Color::Black);
}

void CActiveGanttVCCtl::mp_CHKPXPFigures(void)
{
	mp_CHKPXPStart(Color::White);
	clsG->mp_DrawFigure(200, 20, 20, 20, FT_ARROWDOWN, Color::Blue, Color::Black, LDS_SOLID);
	clsG->mp_DrawFigure(230, 20, 20, 20, FT_ARROWUP, Color::Blue, Color::Black, LDS_SOLID);

	clsG->mp_DrawFigure(200, 60, 20, 20, FT_CIRCLEARROWDOWN, Color::Blue, Color::Black, LDS_SOLID);
	clsG->mp_DrawFigure(230, 60, 20, 20, FT_CIRCLEARROWUP, Color::Blue, Color::Black, LDS_SOLID);

	clsG->mp_DrawFigure(200, 100, 20, 20, FT_CIRCLETRIANGLEDOWN, Color::Blue, Color::Black, LDS_SOLID);
	clsG->mp_DrawFigure(230, 100, 20, 20, FT_CIRCLETRIANGLEUP, Color::Blue, Color::Black, LDS_SOLID);

	clsG->mp_DrawFigure(200, 140, 20, 20, FT_PROJECTDOWN, Color::Blue, Color::Black, LDS_SOLID);
	clsG->mp_DrawFigure(230, 140, 20, 20, FT_PROJECTUP, Color::Blue, Color::Black, LDS_SOLID);

	clsG->mp_DrawFigure(200, 180, 20, 20, FT_SMALLPROJECTDOWN, Color::Blue, Color::Black, LDS_SOLID);
	clsG->mp_DrawFigure(230, 180, 20, 20, FT_SMALLPROJECTUP, Color::Blue, Color::Black, LDS_SOLID);

	clsG->mp_DrawFigure(200, 220, 20, 20, FT_TRIANGLEDOWN, Color::Blue, Color::Black, LDS_SOLID);
	clsG->mp_DrawFigure(230, 220, 20, 20, FT_TRIANGLEUP, Color::Blue, Color::Black, LDS_SOLID);

	clsG->mp_DrawFigure(200, 260, 20, 20, FT_TRIANGLELEFT, Color::Blue, Color::Black, LDS_SOLID);
	clsG->mp_DrawFigure(230, 260, 20, 20, FT_TRIANGLERIGHT, Color::Blue, Color::Black, LDS_SOLID);

	clsG->mp_DrawFigure(200, 300, 20, 20, FT_SQUARE, Color::Blue, Color::Black, LDS_SOLID);
	clsG->mp_DrawFigure(230, 300, 20, 20, FT_CIRCLE, Color::Blue, Color::Black, LDS_SOLID);

	clsG->mp_DrawFigure(200, 340, 20, 20, FT_DIAMOND, Color::Blue, Color::Black, LDS_SOLID);
	clsG->mp_DrawFigure(230, 340, 20, 20, FT_RECTANGLE, Color::Blue, Color::Black, LDS_SOLID);

	clsG->mp_DrawFigure(200, 380, 20, 20, FT_NONE, Color::Blue, Color::Black, LDS_SOLID);
	clsG->mp_DrawFigure(230, 380, 20, 20, FT_NONE, Color::Blue, Color::Black, LDS_SOLID);
}

void CActiveGanttVCCtl::mp_CHKPXPGradients(void)
{
	mp_CHKPXPStart(Color::White);
	clsG->mp_GradientFill(20, 250, 120, 320, Color::Black, Color::Blue, GDT_HORIZONTAL);
	clsG->mp_GradientFill(20, 340, 120, 400, Color::Black, Color::Blue, GDT_VERTICAL);
}

void CActiveGanttVCCtl::mp_CHKPXPText(void)
{
	mp_CHKPXPStart(Color::White);
	StringFormat oFlags;
	oFlags.SetAlignment(StringAlignmentFar);
	oFlags.SetLineAlignment(StringAlignmentFar);
	clsG->mp_DrawTextEx(200, 100, 400, 300, "M Hello World", &oFlags, Color::Black, ::new Font(::new FontFamily(L"Arial"), 10), TRUE);
}

void CActiveGanttVCCtl::mp_CHKPXPHatch(void)
{
	mp_CHKPXPStart(Color::White);
	LONG i;
	LONG j;
	j = 0;
	for (i = 0; i <= 9; i++)
	{
		clsG->mp_HatchFill(0, j * 30, 100, (j * 30) + 20, Color::Black, Color::White, (GRE_HATCHSTYLE)i);
		j = j + 1;
	}
	j = 0;
	for (i = 10; i <= 19; i++)
	{
		clsG->mp_HatchFill(120, j * 30, 220, (j * 30) + 20, Color::Black, Color::White, (GRE_HATCHSTYLE)i);
		j = j + 1;
	}
	j = 0;
	for (i = 20; i <= 29; i++)
	{
		clsG->mp_HatchFill(240, j * 30, 340, (j * 30) + 20, Color::Black, Color::White, (GRE_HATCHSTYLE)i);
		j = j + 1;
	}
	j = 0;
	for (i = 30; i <= 39; i++)
	{
		clsG->mp_HatchFill(360, j * 30, 460, (j * 30) + 20, Color::Black, Color::White, (GRE_HATCHSTYLE)i);
		j = j + 1;
	}
	j = 0;
	for (i = 40; i <= 49; i++)
	{
		clsG->mp_HatchFill(480, j * 30, 580, (j * 30) + 20, Color::Black, Color::White, (GRE_HATCHSTYLE)i);
		j = j + 1;
	}
	j = 0;
	for (i = 50; i <= 52; i++)
	{
		clsG->mp_HatchFill(600, j * 30, 700, (j * 30) + 20, Color::Black, Color::White, (GRE_HATCHSTYLE)i);
		j = j + 1;
	}
}

HWND CActiveGanttVCCtl::f_Hwnd()
{
	return this->m_hWnd;
}

Gdiplus::Color CActiveGanttVCCtl::FromOLEColor(LONG clrOLEColor)
{
	Gdiplus::Color clrReturn;
	clrReturn.SetFromCOLORREF(TranslateColor(clrOLEColor));
	return clrReturn;
}

int CActiveGanttVCCtl::OnCreate(LPCREATESTRUCT lpCreateStruct) 
{
	if (COleControl::OnCreate(lpCreateStruct) == -1)
	{
		return -1;
	}
	return 0;
}

void CActiveGanttVCCtl::SetTimerEx(UINT_PTR nIDEvent, UINT nElapse)
{
	mp_oTimer = SetTimer(nIDEvent, nElapse, NULL);
}

void CActiveGanttVCCtl::OnTimer(UINT nIDEvent)
{
	switch(nIDEvent - 1000)
	{
	case SCR_VERTICAL:
		{
			GetVerticalScrollBar()->GetScrollBar()->mp_oTimer_Tick();
			break;
		}
	case SCR_HORIZONTAL2:
		{
			GetCurrentViewObject()->GetTimeLine()->GetTimeLineScrollBar()->GetScrollBar()->mp_oTimer_Tick();
			break;
		}
	case SCR_HORIZONTAL1:
		{
			GetHorizontalScrollBar()->GetScrollBar()->mp_oTimer_Tick();
			break;
		}
	}
}

void CActiveGanttVCCtl::KillTimerEx()
{
	KillTimer(mp_oTimer);
}



LONG CActiveGanttVCCtl::GetErrorReports(void)
{
	return mp_yErrorReports;
}

void CActiveGanttVCCtl::SetErrorReports(LONG Value)
{
	mp_yErrorReports = Value;
}

CString CActiveGanttVCCtl::GetCurrentLayer(void)
{
	return mp_sCurrentLayer;
}

void CActiveGanttVCCtl::SetCurrentLayer(CString Value)
{
	mp_sCurrentLayer = Value;
}

CString CActiveGanttVCCtl::GetCurrentView(void)
{
	return mp_sCurrentView;
}

void CActiveGanttVCCtl::SetCurrentView(CString Value)
{
	if (Value == "")
	{
		Value = "0";
	}
	mp_oCurrentView = GetViews()->FItem(Value);
	mp_sCurrentView = Value;
}

clsView* CActiveGanttVCCtl::GetCurrentViewObject(void)
{
	return mp_oCurrentView;
}

clsToolTip* CActiveGanttVCCtl::GetToolTip(void)
{
	return MouseKeyboardEvents->mp_oToolTip;
}

LONG CActiveGanttVCCtl::GetScrollBarBehaviour(void)
{
	return mp_yScrollBarBehaviour;
}

void CActiveGanttVCCtl::SetScrollBarBehaviour(LONG Value)
{
	mp_yScrollBarBehaviour = Value;
}

LONG CActiveGanttVCCtl::GetTimeBlockBehaviour(void)
{
	return mp_yTimeBlockBehaviour;
}

void CActiveGanttVCCtl::SetTimeBlockBehaviour(LONG Value)
{
	mp_yTimeBlockBehaviour = Value;
}

LONG CActiveGanttVCCtl::GetSelectedTaskIndex(void)
{
	return mp_lSelectedTaskIndex;
}

void CActiveGanttVCCtl::SetSelectedTaskIndex(LONG Value)
{
	mp_lSelectedTaskIndex = Value;
}

LONG CActiveGanttVCCtl::GetSelectedColumnIndex(void)
{
	return mp_lSelectedColumnIndex;
}

void CActiveGanttVCCtl::SetSelectedColumnIndex(LONG Value)
{
	mp_lSelectedColumnIndex = Value;
}

LONG CActiveGanttVCCtl::GetSelectedRowIndex(void)
{
	return mp_lSelectedRowIndex;
}

void CActiveGanttVCCtl::SetSelectedRowIndex(LONG Value)
{
	mp_lSelectedRowIndex = Value;
}

LONG CActiveGanttVCCtl::GetSelectedCellIndex(void)
{
	return mp_lSelectedCellIndex;
}

void CActiveGanttVCCtl::SetSelectedCellIndex(LONG Value)
{
	mp_lSelectedCellIndex = Value;
}

LONG CActiveGanttVCCtl::GetSelectedPercentageIndex(void)
{
	return mp_lSelectedPercentageIndex;
}

void CActiveGanttVCCtl::SetSelectedPercentageIndex(LONG Value)
{
	mp_lSelectedPercentageIndex = Value;
}

LONG CActiveGanttVCCtl::GetTreeviewColumnIndex(void)
{
    if (GetColumns()->GetCount() == 0)
    {
        return 0;
    }
    else if (mp_lTreeviewColumnIndex > GetColumns()->GetCount())
    {
        return 0;
    }
    else if (mp_lTreeviewColumnIndex < 0)
    {
        return 0;
    }
    else
    {
        return mp_lTreeviewColumnIndex;
    }
}

void CActiveGanttVCCtl::SetTreeviewColumnIndex(LONG Value)
{
    if (Value <= 0)
    {
        Value = 0;
    }
    else if (Value > GetColumns()->GetCount())
    {
        Value = GetColumns()->GetCount();
    }
    mp_lTreeviewColumnIndex = Value;
}

LONG CActiveGanttVCCtl::GetSelectedPredecessorIndex(void)
{
	return mp_lSelectedPredecessorIndex;
}

void CActiveGanttVCCtl::SetSelectedPredecessorIndex(LONG Value)
{
	mp_lSelectedPredecessorIndex = Value;
}

clsCultureInfo* CActiveGanttVCCtl::GetCulture(void)
{
	return mp_oCulture;
}

CString CActiveGanttVCCtl::GetModuleCompletePath(void)
{
	return mp_sModuleCompletePath;
}

CString CActiveGanttVCCtl::GetVersion(void)
{
	return sOCXVERSION.AllocSysString();
}

BOOL CActiveGanttVCCtl::GetAllowSplitterMove(void)
{
	return mp_bAllowSplitterMove;
}

void CActiveGanttVCCtl::SetAllowSplitterMove(BOOL Value)
{
	mp_bAllowSplitterMove = Value;
}

BOOL CActiveGanttVCCtl::GetAllowPredecessorAdd(void)
{
	return mp_bAllowPredecessorAdd;
}

void CActiveGanttVCCtl::SetAllowPredecessorAdd(BOOL Value)
{
	mp_bAllowPredecessorAdd = Value;
}

BOOL CActiveGanttVCCtl::GetAllowAdd(void)
{
	return mp_bAllowAdd;
}

void CActiveGanttVCCtl::SetAllowAdd(BOOL Value)
{
	mp_bAllowAdd = Value;
}

BOOL CActiveGanttVCCtl::GetAllowEdit(void)
{
	return mp_bAllowEdit;
}

void CActiveGanttVCCtl::SetAllowEdit(BOOL Value)
{
	mp_bAllowEdit = Value;
}

BOOL CActiveGanttVCCtl::GetAllowColumnSize(void)
{
	return mp_bAllowColumnSize;
}

void CActiveGanttVCCtl::SetAllowColumnSize(BOOL Value)
{
	mp_bAllowColumnSize = Value;
}

BOOL CActiveGanttVCCtl::GetAllowColumnMove(void)
{
	return mp_bAllowColumnMove;
}

void CActiveGanttVCCtl::SetAllowColumnMove(BOOL Value)
{
	mp_bAllowColumnMove = Value;
}

BOOL CActiveGanttVCCtl::GetAllowRowSize(void)
{
	return mp_bAllowRowSize;
}

void CActiveGanttVCCtl::SetAllowRowSize(BOOL Value)
{
	mp_bAllowRowSize = Value;
}

BOOL CActiveGanttVCCtl::GetAllowRowMove(void)
{
	return mp_bAllowRowMove;
}

void CActiveGanttVCCtl::SetAllowRowMove(BOOL Value)
{
	mp_bAllowRowMove = Value;
}

BOOL CActiveGanttVCCtl::GetAllowTimeLineScroll(void)
{
	return mp_bAllowTimeLineScroll;
}

void CActiveGanttVCCtl::SetAllowTimeLineScroll(BOOL Value)
{
	mp_bAllowTimeLineScroll = Value;
}

BOOL CActiveGanttVCCtl::GetEnforcePredecessors(void)
{
	return mp_bEnforcePredecessors;
}

void CActiveGanttVCCtl::SetEnforcePredecessors(BOOL Value)
{
	mp_bEnforcePredecessors= Value;
}

LONG CActiveGanttVCCtl::GetAddMode(void)
{
	return mp_yAddMode;
}

void CActiveGanttVCCtl::SetAddMode(LONG Value)
{
	mp_yAddMode = Value;
}

LONG CActiveGanttVCCtl::GetPredecessorMode(void)
{
	return mp_yPredecessorMode;
}

void CActiveGanttVCCtl::SetPredecessorMode(LONG Value)
{
	mp_yPredecessorMode = Value;
}

void CActiveGanttVCCtl::ApplyTemplate(LONG ControlTemplate, LONG ObjectTemplate)
{
    mp_yControlTemplate = ControlTemplate;
    mp_yObjectTemplate = ObjectTemplate;
    g_ApplyTemplate(this, ControlTemplate, ObjectTemplate);
}

void CActiveGanttVCCtl::ApplyTemplateSolid(ControlTemplateSolid &ControlTemplate, LONG ObjectTemplate)
{
    mp_yControlTemplate = STC_NONE;
    mp_yObjectTemplate = ObjectTemplate;
    g_ApplyTemplate_Solid(this, ControlTemplate, ObjectTemplate);
}

void CActiveGanttVCCtl::ApplyTemplateGradient(ControlTemplateGradient &ControlTemplate, LONG ObjectTemplate)
{
    mp_yControlTemplate = STC_NONE;
    mp_yObjectTemplate = ObjectTemplate;
    g_ApplyTemplate_Gradient(this, ControlTemplate, ObjectTemplate);
}

LONG CActiveGanttVCCtl::GetControlTemplate(void)
{
	return mp_yControlTemplate;
}

LONG CActiveGanttVCCtl::GetObjectTemplate(void)
{
	return mp_yObjectTemplate;
}

LONG CActiveGanttVCCtl::GetLayerEnableObjects(void)
{
	return mp_yLayerEnableObjects;
}

void CActiveGanttVCCtl::SetLayerEnableObjects(LONG Value)
{
	mp_yLayerEnableObjects = Value;
}

BOOL CActiveGanttVCCtl::GetDoubleBuffering(void)
{
	return mp_bDoubleBuffering;
}

void CActiveGanttVCCtl::SetDoubleBuffering(BOOL Value)
{
	mp_bDoubleBuffering = Value;
}

LONG CActiveGanttVCCtl::GetMinRowHeight(void)
{
	return mp_lMinRowHeight;
}

void CActiveGanttVCCtl::SetMinRowHeight(LONG Value)
{
	mp_lMinRowHeight = Value;
}

LONG CActiveGanttVCCtl::GetMinColumnWidth(void)
{
	return mp_lMinColumnWidth;
}

void CActiveGanttVCCtl::SetMinColumnWidth(LONG Value)
{
	mp_lMinColumnWidth = Value;
}

CString CActiveGanttVCCtl::GetControlTag(void)
{
	return mp_sControlTag;
}

void CActiveGanttVCCtl::SetControlTag(CString Value)
{
	mp_sControlTag = Value;
}

LONG CActiveGanttVCCtl::GetTierAppearanceScope(void)
{
	return mp_yTierAppearanceScope;
}

void CActiveGanttVCCtl::SetTierAppearanceScope(LONG Value)
{
	mp_yTierAppearanceScope = Value;
}

LONG CActiveGanttVCCtl::GetTierFormatScope(void)
{
	return mp_yTierFormatScope;
}

void CActiveGanttVCCtl::SetTierFormatScope(LONG Value)
{
	mp_yTierFormatScope = Value;
}

LONG CActiveGanttVCCtl::GetTimeLineScrollBarScope(void)
{
	return mp_yTimeLineScrollBarScope;
}

void CActiveGanttVCCtl::SetTimeLineScrollBarScope(LONG Value)
{
	mp_yTimeLineScrollBarScope = Value;
}

LONG CActiveGanttVCCtl::GetProgressLineScope(void)
{
	return mp_yProgressLineScope;
}

void CActiveGanttVCCtl::SetProgressLineScope(LONG Value)
{
	mp_yProgressLineScope = Value;
}

CString CActiveGanttVCCtl::GetStyleIndex(void)
{
	if (mp_sStyleIndex == "DS_CONTROL") 
	{
		return "";
	}
	else 
	{
		return mp_sStyleIndex;
	}
}

void CActiveGanttVCCtl::SetStyleIndex(CString Value)
{
	Value = Value.Trim();
	if (Value.GetLength() == 0) {Value = "DS_CONTROL";} 
	mp_sStyleIndex = Value;
	mp_oStyle = GetStyles()->FItem(Value);
}

clsStyle* CActiveGanttVCCtl::GetStyle(void)
{
	return mp_oStyle;
}

clsImage* CActiveGanttVCCtl::GetImage(void)
{
	return mp_oImage;
}

CString CActiveGanttVCCtl::GetImageTag(void)
{
	return mp_sImageTag;
}

void CActiveGanttVCCtl::SetImageTag(CString Value)
{
	mp_sImageTag = Value;
}

clsRows* CActiveGanttVCCtl::GetRows(void)
{
	return mp_oRows;
}

clsTasks* CActiveGanttVCCtl::GetTasks(void)
{
	return mp_oTasks;
}

clsColumns* CActiveGanttVCCtl::GetColumns(void)
{
	return mp_oColumns;
}

clsStyles* CActiveGanttVCCtl::GetStyles(void)
{
	return mp_oStyles;
}

clsLayers* CActiveGanttVCCtl::GetLayers(void)
{
	return mp_oLayers;
}

clsPercentages* CActiveGanttVCCtl::GetPercentages(void)
{
	return mp_oPercentages;
}

clsTimeBlocks* CActiveGanttVCCtl::GetTimeBlocks(void)
{
	return mp_oTimeBlocks;
}

clsPredecessors* CActiveGanttVCCtl::GetPredecessors(void)
{
	return mp_oPredecessors;
}

clsViews* CActiveGanttVCCtl::GetViews(void)
{
	return mp_oViews;
}

clsSplitter* CActiveGanttVCCtl::GetSplitter(void)
{
	return mp_oSplitter;
}

clsPrinter* CActiveGanttVCCtl::GetPrinter(void)
{
	return mp_oPrinter;
}

clsTreeview* CActiveGanttVCCtl::GetTreeview(void)
{
	return mp_oTreeview;
}

clsDrawing* CActiveGanttVCCtl::GetDrawing(void)
{
	return mp_oDrawing;
}

clsMath* CActiveGanttVCCtl::GetMathLib(void)
{
	return mp_oMathLib;
}

clsVerticalScrollBar* CActiveGanttVCCtl::GetVerticalScrollBar(void)
{
	return mp_oVerticalScrollBar;
}

clsHorizontalScrollBar* CActiveGanttVCCtl::GetHorizontalScrollBar(void)
{
	return mp_oHorizontalScrollBar;
}

clsTierAppearance* CActiveGanttVCCtl::GetTierAppearance(void)
{
	return mp_oTierAppearance;
}

clsTierFormat* CActiveGanttVCCtl::GetTierFormat(void)
{
	return mp_oTierFormat;
}

clsScrollBarSeparator* CActiveGanttVCCtl::GetScrollBarSeparator(void)
{
	return mp_oScrollBarSeparator;
}

clsTimeLineScrollBar* CActiveGanttVCCtl::GetTimeLineScrollBar(void)
{
	return mp_oTimeLineScrollBar;
}

clsConfiguration* CActiveGanttVCCtl::GetConfiguration(void)
{
	return mp_oConfiguration;
}

TextEditEventArgs* CActiveGanttVCCtl::GetTextEditEventArgs(void)
{
	return mp_oTextEditEventArgs;
}

ToolTipEventArgs* CActiveGanttVCCtl::GetToolTipEventArgs(void)
{
	return mp_oToolTipEventArgs;
}

ObjectAddedEventArgs* CActiveGanttVCCtl::GetObjectAddedEventArgs(void)
{
	return mp_oObjectAddedEventArgs;
}

CustomTierDrawEventArgs* CActiveGanttVCCtl::GetCustomTierDrawEventArgs(void)
{
	return mp_oCustomTierDrawEventArgs;
}

MouseEventArgs* CActiveGanttVCCtl::GetMouseEventArgs(void)
{
	return mp_oMouseEventArgs;
}

KeyEventArgs* CActiveGanttVCCtl::GetKeyEventArgs(void)
{
	return mp_oKeyEventArgs;
}

ScrollEventArgs* CActiveGanttVCCtl::GetScrollEventArgs(void)
{
	return mp_oScrollEventArgs;
}

DrawEventArgs* CActiveGanttVCCtl::GetDrawEventArgs(void)
{
	return mp_oDrawEventArgs;
}

PredecessorDrawEventArgs* CActiveGanttVCCtl::GetPredecessorDrawEventArgs(void)
{
	return mp_oPredecessorDrawEventArgs;
}

PredecessorExceptionEventArgs* CActiveGanttVCCtl::GetPredecessorExceptionEventArgs(void)
{
	return mp_oPredecessorExceptionEventArgs;
}

CustomTickMarkAreaDrawEventArgs* CActiveGanttVCCtl::GetCustomTickMarkAreaDrawEventArgs(void)
{
	return mp_oCustomTickMarkAreaDrawEventArgs;
}

TickMarkTextDrawEventArgs* CActiveGanttVCCtl::GetTickMarkTextDrawEventArgs(void)
{
	return mp_oTickMarkTextDrawEventArgs;
}

PagePrintEventArgs* CActiveGanttVCCtl::GetPagePrintEventArgs(void)
{
	return mp_oPagePrintEventArgs;
}

ObjectSelectedEventArgs* CActiveGanttVCCtl::GetObjectSelectedEventArgs(void)
{
	return mp_oObjectSelectedEventArgs;
}

ObjectStateChangedEventArgs* CActiveGanttVCCtl::GetObjectStateChangedEventArgs(void)
{
	return mp_oObjectStateChangedEventArgs;
}

ErrorEventArgs* CActiveGanttVCCtl::GetErrorEventArgs(void)
{
	return mp_oErrorEventArgs;
}

NodeEventArgs* CActiveGanttVCCtl::GetNodeEventArgs(void)
{
	return mp_oNodeEventArgs;
}

MouseWheelEventArgs* CActiveGanttVCCtl::GetMouseWheelEventArgs(void)
{
	return mp_oMouseWheelEventArgs;
}

BOOL CActiveGanttVCCtl::Getf_UserMode(void)
{
	if (AmbientUserMode() == FALSE)
	{
		return FALSE;
	}
	else
	{
		return TRUE;
	}
}

LONG CActiveGanttVCCtl::Getmt_BorderThickness(void)
{
    switch (mp_oStyle->GetAppearance())
    {
        case SA_RAISED:
            return 2;
        case SA_SUNKEN:
            return 2;
        case SA_FLAT:
            if (mp_oStyle->GetBorderStyle() == SBR_NONE)
            {
                return 0;
            }
            else
            {
                return mp_oStyle->GetBorderWidth();
            }
        case SA_GRAPHICAL:
            return 0;
    }
    return 0;
}

LONG CActiveGanttVCCtl::Getmt_TableBottom(void)
{
	if ( GetHorizontalScrollBar()->GetState() == SS_SHOWN )
	{
		return clsG->Height() - Getmt_BorderThickness() - 1 - GetHorizontalScrollBar()->GetHeight();
	}
	else
	{
		return clsG->Height() - Getmt_BorderThickness() - 1;
	}
}

LONG CActiveGanttVCCtl::Getmt_TopMargin(void)
{
	return Getmt_BorderThickness();
}

LONG CActiveGanttVCCtl::Getmt_LeftMargin(void)
{
	return Getmt_BorderThickness();
}

LONG CActiveGanttVCCtl::Getmt_RightMargin(void)
{
	if ( GetVerticalScrollBar()->GetState() == SS_SHOWN )
	{
		return clsG->Width() - Getmt_BorderThickness() - 1 - GetVerticalScrollBar()->GetWidth();
	}
	else
	{
		return clsG->Width() - Getmt_BorderThickness() - 1;
	}
}

LONG CActiveGanttVCCtl::Getmt_BottomMargin(void)
{
	return clsG->Height() - Getmt_BorderThickness() - 1;
}


LONG CActiveGanttVCCtl::GetAddDurationInterval(void)
{
	return mp_yAddDurationInterval;
}

void CActiveGanttVCCtl::SetAddDurationInterval(LONG Value)
{
	mp_yAddDurationInterval = Value;
}

void CActiveGanttVCCtl::VerticalScrollBar_ValueChanged(LONG Offset)
{
	InvalidateControl();
    GetScrollEventArgs()->Clear();
    GetScrollEventArgs()->SetScrollBarType(SCR_VERTICAL);
    GetScrollEventArgs()->SetOffset(Offset);
	FireControlScroll();
}

void CActiveGanttVCCtl::HorizontalScrollBar_ValueChanged(LONG Offset)
{
	InvalidateControl();
	GetScrollEventArgs()->Clear();
    GetScrollEventArgs()->SetScrollBarType(SCR_HORIZONTAL1);
    GetScrollEventArgs()->SetOffset(Offset);
	FireControlScroll();
}

void CActiveGanttVCCtl::TimeLineScrollBar_ValueChanged(LONG Offset)
{
	InvalidateControl();
	GetScrollEventArgs()->Clear();
    GetScrollEventArgs()->SetScrollBarType(SCR_HORIZONTAL2);
    GetScrollEventArgs()->SetOffset(Offset);
	FireControlScroll();
}




void CActiveGanttVCCtl::mp_PositionScrollBars(void)
{
	if ( Getf_UserMode() == FALSE || clsG->GetCustomPrinting() == TRUE)
	{
		GetVerticalScrollBar()->SetState(SS_CANTDISPLAY);
		GetHorizontalScrollBar()->SetState(SS_CANTDISPLAY);
		f_TimeLineScrollBar()->SetState(SS_CANTDISPLAY);
		return;
	}

	if ( clsG->Height() <= mp_oCurrentView->GetClientArea()->GetTop() )
	{
		GetVerticalScrollBar()->SetState(SS_CANTDISPLAY);
		GetHorizontalScrollBar()->SetState(SS_CANTDISPLAY);
		f_TimeLineScrollBar()->SetState(SS_CANTDISPLAY);
		return;
	}

	// Determine need for HorizontalScrollBar
	LONG lWidth = 0;
	lWidth = GetColumns()->GetWidth();
	if ( lWidth > GetSplitter()->GetRight() )
	{
		if (GetHorizontalScrollBar()->mf_Visible() == TRUE)
		{
			GetHorizontalScrollBar()->SetState(SS_NEEDED);
		}
		else
		{
			GetHorizontalScrollBar()->SetState(SS_NOTNEEDED);
		}
	}
	else
	{
		GetHorizontalScrollBar()->SetState(SS_NOTNEEDED);
	}
	if ( GetSplitter()->GetRight() < 5 )
	{
		GetHorizontalScrollBar()->SetState(SS_CANTDISPLAY);
	}

	// Determine need for GetCurrentViewObject()->GetTimeLine()->ScrollBar
	if ( GetSplitter()->GetRight() < clsG->Width() - (18 + Getmt_BorderThickness()) )
	{
		if ( f_TimeLineScrollBar()->GetEnabled() == TRUE )
		{
			if (f_TimeLineScrollBar()->mf_Visible() == TRUE)
			{
				f_TimeLineScrollBar()->SetState(SS_NEEDED);
			}
			else
			{
				f_TimeLineScrollBar()->SetState(SS_NOTNEEDED);
			}
		}
		else
		{
			f_TimeLineScrollBar()->SetState(SS_NOTNEEDED);
		}
	}
	else
	{
		f_TimeLineScrollBar()->SetState(SS_CANTDISPLAY);
	}

	// Determine need for VerticalScrollBar
	if ( ((GetRows()->Height() + GetCurrentViewObject()->GetClientArea()->GetTop() + GetHorizontalScrollBar()->GetHeight() + Getmt_BorderThickness()) > clsG->Height()) || (GetRows()->GetRealFirstVisibleRow() > 1) )
	{
		if (f_TimeLineScrollBar()->GetState() == SS_CANTDISPLAY)
		{
			GetVerticalScrollBar()->SetState(SS_CANTDISPLAY);
		}
		else
		{
			GetVerticalScrollBar()->SetState(SS_NEEDED);
		}
	}
	else
	{
		GetVerticalScrollBar()->SetState(SS_NOTNEEDED);
	}

	if (GetVerticalScrollBar()->mf_Visible() == FALSE)
	{
		GetVerticalScrollBar()->SetState(SS_CANTDISPLAY);
	}
	if (GetHorizontalScrollBar()->mf_Visible() == FALSE)
	{
		GetHorizontalScrollBar()->SetState(SS_CANTDISPLAY);
	}
	if (f_TimeLineScrollBar()->mf_Visible() == FALSE)
	{
		f_TimeLineScrollBar()->SetState(SS_CANTDISPLAY);
	}
	if ( GetVerticalScrollBar()->GetState() == SS_SHOWN )
	{
		GetVerticalScrollBar()->Position();
	}
	if ( GetHorizontalScrollBar()->GetState() == SS_SHOWN )
	{
		GetHorizontalScrollBar()->Position();
	}
	if ( f_TimeLineScrollBar()->GetState() == SS_SHOWN )
	{
		f_TimeLineScrollBar()->Position();
	}
}

void CActiveGanttVCCtl::WriteXML(CString url)
{
	clsXML oXML(this, "ActiveGanttCtl");
	mp_WriteXML(&oXML);
	oXML.WriteXML(url);
}

void CActiveGanttVCCtl::ReadXML(CString url)
{
	clsXML oXML(this, "ActiveGanttCtl");
	oXML.ReadXML(url);
	mp_ReadXML(&oXML);
}

void CActiveGanttVCCtl::SetXML(CString sXML)
{
	clsXML oXML(this, "ActiveGanttCtl");
	oXML.SetXML(sXML);
	mp_ReadXML(&oXML);
}

CString CActiveGanttVCCtl::GetXML(void)
{
	clsXML oXML(this, "ActiveGanttCtl");
	mp_WriteXML(&oXML);
	return oXML.GetXML().AllocSysString();
}

void CActiveGanttVCCtl::mp_WriteXML(clsXML* oXML)
{
	oXML->InitializeWriter();
	oXML->WriteProperty("Version", "AGVC");
	oXML->WriteProperty("ControlTag", mp_sControlTag);
	oXML->WriteProperty("AddMode", mp_yAddMode);
	oXML->WriteProperty("AddDurationInterval", mp_yAddDurationInterval);
	oXML->WriteProperty("AllowAdd", mp_bAllowAdd);
	oXML->WriteProperty("AllowColumnMove", mp_bAllowColumnMove);
	oXML->WriteProperty("AllowColumnSize", mp_bAllowColumnSize);
	oXML->WriteProperty("AllowEdit", mp_bAllowEdit);
	oXML->WriteProperty("AllowPredecessorAdd", mp_bAllowPredecessorAdd);
	oXML->WriteProperty("AllowRowMove", mp_bAllowRowMove);
	oXML->WriteProperty("AllowRowSize", mp_bAllowRowSize);
	oXML->WriteProperty("AllowSplitterMove", mp_bAllowSplitterMove);
	oXML->WriteProperty("AllowTimeLineScroll", mp_bAllowTimeLineScroll);
	oXML->WriteProperty("EnforcePredecessors", mp_bEnforcePredecessors);
	oXML->WriteProperty("CurrentLayer", mp_sCurrentLayer);
	oXML->WriteProperty("CurrentView", mp_sCurrentView);
	oXML->WriteProperty("DoubleBuffering", mp_bDoubleBuffering);
	oXML->WriteProperty("ErrorReports", mp_yErrorReports);
	oXML->WriteProperty("LayerEnableObjects", mp_yLayerEnableObjects);
	oXML->WriteProperty("MinColumnWidth", mp_lMinColumnWidth);
	oXML->WriteProperty("MinRowHeight", mp_lMinRowHeight);
	oXML->WriteProperty("ScrollBarBehaviour", mp_yScrollBarBehaviour);
	oXML->WriteProperty("SelectedCellIndex", mp_lSelectedCellIndex);
	oXML->WriteProperty("SelectedColumnIndex", mp_lSelectedColumnIndex);
	oXML->WriteProperty("SelectedPercentageIndex", mp_lSelectedPercentageIndex);
	oXML->WriteProperty("SelectedPredecessorIndex", mp_lSelectedPredecessorIndex);
	oXML->WriteProperty("SelectedRowIndex", mp_lSelectedRowIndex);
	oXML->WriteProperty("SelectedTaskIndex", mp_lSelectedTaskIndex);
	oXML->WriteProperty("TreeviewColumnIndex", mp_lTreeviewColumnIndex);
	oXML->WriteProperty("TimeBlockBehaviour", mp_yTimeBlockBehaviour);
	oXML->WriteProperty("TierAppearanceScope", mp_yTierAppearanceScope);
	oXML->WriteProperty("TierFormatScope", mp_yTierFormatScope);
	oXML->WriteProperty("TimeLineScrollBarScope", mp_yTimeLineScrollBarScope);
	oXML->WriteProperty("ProgressLineScope", mp_yProgressLineScope);
	oXML->WriteProperty("PredecessorMode", mp_yPredecessorMode);
	oXML->WriteProperty("StyleIndex", mp_sStyleIndex);
	oXML->WriteProperty("Image", *mp_oImage);
	oXML->WriteProperty("ImageTag", mp_sImageTag);
	oXML->WriteObject(mp_oConfiguration->GetXML());
	oXML->WriteObject(mp_oStyles->GetXML());
	oXML->WriteObject(mp_oRows->GetXML());
	oXML->WriteObject(mp_oColumns->GetXML());
	oXML->WriteObject(mp_oLayers->GetXML());
	oXML->WriteObject(mp_oTasks->GetXML());
	oXML->WriteObject(mp_oPredecessors->GetXML());
	oXML->WriteObject(mp_oTierAppearance->GetXML());////Must go before Views
	oXML->WriteObject(mp_oTierFormat->GetXML()); ////Must go before Views
	oXML->WriteObject(mp_oTimeLineScrollBar->GetXML()); ////Must go before Views
	oXML->WriteObject(mp_oProgressLine->GetXML()); ////Must go before Views
	oXML->WriteObject(mp_oViews->GetXML()); ////*********Views
	oXML->WriteObject(mp_oTimeBlocks->GetXML());
	oXML->WriteObject(mp_oTimeBlocks->CP_GetXML());
	oXML->WriteObject(mp_oPercentages->GetXML());
	oXML->WriteObject(mp_oSplitter->GetXML());
	oXML->WriteObject(mp_oTreeview->GetXML());
	oXML->WriteObject(mp_oScrollBarSeparator->GetXML());
	oXML->WriteObject(mp_oVerticalScrollBar->GetXML());
	oXML->WriteObject(mp_oHorizontalScrollBar->GetXML());
}

void CActiveGanttVCCtl::mp_ReadXML(clsXML* oXML)
{
	CString sVersion = "";
	CString sCurrentView = "";
	Clear();
	oXML->InitializeReader();
	oXML->ReadProperty("Version", sVersion);
	oXML->ReadProperty("ControlTag", mp_sControlTag);
	oXML->ReadProperty("AddMode", mp_yAddMode);
	oXML->ReadProperty("AddDurationInterval", mp_yAddDurationInterval);
	oXML->ReadProperty("AllowAdd", mp_bAllowAdd);
	oXML->ReadProperty("AllowColumnMove", mp_bAllowColumnMove);
	oXML->ReadProperty("AllowColumnSize", mp_bAllowColumnSize);
	oXML->ReadProperty("AllowEdit", mp_bAllowEdit);
	oXML->ReadProperty("AllowPredecessorAdd", mp_bAllowPredecessorAdd);
	oXML->ReadProperty("AllowRowMove", mp_bAllowRowMove);
	oXML->ReadProperty("AllowRowSize", mp_bAllowRowSize);
	oXML->ReadProperty("AllowSplitterMove", mp_bAllowSplitterMove);
	oXML->ReadProperty("AllowTimeLineScroll", mp_bAllowTimeLineScroll);
	oXML->ReadProperty("EnforcePredecessors", mp_bEnforcePredecessors);
	oXML->ReadProperty("CurrentLayer", mp_sCurrentLayer);
	oXML->ReadProperty("CurrentView", sCurrentView);
	oXML->ReadProperty("DoubleBuffering", mp_bDoubleBuffering);
	oXML->ReadProperty("ErrorReports", mp_yErrorReports);
	oXML->ReadProperty("LayerEnableObjects", mp_yLayerEnableObjects);
	oXML->ReadProperty("MinColumnWidth", mp_lMinColumnWidth);
	oXML->ReadProperty("MinRowHeight", mp_lMinRowHeight);
	oXML->ReadProperty("ScrollBarBehaviour", mp_yScrollBarBehaviour);
	oXML->ReadProperty("SelectedCellIndex", mp_lSelectedCellIndex);
	oXML->ReadProperty("SelectedColumnIndex", mp_lSelectedColumnIndex);
	oXML->ReadProperty("SelectedPercentageIndex", mp_lSelectedPercentageIndex);
	oXML->ReadProperty("SelectedPredecessorIndex", mp_lSelectedPredecessorIndex);
	oXML->ReadProperty("SelectedRowIndex", mp_lSelectedRowIndex);
	oXML->ReadProperty("SelectedTaskIndex", mp_lSelectedTaskIndex);
	oXML->ReadProperty("TreeviewColumnIndex", mp_lTreeviewColumnIndex);
	oXML->ReadProperty("TimeBlockBehaviour", mp_yTimeBlockBehaviour);
	oXML->ReadProperty("TierAppearanceScope", mp_yTierAppearanceScope);
    oXML->ReadProperty("TierFormatScope", mp_yTierFormatScope);
	oXML->ReadProperty("TimeLineScrollBarScope", mp_yTimeLineScrollBarScope);
	oXML->ReadProperty("ProgressLineScope", mp_yProgressLineScope);
	oXML->ReadProperty("PredecessorMode", mp_yPredecessorMode);
	oXML->ReadProperty("StyleIndex", mp_sStyleIndex);
	oXML->ReadProperty("Image", *mp_oImage);
	oXML->ReadProperty("ImageTag", mp_sImageTag);
	mp_oConfiguration->SetXML(oXML->ReadObject("Configuration"));
	mp_oStyles->SetXML(oXML->ReadObject("Styles"));
	mp_oRows->SetXML(oXML->ReadObject("Rows"));
	mp_oColumns->SetXML(oXML->ReadObject("Columns"));
	mp_oLayers->SetXML(oXML->ReadObject("Layers"));
	mp_oTasks->SetXML(oXML->ReadObject("Tasks"));
	mp_oPredecessors->SetXML(oXML->ReadObject("Predecessors"));
    mp_oTierAppearance->SetXML(oXML->ReadObject("TierAppearance")); ////Must go before Views
    mp_oTierFormat->SetXML(oXML->ReadObject("TierFormat")); ////Must go before Views
    mp_oTimeLineScrollBar->SetXML(oXML->ReadObject("TimeLineScrollBar")); ////Must go before Views
    mp_oProgressLine->SetXML(oXML->ReadObject("ProgressLine")); ////Must go before Views
    mp_oViews->SetXML(oXML->ReadObject("Views")); ////*********Views
	mp_oTimeBlocks->SetXML(oXML->ReadObject("TimeBlocks"));
	mp_oTimeBlocks->CP_SetXML(oXML->ReadObject("CP_TimeBlocks"));
	mp_oPercentages->SetXML(oXML->ReadObject("Percentages"));
	mp_oSplitter->SetXML(oXML->ReadObject("Splitter"));
	mp_oTreeview->SetXML(oXML->ReadObject("Treeview"));
	mp_oTierAppearance->SetXML(oXML->ReadObject("TierAppearance"));
	mp_oTierFormat->SetXML(oXML->ReadObject("TierFormat"));
	mp_oScrollBarSeparator->SetXML(oXML->ReadObject("ScrollBarSeparator"));
	mp_oVerticalScrollBar->SetXML(oXML->ReadObject("VerticalScrollBar"));
	mp_oHorizontalScrollBar->SetXML(oXML->ReadObject("HorizontalScrollBar"));
	SetStyleIndex(mp_sStyleIndex);
	GetRows()->UpdateTree();
	SetCurrentView(sCurrentView);
	mp_oCurrentView->GetTimeLine()->Position(mp_oCurrentView->GetTimeLine()->GetStartDate());
}

void CActiveGanttVCCtl::mp_ErrorReport(LONG ErrNumber,CString ErrDescription,CString ErrSource)
{
	if (mp_yErrorReports == RE_MSGBOX)
	{
		AfxMessageBox(CStr(ErrNumber) + _T(": ") + ErrDescription + _T(" (") + ErrSource + _T(")"), MB_OK | MB_ICONSTOP | MB_APPLMODAL);
	}
	else if (mp_yErrorReports == RE_HIDE)
	{
	}
	else if (mp_yErrorReports == RE_RAISE)
	{
		COleDispatchException* pOleDispExcep = new COleDispatchException(_T(""),NULL, 0);
		pOleDispExcep->m_wCode = (WORD)ErrNumber;
		pOleDispExcep->m_strSource = ErrSource;
		pOleDispExcep->m_strDescription = ErrDescription;
		throw(pOleDispExcep);
	}
	else if (mp_yErrorReports == RE_RAISEEVENT)
	{
		GetErrorEventArgs()->Clear();
		GetErrorEventArgs()->SetNumber(ErrNumber);
        GetErrorEventArgs()->SetDescription(ErrDescription);
        GetErrorEventArgs()->SetSource(ErrSource);
        FireActiveGanttError();
	}
}

void CActiveGanttVCCtl::ClearSelections(void)
{
	mp_oTextBox->Terminate();
	mp_lSelectedTaskIndex = 0;
	mp_lSelectedColumnIndex = 0;
	mp_lSelectedRowIndex = 0;
	mp_lSelectedCellIndex = 0;
	mp_lSelectedPercentageIndex = 0;
	mp_lSelectedPredecessorIndex = 0;
}

void CActiveGanttVCCtl::Clear(void)
{
	mp_oTextBox->Terminate();
	mp_oTasks->Clear();
	mp_oPercentages->Clear();
	mp_oRows->Clear();
	mp_oStyles->Clear();
	mp_oLayers->Clear();
	mp_oColumns->Clear();
	mp_oTimeBlocks->Clear();
	mp_oViews->Clear();
}

void CActiveGanttVCCtl::CheckPredecessors(void)
{
    int i = 0;
    clsTask* oTask;
    for (i = 1; i <= GetTasks()->GetCount(); i++)
    {
        oTask = (clsTask*)GetTasks()->mp_oCollection->m_oReturnArrayElement(i);
        oTask->mp_bWarning = FALSE;
    }
    if (GetPredecessors()->GetCount() == 0)
    {
        return;
    }
    clsPredecessor* oPredecessor;
    for (i = 1; i <= GetPredecessors()->GetCount(); i++)
    {
        oPredecessor = (clsPredecessor*)GetPredecessors()->mp_oCollection->m_oReturnArrayElement(i);
        oPredecessor->Check(mp_yPredecessorMode);
    }
}

void CActiveGanttVCCtl::ForceEndTextEdit(void)
{
	mp_oTextBox->Terminate();
}

void CActiveGanttVCCtl::SaveToImage(CString Path, LONG Format)
{
	Gdiplus::Bitmap* oBitmap = new Gdiplus::Bitmap(clsG->Width(), clsG->Height(), PixelFormat32bppARGB);
	mp_oGraphics = Graphics::FromImage(oBitmap);
	mp_Draw();
	CLSID lClsid;
	if (Format == IMF_BMP)
	{
		GetEncoderClsid(_T("image/bmp"), &lClsid);
	}
	else if (Format == IMF_JPEG)
	{
		GetEncoderClsid(_T("image/jpeg"), &lClsid);
	}
	else if (Format == IMF_GIF)
	{
		GetEncoderClsid(_T("image/gif"), &lClsid);
	}
	else if (Format == IMF_TIFF)
	{
		GetEncoderClsid(_T("image/tiff"), &lClsid);
	}
	else if (Format == IMF_PNG)
	{
		GetEncoderClsid(_T("image/png"), &lClsid);
	}
	else
	{
		delete oBitmap;
		oBitmap = NULL;
		return;
	}
	oBitmap->Save(Path, &lClsid, NULL);
	delete oBitmap;
	oBitmap = NULL;
	delete mp_oGraphics;
	mp_oGraphics = NULL;
}

Gdiplus::Bitmap* CActiveGanttVCCtl::GetBitmap(void)
{
	return mp_oBitmap;
}

void CActiveGanttVCCtl::AboutBox(void)
{
	PreModalDialog();
	fAbout oForm(sOCXVERSION);
	oForm.DoModal();
	PostModalDialog();
}

void CActiveGanttVCCtl::Redraw(void)
{
	InvalidateControl();
}

void CActiveGanttVCCtl::FireControlClick(void)
{
	odl_ControlClick(GetMouseEventArgs()->GetIDispatch(TRUE));
}

void CActiveGanttVCCtl::FireControlDblClick(void)
{
	odl_ControlDblClick(GetMouseEventArgs()->GetIDispatch(TRUE));
}

void CActiveGanttVCCtl::FireControlMouseDown(void)
{
	odl_ControlMouseDown(GetMouseEventArgs()->GetIDispatch(TRUE));
}

void CActiveGanttVCCtl::FireControlMouseMove(void)
{
	odl_ControlMouseMove(GetMouseEventArgs()->GetIDispatch(TRUE));
}

void CActiveGanttVCCtl::FireControlMouseUp(void)
{
	odl_ControlMouseUp(GetMouseEventArgs()->GetIDispatch(TRUE));
}

void CActiveGanttVCCtl::FireControlMouseHover(void)
{
	odl_ControlMouseHover(GetMouseEventArgs()->GetIDispatch(TRUE));
}

void CActiveGanttVCCtl::FireControlKeyDown(void)
{
	odl_ControlKeyDown(GetKeyEventArgs()->GetIDispatch(TRUE));
}

void CActiveGanttVCCtl::FireControlKeyUp(void)
{
	odl_ControlKeyUp(GetKeyEventArgs()->GetIDispatch(TRUE));
}

void CActiveGanttVCCtl::FireControlKeyPress(void)
{
	odl_ControlKeyPress(GetKeyEventArgs()->GetIDispatch(TRUE));
}

void CActiveGanttVCCtl::FireControlMouseWheel(void)
{
	odl_ControlMouseWheel(GetMouseWheelEventArgs()->GetIDispatch(TRUE));
}

void CActiveGanttVCCtl::FireBeginTextEdit(void)
{
	odl_BeginTextEdit(GetTextEditEventArgs()->GetIDispatch(TRUE));
}

void CActiveGanttVCCtl::FireEndTextEdit(void)
{
	odl_EndTextEdit(GetTextEditEventArgs()->GetIDispatch(TRUE));
}

void CActiveGanttVCCtl::FireDraw(void)
{
	odl_Draw(GetDrawEventArgs()->GetIDispatch(TRUE));
}

void CActiveGanttVCCtl::FirePredecessorDraw(void)
{
	odl_PredecessorDraw(GetPredecessorDrawEventArgs()->GetIDispatch(TRUE));
}

void CActiveGanttVCCtl::FireCustomTierDraw(void)
{
	odl_CustomTierDraw(GetCustomTierDrawEventArgs()->GetIDispatch(TRUE));
}

void CActiveGanttVCCtl::FireTierTextDraw(void)
{
	odl_TierTextDraw(GetCustomTierDrawEventArgs()->GetIDispatch(TRUE));
}

void CActiveGanttVCCtl::FireObjectAdded(void)
{
	odl_ObjectAdded(GetObjectAddedEventArgs()->GetIDispatch(TRUE));
}

void CActiveGanttVCCtl::FireObjectSelected(void)
{
	odl_ObjectSelected(GetObjectSelectedEventArgs()->GetIDispatch(TRUE));
}

void CActiveGanttVCCtl::FireBeginObjectChange(void)
{
	odl_BeginObjectChange(GetObjectStateChangedEventArgs()->GetIDispatch(TRUE));
}

void CActiveGanttVCCtl::FireObjectChange(void)
{
	odl_ObjectChange(GetObjectStateChangedEventArgs()->GetIDispatch(TRUE));
}

void CActiveGanttVCCtl::FireEndObjectChange(void)
{
	odl_EndObjectChange(GetObjectStateChangedEventArgs()->GetIDispatch(TRUE));
}

void CActiveGanttVCCtl::FireCompleteObjectChange(void)
{
	odl_CompleteObjectChange(GetObjectStateChangedEventArgs()->GetIDispatch(TRUE));
}

void CActiveGanttVCCtl::FireActiveGanttError(void)
{
	odl_ActiveGanttError(GetErrorEventArgs()->GetIDispatch(TRUE));
}

void CActiveGanttVCCtl::FireControlScroll(void)
{
	odl_ControlScroll(GetScrollEventArgs()->GetIDispatch(TRUE));
}

void CActiveGanttVCCtl::FireNodeChecked(void)
{
	odl_NodeChecked(GetNodeEventArgs()->GetIDispatch(TRUE));
}

void CActiveGanttVCCtl::FireNodeExpanded(void)
{
	odl_NodeExpanded(GetNodeEventArgs()->GetIDispatch(TRUE));
}

void CActiveGanttVCCtl::FireControlDraw(void)
{
	odl_ControlDraw();
}

void CActiveGanttVCCtl::FireControlRedrawn(void)
{
	odl_ControlRedrawn();
}

void CActiveGanttVCCtl::FireTimeLineChanged(void)
{
	odl_TimeLineChanged();
}

void CActiveGanttVCCtl::FirePredecessorException(void)
{
	odl_PredecessorException(GetPredecessorExceptionEventArgs()->GetIDispatch(TRUE));
}

void CActiveGanttVCCtl::FireCustomTickMarkAreaDraw(void)
{
	odl_CustomTickMarkAreaDraw(GetCustomTickMarkAreaDrawEventArgs()->GetIDispatch(TRUE));
}

void CActiveGanttVCCtl::FireCustomTickMarkDraw(void)
{
	odl_CustomTickMarkDraw(GetCustomTickMarkAreaDrawEventArgs()->GetIDispatch(TRUE));
}

void CActiveGanttVCCtl::FireTickMarkTextDraw(void)
{
	odl_TickMarkTextDraw(GetTickMarkTextDrawEventArgs()->GetIDispatch(TRUE));
}

void CActiveGanttVCCtl::FirePagePrint(void)
{
	odl_PagePrint(GetPagePrintEventArgs()->GetIDispatch(TRUE));
}

void CActiveGanttVCCtl::FireToolTipOnMouseHover(LONG EventTarget)
{
	if (mp_oCurrentView->GetClientArea()->GetToolTipsVisible() == FALSE) 
	{
		return;
	}
	mp_oToolTipEventArgs->SetToolTipType(TPT_HOVER);
	mp_oToolTipEventArgs->SetEventTarget(EventTarget);
	odl_ToolTipOnMouseHover(GetToolTipEventArgs()->GetIDispatch(TRUE));
}

void CActiveGanttVCCtl::FireToolTipOnMouseMove(LONG Operation)
{
	if (mp_oCurrentView->GetClientArea()->GetToolTipsVisible() == FALSE) 
	{
		return;
	}
	mp_oToolTipEventArgs->SetToolTipType(TPT_MOVEMENT);
	mp_oToolTipEventArgs->SetOperation(Operation);
	odl_ToolTipOnMouseMove(GetToolTipEventArgs()->GetIDispatch(TRUE));
}

void CActiveGanttVCCtl::FireOnMouseHoverToolTipDraw(LONG EventTarget)
{
	if (mp_oCurrentView->GetClientArea()->GetToolTipsVisible() == FALSE) 
	{
		return;
	}
	mp_oToolTipEventArgs->SetX1(GetToolTip()->GetLeft());
	mp_oToolTipEventArgs->SetY1(GetToolTip()->GetTop());
	mp_oToolTipEventArgs->SetX2(GetToolTip()->GetRight() - 1);
	mp_oToolTipEventArgs->SetY2(GetToolTip()->GetBottom() - 1);
	odl_OnMouseHoverToolTipDraw(GetToolTipEventArgs()->GetIDispatch(TRUE));
}

void CActiveGanttVCCtl::FireOnMouseMoveToolTipDraw(LONG Operation)
{
	if (mp_oCurrentView->GetClientArea()->GetToolTipsVisible() == FALSE) 
	{
		return;
	}
	mp_oToolTipEventArgs->SetX1(GetToolTip()->GetLeft());
	mp_oToolTipEventArgs->SetY1(GetToolTip()->GetTop());
	mp_oToolTipEventArgs->SetX2(GetToolTip()->GetRight() - 1);
	mp_oToolTipEventArgs->SetY2(GetToolTip()->GetBottom() - 1);
	odl_OnMouseMoveToolTipDraw(GetToolTipEventArgs()->GetIDispatch(TRUE));
}

clsTimeBlocks* CActiveGanttVCCtl::TempTimeBlocks(void)
{
	return tmpTimeBlocks;
}

void CActiveGanttVCCtl::mp_Draw(void)
{
	FireControlDraw();
 
	clsG->mp_ResetFocusRectangle();
	clsG->mp_ClipRegion(0, 0, clsG->Width(), clsG->Height(), FALSE);
	clsG->mp_DrawItem(0, 0, clsG->Width() - 1, clsG->Height() - 1, "", "", FALSE, GetImage(), 0, 0, GetStyle());
	mp_oCurrentView->GetTimeLine()->Calculate();
	mp_PositionScrollBars();
	mp_oColumns->Position();
	mp_oRows->InitializePosition();
	mp_oRows->PositionRows();
	mp_oColumns->Draw();
	mp_oRows->Draw();
	GetTreeview()->Draw();
	mp_oCurrentView->GetTimeLine()->Draw();
	mp_oTimeBlocks->CreateTemporaryTimeBlocks();
	mp_oTimeBlocks->Draw();
	mp_oCurrentView->GetClientArea()->GetGrid()->Draw();
	mp_oCurrentView->GetClientArea()->Draw();
	mp_oPredecessors->Draw();
	mp_oTasks->Draw();
	mp_oPercentages->Draw();
	f_ProgressLine()->Draw();
	mp_oSplitter->Draw();
	clsG->mp_ClipRegion(0, 0, clsG->Width(), clsG->Height(), TRUE);
    if (GetVerticalScrollBar()->GetState() == SS_SHOWN)
    {
        clsG->mp_DrawItem(GetVerticalScrollBar()->GetLeft(), GetVerticalScrollBar()->GetTop() + GetVerticalScrollBar()->GetHeight(), GetVerticalScrollBar()->GetLeft() + 16, GetVerticalScrollBar()->GetTop() + GetVerticalScrollBar()->GetHeight() + 16, "", "", FALSE, NULL, 0, 0, GetScrollBarSeparator()->GetStyle());
        clsG->mp_ClipRegion(0, 0, clsG->Width(), clsG->Height(), FALSE);
    }
    else if (f_TimeLineScrollBar()->GetState() == SS_SHOWN)
    {
        clsG->mp_DrawItem(f_TimeLineScrollBar()->GetLeft() + f_TimeLineScrollBar()->GetWidth(), f_TimeLineScrollBar()->GetTop(), f_TimeLineScrollBar()->GetLeft() + f_TimeLineScrollBar()->GetWidth() + 16, f_TimeLineScrollBar()->GetTop() + 16, "", "", FALSE, NULL, 0, 0, GetScrollBarSeparator()->GetStyle());
        clsG->mp_ClipRegion(0, 0, clsG->Width(), clsG->Height(), FALSE);
    }
	mp_DrawDebugMetrics();
	if (mp_oVerticalScrollBar->GetState() == SS_SHOWN) 
	{
		mp_oVerticalScrollBar->GetScrollBar()->Draw();
	}
	if (mp_oHorizontalScrollBar->GetState() == SS_SHOWN) 
	{
		mp_oHorizontalScrollBar->GetScrollBar()->Draw();
	}
	if (f_TimeLineScrollBar()->GetState() == SS_SHOWN) 
	{
		f_TimeLineScrollBar()->GetScrollBar()->Draw();
	}

#ifdef _TRIAL
	if (clsG->GetCustomPrinting() == FALSE)
	{
		clsFont oFont;
		oFont.SetFontName("Arial");
		oFont.SetSize(12);
		oFont.SetStyle(FS_FONTSTYLEBOLD);
		Gdiplus::Color oColor = Color::MakeARGB(255, GetMathLib()->Rnd(), GetMathLib()->Rnd(), GetMathLib()->Rnd());
		clsG->mp_DrawAlignedText(20, 20, clsG->Width() - 20, clsG->Height() - 20, _T("ActiveGanttVC Trial Version\r\nFor evaluation purposes only"), HAL_RIGHT, VAL_BOTTOM, oColor, &oFont);
	}
#endif

	FireControlRedrawn();
}

void CActiveGanttVCCtl::mp_DrawDesignMode(void)
{
	//LONG lWidth = clsG->Width();
	//LONG lHeight = clsG->Height();
	//Gdiplus::Color clrTest;
	//clrTest = Color::Blue;

	//clsG->DrawLine(500, 500, clsG->Width(), clsG->Height(), LT_FILLED, clrTest, LDS_SOLID);

	clsG->mp_DrawLine(0, 0, clsG->Width(), clsG->Height(), LT_FILLED, Color::White, LDS_SOLID);

	LONG lLeftBox = 0;
	LONG lTop = 0;
	LONG lRightBox = 0;
	LONG lBottom = 0;
	LONG lLeftCA = 0;
	LONG lRightCA = 0;
	clsFont oFont;
	oFont.SetPtFont(_T("Tahoma"), 8, FS_FONTSTYLEREGULAR);
	mp_oCurrentView->GetTimeLine()->Calculate();

	mp_oCurrentView->GetTimeLine()->Draw();
	clsG->mp_ClearClipRegion();
	lLeftBox = Getmt_LeftMargin();
	lTop = Getmt_TopMargin();
	lRightBox = GetSplitter()->GetLeft();
	lBottom = mp_oCurrentView->GetTimeLine()->GetBottom();
	clsG->mp_DrawEdge(lLeftBox, lTop, lRightBox, lBottom, Color::MakeARGB(255, 192, 192, 192), BT_NORMALWINDOWS, ET_RAISED, TRUE, NULL);
	clsG->mp_DrawAlignedText(lLeftBox, lTop, lRightBox, lBottom, "Column", HAL_CENTER, VAL_CENTER, Color::Black, &oFont);
	lLeftBox = Getmt_LeftMargin();
	lTop = mp_oCurrentView->GetClientArea()->GetTop();
	lRightBox = GetSplitter()->GetLeft();
	lBottom = mp_oCurrentView->GetClientArea()->GetTop() + 40;
	lLeftCA = GetSplitter()->GetRight();
	lRightCA = Getmt_RightMargin();
	clsG->mp_DrawEdge(lLeftBox, lTop, lRightBox, lBottom, Color::MakeARGB(255, 192, 192, 192), BT_NORMALWINDOWS, ET_RAISED, TRUE, NULL);
	clsG->mp_DrawAlignedText(lLeftBox, lTop, lRightBox, lBottom, "Cell", HAL_CENTER, VAL_CENTER, Color::Black, &oFont);
    clsG->mp_DrawLine(lLeftCA, lBottom, lRightCA, lBottom, LT_NORMAL, mp_oCurrentView->GetClientArea()->GetGrid()->GetColor(), LDS_SOLID);
	mp_oRows->SetTopOffset(mp_oCurrentView->GetClientArea()->GetTop() + 40);
	mp_oCurrentView->GetClientArea()->Setf_LastVisibleRow(1);
	mp_oSplitter->Draw();

	clsG->mp_ClipRegion(0, 0, clsG->Width(), clsG->Height(), TRUE);
	clsG->mp_DrawEdge(0, 0, clsG->Width() - 1, clsG->Height() - 1, Color::Black, BT_NORMALWINDOWS, ET_SUNKEN, FALSE, NULL);
}

void CActiveGanttVCCtl::mp_DrawDebugMetrics(void)
{
}

clsProgressLine* CActiveGanttVCCtl::f_ProgressLine(void)
{
    if (mp_yProgressLineScope == OS_CONTROL)
    {
        return mp_oProgressLine;
    }
    else
    {
        return GetCurrentViewObject()->GetTimeLine()->GetProgressLine();
    }
}

clsTimeLineScrollBar* CActiveGanttVCCtl::f_TimeLineScrollBar(void)
{
	if (mp_yTimeLineScrollBarScope == OS_CONTROL)
    {
        return mp_oTimeLineScrollBar;
    }
    else
    {
        return GetCurrentViewObject()->GetTimeLine()->GetTimeLineScrollBar();
    }
}

clsGDIGraphics* CActiveGanttVCCtl::GetGraphicsObject(void)
{
	return NULL;
}

void CActiveGanttVCCtl::ReleaseGraphicsObject(void)
{
}

void CActiveGanttVCCtl::f_Draw(void)
{
	mp_Draw();
}

HBRUSH CActiveGanttVCCtl::OnCtlColor(CDC* pDC, CWnd* pWnd, UINT nCtlColor)
{
	HBRUSH hbr = COleControl::OnCtlColor(pDC, pWnd, nCtlColor);

	if (pWnd->GetDlgCtrlID() == 15620) //clsTextBox
	{
		pDC->SetTextColor(mp_oTextBox->mp_clrForeColor.ToCOLORREF());
		pDC->SetBkColor(mp_oTextBox->mp_clrBackColor.ToCOLORREF());
		mp_oBrush.Detach();
		mp_oBrush.CreateSolidBrush(mp_oTextBox->mp_clrBackColor.ToCOLORREF());
		return (HBRUSH)mp_oBrush.GetSafeHandle();
	}
	return hbr;
}

void CActiveGanttVCCtl::OnUpdateEdit()
{
	mp_oTextBox->OnUpdateEdit();
}