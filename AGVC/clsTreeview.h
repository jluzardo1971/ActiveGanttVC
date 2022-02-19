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

class clsHeader;

class clsTreeview : public CCmdTarget
{
	DECLARE_DYNCREATE(clsTreeview)
	clsTreeview();

public:
	CActiveGanttVCCtl* mp_oControl;

	clsTreeview(CActiveGanttVCCtl* oControl);
	virtual ~clsTreeview();
	virtual void OnFinalRelease();

// Member Variables
public:
	LONG mp_lLastVisibleNode;
	LONG mp_lIndentation;
	Gdiplus::Color mp_clrBackColor;
	Gdiplus::Color mp_clrCheckBoxBorderColor;
	Gdiplus::Color mp_clrCheckBoxBackColor;
	Gdiplus::Color mp_clrCheckBoxMarkColor;
	LONG mp_yCheckBoxBackgroundMode;
	Gdiplus::Color mp_clrSelectedBackColor;
	Gdiplus::Color mp_clrSelectedForeColor;
	Gdiplus::Color mp_clrTreeLineColor;
	Gdiplus::Color mp_clrPlusMinusBorderColor;
	Gdiplus::Color mp_clrPlusMinusSignColor;
	Gdiplus::Color mp_clrPlusMinusBackColor;
	Gdiplus::Color mp_clrExpandIconForeColor;
	Gdiplus::Color mp_clrExpandIconBackColor;
	Gdiplus::Color mp_clrExpandIconDropShadowColor;
	Gdiplus::Color mp_clrCollapseIconForeColor;
	Gdiplus::Color mp_clrCollapseIconDropShadowColor;
	LONG mp_yPlusMinusBackgroundMode;
	BOOL mp_bCheckBoxes;
	BOOL mp_bTreeLines;
	BOOL mp_bImages;
	LONG mp_yTreeviewMode;
	BOOL mp_bFullColumnSelect;
	BOOL mp_bExpansionOnSelection;
	CString mp_sPathSeparator;
	LONG mp_yOperation;
	LONG mp_lFirstVisibleNode;								

//Internal Properties (Extensions to ODL)
public:
	LONG GetFirstVisibleNode(void);
	void SetFirstVisibleNode(LONG Value);
	LONG GetLastVisibleNode(void);
	LONG GetIndentation(void);
	void SetIndentation(LONG Value);
	Gdiplus::Color GetCheckBoxBorderColor(void);
	void SetCheckBoxBorderColor(Gdiplus::Color Value);
	Gdiplus::Color GetCheckBoxBackColor(void);
	void SetCheckBoxBackColor(Gdiplus::Color Value);
	Gdiplus::Color GetCheckBoxMarkColor(void);
	void SetCheckBoxMarkColor(Gdiplus::Color Value);
	Gdiplus::Color GetBackColor(void);
	void SetBackColor(Gdiplus::Color Value);
	CString GetPathSeparator(void);
	void SetPathSeparator(CString Value);
	BOOL GetTreeLines(void);
	void SetTreeLines(BOOL Value);
	LONG GetTreeviewMode(void);
	void SetTreeviewMode(LONG Value);
	BOOL GetImages(void);
	void SetImages(BOOL Value);
	BOOL GetCheckBoxes(void);
	void SetCheckBoxes(BOOL Value);
	BOOL GetFullColumnSelect(void);
	void SetFullColumnSelect(BOOL Value);
	BOOL GetExpansionOnSelection(void);
	void SetExpansionOnSelection(BOOL Value);
	Gdiplus::Color GetSelectedBackColor(void);
	void SetSelectedBackColor(Gdiplus::Color Value);
	Gdiplus::Color GetSelectedForeColor(void);
	void SetSelectedForeColor(Gdiplus::Color Value);
	Gdiplus::Color GetTreeLineColor(void);
	void SetTreeLineColor(Gdiplus::Color Value);
	Gdiplus::Color GetPlusMinusBorderColor(void);
	void SetPlusMinusBorderColor(Gdiplus::Color Value);
	Gdiplus::Color GetPlusMinusSignColor(void);
	void SetPlusMinusSignColor(Gdiplus::Color Value);
	clsHeader* GetHeader(void);
	CString GetStyleIndex(void);
	void SetStyleIndex(CString Value);
	clsStyle* GetStyle(void);
	clsImage* GetImage(void);
	CString GetImageTag(void);
	void SetImageTag(CString Value);
	BOOL mp_bCursorEditTextNode(int X, int Y);
	BOOL mp_bShowEditTextNode(int X, int Y);
	LONG GetCheckBoxBackgroundMode(void);
	void SetCheckBoxBackgroundMode(LONG Value);
	Gdiplus::Color GetPlusMinusBackColor(void);
	void SetPlusMinusBackColor(Gdiplus::Color Value);
	LONG GetPlusMinusBackgroundMode(void);
	void SetPlusMinusBackgroundMode(LONG Value);
	Gdiplus::Color GetExpandIconForeColor(void);
	void SetExpandIconForeColor(Gdiplus::Color Value);
	Gdiplus::Color GetExpandIconBackColor(void);
	void SetExpandIconBackColor(Gdiplus::Color Value);
	Gdiplus::Color GetExpandIconDropShadowColor(void);
	void SetExpandIconDropShadowColor(Gdiplus::Color Value);
	Gdiplus::Color GetCollapseIconForeColor(void);
	void SetCollapseIconForeColor(Gdiplus::Color Value);
	Gdiplus::Color GetCollapseIconDropShadowColor(void);
	void SetCollapseIconDropShadowColor(Gdiplus::Color Value);


	LONG GetLeft(void);
	LONG GetRight(void);
	LONG GetLeftTrim(void);
	LONG GetRightTrim(void);


//Internal Properties
public:
	LONG Getf_FirstVisibleNode(void);
	void Setf_LastVisibleNode(LONG Value);
	LONG Getmt_BorderThickness(void);

//Internal Methods
public:
	BOOL OverControl(LONG X,LONG Y);
	void OnMouseHover(LONG X,LONG Y);
	void OnMouseDown(LONG X,LONG Y);
	void OnMouseMove(LONG X,LONG Y);
	void OnMouseUp(LONG X,LONG Y);
	LONG CursorPosition(LONG X,LONG Y);
	void mp_EO_CHECKBOXCLICK(LONG yMouseKeyBoardEvent,LONG X,LONG Y);
	void mp_EO_SIGNCLICK(LONG yMouseKeyBoardEvent,LONG X,LONG Y);
	void mp_EO_ROWMOVEMENT(LONG yMouseKeyBoardEvent,LONG X,LONG Y);
	void mp_EO_ROWSIZING(LONG yMouseKeyBoardEvent,LONG X,LONG Y);
	void mp_EO_ROWSELECTION(LONG yMouseKeyBoardEvent,LONG X,LONG Y);
	BOOL mp_bOverCheckBox(LONG X,LONG Y);
	BOOL mp_bOverPlusMinusSign(LONG X,LONG Y);
	void Draw(void);
	void ClearSelections(void);
	void SetXML(CString sXML);
	CString GetXML(void);

	struct S_CHECKBOXCLICK
	{
		LONG lNodeIndex;
		void Clear()
		{
			lNodeIndex = 0;
		}
	};

	struct S_SIGNCLICK
	{
		LONG lNodeIndex;
		void Clear()
		{
			lNodeIndex = 0;
		}
	};

	struct S_ROWMOVEMENT
	{
		LONG lRowIndex;
		LONG lDestinationRowIndex;
		void Clear()
		{
			lRowIndex = 0;
			lDestinationRowIndex = 0;
		}
	};

	struct S_ROWSIZING
	{
		LONG lRowIndex;
		void Clear()
		{
			lRowIndex = 0;
		}
	};

	struct S_ROWSELECTION
	{
		LONG lRowIndex;
		LONG lCellIndex;
		void Clear()
		{
			lRowIndex = 0;
			lCellIndex = 0;
		}
	};

	S_CHECKBOXCLICK s_chkCLK;
	S_SIGNCLICK s_sgnCLK;
	S_ROWMOVEMENT s_rowMVT;
	S_ROWSIZING s_rowSZ;
	S_ROWSELECTION s_rowSEL;

protected:
	DECLARE_DISPATCH_MAP()
	DECLARE_MESSAGE_MAP()
	DECLARE_OLECREATE(clsTreeview)
	DECLARE_INTERFACE_MAP()

LONG odl_GetFirstVisibleNode(void)
{
	
	return GetFirstVisibleNode();
}

void odl_SetFirstVisibleNode(LONG Value)
{
	
	SetFirstVisibleNode(Value);
}

LONG odl_GetLastVisibleNode(void)
{
	
	return GetLastVisibleNode();
}

LONG odl_GetIndentation(void)
{
	
	return GetIndentation();
}

void odl_SetIndentation(LONG Value)
{
	
	SetIndentation(Value);
}

OLE_COLOR odl_GetCheckBoxBorderColor(void)
{
	
	return (OLE_COLOR) GetCheckBoxBorderColor().ToCOLORREF();
}

void odl_SetCheckBoxBorderColor(OLE_COLOR Value)
{
	
	Gdiplus::Color oColor;
	oColor.SetFromCOLORREF(g_TranslateColor(Value));
	SetCheckBoxBorderColor(oColor);
}

OLE_COLOR odl_GetCheckBoxColor(void)
{
	
	return (OLE_COLOR) GetCheckBoxBackColor().ToCOLORREF();
}

void odl_SetCheckBoxColor(OLE_COLOR Value)
{
	
	Gdiplus::Color oColor;
	oColor.SetFromCOLORREF(g_TranslateColor(Value));
	SetCheckBoxBackColor(oColor);
}

OLE_COLOR odl_GetCheckBoxMarkColor(void)
{
	
	return (OLE_COLOR) GetCheckBoxMarkColor().ToCOLORREF();
}

void odl_SetCheckBoxMarkColor(OLE_COLOR Value)
{
	
	Gdiplus::Color oColor;
	oColor.SetFromCOLORREF(g_TranslateColor(Value));
	SetCheckBoxMarkColor(oColor);
}

OLE_COLOR odl_GetBackColor(void)
{
	
	return (OLE_COLOR) GetBackColor().ToCOLORREF();
}

void odl_SetBackColor(OLE_COLOR Value)
{
	
	Gdiplus::Color oColor;
	oColor.SetFromCOLORREF(g_TranslateColor(Value));
	SetBackColor(oColor);
}

BSTR odl_GetPathSeparator(void)
{
	
	return GetPathSeparator().AllocSysString();
}

void odl_SetPathSeparator(LPCTSTR Value)
{
	
	SetPathSeparator(Value);
}

BOOL odl_GetTreeLines(void)
{
	
	return GetTreeLines();
}

void odl_SetTreeLines(BOOL Value)
{
	
	
	SetTreeLines(Value);
}

LONG odl_GetTreeviewMode(void)
{
	
	return GetTreeviewMode();
}

void odl_SetTreeviewMode(LONG Value)
{
	
	
	SetTreeviewMode(Value);
}

BOOL odl_GetImages(void)
{
	
	return GetImages();
}

void odl_SetImages(BOOL Value)
{
	
	
	SetImages(Value);
}

BOOL odl_GetCheckBoxes(void)
{
	
	return GetCheckBoxes();
}

void odl_SetCheckBoxes(BOOL Value)
{
	
	
	SetCheckBoxes(Value);
}

BOOL odl_GetFullColumnSelect(void)
{
	
	return GetFullColumnSelect();
}

void odl_SetFullColumnSelect(BOOL Value)
{
	
	
	SetFullColumnSelect(Value);
}

BOOL odl_GetExpansionOnSelection(void)
{
	
	return GetExpansionOnSelection();
}

void odl_SetExpansionOnSelection(BOOL Value)
{
	
	
	SetExpansionOnSelection(Value);
}

OLE_COLOR odl_GetSelectedBackColor(void)
{
	
	return (OLE_COLOR) GetSelectedBackColor().ToCOLORREF();
}

void odl_SetSelectedBackColor(OLE_COLOR Value)
{
	
	Gdiplus::Color oColor;
	oColor.SetFromCOLORREF(g_TranslateColor(Value));
	SetSelectedBackColor(oColor);
}

OLE_COLOR odl_GetSelectedForeColor(void)
{
	
	return (OLE_COLOR) GetSelectedForeColor().ToCOLORREF();
}

void odl_SetSelectedForeColor(OLE_COLOR Value)
{
	
	Gdiplus::Color oColor;
	oColor.SetFromCOLORREF(g_TranslateColor(Value));
	SetSelectedForeColor(oColor);
}

OLE_COLOR odl_GetTreeLineColor(void)
{
	
	return (OLE_COLOR) GetTreeLineColor().ToCOLORREF();
}

void odl_SetTreeLineColor(OLE_COLOR Value)
{
	
	Gdiplus::Color oColor;
	oColor.SetFromCOLORREF(g_TranslateColor(Value));
	SetTreeLineColor(oColor);
}

OLE_COLOR odl_GetPlusMinusBorderColor(void)
{
	
	return (OLE_COLOR) GetPlusMinusBorderColor().ToCOLORREF();
}

void odl_SetPlusMinusBorderColor(OLE_COLOR Value)
{
	
	Gdiplus::Color oColor;
	oColor.SetFromCOLORREF(g_TranslateColor(Value));
	SetPlusMinusBorderColor(oColor);
}

OLE_COLOR odl_GetPlusMinusSignColor(void)
{
	
	return (OLE_COLOR) GetPlusMinusSignColor().ToCOLORREF();
}

void odl_SetPlusMinusSignColor(OLE_COLOR Value)
{
	
	Gdiplus::Color oColor;
	oColor.SetFromCOLORREF(g_TranslateColor(Value));
	SetPlusMinusSignColor(oColor);
}

OLE_COLOR odl_GetPlusMinusBackColor(void)
{
	
	return (OLE_COLOR) GetPlusMinusBackColor().ToCOLORREF();
}

void odl_SetPlusMinusBackColor(OLE_COLOR Value)
{
	
	Gdiplus::Color oColor;
	oColor.SetFromCOLORREF(g_TranslateColor(Value));
	SetPlusMinusBackColor(oColor);
}
	
LONG odl_GetPlusMinusBackgroundMode(void)
{
	
	return GetPlusMinusBackgroundMode();
}

void odl_SetPlusMinusBackgroundMode(LONG Value)
{
	
	SetPlusMinusBackgroundMode(Value);
}

void odl_ClearSelections(void)
{
	
	ClearSelections();
}

void odl_SetXML(LPCTSTR sXML)
{
	
	SetXML(sXML);
}

BSTR odl_GetXML(void)
{
	
	return GetXML().AllocSysString();
}

LONG odl_GetCheckBoxBackgroundMode(void)
{
	
	return GetCheckBoxBackgroundMode();
}

void odl_SetCheckBoxBackgroundMode(LONG Value)
{
	
	SetCheckBoxBackgroundMode(Value);
}

 
 
 
 
 

OLE_COLOR odl_GetExpandIconForeColor(void)
{
	
	return (OLE_COLOR) GetExpandIconForeColor().ToCOLORREF();
}

void odl_SetExpandIconForeColor(OLE_COLOR Value)
{
	
	Gdiplus::Color oColor;
	oColor.SetFromCOLORREF(g_TranslateColor(Value));
	SetExpandIconForeColor(oColor);
}

OLE_COLOR odl_GetExpandIconBackColor(void)
{
	
	return (OLE_COLOR) GetExpandIconBackColor().ToCOLORREF();
}

void odl_SetExpandIconBackColor(OLE_COLOR Value)
{
	
	Gdiplus::Color oColor;
	oColor.SetFromCOLORREF(g_TranslateColor(Value));
	SetExpandIconBackColor(oColor);
}

OLE_COLOR odl_GetExpandIconDropShadowColor(void)
{
	
	return (OLE_COLOR) GetExpandIconDropShadowColor().ToCOLORREF();
}

void odl_SetExpandIconDropShadowColor(OLE_COLOR Value)
{
	
	Gdiplus::Color oColor;
	oColor.SetFromCOLORREF(g_TranslateColor(Value));
	SetExpandIconDropShadowColor(oColor);
}

OLE_COLOR odl_GetCollapseIconForeColor(void)
{
	
	return (OLE_COLOR) GetCollapseIconForeColor().ToCOLORREF();
}

void odl_SetCollapseIconForeColor(OLE_COLOR Value)
{
	
	Gdiplus::Color oColor;
	oColor.SetFromCOLORREF(g_TranslateColor(Value));
	SetCollapseIconForeColor(oColor);
}

OLE_COLOR odl_GetCollapseIconDropShadowColor(void)
{
	
	return (OLE_COLOR) GetCollapseIconDropShadowColor().ToCOLORREF();
}

void odl_SetCollapseIconDropShadowColor(OLE_COLOR Value)
{
	
	Gdiplus::Color oColor;
	oColor.SetFromCOLORREF(g_TranslateColor(Value));
	SetCollapseIconDropShadowColor(oColor);
}




};