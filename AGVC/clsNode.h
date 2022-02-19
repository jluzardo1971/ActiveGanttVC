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

class clsRow;
class clsFont;
class clsImage;
class clsNode;

class clsNode : public CCmdTarget
{
	DECLARE_DYNCREATE(clsNode)
	clsNode();

public:
	CActiveGanttVCCtl* mp_oControl;

	clsNode(CActiveGanttVCCtl* oControl, clsRow* oRow);
	virtual ~clsNode();
	virtual void OnFinalRelease();

// Member Variables
public:
	BOOL mp_bExpanded;
	CString mp_sTag;
	clsImage* mp_oImage;
	clsImage* mp_oExpandedImage;
	clsImage* mp_oSelectedImage;
	BOOL mp_bChecked;
	LONG mp_lDepth;
	clsRow* mp_oRow;
	BOOL mp_bCheckBoxVisible;
	BOOL mp_bImageVisible;
	clsNode* mp_oParent;
	CString mp_sExpandedImageTag;
	CString mp_sSelectedImageTag;
	CString mp_sImageTag;
	CString mp_sStyleIndex;
	clsStyle* mp_oStyle;
	LONG mp_lTextLeft;
    LONG mp_lTextTop;
    LONG mp_lTextRight;
    LONG mp_lTextBottom;
	BOOL mp_bAllowTextEdit;

//Internal Properties (Extensions to ODL)
public:
	clsRow* GetRow(void);
	BOOL GetCheckBoxVisible(void);
	void SetCheckBoxVisible(BOOL Value);
	BOOL GetImageVisible(void);
	void SetImageVisible(BOOL Value);
	LONG GetLeft(void);
	LONG GetTop(void);
	LONG GetBottom(void);
	LONG GetDepth(void);
	void SetDepth(LONG Value);
	BOOL GetExpanded(void);
	void SetExpanded(BOOL Value);
	BOOL GetSelected(void);
	void SetSelected(BOOL Value);
	clsImage* GetExpandedImage(void);
	clsImage* GetSelectedImage(void);
	clsImage* GetImage(void);
	CString GetTag(void);
	void SetTag(CString Value);
	CString GetText(void);
	void SetText(CString Value);
	BOOL GetChecked(void);
	void SetChecked(BOOL Value);
	LONG GetHeight(void);
	void SetHeight(LONG Value);
	BOOL GetHidden(void);
	CString GetExpandedImageTag(void);
	void SetExpandedImageTag(CString Value);
	CString GetSelectedImageTag(void);
	void SetSelectedImageTag(CString Value);
	CString GetImageTag(void);
	void SetImageTag(CString Value);
	CString GetStyleIndex(void);
	void SetStyleIndex(CString Value);
	clsStyle* GetStyle(void);
	BOOL GetAllowTextEdit(void);
	void SetAllowTextEdit(BOOL Value);

//Internal Properties
public:
	LONG GetYCenter(void);
	LONG Getmt_TextLeft(void);
	LONG Getmt_TextTop(void);
	LONG Getmt_TextRight(void);
	LONG Getmt_TextBottom(void);
	LONG GetCheckBoxLeft(void);
	LONG GetImageLeft(void);
	LONG GetImageTop(void);
	LONG GetImageRight(void);
	LONG GetImageBottom(void);
	LONG GetIndex(void);

//Internal Methods
public:
	clsNode* NextSibling(void);
	clsNode* PreviousSibling(void);
	BOOL IsFirstSibling(void);
	clsNode* FirstSibling(void);
	BOOL IsLastSibling(void);
	clsNode* LastSibling(void);
	clsNode* Child(void);
	clsNode* Parent(void);
	BOOL IsRoot(void);
	CString FullPath(void);
	LONG Children(void);
	BOOL RecurseHidden(clsNode* oNode,BOOL bHidden);
	CString GetXML(void);
	void SetXML(CString sXML);

protected:
	DECLARE_DISPATCH_MAP()
	DECLARE_MESSAGE_MAP()
	DECLARE_OLECREATE(clsNode)
	DECLARE_INTERFACE_MAP()

IDispatch* odl_GetRow(void); //Circular Reference

BOOL odl_GetAllowTextEdit(void)
{
	
	return GetAllowTextEdit();
}

void odl_SetAllowTextEdit(BOOL Value)
{
	
	
	SetAllowTextEdit(Value);
}




BOOL odl_GetCheckBoxVisible(void)
{
	
	return GetCheckBoxVisible();
}

void odl_SetCheckBoxVisible(BOOL Value)
{
	
	
	SetCheckBoxVisible(Value);
}

BOOL odl_GetImageVisible(void)
{
	
	return GetImageVisible();
}

void odl_SetImageVisible(BOOL Value)
{
	
	
	SetImageVisible(Value);
}

LONG odl_GetLeft(void)
{
	
	return GetLeft();
}

LONG odl_GetTop(void)
{
	
	return GetTop();
}

LONG odl_GetBottom(void)
{
	
	return GetBottom();
}

LONG odl_GetDepth(void)
{
	
	return GetDepth();
}

void odl_SetDepth(LONG Value)
{
	
	SetDepth(Value);
}

BOOL odl_GetExpanded(void)
{
	
	return GetExpanded();
}

void odl_SetExpanded(BOOL Value)
{
	
	
	SetExpanded(Value);
}

BOOL odl_GetSelected(void)
{
	
	return GetSelected();
}

void odl_SetSelected(BOOL Value)
{
	
	
	SetSelected(Value);
}

IDispatch* odl_GetExpandedImage(void)
{
	
	return GetExpandedImage()->GetIDispatch(TRUE);
}

IDispatch* odl_GetSelectedImage(void)
{
	
	return GetSelectedImage()->GetIDispatch(TRUE);
}

IDispatch* odl_GetImage(void)
{
	
	return GetImage()->GetIDispatch(TRUE);
}

BSTR odl_GetTag(void)
{
	
	return GetTag().AllocSysString();
}

void odl_SetTag(LPCTSTR Value)
{
	
	SetTag(Value);
}

BSTR odl_GetText(void)
{
	
	return GetText().AllocSysString();
}

void odl_SetText(LPCTSTR Value)
{
	
	SetText(Value);
}

BOOL odl_GetChecked(void)
{
	
	return GetChecked();
}

void odl_SetChecked(BOOL Value)
{
	
	
	SetChecked(Value);
}

LONG odl_GetHeight(void)
{
	
	return GetHeight();
}

void odl_SetHeight(LONG Value)
{
	
	SetHeight(Value);
}

BOOL odl_GetHidden(void)
{
	
	return GetHidden();
}

BSTR odl_GetExpandedImageTag(void)
{
	
	return GetExpandedImageTag().AllocSysString();
}

void odl_SetExpandedImageTag(LPCTSTR Value)
{
	
	SetExpandedImageTag(Value);
}

BSTR odl_GetSelectedImageTag(void)
{
	
	return GetSelectedImageTag().AllocSysString();
}

void odl_SetSelectedImageTag(LPCTSTR Value)
{
	
	SetSelectedImageTag(Value);
}

BSTR odl_GetImageTag(void)
{
	
	return GetImageTag().AllocSysString();
}

void odl_SetImageTag(LPCTSTR Value)
{
	
	SetImageTag(Value);
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

IDispatch* odl_NextSibling(void)
{
	
	clsNode* oNode = NextSibling();
	if (oNode != NULL)
	{
		return oNode->GetIDispatch(TRUE);
	}
	else
	{
		return NULL;
	}
}

IDispatch* odl_PreviousSibling(void)
{
	
	clsNode* oNode = PreviousSibling();
	if (oNode != NULL)
	{
		return oNode->GetIDispatch(TRUE);
	}
	else
	{
		return NULL;
	}
}

BOOL odl_IsFirstSibling(void)
{
	
	return IsFirstSibling();
}

IDispatch* odl_FirstSibling(void)
{
	
	clsNode* oNode = FirstSibling();
	if (oNode != NULL)
	{
		return oNode->GetIDispatch(TRUE);
	}
	else
	{
		return NULL;
	}
}

BOOL odl_IsLastSibling(void)
{
	
	return IsLastSibling();
}

IDispatch* odl_LastSibling(void)
{
	
	clsNode* oNode = LastSibling();
	if (oNode != NULL)
	{
		return oNode->GetIDispatch(TRUE);
	}
	else
	{
		return NULL;
	}
}

IDispatch* odl_Child(void)
{
	
	clsNode* oNode = Child();
	if (oNode != NULL)
	{
		return oNode->GetIDispatch(TRUE);
	}
	else
	{
		return NULL;
	}
	
}

IDispatch* odl_Parent(void)
{
	
	clsNode* oNode = Parent();
	if (oNode != NULL)
	{
		return oNode->GetIDispatch(TRUE);
	}
	else
	{
		return NULL;
	}
}

BOOL odl_IsRoot(void)
{
	
	return IsRoot();
}

BSTR odl_FullPath(void)
{
	
	return FullPath().AllocSysString();
}

LONG odl_Children(void)
{
	
	return Children();
}

BSTR odl_GetXML(void)
{
	
	return GetXML().AllocSysString();
}

void odl_SetXML(LPCTSTR sXML)
{
	
	SetXML(sXML);
}


};