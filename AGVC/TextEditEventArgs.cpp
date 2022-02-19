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
#include "TextEditEventArgs.h"


// TextEditEventArgs

IMPLEMENT_DYNCREATE(TextEditEventArgs, CCmdTarget)


TextEditEventArgs::TextEditEventArgs()
{
	EnableAutomation();
	AfxOleLockApp();
	Clear();
}

TextEditEventArgs::~TextEditEventArgs()
{
	AfxOleUnlockApp();
}


void TextEditEventArgs::OnFinalRelease()
{
	// When the last reference for an automation object is released
	// OnFinalRelease is called.  The base class will automatically
	// deletes the object.  Add additional cleanup required for your
	// object before calling the base class.

	CCmdTarget::OnFinalRelease();
}


BEGIN_MESSAGE_MAP(TextEditEventArgs, CCmdTarget)
END_MESSAGE_MAP()


BEGIN_DISPATCH_MAP(TextEditEventArgs, CCmdTarget)
	DISP_PROPERTY_EX_ID(TextEditEventArgs, "ObjectType", 1, odl_GetObjectType, odl_SetObjectType, VT_I4) 
	DISP_PROPERTY_EX_ID(TextEditEventArgs, "ObjectIndex", 2, odl_GetObjectIndex, odl_SetObjectIndex, VT_I4) 
	DISP_PROPERTY_EX_ID(TextEditEventArgs, "ParentObjectIndex", 3, odl_GetParentObjectIndex, odl_SetParentObjectIndex, VT_I4)
	DISP_PROPERTY_EX_ID(TextEditEventArgs, "Text", 4, odl_GetText, odl_SetText, VT_BSTR) 
END_DISPATCH_MAP()

// Note: we add support for IID_ITextEditEventArgs to support typesafe binding
//  from VBA.  This IID must match the GUID that is attached to the 
//  dispinterface in the .IDL file.

//{97259A41-067C-423D-B081-CFB9F17E4D74}
static const IID IID_ITextEditEventArgs = { 0x97259A41, 0x067C, 0x423D, { 0xB0, 0x81, 0xCF, 0xB9, 0xF1, 0x7E, 0x4D, 0x74} };


BEGIN_INTERFACE_MAP(TextEditEventArgs, CCmdTarget)
	INTERFACE_PART(TextEditEventArgs, IID_ITextEditEventArgs, Dispatch)
END_INTERFACE_MAP()

//{DE9924E5-C09C-4EA0-8FD6-BB0E9D5C3B92}
IMPLEMENT_OLECREATE_FLAGS(TextEditEventArgs, "AGVC.TextEditEventArgs", afxRegApartmentThreading, 0xDE9924E5, 0xC09C, 0x4EA0, 0x8F, 0xD6, 0xBB, 0x0E, 0x9D, 0x5C, 0x3B, 0x92)

LONG TextEditEventArgs::GetObjectType(void)
{
	return mp_yObjectType;
}

void TextEditEventArgs::SetObjectType(LONG Value)
{
	mp_yObjectType = Value;
}

LONG TextEditEventArgs::GetObjectIndex(void)
{
	return mp_lObjectIndex;
}

void TextEditEventArgs::SetObjectIndex(LONG Value)
{
	mp_lObjectIndex = Value;
}

LONG TextEditEventArgs::GetParentObjectIndex(void)
{
	return mp_lParentObjectIndex;
}

void TextEditEventArgs::SetParentObjectIndex(LONG Value)
{
	mp_lParentObjectIndex = Value;
}

void TextEditEventArgs::Clear(void)
{
	mp_yObjectType = 0;
	mp_lObjectIndex = 0;
	mp_lParentObjectIndex = 0;
	mp_sText = "";
}

CString TextEditEventArgs::GetText(void)
{
	return mp_sText;
}

void TextEditEventArgs::SetText(CString Value)
{
	mp_sText = Value;
}