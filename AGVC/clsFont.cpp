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
#include "clsFont.h"

IMPLEMENT_DYNCREATE(clsFont, CCmdTarget)

clsFont::clsFont()
{
	EnableAutomation();
	AfxOleLockApp();
	mp_dEmSize = 8;
	mp_sFontName = _T("Tahoma");
	mp_lStyle = 0;
	mp_lUnit = FS_FONTSTYLEREGULAR;
	mp_oFont = new Gdiplus::Font(mp_sFontName, mp_dEmSize, mp_lStyle,(Gdiplus::Unit)mp_lUnit, NULL);
}

void clsFont::OnFinalRelease()
{
	CCmdTarget::OnFinalRelease();
}

clsFont::~clsFont()
{
	delete mp_oFont;
	mp_oFont = NULL;

	AfxOleUnlockApp();
}

clsFont::clsFont(const clsFont& oFont)
{
	EnableAutomation();
	AfxOleLockApp();
	mp_oFont = oFont.mp_oFont;
	mp_dEmSize = oFont.mp_dEmSize;
	mp_sFontName = oFont.mp_sFontName;
	mp_lStyle = oFont.mp_lStyle;
	mp_lUnit = oFont.mp_lUnit;
}

clsFont& clsFont::operator=(const clsFont& oFont)
{
	if (this != &oFont)
	{
		delete mp_oFont;
		mp_oFont = NULL;

		mp_oFont = oFont.mp_oFont;
		mp_dEmSize = oFont.mp_dEmSize;
		mp_sFontName = oFont.mp_sFontName;
		mp_lStyle = oFont.mp_lStyle;
		mp_lUnit = oFont.mp_lUnit;
	}
	return *this;
}

void clsFont::Clone(clsFont* oClone)
{
	delete oClone->mp_oFont;
	oClone->mp_oFont = NULL;

	oClone->mp_dEmSize = mp_dEmSize;
	oClone->mp_sFontName = mp_sFontName;
	oClone->mp_lStyle = mp_lStyle;
	oClone->mp_lUnit = mp_lUnit;
	oClone->mp_oFont = new Gdiplus::Font(mp_sFontName, mp_dEmSize, mp_lStyle,(Gdiplus::Unit)mp_lUnit, NULL);
}

BEGIN_MESSAGE_MAP(clsFont, CCmdTarget)
END_MESSAGE_MAP()


BEGIN_DISPATCH_MAP(clsFont, CCmdTarget)
	DISP_PROPERTY_EX_ID(clsFont, "FontName", 1, odl_GetFontName, odl_SetFontName, VT_BSTR)
	DISP_PROPERTY_EX_ID(clsFont, "Size", 2, odl_GetSize, odl_SetSize, VT_R4)
	DISP_PROPERTY_EX_ID(clsFont, "Style", 3, GetStyle, SetStyle, VT_I4)
	DISP_PROPERTY_EX_ID(clsFont, "Unit", 4, GetUnit, SetUnit, VT_I4)
END_DISPATCH_MAP()

// {F22CBFF7-1AEC-47E8-8811-2FEAFD9F32A4}
static const IID IID_IclsFont = { 0xF22CBFF7, 0x1AEC, 0x47E8, { 0x88, 0x11, 0x2F, 0xEA, 0xFD, 0x9F, 0x32, 0xA4} };

BEGIN_INTERFACE_MAP(clsFont, CCmdTarget)
	INTERFACE_PART(clsFont, IID_IclsFont, Dispatch)
END_INTERFACE_MAP()

// {0FD2F46A-F54E-4BAA-889E-2AD4AFE20ED8}
IMPLEMENT_OLECREATE_FLAGS(clsFont, "AGVC.clsFont", afxRegApartmentThreading, 0x0FD2F46A, 0xF54E, 0x4BAA, 0x88, 0x9E, 0x2A, 0xD4, 0xAF, 0xE2, 0x0E, 0xD8)

Gdiplus::Font* clsFont::GetFontP()
{
	return mp_oFont;
}

CString clsFont::GetFontName()
{
	Gdiplus::FontFamily oFontFamily;
	WCHAR sFamilyName[LF_FACESIZE];
	CString sReturn;
	mp_oFont->GetFamily(&oFontFamily);
	oFontFamily.GetFamilyName(sFamilyName);
	sReturn = sFamilyName;
	return sReturn;
}

void clsFont::SetFontName(CString sFontName)
{
	mp_sFontName = sFontName;

	delete mp_oFont;
	mp_oFont = NULL;

	mp_oFont = new Gdiplus::Font(mp_sFontName, mp_dEmSize, mp_lStyle,(Gdiplus::Unit)mp_lUnit, NULL);
}

FLOAT clsFont::GetSize()
{
	return mp_dEmSize;
}

void clsFont::SetSize(FLOAT dSize)
{
	mp_dEmSize = dSize;

	delete mp_oFont;
	mp_oFont = NULL;

	mp_oFont = new Gdiplus::Font(mp_sFontName, mp_dEmSize, mp_lStyle,(Gdiplus::Unit)mp_lUnit, NULL);

}

LONG clsFont::GetStyle(void)
{
	
	return mp_lStyle;
}

void clsFont::SetStyle(LONG Value)
{
	
	mp_lStyle = Value;

	delete mp_oFont;
	mp_oFont = NULL;

	mp_oFont = new Gdiplus::Font(mp_sFontName, mp_dEmSize, mp_lStyle,(Gdiplus::Unit)mp_lUnit, NULL);
}

void clsFont::SetPtFont(CString sFontName, FLOAT Size, LONG Style)
{
	mp_sFontName = sFontName;
	mp_dEmSize = Size;
	mp_lStyle = Style;
	mp_lUnit = 3;

	delete mp_oFont;
	mp_oFont = NULL;

	mp_oFont = new Gdiplus::Font(sFontName, Size, Style,(Gdiplus::Unit)mp_lUnit, NULL);
}

LONG clsFont::GetUnit(void)
{
	
	return mp_lUnit;
}

void clsFont::SetUnit(LONG Value)
{
	mp_lUnit = Value;

	delete mp_oFont;
	mp_oFont = NULL;

	mp_oFont = new Gdiplus::Font(mp_sFontName, mp_dEmSize, mp_lStyle,(Gdiplus::Unit)mp_lUnit, NULL);
}

BOOL clsFont::GetBold()
{
	if ((mp_lStyle & FontStyleBold) == FontStyleBold)
	{
		return TRUE;
	}
	else
	{
		return FALSE;
	}
}

BOOL clsFont::GetItalic()
{
	if ((mp_lStyle & FontStyleItalic) == FontStyleItalic)
	{
		return TRUE;
	}
	else
	{
		return FALSE;
	}
}
BOOL clsFont::GetUnderline()
{
	if ((mp_lStyle & FontStyleUnderline) == FontStyleUnderline)
	{
		return TRUE;
	}
	else
	{
		return FALSE;
	}
}