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

class clsFont : public CCmdTarget
{
	DECLARE_DYNCREATE(clsFont)

public:
	clsFont();
	virtual ~clsFont();
	virtual void OnFinalRelease();
	clsFont(const clsFont& oFont);
	clsFont& operator=(const clsFont& oFont);
	BOOL GetBold();
	BOOL GetItalic();
	BOOL GetUnderline();
	void Clone(clsFont* oClone);

protected:
	Gdiplus::Font* mp_oFont;
	Gdiplus::REAL mp_dEmSize;
	CString mp_sFontName;
	LONG mp_lStyle;
	LONG mp_lUnit;
public:
	Gdiplus::Font* GetFontP();
	CString GetFontName();
	void SetFontName(CString sFontName);
	FLOAT GetSize();
	void SetSize(FLOAT dSize);
	void SetPtFont(CString sFontName, FLOAT Size, LONG Style);

protected:
	DECLARE_MESSAGE_MAP()
	DECLARE_OLECREATE(clsFont)
	DECLARE_DISPATCH_MAP()
	DECLARE_INTERFACE_MAP()
public:

	LONG GetStyle(void);
	void SetStyle(LONG Value);
	LONG GetUnit(void);
	void SetUnit(LONG Value);


BSTR odl_GetFontName(void)
{
	
	return GetFontName().AllocSysString();
}

void odl_SetFontName(LPCTSTR Value)
{
	
	SetFontName(Value);
}

FLOAT odl_GetSize(void)
{
	
	return GetSize();
}

void odl_SetSize(FLOAT Value)
{
	
	SetSize(Value);
}

};