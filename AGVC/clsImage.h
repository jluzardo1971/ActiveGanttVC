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

class clsImage : public CCmdTarget
{
	DECLARE_DYNCREATE(clsImage)

public:
	clsImage();
	virtual ~clsImage();
	virtual void OnFinalRelease();
	clsImage(const clsImage& oImage);
	clsImage& clsImage::operator=(const clsImage& oImage);

	CString mp_sFilename;
	Gdiplus::Image* mp_oImage;
	BOOL mp_bEmbeddedColorManagement;

	LONG GetWidth(void);
	LONG GetHeight(void);
	Gdiplus::Image* GetImageP();
	void Clear(void);
	LONG hImage(void);
	BSTR GetFilename(void);
	BOOL GetEmbeddedColorManagement(void);
	void FromFile(LPCTSTR Filename, BOOL UseEmbeddedColorManagement);
	void Clone(clsImage* oClone);


protected:
	DECLARE_MESSAGE_MAP()
	DECLARE_OLECREATE(clsImage)
	DECLARE_DISPATCH_MAP()
	DECLARE_INTERFACE_MAP()

};