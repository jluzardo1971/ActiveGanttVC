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
#include "clsImage.h"


// clsImage

IMPLEMENT_DYNCREATE(clsImage, CCmdTarget)

clsImage::clsImage()
{
	EnableAutomation();
	AfxOleLockApp();
	mp_oImage = NULL;
	mp_sFilename = "";
	mp_bEmbeddedColorManagement = FALSE;
}

void clsImage::OnFinalRelease()
{
	CCmdTarget::OnFinalRelease();
}

clsImage::~clsImage()
{

	delete mp_oImage;
	mp_oImage = NULL;
	AfxOleUnlockApp();
}


clsImage::clsImage(const clsImage& oImage)
{
	EnableAutomation();
	AfxOleLockApp();
	mp_oImage = oImage.mp_oImage;
	mp_sFilename = oImage.mp_sFilename;
	mp_bEmbeddedColorManagement = oImage.mp_bEmbeddedColorManagement;
}

clsImage& clsImage::operator=(const clsImage& oImage)
{
	if (this != &oImage)
	{

		delete mp_oImage;
		mp_oImage = NULL;

		mp_oImage = oImage.mp_oImage;
		mp_sFilename = oImage.mp_sFilename;
		mp_bEmbeddedColorManagement = oImage.mp_bEmbeddedColorManagement;
	}
	return *this;
}

BEGIN_MESSAGE_MAP(clsImage, CCmdTarget)
END_MESSAGE_MAP()


BEGIN_DISPATCH_MAP(clsImage, CCmdTarget)
	DISP_FUNCTION_ID(clsImage, "FromFile", 1, FromFile, VT_EMPTY, VTS_BSTR VTS_BOOL)
	DISP_PROPERTY_EX_ID(clsImage, "Width", 2, GetWidth, SetNotSupported, VT_I4)
	DISP_PROPERTY_EX_ID(clsImage, "Height", 3, GetHeight, SetNotSupported, VT_I4)
	DISP_FUNCTION_ID(clsImage, "Clear", 4, Clear, VT_EMPTY, VTS_NONE)
	DISP_FUNCTION_ID(clsImage, "hImage", 5, hImage, VT_I4, VTS_NONE)
	DISP_PROPERTY_EX_ID(clsImage, "Filename", 6, GetFilename, SetNotSupported, VT_BSTR)
	DISP_PROPERTY_EX_ID(clsImage, "EmbeddedColorManagement", 7, GetEmbeddedColorManagement, SetNotSupported, VT_BOOL)
END_DISPATCH_MAP()

// {00AD05DB-9D30-40A4-9B01-90B12B651348}
static const IID IID_IclsImage = { 0x00AD05DB, 0x9D30, 0x40A4, { 0x9B, 0x01, 0x90, 0xB1, 0x2B, 0x65, 0x13, 0x48} };

BEGIN_INTERFACE_MAP(clsImage, CCmdTarget)
	INTERFACE_PART(clsImage, IID_IclsImage, Dispatch)
END_INTERFACE_MAP()

// {25A3A87E-6063-4D12-959A-4209A5C16293}
IMPLEMENT_OLECREATE_FLAGS(clsImage, "AGVC.clsImage", afxRegApartmentThreading, 0x25A3A87E, 0x6063, 0x4D12, 0x95, 0x9A, 0x42, 0x09, 0xA5, 0xC1, 0x62, 0x93)


// clsImage message handlers

void clsImage::FromFile(LPCTSTR Filename, BOOL UseEmbeddedColorManagement)
{
	mp_sFilename = Filename;

	delete mp_oImage;
	mp_oImage = NULL;

	mp_oImage = Image::FromFile(Filename, UseEmbeddedColorManagement);
}

LONG clsImage::GetWidth(void)
{
	
	return mp_oImage->GetWidth();
}

LONG clsImage::GetHeight(void)
{
	
	return mp_oImage->GetHeight();
}

Gdiplus::Image* clsImage::GetImageP()
{
	return mp_oImage;
}

void clsImage::Clear()
{

	delete mp_oImage;
	mp_oImage = NULL;

}

LONG clsImage::hImage(void)
{
	return (LONG)mp_oImage;
}

BSTR clsImage::GetFilename(void)
{
	return mp_sFilename.AllocSysString();
}

BOOL clsImage::GetEmbeddedColorManagement(void)
{
	return mp_bEmbeddedColorManagement;
}

void clsImage::Clone(clsImage* oClone)
{

	delete oClone->mp_oImage;
	oClone->mp_oImage = NULL;

	if (mp_oImage != NULL)
	{
		oClone->mp_oImage = mp_oImage->Clone();
	}
	oClone->mp_sFilename = mp_sFilename;
	oClone->mp_bEmbeddedColorManagement = mp_bEmbeddedColorManagement;
}