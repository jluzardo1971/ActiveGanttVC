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
#include "GlobalFunctions.h"
#include "NumericFormat.h"

int GetEncoderClsid(const WCHAR* format, CLSID* pClsid)
{
   UINT  num = 0;          // number of image encoders
   UINT  size = 0;         // size of the image encoder array in bytes

   ImageCodecInfo* pImageCodecInfo = NULL;

   GetImageEncodersSize(&num, &size);
   if(size == 0)
      return -1;  // Failure

   pImageCodecInfo = (ImageCodecInfo*)(malloc(size));
   if(pImageCodecInfo == NULL)
      return -1;  // Failure

   GetImageEncoders(num, size, pImageCodecInfo);

   for(UINT j = 0; j < num; ++j)
   {
      if( wcscmp(pImageCodecInfo[j].MimeType, format) == 0 )
      {
         *pClsid = pImageCodecInfo[j].Clsid;
         free(pImageCodecInfo);
         return j;  // Success
      }    
   }

   free(pImageCodecInfo);
   return -1;  // Failure
}

VARIANT F_NULL_VARIANT()
{
	VARIANT vReturn;
	VariantInit(&vReturn);
	vReturn.vt = VT_ERROR;
	return vReturn;
}

COleDateTime F_COleDateTime(DATE dtParam)
{
	COleDateTime dtReturn = dtParam;
	return dtReturn;
}

VARIANT F_BSTR_VARIANT(CString sParam)
{
	VARIANT sReturn;
	VariantInit(&sReturn);
	sReturn.vt = VT_BSTR;
	sReturn.bstrVal = sParam.AllocSysString();
	return sReturn;
}

HBITMAP F_HBITMAP_CPictureHolder(CPictureHolder* oParam)
{
	OLE_HANDLE hPictureBuff;
	oParam->m_pPict->get_Handle(&hPictureBuff);
	return (HBITMAP) hPictureBuff;
}

int F_Width_CPictureHolder(CPictureHolder* oParam)
{

	HBITMAP hPicture = F_HBITMAP_CPictureHolder(oParam);
	BITMAP rPicture;
	GetObject(hPicture, sizeof(rPicture), &rPicture);
	return rPicture.bmWidth;
}

int F_Height_CPictureHolder(CPictureHolder* oParam)
{

	HBITMAP hPicture = F_HBITMAP_CPictureHolder(oParam);
	BITMAP rPicture;
	GetObject(hPicture, sizeof(rPicture), &rPicture);
	return rPicture.bmHeight;
}

BOOL IsCPictureHolderEmpty(CPictureHolder* oParam)
{
	HBITMAP hPicture = F_HBITMAP_CPictureHolder(oParam);
	if (hPicture == 0)
	{
		return TRUE;
	}
	else
	{
		return FALSE;
	}
}

BOOL g_TO_BOOL(BOOL bVariant)
{
	if (bVariant == FALSE)
	{
		return FALSE;
	}
	else
	{
		return TRUE;
	}
}

CString F_TOUCSTR(LPCSTR sParam)
{
	int n = ::MultiByteToWideChar(CP_UTF8, 0, sParam, -1, NULL, 0);
	CString sReturn;
	LPWSTR p = sReturn.GetBuffer(n);
	::MultiByteToWideChar(CP_UTF8, 0, sParam, -1, p, n);
	sReturn.ReleaseBuffer();
	return sReturn;
}

clsFont* g_GetFontFromIDispatch(IDispatch* dispFont)
{
	ActiveGanttVC::CclsFont dFont(dispFont);
	clsFont* oFont = new clsFont();

	CString sFontName = dFont.GetFontName();
	oFont->SetFontName(sFontName);

	FLOAT fSize = dFont.GetSize();
	oFont->SetSize(fSize);

	LONG lStyle = dFont.GetStyle();
	oFont->SetStyle(lStyle);

	LONG lUnit = dFont.GetUnit();
	oFont->SetUnit(lUnit);
	dFont.DetachDispatch();

	return oFont;
}

clsImage* g_GetImageFromIDispatch(IDispatch* dispImage)
{
	ActiveGanttVC::CclsImage dImage(dispImage);
	clsImage* oImage = new clsImage();

	CString sFilename = dImage.GetFilename();
	BOOL bEmbeddedColorManagement = dImage.GetEmbeddedColorManagement();

	oImage->FromFile(sFilename, bEmbeddedColorManagement);

	dImage.DetachDispatch();
	return oImage;
}


clsTextFlags* g_GetTextFlagsFromIDispatch(IDispatch* dispTextFlags)
{
	ActiveGanttVC::CclsTextFlags dTextFlags(dispTextFlags);
	clsTextFlags* oTextFlags = new clsTextFlags();

	LONG lVerticalAlignment = dTextFlags.GetVerticalAlignment();
	oTextFlags->SetVerticalAlignment(lVerticalAlignment);

	LONG lHorizontalAlignment = dTextFlags.GetHorizontalAlignment();
	oTextFlags->SetHorizontalAlignment(lHorizontalAlignment);

	BOOL bWordWrap = dTextFlags.GetWordWrap();
	oTextFlags->SetWordWrap(bWordWrap);

	BOOL bRightToLeft = dTextFlags.GetRightToLeft();
	oTextFlags->SetRightToLeft(bRightToLeft);

	LONG lOffsetBottom = dTextFlags.GetOffsetBottom();
	oTextFlags->SetOffsetBottom(lOffsetBottom);

	LONG lOffsetLeft = dTextFlags.GetOffsetLeft();
	oTextFlags->SetOffsetLeft(lOffsetLeft);

	LONG lOffsetRight = dTextFlags.GetOffsetRight();
	oTextFlags->SetOffsetRight(lOffsetRight);

	LONG lOffsetTop = dTextFlags.GetOffsetTop();
	oTextFlags->SetOffsetTop(lOffsetTop);

	dTextFlags.DetachDispatch();
	return oTextFlags;
}

DECIMAL g_InitDec(LONG nNewValue)
{
	DECIMAL cReturn;
	VarDecFromI4(nNewValue, &cReturn);
	return cReturn;
}

HRESULT g_VarDecCmpEx(LPDECIMAL  pdecLeft, LONG nValue)
{
	DECIMAL cDecRight;
	VarDecFromI4(nValue, &cDecRight);
	return VarDecCmp(pdecLeft, &cDecRight);
}

CString CStr(int iValue)
{
	CString sReturn;
	sReturn.Format(_T("%d"), iValue);
	return sReturn;
}

CString CStr(unsigned int iValue)
{
	CString sReturn;
	sReturn.Format(_T("%d"), iValue);
	return sReturn;
}

CString CStr(float fValue)
{
	CString sReturn;
	sReturn.Format(_T("%f"), fValue);
	return sReturn;
}

LONG CInt32(CString sParam)
{
	LONG lReturn;
	lReturn = _wtoi(sParam);
	return lReturn;
}

LONG CInt32(DOUBLE dParam)
{
	LONG lReturn;
	if (dParam < 0.0)
	{
		lReturn = (LONG)ceil(dParam - 0.5);
	}
	else
	{
		lReturn = (LONG)floor(dParam + 0.51);
	}
	return lReturn;
}

double CDbl(int lValue)
{
	return (double)lValue;
}

Gdiplus::REAL CReal(LONG lParam)
{
	return (Gdiplus::REAL)lParam;
}

Gdiplus::REAL CReal(FLOAT fParam)
{
	return (Gdiplus::REAL)fParam;
}

COLORREF g_TranslateColor(OLE_COLOR oColor)
{
	COLORREF oReturn;
	OleTranslateColor(oColor, 0, &oReturn);
	return oReturn;
}

BOOL g_IsNumeric(CString Expression)
{
	if (Expression.GetLength() == 0)
	{
		return FALSE;
	}
	CString sBase = _T("0123456789");
	int i;
	for (i = 0; i <= Expression.GetLength() - 1; i++)
	{
		CString sChar = Expression.GetAt(i);
		if (sBase.Find(sChar, 0) == -1)
		{
			return FALSE;
		}
	}
	return TRUE;
}

CString g_Trim(CString Expression)
{
	Expression.TrimLeft();
	Expression.TrimRight();
	return Expression;
}

CString g_Replace(CString Expression,CString String1,CString String2)
{
	Expression.Replace(String1, String2);
	return Expression;
}

CString g_GetDecimalSeparator()
{
	return GetLocaleInfoEx(LOCALE_SDECIMAL);
}

CString g_Format(FLOAT Expression, CString sFormat)
{
	xAuxCNumericFormat oFormat;
	CString sBuff;
	if (!sFormat.IsEmpty())
	{
		oFormat.SetFormat(sFormat);
		sBuff = oFormat.SPrintNumber(Expression, TRUE);
		return sBuff;
	}
	else
	{
		sBuff.Format(_T("%.2f"), Expression);
		return  sBuff;
	}
}

CString g_Format(LONG Expression, CString sFormat)
{
	return g_Format((FLOAT)Expression, sFormat);
}