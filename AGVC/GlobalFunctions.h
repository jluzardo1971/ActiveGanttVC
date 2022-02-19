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

#include <gdiplus.h>



    struct S_Rectangle
    {
        int X1;
        int Y1;
        int X2;
        int Y2;

        BOOL mp_bInRect(int X, int Y)
        {
            if (X >= X1 && X <= X2 && Y >= Y1 && Y <= Y2)
            {
                return TRUE;
            }
            else
            {
                return FALSE;
            }
        }

    };

    struct S_Point
    {
        int X;
        int Y;
    };

    struct S_TIMEBLOCK
    {
        COleDateTime dtStart;
        COleDateTime dtEnd;
    };

class clsFont;
class clsImage;
class clsTextFlags;
class CActiveGanttVCCtl;

int GetEncoderClsid(const WCHAR* format, CLSID* pClsid);
VARIANT F_BSTR_VARIANT(CString sParam);
COleDateTime F_COleDateTime(DATE dtParam);
VARIANT F_NULL_VARIANT();
HBITMAP F_HBITMAP_CPictureHolder(CPictureHolder* oParam);
int F_Width_CPictureHolder(CPictureHolder* oParam);
int F_Height_CPictureHolder(CPictureHolder* oParam);
BOOL IsCPictureHolderEmpty(CPictureHolder* oParam);
BOOL g_TO_BOOL(BOOL bVariant);
CString F_TOUCSTR(LPCSTR sParam);
clsFont* g_GetFontFromIDispatch(IDispatch* dispFont);
clsImage* g_GetImageFromIDispatch(IDispatch* dispImage);
clsTextFlags* g_GetTextFlagsFromIDispatch(IDispatch* dispTextFlags);


DECIMAL g_InitDec(LONG nNewValue);
HRESULT g_VarDecCmpEx(LPDECIMAL  pdecLeft, LONG nValue);

CString CStr(int iValue);
CString CStr(float fValue);
CString CStr(unsigned int iValue);
LONG CInt32(CString sParam);
LONG CInt32(DOUBLE dParam);
Gdiplus::REAL CReal(LONG lParam);
Gdiplus::REAL CReal(FLOAT fParam);
double CDbl(int lValue);

COLORREF g_TranslateColor(OLE_COLOR oColor);

BOOL g_IsNumeric(CString Expression);
CString g_Trim(CString Expression);
CString g_Replace(CString Expression,CString String1,CString String2);
CString g_GetDecimalSeparator();

CString g_Format(FLOAT Expression, CString sFormat);
CString g_Format(LONG Expression, CString sFormat);