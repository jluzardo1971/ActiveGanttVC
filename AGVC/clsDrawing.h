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


class clsGDIGraphics;
class clsStringFormat;
class clsFont;
class clsImage;

class clsDrawing : public CCmdTarget
{
	DECLARE_DYNCREATE(clsDrawing)
	clsDrawing();

public:
	CActiveGanttVCCtl* mp_oControl;

	clsDrawing(CActiveGanttVCCtl* oControl);
	virtual ~clsDrawing();
	virtual void OnFinalRelease();

// Member Variables
public:
	clsGDIGraphics mp_oGraphics;
	clsTextFlags* mp_oTextFlags;
	clsImage* mp_oImage;
	clsFont mp_oFont;

//Internal Properties (Extensions to ODL)
public:

//Internal Properties
public:


//Internal Methods
public:
	clsGDIGraphics* GraphicsInfo(void);
	void DrawLine(LONG X1,LONG Y1,LONG X2,LONG Y2,Gdiplus::Color LineColor,LONG LineStyle,LONG LineWidth);
	void DrawBorder(LONG X1,LONG Y1,LONG X2,LONG Y2,Gdiplus::Color LineColor,LONG LineStyle,LONG LineWidth);
	void DrawRectangle(LONG X1,LONG Y1,LONG X2,LONG Y2,Gdiplus::Color LineColor,LONG LineStyle,LONG LineWidth);
	void DrawText(LONG X1,LONG Y1,LONG X2,LONG Y2,CString Text,StringFormat* oFlags, Gdiplus::Color TextColor,clsFont* TextFont);
	void DrawAlignedText(LONG X1,LONG Y1,LONG X2,LONG Y2,CString Text,LONG HorizontalAlignment,LONG VerticalAlignment,Gdiplus::Color TextColor,clsFont* TextFont);
	void PaintImage(clsImage* Image,LONG X1,LONG Y1,LONG X2,LONG Y2);
	clsTextFlags* GetTextFlags(void);
	clsImage* GetImage(void);
	clsFont* clsDrawing::GetFont(void);
	void DrawObject(LONG X1, LONG Y1, LONG X2, LONG Y2, CString StyleIndex, CString Text, BOOL Selected, clsImage* Image, LONG ObjectType);

protected:
	DECLARE_DISPATCH_MAP()
	DECLARE_MESSAGE_MAP()
	DECLARE_OLECREATE(clsDrawing)
	DECLARE_INTERFACE_MAP()


	IDispatch* odl_GraphicsInfo(void)
	{
		
		return GraphicsInfo()->GetIDispatch(TRUE);
	}

	void odl_DrawLine(LONG X1,LONG Y1,LONG X2,LONG Y2,OLE_COLOR LineColor,LONG LineStyle,LONG LineWidth)
	{
		
		Gdiplus::Color oColor;
		oColor.SetFromCOLORREF(g_TranslateColor(LineColor));
		DrawLine(X1,Y1,X2,Y2,oColor,LineStyle,LineWidth);
	}

	void odl_DrawBorder(LONG X1,LONG Y1,LONG X2,LONG Y2,OLE_COLOR LineColor,LONG LineStyle,LONG LineWidth)
	{
		
		Gdiplus::Color oColor;
		oColor.SetFromCOLORREF(g_TranslateColor(LineColor));
		DrawBorder(X1,Y1,X2,Y2,oColor,LineStyle,LineWidth);
	}

	void odl_DrawRectangle(LONG X1,LONG Y1,LONG X2,LONG Y2,OLE_COLOR LineColor,LONG LineStyle,LONG LineWidth)
	{
		
		Gdiplus::Color oColor;
		oColor.SetFromCOLORREF(g_TranslateColor(LineColor));
		DrawRectangle(X1,Y1,X2,Y2,oColor,LineStyle,LineWidth);
	}

	void odl_DrawText(LONG X1,LONG Y1,LONG X2,LONG Y2,LPCTSTR Text,OLE_COLOR TextColor,IDispatch* TextFont)
	{
		
		Gdiplus::Color oColor;
		oColor.SetFromCOLORREF(g_TranslateColor(TextColor));
		clsFont* oFont = g_GetFontFromIDispatch(TextFont);
		//clsTextFlags oTextFlags = g_GetTextFlagsFromIDispatch(TextFlags);
		Gdiplus::StringFormat oFlags;
		oFlags.SetLineAlignment((Gdiplus::StringAlignment)(mp_oTextFlags->GetVerticalAlignment() - 1));
		oFlags.SetAlignment((Gdiplus::StringAlignment)(mp_oTextFlags->GetHorizontalAlignment() -1));
		if (mp_oTextFlags->GetWordWrap() == FALSE)
		{
			oFlags.SetFormatFlags(oFlags.GetFormatFlags() | StringFormatFlagsNoWrap);
		}
		if (mp_oTextFlags->GetRightToLeft() == TRUE)
		{
			oFlags.SetFormatFlags(oFlags.GetFormatFlags() | StringFormatFlagsDirectionRightToLeft);
		}
		DrawText(X1 + mp_oTextFlags->GetOffsetLeft(),Y1 + mp_oTextFlags->GetOffsetTop(),X2 + mp_oTextFlags->GetOffsetRight(),Y2 +  mp_oTextFlags->GetOffsetBottom(),Text,&oFlags,oColor, oFont);
		delete oFont;
		oFont = NULL;
	}

	void odl_DrawAlignedText(LONG X1,LONG Y1,LONG X2,LONG Y2,LPCTSTR Text,LONG HorizontalAlignment,LONG VerticalAlignment,OLE_COLOR TextColor,IDispatch* TextFont)
	{
		
		Gdiplus::Color oColor;
		oColor.SetFromCOLORREF(g_TranslateColor(TextColor));
		clsFont* oFont = g_GetFontFromIDispatch(TextFont);
		DrawAlignedText(X1,Y1,X2,Y2,Text,HorizontalAlignment,VerticalAlignment,oColor, oFont);
		delete oFont;
		oFont = NULL;
	}

	void odl_PaintImage(IDispatch* Image,LONG X1,LONG Y1,LONG X2,LONG Y2)
	{
		
		clsImage* oImage = g_GetImageFromIDispatch(Image);
		PaintImage(oImage,X1,Y1,X2,Y2);
		delete oImage;
		oImage = NULL;
	}

	IDispatch* odl_GetFont(void)
	{
		
		return GetFont()->GetIDispatch(TRUE);
	}

	IDispatch* odl_GetImage(void)
	{
		
		return GetImage()->GetIDispatch(TRUE);
	}

	IDispatch* odl_GetTextFlags(void)
	{
		
		return GetTextFlags()->GetIDispatch(TRUE);
	}

	void odl_DrawObject(LONG X1, LONG Y1, LONG X2, LONG Y2, LPCTSTR StyleIndex, LPCTSTR Text, BOOL Selected, IDispatch* Image, LONG ObjectType)
	{
		clsImage* oImage = g_GetImageFromIDispatch(Image);
		DrawObject(X1, Y1, X2, Y2, StyleIndex, Text, Selected, oImage, ObjectType);
		delete oImage;
		oImage = NULL;
	}


};