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


using namespace Gdiplus;

class clsTask;
class clsRow;
class clsStyle;

class clsGraphics  
{
public:

	CActiveGanttVCCtl* mp_oControl;

	clsGraphics(CActiveGanttVCCtl* oControl);
	virtual ~clsGraphics();

	CPtrArray mp_audtActiveReversibleFrames;
	CPtrArray mp_audtActiveReversibleLines;
	BOOL mp_bCustomPrinting;
	Graphics* mp_lCustomDC;
	LONG mp_lPWidth;
	LONG mp_lPHeight;
	LONG mp_lFocusLeft;
	LONG mp_lFocusTop;
	LONG mp_lFocusRight;
	LONG mp_lFocusBottom;
	BOOL mp_bEnableClipRegions;
	BOOL bToolTipGraphics;
	RECTD mp_oTextFinalLayout;

	LONG mp_lX1;
	LONG mp_lY1;
	LONG mp_lX2;
	LONG mp_lY2;

	RECT mp_udtPrintingClipRegion;
	RECT mp_udtFocusRect;
	RECT mp_udtPreviousClipRegion;

	LONG mp_lBackupDC;
	float mp_fFontRatio;

	LONG Getf_FocusLeft(void);
	void Setf_FocusLeft(LONG Value);
	LONG Getf_FocusTop(void);
	void Setf_FocusTop(LONG Value);
	LONG Getf_FocusRight(void);
	void Setf_FocusRight(LONG Value);
	LONG Getf_FocusBottom(void);
	void Setf_FocusBottom(LONG Value);


	Graphics* oGraphics();
	void SetCustomPrinting(BOOL nValue);
	BOOL GetCustomPrinting();
	LONG Width();
	LONG Height();
	void SetCustomWidth(LONG nValue);
	LONG GetCustomWidth();
	void SetCustomHeight(LONG nValue);
	LONG GetCustomHeight();
	void mp_DrawPolygon(Color clrColor, Point* r_oPoints, LONG lLength);
	void mp_DrawEdge(LONG X1, LONG Y1, LONG X2, LONG Y2, Color clrBackColor, GRE_BUTTONSTYLE yButtonStyle, GRE_EDGETYPE yEdgeType, BOOL bFilled, clsStyle* oStyle);
	void mp_DrawLine(LONG X1, LONG Y1, LONG X2, LONG Y2, GRE_LINETYPE yStyle, Color clrColor, GRE_LINEDRAWSTYLE yDrawStyle, LONG lWidth = 1, BOOL bCreatePens = TRUE);
	void mp_DrawFigure(LONG X, LONG Y, LONG lDx, LONG lDy, GRE_FIGURETYPE yFigureType, Color lBorderColor, Color lFillColor, GRE_LINEDRAWSTYLE BorderStyle);
	void mp_DrawFigureAux(Brush* oBrush, Pen* oPen, Point* oPoints, int Count);
	void mp_DrawPattern(LONG X1, LONG Y1, LONG X2, LONG Y2, Color clrColor, GRE_PATTERN yDrawStyle, LONG lPatternFactor);
	void mp_DrawPolyLine(Color clrColor, LONG lWidth, GRE_LINEDRAWSTYLE yDrawStyle, Point *r_oPoints, LONG lLength);
	void mp_DrawTextEx(LONG X1, LONG Y1, LONG X2, LONG Y2, CString sParam, StringFormat* yFlags, Color clrColor, Font* oFont, BOOL bClip = TRUE);
	void mp_DrawAlignedText(LONG X1, LONG Y1, LONG X2, LONG Y2, CString sParam, GRE_HORIZONTALALIGNMENT yHPos, GRE_VERTICALALIGNMENT yVPos, Color clrColor, clsFont* oFont, BOOL bClip = TRUE);
	void mp_ClipRegion(LONG X1, LONG Y1, LONG X2, LONG Y2, BOOL bStore);
	void mp_RestorePreviousClipRegion();
	void mp_ClearClipRegion();
	void mp_TileImageHorizontal(clsImage* oImageHandle, LONG X1, LONG Y1, LONG X2, LONG Y2, BOOL bTransparent);
	void mp_PaintImage(clsImage* oImageHandle, LONG X1, LONG Y1, LONG X2, LONG Y2, LONG xOrigin, LONG yOrigin, BOOL bUseMask);
	void mp_DrawImage(clsImage* oImage, GRE_HORIZONTALALIGNMENT yHorizontalAlignment, GRE_VERTICALALIGNMENT yVerticalAlignment, LONG lImageXMargin, LONG lImageYMargin, LONG X1, LONG X2, LONG Y1, LONG Y2, BOOL bTransparent);
	void mp_DrawFocusRectangle(LONG X1, LONG Y1, LONG X2, LONG Y2);
	void mp_GradientFill(LONG X1, LONG Y1, LONG X2, LONG Y2, Color clrStartColor, Color clrEndColor, GRE_GRADIENTFILLMODE yGradientType);
	void mp_HatchFill(LONG X1, LONG Y1, LONG X2, LONG Y2, Color clrForeColor, Color clrBackColor, GRE_HATCHSTYLE yHatchStyle);
	void mp_ResetFocusRectangle();
	void mp_DrawReversibleFrameEx();
	void mp_DrawReversibleLine(LONG X1, LONG Y1, LONG X2, LONG Y2);
	void mp_EraseReversibleLines();
	void mp_DrawReversibleFrame(LONG X1, LONG Y1, LONG X2, LONG Y2);
	void mp_EraseReversibleFrames();
	Color mp_ConvertColor(OLE_COLOR dwOleColour);
	BOOL mp_LineIntersection(LONG X1, LONG Y1, LONG X2, LONG Y2);
	void mp_DrawItemI(LONG X1, LONG Y1, LONG X2, LONG Y2, CString sStyleIndex, CString sText, BOOL bSelected, clsImage* oImage, LONG lLeftTrim, LONG lRightTrim, clsStyle* v_oStyle);
	void mp_DrawItemI(clsTask* oTask, CString sStyleIndex, BOOL Selected, clsStyle* v_oStyle);
	void mp_DrawItemEx(LONG X1, LONG Y1, LONG X2, LONG Y2, CString sText, BOOL bIsSelected, clsImage* oImage, LONG lLeftTrim, LONG lRightTrim, clsStyle* v_oStyle, Color clrBackColor, Color clrForeColor, Color clrStartGradientColor, Color clrEndGradientColor, Color clrHatchBackColor, Color clrHatchForeColor);
	void mp_DrawItemRow(LONG X1, LONG Y1, LONG X2, LONG Y2, CString sText, BOOL bIsSelected, clsImage* oImage, LONG lLeftTrim, LONG lRightTrim, clsStyle* v_oStyle, clsRow* oRow);
	void mp_DrawItem(LONG X1, LONG Y1, LONG X2, LONG Y2, CString sStyleIndex, CString sText, BOOL bIsSelected, clsImage* oImage, LONG lLeftTrim, LONG lRightTrim, clsStyle* v_oStyle);
	void mp_DrawSelectionRectangle(LONG X1, LONG Y1, LONG X2, LONG Y2, clsStyle* oStyle);
	void mp_DrawItemText(LONG X1, LONG Y1, LONG X2, LONG Y2, LONG lLeftTrim, LONG lRightTrim, clsStyle* oStyle, CString sText);
	BOOL mp_bPositionItem(LONG* r_lTop, LONG* r_lBottom, clsStyle* v_oStyle);
	void mp_DrawButton(Rect oRect, E_SCROLLBUTTONSTATE state);
	void mp_DrawScrollButton(LONG X1, LONG Y1, LONG width, LONG height, E_SCROLLBUTTON button, E_SCROLLBUTTONSTATE state);
	HBITMAP mp_GetBitmapHandle(LPDISPATCH oImageList, LONG lIndex);
	LONG mp_GetImageListCount(LPDISPATCH oImageList);
	Color mp_ConvGDICol(OLE_COLOR lColor);
	void SetCustomDC(Graphics* oGraphics);

	//void DrawPoint(LONG X, LONG Y, Color clrColor);

	void mp_DrawPoint(int X, int Y, Gdiplus::Color clrColor);
	void mp_DrawArrow(int X, int Y, GRE_ARROWDIRECTION yArrowDirection, int lArrowSize, Gdiplus::Color clrColor);

	LONG mp_lStrWidth(CString sString,clsFont* r_oFont);
	LONG mp_lStrWidth(CString sString, Font* r_oFont);
	LONG mp_lStrHeight(CString sString,clsFont* r_oFont);
	LONG mp_lStrHeight(CString sString, Font* r_oFont);

	void mp_DrawExpandIcon(LONG X, LONG Y);
	void mp_DrawCollapseIcon(LONG X, LONG Y);

	void mp_DrawExpandIcon(LONG X, LONG Y, Gdiplus::Color oColor, Gdiplus::Color oDropShadowColor, Gdiplus::Color oBackgroundColor);
	void mp_DrawCollapseIcon(LONG X, LONG Y, Gdiplus::Color oColor, Gdiplus::Color oDropShadowColor);

	void oGraphics_DrawLine(Gdiplus::Color oColor, LONG X1, LONG Y1, LONG X2, LONG Y2);
	void oGraphics_FillRectangle(Gdiplus::Color oColor, LONG X, LONG Y, LONG Width, LONG Height);


};