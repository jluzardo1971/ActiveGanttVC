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
#include "clsGraphics.h"

#ifdef _DEBUG
#undef THIS_FILE
static char THIS_FILE[]=__FILE__;
#define new DEBUG_NEW
#endif



clsGraphics::clsGraphics(CActiveGanttVCCtl* oControl)
{
	mp_oControl = oControl;
	mp_audtActiveReversibleFrames.SetSize(0);
	mp_audtActiveReversibleLines.SetSize(0);
	mp_bCustomPrinting = FALSE;
	mp_lFocusLeft = 0;
	mp_lFocusTop = 0;
	mp_lFocusRight = 0;
	mp_lFocusBottom = 0;
	bToolTipGraphics = FALSE;
	mp_lPWidth = 0;
	mp_lPHeight = 0;
	mp_bEnableClipRegions = TRUE;
}

clsGraphics::~clsGraphics()
{

}

Graphics* clsGraphics::oGraphics()
{
	if (mp_bCustomPrinting == FALSE)
	{
		if (bToolTipGraphics == FALSE)
		{
			return mp_oControl->mp_oGraphics;
		}
		else
		{
			return mp_oControl->GetToolTip()->mp_oToolTipGraphics;
		}
		
	}
	else
	{
		return mp_lCustomDC;
	}
}

void clsGraphics::SetCustomPrinting(BOOL nValue)
{
	mp_bCustomPrinting = nValue;
}

BOOL clsGraphics::GetCustomPrinting()
{
	return mp_bCustomPrinting;
}

LONG clsGraphics::Width()
{
	if (mp_bCustomPrinting == FALSE)
	{
		int lHeight;
		int lWidth;
		mp_oControl->GetControlSize(&lWidth,&lHeight);
		return lWidth;
	}
	else
	{
		return mp_lPWidth;
	}
}

LONG clsGraphics::Height()
{
	if (mp_bCustomPrinting == FALSE)
	{
		int lHeight;
		int lWidth;
		mp_oControl->GetControlSize(&lWidth,&lHeight);
		return lHeight;
	}
	else
	{
		return mp_lPHeight;
	}
}

LONG clsGraphics::GetCustomWidth()
{
	return mp_lPWidth;
}

LONG clsGraphics::GetCustomHeight()
{
	return mp_lPHeight;
}

void clsGraphics::SetCustomWidth(LONG nValue)
{
	mp_lPWidth = nValue;
}

void clsGraphics::SetCustomHeight(LONG nValue)
{
	mp_lPHeight = nValue;
}

void clsGraphics::mp_DrawPolygon(Color clrColor, Point *r_oPoints, LONG lLength)
{
	Color oColor(clrColor);
	Pen hPen(oColor, 1);
	oGraphics()->DrawPolygon(&hPen, r_oPoints, lLength);
}

void clsGraphics::mp_DrawEdge(LONG X1, LONG Y1, LONG X2, LONG Y2, Color clrBackColor, GRE_BUTTONSTYLE yButtonStyle, GRE_EDGETYPE yEdgeType, BOOL bFilled, clsStyle* oStyle)
{
	Gdiplus::Color lExteriorLeftTopColor = Color::White;
	Gdiplus::Color lInteriorLeftTopColor = Color::White;
	Gdiplus::Color lExteriorRightBottomColor = Color::White;
	Gdiplus::Color lInteriorRightBottomColor = Color::White;
	if (yButtonStyle == BT_NORMALWINDOWS)
	{
		switch (yEdgeType)
		{
		case ET_RAISED:
			if (oStyle == NULL)
			{
				lExteriorLeftTopColor = Color::MakeARGB(255, 240, 240, 240);
				lInteriorLeftTopColor = Color::MakeARGB(255, 192, 192, 192);
				lInteriorRightBottomColor = Color::Gray;
				lExteriorRightBottomColor = Color::MakeARGB(255, 64, 64, 64);
			}
			else
			{
				lExteriorLeftTopColor = oStyle->GetButtonBorderStyle()->GetRaisedExteriorLeftTopColor();
				lInteriorLeftTopColor = oStyle->GetButtonBorderStyle()->GetRaisedInteriorLeftTopColor();
				lInteriorRightBottomColor = oStyle->GetButtonBorderStyle()->GetRaisedInteriorRightBottomColor();
				lExteriorRightBottomColor = oStyle->GetButtonBorderStyle()->GetRaisedExteriorRightBottomColor();
			}
			break;
		case ET_SUNKEN:
			if (oStyle == NULL)
			{
				lExteriorLeftTopColor = Color::Gray;
				lInteriorLeftTopColor = Color::MakeARGB(255, 64, 64, 64);
				lInteriorRightBottomColor = Color::MakeARGB(255, 192, 192, 192);
				lExteriorRightBottomColor = Color::MakeARGB(255, 240, 240, 240);
			}
			else
			{
                lExteriorLeftTopColor = oStyle->GetButtonBorderStyle()->GetSunkenExteriorLeftTopColor();
                lInteriorLeftTopColor = oStyle->GetButtonBorderStyle()->GetSunkenInteriorLeftTopColor();
                lInteriorRightBottomColor = oStyle->GetButtonBorderStyle()->GetSunkenInteriorRightBottomColor();
                lExteriorRightBottomColor = oStyle->GetButtonBorderStyle()->GetSunkenExteriorRightBottomColor();
			}
			break;
		}

		// Exterior Left
		mp_DrawLine(X1, Y1, X1, Y2, LT_NORMAL, lExteriorLeftTopColor, LDS_SOLID);
		// Exterior Top
		mp_DrawLine(X1, Y1, X2, Y1, LT_NORMAL, lExteriorLeftTopColor, LDS_SOLID);
		// Exterior Right
		mp_DrawLine(X2, Y2, X2, Y1, LT_NORMAL, lExteriorRightBottomColor, LDS_SOLID);
		// Exterior Bottom
		mp_DrawLine(X1, Y2, X2, Y2, LT_NORMAL, lExteriorRightBottomColor, LDS_SOLID);

		// Interior Left
		mp_DrawLine(X1 + 1, Y1 + 1, X1 + 1, Y2 - 1, LT_NORMAL, lInteriorLeftTopColor, LDS_SOLID);
		// Interior Top
		mp_DrawLine(X1 + 1, Y1 + 1, X2 - 1, Y1 + 1, LT_NORMAL, lInteriorLeftTopColor, LDS_SOLID);
		// Interior Right
		mp_DrawLine(X2 - 1, Y2 - 1, X2 - 1, Y1 + 1, LT_NORMAL, lInteriorRightBottomColor, LDS_SOLID);
		// Interior Bottom
		mp_DrawLine(X1 + 1, Y2 - 1, X2 - 1, Y2 - 1, LT_NORMAL, lInteriorRightBottomColor, LDS_SOLID);
		
		if (bFilled == TRUE) 
		{
			mp_DrawLine(X1 + 2, Y1 + 2, X2 - 2, Y2 - 2, LT_FILLED, clrBackColor, LDS_SOLID);
		}
	}
	else
	{
		switch (yEdgeType)
		{
		case ET_RAISED:
			if (oStyle == NULL)
			{
				lExteriorLeftTopColor = Color::White;
				lExteriorRightBottomColor = Color::MakeARGB(255, 64, 64, 64);
			}
			else
			{
                lExteriorLeftTopColor = oStyle->GetButtonBorderStyle()->GetRaisedExteriorLeftTopColor();
                lExteriorRightBottomColor = oStyle->GetButtonBorderStyle()->GetRaisedExteriorRightBottomColor();
			}

			break;
		case ET_SUNKEN:
			if (oStyle == NULL)
			{
				lExteriorLeftTopColor = Color::Gray;
				lExteriorRightBottomColor = Color::WhiteSmoke;
			}
			else
			{
                lExteriorLeftTopColor = oStyle->GetButtonBorderStyle()->GetSunkenExteriorLeftTopColor();
                lExteriorRightBottomColor = oStyle->GetButtonBorderStyle()->GetSunkenExteriorRightBottomColor();
			}
			break;
		}
		mp_DrawLine(X1, Y1, X2, Y1, LT_NORMAL, lExteriorLeftTopColor, LDS_SOLID);
		mp_DrawLine(X1, Y1, X1, Y2, LT_NORMAL, lExteriorLeftTopColor, LDS_SOLID);
		mp_DrawLine(X1, Y2, X2, Y2, LT_NORMAL, lExteriorRightBottomColor, LDS_SOLID);
		mp_DrawLine(X2, Y2, X2, Y1 - 1, LT_NORMAL, lExteriorRightBottomColor, LDS_SOLID);
		if (bFilled == TRUE)
		{
			mp_DrawLine(X1 + 1, Y1 + 1, X2 - 1, Y2 - 1, LT_FILLED, clrBackColor, LDS_SOLID);
		}
	}

}

void clsGraphics::mp_DrawLine(LONG X1, LONG Y1, LONG X2, LONG Y2, GRE_LINETYPE yStyle, Color clrColor, GRE_LINEDRAWSTYLE yDrawStyle, LONG lWidth, BOOL bCreatePens)
{

	Pen mp_ucPen(clrColor, (REAL)lWidth);
	SolidBrush mp_ucBrush(clrColor);
	Point Points[5];
	switch (yStyle)
	{
	case LT_NORMAL:
		Points[0].X = X1;
		Points[0].Y = Y1;
		Points[1].X = X2;
		Points[1].Y = Y2;
		oGraphics()->DrawPolygon(&mp_ucPen,&Points[0], 2);
		break;
	case LT_BORDER:
		Points[0].X = X1;
		Points[0].Y = Y1;
		Points[1].X = X2;
		Points[1].Y = Y1;
		Points[2].X = X2;
		Points[2].Y = Y2;
		Points[3].X = X1;
		Points[3].Y = Y2;
		Points[4].X = X1;
		Points[4].Y = Y1;
		oGraphics()->DrawPolygon(&mp_ucPen, &Points[0], 5);
		break;
	case LT_FILLED:
		Points[0].X = X1;
		Points[0].Y = Y1;
		Points[1].X = X2 + 1;
		Points[1].Y = Y1;
		Points[2].X = X2 + 1;
		Points[2].Y = Y2 + 1;
		Points[3].X = X1;
		Points[3].Y = Y2 + 1;
		Points[4].X = X1;
		Points[4].Y = Y1;
		oGraphics()->FillPolygon(&mp_ucBrush, &Points[0], 5);
		break;
	}
}

void clsGraphics::mp_DrawFigure(LONG X, LONG Y, LONG lDx, LONG lDy, GRE_FIGURETYPE yFigureType, Color lBorderColor, Color lFillColor, GRE_LINEDRAWSTYLE BorderStyle)
{
	Pen oPen(lBorderColor);
	SolidBrush oBrush(lFillColor);
	Point Points[10];
	LONG NumPoints;
	if (lDx % 2 != 0)
	{
		lDx = lDx + 1;
		lDy = lDy + 1;
	}
	switch (yFigureType)
	{
	case FT_PROJECTUP: 
		NumPoints = 5;
		Points[0].X = X;
		Points[0].Y = Y;
		Points[1].X = X + lDx / 2;
		Points[1].Y = Y + lDy / 2;
		Points[2].X = X + lDx / 2;
		Points[2].Y = Y + lDy;
		Points[3].X = X - lDx / 2;
		Points[3].Y = Y + lDy;
		Points[4].X = X - lDx / 2;
		Points[4].Y = Y + lDy / 2;
		mp_DrawFigureAux(&oBrush, &oPen, &Points[0], 5);
		break;
	case FT_PROJECTDOWN:
		NumPoints = 5;
		Points[0].X = X + lDx / 2;
		Points[0].Y = Y;
		Points[1].X = X + lDx / 2;
		Points[1].Y = Y + lDy / 2;
		Points[2].X = X;
		Points[2].Y = Y + lDy;
		Points[3].X = X - lDx / 2;
		Points[3].Y = Y + lDy / 2;
		Points[4].X = X - lDx / 2;
		Points[4].Y = Y;
		mp_DrawFigureAux(&oBrush, &oPen, &Points[0], 5);
		break;
	case FT_DIAMOND:
		NumPoints = 4;
		Points[0].X = X;
		Points[0].Y = Y;
		Points[1].X = X + lDx / 2;
		Points[1].Y = Y + lDy / 2;
		Points[2].X = X;
		Points[2].Y = Y + lDy;
		Points[3].X = X - lDx / 2;
		Points[3].Y = Y + lDy / 2;
		mp_DrawFigureAux(&oBrush, &oPen, &Points[0], 4);
		break;
	case FT_CIRCLEDIAMOND:
		NumPoints = 4;
		Points[0].X = X;
		Points[0].Y = Y + lDy / 4;
		Points[1].X = X + lDx / 4;
		Points[1].Y = Y + lDy / 2;
		Points[2].X = X;
		Points[2].Y = Y + (3 * lDy) / 4;
		Points[3].X = X - lDx / 4;
		Points[3].Y = Y + lDy / 2;
		oGraphics()->DrawEllipse(&oPen, CInt32(X - lDx / 2), Y, lDx, lDy);
		mp_DrawFigureAux(&oBrush, &oPen, &Points[0], 4);
		break;
	case FT_TRIANGLEUP:
		NumPoints = 3;
		Points[0].X = X;
		Points[0].Y = Y;
		Points[1].X = X + lDx / 2;
		Points[1].Y = Y + lDy;
		Points[2].X = X - lDx / 2;
		Points[2].Y = Y + lDy;
		mp_DrawFigureAux(&oBrush, &oPen, &Points[0], 3);
		break;
	case FT_TRIANGLEDOWN:
		NumPoints = 3; 
		Points[0].X = X + lDx / 2;
		Points[0].Y = Y;
		Points[1].X = X - lDx / 2;
		Points[1].Y = Y;
		Points[2].X = X;
		Points[2].Y = Y + lDy;
		mp_DrawFigureAux(&oBrush, &oPen, &Points[0], 3);
		break;
	case FT_TRIANGLERIGHT:
		NumPoints = 3;
		Points[0].X = X;
		Points[0].Y = Y;
		Points[1].X = X;
		Points[1].Y = Y + lDy;
		Points[2].X = X + lDx;
		Points[2].Y = Y + lDy / 2;
		mp_DrawFigureAux(&oBrush, &oPen, &Points[0], 3);
		break;
	case FT_TRIANGLELEFT:
		NumPoints = 3;
		Points[0].X = X;
		Points[0].Y = Y;
		Points[1].X = X;
		Points[1].Y = Y + lDy;
		Points[2].X = X - lDx;
		Points[2].Y = Y + lDy / 2;
		mp_DrawFigureAux(&oBrush, &oPen, &Points[0], 3);
		break;
	case FT_CIRCLETRIANGLEUP:
		NumPoints = 3;
		Points[0].X = X;
		Points[0].Y = Y + lDy / 4;
		Points[1].X = X + lDx / 4;
		Points[1].Y = Y + (3 * lDy) / 4;
		Points[2].X = X - lDx / 4;
		Points[2].Y = Y + (3 * lDy) / 4;
		oGraphics()->DrawEllipse(&oPen, CInt32(X - lDx / 2), Y, lDx, lDy);
		mp_DrawFigureAux(&oBrush, &oPen, &Points[0], 3);
		break;
	case FT_CIRCLETRIANGLEDOWN:
		NumPoints = 3;
		Points[0].X = X - lDx / 4;
		Points[0].Y = Y + lDy / 4;
		Points[1].X = X + lDx / 4;
		Points[1].Y = Y + lDy / 4;
		Points[2].X = X;
		Points[2].Y = Y + (3 * lDy) / 4;
		oGraphics()->DrawEllipse(&oPen, CInt32(X - lDx / 2), Y, lDx, lDy);
		mp_DrawFigureAux(&oBrush, &oPen, &Points[0], 3);
		break;
	case FT_ARROWUP:
		NumPoints = 7;
		Points[0].X = X;
		Points[0].Y = Y;
		Points[1].X = X + lDx / 2;
		Points[1].Y = Y + lDy / 2;
		Points[2].X = X + lDx / 4;
		Points[2].Y = Y + lDy / 2;
		Points[3].X = X + lDx / 4;
		Points[3].Y = Y + lDy;
		Points[4].X = X - lDx / 4;
		Points[4].Y = Y + lDy;
		Points[5].X = X - lDx / 4;
		Points[5].Y = Y + lDy / 2;
		Points[6].X = X - lDx / 2;
		Points[6].Y = Y + lDy / 2;
		mp_DrawFigureAux(&oBrush, &oPen, &Points[0], 7);
		break;
	case FT_ARROWDOWN:
		NumPoints = 7;
		Points[0].X = X - lDx / 4;
		Points[0].Y = Y;
		Points[1].X = X + lDx / 4;
		Points[1].Y = Y;
		Points[2].X = X + lDx / 4;
		Points[2].Y = Y + lDy / 2;
		Points[3].X = X + lDx / 2;
		Points[3].Y = Y + lDy / 2;
		Points[4].X = X;
		Points[4].Y = Y + lDy;
		Points[5].X = X - lDx / 2;
		Points[5].Y = Y + lDy / 2;
		Points[6].X = X - lDx / 4;
		Points[6].Y = Y + lDy / 2;
		mp_DrawFigureAux(&oBrush, &oPen, &Points[0], 7);
		break;
	case FT_CIRCLEARROWUP:
		NumPoints = 7;
		Points[0].X = X;
		Points[0].Y = Y + lDy / 4;
		Points[1].X = X + lDx / 4;
		Points[1].Y = Y + lDy / 2;
		Points[2].X = X + lDx / 8;
		Points[2].Y = Y + lDy / 2;
		Points[3].X = X + lDx / 8;
		Points[3].Y = Y + (3 * lDy) / 4;
		Points[4].X = X - lDx / 8;
		Points[4].Y = Y + (3 * lDy) / 4;
		Points[5].X = X - lDx / 8;
		Points[5].Y = Y + lDy / 2;
		Points[6].X = X - lDx / 4;
		Points[6].Y = Y + lDy / 2;
		oGraphics()->DrawEllipse(&oPen, CInt32(X - lDx / 2), Y, lDx, lDy);
		mp_DrawFigureAux(&oBrush, &oPen, &Points[0], 7);
		break;
	case FT_CIRCLEARROWDOWN:
		NumPoints = 7;
		Points[0].X = X - lDx / 8;
		Points[0].Y = Y + lDy / 4;
		Points[1].X = X + lDx / 8;
		Points[1].Y = Y + lDy / 4;
		Points[2].X = X + lDx / 8;
		Points[2].Y = Y + lDy / 2;
		Points[3].X = X + lDx / 4;
		Points[3].Y = Y + lDy / 2;
		Points[4].X = X;
		Points[4].Y = Y + (3 * lDy) / 4;
		Points[5].X = X - lDx / 4;
		Points[5].Y = Y + lDy / 2;
		Points[6].X = X - lDx / 8;
		Points[6].Y = Y + lDy / 2;
		oGraphics()->DrawEllipse(&oPen, CInt32(X - lDx / 2), Y, lDx, lDy);
		mp_DrawFigureAux(&oBrush, &oPen, &Points[0], 7);
		break;
	case FT_SMALLPROJECTUP:
		NumPoints = 5;
		Points[0].X = X;
		Points[0].Y = Y + lDy / 2;
		Points[1].X = X + lDx / 4;
		Points[1].Y = Y + (3 * lDy) / 4;
		Points[2].X = X + lDx / 4;
		Points[2].Y = Y + lDy;
		Points[3].X = X - lDx / 4;
		Points[3].Y = Y + lDy;
		Points[4].X = X - lDx / 4;
		Points[4].Y = Y + (3 * lDy) / 4;
		mp_DrawFigureAux(&oBrush, &oPen, &Points[0], 5);
		break;
	case FT_SMALLPROJECTDOWN:
		NumPoints = 5;
		Points[0].X = X + lDx / 4;
		Points[0].Y = Y;
		Points[1].X = X + lDx / 4;
		Points[1].Y = Y + lDy / 4;
		Points[2].X = X;
		Points[2].Y = Y + lDy / 2;
		Points[3].X = X - lDx / 4;
		Points[3].Y = Y + lDy / 4;
		Points[4].X = X - lDx / 4;
		Points[4].Y = Y;
		mp_DrawFigureAux(&oBrush, &oPen, &Points[0], 5);
		break;
	case FT_RECTANGLE:
		NumPoints = 4;
		Points[0].X = X - lDx / 8;
		Points[0].Y = Y;
		Points[1].X = X + lDx / 8;
		Points[1].Y = Y;
		Points[2].X = X + lDx / 8;
		Points[2].Y = Y + lDy;
		Points[3].X = X - lDx / 8;
		Points[3].Y = Y + lDy;
		mp_DrawFigureAux(&oBrush, &oPen, &Points[0], 4);
		break;
	case FT_SQUARE:
		NumPoints = 4;
		Points[0].X = X - lDx / 4;
		Points[0].Y = Y + lDx / 4;
		Points[1].X = X + lDx / 4;
		Points[1].Y = Y + lDx / 4;
		Points[2].X = X + lDx / 4;
		Points[2].Y = Y + (3 * lDy) / 4;
		Points[3].X = X - lDx / 4;
		Points[3].Y = Y + (3 * lDy) / 4;
		mp_DrawFigureAux(&oBrush, &oPen, &Points[0], 4);
		break;
	case FT_CIRCLE:
		oGraphics()->FillEllipse(&oBrush, (INT)X - lDx / 2, (INT)Y, (INT)lDx, (INT)lDy);
		break;
	default:
		return;
	}

}

void clsGraphics::mp_DrawFigureAux(Brush* oBrush, Pen* oPen, Point* oPoints, int Count)
{
	oGraphics()->FillPolygon(oBrush, oPoints, Count);
	oGraphics()->DrawPolygon(oPen, oPoints, Count);
}

void clsGraphics::mp_DrawPattern(LONG X1, LONG Y1, LONG X2, LONG Y2, Color clrColor, GRE_PATTERN yDrawStyle, LONG lPatternFactor)
{
	LONG tmp = 0;
	LONG c = 0;
	LONG c1 = 0;
	LONG c2 = 0;
	LONG i1 = 0;
	LONG j1 = 0;
	LONG i2 = 0;
	LONG j2 = 0;
	if (X1 > X2) 
	{
		tmp = X1;
		X1 = X2;
		X2 = tmp;
	}
	if (Y1 > Y2) 
	{
		tmp = Y1;
		Y1 = Y2;
		Y2 = tmp;
	}
	if (yDrawStyle == FP_HORIZONTALLINE || yDrawStyle == FP_CROSS) 
	{
		for (j1 = (Y1 + lPatternFactor); j1 <= Y2; j1 += lPatternFactor) 
		{
			mp_DrawLine(X1, j1, X2, j1, LT_NORMAL, clrColor, LDS_SOLID, 1, TRUE);
		}
	}
	if (yDrawStyle == FP_VERTICALLINE || yDrawStyle == FP_CROSS) 
	{
		for (j1 = (X1 + lPatternFactor); j1 <= X2; j1 += lPatternFactor) 
		{
			mp_DrawLine(j1, Y1, j1, Y2, LT_NORMAL, clrColor, LDS_SOLID, 1, TRUE);
		}
	}
	if (yDrawStyle == FP_UPWARDDIAGONAL || yDrawStyle == FP_DIAGONALCROSS) 
	{
		c1 = CInt32((DOUBLE)(Y1 + X1) / lPatternFactor + 1);
		c2 = CInt32((DOUBLE)(Y2 + X2) / lPatternFactor);
		for (c = c1; c <= c2; c++) 
		{
			i1 = X1;
			i2 = X2;
			j1 = c * lPatternFactor - i1;
			j2 = c * lPatternFactor - i2;
			if (j2 < Y1) 
			{
				i2 = c * lPatternFactor - Y1;
				j2 = c * lPatternFactor - i2;
			}
			if (j1 > Y2) 
			{
				i1 = c * lPatternFactor - Y2;
				j1 = c * lPatternFactor - i1;
			}
			mp_DrawLine(i1, j1, i2, j2, LT_NORMAL, clrColor, LDS_SOLID, 1, FALSE);
		}
	}
	if (yDrawStyle == FP_DOWNWARDDIAGONAL || yDrawStyle == FP_DIAGONALCROSS) 
	{
		c1 = CInt32((DOUBLE)(Y1 - X2) / lPatternFactor + 1);
		c2 = CInt32((DOUBLE)(Y2 - X1) / lPatternFactor);
		for (c = c1; c <= c2; c++) 
		{
			i1 = X1;
			i2 = X2;
			j1 = i1 + c * lPatternFactor;
			j2 = i2 + c * lPatternFactor;
			if (j1 < Y1) 
			{
				i1 = Y1 - c * lPatternFactor;
				j1 = i1 + c * lPatternFactor;
			}
			if (j2 > Y2) 
			{
				i2 = Y2 - c * lPatternFactor;
				j2 = i2 + c * lPatternFactor;
			}
			mp_DrawLine(i1, j1, i2, j2, LT_NORMAL, clrColor, LDS_SOLID, 1, FALSE);
		}
	}
	if (yDrawStyle == FP_LIGHT) 
	{
		for (j1 = (Y1 + 1); j1 <= (Y2 - 1); j1++) 
		{
			if (j1 % 2 == 0) 
			{
				for (j2 = (X1 + 1); j2 <= (X2 - 1); j2 += 4) 
				{
					mp_DrawLine(j2, j1, j2 + 1, j1, LT_NORMAL, clrColor, LDS_SOLID, 1, TRUE);
				}
			}
			else 
			{
				for (j2 = (X1 + 3); j2 <= (X2 - 1); j2 += 4) 
				{
					mp_DrawLine(j2, j1, j2 + 1, j1, LT_NORMAL, clrColor, LDS_SOLID, 1, TRUE);
				}
			}
		}
	}
	if (yDrawStyle == FP_MEDIUM) 
	{
		for (j1 = (Y1 + 1); j1 <= (Y2 - 1); j1++) 
		{
			if (j1 % 2 == 0) 
			{
				for (j2 = (X1 + 1); j2 <= (X2 - 1); j2 += 2) 
				{
					mp_DrawLine(j2, j1, j2 + 1, j1, LT_NORMAL, clrColor, LDS_SOLID, 1, TRUE);
				}
			}
			else 
			{
				for (j2 = (X1 + 2); j2 <= (X2 - 1); j2 += 2) 
				{
					mp_DrawLine(j2, j1, j2 + 1, j1, LT_NORMAL, clrColor, LDS_SOLID, 1, TRUE);
				}
			}
		}
	}
	if (yDrawStyle == FP_DARK) 
	{
		for (j1 = (Y1 + 1); j1 <= (Y2 - 1); j1++) 
		{
			if (j1 % 2 == 0) 
			{
				for (j2 = (X1 + 1); j2 <= (X2 - 1); j2 += 4) 
				{
					if (j2 + 3 < X2) 
					{
						mp_DrawLine(j2, j1, j2 + 3, j1, LT_NORMAL, clrColor, LDS_SOLID, 1, TRUE);
					}
					else 
					{
						mp_DrawLine(j2, j1, X2, j1, LT_NORMAL, clrColor, LDS_SOLID, 1, TRUE);
					}
				}
			}
			else 
			{
				mp_DrawLine(X1, j1, X1 + 2, j1, LT_NORMAL, clrColor, LDS_SOLID, 1, TRUE);
				for (j2 = (X1 + 3); j2 <= (X2 - 1); j2 += 4) 
				{
					if (j2 + 3 < X2) 
					{
						mp_DrawLine(j2, j1, j2 + 3, j1, LT_NORMAL, clrColor, LDS_SOLID, 1, TRUE);
					}
					else 
					{
						mp_DrawLine(j2, j1, X2, j1, LT_NORMAL, clrColor, LDS_SOLID, 1, TRUE);
					}
				}
			}
		}
	}
}

void clsGraphics::mp_DrawPolyLine(Color clrColor, LONG lWidth, GRE_LINEDRAWSTYLE yDrawStyle, Point *r_oPoints, LONG lLength)
{
	Pen hPen(clrColor, (REAL) lWidth);
	oGraphics()->DrawLines(&hPen, r_oPoints, lLength);
}

void clsGraphics::mp_DrawTextEx(LONG X1, LONG Y1, LONG X2, LONG Y2, CString sParam, StringFormat* yFlags, Color clrColor, Font* oFont, BOOL bClip)
{
	SolidBrush mp_ucBrush(clrColor);
	RectF oRect;
	
	oRect.X = (REAL) X1;
	oRect.Y = (REAL) Y1;
	oRect.Width = (REAL) (X2 - X1);
	oRect.Height = (REAL) (Y2 - Y1);
	oGraphics()->DrawString(sParam, sParam.GetLength(), oFont, oRect, yFlags, &mp_ucBrush);
	if (sParam.GetLength() > 0)
	{
		Rect oRectX;
		CharacterRange aRanges[1] = { CharacterRange(0, sParam.GetLength()) };
		Region aRegion[2];
		yFlags->SetMeasurableCharacterRanges(1, aRanges);
		oGraphics()->MeasureCharacterRanges(sParam, sParam.GetLength(), oFont, oRect, yFlags, 2, aRegion);
		aRegion[0].GetBounds(&oRectX, oGraphics());
		mp_oTextFinalLayout.Left = oRectX.GetLeft();
		mp_oTextFinalLayout.Top = oRectX.GetTop();
		mp_oTextFinalLayout.Width = oRectX.Width + mp_lStrWidth(_T("W"), oFont);
		mp_oTextFinalLayout.Height = oRectX.Height;
	}
}

void clsGraphics::mp_DrawAlignedText(LONG X1, LONG Y1, LONG X2, LONG Y2, CString sParam, GRE_HORIZONTALALIGNMENT yHPos, GRE_VERTICALALIGNMENT yVPos, Color clrColor, clsFont* oFont, BOOL bClip)
{
	LONG lHeight = 0;
	LONG lWidth = 0;
	SolidBrush mp_ucBrush(clrColor);
	lHeight = mp_lStrHeight(sParam, oFont);
	lWidth = mp_lStrWidth(sParam, oFont);
	//PointF oPoint;
	RectF oRect;
	StringFormat oFormat;
//	oFormat.SetFormatFlags(StringFormatFlagsNoWrap);

	//bClip = TRUE;
	if (bClip == FALSE && lWidth > (X2 - X1)) return; 
	if (bClip == FALSE && lHeight > (Y2 - Y1)) return;
	//oPoint.X = (REAL) X1;
	//oPoint.Y = (REAL) Y1;
	oRect.X = (REAL) X1;
	oRect.Y = (REAL) Y1;
	oRect.Width = (REAL) (X2 - X1);
	oRect.Height = (REAL) (Y2 - Y1);
	switch (yHPos) 
	{
	case HAL_LEFT:
		oFormat.SetAlignment(StringAlignmentNear);
		break;
	case HAL_CENTER:
		oFormat.SetAlignment(StringAlignmentCenter);
		break;
	case HAL_RIGHT:
		oFormat.SetAlignment(StringAlignmentFar);
		break;
	}
	switch (yVPos) 
	{
	case VAL_TOP:
		oFormat.SetLineAlignment(StringAlignmentNear);
		break;
	case VAL_CENTER:
		oFormat.SetLineAlignment(StringAlignmentCenter);
		break;
	case VAL_BOTTOM:
		oFormat.SetLineAlignment(StringAlignmentFar);
		break;
	}
	oFormat.SetFormatFlags(oFormat.GetFormatFlags() | StringFormatFlagsNoWrap);
	oGraphics()->DrawString(sParam, sParam.GetLength(), oFont->GetFontP(), oRect, &oFormat, &mp_ucBrush);
	if (sParam.GetLength() > 0)
	{
		Rect oRectX;
		CharacterRange aRanges[1] = { CharacterRange(0, sParam.GetLength()) };
		Region aRegion[2];
		oFormat.SetMeasurableCharacterRanges(1, aRanges);
		oGraphics()->MeasureCharacterRanges(sParam, sParam.GetLength(), oFont->GetFontP(), oRect, &oFormat, 2, aRegion);
		aRegion[0].GetBounds(&oRectX, oGraphics());
		mp_oTextFinalLayout.Left = oRectX.GetLeft();
		mp_oTextFinalLayout.Top = oRectX.GetTop();
		mp_oTextFinalLayout.Width = oRectX.Width + mp_lStrWidth(_T("W"), oFont);
		mp_oTextFinalLayout.Height = oRectX.Height;
	}
}

void clsGraphics::mp_ClipRegion(LONG X1, LONG Y1, LONG X2, LONG Y2, BOOL bStore)
{

	if ((mp_bEnableClipRegions == FALSE)) 
	{
		return;
	}
	if (mp_bCustomPrinting == TRUE) 
	{
		if (X1 < mp_lX1) 
		{
			X1 = mp_lX1;
		}
		if (Y1 < mp_lY1) 
		{
			Y1 = mp_lY1;
		}
		if (X2 > mp_lX2) 
		{
			X2 = mp_lX2;
		}
		if (Y2 > mp_lY2) 
		{
			Y2 = mp_lY2;
		}
		if ((X1 > mp_lX2) || (Y1 > mp_lY2) || (X2 < mp_lX1) || (Y2 < mp_lY1)) 
		{
			X1 = 0;
			Y1 = 0;
			X2 = 0;
			Y2 = 0;
		}
	}
	Rect oRect(X1, Y1, (X2 - X1 + 1), (Y2 - Y1 + 1));
	Region oRegion(oRect);
	if (bStore == TRUE) 
	{
		mp_udtPreviousClipRegion.left = X1;
		mp_udtPreviousClipRegion.right = X2;
		mp_udtPreviousClipRegion.top = Y1;
		mp_udtPreviousClipRegion.bottom = Y2;
	}
	oGraphics()->SetClip(&oRegion);
}

void clsGraphics::mp_RestorePreviousClipRegion()
{
	if ((mp_bEnableClipRegions == FALSE)) 
	{
		return;
	}
	mp_ClipRegion(mp_udtPreviousClipRegion.left, mp_udtPreviousClipRegion.top, mp_udtPreviousClipRegion.right, mp_udtPreviousClipRegion.bottom, FALSE);
}

void clsGraphics::mp_ClearClipRegion()
{
	oGraphics()->ResetClip();
}

void clsGraphics::mp_TileImageHorizontal(clsImage* oImageHandle, LONG X1, LONG Y1, LONG X2, LONG Y2, BOOL bTransparent)
{
	LONG XBuff = 0;
	LONG lImageWidth = 0;
	LONG lImageHeight = 0;
	lImageHeight = oImageHandle->GetHeight();
	lImageWidth = oImageHandle->GetWidth();
	while (XBuff < (X2 - X1)) 
	{
		if ((XBuff + lImageWidth) > (X2 - X1)) 
		{
			mp_PaintImage(oImageHandle, X2 - lImageWidth, Y1, X2, Y1 + lImageHeight, 0, 0, bTransparent);
		}
		else 
		{
			mp_PaintImage(oImageHandle, X1 + XBuff, Y1, X1 + XBuff + lImageWidth, Y1 + lImageHeight, 0, 0, bTransparent);
		}
		XBuff = XBuff + lImageWidth;
	}
}

void clsGraphics::mp_PaintImage(clsImage* oImageHandle, LONG X1, LONG Y1, LONG X2, LONG Y2, LONG xOrigin, LONG yOrigin, BOOL bUseMask)
{
	if (oImageHandle->GetImageP() == NULL)
	{
		return;
	}
	RectF DestRect;
	RectF OriginRect;
	OriginRect.X = (REAL) xOrigin;
	OriginRect.Y = (REAL) yOrigin;
	OriginRect.Width = (REAL) (oImageHandle->GetWidth() - xOrigin);
	OriginRect.Height = (REAL) (oImageHandle->GetHeight() - yOrigin);
	DestRect.X = (REAL) X1;
	DestRect.Y = (REAL) Y1;
	DestRect.Width = (REAL) (X2 - X1);
	DestRect.Height = (REAL) (Y2 - Y1);
	if (bUseMask == FALSE) 
	{
		oGraphics()->DrawImage(oImageHandle->mp_oImage, DestRect, OriginRect.X, OriginRect.Y, OriginRect.Width, OriginRect.Height, UnitPixel);
	}
	else 
	{
		oGraphics()->DrawImage(oImageHandle->mp_oImage, DestRect, OriginRect.X, OriginRect.Y, OriginRect.Width, OriginRect.Height, UnitPixel);

		//Bitmap oBitmap;
		//oBitmap = (Bitmap) oImageHandle->;
		//oBitmap.MakeTransparent(Color::White);
		//oGraphics.DrawImage(oBitmap, DestRect, OriginRect, UnitPixel);
	}
}

void clsGraphics::mp_DrawImage(clsImage* oImage, GRE_HORIZONTALALIGNMENT yHorizontalAlignment, GRE_VERTICALALIGNMENT yVerticalAlignment, LONG lImageXMargin, LONG lImageYMargin, LONG X1, LONG X2, LONG Y1, LONG Y2, BOOL bTransparent)
{
	if (oImage->GetImageP() == NULL)
	{
		return;
	}
	BOOL bDrawImage = FALSE;
	BOOL bHorizontalSmall = FALSE;
	BOOL bVerticalSmall = FALSE;
	LONG XOrigin = 0;
	LONG YOrigin = 0;
	LONG xDest = 0;
	LONG yDest = 0;
	LONG lxWidth = 0;
	LONG lyHeight = 0;
	LONG lImageHeight = 0;
	LONG lImageWidth = 0;
	if ((oImage == NULL)) 
	{
		return;
	}
	lImageHeight = oImage->GetHeight();
	lImageWidth = oImage->GetWidth();
	if (yHorizontalAlignment == HAL_CENTER) 
	{
		lImageXMargin = 0;
	}
	if (yVerticalAlignment == VAL_CENTER) 
	{
		lImageYMargin = 0;
	}
	bDrawImage = TRUE;
	if ((X2 - X1) < (lImageWidth + lImageXMargin)) 
	{
		lxWidth = X2 - X1 - lImageXMargin;
		if (lxWidth <= 0) bDrawImage = FALSE; 
		bHorizontalSmall = TRUE;
	}
	else 
	{
		lxWidth = lImageWidth;
		bHorizontalSmall = FALSE;
	}
	if ((Y2 - Y1) < (lImageHeight + lImageYMargin)) 
	{
		lyHeight = Y2 - Y1 - lImageYMargin;
		if (lyHeight <= 0) bDrawImage = FALSE; 
		bVerticalSmall = TRUE;
	}
	else 
	{
		lyHeight = lImageHeight;
		bVerticalSmall = FALSE;
	}
	if (bHorizontalSmall == FALSE) 
	{
		switch (yHorizontalAlignment) 
		{
		case HAL_LEFT:
			xDest = X1 + lImageXMargin;
			break;
		case HAL_CENTER:
			xDest = ((X2 - X1) - lImageWidth) / 2 + X1;
			break;
		case HAL_RIGHT:
			xDest = X2 - lImageWidth - lImageXMargin;
			break;
		}
		XOrigin = 0;
	}
	else 
	{
		switch (yHorizontalAlignment) 
		{
		case HAL_LEFT:
			XOrigin = 0;
			xDest = X1 + lImageXMargin;
			break;
		case HAL_CENTER:
			XOrigin = (lImageWidth - lxWidth) / 2;
			xDest = X1;
			break;
		case HAL_RIGHT:
			XOrigin = lImageWidth - lxWidth;
			xDest = X2 - lxWidth - lImageXMargin;
			break;
		}
	}
	if (bVerticalSmall == FALSE) 
	{
		switch (yVerticalAlignment) 
		{
		case VAL_TOP:
			yDest = Y1 + lImageYMargin;
			break;
		case VAL_CENTER:
			yDest = ((Y2 - Y1) - lImageHeight) / 2 + Y1;
			break;
		case VAL_BOTTOM:
			yDest = Y2 - lImageHeight - lImageYMargin;
			break;
		}
		YOrigin = 0;
	}
	else 
	{
		switch (yVerticalAlignment) 
		{
		case VAL_TOP:
			YOrigin = 0;
			yDest = Y1 + lImageYMargin;
			break;
		case VAL_CENTER:
			YOrigin = (lImageHeight - lyHeight) / 2;
			yDest = Y1;
			break;
		case VAL_BOTTOM:
			YOrigin = lImageHeight - lyHeight;
			yDest = Y2 - lyHeight - lImageYMargin;
			break;
		}
	}
	if (bDrawImage == TRUE) 
	{
		mp_PaintImage(oImage, xDest, yDest, xDest + lxWidth, yDest + lyHeight, XOrigin, YOrigin, bTransparent);
	}
}

void clsGraphics::mp_DrawFocusRectangle(LONG X1, LONG Y1, LONG X2, LONG Y2)
{
	if (mp_bCustomPrinting == TRUE)
	{
		return;
	}
	CPen oPen;
	CPen* holdpen;
	RECT udtRect;
	HDC lHdc;
	udtRect.left = X1;
	udtRect.top = Y1;
	udtRect.right = X2;
	udtRect.bottom = Y2;
	oPen.CreatePen(LDS_DOT,1, CLR_WHITE);
	lHdc = oGraphics()->GetHDC();
	holdpen = (CPen*) SelectObject(lHdc, &oPen);
    DrawFocusRect(lHdc, &udtRect);
	SelectObject(lHdc, holdpen);
	DeleteObject(oPen);
	oGraphics()->ReleaseHDC(lHdc);
}

void clsGraphics::mp_GradientFill(LONG X1, LONG Y1, LONG X2, LONG Y2, Color clrStartColor, Color clrEndColor, GRE_GRADIENTFILLMODE yGradientType)
{
	if ((X2 - X1) <= 0)
	{
		return;
	}
	if ((Y2 - Y1) <= 0)
	{
		return;
	}
	RectF oRect((REAL)X1, (REAL)Y1, (REAL)(X2 - X1), (REAL)(Y2 - Y1));

	Point Points[5];
	Points[0].X = X1;
	Points[0].Y = Y1;
	Points[1].X = X2 + 1;
	Points[1].Y = Y1;
	Points[2].X = X2 + 1;
	Points[2].Y = Y2 + 1;
	Points[3].X = X1;
	Points[3].Y = Y2 + 1;
	Points[4].X = X1;
	Points[4].Y = Y1;
	if ((yGradientType == GDT_VERTICAL)) 
	{
		LinearGradientBrush mp_ucBrushV(oRect, clrStartColor, clrEndColor, LinearGradientModeVertical);
		oGraphics()->FillPolygon(&mp_ucBrushV, Points, 5);
	}
	else if ((yGradientType == GDT_HORIZONTAL)) 
	{
		LinearGradientBrush mp_ucBrushH(oRect, clrStartColor, clrEndColor, LinearGradientModeHorizontal);
		oGraphics()->FillPolygon(&mp_ucBrushH, Points, 5);
	}
	
}

void clsGraphics::mp_HatchFill(LONG X1, LONG Y1, LONG X2, LONG Y2, Color clrForeColor, Color clrBackColor, GRE_HATCHSTYLE yHatchStyle)
{
	HatchBrush mp_ucBrush((HatchStyle)yHatchStyle, clrForeColor, clrBackColor);
	Point Points[5];
	Points[0].X = X1;
	Points[0].Y = Y1;
	Points[1].X = X2 + 1;
	Points[1].Y = Y1;
	Points[2].X = X2 + 1;
	Points[2].Y = Y2 + 1;
	Points[3].X = X1;
	Points[3].Y = Y2 + 1;
	Points[4].X = X1;
	Points[4].Y = Y1;
	oGraphics()->FillPolygon(&mp_ucBrush, Points, 5);
}

void clsGraphics::mp_ResetFocusRectangle()
{
	mp_EraseReversibleFrames();
	mp_lFocusLeft = 0;
	mp_lFocusTop = 0;
	mp_lFocusRight = 0;
	mp_lFocusLeft = 0;
}

void clsGraphics::mp_DrawReversibleFrameEx()
{
	mp_DrawReversibleFrame(mp_lFocusLeft, mp_lFocusTop, mp_lFocusRight, mp_lFocusBottom);
}

void clsGraphics::mp_DrawReversibleLine(LONG X1, LONG Y1, LONG X2, LONG Y2)
{
	CPen oPen;
	CPen* holdpen;
	CRect* udtRect;
	HDC lHdc;
	LONG lARFUBound;
	udtRect = new CRect();
	udtRect->left = X1;
	udtRect->top = Y1;
	udtRect->right = X2;
	udtRect->bottom = Y2;
	lHdc = GetDC(mp_oControl->f_Hwnd());
	oPen.CreatePen(LDS_SOLID,1, CLR_BLACK);
	holdpen = (CPen*) SelectObject(lHdc, &oPen);
	SetROP2(lHdc, R2_NOTXORPEN);
	MoveToEx(lHdc, udtRect->left, udtRect->top, NULL);
	LineTo(lHdc, udtRect->right, udtRect->bottom);
	SelectObject(lHdc, holdpen);
	DeleteObject(oPen);
	ReleaseDC(mp_oControl->f_Hwnd(), lHdc);
	lARFUBound = mp_audtActiveReversibleLines.GetSize() + 1;
	mp_audtActiveReversibleLines.SetSize(lARFUBound);
	mp_audtActiveReversibleLines.SetAt(lARFUBound - 1, udtRect);
}
void clsGraphics::mp_EraseReversibleLines()
{
	CPen oPen;
	CPen* holdpen;
	RECT* udtRect;
	HDC lHdc;
	LONG lIndex;
	if (mp_audtActiveReversibleLines.GetSize() == 0) 
	{
		return;
	}
	for (lIndex = mp_audtActiveReversibleLines.GetSize() - 1; lIndex >= 0; lIndex--)
	{
		udtRect = (RECT*) mp_audtActiveReversibleLines[lIndex];
		lHdc = GetDC(mp_oControl->f_Hwnd());
		oPen.CreatePen(LDS_SOLID,1, CLR_BLACK);
		holdpen = (CPen*) SelectObject(lHdc, &oPen);
		SetROP2(lHdc, R2_NOTXORPEN);
		MoveToEx(lHdc, udtRect->left, udtRect->top, NULL);
		LineTo(lHdc, udtRect->right, udtRect->bottom);
		SelectObject(lHdc, holdpen);
		DeleteObject(oPen);
		ReleaseDC(mp_oControl->f_Hwnd(), lHdc);
		delete(udtRect);
		udtRect = NULL;
	}
	mp_audtActiveReversibleLines.RemoveAll();
	mp_audtActiveReversibleLines.SetSize(0);
}

void clsGraphics::mp_DrawReversibleFrame(LONG X1, LONG Y1, LONG X2, LONG Y2)
{
	CPen oPen;
	CPen* holdpen;
	CRect* udtRect;
	HDC lHdc;
	LONG lARFUBound;
	udtRect = new CRect();
	udtRect->left = X1;
	udtRect->top = Y1;
	udtRect->right = X2;
	udtRect->bottom = Y2;
	lHdc = GetDC(mp_oControl->f_Hwnd());
	oPen.CreatePen(LDS_SOLID,1, CLR_BLACK);
	holdpen = (CPen*) SelectObject(lHdc, &oPen);
	SetROP2(lHdc, R2_NOTXORPEN);
	Rectangle(lHdc, udtRect->left, udtRect->top, udtRect->right, udtRect->bottom);
	SelectObject(lHdc, holdpen);
	DeleteObject(oPen);
	ReleaseDC(mp_oControl->f_Hwnd(), lHdc);
	lARFUBound = mp_audtActiveReversibleFrames.GetSize() + 1;
	mp_audtActiveReversibleFrames.SetSize(lARFUBound);
	mp_audtActiveReversibleFrames.SetAt(lARFUBound - 1, udtRect);
}

void clsGraphics::mp_EraseReversibleFrames()
{
	CPen oPen;
	CPen* holdpen;
	RECT* udtRect;
	HDC lHdc;
	LONG lIndex;
	if (mp_audtActiveReversibleFrames.GetSize() == 0) 
	{
		return;
	}
	for (lIndex = mp_audtActiveReversibleFrames.GetSize() - 1; lIndex >= 0; lIndex--)
	{
		udtRect = (RECT*) mp_audtActiveReversibleFrames[lIndex];
		lHdc = GetDC(mp_oControl->f_Hwnd());
		oPen.CreatePen(LDS_SOLID,1, CLR_BLACK);
		holdpen = (CPen*) SelectObject(lHdc, &oPen);
		SetROP2(lHdc, R2_NOTXORPEN);
		Rectangle(lHdc, udtRect->left, udtRect->top, udtRect->right, udtRect->bottom);
		SelectObject(lHdc, holdpen);
		DeleteObject(oPen);
		ReleaseDC(mp_oControl->f_Hwnd(), lHdc);
		delete(udtRect);
		udtRect = NULL;
	}
	mp_audtActiveReversibleFrames.RemoveAll();
	mp_audtActiveReversibleFrames.SetSize(0);

}

Color clsGraphics::mp_ConvertColor(OLE_COLOR dwOleColour)
{
	Color oColor;
	COLORREF rColor;
	OleTranslateColor(dwOleColour, 0, &rColor);
	oColor.SetFromCOLORREF(rColor);
	return oColor;
}

BOOL clsGraphics::mp_LineIntersection(LONG X1, LONG Y1, LONG X2, LONG Y2)
{
	return TRUE;
}

void clsGraphics::mp_DrawItemI(LONG X1, LONG Y1, LONG X2, LONG Y2, CString sStyleIndex, CString sText, BOOL bSelected, clsImage* oImage, LONG lLeftTrim, LONG lRightTrim, clsStyle* v_oStyle)
{
	clsStyle* oStyle;
	clsMilestoneStyle* oMilestoneStyle;
	if ((v_oStyle == NULL)) 
	{
		if (g_IsNumeric(sStyleIndex)) 
		{
			if (CInt32(sStyleIndex) < 0 || CInt32(sStyleIndex) > mp_oControl->GetStyles()->GetCount()) 
			{
				mp_oControl->mp_ErrorReport(STYLE_INVALID_INDEX, "Style object element not found when preparing to draw, invalid index", "mp_DrawItemI");
				return;
			}
		}
		else 
		{
			if (mp_oControl->GetStyles()->mp_oCollection->m_bDoesKeyExist(sStyleIndex) == FALSE) 
			{
				mp_oControl->mp_ErrorReport(STYLE_INVALID_KEY, _T("Style object element not found when preparing to draw, invalid key (\"")  + sStyleIndex + _T("\")"), _T("mp_DrawItemI"));
				return;
			}
		}
		oStyle = mp_oControl->GetStyles()->FItem(sStyleIndex);
	}
	else 
	{
		oStyle = v_oStyle;
	}
	switch (oStyle->GetAppearance()) 
	{
	case SA_FLAT:
		oMilestoneStyle = oStyle->mp_oMilestoneStyle;
		mp_DrawFigure(CInt32(CDbl(X1) + ((CDbl(X2) - CDbl(X1)) / CDbl(2))), Y1, Y2 - Y1, Y2 - Y1, (GRE_FIGURETYPE) oMilestoneStyle->GetShapeIndex(), oMilestoneStyle->GetBorderColor(), oMilestoneStyle->GetFillColor(), LDS_SOLID);
		break;
	case SA_GRAPHICAL:
		if (oStyle->mp_oMilestoneStyle->GetImage() != NULL)
		{
			mp_DrawImage(oStyle->mp_oMilestoneStyle->GetImage(), (GRE_HORIZONTALALIGNMENT) oStyle->GetImageAlignmentHorizontal(), (GRE_VERTICALALIGNMENT) oStyle->GetImageAlignmentVertical(), oStyle->GetImageXMargin(), oStyle->GetImageYMargin(), X1, Y1, X2, Y2, oStyle->GetUseMask());
		}
		else if (oImage != NULL)
		{
			mp_DrawImage(oImage, (GRE_HORIZONTALALIGNMENT) oStyle->GetImageAlignmentHorizontal(), (GRE_VERTICALALIGNMENT) oStyle->GetImageAlignmentVertical(), oStyle->GetImageXMargin(), oStyle->GetImageYMargin(), X1, Y1, X2, Y2, oStyle->GetUseMask());
		}
		break;
	default:
		oMilestoneStyle = oStyle->mp_oMilestoneStyle;
		mp_DrawFigure(CInt32(CDbl(X1) + ((CDbl(X2) - CDbl(X1)) / CDbl(2))), Y1, Y2 - Y1, Y2 - Y1, (GRE_FIGURETYPE) oMilestoneStyle->GetShapeIndex(), oMilestoneStyle->GetBorderColor(), oMilestoneStyle->GetFillColor(), LDS_SOLID);
		break;
	}
	mp_DrawItemText(X1, Y1, X2, Y2, lLeftTrim, lRightTrim, oStyle, sText);
	
    if ((oStyle->GetSelectionRectangleStyle()->GetVisible() == TRUE) && (bSelected == TRUE))
    {
        if (oStyle->GetSelectionRectangleStyle()->GetMode() == SRM_DOTTED)
        {
            mp_DrawFocusRectangle(X1, Y1, X2, Y2);
        }
        else if (oStyle->GetSelectionRectangleStyle()->GetMode() == SRM_COLOR)
        {
            mp_DrawLine(X1, Y1, X2, Y2, LT_BORDER, oStyle->GetSelectionRectangleStyle()->GetColor(), LDS_SOLID, oStyle->GetSelectionRectangleStyle()->GetBorderWidth());
        }
    }
}

void clsGraphics::mp_DrawItemI(clsTask* oTask, CString sStyleIndex, BOOL Selected, clsStyle* v_oStyle)
{
	mp_DrawItemI(oTask->GetLeft(), oTask->GetTop(), oTask->GetRight(), oTask->GetBottom(), oTask->GetStyleIndex(), oTask->GetText(), Selected, oTask->GetStyle()->GetMilestoneStyle()->GetImage(), oTask->GetLeftTrim(), oTask->GetRightTrim(), v_oStyle);
}

void clsGraphics::mp_DrawItemEx(LONG X1, LONG Y1, LONG X2, LONG Y2, CString sText, BOOL bIsSelected, clsImage* oImage, LONG lLeftTrim, LONG lRightTrim, clsStyle* v_oStyle, Color clrBackColor, Color clrForeColor, Color clrStartGradientColor, Color clrEndGradientColor, Color clrHatchBackColor, Color clrHatchForeColor)
{
	clsStyle* oStyle;
	clsTaskStyle* oTaskStyle;
	if (v_oStyle == NULL)
	{
		mp_oControl->mp_ErrorReport(STYLE_NULL, "Style object is null when preparing to draw.", "mp_DrawItemEx");
		return;
	}
	else
	{
		oStyle = v_oStyle;
	}
	oTaskStyle = oStyle->GetTaskStyle();
	switch (oStyle->GetAppearance()) 
	{
	case SA_RAISED:
		{
			mp_DrawEdge(X1, Y1, X2, Y2, clrBackColor, (GRE_BUTTONSTYLE) oStyle->GetButtonStyle(), ET_RAISED, TRUE, v_oStyle);
			break;
		}
	case SA_SUNKEN:
		{
			mp_DrawEdge(X1, Y1, X2, Y2, clrBackColor, (GRE_BUTTONSTYLE) oStyle->GetButtonStyle(), ET_SUNKEN, TRUE, v_oStyle);
			break;
		}
	case SA_FLAT:
		{
			LONG lTop = 0;
			LONG lBottom = 0;
			lTop = Y1;
			lBottom = Y2;
			switch (oStyle->GetFillMode()) 
			{
			case FM_COMPLETELYFILLED:
				break;
			case FM_UPPERHALFFILLED:
				lBottom = Y1 + ((Y2 - Y1) / 2);
				break;
			case FM_LOWERHALFFILLED:
				lTop = Y2 - ((Y2 - Y1) / 2);
				break;
			}
			if ((oStyle->GetBackgroundMode() == FP_SOLID)) 
			{
				mp_DrawLine(X1, lTop, X2, lBottom, LT_FILLED, clrBackColor, LDS_SOLID, 1, TRUE);
			}
			else if ((oStyle->GetBackgroundMode() == FP_GRADIENT)) 
			{
				mp_GradientFill(X1, lTop, X2, lBottom, clrStartGradientColor, clrEndGradientColor, (GRE_GRADIENTFILLMODE) oStyle->GetGradientFillMode());
			}
			else if ((oStyle->GetBackgroundMode() == FP_PATTERN)) 
			{
				mp_DrawPattern(X1, lTop, X2, lBottom, clrBackColor, (GRE_PATTERN) oStyle->GetPattern(), oStyle->GetPatternFactor());
			}
			else if ((oStyle->GetBackgroundMode() == FP_HATCH)) 
			{
				mp_HatchFill(X1, lTop, X2, lBottom, clrHatchForeColor, clrHatchBackColor, (GRE_HATCHSTYLE) oStyle->GetHatchStyle());
			}

			if (oStyle->GetBorderStyle() == SBR_SINGLE) 
			{
				mp_DrawLine(X1, lTop, X2, lBottom, LT_BORDER, oStyle->GetBorderColor(), LDS_SOLID, oStyle->GetBorderWidth(), TRUE);
			}
			else if (oStyle->GetBorderStyle() == SBR_CUSTOM) 
			{
				if (oStyle->mp_oCustomBorderStyle->GetLeft() == TRUE) 
				{
					mp_DrawLine(X1, lTop, X1, lBottom, LT_NORMAL, oStyle->GetBorderColor(), LDS_SOLID, oStyle->GetBorderWidth(), TRUE);
				}
				if (oStyle->mp_oCustomBorderStyle->GetTop() == TRUE) 
				{
					mp_DrawLine(X1, lTop, X2, lTop, LT_NORMAL, oStyle->GetBorderColor(), LDS_SOLID, oStyle->GetBorderWidth(), TRUE);
				}
				if (oStyle->mp_oCustomBorderStyle->GetRight() == TRUE) 
				{
					mp_DrawLine(X2, lTop, X2, lBottom, LT_NORMAL, oStyle->GetBorderColor(), LDS_SOLID, oStyle->GetBorderWidth(), TRUE);
				}
				if (oStyle->mp_oCustomBorderStyle->GetBottom() == TRUE) 
				{
					mp_DrawLine(X1, lBottom, X2, lBottom, LT_NORMAL, oStyle->GetBorderColor(), LDS_SOLID, oStyle->GetBorderWidth(), TRUE);
				}
			}
			mp_DrawFigure(X2, Y1, Y2 - Y1, Y2 - Y1,(GRE_FIGURETYPE) oTaskStyle->GetEndShapeIndex(), oTaskStyle->GetEndBorderColor(), oTaskStyle->GetEndFillColor(), LDS_SOLID);
			mp_DrawFigure(X1, Y1, Y2 - Y1, Y2 - Y1,(GRE_FIGURETYPE) oTaskStyle->GetStartShapeIndex(), oTaskStyle->GetStartBorderColor(), oTaskStyle->GetStartFillColor(), LDS_SOLID);
			break;
		}
	case SA_GRAPHICAL:
		{
			if ((oTaskStyle->GetMiddleImage() == NULL) || (oTaskStyle->GetStartImage() == NULL) || (oTaskStyle->GetEndImage() == NULL))
			{
			}
			else
			{
				LONG lImageHeight = 0;
				LONG lImageWidth = 0;
				lImageHeight = oTaskStyle->GetMiddleImage()->GetHeight();
				lImageWidth = oTaskStyle->GetMiddleImage()->GetWidth();
				mp_TileImageHorizontal(oTaskStyle->GetMiddleImage(), X1, Y1, X2, Y2, oStyle->GetUseMask());
				//// Exit if the start and end sections don't fit
				if ((X2 - X1) > (lImageWidth * 2)) 
				{
					// Left Section
					mp_PaintImage(oTaskStyle->GetStartImage(), X1, Y1, X1 + lImageWidth, Y1 + lImageHeight, 0, 0, oStyle->GetUseMask());
					// Right Section
					mp_PaintImage(oTaskStyle->GetEndImage(), X2 - lImageWidth, Y1, X2, Y1 + lImageHeight, 0, 0, oStyle->GetUseMask());
				}
			}
			break;
		}
	}
	if ((oImage != NULL)) 
	{
		mp_DrawImage(oImage, (GRE_HORIZONTALALIGNMENT) oStyle->GetImageAlignmentHorizontal(), (GRE_VERTICALALIGNMENT) oStyle->GetImageAlignmentVertical(), oStyle->GetImageXMargin(), oStyle->GetImageYMargin(), X1, X2, Y1, Y2, oStyle->GetUseMask());
	}
	mp_DrawItemText(X1, Y1, X2, Y2, lLeftTrim, lRightTrim, oStyle, sText);
    if (oStyle->GetSelectionRectangleStyle()->GetVisible() == TRUE && bIsSelected)
    {
		mp_DrawSelectionRectangle(X1, Y1, X2, Y2, oStyle);
    }
}

void clsGraphics::mp_DrawSelectionRectangle(LONG X1, LONG Y1, LONG X2, LONG Y2, clsStyle* oStyle)
{
    if (oStyle->GetSelectionRectangleStyle()->GetMode() == SRM_DOTTED)
    {
        mp_DrawFocusRectangle(X1 + oStyle->GetSelectionRectangleStyle()->GetOffsetLeft(), Y1 + oStyle->GetSelectionRectangleStyle()->GetOffsetTop(), X2 - oStyle->GetSelectionRectangleStyle()->GetOffsetRight(), Y2 - oStyle->GetSelectionRectangleStyle()->GetOffsetBottom());
    }
    else if (oStyle->GetSelectionRectangleStyle()->GetMode() == SRM_COLOR)
    {
        mp_DrawLine(X1 + oStyle->GetSelectionRectangleStyle()->GetOffsetLeft(), Y1 + oStyle->GetSelectionRectangleStyle()->GetOffsetTop(), X2 - oStyle->GetSelectionRectangleStyle()->GetOffsetRight(), Y2 - oStyle->GetSelectionRectangleStyle()->GetOffsetBottom(), LT_BORDER, oStyle->GetSelectionRectangleStyle()->GetColor(), LDS_SOLID, oStyle->GetSelectionRectangleStyle()->GetBorderWidth());
    }
}

void clsGraphics::mp_DrawItemRow(LONG X1, LONG Y1, LONG X2, LONG Y2, CString sText, BOOL bIsSelected, clsImage* oImage, LONG lLeftTrim, LONG lRightTrim, clsStyle* v_oStyle, clsRow* oRow)
{
	LONG yBackgroundMode;
    if (oRow->GetHighlighted() == TRUE)
	{
        mp_DrawLine(X1, Y1, X2, Y2, LT_FILLED, mp_oControl->GetConfiguration()->GetRowHighlightedBackColor(), LDS_SOLID);
        yBackgroundMode = v_oStyle->GetBackgroundMode();
		if (v_oStyle->GetBackgroundMode() != FP_TRANSPARENT)
		{
            v_oStyle->SetBackgroundMode(FP_TRANSPARENT);
		}
        mp_DrawItem(X1, Y1, X2, Y2, "", sText, bIsSelected, oImage, lLeftTrim, lRightTrim, v_oStyle);
		v_oStyle->SetBackgroundMode(yBackgroundMode);
	}
	else
	{
        if (mp_oControl->GetConfiguration()->GetAlternatingRowStyles() == TRUE)
		{
			yBackgroundMode = v_oStyle->GetBackgroundMode();
            if (v_oStyle->GetBackgroundMode() != FP_TRANSPARENT)
			{
                v_oStyle->SetBackgroundMode(FP_TRANSPARENT);
			}
            if (oRow->GetIndex() % 2 == 1)
			{
                //Odd
                mp_DrawLine(X1, Y1, X2, Y2, LT_FILLED, mp_oControl->GetConfiguration()->GetRowOddBackColor(), LDS_SOLID);
                mp_DrawItem(X1, Y1, X2, Y2, "", sText, bIsSelected, oImage, lLeftTrim, lRightTrim, v_oStyle);
			}
            else
			{
                //Even
                mp_DrawLine(X1, Y1, X2, Y2, LT_FILLED, mp_oControl->GetConfiguration()->GetRowEvenBackColor(), LDS_SOLID);
                mp_DrawItem(X1, Y1, X2, Y2, "", sText, bIsSelected, oImage, lLeftTrim, lRightTrim, v_oStyle);
			}
			v_oStyle->SetBackgroundMode(yBackgroundMode);
		}
        else
		{
            mp_DrawItem(X1, Y1, X2, Y2, "", sText, bIsSelected, oImage, lLeftTrim, lRightTrim, v_oStyle);
		}
	}
}


void clsGraphics::mp_DrawItem(LONG X1, LONG Y1, LONG X2, LONG Y2, CString sStyleIndex, CString sText, BOOL bIsSelected, clsImage* oImage, LONG lLeftTrim, LONG lRightTrim, clsStyle* v_oStyle)
{
	clsStyle* oStyle;
	if ((v_oStyle == NULL)) 
	{
		if (g_IsNumeric(sStyleIndex)) 
		{
			if (CInt32(sStyleIndex) < 0 || CInt32(sStyleIndex) > mp_oControl->GetStyles()->GetCount()) 
			{
				mp_oControl->mp_ErrorReport(STYLE_INVALID_INDEX, "Style object element not found when preparing to draw, invalid index", "mp_DrawItem");
				return;
			}
		}
		else 
		{
			if (mp_oControl->GetStyles()->mp_oCollection->m_bDoesKeyExist(sStyleIndex) == FALSE) 
			{
				mp_oControl->mp_ErrorReport(STYLE_INVALID_KEY, _T("Style object element not found when preparing to draw, invalid key (\"")  + sStyleIndex + _T("\")"), _T("mp_DrawItem"));
				return;
			}
		}
		oStyle = mp_oControl->GetStyles()->FItem(sStyleIndex);
	}
	else 
	{
		oStyle = v_oStyle;
	}
	mp_DrawItemEx(X1, Y1, X2, Y2, sText, bIsSelected, oImage, lLeftTrim, lRightTrim, oStyle, oStyle->GetBackColor(), oStyle->GetForeColor(), oStyle->GetStartGradientColor(), oStyle->GetEndGradientColor(), oStyle->GetHatchBackColor(), oStyle->GetHatchForeColor());
}



void clsGraphics::mp_DrawItemText(LONG X1, LONG Y1, LONG X2, LONG Y2, LONG lLeftTrim, LONG lRightTrim, clsStyle *oStyle, CString sText)
{
	LONG lTextLeft = 0;
	LONG lTextRight = 0;
	LONG lTextTop = 0;
	LONG lTextBottom = 0;
	Gdiplus::StringFormat oFlags;
	if (oStyle->GetTextVisible() == FALSE) 
	{
		return;
	}
	if (sText == "") 
	{
		return;
	}
	switch (oStyle->GetTextPlacement()) 
	{
	case SCP_OBJECTEXTENTSPLACEMENT:
		if ((oStyle->GetDrawTextInVisibleArea() == FALSE)) 
		{
			lTextLeft = X1;
			lTextRight = X2;
		}
		else 
		{
			lTextLeft = lLeftTrim;
			lTextRight = lRightTrim;
		}

		lTextTop = Y1;
		lTextBottom = Y2;
		if (oStyle->GetTextAlignmentHorizontal() == HAL_LEFT) 
		{
			lTextLeft = X1 + oStyle->GetTextXMargin();
		}

		if (oStyle->GetTextAlignmentHorizontal() == HAL_RIGHT) 
		{
			lTextRight = X2 - oStyle->GetTextXMargin();
		}

		if (oStyle->GetTextAlignmentVertical() == VAL_TOP) 
		{
			lTextTop = Y1 + oStyle->GetTextYMargin();
		}

		if (oStyle->GetTextAlignmentVertical() == VAL_BOTTOM) 
		{
			lTextBottom = Y2 - oStyle->GetTextYMargin();
		}

		mp_DrawAlignedText(lTextLeft, lTextTop, lTextRight, lTextBottom, sText, (GRE_HORIZONTALALIGNMENT) oStyle->GetTextAlignmentHorizontal(), (GRE_VERTICALALIGNMENT) oStyle->GetTextAlignmentVertical(), oStyle->GetForeColor(), oStyle->GetFont(), oStyle->GetClipText());
		break;
	case SCP_OFFSETPLACEMENT:

		oFlags.SetLineAlignment((Gdiplus::StringAlignment)(oStyle->GetTextFlags()->GetVerticalAlignment() - 1));
		oFlags.SetAlignment((Gdiplus::StringAlignment)(oStyle->GetTextFlags()->GetHorizontalAlignment() -1));
		if (oStyle->GetTextFlags()->GetWordWrap() == FALSE)
		{
			oFlags.SetFormatFlags(oFlags.GetFormatFlags() | StringFormatFlagsNoWrap);
		}
		if (oStyle->GetTextFlags()->GetRightToLeft() == TRUE)
		{
			oFlags.SetFormatFlags(oFlags.GetFormatFlags() | StringFormatFlagsDirectionRightToLeft);
		}
		mp_DrawTextEx(X1 + oStyle->GetTextFlags()->GetOffsetLeft(), Y1 + oStyle->GetTextFlags()->GetOffsetTop(), X2 - oStyle->GetTextFlags()->GetOffsetRight(), Y2 - oStyle->GetTextFlags()->GetOffsetBottom(), sText, &oFlags, oStyle->GetForeColor(), oStyle->GetFont()->GetFontP(), oStyle->GetClipText());
		break;
	case SCP_EXTERIORPLACEMENT:
		if (oStyle->GetTextAlignmentHorizontal() == HAL_LEFT) 
		{
			lTextLeft = X1 - mp_lStrWidth(sText, oStyle->GetFont()) - oStyle->GetTextXMargin();
			lTextRight = X1 - oStyle->GetTextXMargin() + 1;
		}

		if (oStyle->GetTextAlignmentHorizontal() == HAL_RIGHT) 
		{
			lTextLeft = X2 + oStyle->GetTextXMargin();
			lTextRight = X2 + mp_lStrWidth(sText, oStyle->GetFont()) + oStyle->GetTextXMargin() + 1;
		}

		if (oStyle->GetTextAlignmentHorizontal() == HAL_CENTER) 
		{
			lTextLeft = X1;
			lTextRight = X2 + 1;
		}

		if (oStyle->GetTextAlignmentVertical() == VAL_TOP) 
		{
			lTextTop = Y1 - mp_lStrHeight(sText, oStyle->GetFont()) - oStyle->GetTextYMargin();
			lTextBottom = Y1 - oStyle->GetTextYMargin() + 3;
		}

		if (oStyle->GetTextAlignmentVertical() == VAL_BOTTOM) 
		{
			lTextTop = Y2 + oStyle->GetTextYMargin();
			lTextBottom = Y2 + mp_lStrHeight(sText, oStyle->GetFont()) + oStyle->GetTextYMargin() + 3;
		}

		if (oStyle->GetTextAlignmentVertical() == VAL_CENTER) 
		{
			lTextTop = Y1;
			lTextBottom = Y2 + 3;
		}
		mp_DrawAlignedText(lTextLeft, lTextTop, lTextRight, lTextBottom, sText, HAL_LEFT, VAL_TOP, oStyle->GetForeColor(), oStyle->GetFont(), oStyle->GetClipText());
		break;
	}


}

BOOL clsGraphics::mp_bPositionItem(LONG *r_lTop, LONG *r_lBottom, clsStyle *v_oStyle)
{
	if (v_oStyle->GetPlacement() == PLC_ROWEXTENTSPLACEMENT) 
	{
		return TRUE;
	}
	if (v_oStyle->GetPlacement() == PLC_OFFSETPLACEMENT) 
	{
		LONG lTop = 0;
		LONG lBottom = 0;
		lTop = *r_lTop;
		lBottom = *r_lBottom;
		switch (v_oStyle->GetFillMode()) 
		{
		case FM_COMPLETELYFILLED:
			*r_lTop = *r_lTop + v_oStyle->GetOffsetTop();
			*r_lBottom = *r_lTop + v_oStyle->GetOffsetBottom();
			break;
		case FM_UPPERHALFFILLED:
			*r_lTop = *r_lTop + v_oStyle->GetOffsetTop();
			*r_lBottom = *r_lTop + (v_oStyle->GetOffsetBottom() / 2);
			break;
		case FM_LOWERHALFFILLED:
			*r_lTop = *r_lTop + v_oStyle->GetOffsetTop() + (v_oStyle->GetOffsetBottom() / 2);
			*r_lBottom = *r_lTop + (v_oStyle->GetOffsetBottom() / 2);
			break;
		}
		if (!(*r_lTop > lTop && *r_lTop < lBottom)) 
		{
			return FALSE;
		}
		return TRUE;
	}
	return FALSE;
}

void clsGraphics::mp_DrawButton(Rect oRect, E_SCROLLBUTTONSTATE state)
{

	Color clrLightGray = Color(192, 192, 192);
	Color clrMediumGray = Color(128, 128, 128);
	Color clrDarkGray = Color(64, 64, 64);
	mp_DrawLine(oRect.X + 1, oRect.Y + 1, oRect.X + oRect.Width - 3, oRect.Y + oRect.Height - 3, LT_FILLED, clrLightGray, LDS_SOLID);

	mp_DrawLine(oRect.X, oRect.Y, oRect.X + oRect.Width - 2, oRect.Y, LT_NORMAL, Color::White, LDS_SOLID);
	mp_DrawLine(oRect.X, oRect.Y, oRect.X, oRect.Y + oRect.Height - 2, LT_NORMAL, Color::White, LDS_SOLID);
	mp_DrawLine(oRect.X, oRect.Y + oRect.Height - 1, oRect.X + oRect.Width - 1, oRect.Y + oRect.Height - 1, LT_NORMAL, clrDarkGray, LDS_SOLID);
	mp_DrawLine(oRect.X + oRect.Width - 1, oRect.Y, oRect.X + oRect.Width - 1, oRect.Y + oRect.Height - 1, LT_NORMAL, clrDarkGray, LDS_SOLID);

	mp_DrawLine(oRect.X + 1, oRect.Y + oRect.Height - 2, oRect.X + oRect.Width - 2, oRect.Y + oRect.Height - 2, LT_NORMAL, clrMediumGray, LDS_SOLID);
	mp_DrawLine(oRect.X + oRect.Width - 2, oRect.Y + 1, oRect.X + oRect.Width - 2, oRect.Y + oRect.Height - 2, LT_NORMAL, clrMediumGray, LDS_SOLID);
}

void clsGraphics::mp_DrawScrollButton(LONG X1, LONG Y1, LONG width, LONG height, E_SCROLLBUTTON button, E_SCROLLBUTTONSTATE state)
{

	Color clrLightGray = Color(192, 192, 192);
	Color clrMediumGray = Color(128, 128, 128);
	Color clrDarkGray = Color(64, 64, 64);
	LONG lOffset = 0;
	LONG lOffsetI = 0;
	Color clrArrowColor = Color::Black;
	mp_DrawLine(X1 + 2, Y1 + 2, X1 + width - 3, Y1 + height - 3, LT_FILLED, clrLightGray, LDS_SOLID);
	switch (state)
	{
	case BS_NORMAL:
	case BS_INACTIVE:
		mp_DrawLine(X1, Y1, X1 + width - 2, Y1, LT_NORMAL, clrLightGray, LDS_SOLID);
		mp_DrawLine(X1, Y1, X1, Y1 + height - 2, LT_NORMAL, clrLightGray, LDS_SOLID);
		mp_DrawLine(X1, Y1 + height - 1, X1 + width - 1, Y1 + height - 1, LT_NORMAL, clrDarkGray, LDS_SOLID);
		mp_DrawLine(X1 + width - 1, Y1, X1 + width - 1, Y1 + height - 1, LT_NORMAL, clrDarkGray, LDS_SOLID);

		mp_DrawLine(X1 + 1, Y1 + 1, X1 + width - 3, Y1 + 1, LT_NORMAL, Color::White, LDS_SOLID);
		mp_DrawLine(X1 + 1, Y1 + 1, X1 + 1, Y1 + height - 3, LT_NORMAL, Color::White, LDS_SOLID);

		mp_DrawLine(X1 + 1, Y1 + height - 2, X1 + width - 2, Y1 + height - 2, LT_NORMAL, clrMediumGray, LDS_SOLID);
		mp_DrawLine(X1 + width - 2, Y1 + 1, X1 + width - 2, Y1 + height - 2, LT_NORMAL, clrMediumGray, LDS_SOLID);
		break;
	case BS_PUSHED:
		mp_DrawLine(X1, Y1, X1 + width - 2, Y1, LT_NORMAL, clrMediumGray, LDS_SOLID);
		mp_DrawLine(X1, Y1, X1, Y1 + height - 2, LT_NORMAL, clrMediumGray, LDS_SOLID);
		mp_DrawLine(X1, Y1 + height - 1, X1 + width - 1, Y1 + height - 1, LT_NORMAL, Color::White, LDS_SOLID);
		mp_DrawLine(X1 + width - 1, Y1, X1 + width - 1, Y1 + height - 1, LT_NORMAL, Color::White, LDS_SOLID);

		mp_DrawLine(X1 + 1, Y1 + 1, X1 + width - 3, Y1 + 1, LT_NORMAL, clrDarkGray, LDS_SOLID);
		mp_DrawLine(X1 + 1, Y1 + 1, X1 + 1, Y1 + height - 3, LT_NORMAL, clrDarkGray, LDS_SOLID);

		mp_DrawLine(X1 + 1, Y1 + height - 2, X1 + width - 2, Y1 + height - 2, LT_NORMAL, clrLightGray, LDS_SOLID);
		mp_DrawLine(X1 + width - 2, Y1 + 1, X1 + width - 2, Y1 + height - 2, LT_NORMAL, clrLightGray, LDS_SOLID);
		break;
	}
	if (state == BS_PUSHED)
	{
		lOffset = 1;
	}
	if (state == BS_INACTIVE)
	{
		clrArrowColor = clrMediumGray;
		lOffsetI = 1;
	}
	switch (button)
	{
	case SBC_UP:
		if (state == BS_INACTIVE)
		{
			mp_DrawPoint(X1 + 9 + lOffsetI, Y1 + 7 + lOffsetI, Color::White);
			mp_DrawLine(X1 + 7 + lOffsetI, Y1 + 7 + lOffsetI, X1 + width - 8 + lOffsetI, Y1 + 7 + lOffsetI, LT_NORMAL, Color::White, LDS_SOLID);
			mp_DrawLine(X1 + 6 + lOffsetI, Y1 + 8 + lOffsetI, X1 + width - 7 + lOffsetI, Y1 + 8 + lOffsetI, LT_NORMAL, Color::White, LDS_SOLID);
			mp_DrawLine(X1 + 5 + lOffsetI, Y1 + 9 + lOffsetI, X1 + width - 6 + lOffsetI, Y1 + 9 + lOffsetI, LT_NORMAL, Color::White, LDS_SOLID);
		}

		mp_DrawPoint(X1 + 9 + lOffset, Y1 + 7 + lOffset, clrArrowColor);
		mp_DrawLine(X1 + 7 + lOffset, Y1 + 7 + lOffset, X1 + width - 8 + lOffset, Y1 + 7 + lOffset, LT_NORMAL, clrArrowColor, LDS_SOLID);
		mp_DrawLine(X1 + 6 + lOffset, Y1 + 8 + lOffset, X1 + width - 7 + lOffset, Y1 + 8 + lOffset, LT_NORMAL, clrArrowColor, LDS_SOLID);
		mp_DrawLine(X1 + 5 + lOffset, Y1 + 9 + lOffset, X1 + width - 6 + lOffset, Y1 + 9 + lOffset, LT_NORMAL, clrArrowColor, LDS_SOLID);
		break;
	case SBC_DOWN:
		if (state == BS_INACTIVE)
		{
			mp_DrawLine(X1 + 5 + lOffsetI, Y1 + 7 + lOffsetI, X1 + width - 6 + lOffsetI, Y1 + 7 + lOffsetI, LT_NORMAL, Color::White, LDS_SOLID);
			mp_DrawLine(X1 + 6 + lOffsetI, Y1 + 8 + lOffsetI, X1 + width - 7 + lOffsetI, Y1 + 8 + lOffsetI, LT_NORMAL, Color::White, LDS_SOLID);
			mp_DrawLine(X1 + 7 + lOffsetI, Y1 + 9 + lOffsetI, X1 + width - 8 + lOffsetI, Y1 + 9 + lOffsetI, LT_NORMAL, Color::White, LDS_SOLID);
			mp_DrawPoint(X1 + 8 + lOffsetI, Y1 + 10 + lOffsetI, Color::White);
		}

		mp_DrawLine(X1 + 5 + lOffset, Y1 + 7 + lOffset, X1 + width - 6 + lOffset, Y1 + 7 + lOffset, LT_NORMAL, clrArrowColor, LDS_SOLID);
		mp_DrawLine(X1 + 6 + lOffset, Y1 + 8 + lOffset, X1 + width - 7 + lOffset, Y1 + 8 + lOffset, LT_NORMAL, clrArrowColor, LDS_SOLID);
		mp_DrawLine(X1 + 7 + lOffset, Y1 + 9 + lOffset, X1 + width - 8 + lOffset, Y1 + 9 + lOffset, LT_NORMAL, clrArrowColor, LDS_SOLID);
		mp_DrawPoint(X1 + 8 + lOffset, Y1 + 10 + lOffset, clrArrowColor);
		break;
	case SBC_LEFT:
		if (state == BS_INACTIVE)
		{
			mp_DrawPoint(X1 + 6 + lOffsetI, Y1 + 9 + lOffsetI, Color::White);
			mp_DrawLine(X1 + 6 + lOffsetI, Y1 + 7 + lOffsetI, X1 + 6 + lOffsetI, Y1 + height - 8 + lOffsetI, LT_NORMAL, Color::White, LDS_SOLID);
			mp_DrawLine(X1 + 7 + lOffsetI, Y1 + 6 + lOffsetI, X1 + 7 + lOffsetI, Y1 + height - 7 + lOffsetI, LT_NORMAL, Color::White, LDS_SOLID);
			mp_DrawLine(X1 + 8 + lOffsetI, Y1 + 5 + lOffsetI, X1 + 8 + lOffsetI, Y1 + height - 6 + lOffsetI, LT_NORMAL, Color::White, LDS_SOLID);
		}

		mp_DrawPoint(X1 + 6 + lOffset, Y1 + 9 + lOffset, clrArrowColor);
		mp_DrawLine(X1 + 6 + lOffset, Y1 + 7 + lOffset, X1 + 6 + lOffset, Y1 + height - 8 + lOffset, LT_NORMAL, clrArrowColor, LDS_SOLID);
		mp_DrawLine(X1 + 7 + lOffset, Y1 + 6 + lOffset, X1 + 7 + lOffset, Y1 + height - 7 + lOffset, LT_NORMAL, clrArrowColor, LDS_SOLID);
		mp_DrawLine(X1 + 8 + lOffset, Y1 + 5 + lOffset, X1 + 8 + lOffset, Y1 + height - 6 + lOffset, LT_NORMAL, clrArrowColor, LDS_SOLID);
		break;
	case SBC_RIGHT:
		if (state == BS_INACTIVE)
		{
			mp_DrawLine(X1 + 7 + lOffsetI, Y1 + 5 + lOffsetI, X1 + 7 + lOffsetI, Y1 + height - 6 + lOffsetI, LT_NORMAL, Color::White, LDS_SOLID);
			mp_DrawLine(X1 + 8 + lOffsetI, Y1 + 6 + lOffsetI, X1 + 8 + lOffsetI, Y1 + height - 7 + lOffsetI, LT_NORMAL, Color::White, LDS_SOLID);
			mp_DrawLine(X1 + 9 + lOffsetI, Y1 + 7 + lOffsetI, X1 + 9 + lOffsetI, Y1 + height - 8 + lOffsetI, LT_NORMAL, Color::White, LDS_SOLID);
			mp_DrawPoint(X1 + 10 + lOffsetI, Y1 + 8 + lOffsetI, Color::White);
		}

		mp_DrawLine(X1 + 7 + lOffset, Y1 + 5 + lOffset, X1 + 7 + lOffset, Y1 + height - 6 + lOffset, LT_NORMAL, clrArrowColor, LDS_SOLID);
		mp_DrawLine(X1 + 8 + lOffset, Y1 + 6 + lOffset, X1 + 8 + lOffset, Y1 + height - 7 + lOffset, LT_NORMAL, clrArrowColor, LDS_SOLID);
		mp_DrawLine(X1 + 9 + lOffset, Y1 + 7 + lOffset, X1 + 9 + lOffset, Y1 + height - 8 + lOffset, LT_NORMAL, clrArrowColor, LDS_SOLID);
		mp_DrawPoint(X1 + 10 + lOffset, Y1 + 8 + lOffset, clrArrowColor);
		break;
	}
}

HBITMAP clsGraphics::mp_GetBitmapHandle(LPDISPATCH oImageList, LONG lIndex)
{
	HBITMAP hResult;
	LPDISPATCH oListImages;
	LPDISPATCH oListImage;
	LPDISPATCH oPicture;
	DISPID oPropertyID = NULL;
	VARIANT FAR *pVarResult = new VARIANT;
	DISPPARAMS dispparamsNoArgs = {NULL, NULL, 0, 0};
	CString propName = _T("ListImages");
	LPOLESTR sPropertyName = propName.AllocSysString();
	oImageList->GetIDsOfNames(IID_NULL, &sPropertyName , 1, LOCALE_USER_DEFAULT, &oPropertyID);
	oImageList->Invoke(oPropertyID, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYGET, &dispparamsNoArgs, pVarResult, NULL, NULL);
	oListImages = pVarResult->pdispVal;
	propName = _T("Item");
	sPropertyName = propName.AllocSysString();
	DISPPARAMS dispparamsitem;
	dispparamsitem.rgvarg = new VARIANT[1]; 
	dispparamsitem.rgvarg[0].vt = VT_I2;
	dispparamsitem.rgvarg[0].intVal = lIndex;
	dispparamsitem.cArgs = 1;
	dispparamsitem.cNamedArgs = 0;
	dispparamsitem.rgdispidNamedArgs = 0; 
	oListImages->GetIDsOfNames(IID_NULL, &sPropertyName , 1, LOCALE_USER_DEFAULT, &oPropertyID);
	oListImages->Invoke(oPropertyID, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYGET, &dispparamsitem, pVarResult, NULL, NULL);
	oListImage = pVarResult->pdispVal;
	propName = _T("Picture");
	sPropertyName = propName.AllocSysString();
	oListImage->GetIDsOfNames(IID_NULL, &sPropertyName , 1, LOCALE_USER_DEFAULT, &oPropertyID);
	oListImage->Invoke(oPropertyID, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYGET, &dispparamsNoArgs, pVarResult, NULL, NULL);
	oPicture = pVarResult->pdispVal;
	propName = _T("Handle");
	sPropertyName = propName.AllocSysString();
	oPicture->GetIDsOfNames(IID_NULL, &sPropertyName , 1, LOCALE_USER_DEFAULT, &oPropertyID);
	oPicture->Invoke(oPropertyID, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYGET, &dispparamsNoArgs, pVarResult, NULL, NULL);
	hResult = (HBITMAP) pVarResult->lVal;
	delete(pVarResult);
	pVarResult = NULL;
	delete(dispparamsitem.rgvarg);
	dispparamsitem.rgvarg = NULL;
	return hResult;
}

LONG clsGraphics::mp_GetImageListCount(LPDISPATCH oImageList)
{
	LONG lResult;
	LPDISPATCH oListImages;
	DISPID oPropertyID = NULL;
	VARIANT FAR *pVarResult = new VARIANT;
	DISPPARAMS dispparamsNoArgs = {NULL, NULL, 0, 0};
	CString propName = _T("ListImages");
	LPOLESTR sPropertyName = propName.AllocSysString();
	oImageList->GetIDsOfNames(IID_NULL, &sPropertyName , 1, LOCALE_USER_DEFAULT, &oPropertyID);
	oImageList->Invoke(oPropertyID, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYGET, &dispparamsNoArgs, pVarResult, NULL, NULL);
	oListImages = pVarResult->pdispVal;
	propName = _T("Count");
	sPropertyName = propName.AllocSysString();
	oListImages->GetIDsOfNames(IID_NULL, &sPropertyName , 1, LOCALE_USER_DEFAULT, &oPropertyID);
	oListImages->Invoke(oPropertyID, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYGET, &dispparamsNoArgs, pVarResult, NULL, NULL);
	lResult = pVarResult->intVal;
	delete(pVarResult);
	pVarResult = NULL;
	return lResult;
}

Color clsGraphics::mp_ConvGDICol(OLE_COLOR lColor)
{
	Color oColor(lColor);
	return Color();
}

//void clsGraphics::DrawPoint(LONG X, LONG Y, Color clrColor)
//{
//	Gdiplus::Pen oPen(clrColor, 1);
//	oGraphics()->DrawLine(&oPen, (REAL)(X-1), (REAL)(Y-1), (REAL)(X), (REAL)(Y));
//}

LONG clsGraphics::Getf_FocusLeft(void)
{
	return mp_lFocusLeft;
}

void clsGraphics::Setf_FocusLeft(LONG Value)
{
	mp_lFocusLeft = Value;
}

LONG clsGraphics::Getf_FocusTop(void)
{
	return mp_lFocusTop;
}

void clsGraphics::Setf_FocusTop(LONG Value)
{
	mp_lFocusTop = Value;
}

LONG clsGraphics::Getf_FocusRight(void)
{
	return mp_lFocusRight;
}

void clsGraphics::Setf_FocusRight(LONG Value)
{
	mp_lFocusRight = Value;
}

LONG clsGraphics::Getf_FocusBottom(void)
{
	return mp_lFocusBottom;
}

void clsGraphics::Setf_FocusBottom(LONG Value)
{
	mp_lFocusBottom = Value;
}

void clsGraphics::SetCustomDC(Graphics* oGraphics)
{
	mp_lCustomDC = oGraphics;
}

void clsGraphics::mp_DrawPoint(int X, int Y, Gdiplus::Color clrColor)
{
	Gdiplus::SolidBrush mp_ucBrush(clrColor);
    oGraphics()->FillRectangle(&mp_ucBrush, X, Y, 1, 1);
}

void clsGraphics::mp_DrawArrow(int X, int Y, GRE_ARROWDIRECTION yArrowDirection, int lArrowSize, Gdiplus::Color clrColor)
{
    int i = 0;
    switch (yArrowDirection)
    {
        case AWD_LEFT:
            mp_DrawPoint(X, Y, clrColor);
            for (i = 1; i <= lArrowSize; i++)
            {
                X = X + 1;
                mp_DrawLine(X, Y - i, X, Y + i, LT_NORMAL, clrColor, LDS_SOLID);
            }
            break;
        case AWD_RIGHT:
            mp_DrawPoint(X, Y, clrColor);
            for (i = 1; i <= lArrowSize; i++)
            {
                X = X - 1;
                mp_DrawLine(X, Y - i, X, Y + i, LT_NORMAL, clrColor, LDS_SOLID);
            }
            break;
        case AWD_UP:
            mp_DrawPoint(X, Y, clrColor);
            for (i = 1; i <= lArrowSize; i++)
            {
                Y = Y + 1;
                mp_DrawLine(X - i, Y, X + i, Y, LT_NORMAL, clrColor, LDS_SOLID);
            }
            break;
        case AWD_DOWN:
            mp_DrawPoint(X, Y, clrColor);
            for (i = 1; i <= lArrowSize; i++)
            {
                Y = Y - 1;
                mp_DrawLine(X - i, Y, X + i, Y, LT_NORMAL, clrColor, LDS_SOLID);
            }
            break;
    }
}

LONG clsGraphics::mp_lStrWidth(CString sString,clsFont* r_oFont)
{
	Gdiplus::RectF layoutRect(0, 0, 0, 0);
	Gdiplus::RectF boundRect;
	oGraphics()->MeasureString(sString, sString.GetLength(), r_oFont->GetFontP(), layoutRect, &boundRect);
	return CInt32(boundRect.GetRight() - boundRect.GetLeft());
}

LONG clsGraphics::mp_lStrWidth(CString sString, Font* r_oFont)
{
	Gdiplus::RectF layoutRect(0, 0, 0, 0);
	Gdiplus::RectF boundRect;
	oGraphics()->MeasureString(sString, sString.GetLength(), r_oFont, layoutRect, &boundRect);
	return CInt32(boundRect.GetRight() - boundRect.GetLeft());
}

LONG clsGraphics::mp_lStrHeight(CString sString,clsFont* r_oFont)
{
	Gdiplus::RectF layoutRect(0, 0, 0, 0);
	Gdiplus::RectF boundRect;
	oGraphics()->MeasureString(sString, sString.GetLength(), r_oFont->GetFontP(), layoutRect, &boundRect);
	return CInt32(boundRect.GetBottom() - boundRect.GetTop());
}

LONG clsGraphics::mp_lStrHeight(CString sString, Font* r_oFont)
{
	Gdiplus::RectF layoutRect(0, 0, 0, 0);
	Gdiplus::RectF boundRect;
	oGraphics()->MeasureString(sString, sString.GetLength(), r_oFont, layoutRect, &boundRect);
	return CInt32(boundRect.GetBottom() - boundRect.GetTop());
}

void clsGraphics::mp_DrawExpandIcon(LONG X, LONG Y, Gdiplus::Color oColor, Gdiplus::Color oDropShadowColor, Gdiplus::Color oBackgroundColor)
{
    oGraphics_DrawLine(oColor, X + 1, Y, X + 1, Y + 8);
    oGraphics_DrawLine(oColor, X + 1, Y, X + 5, Y + 4);
    oGraphics_DrawLine(oColor, X + 1, Y + 8, X + 5, Y + 4);
    oGraphics_DrawLine(oDropShadowColor, X + 2, Y, X + 6, Y + 4);
    oGraphics_DrawLine(oDropShadowColor, X + 2, Y + 8, X + 6, Y + 4);
    oGraphics_DrawLine(oBackgroundColor, X + 2, Y + 2, X + 2, Y + 6);
    oGraphics_DrawLine(oBackgroundColor, X + 3, Y + 3, X + 3, Y + 5);
    oGraphics_FillRectangle(oBackgroundColor, X + 4, Y + 4, 1, 1);
}

void clsGraphics::mp_DrawCollapseIcon(LONG X, LONG Y, Gdiplus::Color oColor, Gdiplus::Color oDropShadowColor)
{
	oGraphics_DrawLine(oColor, X + 5, Y + 2, X + 6, Y + 2);
    oGraphics_DrawLine(oColor, X + 4, Y + 3, X + 6, Y + 3);
    oGraphics_DrawLine(oColor, X + 3, Y + 4, X + 6, Y + 4);
    oGraphics_DrawLine(oColor, X + 2, Y + 5, X + 6, Y + 5);
    oGraphics_DrawLine(oColor, X + 1, Y + 6, X + 6, Y + 6);
    oGraphics_DrawLine(oDropShadowColor, X + 6, Y + 1, X + 1, Y + 6);
}

void clsGraphics::oGraphics_DrawLine(Gdiplus::Color oColor, LONG X1, LONG Y1, LONG X2, LONG Y2)
{
	Gdiplus::Pen oPen(oColor);
	Gdiplus::Point oPoint1(X1, Y1);
	Gdiplus::Point oPoint2(X2, Y2);
	oGraphics()->DrawLine(&oPen, oPoint1, oPoint2);
}

void clsGraphics::oGraphics_FillRectangle(Gdiplus::Color oColor, LONG X, LONG Y, LONG Width, LONG Height)
{
	Gdiplus::SolidBrush mp_ucBrush(oColor);
	oGraphics()->FillRectangle(&mp_ucBrush, X, Y, 1, 1);
}