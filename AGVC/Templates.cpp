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
#include "Templates.h"
#include "ActiveGanttVCCtl.h"


void g_ApplyTemplate(CActiveGanttVCCtl* oControl, LONG yControlTemplate, LONG yObjectTemplate)
{
	oControl->GetStyles()->Clear();
	mp_ControlTemplateSelector(oControl, yControlTemplate);
	mp_ObjectTemplateSelector(oControl, yObjectTemplate);
}

void mp_ControlTemplateSelector(CActiveGanttVCCtl* oControl, LONG yControlTemplate)
{
	switch (yControlTemplate) 
	{
	case STC_CH_SOLID_WHITE:
		{
			ControlTemplateSolid oTemplate;
			oTemplate.ControlBorderColor = Color::Black;
			oTemplate.HeadingBorderColor = Color::Black;
			oTemplate.HeadingForeColor = Color::Black;
			oTemplate.RowForeColor = Color::Black;
			oTemplate.HeadingBackColor = Color::White;
			oTemplate.TreeviewColor = Color::Black;
			oTemplate.RowBorderColor = Color::MakeARGB(255, 192, 192, 192);
			oTemplate.TimeBlockBackColor = Color::MakeARGB(255, 192, 192, 192);

			mp_ApplyControlTemplate_Solid(oControl, oTemplate);
		}

		break;
	case STC_CH_SOLID_DARK_BLUE:
		{
			ControlTemplateSolid oTemplate;
			oTemplate.ControlBorderColor = Color::Black;
			oTemplate.HeadingBorderColor = Color::White;
			oTemplate.HeadingForeColor = Color::White;
			oTemplate.RowForeColor = Color::Black;
			oTemplate.HeadingBackColor = Color::MakeARGB(255, 63, 68, 90);
			oTemplate.TreeviewColor = Color::Black;
			oTemplate.RowBorderColor = Color::MakeARGB(255, 192, 192, 192);
			oTemplate.TimeBlockBackColor = Color::MakeARGB(255, 115, 119, 135);

			mp_ApplyControlTemplate_Solid(oControl, oTemplate);
		}

		break;
	case STC_CH_SOLID_VIOLET:
		{

			ControlTemplateSolid oTemplate;
			oTemplate.ControlBorderColor = Color::MakeARGB(255, 93, 71, 139);
			oTemplate.HeadingBorderColor = Color::White;
			oTemplate.HeadingForeColor = Color::White;
			oTemplate.RowForeColor = Color::Black;
			oTemplate.HeadingBackColor = Color::MakeARGB(255, 93, 71, 139);
			oTemplate.TreeviewColor = Color::Black;
			oTemplate.RowBorderColor = Color::MakeARGB(255, 192, 192, 192);
			oTemplate.TimeBlockBackColor = Color::MakeARGB(255, 140, 124, 173);

			mp_ApplyControlTemplate_Solid(oControl, oTemplate);
		}

		break;
	case STC_CH_SOLID_GREEN:
		{
			ControlTemplateSolid oTemplate;
			oTemplate.ControlBorderColor = Color::Black;
			oTemplate.HeadingBorderColor = Color::White;
			oTemplate.HeadingForeColor = Color::White;
			oTemplate.RowForeColor = Color::Black;
			oTemplate.HeadingBackColor = Color::MakeARGB(255, 105, 152, 105);
			oTemplate.TreeviewColor = Color::MakeARGB(255, 105, 152, 105);
			oTemplate.RowBorderColor = Color::MakeARGB(255, 192, 192, 192);
			oTemplate.TimeBlockBackColor = Color::MakeARGB(255, 177, 201, 177);

			mp_ApplyControlTemplate_Solid(oControl, oTemplate);
		}

		break;
	case STC_CH_SOLID_RED:
		{
			ControlTemplateSolid oTemplate;
			oTemplate.ControlBorderColor = Color::MakeARGB(255, 180, 2, 52);
			oTemplate.HeadingBorderColor = Color::White;
			oTemplate.HeadingForeColor = Color::White;
			oTemplate.RowForeColor = Color::Black;
			oTemplate.HeadingBackColor = Color::MakeARGB(255, 180, 2, 52);
			oTemplate.TreeviewColor = Color::MakeARGB(255, 180, 2, 52);
			oTemplate.RowBorderColor = Color::MakeARGB(255, 192, 192, 192);
			oTemplate.TimeBlockBackColor = Color::MakeARGB(255, 202, 77, 113);

			mp_ApplyControlTemplate_Solid(oControl, oTemplate);
		}

		break;
	case STC_CH_SOLID_LIGHT_BLUE:
		{
			ControlTemplateSolid oTemplate;
			oTemplate.ControlBorderColor = Color::MakeARGB(255, 153, 153, 221);
			oTemplate.HeadingBorderColor = Color::White;
			oTemplate.HeadingForeColor = Color::White;
			oTemplate.RowForeColor = Color::Black;
			oTemplate.HeadingBackColor = Color::MakeARGB(255, 153, 153, 221);
			oTemplate.TreeviewColor = Color::MakeARGB(255, 153, 153, 221);
			oTemplate.RowBorderColor = Color::MakeARGB(255, 192, 192, 192);
			oTemplate.TimeBlockBackColor = Color::MakeARGB(255, 204, 204, 255);

			mp_ApplyControlTemplate_Solid(oControl, oTemplate);
		}

		break;
	case STC_CH_SOLID_GREY:
		{
			ControlTemplateSolid oTemplate;
			oTemplate.ControlBorderColor = Color::Black;
			oTemplate.HeadingBorderColor = Color::White;
			oTemplate.HeadingForeColor = Color::White;
			oTemplate.RowForeColor = Color::Black;
			oTemplate.HeadingBackColor = Color::MakeARGB(255, 142, 142, 142);
			oTemplate.TreeviewColor = Color::Black;
			oTemplate.RowBorderColor = Color::MakeARGB(255, 192, 192, 192);
			oTemplate.TimeBlockBackColor = Color::MakeARGB(255, 172, 172, 172);

			mp_ApplyControlTemplate_Solid(oControl, oTemplate);
		}

		break;
	case STC_CH_SOLID_LIGHT_STEEL_BLUE:
		{
			ControlTemplateSolid oTemplate;
			oTemplate.ControlBorderColor = Color::MakeARGB(255, 162, 181, 205);
			oTemplate.HeadingBorderColor = Color::Black;
			oTemplate.HeadingForeColor = Color::Black;
			oTemplate.RowForeColor = Color::Black;
			oTemplate.HeadingBackColor = Color::MakeARGB(255, 162, 181, 205);
			oTemplate.TreeviewColor = Color::MakeARGB(255, 162, 181, 205);
			oTemplate.RowBorderColor = Color::MakeARGB(255, 192, 192, 192);
			oTemplate.TimeBlockBackColor = Color::MakeARGB(255, 187, 201, 219);

			mp_ApplyControlTemplate_Solid(oControl, oTemplate);
		}

		break;
	case STC_CH_VGRAD_YELLOW:
		{
			ControlTemplateGradient oTemplate;
			oTemplate.ControlBorderColor = Color::Black;
			oTemplate.HeadingBorderColor = Color::Black;
			oTemplate.HeadingForeColor = Color::Black;
			oTemplate.RowForeColor = Color::Black;
			oTemplate.GradientFillMode = GDT_VERTICAL;
			oTemplate.StartGradientColor = Color::MakeARGB(255, 209, 164, 2);
			oTemplate.EndGradientColor = Color::MakeARGB(255, 229, 203, 5);
			oTemplate.TreeviewColor = Color::MakeARGB(255, 189, 167, 4);
			oTemplate.RowBorderColor = Color::MakeARGB(255, 192, 192, 192);

			mp_ApplyControlTemplate_Gradient(oControl, oTemplate);
		}

		break;
	case STC_CH_VGRAD_ORANGE:
		{
			ControlTemplateGradient oTemplate;
			oTemplate.ControlBorderColor = Color::Black;
			oTemplate.HeadingBorderColor = Color::MakeARGB(255, 197, 206, 216);
			oTemplate.HeadingForeColor = Color::Black;
			oTemplate.RowForeColor = Color::Black;
			oTemplate.GradientFillMode = GDT_VERTICAL;
			oTemplate.StartGradientColor = Color::MakeARGB(255, 202, 116, 38);
			oTemplate.EndGradientColor = Color::MakeARGB(255, 201, 109, 36);
			oTemplate.TreeviewColor = Color::MakeARGB(255, 202, 116, 38);
			oTemplate.RowBorderColor = Color::MakeARGB(255, 192, 192, 192);

			mp_ApplyControlTemplate_Gradient(oControl, oTemplate);
		}

		break;
	case STC_CH_VGRAD_RED:
		{
			ControlTemplateGradient oTemplate;
			oTemplate.ControlBorderColor = Color::MakeARGB(255, 142, 2, 37);
			oTemplate.HeadingBorderColor = Color::MakeARGB(255, 197, 206, 216);
			oTemplate.HeadingForeColor = Color::MakeARGB(255, 230, 230, 230);
			oTemplate.RowForeColor = Color::Black;
			oTemplate.GradientFillMode = GDT_VERTICAL;
			oTemplate.StartGradientColor = Color::MakeARGB(255, 142, 2, 37);
			oTemplate.EndGradientColor = Color::MakeARGB(255, 180, 2, 52);
			oTemplate.TreeviewColor = Color::MakeARGB(255, 142, 2, 37);
			oTemplate.RowBorderColor = Color::MakeARGB(255, 192, 192, 192);

			mp_ApplyControlTemplate_Gradient(oControl, oTemplate);
		}

		break;
	case STC_CH_VGRAD_CRIMSON:
		{
			ControlTemplateGradient oTemplate;
			oTemplate.ControlBorderColor = Color::MakeARGB(255, 157, 3, 57);
			oTemplate.HeadingBorderColor = Color::MakeARGB(255, 197, 206, 216);
			oTemplate.HeadingForeColor = Color::MakeARGB(255, 230, 230, 230);
			oTemplate.RowForeColor = Color::Black;
			oTemplate.GradientFillMode = GDT_VERTICAL;
			oTemplate.StartGradientColor = Color::MakeARGB(255, 157, 3, 57);
			oTemplate.EndGradientColor = Color::MakeARGB(255, 206, 2, 73);
			oTemplate.TreeviewColor = Color::MakeARGB(255, 157, 3, 57);
			oTemplate.RowBorderColor = Color::MakeARGB(255, 192, 192, 192);

			mp_ApplyControlTemplate_Gradient(oControl, oTemplate);
		}

		break;
	case STC_CH_VGRAD_MAGENTA:
		{
			ControlTemplateGradient oTemplate;
			oTemplate.ControlBorderColor = Color::MakeARGB(255, 171, 4, 96);
			oTemplate.HeadingBorderColor = Color::MakeARGB(255, 197, 206, 216);
			oTemplate.HeadingForeColor = Color::Black;
			oTemplate.RowForeColor = Color::Black;
			oTemplate.GradientFillMode = GDT_VERTICAL;
			oTemplate.StartGradientColor = Color::MakeARGB(255, 171, 4, 96);
			oTemplate.EndGradientColor = Color::MakeARGB(255, 199, 3, 111);
			oTemplate.TreeviewColor = Color::MakeARGB(255, 171, 4, 96);
			oTemplate.RowBorderColor = Color::MakeARGB(255, 192, 192, 192);

			mp_ApplyControlTemplate_Gradient(oControl, oTemplate);
		}

		break;
	case STC_CH_VGRAD_MULBERRY:
		{
			ControlTemplateGradient oTemplate;
			oTemplate.ControlBorderColor = Color::MakeARGB(255, 156, 2, 124);
			oTemplate.HeadingBorderColor = Color::MakeARGB(255, 197, 206, 216);
			oTemplate.HeadingForeColor = Color::Black;
			oTemplate.RowForeColor = Color::Black;
			oTemplate.GradientFillMode = GDT_VERTICAL;
			oTemplate.StartGradientColor = Color::MakeARGB(255, 156, 2, 124);
			oTemplate.EndGradientColor = Color::MakeARGB(255, 191, 3, 150);
			oTemplate.TreeviewColor = Color::MakeARGB(255, 156, 2, 124);
			oTemplate.RowBorderColor = Color::MakeARGB(255, 192, 192, 192);

			mp_ApplyControlTemplate_Gradient(oControl, oTemplate);
		}

		break;
	case STC_CH_VGRAD_BLUE_VIOLET:
		{
			ControlTemplateGradient oTemplate;
			oTemplate.ControlBorderColor = Color::MakeARGB(255, 100, 145, 204);
			oTemplate.HeadingBorderColor = Color::MakeARGB(255, 197, 206, 216);
			oTemplate.HeadingForeColor = Color::White;
			oTemplate.RowForeColor = Color::Black;
			oTemplate.GradientFillMode = GDT_VERTICAL;
			oTemplate.StartGradientColor = Color::MakeARGB(255, 112, 52, 197);
			oTemplate.EndGradientColor = Color::MakeARGB(255, 142, 63, 217);
			oTemplate.TreeviewColor = Color::MakeARGB(255, 100, 145, 204);
			oTemplate.RowBorderColor = Color::MakeARGB(255, 192, 192, 192);

			mp_ApplyControlTemplate_Gradient(oControl, oTemplate);
		}

		break;
	case STC_CH_VGRAD_ANAKIWA_BLUE:
		{
			ControlTemplateGradient oTemplate;
			oTemplate.ControlBorderColor = Color::MakeARGB(255, 100, 145, 204);
			oTemplate.HeadingBorderColor = Color::MakeARGB(255, 197, 206, 216);
			oTemplate.HeadingForeColor = Color::Black;
			oTemplate.RowForeColor = Color::Black;
			oTemplate.GradientFillMode = GDT_VERTICAL;
			oTemplate.StartGradientColor = Color::MakeARGB(255, 179, 206, 235);
			oTemplate.EndGradientColor = Color::MakeARGB(255, 161, 193, 232);
			oTemplate.TreeviewColor = Color::MakeARGB(255, 100, 145, 204);
			oTemplate.RowBorderColor = Color::MakeARGB(255, 192, 192, 192);

			mp_ApplyControlTemplate_Gradient(oControl, oTemplate);
		}

		break;
	case STC_CH_VGRAD_BLUE_BELL:
		{
			ControlTemplateGradient oTemplate;
			oTemplate.ControlBorderColor = Color::MakeARGB(255, 100, 145, 204);
			oTemplate.HeadingBorderColor = Color::MakeARGB(255, 197, 206, 216);
			oTemplate.HeadingForeColor = Color::Black;
			oTemplate.RowForeColor = Color::Black;
			oTemplate.GradientFillMode = GDT_VERTICAL;
			oTemplate.StartGradientColor = Color::MakeARGB(255, 150, 166, 191);
			oTemplate.EndGradientColor = Color::MakeARGB(255, 177, 198, 227);
			oTemplate.TreeviewColor = Color::MakeARGB(255, 100, 145, 204);
			oTemplate.RowBorderColor = Color::MakeARGB(255, 192, 192, 192);

			mp_ApplyControlTemplate_Gradient(oControl, oTemplate);
		}

		break;
	case STC_CH_VGRAD_BLUE:
		{
			ControlTemplateGradient oTemplate;
			oTemplate.ControlBorderColor = Color::MakeARGB(255, 100, 145, 204);
			oTemplate.HeadingBorderColor = Color::MakeARGB(255, 197, 206, 216);
			oTemplate.HeadingForeColor = Color::White;
			oTemplate.RowForeColor = Color::Black;
			oTemplate.GradientFillMode = GDT_VERTICAL;
			oTemplate.StartGradientColor = Color::MakeARGB(255, 0, 102, 153);
			oTemplate.EndGradientColor = Color::MakeARGB(255, 3, 167, 208);
			oTemplate.TreeviewColor = Color::MakeARGB(255, 100, 145, 204);
			oTemplate.RowBorderColor = Color::MakeARGB(255, 192, 192, 192);

			mp_ApplyControlTemplate_Gradient(oControl, oTemplate);
		}

		break;
	case STC_CH_VGRAD_AERO:
		{
			ControlTemplateGradient oTemplate;
			oTemplate.ControlBorderColor = Color::MakeARGB(255, 100, 145, 204);
			oTemplate.HeadingBorderColor = Color::MakeARGB(255, 197, 206, 216);
			oTemplate.HeadingForeColor = Color::Black;
			oTemplate.RowForeColor = Color::Black;
			oTemplate.GradientFillMode = GDT_VERTICAL;
			oTemplate.StartGradientColor = Color::MakeARGB(255, 2, 157, 177);
			oTemplate.EndGradientColor = Color::MakeARGB(255, 3, 199, 211);
			oTemplate.TreeviewColor = Color::MakeARGB(255, 100, 145, 204);
			oTemplate.RowBorderColor = Color::MakeARGB(255, 192, 192, 192);

			mp_ApplyControlTemplate_Gradient(oControl, oTemplate);
		}

		break;
	case STC_CH_HGRAD_YELLOW:
		{
			ControlTemplateGradient oTemplate;
			oTemplate.ControlBorderColor = Color::Black;
			oTemplate.HeadingBorderColor = Color::Black;
			oTemplate.HeadingForeColor = Color::Black;
			oTemplate.RowForeColor = Color::Black;
			oTemplate.GradientFillMode = GDT_HORIZONTAL;
			oTemplate.StartGradientColor = Color::MakeARGB(255, 209, 164, 2);
			oTemplate.EndGradientColor = Color::MakeARGB(255, 229, 203, 5);
			oTemplate.TreeviewColor = Color::MakeARGB(255, 189, 167, 4);
			oTemplate.RowBorderColor = Color::MakeARGB(255, 192, 192, 192);

			mp_ApplyControlTemplate_Gradient(oControl, oTemplate);
		}

		break;
	case STC_CH_HGRAD_ORANGE:
		{
			ControlTemplateGradient oTemplate;
			oTemplate.ControlBorderColor = Color::Black;
			oTemplate.HeadingBorderColor = Color::MakeARGB(255, 197, 206, 216);
			oTemplate.HeadingForeColor = Color::Black;
			oTemplate.RowForeColor = Color::Black;
			oTemplate.GradientFillMode = GDT_HORIZONTAL;
			oTemplate.StartGradientColor = Color::MakeARGB(255, 202, 116, 38);
			oTemplate.EndGradientColor = Color::MakeARGB(255, 201, 109, 36);
			oTemplate.TreeviewColor = Color::MakeARGB(255, 202, 116, 38);
			oTemplate.RowBorderColor = Color::MakeARGB(255, 192, 192, 192);

			mp_ApplyControlTemplate_Gradient(oControl, oTemplate);
		}

		break;
	case STC_CH_HGRAD_RED:
		{
			ControlTemplateGradient oTemplate;
			oTemplate.ControlBorderColor = Color::MakeARGB(255, 142, 2, 37);
			oTemplate.HeadingBorderColor = Color::MakeARGB(255, 197, 206, 216);
			oTemplate.HeadingForeColor = Color::MakeARGB(255, 230, 230, 230);
			oTemplate.RowForeColor = Color::Black;
			oTemplate.GradientFillMode = GDT_HORIZONTAL;
			oTemplate.StartGradientColor = Color::MakeARGB(255, 142, 2, 37);
			oTemplate.EndGradientColor = Color::MakeARGB(255, 180, 2, 52);
			oTemplate.TreeviewColor = Color::MakeARGB(255, 142, 2, 37);
			oTemplate.RowBorderColor = Color::MakeARGB(255, 192, 192, 192);

			mp_ApplyControlTemplate_Gradient(oControl, oTemplate);
		}

		break;
	case STC_CH_HGRAD_CRIMSON:
		{
			ControlTemplateGradient oTemplate;
			oTemplate.ControlBorderColor = Color::MakeARGB(255, 157, 3, 57);
			oTemplate.HeadingBorderColor = Color::MakeARGB(255, 197, 206, 216);
			oTemplate.HeadingForeColor = Color::MakeARGB(255, 230, 230, 230);
			oTemplate.RowForeColor = Color::Black;
			oTemplate.GradientFillMode = GDT_HORIZONTAL;
			oTemplate.StartGradientColor = Color::MakeARGB(255, 157, 3, 57);
			oTemplate.EndGradientColor = Color::MakeARGB(255, 206, 2, 73);
			oTemplate.TreeviewColor = Color::MakeARGB(255, 157, 3, 57);
			oTemplate.RowBorderColor = Color::MakeARGB(255, 192, 192, 192);

			mp_ApplyControlTemplate_Gradient(oControl, oTemplate);
		}

		break;
	case STC_CH_HGRAD_MAGENTA:
		{
			ControlTemplateGradient oTemplate;
			oTemplate.ControlBorderColor = Color::MakeARGB(255, 171, 4, 96);
			oTemplate.HeadingBorderColor = Color::MakeARGB(255, 197, 206, 216);
			oTemplate.HeadingForeColor = Color::Black;
			oTemplate.RowForeColor = Color::Black;
			oTemplate.GradientFillMode = GDT_HORIZONTAL;
			oTemplate.StartGradientColor = Color::MakeARGB(255, 171, 4, 96);
			oTemplate.EndGradientColor = Color::MakeARGB(255, 199, 3, 111);
			oTemplate.TreeviewColor = Color::MakeARGB(255, 171, 4, 96);
			oTemplate.RowBorderColor = Color::MakeARGB(255, 192, 192, 192);

			mp_ApplyControlTemplate_Gradient(oControl, oTemplate);
		}

		break;
	case STC_CH_HGRAD_MULBERRY:
		{
			ControlTemplateGradient oTemplate;
			oTemplate.ControlBorderColor = Color::MakeARGB(255, 156, 2, 124);
			oTemplate.HeadingBorderColor = Color::MakeARGB(255, 197, 206, 216);
			oTemplate.HeadingForeColor = Color::Black;
			oTemplate.RowForeColor = Color::Black;
			oTemplate.GradientFillMode = GDT_HORIZONTAL;
			oTemplate.StartGradientColor = Color::MakeARGB(255, 156, 2, 124);
			oTemplate.EndGradientColor = Color::MakeARGB(255, 191, 3, 150);
			oTemplate.TreeviewColor = Color::MakeARGB(255, 156, 2, 124);
			oTemplate.RowBorderColor = Color::MakeARGB(255, 192, 192, 192);

			mp_ApplyControlTemplate_Gradient(oControl, oTemplate);
		}

		break;
	case STC_CH_HGRAD_BLUE_VIOLET:
		{
			ControlTemplateGradient oTemplate;
			oTemplate.ControlBorderColor = Color::MakeARGB(255, 100, 145, 204);
			oTemplate.HeadingBorderColor = Color::MakeARGB(255, 197, 206, 216);
			oTemplate.HeadingForeColor = Color::White;
			oTemplate.RowForeColor = Color::Black;
			oTemplate.GradientFillMode = GDT_HORIZONTAL;
			oTemplate.StartGradientColor = Color::MakeARGB(255, 112, 52, 197);
			oTemplate.EndGradientColor = Color::MakeARGB(255, 142, 63, 217);
			oTemplate.TreeviewColor = Color::MakeARGB(255, 100, 145, 204);
			oTemplate.RowBorderColor = Color::MakeARGB(255, 192, 192, 192);

			mp_ApplyControlTemplate_Gradient(oControl, oTemplate);
		}

		break;
	case STC_CH_HGRAD_ANAKIWA_BLUE:
		{
			ControlTemplateGradient oTemplate;
			oTemplate.ControlBorderColor = Color::MakeARGB(255, 100, 145, 204);
			oTemplate.HeadingBorderColor = Color::MakeARGB(255, 197, 206, 216);
			oTemplate.HeadingForeColor = Color::Black;
			oTemplate.RowForeColor = Color::Black;
			oTemplate.GradientFillMode = GDT_HORIZONTAL;
			oTemplate.StartGradientColor = Color::MakeARGB(255, 179, 206, 235);
			oTemplate.EndGradientColor = Color::MakeARGB(255, 161, 193, 232);
			oTemplate.TreeviewColor = Color::MakeARGB(255, 100, 145, 204);
			oTemplate.RowBorderColor = Color::MakeARGB(255, 192, 192, 192);

			mp_ApplyControlTemplate_Gradient(oControl, oTemplate);
		}

		break;
	case STC_CH_HGRAD_BLUE_BELL:
		{
			ControlTemplateGradient oTemplate;
			oTemplate.ControlBorderColor = Color::MakeARGB(255, 100, 145, 204);
			oTemplate.HeadingBorderColor = Color::MakeARGB(255, 197, 206, 216);
			oTemplate.HeadingForeColor = Color::Black;
			oTemplate.RowForeColor = Color::Black;
			oTemplate.GradientFillMode = GDT_HORIZONTAL;
			oTemplate.StartGradientColor = Color::MakeARGB(255, 150, 166, 191);
			oTemplate.EndGradientColor = Color::MakeARGB(255, 177, 198, 227);
			oTemplate.TreeviewColor = Color::MakeARGB(255, 100, 145, 204);
			oTemplate.RowBorderColor = Color::MakeARGB(255, 192, 192, 192);

			mp_ApplyControlTemplate_Gradient(oControl, oTemplate);
		}
		break;
	case STC_CH_HGRAD_BLUE:
		{
			ControlTemplateGradient oTemplate;
			oTemplate.ControlBorderColor = Color::MakeARGB(255, 100, 145, 204);
			oTemplate.HeadingBorderColor = Color::MakeARGB(255, 197, 206, 216);
			oTemplate.HeadingForeColor = Color::White;
			oTemplate.RowForeColor = Color::Black;
			oTemplate.GradientFillMode = GDT_HORIZONTAL;
			oTemplate.StartGradientColor = Color::MakeARGB(255, 0, 102, 153);
			oTemplate.EndGradientColor = Color::MakeARGB(255, 3, 167, 208);
			oTemplate.TreeviewColor = Color::MakeARGB(255, 100, 145, 204);
			oTemplate.RowBorderColor = Color::MakeARGB(255, 192, 192, 192);

			mp_ApplyControlTemplate_Gradient(oControl, oTemplate);
		}

		break;
	case STC_CH_HGRAD_AERO:
		{
			ControlTemplateGradient oTemplate;
			oTemplate.ControlBorderColor = Color::MakeARGB(255, 100, 145, 204);
			oTemplate.HeadingBorderColor = Color::MakeARGB(255, 197, 206, 216);
			oTemplate.HeadingForeColor = Color::Black;
			oTemplate.RowForeColor = Color::Black;
			oTemplate.GradientFillMode = GDT_HORIZONTAL;
			oTemplate.StartGradientColor = Color::MakeARGB(255, 2, 157, 177);
			oTemplate.EndGradientColor = Color::MakeARGB(255, 3, 199, 211);
			oTemplate.TreeviewColor = Color::MakeARGB(255, 100, 145, 204);
			oTemplate.RowBorderColor = Color::MakeARGB(255, 192, 192, 192);

			mp_ApplyControlTemplate_Gradient(oControl, oTemplate);
		}

		break;
	}
}

void mp_ObjectTemplateSelector(CActiveGanttVCCtl* oControl, LONG yObjectTemplate)
{
	switch (yObjectTemplate) 
	{
	case STO_BW_HATCH:
		{

			S_ObjectTemplate oTemplate;

			oTemplate.yBackgroundMode_Tasks = FP_HATCH;

			oTemplate.BorderColor_T1 = Color::Black;
			oTemplate.HatchStyle_T1 = HS_PERCENT50;
			oTemplate.HatchBackColor_T1 = Color::White;
			oTemplate.HatchForeColor_T1 = Color::Black;

			oTemplate.BorderColor_T2 = Color::Black;
			oTemplate.HatchStyle_T2 = HS_LIGHTUPWARDDIAGONAL;
			oTemplate.HatchBackColor_T2 = Color::White;
			oTemplate.HatchForeColor_T2 = Color::Black;

			oTemplate.BorderColor_T3 = Color::Black;
			oTemplate.HatchStyle_T3 = HS_LIGHTDOWNWARDDIAGONAL;
			oTemplate.HatchBackColor_T3 = Color::White;
			oTemplate.HatchForeColor_T3 = Color::Black;

			oTemplate.BorderColor_T4 = Color::Black;
			oTemplate.HatchStyle_T4 = HS_NARROWVERTICAL;
			oTemplate.HatchBackColor_T4 = Color::White;
			oTemplate.HatchForeColor_T4 = Color::Black;

			mp_ApplyObjectTemplate(oControl, oTemplate);
		}
		break;
	case STO_COLOR_HATCH:
		{
			S_ObjectTemplate oTemplate;

			oTemplate.yBackgroundMode_Tasks = FP_HATCH;

			oTemplate.BorderColor_T1 = Color::MakeARGB(255, 121, 163, 213);
			oTemplate.HatchStyle_T1 = HS_PERCENT50;
			oTemplate.HatchBackColor_T1 = Color::White;
			oTemplate.HatchForeColor_T1 = Color::MakeARGB(255, 121, 163, 213);

			oTemplate.BorderColor_T2 = Color::MakeARGB(255, 82, 94, 119);
			oTemplate.HatchStyle_T2 = HS_LIGHTUPWARDDIAGONAL;
			oTemplate.HatchBackColor_T2 = Color::White;
			oTemplate.HatchForeColor_T2 = Color::MakeARGB(255, 82, 94, 119);

			oTemplate.BorderColor_T3 = Color::MakeARGB(255, 230, 150, 24);
			oTemplate.HatchStyle_T3 = HS_LIGHTDOWNWARDDIAGONAL;
			oTemplate.HatchBackColor_T3 = Color::White;
			oTemplate.HatchForeColor_T3 = Color::MakeARGB(255, 230, 150, 24);

			oTemplate.BorderColor_T4 = Color::MakeARGB(255, 168, 121, 213);
			oTemplate.HatchStyle_T4 = HS_NARROWVERTICAL;
			oTemplate.HatchBackColor_T4 = Color::White;
			oTemplate.HatchForeColor_T4 = Color::MakeARGB(255, 168, 121, 213);

			mp_ApplyObjectTemplate(oControl, oTemplate);
		}
		break;
	case STO_GREY_SOLID:
		{
			S_ObjectTemplate oTemplate;

			oTemplate.yBackgroundMode_Tasks = FP_SOLID;

			oTemplate.BorderColor_T1 = Color::Black;
			oTemplate.BackColor_T1 = Color::MakeARGB(255, 200, 200, 200);

			oTemplate.BorderColor_T2 = Color::Black;
			oTemplate.BackColor_T2 = Color::MakeARGB(255, 166, 166, 166);

			oTemplate.BorderColor_T3 = Color::Black;
			oTemplate.BackColor_T3 = Color::MakeARGB(255, 133, 133, 133);

			oTemplate.BorderColor_T4 = Color::Black;
			oTemplate.BackColor_T4 = Color::MakeARGB(255, 100, 100, 100);

			mp_ApplyObjectTemplate(oControl, oTemplate);
		}
		break;
	case STO_COLORS_SOLID:
		{
			S_ObjectTemplate oTemplate;

			oTemplate.yBackgroundMode_Tasks = FP_SOLID;

			oTemplate.BackColor_T1 = Color::MakeARGB(255, 88, 127, 196);
			oTemplate.BorderColor_T1 = Color::MakeARGB(255, 57, 109, 182);

			oTemplate.BackColor_T2 = Color::MakeARGB(255, 195, 82, 76);
			oTemplate.BorderColor_T2 = Color::MakeARGB(255, 189, 55, 56);

			oTemplate.BackColor_T3 = Color::MakeARGB(255, 145, 105, 168);
			oTemplate.BorderColor_T3 = Color::MakeARGB(255, 82, 44, 103);

			oTemplate.BackColor_T4 = Color::MakeARGB(255, 248, 151, 70);
			oTemplate.BorderColor_T4 = Color::MakeARGB(255, 252, 114, 1);

			mp_ApplyObjectTemplate(oControl, oTemplate);
		}
		break;
	case STO_DEFAULT:
		{
			S_ObjectTemplate oTemplate;

			oTemplate.yBackgroundMode_Tasks = FP_GRADIENT;
			oTemplate.yGradientFillMode_Tasks = GDT_HORIZONTAL;
			oTemplate.BorderColor_T1 = Color::Blue;
			oTemplate.StartGradientColor_T1 = Color::White;
			oTemplate.EndGradientColor_T1 = Color::MakeARGB(255, 100, 100, 204);
			oTemplate.BorderColor_T2 = Color::Green;
			oTemplate.StartGradientColor_T2 = Color::White;
			oTemplate.EndGradientColor_T2 = Color::MakeARGB(255, 100, 204, 100);
			oTemplate.BorderColor_T3 = Color::Red;
			oTemplate.StartGradientColor_T3 = Color::White;
			oTemplate.EndGradientColor_T3 = Color::MakeARGB(255, 204, 100, 100);
			oTemplate.BorderColor_T4 = Color::Black;
			oTemplate.StartGradientColor_T4 = Color::White;
			oTemplate.EndGradientColor_T4 = Color::MakeARGB(255, 51, 51, 51);

			mp_ApplyObjectTemplate(oControl, oTemplate);
		}
		break;
	case STO_VARIATION_1:
		{
			S_ObjectTemplate oTemplate;

			oTemplate.yBackgroundMode_Tasks = FP_GRADIENT;
			oTemplate.yGradientFillMode_Tasks = GDT_HORIZONTAL;
			oTemplate.BorderColor_T1 = Color::Black;
			oTemplate.ForeColor_T1 = Color::White;
			oTemplate.StartGradientColor_T1 = Color::MakeARGB(255, 162, 78, 50);
			oTemplate.EndGradientColor_T1 = Color::MakeARGB(255, 215, 92, 54);
			oTemplate.BorderColor_T2 = Color::Black;
			oTemplate.ForeColor_T2 = Color::White;
			oTemplate.StartGradientColor_T2 = Color::MakeARGB(255, 109, 122, 136);
			oTemplate.EndGradientColor_T2 = Color::MakeARGB(255, 179, 199, 229);
			oTemplate.BorderColor_T3 = Color::Black;
			oTemplate.ForeColor_T3 = Color::White;
			oTemplate.StartGradientColor_T3 = Color::MakeARGB(255, 255, 77, 1);
			oTemplate.EndGradientColor_T3 = Color::MakeARGB(255, 244, 172, 43);
			oTemplate.BorderColor_T4 = Color::Black;
			oTemplate.StartGradientColor_T4 = Color::White;
			oTemplate.EndGradientColor_T4 = Color::MakeARGB(255, 51, 51, 51);

			mp_ApplyObjectTemplate(oControl, oTemplate);
		}
		break;

	}
}

void mp_ApplyObjectTemplate(CActiveGanttVCCtl* oControl, S_ObjectTemplate &oTemplate)
{
	clsStyle* oStyle;

	oStyle = oControl->GetConfiguration()->GetDefaultTaskStyle();
	oStyle->SetAppearance(SA_FLAT);
	oStyle->SetPlacement(PLC_OFFSETPLACEMENT);
	oStyle->SetBorderStyle(SBR_SINGLE);

	oStyle->SetBackColor(oTemplate.BackColor_T1);
	oStyle->SetForeColor(oTemplate.ForeColor_T1);
	oStyle->SetBorderColor(oTemplate.BorderColor_T1);

	oStyle->GetSelectionRectangleStyle()->SetOffsetTop(0);
	oStyle->GetSelectionRectangleStyle()->SetOffsetLeft(0);
	oStyle->GetSelectionRectangleStyle()->SetOffsetRight(0);
	oStyle->GetSelectionRectangleStyle()->SetOffsetBottom(0);
	oStyle->SetOffsetTop(5);
	oStyle->SetOffsetBottom(10);
	oStyle->SetBackgroundMode(oTemplate.yBackgroundMode_Tasks);
	oStyle->SetHatchStyle(oTemplate.HatchStyle_T1);
	oStyle->SetHatchBackColor(oTemplate.HatchBackColor_T1);
	oStyle->SetHatchForeColor(oTemplate.HatchForeColor_T1);
	oStyle->SetGradientFillMode(oTemplate.yGradientFillMode_Tasks);
	oStyle->SetStartGradientColor(oTemplate.StartGradientColor_T1);
	oStyle->SetEndGradientColor(oTemplate.EndGradientColor_T1);
	oStyle->GetMilestoneStyle()->SetShapeIndex(FT_DIAMOND);
	oStyle->GetMilestoneStyle()->SetBorderColor(oTemplate.BorderColor_T1);
	oStyle->GetMilestoneStyle()->SetFillColor(oTemplate.BorderColor_T1);

	oStyle->GetPredecessorStyle()->SetLineColor(oTemplate.BorderColor_T1);


	oStyle = oControl->GetConfiguration()->GetDefaultTaskStyle()->Clone("T1");

	oStyle = oControl->GetConfiguration()->GetDefaultTaskStyle()->Clone("T2");
	oStyle->SetBackColor(oTemplate.BackColor_T2);
	oStyle->SetForeColor(oTemplate.ForeColor_T2);
	oStyle->SetBorderColor(oTemplate.BorderColor_T2);

	oStyle->SetHatchStyle(oTemplate.HatchStyle_T2);
	oStyle->SetHatchBackColor(oTemplate.HatchBackColor_T2);
	oStyle->SetHatchForeColor(oTemplate.HatchForeColor_T2);
	oStyle->SetStartGradientColor(oTemplate.StartGradientColor_T2);
	oStyle->SetEndGradientColor(oTemplate.EndGradientColor_T2);
	oStyle->GetPredecessorStyle()->SetLineColor(oTemplate.BorderColor_T2);
	oStyle->GetMilestoneStyle()->SetBorderColor(oTemplate.BorderColor_T2);
	oStyle->GetMilestoneStyle()->SetFillColor(oTemplate.BorderColor_T2);

	oStyle = oControl->GetConfiguration()->GetDefaultTaskStyle()->Clone("T3");
	oStyle->SetBackColor(oTemplate.BackColor_T3);
	oStyle->SetForeColor(oTemplate.ForeColor_T3);
	oStyle->SetBorderColor(oTemplate.BorderColor_T3);

	oStyle->SetHatchStyle(oTemplate.HatchStyle_T3);
	oStyle->SetHatchBackColor(oTemplate.HatchBackColor_T3);
	oStyle->SetHatchForeColor(oTemplate.HatchForeColor_T3);
	oStyle->SetStartGradientColor(oTemplate.StartGradientColor_T3);
	oStyle->SetEndGradientColor(oTemplate.EndGradientColor_T3);
	oStyle->GetPredecessorStyle()->SetLineColor(oTemplate.BorderColor_T3);
	oStyle->GetMilestoneStyle()->SetBorderColor(oTemplate.BorderColor_T3);
	oStyle->GetMilestoneStyle()->SetFillColor(oTemplate.BorderColor_T3);

	oStyle = oControl->GetConfiguration()->GetDefaultTaskStyle()->Clone("T4");
	oStyle->SetBackColor(oTemplate.BackColor_T4);
	oStyle->SetForeColor(oTemplate.ForeColor_T4);
	oStyle->SetBorderColor(oTemplate.BorderColor_T4);

	oStyle->SetHatchStyle(oTemplate.HatchStyle_T4);
	oStyle->SetHatchBackColor(oTemplate.HatchBackColor_T4);
	oStyle->SetHatchForeColor(oTemplate.HatchForeColor_T4);
	oStyle->SetStartGradientColor(oTemplate.StartGradientColor_T4);
	oStyle->SetEndGradientColor(oTemplate.EndGradientColor_T4);
	oStyle->GetPredecessorStyle()->SetLineColor(oTemplate.BorderColor_T4);
	oStyle->GetMilestoneStyle()->SetBorderColor(oTemplate.BorderColor_T4);
	oStyle->GetMilestoneStyle()->SetFillColor(oTemplate.BorderColor_T4);

	oStyle = oControl->GetStyles()->Item("T1")->Clone("TW1");
	oStyle->SetBorderColor(Color::Red);
	oStyle->GetPredecessorStyle()->SetLineColor(Color::Red);

	oStyle = oControl->GetStyles()->Item("T2")->Clone("TW2");
	oStyle->SetBorderColor(Color::Red);
	oStyle->GetPredecessorStyle()->SetLineColor(Color::Red);

	oStyle = oControl->GetStyles()->Item("T3")->Clone("TW3");
	oStyle->SetBorderColor(Color::Yellow);
	oStyle->GetPredecessorStyle()->SetLineColor(Color::Red);

	oStyle = oControl->GetStyles()->Item("T4")->Clone("TW4");
	oStyle->SetBorderColor(Color::Red);
	oStyle->GetPredecessorStyle()->SetLineColor(Color::Red);

	oStyle = oControl->GetStyles()->Item("T1")->Clone("S1");
	oStyle->GetSelectionRectangleStyle()->SetVisible(FALSE);
	oStyle->GetTaskStyle()->SetStartShapeIndex(FT_PROJECTDOWN);
	oStyle->GetTaskStyle()->SetEndShapeIndex(FT_PROJECTDOWN);
	oStyle->GetTaskStyle()->SetStartBorderColor(oTemplate.BorderColor_T1);
	oStyle->GetTaskStyle()->SetStartFillColor(oTemplate.BorderColor_T1);
	oStyle->GetTaskStyle()->SetEndBorderColor(oTemplate.BorderColor_T1);
	oStyle->GetTaskStyle()->SetEndFillColor(oTemplate.BorderColor_T1);
	oStyle->SetFillMode(FM_UPPERHALFFILLED);

	oStyle = oControl->GetStyles()->Item("T2")->Clone("S2");
	oStyle->GetSelectionRectangleStyle()->SetVisible(FALSE);
	oStyle->GetTaskStyle()->SetStartShapeIndex(FT_PROJECTDOWN);
	oStyle->GetTaskStyle()->SetEndShapeIndex(FT_PROJECTDOWN);
	oStyle->GetTaskStyle()->SetStartBorderColor(oTemplate.BorderColor_T2);
	oStyle->GetTaskStyle()->SetStartFillColor(oTemplate.BorderColor_T2);
	oStyle->GetTaskStyle()->SetEndBorderColor(oTemplate.BorderColor_T2);
	oStyle->GetTaskStyle()->SetEndFillColor(oTemplate.BorderColor_T2);
	oStyle->SetFillMode(FM_UPPERHALFFILLED);

	oStyle = oControl->GetStyles()->Item("T3")->Clone("S3");
	oStyle->GetSelectionRectangleStyle()->SetVisible(FALSE);
	oStyle->GetTaskStyle()->SetStartShapeIndex(FT_PROJECTDOWN);
	oStyle->GetTaskStyle()->SetEndShapeIndex(FT_PROJECTDOWN);
	oStyle->GetTaskStyle()->SetStartBorderColor(oTemplate.BorderColor_T3);
	oStyle->GetTaskStyle()->SetStartFillColor(oTemplate.BorderColor_T3);
	oStyle->GetTaskStyle()->SetEndBorderColor(oTemplate.BorderColor_T3);
	oStyle->GetTaskStyle()->SetEndFillColor(oTemplate.BorderColor_T3);
	oStyle->SetFillMode(FM_UPPERHALFFILLED);

	oStyle = oControl->GetStyles()->Item("T4")->Clone("S4");
	oStyle->GetSelectionRectangleStyle()->SetVisible(FALSE);
	oStyle->GetTaskStyle()->SetStartShapeIndex(FT_PROJECTDOWN);
	oStyle->GetTaskStyle()->SetEndShapeIndex(FT_PROJECTDOWN);
	oStyle->GetTaskStyle()->SetStartBorderColor(oTemplate.BorderColor_T4);
	oStyle->GetTaskStyle()->SetStartFillColor(oTemplate.BorderColor_T4);
	oStyle->GetTaskStyle()->SetEndBorderColor(oTemplate.BorderColor_T4);
	oStyle->GetTaskStyle()->SetEndFillColor(oTemplate.BorderColor_T4);
	oStyle->SetFillMode(FM_UPPERHALFFILLED);

	oStyle = oControl->GetConfiguration()->GetDefaultPercentageStyle();
	oStyle->SetAppearance(SA_FLAT);
	oStyle->SetPlacement(PLC_OFFSETPLACEMENT);
	oStyle->SetBorderStyle(SBR_SINGLE);
	oStyle->GetSelectionRectangleStyle()->SetOffsetTop(0);
	oStyle->GetSelectionRectangleStyle()->SetOffsetLeft(0);
	oStyle->GetSelectionRectangleStyle()->SetOffsetRight(0);
	oStyle->GetSelectionRectangleStyle()->SetOffsetBottom(0);
	oStyle->SetOffsetTop(8);
	oStyle->SetOffsetBottom(4);
	oStyle->GetSelectionRectangleStyle()->SetVisible(TRUE);
	oStyle->SetTextVisible(FALSE);
	oStyle->SetBackgroundMode(FP_SOLID);
	oStyle->SetBackColor(oTemplate.BorderColor_T1);
	oStyle->SetBorderColor(oTemplate.BorderColor_T1);

	oStyle = oControl->GetConfiguration()->GetDefaultPercentageStyle()->Clone("P1");

	oStyle = oControl->GetConfiguration()->GetDefaultPercentageStyle()->Clone("P2");
	oStyle->SetBackColor(oTemplate.BorderColor_T2);
	oStyle->SetBorderColor(oTemplate.BorderColor_T2);

	oStyle = oControl->GetConfiguration()->GetDefaultPercentageStyle()->Clone("P3");
	oStyle->SetBackColor(oTemplate.BorderColor_T3);
	oStyle->SetBorderColor(oTemplate.BorderColor_T3);

	oStyle = oControl->GetConfiguration()->GetDefaultPercentageStyle()->Clone("P4");
	oStyle->SetBackColor(oTemplate.BorderColor_T4);
	oStyle->SetBorderColor(oTemplate.BorderColor_T4);

	oStyle = oControl->GetConfiguration()->GetDefaultPercentageStyle()->Clone("SP1");
	oStyle->SetBackColor(oTemplate.BorderColor_T1);
	oStyle->SetBorderColor(oTemplate.BorderColor_T1);
	oStyle->SetOffsetTop(5);
	oStyle->SetOffsetBottom(5);
	oStyle->GetSelectionRectangleStyle()->SetVisible(FALSE);
	oStyle->SetBackgroundMode(FP_SOLID);

	oStyle = oStyle->Clone("SP2");
	oStyle->SetBackColor(oTemplate.BorderColor_T2);
	oStyle->SetBorderColor(oTemplate.BorderColor_T2);

	oStyle = oStyle->Clone("SP3");
	oStyle->SetBackColor(oTemplate.BorderColor_T3);
	oStyle->SetBorderColor(oTemplate.BorderColor_T3);

	oStyle = oStyle->Clone("SP4");
	oStyle->SetBackColor(oTemplate.BorderColor_T4);
	oStyle->SetBorderColor(oTemplate.BorderColor_T4);



	oStyle = oControl->GetStyles()->Add("RET1");
	oStyle->SetAppearance(SA_FLAT);
	oStyle->SetBackgroundMode(oTemplate.yBackgroundMode_Tasks);
	oStyle->SetGradientFillMode(oTemplate.yGradientFillMode_Tasks);
	oStyle->SetBorderStyle(SBR_SINGLE);
	oStyle->SetPlacement(PLC_ROWEXTENTSPLACEMENT);
	oStyle->SetBackColor(oTemplate.BackColor_T1);
	oStyle->SetForeColor(oTemplate.ForeColor_T1);
	oStyle->SetBorderColor(oTemplate.BorderColor_T1);
	oStyle->SetHatchStyle(oTemplate.HatchStyle_T1);
	oStyle->SetHatchBackColor(oTemplate.HatchBackColor_T1);
	oStyle->SetHatchForeColor(oTemplate.HatchForeColor_T1);
	oStyle->SetStartGradientColor(oTemplate.StartGradientColor_T1);
	oStyle->SetEndGradientColor(oTemplate.EndGradientColor_T1);
	oStyle->GetPredecessorStyle()->SetLineColor(oTemplate.BorderColor_T1);
	oStyle->GetMilestoneStyle()->SetBorderColor(oTemplate.BorderColor_T1);
	oStyle->GetMilestoneStyle()->SetFillColor(oTemplate.BorderColor_T1);

	oStyle = oControl->GetStyles()->Item("RET1")->Clone("RET2");
	oStyle->SetBackColor(oTemplate.BackColor_T2);
	oStyle->SetForeColor(oTemplate.ForeColor_T2);
	oStyle->SetBorderColor(oTemplate.BorderColor_T2);
	oStyle->SetHatchStyle(oTemplate.HatchStyle_T2);
	oStyle->SetHatchBackColor(oTemplate.HatchBackColor_T2);
	oStyle->SetHatchForeColor(oTemplate.HatchForeColor_T2);
	oStyle->SetStartGradientColor(oTemplate.StartGradientColor_T2);
	oStyle->SetEndGradientColor(oTemplate.EndGradientColor_T2);
	oStyle->GetPredecessorStyle()->SetLineColor(oTemplate.BorderColor_T2);
	oStyle->GetMilestoneStyle()->SetBorderColor(oTemplate.BorderColor_T2);
	oStyle->GetMilestoneStyle()->SetFillColor(oTemplate.BorderColor_T2);

	oStyle = oControl->GetStyles()->Item("RET1")->Clone("RET3");
	oStyle->SetBackColor(oTemplate.BackColor_T3);
	oStyle->SetForeColor(oTemplate.ForeColor_T3);
	oStyle->SetBorderColor(oTemplate.BorderColor_T3);
	oStyle->SetHatchStyle(oTemplate.HatchStyle_T3);
	oStyle->SetHatchBackColor(oTemplate.HatchBackColor_T3);
	oStyle->SetHatchForeColor(oTemplate.HatchForeColor_T3);
	oStyle->SetStartGradientColor(oTemplate.StartGradientColor_T3);
	oStyle->SetEndGradientColor(oTemplate.EndGradientColor_T3);
	oStyle->GetPredecessorStyle()->SetLineColor(oTemplate.BorderColor_T3);
	oStyle->GetMilestoneStyle()->SetBorderColor(oTemplate.BorderColor_T3);
	oStyle->GetMilestoneStyle()->SetFillColor(oTemplate.BorderColor_T3);

	oStyle = oControl->GetStyles()->Item("RET1")->Clone("RET4");
	oStyle->SetBackColor(oTemplate.BackColor_T4);
	oStyle->SetForeColor(oTemplate.ForeColor_T4);
	oStyle->SetBorderColor(oTemplate.BorderColor_T4);
	oStyle->SetHatchStyle(oTemplate.HatchStyle_T4);
	oStyle->SetHatchBackColor(oTemplate.HatchBackColor_T4);
	oStyle->SetHatchForeColor(oTemplate.HatchForeColor_T4);
	oStyle->SetStartGradientColor(oTemplate.StartGradientColor_T4);
	oStyle->SetEndGradientColor(oTemplate.EndGradientColor_T4);
	oStyle->GetPredecessorStyle()->SetLineColor(oTemplate.BorderColor_T4);
	oStyle->GetMilestoneStyle()->SetBorderColor(oTemplate.BorderColor_T4);
	oStyle->GetMilestoneStyle()->SetFillColor(oTemplate.BorderColor_T4);
}

void g_ApplyTemplate_Gradient(CActiveGanttVCCtl* oControl, ControlTemplateGradient &oTemplate, LONG yObjectTemplate)
{
	oControl->GetStyles()->Clear();
	mp_ApplyControlTemplate_Gradient(oControl, oTemplate);
	mp_ObjectTemplateSelector(oControl, yObjectTemplate);
}

void mp_ApplyControlTemplate_Gradient(CActiveGanttVCCtl* oControl, ControlTemplateGradient &oTemplate)
{
	clsStyle* oStyle;

	oStyle = oControl->GetConfiguration()->GetDefaultControlStyle();
	oStyle->SetAppearance(SA_FLAT);
	oStyle->SetBorderStyle(SBR_CUSTOM);
	oStyle->SetBorderColor(oTemplate.ControlBorderColor);
	oStyle->SetBackColor(Color::White);

	oStyle = oControl->GetConfiguration()->GetDefaultColumnStyle();
	oStyle->GetFont()->SetPtFont(oControl->GetConfiguration()->GetDefaultFont()->GetFontName(), 8, FS_FONTSTYLEBOLD);
	oStyle->SetAppearance(SA_FLAT);
	oStyle->SetBackgroundMode(FP_GRADIENT);
	oStyle->SetGradientFillMode(oTemplate.GradientFillMode);
	oStyle->SetStartGradientColor(oTemplate.StartGradientColor);
	oStyle->SetEndGradientColor(oTemplate.EndGradientColor);
	oStyle->GetCustomBorderStyle()->SetLeft(FALSE);
	oStyle->GetCustomBorderStyle()->SetTop(FALSE);
	oStyle->GetCustomBorderStyle()->SetRight(TRUE);
	oStyle->GetCustomBorderStyle()->SetBottom(TRUE);
	oStyle->SetBorderStyle(SBR_CUSTOM);
	oStyle->SetBorderColor(oTemplate.HeadingBorderColor);
	oStyle->SetForeColor(oTemplate.HeadingForeColor);

	oStyle = oControl->GetConfiguration()->GetDefaultTimeLineStyle();
	oStyle->SetAppearance(SA_FLAT);
	oStyle->SetBackgroundMode(FP_GRADIENT);
	oStyle->SetBorderStyle(SBR_NONE);
	oStyle->SetGradientFillMode(oTemplate.GradientFillMode);
	oStyle->SetStartGradientColor(oTemplate.StartGradientColor);
	oStyle->SetEndGradientColor(oTemplate.EndGradientColor);


	oStyle = oControl->GetConfiguration()->GetDefaultTierStyle();
	oStyle->GetFont()->SetPtFont(oControl->GetConfiguration()->GetDefaultFont()->GetFontName(), 7, FS_FONTSTYLEREGULAR);
	oStyle->SetAppearance(SA_FLAT);
	if (oTemplate.GradientFillMode == GDT_VERTICAL) {
		oStyle->SetBackgroundMode(FP_TRANSPARENT);
	} else {
		oStyle->SetBackgroundMode(FP_GRADIENT);
		oStyle->SetGradientFillMode(GDT_HORIZONTAL);
		oStyle->SetStartGradientColor(oTemplate.StartGradientColor);
		oStyle->SetEndGradientColor(oTemplate.EndGradientColor);
	}
	oStyle->GetCustomBorderStyle()->SetLeft(TRUE);
	oStyle->GetCustomBorderStyle()->SetTop(FALSE);
	oStyle->GetCustomBorderStyle()->SetRight(FALSE);
	oStyle->GetCustomBorderStyle()->SetBottom(TRUE);
	oStyle->SetBorderStyle(SBR_CUSTOM);
	oStyle->SetBorderColor(oTemplate.HeadingBorderColor);
	oStyle->SetClipText(FALSE);
	oStyle->SetForeColor(oTemplate.HeadingForeColor);
	oControl->GetCurrentViewObject()->GetTimeLine()->GetTierArea()->GetUpperTier()->SetBackgroundMode(ETBGM_STYLE);
	oControl->GetCurrentViewObject()->GetTimeLine()->GetTierArea()->GetMiddleTier()->SetBackgroundMode(ETBGM_STYLE);
	oControl->GetCurrentViewObject()->GetTimeLine()->GetTierArea()->GetLowerTier()->SetBackgroundMode(ETBGM_STYLE);

	oStyle = oControl->GetConfiguration()->GetDefaultTickMarkAreaStyle();
	oStyle->GetFont()->SetPtFont(oControl->GetConfiguration()->GetDefaultFont()->GetFontName(), 7, FS_FONTSTYLEREGULAR);
	oStyle->SetAppearance(SA_FLAT);
	oStyle->SetBackgroundMode(FP_TRANSPARENT);
	oStyle->GetCustomBorderStyle()->SetLeft(TRUE);
	oStyle->GetCustomBorderStyle()->SetTop(FALSE);
	oStyle->GetCustomBorderStyle()->SetRight(FALSE);
	oStyle->GetCustomBorderStyle()->SetBottom(TRUE);
	oStyle->SetBorderStyle(SBR_CUSTOM);
	oStyle->SetBorderColor(oTemplate.HeadingBorderColor);
	oStyle->SetForeColor(oTemplate.HeadingForeColor);

	oStyle = oControl->GetConfiguration()->GetDefaultRowStyle();
	oStyle->GetFont()->SetPtFont(oControl->GetConfiguration()->GetDefaultFont()->GetFontName(), 8, FS_FONTSTYLEREGULAR);
	oStyle->SetAppearance(SA_FLAT);
	oStyle->SetBackColor(Color::White);
	oStyle->SetBorderColor(oTemplate.RowBorderColor);
	oStyle->SetBorderStyle(SBR_CUSTOM);
	oStyle->GetCustomBorderStyle()->SetTop(FALSE);
	oStyle->GetCustomBorderStyle()->SetLeft(FALSE);
	oStyle->SetForeColor(oTemplate.RowForeColor);

	oStyle = oControl->GetConfiguration()->GetDefaultCellStyle();
	oStyle->GetFont()->SetPtFont(oControl->GetConfiguration()->GetDefaultFont()->GetFontName(), 8, FS_FONTSTYLEREGULAR);
	oStyle->SetAppearance(SA_FLAT);
	oStyle->SetBackColor(Color::White);
	oStyle->SetBorderColor(oTemplate.RowBorderColor);
	oStyle->SetBorderStyle(SBR_CUSTOM);
	oStyle->GetCustomBorderStyle()->SetTop(FALSE);
	oStyle->GetCustomBorderStyle()->SetLeft(FALSE);
	oStyle->SetForeColor(oTemplate.RowForeColor);

	oStyle = oControl->GetConfiguration()->GetDefaultCellStyle()->Clone("CL");

	oStyle = oControl->GetConfiguration()->GetDefaultCellStyle()->Clone("CR");
	oStyle->SetTextAlignmentHorizontal(HAL_RIGHT);

	oStyle = oControl->GetConfiguration()->GetDefaultCellStyle()->Clone("CH");
	oStyle->SetTextAlignmentHorizontal(HAL_RIGHT);
	oStyle->SetBorderColor(oTemplate.ControlBorderColor);


	oStyle = oControl->GetConfiguration()->GetDefaultNodeStyle();
	oStyle->GetFont()->SetPtFont(oControl->GetConfiguration()->GetDefaultFont()->GetFontName(), 8, FS_FONTSTYLEREGULAR);
	oStyle->SetAppearance(SA_FLAT);
	oStyle->SetBackgroundMode(FP_SOLID);
	oStyle->SetBackColor(Color::White);
	oStyle->SetBorderColor(oTemplate.RowBorderColor);
	oStyle->SetBorderStyle(SBR_CUSTOM);
	oStyle->GetCustomBorderStyle()->SetTop(FALSE);
	oStyle->GetCustomBorderStyle()->SetLeft(FALSE);
	oStyle->SetForeColor(oTemplate.RowForeColor);

	oStyle = oControl->GetConfiguration()->GetDefaultNodeStyle()->Clone("NR");

	oStyle = oControl->GetConfiguration()->GetDefaultNodeStyle()->Clone("NB");
	oStyle->GetFont()->SetPtFont(oControl->GetConfiguration()->GetDefaultFont()->GetFontName(), 8, FS_FONTSTYLEBOLD);

	oStyle = oControl->GetConfiguration()->GetDefaultClientAreaStyle();
	oStyle->SetAppearance(SA_FLAT);
	oStyle->SetBackColor(Color::White);
	oStyle->SetBorderColor(oTemplate.RowBorderColor);
	oStyle->SetBorderStyle(SBR_CUSTOM);
	oStyle->GetCustomBorderStyle()->SetTop(FALSE);
	oStyle->GetCustomBorderStyle()->SetLeft(FALSE);
	oStyle->GetCustomBorderStyle()->SetRight(FALSE);

	oControl->GetTreeview()->SetPlusMinusBorderColor(oTemplate.TreeviewColor);
	oControl->GetTreeview()->SetPlusMinusSignColor(oTemplate.TreeviewColor);
	oControl->GetTreeview()->SetCheckBoxBorderColor(oTemplate.TreeviewColor);
	oControl->GetTreeview()->SetCheckBoxMarkColor(oTemplate.TreeviewColor);
	oControl->GetTreeview()->SetTreeLineColor(oTemplate.TreeviewColor);
	oControl->GetTreeview()->SetCheckBoxBackColor(oTemplate.TreeviewColor);
	oControl->GetTreeview()->SetPlusMinusBackColor(oTemplate.TreeviewColor);

	oControl->GetSplitter()->SetType(SA_USERDEFINED);
	oControl->GetSplitter()->SetWidth(1);
	oControl->GetSplitter()->SetColor(1, oTemplate.ControlBorderColor);

	oStyle = oControl->GetConfiguration()->GetDefaultSBSeparatorStyle();
	oStyle->SetAppearance(SA_FLAT);
	oStyle->SetBackgroundMode(FP_SOLID);
	oStyle->SetBackColor(Color::White);
	oStyle->SetBorderStyle(SBR_SINGLE);
	oStyle->SetBorderColor(oTemplate.RowBorderColor);

	oStyle = oControl->GetConfiguration()->GetDefaultScrollBarStyle();
	oStyle->SetAppearance(SA_FLAT);
	oStyle->SetBackgroundMode(FP_SOLID);
	oStyle->SetBackColor(Color::White);
	oStyle->SetBorderStyle(SBR_SINGLE);
	oStyle->SetBorderColor(oTemplate.RowBorderColor);
	oStyle->GetScrollBarStyle()->SetArrowColor(Color::Black);

	oStyle = oControl->GetStyles()->Add("AB");
	oStyle->SetAppearance(SA_FLAT);
	oStyle->SetBackgroundMode(FP_SOLID);
	oStyle->SetBackColor(Color::White);
	oStyle->SetBorderStyle(SBR_SINGLE);
	oStyle->SetBorderColor(oTemplate.RowBorderColor);
	oStyle->GetScrollBarStyle()->SetArrowColor(Color::Black);

	oStyle = oControl->GetStyles()->Add("TBH");
	oStyle->SetAppearance(SA_FLAT);
	oStyle->SetBackgroundMode(FP_GRADIENT);
	oStyle->SetGradientFillMode(GDT_VERTICAL);
	oStyle->SetStartGradientColor(Color::White);
	oStyle->SetEndGradientColor(oTemplate.EndGradientColor);
	oStyle->SetBorderStyle(SBR_SINGLE);
	oStyle->SetBorderColor(oTemplate.RowBorderColor);
	oStyle->SetBackColor(Color::White);

	oStyle = oControl->GetStyles()->Item("TBH")->Clone("TBHP");
	oStyle->SetStartGradientColor(oTemplate.EndGradientColor);
	oStyle->SetEndGradientColor(Color::White);

	oStyle = oControl->GetStyles()->Add("TBV");
	oStyle->SetAppearance(SA_FLAT);
	oStyle->SetBackgroundMode(FP_GRADIENT);
	oStyle->SetGradientFillMode(GDT_HORIZONTAL);
	oStyle->SetStartGradientColor(Color::White);
	oStyle->SetEndGradientColor(oTemplate.EndGradientColor);
	oStyle->SetBorderStyle(SBR_SINGLE);
	oStyle->SetBorderColor(oTemplate.RowBorderColor);
	oStyle->SetBackColor(Color::White);

	oStyle = oControl->GetStyles()->Item("TBV")->Clone("TBVP");
	oStyle->SetStartGradientColor(oTemplate.EndGradientColor);
	oStyle->SetEndGradientColor(Color::White);

	oControl->GetVerticalScrollBar()->GetScrollBar()->GetArrowButtons()->SetNormalStyleIndex("AB");
	oControl->GetVerticalScrollBar()->GetScrollBar()->GetArrowButtons()->SetPressedStyleIndex("AB");
	oControl->GetVerticalScrollBar()->GetScrollBar()->GetArrowButtons()->SetDisabledStyleIndex("AB");
	oControl->GetVerticalScrollBar()->GetScrollBar()->GetThumbButton()->SetNormalStyleIndex("TBV");
	oControl->GetVerticalScrollBar()->GetScrollBar()->GetThumbButton()->SetPressedStyleIndex("TBVP");
	oControl->GetVerticalScrollBar()->GetScrollBar()->GetThumbButton()->SetDisabledStyleIndex("TBV");

	oControl->GetHorizontalScrollBar()->GetScrollBar()->GetArrowButtons()->SetNormalStyleIndex("AB");
	oControl->GetHorizontalScrollBar()->GetScrollBar()->GetArrowButtons()->SetPressedStyleIndex("AB");
	oControl->GetHorizontalScrollBar()->GetScrollBar()->GetArrowButtons()->SetDisabledStyleIndex("AB");
	oControl->GetHorizontalScrollBar()->GetScrollBar()->GetThumbButton()->SetNormalStyleIndex("TBH");
	oControl->GetHorizontalScrollBar()->GetScrollBar()->GetThumbButton()->SetPressedStyleIndex("TBHP");
	oControl->GetHorizontalScrollBar()->GetScrollBar()->GetThumbButton()->SetDisabledStyleIndex("TBH");

	oControl->GetTimeLineScrollBar()->GetScrollBar()->GetArrowButtons()->SetNormalStyleIndex("AB");
	oControl->GetTimeLineScrollBar()->GetScrollBar()->GetArrowButtons()->SetPressedStyleIndex("AB");
	oControl->GetTimeLineScrollBar()->GetScrollBar()->GetArrowButtons()->SetDisabledStyleIndex("AB");
	oControl->GetTimeLineScrollBar()->GetScrollBar()->GetThumbButton()->SetNormalStyleIndex("TBH");
	oControl->GetTimeLineScrollBar()->GetScrollBar()->GetThumbButton()->SetPressedStyleIndex("TBHP");
	oControl->GetTimeLineScrollBar()->GetScrollBar()->GetThumbButton()->SetDisabledStyleIndex("TBH");

	clsTimeLine* oTimeLine = oControl->GetCurrentViewObject()->GetTimeLine();

	oTimeLine->GetTimeLineScrollBar()->GetScrollBar()->GetArrowButtons()->SetNormalStyleIndex("AB");
	oTimeLine->GetTimeLineScrollBar()->GetScrollBar()->GetArrowButtons()->SetPressedStyleIndex("AB");
	oTimeLine->GetTimeLineScrollBar()->GetScrollBar()->GetArrowButtons()->SetDisabledStyleIndex("AB");
	oTimeLine->GetTimeLineScrollBar()->GetScrollBar()->GetThumbButton()->SetNormalStyleIndex("TBH");
	oTimeLine->GetTimeLineScrollBar()->GetScrollBar()->GetThumbButton()->SetPressedStyleIndex("TBHP");
	oTimeLine->GetTimeLineScrollBar()->GetScrollBar()->GetThumbButton()->SetDisabledStyleIndex("TBH");

	oStyle = oControl->GetConfiguration()->GetDefaultTimeBlockStyle();
	oStyle->SetBorderStyle(SBR_NONE);
	oStyle->SetAppearance(SA_FLAT);
	oStyle->SetBackgroundMode(FP_GRADIENT);
	oStyle->SetGradientFillMode(GDT_HORIZONTAL);
	oStyle->SetStartGradientColor(oTemplate.StartGradientColor);
	oStyle->SetEndGradientColor(oTemplate.EndGradientColor);
}

void g_ApplyTemplate_Solid(CActiveGanttVCCtl* oControl, ControlTemplateSolid &oTemplate, LONG yObjectTemplate)
{
	oControl->GetStyles()->Clear();
	mp_ApplyControlTemplate_Solid(oControl, oTemplate);
	mp_ObjectTemplateSelector(oControl, yObjectTemplate);

}

void mp_ApplyControlTemplate_Solid(CActiveGanttVCCtl* oControl, ControlTemplateSolid &oTemplate)
{
	clsStyle* oStyle;

	oStyle = oControl->GetConfiguration()->GetDefaultControlStyle();
	oStyle->SetAppearance(SA_FLAT);
	oStyle->SetBorderStyle(SBR_CUSTOM);
	oStyle->SetBorderColor(oTemplate.ControlBorderColor);
	oStyle->SetBackColor(Color::White);

	oStyle = oControl->GetConfiguration()->GetDefaultColumnStyle();
	oStyle->GetFont()->SetPtFont(oControl->GetConfiguration()->GetDefaultFont()->GetFontName(), 8, FS_FONTSTYLEBOLD);
	oStyle->SetAppearance(SA_FLAT);
	oStyle->SetBackgroundMode(FP_SOLID);
	oStyle->SetBackColor(oTemplate.HeadingBackColor);
	oStyle->GetCustomBorderStyle()->SetLeft(FALSE);
	oStyle->GetCustomBorderStyle()->SetTop(FALSE);
	oStyle->GetCustomBorderStyle()->SetRight(TRUE);
	oStyle->GetCustomBorderStyle()->SetBottom(TRUE);
	oStyle->SetBorderStyle(SBR_CUSTOM);
	oStyle->SetBorderColor(oTemplate.HeadingBorderColor);
	oStyle->SetForeColor(oTemplate.HeadingForeColor);

	oStyle = oControl->GetConfiguration()->GetDefaultTimeLineStyle();
	oStyle->SetAppearance(SA_FLAT);
	oStyle->SetBorderStyle(SBR_NONE);
	oStyle->SetBackgroundMode(FP_SOLID);
	oStyle->SetBackColor(oTemplate.HeadingBackColor);


	oStyle = oControl->GetConfiguration()->GetDefaultTierStyle();
	oStyle->GetFont()->SetPtFont(oControl->GetConfiguration()->GetDefaultFont()->GetFontName(), 7, FS_FONTSTYLEREGULAR);
	oStyle->SetAppearance(SA_FLAT);
	oStyle->SetBackgroundMode(FP_SOLID);
	oStyle->SetBackColor(oTemplate.HeadingBackColor);
	oStyle->GetCustomBorderStyle()->SetLeft(TRUE);
	oStyle->GetCustomBorderStyle()->SetTop(FALSE);
	oStyle->GetCustomBorderStyle()->SetRight(FALSE);
	oStyle->GetCustomBorderStyle()->SetBottom(TRUE);
	oStyle->SetBorderStyle(SBR_CUSTOM);
	oStyle->SetBorderColor(oTemplate.HeadingBorderColor);
	oStyle->SetClipText(FALSE);
	oStyle->SetForeColor(oTemplate.HeadingForeColor);
	oControl->GetCurrentViewObject()->GetTimeLine()->GetTierArea()->GetUpperTier()->SetBackgroundMode(ETBGM_STYLE);
	oControl->GetCurrentViewObject()->GetTimeLine()->GetTierArea()->GetMiddleTier()->SetBackgroundMode(ETBGM_STYLE);
	oControl->GetCurrentViewObject()->GetTimeLine()->GetTierArea()->GetLowerTier()->SetBackgroundMode(ETBGM_STYLE);

	oStyle = oControl->GetConfiguration()->GetDefaultTickMarkAreaStyle();
	oStyle->GetFont()->SetPtFont(oControl->GetConfiguration()->GetDefaultFont()->GetFontName(), 7, FS_FONTSTYLEREGULAR);
	oStyle->SetAppearance(SA_FLAT);
	oStyle->SetBackgroundMode(FP_TRANSPARENT);
	oStyle->GetCustomBorderStyle()->SetLeft(TRUE);
	oStyle->GetCustomBorderStyle()->SetTop(FALSE);
	oStyle->GetCustomBorderStyle()->SetRight(FALSE);
	oStyle->GetCustomBorderStyle()->SetBottom(TRUE);
	oStyle->SetBorderStyle(SBR_CUSTOM);
	oStyle->SetBorderColor(oTemplate.HeadingBorderColor);
	oStyle->SetForeColor(oTemplate.HeadingForeColor);

	oStyle = oControl->GetConfiguration()->GetDefaultRowStyle();
	oStyle->GetFont()->SetPtFont(oControl->GetConfiguration()->GetDefaultFont()->GetFontName(), 7, FS_FONTSTYLEREGULAR);
	oStyle->SetAppearance(SA_FLAT);
	oStyle->SetBackColor(Color::White);
	oStyle->SetBorderColor(oTemplate.RowBorderColor);
	oStyle->SetBorderStyle(SBR_CUSTOM);
	oStyle->GetCustomBorderStyle()->SetTop(FALSE);
	oStyle->GetCustomBorderStyle()->SetLeft(FALSE);
	oStyle->SetForeColor(oTemplate.RowForeColor);

	oStyle = oControl->GetConfiguration()->GetDefaultCellStyle();
	oStyle->GetFont()->SetPtFont(oControl->GetConfiguration()->GetDefaultFont()->GetFontName(), 7, FS_FONTSTYLEREGULAR);
	oStyle->SetAppearance(SA_FLAT);
	oStyle->SetBackColor(Color::White);
	oStyle->SetBorderColor(oTemplate.RowBorderColor);
	oStyle->SetBorderStyle(SBR_CUSTOM);
	oStyle->GetCustomBorderStyle()->SetTop(FALSE);
	oStyle->GetCustomBorderStyle()->SetLeft(FALSE);
	oStyle->SetForeColor(oTemplate.RowForeColor);

	oStyle = oControl->GetConfiguration()->GetDefaultCellStyle()->Clone("CL");

	oStyle = oControl->GetConfiguration()->GetDefaultCellStyle()->Clone("CR");
	oStyle->SetTextAlignmentHorizontal(HAL_RIGHT);

	oStyle = oControl->GetConfiguration()->GetDefaultCellStyle()->Clone("CH");
	oStyle->SetTextAlignmentHorizontal(HAL_RIGHT);
	oStyle->SetBorderColor(oTemplate.ControlBorderColor);


	oStyle = oControl->GetConfiguration()->GetDefaultNodeStyle();
	oStyle->GetFont()->SetPtFont(oControl->GetConfiguration()->GetDefaultFont()->GetFontName(), 8, FS_FONTSTYLEREGULAR);
	oStyle->SetAppearance(SA_FLAT);
	oStyle->SetBackgroundMode(FP_SOLID);
	oStyle->SetBackColor(Color::White);
	oStyle->SetBorderColor(oTemplate.RowBorderColor);
	oStyle->SetBorderStyle(SBR_CUSTOM);
	oStyle->GetCustomBorderStyle()->SetTop(FALSE);
	oStyle->GetCustomBorderStyle()->SetLeft(FALSE);
	oStyle->SetForeColor(oTemplate.RowForeColor);

	oStyle = oControl->GetConfiguration()->GetDefaultNodeStyle()->Clone("NR");

	oStyle = oControl->GetConfiguration()->GetDefaultNodeStyle()->Clone("NB");
	oStyle->GetFont()->SetPtFont(oControl->GetConfiguration()->GetDefaultFont()->GetFontName(), 8, FS_FONTSTYLEBOLD);


	oStyle = oControl->GetConfiguration()->GetDefaultClientAreaStyle();
	oStyle->SetAppearance(SA_FLAT);
	oStyle->SetBackColor(Color::White);
	oStyle->SetBorderColor(oTemplate.RowBorderColor);
	oStyle->SetBorderStyle(SBR_CUSTOM);
	oStyle->GetCustomBorderStyle()->SetTop(FALSE);
	oStyle->GetCustomBorderStyle()->SetLeft(FALSE);
	oStyle->GetCustomBorderStyle()->SetRight(FALSE);

	oControl->GetTreeview()->SetPlusMinusBorderColor(oTemplate.TreeviewColor);
	oControl->GetTreeview()->SetPlusMinusSignColor(oTemplate.TreeviewColor);
	oControl->GetTreeview()->SetCheckBoxBorderColor(oTemplate.TreeviewColor);
	oControl->GetTreeview()->SetCheckBoxMarkColor(oTemplate.TreeviewColor);
	oControl->GetTreeview()->SetTreeLineColor(oTemplate.TreeviewColor);
	oControl->GetTreeview()->SetCheckBoxBackColor(oTemplate.TreeviewColor);
	oControl->GetTreeview()->SetPlusMinusBackColor(oTemplate.TreeviewColor);

	oControl->GetSplitter()->SetType(SA_USERDEFINED);
	oControl->GetSplitter()->SetWidth(1);
	oControl->GetSplitter()->SetColor(1, oTemplate.ControlBorderColor);

	oStyle = oControl->GetConfiguration()->GetDefaultSBSeparatorStyle();
	oStyle->SetAppearance(SA_FLAT);
	oStyle->SetBackgroundMode(FP_SOLID);
	oStyle->SetBackColor(Color::White);
	oStyle->SetBorderStyle(SBR_SINGLE);
	oStyle->SetBorderColor(oTemplate.RowBorderColor);

	oStyle = oControl->GetConfiguration()->GetDefaultScrollBarStyle();
	oStyle->SetAppearance(SA_FLAT);
	oStyle->SetBackgroundMode(FP_SOLID);
	oStyle->SetBackColor(Color::White);
	oStyle->SetBorderStyle(SBR_SINGLE);
	oStyle->SetBorderColor(oTemplate.RowBorderColor);
	oStyle->GetScrollBarStyle()->SetArrowColor(Color::Black);

	oStyle = oControl->GetStyles()->Add("AB");
	oStyle->SetAppearance(SA_FLAT);
	oStyle->SetBackgroundMode(FP_SOLID);
	oStyle->SetBackColor(Color::White);
	oStyle->SetBorderStyle(SBR_SINGLE);
	oStyle->SetBorderColor(oTemplate.RowBorderColor);
	oStyle->GetScrollBarStyle()->SetArrowColor(Color::Black);

	oStyle = oControl->GetStyles()->Add("TBH");
	oStyle->SetAppearance(SA_FLAT);
	oStyle->SetBackgroundMode(FP_GRADIENT);
	oStyle->SetGradientFillMode(GDT_VERTICAL);
	oStyle->SetStartGradientColor(Color::White);
	oStyle->SetEndGradientColor(oTemplate.HeadingBackColor);
	oStyle->SetBorderStyle(SBR_SINGLE);
	oStyle->SetBorderColor(oTemplate.RowBorderColor);
	oStyle->SetBackColor(Color::White);

	oStyle = oControl->GetStyles()->Item("TBH")->Clone("TBHP");
	oStyle->SetStartGradientColor(oTemplate.HeadingBackColor);
	oStyle->SetEndGradientColor(Color::White);

	oStyle = oControl->GetStyles()->Add("TBV");
	oStyle->SetAppearance(SA_FLAT);
	oStyle->SetBackgroundMode(FP_GRADIENT);
	oStyle->SetGradientFillMode(GDT_HORIZONTAL);
	oStyle->SetStartGradientColor(Color::White);
	oStyle->SetEndGradientColor(oTemplate.HeadingBackColor);
	oStyle->SetBorderStyle(SBR_SINGLE);
	oStyle->SetBorderColor(oTemplate.RowBorderColor);
	oStyle->SetBackColor(Color::White);

	oStyle = oControl->GetStyles()->Item("TBV")->Clone("TBVP");
	oStyle->SetStartGradientColor(oTemplate.HeadingBackColor);
	oStyle->SetEndGradientColor(Color::White);

	oControl->GetVerticalScrollBar()->GetScrollBar()->GetArrowButtons()->SetNormalStyleIndex("AB");
	oControl->GetVerticalScrollBar()->GetScrollBar()->GetArrowButtons()->SetPressedStyleIndex("AB");
	oControl->GetVerticalScrollBar()->GetScrollBar()->GetArrowButtons()->SetDisabledStyleIndex("AB");
	oControl->GetVerticalScrollBar()->GetScrollBar()->GetThumbButton()->SetNormalStyleIndex("TBV");
	oControl->GetVerticalScrollBar()->GetScrollBar()->GetThumbButton()->SetPressedStyleIndex("TBVP");
	oControl->GetVerticalScrollBar()->GetScrollBar()->GetThumbButton()->SetDisabledStyleIndex("TBV");

	oControl->GetHorizontalScrollBar()->GetScrollBar()->GetArrowButtons()->SetNormalStyleIndex("AB");
	oControl->GetHorizontalScrollBar()->GetScrollBar()->GetArrowButtons()->SetPressedStyleIndex("AB");
	oControl->GetHorizontalScrollBar()->GetScrollBar()->GetArrowButtons()->SetDisabledStyleIndex("AB");
	oControl->GetHorizontalScrollBar()->GetScrollBar()->GetThumbButton()->SetNormalStyleIndex("TBH");
	oControl->GetHorizontalScrollBar()->GetScrollBar()->GetThumbButton()->SetPressedStyleIndex("TBHP");
	oControl->GetHorizontalScrollBar()->GetScrollBar()->GetThumbButton()->SetDisabledStyleIndex("TBH");

	oControl->GetTimeLineScrollBar()->GetScrollBar()->GetArrowButtons()->SetNormalStyleIndex("AB");
	oControl->GetTimeLineScrollBar()->GetScrollBar()->GetArrowButtons()->SetPressedStyleIndex("AB");
	oControl->GetTimeLineScrollBar()->GetScrollBar()->GetArrowButtons()->SetDisabledStyleIndex("AB");
	oControl->GetTimeLineScrollBar()->GetScrollBar()->GetThumbButton()->SetNormalStyleIndex("TBH");
	oControl->GetTimeLineScrollBar()->GetScrollBar()->GetThumbButton()->SetPressedStyleIndex("TBHP");
	oControl->GetTimeLineScrollBar()->GetScrollBar()->GetThumbButton()->SetDisabledStyleIndex("TBH");

	clsTimeLine* oTimeLine = oControl->GetCurrentViewObject()->GetTimeLine();

	oTimeLine->GetTimeLineScrollBar()->GetScrollBar()->GetArrowButtons()->SetNormalStyleIndex("AB");
	oTimeLine->GetTimeLineScrollBar()->GetScrollBar()->GetArrowButtons()->SetPressedStyleIndex("AB");
	oTimeLine->GetTimeLineScrollBar()->GetScrollBar()->GetArrowButtons()->SetDisabledStyleIndex("AB");
	oTimeLine->GetTimeLineScrollBar()->GetScrollBar()->GetThumbButton()->SetNormalStyleIndex("TBH");
	oTimeLine->GetTimeLineScrollBar()->GetScrollBar()->GetThumbButton()->SetPressedStyleIndex("TBHP");
	oTimeLine->GetTimeLineScrollBar()->GetScrollBar()->GetThumbButton()->SetDisabledStyleIndex("TBH");

	oStyle = oControl->GetConfiguration()->GetDefaultTimeBlockStyle();
	oStyle->SetBorderStyle(SBR_NONE);
	oStyle->SetAppearance(SA_FLAT);
	oStyle->SetBackgroundMode(FP_SOLID);
	oStyle->SetBackColor(oTemplate.TimeBlockBackColor);
}