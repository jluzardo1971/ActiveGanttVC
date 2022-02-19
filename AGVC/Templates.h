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

class S_ObjectTemplate
{
public:

	S_ObjectTemplate()
	{

	}

	LONG yBackgroundMode_Tasks;
	LONG yGradientFillMode_Tasks;
		
	Gdiplus::Color BackColor_T1;
	Gdiplus::Color ForeColor_T1;
	Gdiplus::Color BorderColor_T1;
	Gdiplus::Color StartGradientColor_T1;
	Gdiplus::Color EndGradientColor_T1;
	LONG HatchStyle_T1;
	Gdiplus::Color HatchBackColor_T1;

	Gdiplus::Color HatchForeColor_T1;
	Gdiplus::Color BackColor_T2;
	Gdiplus::Color ForeColor_T2;
	Gdiplus::Color BorderColor_T2;
	Gdiplus::Color StartGradientColor_T2;
	Gdiplus::Color EndGradientColor_T2;
	LONG HatchStyle_T2;
	Gdiplus::Color HatchBackColor_T2;

	Gdiplus::Color HatchForeColor_T2;
	Gdiplus::Color BackColor_T3;
	Gdiplus::Color ForeColor_T3;
	Gdiplus::Color BorderColor_T3;
	Gdiplus::Color StartGradientColor_T3;
	Gdiplus::Color EndGradientColor_T3;
	LONG HatchStyle_T3;
	Gdiplus::Color HatchBackColor_T3;

	Gdiplus::Color HatchForeColor_T3;
	Gdiplus::Color BackColor_T4;
	Gdiplus::Color ForeColor_T4;
	Gdiplus::Color BorderColor_T4;
	Gdiplus::Color StartGradientColor_T4;
	Gdiplus::Color EndGradientColor_T4;
	LONG HatchStyle_T4;
	Gdiplus::Color HatchBackColor_T4;
	Gdiplus::Color HatchForeColor_T4;
};

void g_ApplyTemplate(CActiveGanttVCCtl* oControl, LONG yControlTemplate, LONG yObjectTemplate);

void mp_ControlTemplateSelector(CActiveGanttVCCtl* oControl, LONG yControlTemplate);

void mp_ObjectTemplateSelector(CActiveGanttVCCtl* oControl, LONG yObjectTemplate);

void mp_ApplyObjectTemplate(CActiveGanttVCCtl* oControl, S_ObjectTemplate &oTemplate);

void g_ApplyTemplate_Gradient(CActiveGanttVCCtl* oControl, ControlTemplateGradient &oTemplate, LONG yObjectTemplate);

void mp_ApplyControlTemplate_Gradient(CActiveGanttVCCtl* oControl, ControlTemplateGradient &oTemplate);

void g_ApplyTemplate_Solid(CActiveGanttVCCtl* oControl, ControlTemplateSolid &oTemplate, LONG yObjectTemplate);

void mp_ApplyControlTemplate_Solid(CActiveGanttVCCtl* oControl, ControlTemplateSolid &oTemplate);