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


// ControlTemplateGradient command target

class ControlTemplateGradient : public CCmdTarget
{
	DECLARE_DYNCREATE(ControlTemplateGradient)

public:
	ControlTemplateGradient();
	virtual ~ControlTemplateGradient();

	virtual void OnFinalRelease();

    Gdiplus::Color ControlBorderColor;
    Gdiplus::Color HeadingBorderColor;
    Gdiplus::Color HeadingForeColor;
    Gdiplus::Color RowForeColor;
    LONG GradientFillMode;
    Gdiplus::Color StartGradientColor;
    Gdiplus::Color EndGradientColor;
    Gdiplus::Color TreeviewColor;
    Gdiplus::Color RowBorderColor;





protected:
	DECLARE_MESSAGE_MAP()
	DECLARE_OLECREATE(ControlTemplateGradient)
	DECLARE_DISPATCH_MAP()
	DECLARE_INTERFACE_MAP()

	OLE_COLOR odl_GetControlBorderColor(void)
	{
		
		return (OLE_COLOR) ControlBorderColor.ToCOLORREF();
	}

	void odl_SetControlBorderColor(OLE_COLOR Value)
	{
		
		ControlBorderColor.SetFromCOLORREF(g_TranslateColor(Value));
	}

    OLE_COLOR odl_GetHeadingBorderColor(void)
	{
		
		return (OLE_COLOR) HeadingBorderColor.ToCOLORREF();
	}

	void odl_SetHeadingBorderColor(OLE_COLOR Value)
	{
		
		HeadingBorderColor.SetFromCOLORREF(g_TranslateColor(Value));
	}

    OLE_COLOR odl_GetHeadingForeColor(void)
	{
		
		return (OLE_COLOR) HeadingForeColor.ToCOLORREF();
	}

	void odl_SetHeadingForeColor(OLE_COLOR Value)
	{
		
		HeadingForeColor.SetFromCOLORREF(g_TranslateColor(Value));
	}

    OLE_COLOR odl_GetRowForeColor(void)
	{
		
		return (OLE_COLOR) RowForeColor.ToCOLORREF();
	}

	void odl_SetRowForeColor(OLE_COLOR Value)
	{
		
		RowForeColor.SetFromCOLORREF(g_TranslateColor(Value));
	}

    OLE_COLOR odl_GetStartGradientColor(void)
	{
		
		return (OLE_COLOR) StartGradientColor.ToCOLORREF();
	}

	void odl_SetStartGradientColor(OLE_COLOR Value)
	{
		
		StartGradientColor.SetFromCOLORREF(g_TranslateColor(Value));
	}

    OLE_COLOR odl_GetEndGradientColor(void)
	{
		
		return (OLE_COLOR) EndGradientColor.ToCOLORREF();
	}

	void odl_SetEndGradientColor(OLE_COLOR Value)
	{
		
		EndGradientColor.SetFromCOLORREF(g_TranslateColor(Value));
	}

    OLE_COLOR odl_GetTreeviewColor(void)
	{
		
		return (OLE_COLOR) TreeviewColor.ToCOLORREF();
	}

	void odl_SetTreeviewColor(OLE_COLOR Value)
	{
		
		TreeviewColor.SetFromCOLORREF(g_TranslateColor(Value));
	}

    OLE_COLOR odl_GetRowBorderColor(void)
	{
		
		return (OLE_COLOR) RowBorderColor.ToCOLORREF();
	}

	void odl_SetRowBorderColor(OLE_COLOR Value)
	{
		
		RowBorderColor.SetFromCOLORREF(g_TranslateColor(Value));
	}

	LONG odl_GetGradientFillMode(void)
	{
		
		return GradientFillMode;
	}

	void odl_SetGradientFillMode(LONG Value)
	{
		
		GradientFillMode = Value;
	}


};