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

#include <msxml2.h>
#include <iostream>
#include <fstream>

typedef enum PE_LEVEL
{ 
    LVL_CONTROL = 0,
    LVL_FONT = 5,
	LVL_DATETIME = 6,
}PE_LEVEL;

class clsXML  
{
public:
	IXMLDOMDocument2* xDoc;
    IXMLDOMElement* oControlElement; 
    IXMLDOMElement* oFontElement;
	IXMLDOMElement* oDateTimeElement;
	CString mp_sObject;
	PE_LEVEL mp_yLevel;
	HRESULT mp_hr;
	CString mp_sNamespace;

	BOOL mp_bBoolsAreNumeric;
	BOOL mp_bSupportOptional;

	clsXML(CString sObject);
	clsXML();
	virtual ~clsXML();


	void InitializeWriter();
	void InitializeReader();

	void WriteXML(CString sPath);
	void ReadXML(CString sPath);

	IXMLDOMElement* ParentElement();
	IXMLDOMElement* mp_oCreateEmptyDOMElement(CString sElementName);
	IXMLDOMElement* GetDocumentElement(CString sTagName, int lIndex);
	IXMLDOMElement* GetElement(CString sTagName, int lIndex);

	void SetXML(CString sXML);
	CString GetXML();
	
	long ReadCollectionCount();
	
	CString ReadObject(CString sObjectName);
	CString ReadCollectionObject(int lIndex);
	void ReadProperty(CString sTagName, long& sElementValue);
	void ReadProperty(CString sTagName, BYTE& sElementValue);
	void ReadProperty(CString sTagName, short& sElementValue);
	void ReadProperty(CString sTagName, CString& sElementValue);
	void ReadProperty(CString sTagName, BOOL& sElementValue);
	void ReadProperty(CString sTagName, COleDateTime& sElementValue);
	void ReadProperty(CString sTagName, COleCurrency& sElementValue);
	void ReadProperty(CString sTagName, float& sElementValue);
	void ReadProperty(CString sTagName, double& sElementValue);
	void ReadProperty(CString sTagName, Time* oElementValue);
	void ReadProperty(CString sTagName, Duration* oElementValue);
	void ReadProperty(CString sTagName, DECIMAL oElementValue);


	void WriteObject(CString sObjectText);
	void WriteProperty(CString sElementName, long lElementValue);
	void WriteProperty(CString sElementName, BYTE lElementValue);
	void WriteProperty(CString sElementName, short iElementValue);
	void WriteProperty(CString sElementName, CString sElementValue);
	void WriteProperty(CString sElementName, BOOL bElementValue);
	void WriteProperty(CString sElementName, COleDateTime dtElementValue);
	void WriteProperty(CString sElementName, Time* oElementValue);
	void WriteProperty(CString sElementName, Duration* oElementValue);
	void WriteProperty(CString sElementName, DECIMAL oElementValue);

	void WriteProperty(CString sElementName, float fElementValue);
	void WriteProperty(CString sElementName, double fElementValue);

	void SetBoolsAreNumeric(BOOL newValue);
	BOOL GetBoolsAreNumeric(void);
	void SetSupportOptional(BOOL newValue);
	BOOL GetSupportOptional(void);

	
	CString GetCollectionObjectName(LONG lIndex);
	void RemoveNamespace(void);
	void AddNamespace(void);
	

	bool FormatDOMDocument (IXMLDOMDocument *pDoc, IStream *pStream);

};