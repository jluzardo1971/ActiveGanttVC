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
#include "clsXML.h"
#include <Shlwapi.h>
#if _MSC_VER <= 1200
	#include <fstream.h>
#else
	#include <fstream>
	using namespace std;
#endif
#include <sys/types.h>
#include <sys/stat.h>
#include <wchar.h>
#include <stdlib.h>

#ifdef _DEBUG
#undef THIS_FILE
static char THIS_FILE[]=__FILE__;
//#define new DEBUG_NEW
#endif

clsXML::clsXML(CString sObject)
{
	mp_sObject = sObject;
	mp_bBoolsAreNumeric = FALSE;
	mp_bSupportOptional = FALSE;
	mp_sNamespace = _T("");
	CoInitialize(NULL);
	CoCreateInstance(CLSID_DOMDocument60, NULL, CLSCTX_INPROC_SERVER, IID_IXMLDOMDocument, (void**)&xDoc);
	xDoc->put_async(FALSE);
}

clsXML::clsXML()
{

}

clsXML::~clsXML()
{
	xDoc->Release();
	CoUninitialize();
}

void clsXML::InitializeWriter()
{
	VARIANT_BOOL status;
	CString sBuff;
	sBuff = _T("<" + mp_sObject + "></" + mp_sObject + ">");
	BSTR zBuff = sBuff.AllocSysString();
    xDoc->loadXML(zBuff, &status);
	::SysFreeString(zBuff);
	oControlElement = GetDocumentElement(mp_sObject, 0);
	mp_yLevel = LVL_CONTROL;
}

void clsXML::InitializeReader()
{
	oControlElement = GetDocumentElement(mp_sObject, 0);
    mp_yLevel = LVL_CONTROL;
}

void clsXML::WriteXML(CString sPath)
{
	IStream* oStream;
	SHCreateStreamOnFile(sPath, STGM_WRITE | STGM_CREATE, &oStream);
	FormatDOMDocument(xDoc, oStream);
	oStream->Release();
}

void clsXML::ReadXML(CString sPath)
{
	VARIANT_BOOL status;
	VARIANT vSrc;
	xDoc->put_async(FALSE);
	VariantInit(&vSrc);
	BSTR zPath = sPath.AllocSysString();
	V_BSTR(&vSrc) = zPath;
	V_VT(&vSrc) = VT_BSTR;
	xDoc->load(vSrc, &status);
	::SysFreeString(zPath);
}

IXMLDOMElement* clsXML::ParentElement()
{
    switch(mp_yLevel)
	{
	case LVL_CONTROL:
		return oControlElement;
		break;
	case LVL_FONT:
		return oFontElement;
		break;
	case LVL_DATETIME:
		return oDateTimeElement;
		break;
	default:
		return NULL;
		break;
    }
}

IXMLDOMElement* clsXML::mp_oCreateEmptyDOMElement(CString sElementName)
{
	IXMLDOMElement* oElement;
	IXMLDOMNode* oNode;
	BSTR zElementName = sElementName.AllocSysString();
    xDoc->createElement(zElementName, &oElement);
	::SysFreeString(zElementName);
	oElement->QueryInterface(IID_IXMLDOMNode, (void**)& oNode);
    ParentElement()->appendChild(oNode, &oNode);
	return oElement;
}

IXMLDOMElement* clsXML::GetDocumentElement(CString sTagName, int lIndex)
{
	IXMLDOMNodeList* oNodeList;
	IXMLDOMNode* oNode;
	IXMLDOMElement* oElement;
	BSTR zTagName = sTagName.AllocSysString();
	xDoc->getElementsByTagName(zTagName, &oNodeList);
	::SysFreeString(zTagName);
	oNodeList->get_item(lIndex, &oNode);
	if (oNode != NULL)
	{
		oNode->QueryInterface(IID_IXMLDOMElement, (void**)& oElement);
		return oElement;
	}
	else
	{
		return NULL;
	}
}

IXMLDOMElement* clsXML::GetElement(CString sTagName, int lIndex)
{
	IXMLDOMNodeList* oNodeList;
	IXMLDOMNode* oNode;
	IXMLDOMElement* oElement;
	BSTR zTagName = sTagName.AllocSysString();
	ParentElement()->getElementsByTagName(zTagName, &oNodeList);
	::SysFreeString(zTagName);
	oNodeList->get_item(lIndex, &oNode);
	oNode->QueryInterface(IID_IXMLDOMElement, (void**)& oElement);
	return oElement;
}

CString clsXML::GetCollectionObjectName(LONG lIndex)
{
	IXMLDOMNodeList* oNodeList;
	IXMLDOMNode* oNode;
	ParentElement()->get_childNodes(&oNodeList);
	oNodeList->get_item(lIndex - 1, &oNode);
	BSTR zReturn = NULL;
	oNode->get_nodeName(&zReturn);
	CString sReturn = zReturn;
	::SysFreeString(zReturn);
	return sReturn;
}

bool clsXML::FormatDOMDocument (IXMLDOMDocument *pDoc, IStream *pStream)
{
    CComPtr <IMXWriter> pMXWriter; 
    if (FAILED (pMXWriter.CoCreateInstance(__uuidof (MXXMLWriter), NULL, CLSCTX_ALL))) 
    { 
        return false; 
    } 
    CComPtr <ISAXContentHandler> pISAXContentHandler; 
    if (FAILED (pMXWriter.QueryInterface(&pISAXContentHandler))) 
    { 
        return false; 
    } 
    CComPtr <ISAXErrorHandler> pISAXErrorHandler; 
    if (FAILED (pMXWriter.QueryInterface (&pISAXErrorHandler))) 
    { 
        return false; 
    } 
    CComPtr <ISAXDTDHandler> pISAXDTDHandler; 
    if (FAILED (pMXWriter.QueryInterface (&pISAXDTDHandler))) 
    { 
        return false; 
    } 
	//FAILED (pMXWriter ->put_standalone (VARIANT_FALSE)) ||
    if (FAILED (pMXWriter ->put_omitXMLDeclaration (VARIANT_FALSE)) ||  
        FAILED (pMXWriter ->put_indent (VARIANT_TRUE)) || 
        FAILED (pMXWriter ->put_encoding (L"UTF-8"))) 
    { 
        return false; 
    } 
    CComPtr <ISAXXMLReader> pSAXReader; 
    if (FAILED (pSAXReader.CoCreateInstance (__uuidof (SAXXMLReader), NULL, CLSCTX_ALL))) 
    { 
        return false; 
    } 
 
    if (FAILED (pSAXReader ->putContentHandler (pISAXContentHandler)) || 
        FAILED (pSAXReader ->putDTDHandler (pISAXDTDHandler)) || 
        FAILED (pSAXReader ->putErrorHandler (pISAXErrorHandler)) || 
        FAILED (pSAXReader ->putProperty ( 
        L"http://xml.org/sax/properties/lexical-handler", CComVariant (pMXWriter))) || 
        FAILED (pSAXReader ->putProperty ( 
        L"http://xml.org/sax/properties/declaration-handler", CComVariant (pMXWriter)))) 
    { 
        return false; 
    } 
 
    // Perform the write 
 
    return  
       SUCCEEDED (pMXWriter ->put_output (CComVariant (pStream))) && 
       SUCCEEDED (pSAXReader ->parse (CComVariant (pDoc))); 
}

CString clsXML::GetXML()
{
	CString sReturn = "";
	BSTR zReturn = NULL;
	xDoc->get_xml(&zReturn);
	sReturn = zReturn;
	::SysFreeString(zReturn);
	return sReturn;

}

void clsXML::SetXML(CString sXML)
{
	if (mp_bSupportOptional == FALSE)
	{
		VARIANT_BOOL status;
		CString sBuff;
		xDoc->put_async(FALSE);
		BSTR zXML = sXML.AllocSysString();
		xDoc->loadXML(zXML, &status);
		::SysFreeString(zXML);
	}
	else
	{
		if (sXML.GetLength() > 0)
		{
			VARIANT_BOOL status;
			CString sBuff;
			xDoc->put_async(FALSE);
			BSTR zXML = sXML.AllocSysString();
			xDoc->loadXML(zXML, &status);
			::SysFreeString(zXML);
		}
	}
}

long clsXML::ReadCollectionCount()
{
	if (mp_bSupportOptional == FALSE)
	{
		IXMLDOMNodeList* oNodeList;
		long lReturn = 0;
		ParentElement()->get_childNodes(&oNodeList);
		oNodeList->get_length(&lReturn);
		return lReturn;
	}
	else
	{
		IXMLDOMNodeList* oNodeList;
		long lReturn = 0;
		if (ParentElement() == NULL)
		{
			return 0;
		}
		else
		{
			ParentElement()->get_childNodes(&oNodeList);
			oNodeList->get_length(&lReturn);
			return lReturn;
		}
	}
}

CString clsXML::ReadObject(CString sObjectName)
{
	if (mp_bSupportOptional == FALSE)
	{
		IXMLDOMNodeList* oNodeList;
		IXMLDOMNode* oNode;
		CString sResult;
		sResult = "";
		BSTR zObjectName = sObjectName.AllocSysString();
		ParentElement()->getElementsByTagName(zObjectName, &oNodeList);
		::SysFreeString(zObjectName);
		oNodeList->get_item(0, &oNode);
		BSTR zResult = NULL;
		oNode->get_xml(&zResult);
		sResult = zResult;
		::SysFreeString(zResult);
		return sResult;
	}
	else
	{
		IXMLDOMNodeList* oNodeList;
		IXMLDOMNode* oNode;
		CString sResult;
		sResult = "";
		if (ParentElement() == NULL)
		{
			return "";
		}
		BSTR zObjectName = sObjectName.AllocSysString();
		ParentElement()->getElementsByTagName(zObjectName, &oNodeList);
		::SysFreeString(zObjectName);
		long lListLength = 0;
		oNodeList->get_length(&lListLength);
		if (lListLength > 0)
		{
			oNodeList->get_item(0, &oNode);
			BSTR zResult = NULL;
			oNode->get_xml(&zResult);
			sResult = zResult;
			::SysFreeString(zResult);
			return sResult;
		}
		else
		{
			return "";
		}

	}
}

CString clsXML::ReadCollectionObject(int lIndex)
{
	if (mp_bSupportOptional == FALSE)
	{
		IXMLDOMNodeList* oNodeList;
		IXMLDOMNode* oNode;
		CString sResult;
		
		sResult = "";
		ParentElement()->get_childNodes(&oNodeList);
		oNodeList->get_item(lIndex - 1, &oNode);
		BSTR zResult = NULL;
		oNode->get_xml(&zResult);
		sResult = zResult;
		::SysFreeString(zResult);
		return sResult;
	}
	else
	{
		IXMLDOMNodeList* oNodeList;
		IXMLDOMNode* oNode;
		CString sResult;
		sResult = "";
		if (ParentElement() == NULL)
		{
			return "";
		}
		ParentElement()->get_childNodes(&oNodeList);
		long lListLength = 0;
		oNodeList->get_length(&lListLength);
		if (lListLength > 0)
		{
			oNodeList->get_item(lIndex - 1, &oNode);
			BSTR zResult = NULL;
			oNode->get_xml(&zResult);
			sResult = zResult;
			::SysFreeString(zResult);
			return sResult;
		}
		else
		{
			return "";
		}

	}

}

void clsXML::ReadProperty(CString sTagName, long& sElementValue)
{
	CString sReturn;
	ReadProperty(sTagName, sReturn);
	if (sReturn.GetLength() == 0)
	{
		sReturn = "0";
	}
	sElementValue = g_CInt(sReturn);
}

void clsXML::ReadProperty(CString sTagName, BYTE& sElementValue)
{
	CString sReturn;
	ReadProperty(sTagName, sReturn);
	if (sReturn.GetLength() == 0)
	{
		sReturn = "0";
	}
	sElementValue = (BYTE) g_CInt(sReturn);
}

void clsXML::ReadProperty(CString sTagName, short& sElementValue)
{
	CString sReturn;
	ReadProperty(sTagName, sReturn);
	if (sReturn.GetLength() == 0)
	{
		sReturn = "0";
	}
	sElementValue = (short) g_CInt(sReturn);
}

void clsXML::ReadProperty(CString sTagName, BOOL& sElementValue)
{
	CString sReturn;
	ReadProperty(sTagName, sReturn);
	if (sReturn.GetLength() == 0)
	{
		sReturn = "0";
	}
	if (sReturn == _T("false") || sReturn == _T("0"))
	{
		sElementValue = FALSE;
	}
	else 
	{
		sElementValue = TRUE;
	}

}

void clsXML::ReadProperty(CString sTagName, CString& sElementValue)
{
	if (mp_bSupportOptional == FALSE)
	{
		IXMLDOMNodeList* oNodeList;
		IXMLDOMNode* oNode;
		VARIANT vReturn;
		VariantInit(&vReturn);
		BSTR zTagName = sTagName.AllocSysString();
		ParentElement()->getElementsByTagName(zTagName, &oNodeList);
		::SysFreeString(zTagName);
		oNodeList->get_item(0, &oNode);
		oNode->get_nodeTypedValue(&vReturn);
		sElementValue = vReturn.bstrVal;
	}
	else
	{
		IXMLDOMNodeList* oNodeList;
		IXMLDOMNode* oNode;
		IXMLDOMNode* oParentNode;
		VARIANT vReturn;
		VariantInit(&vReturn);
		BSTR sParentName1 = NULL;
		BSTR sParentName2 = NULL;
		if (ParentElement() == NULL)
		{
			return;
		}
		ParentElement()->get_nodeName(&sParentName1);
		BSTR zTagName = sTagName.AllocSysString();
		ParentElement()->getElementsByTagName(zTagName, &oNodeList);
		::SysFreeString(zTagName);
		long lListLength = 0;
		oNodeList->get_length(&lListLength);
		if (lListLength > 0)
		{
			oNodeList->get_item(0, &oNode);
			oNode->get_nodeTypedValue(&vReturn);
			oNode->get_parentNode(&oParentNode);
			oParentNode->get_nodeName(&sParentName2);
			if (_tcscmp(OLE2T(sParentName1), OLE2T(sParentName2)) == 0)
			{
				sElementValue = vReturn.bstrVal;
			}
			::SysFreeString(sParentName2);
		}
		::SysFreeString(sParentName1);
	}

}

void clsXML::ReadProperty(CString sTagName, COleDateTime& sElementValue)
{
	COleDateTime dtReturn;
	CString sReturn;
	CString sBuff;
	int lYear;
	int lMonth;
	int lDay;
	int lHour;
	int lMinute;
	int lSecond;
	ReadProperty(sTagName, sBuff);
	if (sBuff.GetLength() != 19)
	{
		sElementValue = (DATE)0;
	}
	else
	{
		sReturn = sBuff.Left(4);
		lYear = g_CInt(sReturn);
		sReturn = sBuff.Mid(5,2);
		lMonth = g_CInt(sReturn);
		sReturn = sBuff.Mid(8,2);
		lDay = g_CInt(sReturn);
		sReturn = sBuff.Mid(11,2);
		lHour = g_CInt(sReturn);
		sReturn = sBuff.Mid(14,2);
		lMinute = g_CInt(sReturn);
		sReturn = sBuff.Mid(17,2);
		lSecond = g_CInt(sReturn);
		dtReturn.SetDateTime(lYear,lMonth,lDay,lHour,lMinute,lSecond);
		sElementValue = dtReturn;
	}


}

void clsXML::ReadProperty(CString sTagName, COleCurrency& sElementValue)
{
	CString sBuff;
	COleCurrency cReturn;
	ReadProperty(sTagName, sBuff);
	if (sBuff.GetLength() == 0)
	{
		sBuff = "0";
	}
	cReturn.ParseCurrency(sBuff);
	sElementValue = cReturn;
}

void clsXML::ReadProperty(CString sTagName, float& sElementValue)
{
	CString sReturn;
	ReadProperty(sTagName, sReturn);
	if (sReturn.GetLength() == 0)
	{
		sReturn = "0";
	}
	sElementValue = (float) _tstof(sReturn);
}

void clsXML::ReadProperty(CString sTagName, double& sElementValue)
{
	CString sReturn;
	ReadProperty(sTagName, sReturn);
	if (sReturn.GetLength() == 0)
	{
		sReturn = "0";
	}
	sElementValue = _tstof(sReturn);
}

void clsXML::ReadProperty(CString sTagName, Time* oElementValue)
{
	CString sReturn;
	ReadProperty(sTagName, sReturn);
	oElementValue->FromString(sReturn);
}

void clsXML::ReadProperty(CString sTagName, Duration* oElementValue)
{
	CString sReturn;
	ReadProperty(sTagName, sReturn);
	oElementValue->FromString(sReturn);
}

void clsXML::ReadProperty(CString sTagName, DECIMAL oElementValue)
{
	CString sReturn;
	COleCurrency cReturn;
	ReadProperty(sTagName, sReturn);
	if (sReturn.GetLength() == 0)
	{
		sReturn = "0";
	}
	cReturn.ParseCurrency(sReturn);
	VarDecFromCy(cReturn, &oElementValue);
}

void clsXML::WriteObject(CString sObjectText)
{
	VARIANT_BOOL status;
	IXMLDOMDocument2* xDoc1;
	IXMLDOMElement* oNodeBuff;
	IXMLDOMNode* oNodeBuff1;
	mp_hr = CoCreateInstance(CLSID_DOMDocument60, NULL, CLSCTX_INPROC_SERVER, IID_IXMLDOMDocument, (void**)&xDoc1);
	xDoc1->put_async(FALSE);
	BSTR zObjectText = sObjectText.AllocSysString();
	xDoc1->loadXML(zObjectText, &status);
	::SysFreeString(zObjectText);
	xDoc1->get_documentElement(&oNodeBuff);
	oNodeBuff->QueryInterface(IID_IXMLDOMElement, (void**)& oNodeBuff1);
	ParentElement()->appendChild(oNodeBuff1, &oNodeBuff1);
	xDoc1->Release();
	CoUninitialize();
}

void clsXML::WriteProperty(CString sElementName, long lElementValue)
{
	CString sBuff;
	sBuff.Format(_T("%d"), lElementValue);
	WriteProperty(sElementName, sBuff);
}

void clsXML::WriteProperty(CString sElementName, BYTE lElementValue)
{
	CString sBuff;
	sBuff.Format(_T("%d"), lElementValue);
	WriteProperty(sElementName, sBuff);
}

void clsXML::WriteProperty(CString sElementName, short iElementValue)
{
	CString sBuff;
	sBuff.Format(_T("%d"), iElementValue);
	WriteProperty(sElementName, sBuff);
}

void clsXML::WriteProperty(CString sElementName, BOOL bElementValue)
{
	if (bElementValue == TRUE)
	{
		if (mp_bBoolsAreNumeric == TRUE)
		{
			WriteProperty(sElementName, _T("1"));
		}
		else
		{
			WriteProperty(sElementName, _T("true"));
		}
	}
	else
	{
		if (mp_bBoolsAreNumeric == TRUE)
		{
			WriteProperty(sElementName, _T("0"));
		}
		else
		{
			WriteProperty(sElementName, _T("false"));
		}
	}
}

void clsXML::WriteProperty(CString sElementName, CString vElementValue)
{
	//IXMLDOMElement* oNodeBuff;
	IXMLDOMNode* oNodeBuff;
	IXMLDOMNode* oNodeBuff1;
	VARIANT vBuff;
	VariantInit(&vBuff);
	BSTR zElementValue = vElementValue.AllocSysString();
	V_BSTR(&vBuff) = zElementValue;
	V_VT(&vBuff) = VT_BSTR;
	CComVariant varNodeType;
	varNodeType = NODE_ELEMENT;
	CString sNS = _T("");
	BSTR zElementName = sElementName.AllocSysString();
	BSTR zNS = sNS.AllocSysString();
	xDoc->createNode(varNodeType, zElementName, zNS, &oNodeBuff);
	::SysFreeString(zElementName);
	::SysFreeString(zNS);
	oNodeBuff->put_nodeTypedValue(vBuff);
	oNodeBuff->QueryInterface(IID_IXMLDOMNode, (void**)& oNodeBuff1);
    ParentElement()->appendChild(oNodeBuff1, &oNodeBuff1);
	::SysFreeString(zElementValue);
}

void clsXML::WriteProperty(CString sElementName, COleDateTime dtElementValue)
{
	CString sResult;
	CString sBuff;
	sBuff.Format(_T("%d"), dtElementValue.GetYear());
	while (sBuff.GetLength() < 4)
	{
		sBuff = _T("0") + sBuff;
	}
	sResult = sBuff + _T("-");
	sBuff.Format(_T("%d"), dtElementValue.GetMonth());
	while (sBuff.GetLength() < 2)
	{
		sBuff = _T("0") + sBuff;
	}
	sResult = sResult + sBuff + _T("-");
	sBuff.Format(_T("%d"), dtElementValue.GetDay());
	while (sBuff.GetLength() < 2)
	{
		sBuff = _T("0") + sBuff;
	}
	sResult = sResult + sBuff + _T("T");
	sBuff.Format(_T("%d"), dtElementValue.GetHour());
	while (sBuff.GetLength() < 2)
	{
		sBuff = _T("0") + sBuff;
	}
	sResult = sResult + sBuff + _T(":");
	sBuff.Format(_T("%d"), dtElementValue.GetMinute());
	while (sBuff.GetLength() < 2)
	{
		sBuff = _T("0") + sBuff;
	}
	sResult = sResult + sBuff + _T(":");
	sBuff.Format(_T("%d"), dtElementValue.GetSecond());
	while (sBuff.GetLength() < 2)
	{
		sBuff = _T("0") + sBuff;
	}
	sResult = sResult + sBuff;
	WriteProperty(sElementName, sResult);
}

void clsXML::WriteProperty(CString sElementName, float fElementValue)
{
	CString sBuff;
	sBuff.Format(_T("%.7f"), fElementValue);
	if (sBuff.Find(_T(".")) > -1)
	{
		while ((sBuff.Right(1) == "0" || sBuff.Right(1) == ".") && sBuff.GetLength() > 1)
		{
			CString sLast = sBuff.Right(1);
			sBuff = sBuff.Left(sBuff.GetLength() - 1);
			if (sLast == _T("."))
			{
				break;
			}
		}
	}
	WriteProperty(sElementName, sBuff);
}

void clsXML::WriteProperty(CString sElementName, double fElementValue)
{
	CString sBuff;
	sBuff.Format(_T("%.7f"), fElementValue);
	if (sBuff.Find(_T(".")) > -1)
	{
		while ((sBuff.Right(1) == "0" || sBuff.Right(1) == ".") && sBuff.GetLength() > 1)
		{
			CString sLast = sBuff.Right(1);
			sBuff = sBuff.Left(sBuff.GetLength() - 1);
			if (sLast == _T("."))
			{
				break;
			}
		}
	}
	WriteProperty(sElementName, sBuff);
}

void clsXML::WriteProperty(CString sElementName, Time* oElementValue)
{
	CString sBuff;
	sBuff = oElementValue->ToString();
	WriteProperty(sElementName, sBuff);
}
void clsXML::WriteProperty(CString sElementName, Duration* oElementValue)
{
	CString sBuff;
	sBuff = oElementValue->ToString();
	WriteProperty(sElementName, sBuff);
}
void clsXML::WriteProperty(CString sElementName, DECIMAL oElementValue)
{
	CY cBuff;
	CString sBuff = "";
	VarCyFromDec(&oElementValue, &cBuff);
	COleCurrency cBuff1 = cBuff;
	sBuff = cBuff1.Format();
	WriteProperty(sElementName, sBuff);
}

void clsXML::SetBoolsAreNumeric(BOOL newValue)
{
	mp_bBoolsAreNumeric = newValue;
}

BOOL clsXML::GetBoolsAreNumeric(void)
{
	return mp_bBoolsAreNumeric;
}

void clsXML::SetSupportOptional(BOOL newValue)
{
	mp_bSupportOptional = newValue;
}

BOOL clsXML::GetSupportOptional(void)
{
	return mp_bSupportOptional;
}

void clsXML::AddNamespace(void)
{
	BSTR zXML;
	xDoc->get_xml(&zXML);
	CString sXML = zXML;
	::SysFreeString(zXML);
	sXML.Replace(_T("<Project>"), _T("<Project xmlns=\"http://schemas.microsoft.com/project\">"));
	BSTR zXML1 = sXML.AllocSysString();
	xDoc->loadXML(zXML1, NULL);
	::SysFreeString(zXML1);
}

void clsXML::RemoveNamespace(void)
{
	BSTR zXML;
	xDoc->get_xml(&zXML);
	CString sXML = zXML;
	::SysFreeString(zXML);
	sXML.Replace(_T(" xmlns=\"http://schemas.microsoft.com/project\""), _T(""));
	BSTR zXML1 = sXML.AllocSysString();
	xDoc->loadXML(zXML1, NULL);
	::SysFreeString(zXML1);
}