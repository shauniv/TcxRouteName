#pragma once
#include <windows.h>
#include <objbase.h>
#include <atlbase.h>
#include <objbase.h>


namespace XmlHelpers
{
    HRESULT GetString(IXMLDOMElement* pXmlDomElement, PCWSTR pszQueryString, BSTR* pbstrResult)
    {
        *pbstrResult = NULL;

        CComPtr<IXMLDOMNode> spXmlDomNode;
        HRESULT hr = pXmlDomElement->selectSingleNode(CComBSTR(pszQueryString), &spXmlDomNode);
        if (hr == S_OK)
        {
            hr = spXmlDomNode->get_text(pbstrResult);
        }
        else
        {
            hr = HRESULT_FROM_WIN32(ERROR_NOT_FOUND);
        }
        return hr;
    }

    HRESULT SetString(IXMLDOMElement* pXmlDomElement, PCWSTR pszQueryString, PCWSTR pszValue)
    {
        CComPtr<IXMLDOMNode> spXmlDomNode;
        HRESULT hr = pXmlDomElement->selectSingleNode(CComBSTR(pszQueryString), &spXmlDomNode);
        if (hr == S_OK)
        {
            hr = spXmlDomNode->put_text(CComBSTR(pszValue));
        }
        else
        {
            hr = HRESULT_FROM_WIN32(ERROR_NOT_FOUND);
        }
        return hr;
    }

    HRESULT GetString(IXMLDOMDocument* pXmlDomDocument, PCWSTR pszQueryString, BSTR* pbstrResult)
    {
        *pbstrResult = NULL;

        CComPtr<IXMLDOMElement> spDocumentRoot;
        HRESULT hr = pXmlDomDocument->get_documentElement(&spDocumentRoot);
        if (SUCCEEDED(hr))
        {
            hr = GetString(spDocumentRoot, pszQueryString, pbstrResult);
        }
        return hr;
    }

    HRESULT SetString(IXMLDOMDocument* pXmlDomDocument, PCWSTR pszQueryString, PCWSTR pszValue)
    {
        CComPtr<IXMLDOMElement> spDocumentRoot;
        HRESULT hr = pXmlDomDocument->get_documentElement(&spDocumentRoot);
        if (SUCCEEDED(hr))
        {
            hr = SetString(spDocumentRoot, pszQueryString, pszValue);
        }
        return hr;
    }
}

