#include <iostream>
#include <fstream>
#include <sstream>
#include <string>
#include <fstream>
#include <cctype>
#include <iomanip>
#include <vector>
#include <stdio.h>
#include <stdlib.h>
#include <assert.h>
#include <codecvt>
#include <locale>
#include <thread>
#include <tchar.h>
#import <msxml6.dll>
#include <msxml6.h>
#include "comutil.h"
#include <filesystem>
#include <windows.h>

#define CHK_HR(stmt)        do { hr=(stmt); if (FAILED(hr)) goto CleanUp; } while(0)

std::wstring ReplaceString
(
    std::wstring String1  // ’u‚«Š·‚¦‘ÎÛ
    , std::wstring String2  // ŒŸõ‘ÎÛ
    , std::wstring String3  // ’u‚«Š·‚¦‚é“à—e
){
    std::wstring::size_type  Pos(String1.find(String2));

    while (Pos != std::string::npos)
    {
        String1.replace(Pos, String2.length(), String3);
        Pos = String1.find(String2, Pos + String3.length());
    }

    return String1;
}

void queryNodesSmart(const wchar_t* file)
{
    MSXML2::IXMLDOMDocumentPtr pXMLDom;
    HRESULT hr = pXMLDom.CreateInstance(__uuidof(MSXML2::DOMDocument60), NULL, CLSCTX_INPROC_SERVER);
    if (FAILED(hr)) {
        printf("Failed to instantiate XML DOM 1.\n");
        return;
    }

    try {
        std::string nodeLabelVal = "";
        std::string nodeDescriptionVal = "";
        std::string nodeLabelVal2 = "";
        std::string nodeDescriptionVal2 = "";
        pXMLDom->async = VARIANT_FALSE;
        pXMLDom->validateOnParse = VARIANT_FALSE;
        pXMLDom->resolveExternals = VARIANT_FALSE;

        if (pXMLDom->load(file) != VARIANT_TRUE) {// extension must be .xml
        
            CHK_HR(pXMLDom->parseError->errorCode);
            printf("Failed to load DOM from xml file.\n%s\n",
                (LPCSTR)pXMLDom->parseError->Getreason());
        }
        //pXMLDom.setProperty('SelectionNamespaces',
         //   'xmlns:pf="http://soap.sforce.com/2006/04/metadata/"');

        //pXMLDom.setProperty(L"SelectionNamespaces", L"xmlns="http://soap.sforce.com/2006/04/metadata");
        //pXMLDom.setProperty(L"ProhibitDTD", FALSE);

        // Query root node.
        MSXML2::IXMLDOMNodePtr pNode_root= pXMLDom->selectSingleNode(L"ReportType");
        if (pNode_root) {
            MSXML2::IXMLDOMNodePtr pNode_label = pNode_root->selectSingleNode(L"label");
            if (pNode_label) {
                nodeLabelVal = (LPCSTR)pNode_label->Gettext();
                //printf("Result from selectSingleNode:\nNode, <%s>:\n\t%s\n\n",
                //    (LPCSTR)pNode_label->nodeName, (LPCSTR)pNode_label->Gettext());
            }
            else{
                printf("No label node is fetched in DOM 1.\n");
            }
            //_variant_t strVar
            MSXML2::IXMLDOMNodePtr pNode_description = pNode_root->selectSingleNode(L"description");
            if (pNode_description) {
                nodeDescriptionVal = (LPCSTR)pNode_description->Gettext();
                //printf("Result from selectSingleNode:\nNode, <%s>:\n\t%s\n\n",
                //    (LPCSTR)pNode_description->nodeName, (LPCSTR)pNode_description->Gettext());//(LPCSTR)pNode_description->nodeValue
            }
            else {
                printf("No label node is fetched in DOM 1.\n");
            }
        }
        else {
            printf("No Root node is fetched in DOM 1.\n");
        }

        printf("Result from XML1:\n \tNode\n \tLabel : <%s>\n \tDescription: <%s> \n", 
            nodeLabelVal.c_str(), nodeDescriptionVal.c_str());

        std::wstring file2 = ReplaceString(file, L"ReportType_migfull", L"ReportType_hhtdev2");
        std::wstring file3 = ReplaceString(file, L"ReportType_migfull", L"ReportType_hhtdev3");
        wprintf(L"Out put file path: %s \n\n ", file3.c_str());
       
        MSXML2::IXMLDOMDocumentPtr pXMLDom2;
        HRESULT hr2 = pXMLDom2.CreateInstance(__uuidof(MSXML2::DOMDocument60), NULL, CLSCTX_INPROC_SERVER);
        if (FAILED(hr2)) {
            printf("Failed to instantiate  XML DOM 2.\n");
            return;
        }
        else {
            pXMLDom2->async = VARIANT_FALSE;
            pXMLDom2->validateOnParse = VARIANT_FALSE;
            pXMLDom2->resolveExternals = VARIANT_FALSE;

            if (pXMLDom2->load(file2.c_str()) != VARIANT_TRUE) {// extension must be .xml
                CHK_HR(pXMLDom2->parseError->errorCode);
                printf("Failed to load DOM from xml file 1.\n%s\n",
                    (LPCSTR)pXMLDom2->parseError->Getreason());
            }

            // Query a single node.
            MSXML2::IXMLDOMNodePtr pNode_root2 = pXMLDom2->selectSingleNode(L"ReportType");
            if (pNode_root2) {
                MSXML2::IXMLDOMNodePtr pNode_label2 = pNode_root2->selectSingleNode(L"label");
                if (pNode_label2) {
                    nodeLabelVal2 = (LPCSTR)pNode_label2->Gettext();
                    _bstr_t bs1 = nodeLabelVal.c_str();
                    BSTR bstr = bs1.copy();
                    pNode_label2->put_text(bstr);
                }
                else {
                    printf("No label node is fetched in DOM2.\n");
                }
                MSXML2::IXMLDOMNodePtr pNode_description2 = pNode_root2->selectSingleNode(L"description");
                if (pNode_description2) {
                    nodeDescriptionVal2 = (LPCSTR)pNode_description2->Gettext();
                    _bstr_t bs1 = nodeDescriptionVal.c_str();
                    BSTR bstr = bs1.copy();
                    pNode_description2->put_text(bstr);
                }
                else {
                    printf("No description node is fetched in DOM2.\n");
                }
            }
            else {
                printf("No Root node is fetched in DOM 1.\n");
            }

            printf("Result from XML2:\n \tNode\n \tLabel : <%s>\n \tDescription: <%s> \n\n\n\n",
                nodeLabelVal2.c_str(), nodeDescriptionVal2.c_str());


            _variant_t fil2v(file3.c_str());
            HRESULT hr3 = pXMLDom2->save(fil2v);
            if (FAILED(hr3)) {
                printf("Failed to save XML DOM 2.\n");
                return;
            }
        }


    }
    catch (_com_error errorObject){
        printf("Exception thrown, HRESULT: 0x%08x", errorObject.Error());
    }

CleanUp:
    return;
}


bool ListDirectoryContents(const wchar_t* sDir)
{
    WIN32_FIND_DATA fdFile;
    HANDLE hFind = NULL;

    wchar_t sPath[2048];

    //Specify a file mask. *.* = We want everything!
    wsprintf(sPath, L"%s\\*.*", sDir);

    if ((hFind = FindFirstFile(sPath, &fdFile)) == INVALID_HANDLE_VALUE)
    {
        wprintf(L"Path not found: [%s]\n", sDir);
        return false;
    }
    unsigned int index = 0;
    do
    {
        //Find first file will always return "."
        //    and ".." as the first two directories.
        if (wcscmp(fdFile.cFileName, L".") != 0
            && wcscmp(fdFile.cFileName, L"..") != 0)
        {
            //Build up our file path using the passed in
            //  [sDir] and the file/foldername we just found:
            wsprintf(sPath, L"%s\\%s", sDir, fdFile.cFileName);

            //Is the entity a File or Folder?
            if (fdFile.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY)
            {
                wprintf(L"Directory: %s\n", sPath);
                ListDirectoryContents(sPath);
            }
            else {
                wprintf(L"File from: %s\n", sPath);
                
                //queryNodesSmart(sPath);

                /* rename file Start 
                std::wstring str = sPath;
                int position;
                position = str.find(L".");
                if (position != str.npos){
                    printf("position is : %d\n", position);
                    std::wstring str1 = str.substr(0, position)+L".reportType";
                    _wrename(sPath, str1.c_str());
                }
                 rename file end */

                /* replace line number 2 start  replace line number 2 end*/
                std::wstring temp;
                std::wstring sPath1 = sPath;

                std::wfstream seek1;
                //std::wfstream seek2;
                std::vector<std::wstring> lines;
                seek1.open(sPath1);
                unsigned int counter = 0;
                while (std::getline(seek1, temp)) {
                    if (index == 1) {
                        std::wcout << temp << std::endl;
                    }
                    //if (counter == 1) {
                    //    lines.push_back(L"<ReportType xmlns=\"http://soap.sforce.com/2006/04/metadata\">");
                    //}
                    //else {
                    
                    lines.push_back(temp);
                    //}
                    counter++;
                    
                }
                seek1.close();
                seek1.open(sPath1, std::ios::out | std::ios::trunc);

                for (unsigned int i = 0; i < lines.size(); i++) {
                    std::wstring wstr = lines.at(i);
                    wstr = ReplaceString(wstr, L"\t", L"    ");
                    seek1 << wstr << std::endl;
                }
                


            }
        }
        index++;
    } while (FindNextFile(hFind, &fdFile)); //Find the next file.

    FindClose(hFind);

    return true;
}

int main()
{
    setlocale(LC_ALL, "Japanese");
    std::cout << "Hello ‚ ‚ ‚ !\n";

    HRESULT hr = CoInitialize(NULL);
    if (SUCCEEDED(hr)) {
        
        ListDirectoryContents(L"C:\\work\\migfull\\ReportType_hhtdev3_after\\reportTypes");
        CoUninitialize();
    }
    return 0;

}
