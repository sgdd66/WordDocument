
// stdafx.h : ��׼ϵͳ�����ļ��İ����ļ���
// ���Ǿ���ʹ�õ��������ĵ�
// �ض�����Ŀ�İ����ļ�

#pragma once

#ifndef _SECURE_ATL
#define _SECURE_ATL 1
#endif

#ifndef VC_EXTRALEAN
#define VC_EXTRALEAN            // �� Windows ͷ���ų�����ʹ�õ�����
#endif

#include "targetver.h"

#define _ATL_CSTRING_EXPLICIT_CONSTRUCTORS      // ĳЩ CString ���캯��������ʽ��

// �ر� MFC ��ĳЩ�����������ɷ��ĺ��Եľ�����Ϣ������
#define _AFX_ALL_WARNINGS

#include <afxwin.h>         // MFC ��������ͱ�׼���
#include <afxext.h>         // MFC ��չ


#include <afxdisp.h>        // MFC �Զ�����



#ifndef _AFX_NO_OLE_SUPPORT
#include <afxdtctl.h>           // MFC �� Internet Explorer 4 �����ؼ���֧��
#endif
#ifndef _AFX_NO_AFXCMN_SUPPORT
#include <afxcmn.h>             // MFC �� Windows �����ؼ���֧��
#endif // _AFX_NO_AFXCMN_SUPPORT

#include <afxcontrolbars.h>     // �������Ϳؼ����� MFC ֧��



#ifndef _WORD
#define _WORD 1
#include "WordApp.h"
#include"WordDoc.h"
#include"WordDocs.h"
#include"WordSel.h"
#include"WordFont.h"
#include"WordRange.h"
#include"WordTable.h"
#include"WordTables.h"
#include"Paragraphs.h"
#include"Paragraph.h"
#include"Cells.h"
#include"InlineShape.h"
#include"InlineShapes.h"
#include"Shading.h"
#endif

#ifndef _MYTABLE
#define _MYTABLE 1
#include "MyTable.h"
#endif

#ifndef _TABLEADAPT
#define _TABLEADAPT 1
#include"TableAdapt.h"
#endif

#ifndef TURBINEDOC
#define TURBINEDOC 1
#include"TurbineDoc.h"
#endif

#ifndef WORDNAME
#define WORDNAME 1
#define  wdCell 12
#define wdTable 15
#define wdLine 5
#define wdCharacter 1
#define wdMove 0
#define wdExtend 1
#define wdToggle 9999998
#define wdAlignParagraphCenter 1
#define wdAlignParagraphLeft 0
#define wdStory 6
#define wdItem 16
#define wdRow 10
#define wdColumn 9
#define wdCharacterFormating 13
#define wdParagraphFormating 14
#define wdParagraph 4
#define wdScreen 7
#define wdSection 8
#define wdSentence 3
#define wdStory 6
#define wdWindow 11
#define wdWord 2
#define wdCellAlignVerticalCenter 1

#endif


#ifdef _UNICODE
#if defined _M_IX86
#pragma comment(linker,"/manifestdependency:\"type='win32' name='Microsoft.Windows.Common-Controls' version='6.0.0.0' processorArchitecture='x86' publicKeyToken='6595b64144ccf1df' language='*'\"")
#elif defined _M_X64
#pragma comment(linker,"/manifestdependency:\"type='win32' name='Microsoft.Windows.Common-Controls' version='6.0.0.0' processorArchitecture='amd64' publicKeyToken='6595b64144ccf1df' language='*'\"")
#else
#pragma comment(linker,"/manifestdependency:\"type='win32' name='Microsoft.Windows.Common-Controls' version='6.0.0.0' processorArchitecture='*' publicKeyToken='6595b64144ccf1df' language='*'\"")
#endif
#endif


