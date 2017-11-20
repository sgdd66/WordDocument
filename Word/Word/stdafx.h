
// stdafx.h : 标准系统包含文件的包含文件，
// 或是经常使用但不常更改的
// 特定于项目的包含文件

#pragma once

#ifndef _SECURE_ATL
#define _SECURE_ATL 1
#endif

#ifndef VC_EXTRALEAN
#define VC_EXTRALEAN            // 从 Windows 头中排除极少使用的资料
#endif

#include "targetver.h"

#define _ATL_CSTRING_EXPLICIT_CONSTRUCTORS      // 某些 CString 构造函数将是显式的

// 关闭 MFC 对某些常见但经常可放心忽略的警告消息的隐藏
#define _AFX_ALL_WARNINGS

#include <afxwin.h>         // MFC 核心组件和标准组件
#include <afxext.h>         // MFC 扩展


#include <afxdisp.h>        // MFC 自动化类



#ifndef _AFX_NO_OLE_SUPPORT
#include <afxdtctl.h>           // MFC 对 Internet Explorer 4 公共控件的支持
#endif
#ifndef _AFX_NO_AFXCMN_SUPPORT
#include <afxcmn.h>             // MFC 对 Windows 公共控件的支持
#endif // _AFX_NO_AFXCMN_SUPPORT

#include <afxcontrolbars.h>     // 功能区和控件条的 MFC 支持



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


