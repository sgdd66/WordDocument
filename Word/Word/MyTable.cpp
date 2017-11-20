
#include"stdafx.h"
#ifndef _MYTABLE
#define _MYTABLE 1
#include "MyTable.h"
#endif



CMyTable::~CMyTable(void)
{
		CoUninitialize();
}

CMyTable::CMyTable(CTableAdapt* a)
{
	adapt=a;
	CoInitialize(NULL);
    COleVariant vOpt((long)DISP_E_PARAMNOTFOUND, VT_ERROR);
   
	CoInitialize(NULL);

    //开始一个新的Microsoft Word实例
    if (!oWordApp.CreateDispatch(_T("Word.Application")))
    {
        return;
    }
    oWordApp.put_Visible(TRUE);

    //创建一个新的word文档
    oDocs = oWordApp.get_Documents();
    oDoc = oDocs.Add(vOpt,vOpt, vOpt, vOpt);
    oSel = oWordApp.get_Selection();
	WordFont font(oSel.get_Font());
	font.put_Name(_T("仿宋_GB2312"));
    font.put_Color(RGB(0,0,0));
    font.put_Size(10.f);
 
}





void CMyTable::setTable(void)
{
	COleVariant vOpt((long)DISP_E_PARAMNOTFOUND, VT_ERROR);
    WordRange     Rng  = oSel.get_Range();
    WordTables    Tbls = oDoc.get_Tables(); 
	int hor=adapt->getHor();
	int ver=adapt->getVer();
    WordTable     Tbl  = Tbls.Add(Rng, hor, ver, COleVariant((short)1), vOpt);
			
}


void CMyTable::link(int hor, int ver, bool isRight, int num)
{
	moveHome();
	oSel.MoveRight(COleVariant((short)wdCell),COleVariant((short)ver),COleVariant((short)wdMove));
	oSel.MoveDown(COleVariant((short)wdLine),COleVariant((short)hor),COleVariant((short)wdMove));
	if(isRight)
		oSel.MoveRight(COleVariant((short)wdCharacter),COleVariant((short)num),COleVariant((short)wdExtend));
	else
		oSel.MoveDown(COleVariant((short)wdLine),COleVariant((short)(num-1)),COleVariant((short)wdExtend));

    Cells cell = oSel.get_Cells();
    cell.Merge();
}


void CMyTable::writeTable(void)
{
	moveHome();	
	int hor=adapt->getHor();
	int ver=adapt->getVer();
	for(int i=0;i<hor;i++){
		for(int j=0;j<ver;j++){
			oSel.TypeText(adapt->getString(i,j));
			if(i!=hor-1||j!=ver-1)
				oSel.MoveRight(COleVariant((short)wdCell),COleVariant((short)1),COleVariant((short)wdMove));
		}
	}
		oSel.MoveLeft(COleVariant((short)wdCharacter),COleVariant((short)2),COleVariant((short)wdMove));

}


void CMyTable::setTitle(CString title)
{
	this->title=title;
	WordFont font(oSel.get_Font());
	font.put_Name(_T("仿宋_GB2312"));
    font.put_Color(RGB(0,0,0));
    font.put_Size(19.f);
	font.put_Bold(wdToggle);
	Paragraphs para;
	 //获得对齐方式实例
    para = oSel.get_ParagraphFormat();
	//设置标题行为“居中”
	para.put_Alignment(wdAlignParagraphCenter);     
    //把文本添加到word文档
    oSel.TypeText(title);
    oSel.TypeParagraph();
	para.put_Alignment(wdAlignParagraphLeft); 
    font.put_Size(14.f);
	font.put_Bold(wdToggle);
}


void CMyTable::insertImage(CString file)
{	
	InlineShapes shapes=oDoc.get_InlineShapes();
	CComVariant varRang(oSel.get_Range());
	InlineShape shape=shapes.AddPicture(file,COleVariant((short)false),COleVariant((short)true),&varRang);
}


void CMyTable::moveHome(void)
{
	oSel.StartOf(COleVariant((short)wdTable),COleVariant((short)wdMove));
//	oSel.MoveRight(COleVariant((short)wdCell),COleVariant((short)1),COleVariant((short)wdMove));
//	oSel.MoveLeft(COleVariant((short)wdCell),COleVariant((short)1),COleVariant((short)wdMove));

}


void CMyTable::end(void)
{
	oSel.EndOf(COleVariant((short)wdTable),COleVariant((short)wdMove));
	oSel.MoveRight(COleVariant((short)wdCharacter),COleVariant((short)2),COleVariant((short)wdMove));
	oSel.TypeParagraph();
}


void CMyTable::shading(long RGB)
{
	Shading shading=oSel.get_Shading();
	shading.put_BackgroundPatternColor(RGB);
}


void CMyTable::writeInFont(CString text,double fontSize, CString fontName,bool isBold)
{
	WordFont font(oSel.get_Font());
	font.put_Name(fontName);
    font.put_Size(fontSize);
	if(isBold)font.put_Bold(wdToggle);
	oSel.TypeText(text);

	font.put_Name(_T("仿宋_GB2312"));
    font.put_Size(14.f);
	if(isBold)font.put_Bold(wdToggle);
}

void CMyTable::setParagraph(int style)
{
	Paragraphs para;
	 //获得对齐方式实例
    para = oSel.get_ParagraphFormat();
	//设置标题行为“居中”
	para.put_Alignment(style);  
	Cells cell=oSel.get_Cells();
	cell.put_VerticalAlignment(style);

}


void CMyTable::write(CString text)
{
	oSel.TypeText(text);
}


void CMyTable::moveDown(int num)
{
	oSel.MoveDown(COleVariant((short)wdLine),COleVariant(short(num)),COleVariant((short)wdMove));
}


void CMyTable::moveRight(int num)
{	
	oSel.MoveRight(COleVariant((short)wdCell),COleVariant(short(num)),COleVariant((short)wdMove));
}


void CMyTable::moveLeft(int num)
{
	oSel.MoveLeft(COleVariant((short)wdCell),COleVariant(short(num)),COleVariant((short)wdMove));
}


void CMyTable::moveUp(int num)
{
	oSel.MoveUp(COleVariant((short)wdLine),COleVariant(short(num)),COleVariant((short)wdMove));
}


void CMyTable::setTableAdapt(CTableAdapt* adapt)
{
	this->adapt=adapt;
}


void CMyTable::saveAs(CString fileName)
{
	WordDoc oActiveDoc = oWordApp.get_ActiveDocument();
    COleVariant FileName(fileName);	//文件名
    COleVariant FileFormat((short)0);
    COleVariant LockComments((short)FALSE);
    COleVariant Password(_T(""));
    COleVariant AddToRecentFiles((short)TRUE);
    COleVariant WritePassword(_T(""));
    COleVariant ReadOnlyRecommended((short)FALSE);
    COleVariant EmbedTrueTypeFonts((short)FALSE);
    COleVariant SaveNativePictureFormat((short)FALSE);
    COleVariant SaveFormsData((short)FALSE);
    COleVariant SaveAsAOCELetter((short)FALSE);
	COleVariant vFalse((short)FALSE);
    oActiveDoc.SaveAs(&FileName,&FileFormat,&LockComments,&Password,
        &AddToRecentFiles,&WritePassword,&ReadOnlyRecommended,
        &EmbedTrueTypeFonts,&SaveNativePictureFormat,&SaveFormsData,
		&SaveAsAOCELetter, vFalse, vFalse, vFalse, vFalse, vFalse);
}
