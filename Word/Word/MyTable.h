
#include"stdafx.h"
#ifndef _TABLEADAPT
#define _TABLEADAPT 1
#include"TableAdapt.h"
#endif

#ifndef _WORD
#define _WORD 1
#include "wordapp.h"
#include "worddocs.h"
#include "worddoc.h"
#include "wordsel.h"
#endif




class CMyTable
{
public:
	CMyTable(CTableAdapt* a);
	~CMyTable(void);
private:
	CTableAdapt* adapt;
	WordApp oWordApp;
	WordDocs oDocs;
	WordDoc oDoc;
	WordSel oSel;
	CString title;
public:
	void setTable();
	void link(int hor, int ver, bool isRight, int num);
	void setTitle(CString title);
	void insertImage(CString file);
	void moveHome(void);
	void end(void);
	void writeTable(void);
	void shading(long RGB);
	void writeInFont(CString text, double fontSize, CString fontName,bool isBold);
	void setParagraph(int style);
	void write(CString text);
	void moveDown(int num);
	void moveRight(int num);
	void moveLeft(int num);
	void moveUp(int num);
	void setTableAdapt(CTableAdapt* adapt);
	void saveAs(CString fileName);
};

