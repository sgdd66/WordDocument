#include "stdafx.h"
#ifndef TURBINEDOC
#define TURBINEDOC 1
#include"TurbineDoc.h"
#endif


TurbineDoc::TurbineDoc(int tupleNum)
	: length(28)
{
	this->tupleNum=tupleNum;
	data=new double*[tupleNum];

}


TurbineDoc::~TurbineDoc(void)
{
}


void TurbineDoc::setString(CString* str)
{
	for(int i=0;i<stringNum;i++){
		string[i]=str[i];
	}
}


void TurbineDoc::insertData(double* data, int index)
{
	this->data[index]=new double[length];
	for(int i=0;i<length;i++)
		this->data[index][i]=data[i];
}


void TurbineDoc::makeWord(void)
{
	int i,j;
	CTableAdapt adapt(6,12);
	CMyTable table(&adapt);
	table.setTable();
	table.link(5,0,true,2);
	table.link(5,2,true,10);
	for(i=4;i>0;i--)
	{
		table.link(i,0,true,2);
		table.link(i,2,true,3);
		table.link(i,6,true,2);
		table.link(i,8,true,4);
	}
	table.link(0,8,true,4);
	table.link(0,0,true,8);
	table.link(3,1,false,2);
	table.link(0,1,false,3);
	table.moveHome();

	table.writeInFont(_T("汽轮机三维气动计算结果数据表"),19.f,_T("华文楷体"),true);
	table.setParagraph(wdCellAlignVerticalCenter);
	table.moveRight(1);
	CString strexe,strtmp,stricon;
	int nexe;
	::GetModuleFileName(NULL,strexe.GetBufferSetLength(MAX_PATH+1),MAX_PATH);
	nexe = strexe.GetLength();
	for(i=nexe-1;i>=0;i--)
	{
		if(strexe[i]=='\\')
			break;
	}
	strtmp = strexe.Left(i+1);              
	stricon = strtmp+_T("DEC.PNG");
	table.insertImage(stricon);
	table.setParagraph(wdCellAlignVerticalCenter);
	table.moveHome();
	table.shading(RGB(112,173,204));

	table.moveDown(1);
	table.write(_T("合同号"));
	table.moveRight(1);
	table.write(string[0]);
	table.setParagraph(wdCellAlignVerticalCenter);
	table.moveRight(1);
	table.write(_T("计算"));
	table.moveRight(1);
	table.write(string[1]);
	table.setParagraph(wdCellAlignVerticalCenter);

	table.moveDown(1);
	table.moveLeft(3);
	table.write(_T("任务号"));
	table.moveRight(1);
	table.write(string[2]);
	table.setParagraph(wdCellAlignVerticalCenter);
	table.moveRight(1);
	table.write(_T("校核"));
	table.moveRight(1);
	table.write(string[3]);
	table.setParagraph(wdCellAlignVerticalCenter);

	table.moveDown(1);
	table.moveLeft(3);
	table.write(_T("项目名称"));
	table.moveRight(1);
	table.write(string[4]);
	table.setParagraph(wdCellAlignVerticalCenter);
	table.moveRight(1);
	table.write(_T("日期"));
	table.moveRight(1);
	table.write(string[5]);
	table.setParagraph(wdCellAlignVerticalCenter);

	table.moveDown(1);
	table.moveLeft(3);
	table.write(_T("机型"));
	table.moveRight(1);
	table.write(string[6]);
	table.setParagraph(wdCellAlignVerticalCenter);
	table.moveRight(1);
	table.write(_T("版次"));
	table.moveRight(1);
	table.write(string[7]);
	table.setParagraph(wdCellAlignVerticalCenter);

	table.moveRight(1);
	table.write(_T("东方汽轮机有限公司"));
	table.setParagraph(wdCellAlignVerticalCenter);

	table.moveRight(1);
	table.write(_T("用户名称"));
	table.moveRight(1);
	table.write(string[8]);

	table.end();

	int dataNumInTable=7;
	
	CTableAdapt adapt1(tupleNum+3,dataNumInTable+1);
	table.setTableAdapt(&adapt1);
	CString attriName[10];
	attriName[0]=_T("序号");
	attriName[1]=_T("P");
	attriName[2]=_T("T1P");
	attriName[3]=_T("T2");
	attriName[4]=_T("Y0");
	attriName[5]=_T("HOM");
	attriName[6]=_T("HCM");
	attriName[7]=_T("A2");
	adapt1.insertString(attriName,1);
	attriName[0]=_T("");
	attriName[1]=_T("(MPa)");
	attriName[2]=_T("(℃)");
	attriName[3]=_T("(℃)");
	attriName[4]=_T("(%)");
	attriName[5]=_T("(kj/kg)");
	attriName[6]=_T("(kj/kg)");
	attriName[7]=_T("deg");
	adapt1.insertString(attriName,2);

	CString dataName[10];

	int baseNum=0*dataNumInTable;	
	for(i=0;i<tupleNum;i++){

		dataName[0].Format(_T("%d"),i+1);
		dataName[1].Format(_T("%.3f"),data[i][baseNum+0]);
		dataName[2].Format(_T("%.1f"),data[i][baseNum+1]);		
		dataName[3].Format(_T("%.1f"),data[i][baseNum+2]);		
		dataName[4].Format(_T("%.1f"),data[i][baseNum+3]);
		dataName[5].Format(_T("%.4f"),data[i][baseNum+4]);		
		dataName[6].Format(_T("%.4f"),data[i][baseNum+5]);		
		dataName[7].Format(_T("%.2f"),data[i][baseNum+6]);		
		
		adapt1.insertString(dataName,3+i);
	}


	table.setTable();
	table.link(1,0,false,2);
	table.setParagraph(wdCellAlignVerticalCenter);
	table.shading(RGB(125,125,125));
	for(i=0;i<tupleNum;i++){
		table.moveDown(1);
		table.shading(RGB(125,125,125));
	}
	table.writeTable();
	table.link(0,0,true,dataNumInTable+1);
	table.moveHome();
	table.shading(RGB(112,173,204));
	table.setParagraph(wdCellAlignVerticalCenter);
	table.writeInFont(_T("三维计算结果"),16.f,_T("仿宋_GB2312"),false);
	table.end();
////////////////////////////////////////////////////////////////////////////////////
	attriName[0]=_T("序号");
	attriName[1]=_T("RAH");
	attriName[2]=_T("RAM");
	attriName[3]=_T("RAT");
	attriName[4]=_T("YM1");
	attriName[5]=_T("YM2");
	attriName[6]=_T("lp");
	attriName[7]=_T("ld");
	adapt1.insertString(attriName,1);
	attriName[0]=_T("");
	attriName[1]=_T("(%)");
	attriName[2]=_T("(%)");
	attriName[3]=_T("(%)");
	attriName[4]=_T("");
	attriName[5]=_T("");
	attriName[6]=_T("mm");
	attriName[7]=_T("mm");
	adapt1.insertString(attriName,2);
	baseNum=1*dataNumInTable;
	for(i=0;i<tupleNum;i++){
		dataName[0].Format(_T("%d"),i+1);
		dataName[1].Format(_T("%.4f"),data[i][baseNum+0]);
		dataName[2].Format(_T("%.4f"),data[i][baseNum+1]);		
		dataName[3].Format(_T("%.4f"),data[i][baseNum+2]);		
		dataName[4].Format(_T("%.4f"),data[i][baseNum+3]);
		dataName[5].Format(_T("%.4f"),data[i][baseNum+4]);		
		dataName[6].Format(_T("%.2f"),data[i][baseNum+5]);		
		dataName[7].Format(_T("%.2f"),data[i][baseNum+6]);
		adapt1.insertString(dataName,3+i);
	}


	table.setTable();
	table.link(1,0,false,2);
	table.setParagraph(wdCellAlignVerticalCenter);
	table.shading(RGB(125,125,125));
	for(i=0;i<tupleNum;i++){
		table.moveDown(1);
		table.shading(RGB(125,125,125));
	}
	table.writeTable();
	table.link(0,0,true,dataNumInTable+1);
	table.moveHome();
	table.shading(RGB(112,173,204));
	table.setParagraph(wdCellAlignVerticalCenter);
	table.writeInFont(_T("三维计算结果"),16.f,_T("仿宋_GB2312"),false);
	table.end();
////////////////////////////////////////////////////////////////////////////////////////////
	attriName[0]=_T("序号");
	attriName[1]=_T("UC0H");
	attriName[2]=_T("UC0M");
	attriName[3]=_T("UC0T");
	attriName[4]=_T("avalg");
	attriName[5]=_T("aval");
	attriName[6]=_T("avb2g");
	attriName[7]=_T("avb2");
	adapt1.insertString(attriName,1);
	attriName[0]=_T("");
	attriName[1]=_T("");
	attriName[2]=_T("");
	attriName[3]=_T("");
	attriName[4]=_T("deg");
	attriName[5]=_T("deg");
	attriName[6]=_T("deg");
	attriName[7]=_T("deg");
	adapt1.insertString(attriName,2);
	baseNum=2*dataNumInTable;
	for(i=0;i<tupleNum;i++){
		dataName[0].Format(_T("%d"),i+1);
		dataName[1].Format(_T("%.4f"),data[i][baseNum+0]);
		dataName[2].Format(_T("%.4f"),data[i][baseNum+1]);		
		dataName[3].Format(_T("%.4f"),data[i][baseNum+2]);		
		dataName[4].Format(_T("%.2f"),data[i][baseNum+3]);
		dataName[5].Format(_T("%.2f"),data[i][baseNum+4]);		
		dataName[6].Format(_T("%.2f"),data[i][baseNum+5]);		
		dataName[7].Format(_T("%.2f"),data[i][baseNum+6]);
		adapt1.insertString(dataName,3+i);
	}


	table.setTable();
	table.link(1,0,false,2);
	table.setParagraph(wdCellAlignVerticalCenter);
	table.shading(RGB(125,125,125));
	for(i=0;i<tupleNum;i++){
		table.moveDown(1);
		table.shading(RGB(125,125,125));
	}
	table.writeTable();
	table.link(0,0,true,dataNumInTable+1);
	table.moveHome();
	table.shading(RGB(112,173,204));
	table.setParagraph(wdCellAlignVerticalCenter);
	table.writeInFont(_T("三维计算结果"),16.f,_T("仿宋_GB2312"),false);
	table.end();

///////////////////////////////////////////////////////////////////////////////////////////

	CTableAdapt adapt2(tupleNum+3,dataNumInTable+1);
	table.setTableAdapt(&adapt2);

	attriName[0]=_T("序号");
	attriName[1]=_T("VX/Uh");
	attriName[2]=_T("IhS");
	attriName[3]=_T("ImS");
	attriName[4]=_T("ItS");
	attriName[5]=_T("IhR");
	attriName[6]=_T("ImR");
	attriName[7]=_T("ItR");
	adapt2.insertString(attriName,1);
	attriName[0]=_T("");
	attriName[1]=_T("");
	attriName[2]=_T("deg");
	attriName[3]=_T("deg");
	attriName[4]=_T("deg");
	attriName[5]=_T("deg");
	attriName[6]=_T("deg");
	attriName[7]=_T("deg");
	adapt2.insertString(attriName,2);
	baseNum=2*dataNumInTable;	
	for(i=0;i<tupleNum;i++){
		dataName[0].Format(_T("%d"),i+1);
		dataName[1].Format(_T("%.2f"),data[i][baseNum+0]);
		dataName[2].Format(_T("%.2f"),data[i][baseNum+1]);		
		dataName[3].Format(_T("%.2f"),data[i][baseNum+2]);		
		dataName[4].Format(_T("%.2f"),data[i][baseNum+3]);
		dataName[5].Format(_T("%.2f"),data[i][baseNum+4]);		
		dataName[6].Format(_T("%.2f"),data[i][baseNum+5]);		
		dataName[7].Format(_T("%.2f"),data[i][baseNum+6]);		
		adapt2.insertString(dataName,3+i);
	}


	table.setTable();
	table.link(1,0,false,2);
	table.setParagraph(wdCellAlignVerticalCenter);
	table.shading(RGB(125,125,125));
	for(i=0;i<tupleNum;i++){
		table.moveDown(1);
		table.shading(RGB(125,125,125));
	}
	table.writeTable();
	table.link(0,0,true,dataNumInTable+1);
	table.moveHome();
	table.shading(RGB(112,173,204));
	table.setParagraph(wdCellAlignVerticalCenter);
	table.writeInFont(_T("三维计算结果"),16.f,_T("仿宋_GB2312"),false);

	table.saveAs(file);
}


void TurbineDoc::saveAs(CString fileName)
{
	file=fileName;
}
