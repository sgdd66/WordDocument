#include "stdafx.h"

#ifndef _TABLEADAPT
#define _TABLEADAPT 1
#include"TableAdapt.h"
#endif



CTableAdapt::CTableAdapt(CString** matrix,int hor,int ver)
{
	this->hor=hor;
	this->ver=ver;
	table=new CString*[hor];
	for(int i=0;i<hor;i++)
		table[i]=new CString[ver];
	for(int i=0;i<hor;i++)
		for(int j=0;j<ver;j++)
			table[i][j]=matrix[i][j];

}

CTableAdapt::CTableAdapt(void)
{}


CTableAdapt::~CTableAdapt(void)
{
}


int CTableAdapt::getHor(void)
{
	return hor;
}


int CTableAdapt::getVer(void)
{
	return ver;
}


CTableAdapt::CTableAdapt(const CTableAdapt& adapt)
{
	this->ver=adapt.ver;
	this->hor=adapt.hor;
	table=new CString*[hor];
	for(int i=0;i<hor;i++)
		table[i]=new CString[ver];
	for(int i=0;i<hor;i++)
		for(int j=0;j<ver;j++)
			table[i][j]=adapt.table[i][j];

}


CString CTableAdapt::getString(int hor, int ver)
{
	if(hor>this->hor||ver>this->ver)
		return _T("False");
	return table[hor][ver];
}


void CTableAdapt::setString(CString str, int hor, int ver)
{
	if(hor>this->hor||ver>this->ver)
		return;
	table[hor][ver]=str;
}


CTableAdapt::CTableAdapt(int hor, int ver)
{
	this->hor=hor;
	this->ver=ver;
	table=new CString*[hor];
	for(int i=0;i<hor;i++)
		table[i]=new CString[ver];
}


void CTableAdapt::insertData(double* data, int hor)
{
	for(int i=0;i<this->ver;i++)
		table[hor][i].Format(_T("%f"),data[i]);
}


void CTableAdapt::insertString(CString* attributeName, int hor)
{
	for(int i=0;i<this->ver;i++)
		table[hor][i]=attributeName[i];
}
