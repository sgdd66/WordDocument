#pragma once

#define stringNum 9
#define dataNum 25
class TurbineDoc
{
public:
	TurbineDoc(int tupleNum);
	~TurbineDoc(void);
public:

private:
	CString string[stringNum];
	double** data;
	int tupleNum;
public:
	void setString(CString* str);
	void insertData(double* data, int index);
	void makeWord(void);
	void saveAs(CString fileName);
	CString file;
private:
	int length;
};

