
class CTableAdapt
{
public:
	CTableAdapt(CString** matrix,int hor,int ver);
	~CTableAdapt(void);
	CTableAdapt(const CTableAdapt& adapt);
	CTableAdapt(int hor, int ver);
	CTableAdapt(void);
private:
	CString** table;
	int hor;
	int ver;
public:
	int getHor(void);
	int getVer(void);
	CString getString(int hor, int ver);
	void setString(CString str, int hor, int ver);
	void insertData(double* data, int hor);
	void insertString(CString* attributeName, int hor);
};

