
// WordDlg.cpp : 实现文件
//

#include "stdafx.h"
#include "Word.h"
#include "WordDlg.h"
#include "afxdialogex.h"



#ifdef _DEBUG
#define new DEBUG_NEW
#endif


// 用于应用程序“关于”菜单项的 CAboutDlg 对话框

class CAboutDlg : public CDialogEx
{
public:
	CAboutDlg();

// 对话框数据
	enum { IDD = IDD_ABOUTBOX };

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV 支持

// 实现
protected:
	DECLARE_MESSAGE_MAP()
};

CAboutDlg::CAboutDlg() : CDialogEx(CAboutDlg::IDD)
{
}

void CAboutDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
}

BEGIN_MESSAGE_MAP(CAboutDlg, CDialogEx)
END_MESSAGE_MAP()


// CWordDlg 对话框




CWordDlg::CWordDlg(CWnd* pParent /*=NULL*/)
	: CDialogEx(CWordDlg::IDD, pParent)
{
	m_hIcon = AfxGetApp()->LoadIcon(IDR_MAINFRAME);
}

void CWordDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
}

BEGIN_MESSAGE_MAP(CWordDlg, CDialogEx)
	ON_WM_SYSCOMMAND()
	ON_WM_PAINT()
	ON_WM_QUERYDRAGICON()
	ON_BN_CLICKED(IDOK, &CWordDlg::OnBnClickedOk)
	ON_BN_CLICKED(IDC_WordButton, &CWordDlg::OnClickedWordbutton)
END_MESSAGE_MAP()


// CWordDlg 消息处理程序

BOOL CWordDlg::OnInitDialog()
{
	CDialogEx::OnInitDialog();

	// 将“关于...”菜单项添加到系统菜单中。

	// IDM_ABOUTBOX 必须在系统命令范围内。
	ASSERT((IDM_ABOUTBOX & 0xFFF0) == IDM_ABOUTBOX);
	ASSERT(IDM_ABOUTBOX < 0xF000);

	CMenu* pSysMenu = GetSystemMenu(FALSE);
	if (pSysMenu != NULL)
	{
		BOOL bNameValid;
		CString strAboutMenu;
		bNameValid = strAboutMenu.LoadString(IDS_ABOUTBOX);
		ASSERT(bNameValid);
		if (!strAboutMenu.IsEmpty())
		{
			pSysMenu->AppendMenu(MF_SEPARATOR);
			pSysMenu->AppendMenu(MF_STRING, IDM_ABOUTBOX, strAboutMenu);
		}
	}

	// 设置此对话框的图标。当应用程序主窗口不是对话框时，框架将自动
	//  执行此操作
	SetIcon(m_hIcon, TRUE);			// 设置大图标
	SetIcon(m_hIcon, FALSE);		// 设置小图标

	// TODO: 在此添加额外的初始化代码

	return TRUE;  // 除非将焦点设置到控件，否则返回 TRUE
}

void CWordDlg::OnSysCommand(UINT nID, LPARAM lParam)
{
	if ((nID & 0xFFF0) == IDM_ABOUTBOX)
	{
		CAboutDlg dlgAbout;
		dlgAbout.DoModal();
	}
	else
	{
		CDialogEx::OnSysCommand(nID, lParam);
	}
}

// 如果向对话框添加最小化按钮，则需要下面的代码
//  来绘制该图标。对于使用文档/视图模型的 MFC 应用程序，
//  这将由框架自动完成。

void CWordDlg::OnPaint()
{
	if (IsIconic())
	{
		CPaintDC dc(this); // 用于绘制的设备上下文

		SendMessage(WM_ICONERASEBKGND, reinterpret_cast<WPARAM>(dc.GetSafeHdc()), 0);

		// 使图标在工作区矩形中居中
		int cxIcon = GetSystemMetrics(SM_CXICON);
		int cyIcon = GetSystemMetrics(SM_CYICON);
		CRect rect;
		GetClientRect(&rect);
		int x = (rect.Width() - cxIcon + 1) / 2;
		int y = (rect.Height() - cyIcon + 1) / 2;

		// 绘制图标
		dc.DrawIcon(x, y, m_hIcon);
	}
	else
	{
		CDialogEx::OnPaint();
	}
}

//当用户拖动最小化窗口时系统调用此函数取得光标
//显示。
HCURSOR CWordDlg::OnQueryDragIcon()
{
	return static_cast<HCURSOR>(m_hIcon);
}



void CWordDlg::OnBnClickedOk()
{
	// TODO: 在此添加控件通知处理程序代码
	CDialogEx::OnOK();
	//定义表头的属性名
	CString* attrName=new CString[22];
	attrName[0]="STAGE N0.";
	attrName[1]="Gj kg/h";
	attrName[2]="P2 MPa";
	attrName[3]="t2 C";
	attrName[4]="Hs kj/kg";
	attrName[5]="U/Ca";
	attrName[6]="Rec. %";
	attrName[7]="Eff.";
	attrName[8]="Faip";
	attrName[9]="P2p MPa";
	attrName[10]="Noi kw";
	attrName[11]="Nu kw";
	attrName[12]="Alf2 deg";
	attrName[13]="Alf1 deg";
	attrName[14]="Dpp mm";
	attrName[15]="Lp mm";
	attrName[16]="Fp cm**2";
	attrName[17]="Bet2 deg";
	attrName[18]="Dpd mm";
	attrName[19]="Ld mm";
	attrName[20]="Fd cm**2";
	attrName[21]="Fd/Fp";
	
	//CTableAdapt的构造函数，第一参数为属性列数目（int），第二参数为元组数目（int）
	CTableAdapt tableAdapt(22,4);
	tableAdapt.insertString(attrName,0);
	//使用insert逐行插入元组数据，注意数组维度与属性列数目相同
	double* data=new double[22];
	for(int i=0;i<22;i++)
		data[i]=i;
	for(int i=0;i<4;i++)
		tableAdapt.insertData(data,i+1);//从1开始插入
	//更改某行某列的数据
	tableAdapt.setString(_T("rrrrrrrrrrrr"),2,5);

	//以tableAdapt构造一个CMyTable类
	CMyTable table(&tableAdapt);

	//设定标题
	table.setTitle(_T(" ****** PERFORMANCE OF 1D Design ******"));

	//绘制表格
	table.setTable();

	//link用于合并单元格，参数一和二指定从第0行第2列开始，参数三为false是向下合并，为true是向右合并，参数四指定合并数目
	//table.link(1,2,false,2);



	//灰显
	table.moveHome();

	table.shading(RGB(125,125,125));


	//更改某行某列的数据	
	tableAdapt.setString(_T("yyy"),2,6);

	//write用于写入数据
	table.writeTable();
	//以特定字体写入数据
	table.writeInFont(_T("t"),19.f,_T("宋体"),true);
	//insertImage用于插入图片
	TCHAR path[100];
	::GetCurrentDirectoryW(100,path);
	CString p(path);
	p.Append(_T("/3.jpg"));
	table.insertImage(p);

	//在创建新表格前一定要调用end函数
	table.end();

}

void CWordDlg::OnClickedWordbutton()
{
	// TODO: 在此添加控件通知处理程序代码
	
   TurbineDoc doc(2);//8指数据表共有8行数据
   CString str[9];
   str[0]=_T("20120805");//合同号
   str[1]=_T("何伟");//计算人员
   str[2]=_T("W1294-FW");//任务号
   str[3]=_T("李林");//校核
   str[4]=_T("通风机设计");//项目名称
   str[5]=_T("2015.10.26");//日期
   str[6]=_T("TS-E33");//机型
   str[7]=_T("3");//版次
   str[8]=_T("XXX公司");//用户名称
   doc.setString(str);

	double data[28];//28项数据
	for(int i=0;i<28;i++)
	   data[i]=1.2;
	doc.insertData(data,0);
	for(int i=0;i<28;i++)
	   data[i]=3.4;
	doc.insertData(data,1);
	doc.saveAs(_T("E:\solution.doc"));//所保存的文件名
	doc.makeWord();







}
