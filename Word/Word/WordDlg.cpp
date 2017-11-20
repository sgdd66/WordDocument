
// WordDlg.cpp : ʵ���ļ�
//

#include "stdafx.h"
#include "Word.h"
#include "WordDlg.h"
#include "afxdialogex.h"



#ifdef _DEBUG
#define new DEBUG_NEW
#endif


// ����Ӧ�ó��򡰹��ڡ��˵���� CAboutDlg �Ի���

class CAboutDlg : public CDialogEx
{
public:
	CAboutDlg();

// �Ի�������
	enum { IDD = IDD_ABOUTBOX };

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV ֧��

// ʵ��
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


// CWordDlg �Ի���




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


// CWordDlg ��Ϣ�������

BOOL CWordDlg::OnInitDialog()
{
	CDialogEx::OnInitDialog();

	// ��������...���˵�����ӵ�ϵͳ�˵��С�

	// IDM_ABOUTBOX ������ϵͳ���Χ�ڡ�
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

	// ���ô˶Ի����ͼ�ꡣ��Ӧ�ó��������ڲ��ǶԻ���ʱ����ܽ��Զ�
	//  ִ�д˲���
	SetIcon(m_hIcon, TRUE);			// ���ô�ͼ��
	SetIcon(m_hIcon, FALSE);		// ����Сͼ��

	// TODO: �ڴ���Ӷ���ĳ�ʼ������

	return TRUE;  // ���ǽ��������õ��ؼ������򷵻� TRUE
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

// �����Ի��������С����ť������Ҫ����Ĵ���
//  �����Ƹ�ͼ�ꡣ����ʹ���ĵ�/��ͼģ�͵� MFC Ӧ�ó���
//  �⽫�ɿ���Զ���ɡ�

void CWordDlg::OnPaint()
{
	if (IsIconic())
	{
		CPaintDC dc(this); // ���ڻ��Ƶ��豸������

		SendMessage(WM_ICONERASEBKGND, reinterpret_cast<WPARAM>(dc.GetSafeHdc()), 0);

		// ʹͼ���ڹ����������о���
		int cxIcon = GetSystemMetrics(SM_CXICON);
		int cyIcon = GetSystemMetrics(SM_CYICON);
		CRect rect;
		GetClientRect(&rect);
		int x = (rect.Width() - cxIcon + 1) / 2;
		int y = (rect.Height() - cyIcon + 1) / 2;

		// ����ͼ��
		dc.DrawIcon(x, y, m_hIcon);
	}
	else
	{
		CDialogEx::OnPaint();
	}
}

//���û��϶���С������ʱϵͳ���ô˺���ȡ�ù��
//��ʾ��
HCURSOR CWordDlg::OnQueryDragIcon()
{
	return static_cast<HCURSOR>(m_hIcon);
}



void CWordDlg::OnBnClickedOk()
{
	// TODO: �ڴ���ӿؼ�֪ͨ����������
	CDialogEx::OnOK();
	//�����ͷ��������
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
	
	//CTableAdapt�Ĺ��캯������һ����Ϊ��������Ŀ��int�����ڶ�����ΪԪ����Ŀ��int��
	CTableAdapt tableAdapt(22,4);
	tableAdapt.insertString(attrName,0);
	//ʹ��insert���в���Ԫ�����ݣ�ע������ά������������Ŀ��ͬ
	double* data=new double[22];
	for(int i=0;i<22;i++)
		data[i]=i;
	for(int i=0;i<4;i++)
		tableAdapt.insertData(data,i+1);//��1��ʼ����
	//����ĳ��ĳ�е�����
	tableAdapt.setString(_T("rrrrrrrrrrrr"),2,5);

	//��tableAdapt����һ��CMyTable��
	CMyTable table(&tableAdapt);

	//�趨����
	table.setTitle(_T(" ****** PERFORMANCE OF 1D Design ******"));

	//���Ʊ��
	table.setTable();

	//link���ںϲ���Ԫ�񣬲���һ�Ͷ�ָ���ӵ�0�е�2�п�ʼ��������Ϊfalse�����ºϲ���Ϊtrue�����Һϲ���������ָ���ϲ���Ŀ
	//table.link(1,2,false,2);



	//����
	table.moveHome();

	table.shading(RGB(125,125,125));


	//����ĳ��ĳ�е�����	
	tableAdapt.setString(_T("yyy"),2,6);

	//write����д������
	table.writeTable();
	//���ض�����д������
	table.writeInFont(_T("t"),19.f,_T("����"),true);
	//insertImage���ڲ���ͼƬ
	TCHAR path[100];
	::GetCurrentDirectoryW(100,path);
	CString p(path);
	p.Append(_T("/3.jpg"));
	table.insertImage(p);

	//�ڴ����±��ǰһ��Ҫ����end����
	table.end();

}

void CWordDlg::OnClickedWordbutton()
{
	// TODO: �ڴ���ӿؼ�֪ͨ����������
	
   TurbineDoc doc(2);//8ָ���ݱ���8������
   CString str[9];
   str[0]=_T("20120805");//��ͬ��
   str[1]=_T("��ΰ");//������Ա
   str[2]=_T("W1294-FW");//�����
   str[3]=_T("����");//У��
   str[4]=_T("ͨ������");//��Ŀ����
   str[5]=_T("2015.10.26");//����
   str[6]=_T("TS-E33");//����
   str[7]=_T("3");//���
   str[8]=_T("XXX��˾");//�û�����
   doc.setString(str);

	double data[28];//28������
	for(int i=0;i<28;i++)
	   data[i]=1.2;
	doc.insertData(data,0);
	for(int i=0;i<28;i++)
	   data[i]=3.4;
	doc.insertData(data,1);
	doc.saveAs(_T("E:\solution.doc"));//��������ļ���
	doc.makeWord();







}
