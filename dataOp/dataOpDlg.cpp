
// dataOpDlg.cpp : ʵ���ļ�
//

#include "stdafx.h"
#include "dataOp.h"
#include "dataOpDlg.h"
#include "afxdialogex.h"
#include "excel9.h"

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


// CdataOpDlg �Ի���




CdataOpDlg::CdataOpDlg(CWnd* pParent /*=NULL*/)
	: CDialogEx(CdataOpDlg::IDD, pParent)
{
	m_hIcon = AfxGetApp()->LoadIcon(IDR_MAINFRAME);
}

void CdataOpDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
}

BEGIN_MESSAGE_MAP(CdataOpDlg, CDialogEx)
	ON_WM_SYSCOMMAND()
	ON_WM_PAINT()
	ON_WM_QUERYDRAGICON()
	ON_EN_CHANGE(IDC_SELECTFILE, &CdataOpDlg::OnEnChangeSelectfile)
//	ON_LBN_SELCHANGE(IDC_FILELIST, &CdataOpDlg::OnLbnSelchangeFilelist)
	ON_LBN_DBLCLK(IDC_FILELIST, &CdataOpDlg::OnLbnDblclkFilelist)
	ON_BN_CLICKED(IDC_BUTTON1, &CdataOpDlg::OnBnClickedButton1)
	ON_BN_CLICKED(IDC_BUTTON2, &CdataOpDlg::OnBnClickedButton2)
	ON_BN_CLICKED(IDOK, &CdataOpDlg::OnBnClickedOk)
	ON_BN_CLICKED(IDCANCEL, &CdataOpDlg::OnBnClickedCancel)
END_MESSAGE_MAP()


// CdataOpDlg ��Ϣ�������

BOOL CdataOpDlg::OnInitDialog()
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

	if( ! AfxOleInit() )
	{
		AfxMessageBox(_T("����OLEʧ��"));
		return FALSE;
	}
	if(!this->m_ExcelApp.CreateDispatch(_T("Excel.Application")) )
	{
		AfxMessageBox(_T("����Excel������ʧ��!"));
		return ;
	}

	return TRUE;  // ���ǽ��������õ��ؼ������򷵻� TRUE
}

void CdataOpDlg::OnSysCommand(UINT nID, LPARAM lParam)
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

void CdataOpDlg::OnPaint()
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
HCURSOR CdataOpDlg::OnQueryDragIcon()
{
	return static_cast<HCURSOR>(m_hIcon);
}



void CdataOpDlg::OnEnChangeSelectfile()
{
	CFileFind fileFinder;
	CString path;

	this->m_vFileList.clear();
	this->GetDlgItemTextW(IDC_SELECTFILE, this->m_strPath);

	path = this->m_strPath + _T("\\*.xls");

	bool isExist = fileFinder.FindFile(path);
	while( isExist )
	{
		isExist = fileFinder.FindNextFileW();
		this->m_vFileList.push_back(fileFinder.GetFileName());
		refreshListBox();
	}
}

void CdataOpDlg::refreshListBox()
{
	CListBox* pList= (CListBox *)this->GetDlgItem(IDC_FILELIST);
	pList->ResetContent();
	
	int i;
	for(i=0; i<this->m_vFileList.size(); i++ )
	{
		pList->AddString(this->m_vFileList[i]);
	}
}


void CdataOpDlg::OnLbnDblclkFilelist()
{
	CListBox* pList = (CListBox *)this->GetDlgItem(IDC_FILELIST);
	CArray<int, int> listContent;
	int iCnt = pList->GetSelCount();
	listContent.SetSize(iCnt);
	
	pList->GetSelItems(iCnt, listContent.GetData());

	
	if( IDYES == MessageBox(_T("ȷ�ϲ�������Щ�ļ���"),_T("��ʾ"),MB_YESNO))
	{
		for(int i=listContent.GetSize()-1; i>=0; i-- )
		{
			pList->DeleteString(listContent.GetAt(i));
		}
	}
}


void CdataOpDlg::OnBnClickedButton1()
{
	OnLbnDblclkFilelist();
}


void CdataOpDlg::OnBnClickedButton2()
{
	refreshListBox();
}


void CdataOpDlg::OnBnClickedOk()
{
	CListBox* pList = (CListBox *)this->GetDlgItem(IDC_FILELIST);
	int iCnt = pList->GetCount();

	if( 0 == iCnt )
	{
		AfxMessageBox(_T("��ѡ������һ���ļ�!"));
		return ;
	}

	CString str ;
	str.Format(_T("��һ��ѡ���� %d ���ļ����Ƿ�ʼ����"),iCnt);
	if( IDYES == MessageBox(str, _T("��ʾ"), MB_YESNO) )
	{
		this->m_pthMainProcess = AfxBeginThread(CdataOpDlg::MainProcess, (LPVOID)this);
	}
	else
	{

	}
	//CDialogEx::OnOK();
}


void CdataOpDlg::OnBnClickedCancel()
{
	CDialogEx::OnCancel();
}


UINT CdataOpDlg::MainProcess( LPVOID lParam )
{
	CdataOpDlg* pThis = (CdataOpDlg *)lParam;

	CWnd* pStatus = pThis->GetDlgItem(IDC_STATUS);
	CListBox* pFileList =(CListBox* ) pThis->GetDlgItem(IDC_FILELIST);
	CListCtrl* pOutput =(CListCtrl* ) pThis->GetDlgItem(IDC_OUTPUT);

	int iCnt = pFileList->GetCount();
	for(int i=0; i<iCnt; i++ )
	{
		CString str;
		pFileList->GetText(i,str);
		/*AfxMessageBox(str);*/
		pThis->SetDlgItemTextW(IDC_STATUS,_T("���ڴ����ļ� ") + str + _T("...") );
		CdataOpDlg::dealWith(str, pThis);
	}


	return 0;
}

