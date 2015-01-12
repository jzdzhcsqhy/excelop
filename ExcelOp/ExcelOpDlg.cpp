
// ExcelOpDlg.cpp : 实现文件
//

#include "stdafx.h"
#include "ExcelOp.h"
#include "ExcelOpDlg.h"
#include "afxdialogex.h"
#include <fstream>
#include <string>


using namespace std;

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


// CExcelOpDlg 对话框




CExcelOpDlg::CExcelOpDlg(CWnd* pParent /*=NULL*/)
	: CDialogEx(CExcelOpDlg::IDD, pParent)
{
	m_hIcon = AfxGetApp()->LoadIcon(IDR_MAINFRAME);
}

void CExcelOpDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
}

BEGIN_MESSAGE_MAP(CExcelOpDlg, CDialogEx)
	ON_WM_SYSCOMMAND()
	ON_WM_PAINT()
	ON_WM_QUERYDRAGICON()
	ON_BN_CLICKED(IDOK, &CExcelOpDlg::OnBnClickedOk)
END_MESSAGE_MAP()


// CExcelOpDlg 消息处理程序

BOOL CExcelOpDlg::OnInitDialog()
{
	CDialogEx::OnInitDialog();
	if( !AfxOleInit() )
	{
		AfxMessageBox(_T("aaaa"));
		return FALSE ;
	}

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

void CExcelOpDlg::OnSysCommand(UINT nID, LPARAM lParam)
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

void CExcelOpDlg::OnPaint()
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
HCURSOR CExcelOpDlg::OnQueryDragIcon()
{
	return static_cast<HCURSOR>(m_hIcon);
}



void CExcelOpDlg::OnBnClickedOk()
{
	// TODO: 在此添加控件通知处理程序代码
	_Application app;
	Workbooks books;
	_Workbook book;
	Worksheets sheets;
	_Worksheet sheet;
	Range range;
	//Font font;
	Range cols;
	COleVariant covOPtional((long) DISP_E_PARAMNOTFOUND,VT_ERROR);

	ifstream fin("1.csv");
	string str;
	str.clear();
	if(!app.CreateDispatch(_T("Excel.Application")))
	{
		AfxMessageBox(_T("无法创建Excel应用！"));  
		return;
	}
	books=app.GetWorkbooks();
	book=books.Add(covOPtional);
	sheets=book.GetSheets();
	sheet=sheets.GetItem(COleVariant((short)1));
	while( fin >> str )
	{
		CString str1;
		str1.Format(_T("%s"), str.c_str());
		//AfxMessageBox(str1);
		
		
		
		
		/*range=sheet.GetRange(COleVariant(_T("A1")),COleVariant(_T("A1")));
		range.SetValue(COleVariant(_T("HELLO EXCEL!")));
		range=sheet.GetRange(COleVariant(_T("A2")),COleVariant(_T("A2")));
		range.SetFormula(COleVariant(_T("0000")));
		range.SetNumberFormat(COleVariant(_T("$0.00")));
		*/
		range = sheet.GetRange(COleVariant(_T("A2")),COleVariant(_T("A5")));
		range.Merge(COleVariant(_T("FALSE")));

		cols=range.GetEntireColumn();
		cols.AutoFit();
		str.clear();
		break;
	}
	system("del C:\\Users\\Free\\Desktop\\rs.xls");
	app.SetVisible(TRUE);
	
	book.SaveAs(COleVariant(_T("C:\\Users\\Free\\Desktop\\rs.xls")),covOPtional ,
		covOPtional,covOPtional,
		covOPtional,covOPtional,(long)0,covOPtional,covOPtional,covOPtional,
		covOPtional);
	CDialogEx::OnOK();
}
