#include "stdafx.h"
#include "dataOpDlg.h"
#include "excel9.h"

void CdataOpDlg::dealWith(const CString &filename, CdataOpDlg* pThis)
{
	
	CWorkbooks books;
	CWorkbook book;
	CWorksheets sheets;
	CWorksheet sheet;
	CRange range;
	LPDISPATCH lpDisp = NULL;

	//����Excel ������(����Excel)
	

	/*�жϵ�ǰExcel�İ汾*/
	CString strExcelVersion = this->m_ExcelApp.get_Version();
	int iStart = 0;
	strExcelVersion = strExcelVersion.Tokenize(_T("."), iStart);
	if (_T("11") == strExcelVersion)
	{
		AfxMessageBox(_T("��ǰExcel�İ汾��2003��"));
	}
	else if (_T("12") == strExcelVersion)
	{
		AfxMessageBox(_T("��ǰExcel�İ汾��2007��"));
	}
	else
	{
		AfxMessageBox(_T("��ǰExcel�İ汾�������汾��"));
	}

	this->m_ExcelApp.put_Visible(TRUE);
	this->m_ExcelApp.put_UserControl(FALSE);

	/*�õ�����������*/
	books.AttachDispatch(this->m_ExcelApp.get_Workbooks());

	/*��һ�����������粻���ڣ�������һ��������*/
	CString strBookPath = _T("C:\\tmp.xls");
	try
	{
		/*��һ��������*/
		lpDisp = books.Open(strBookPath, 
			vtMissing, vtMissing, vtMissing, vtMissing, vtMissing,
			vtMissing, vtMissing, vtMissing, vtMissing, vtMissing, 
			vtMissing, vtMissing, vtMissing, vtMissing);
		book.AttachDispatch(lpDisp);
	}
	catch(...)
	{
		/*����һ���µĹ�����*/
		lpDisp = books.Add(vtMissing);
		book.AttachDispatch(lpDisp);
	}


	/*�õ��������е�Sheet������*/
	sheets.AttachDispatch(book.get_Sheets());

	/*��һ��Sheet���粻���ڣ�������һ��Sheet*/
	CString strSheetName = _T("NewSheet");
	try
	{
		/*��һ�����е�Sheet*/
		lpDisp = sheets.get_Item(_variant_t(strSheetName));
		sheet.AttachDispatch(lpDisp);
	}
	catch(...)
	{
		/*����һ���µ�Sheet*/
		lpDisp = sheets.Add(vtMissing, vtMissing, _variant_t((long)1), vtMissing);
		sheet.AttachDispatch(lpDisp);
		sheet.put_Name(strSheetName);
	}

	system("pause");

	/*��Sheet��д������Ԫ��,��ģΪ10*10 */
	lpDisp = sheet.get_Range(_variant_t("A1"), _variant_t("J10"));
	range.AttachDispatch(lpDisp);

	VARTYPE vt = VT_I4; /*����Ԫ�ص����ͣ�long*/
	SAFEARRAYBOUND sabWrite[2]; /*���ڶ��������ά�����±����ʼֵ*/
	sabWrite[0].cElements = 10;
	sabWrite[0].lLbound = 0;
	sabWrite[1].cElements = 10;
	sabWrite[1].lLbound = 0;

	COleSafeArray olesaWrite;
	olesaWrite.Create(vt, sizeof(sabWrite)/sizeof(SAFEARRAYBOUND), sabWrite);

	/*ͨ��ָ�������ָ�����Զ�ά�����Ԫ�ؽ��м�Ӹ�ֵ*/
	long (*pArray)[2] = NULL;
	olesaWrite.AccessData((void **)&pArray);
	memset(pArray, 0, sabWrite[0].cElements * sabWrite[1].cElements * sizeof(long));

	/*�ͷ�ָ�������ָ��*/
	olesaWrite.UnaccessData();
	pArray = NULL;

	/*�Զ�ά�����Ԫ�ؽ��������ֵ*/
	long index[2] = {0, 0};
	long lFirstLBound = 0;
	long lFirstUBound = 0;
	long lSecondLBound = 0;
	long lSecondUBound = 0;
	olesaWrite.GetLBound(1, &lFirstLBound);
	olesaWrite.GetUBound(1, &lFirstUBound);
	olesaWrite.GetLBound(2, &lSecondLBound);
	olesaWrite.GetUBound(2, &lSecondUBound);
	for (long i = lFirstLBound; i <= lFirstUBound; i++)
	{
		index[0] = i;
		for (long j = lSecondLBound; j <= lSecondUBound; j++)
		{
			index[1] = j;
			long lElement = i * sabWrite[1].cElements + j; 
			olesaWrite.PutElement(index, &lElement);
		}
	}

	/*��ColesaWritefeArray����ת��ΪVARIANT,��д�뵽Excel�����*/
	VARIANT varWrite = (VARIANT)olesaWrite;
	range.put_Value2(varWrite);

	system("pause");



	/*��ȡExcel���еĶ����Ԫ���ֵ����listctrl����ʾ*/
	VARIANT varRead = range.get_Value2();
	COleSafeArray olesaRead(varRead);

	VARIANT varItem;
	CString strItem;
	lFirstLBound = 0;
	lFirstUBound = 0;
	lSecondLBound = 0;
	lSecondUBound = 0;
	olesaRead.GetLBound(1, &lFirstLBound);
	olesaRead.GetUBound(1, &lFirstUBound);
	olesaRead.GetLBound(2, &lSecondLBound);
	olesaRead.GetUBound(2, &lSecondUBound);
	memset(index, 0, 2 * sizeof(long));


	/*�ͷ���Դ*/
	sheet.ReleaseDispatch();
	sheets.ReleaseDispatch();
	book.ReleaseDispatch();
	books.ReleaseDispatch();
	this->m_ExcelApp.Quit();
	this->m_ExcelApp.ReleaseDispatch();
}