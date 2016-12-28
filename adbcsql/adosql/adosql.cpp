// adosql.cpp : �������̨Ӧ�ó������ڵ㡣
//

#include "stdafx.h"
#include <windows.h>
#include <string>
#include <iostream>
#include <sqltypes.h>
#include <sql.h>
#include <sqlext.h>
#include <atlbase.h>

#define DBSERVER "(local)"
#define DBNAME "test"
#define DSNNAME "hello"
#define USRID	"sa"

//˵����ҪҪ�޸ĵĺ���Ϊ�� BOOL InitParams() �� BOOL AddDtaAndExecu()


//���ݿ�ִ�и�ʽ ��BATCH�� ���ݿ������F1��F2...���ݿ�������
TCHAR szSQL[512] = TEXT("INSERT INTO BATCH(F1,F2,F3,F4,F5,F6,F7,F8,F9) VALUES(?,?,?,?,?,?,?,?,?)");
//���ݿ��ṹ��Ҫ�����ݿ�ִ�и�ʽһһ��Ӧ������F[8]��ʶF1~F8(Ϊint����),F9��ʶF9(��ṹ��Ϊchar*����)
//cb[9]��ʶ���ϸ���������ռ�ֽ���
typedef struct tagRECORD
{
	int F[8];
	char F9[1024];
	LONG cb[9];
} RECORD, *LPRECORD;

using namespace std;

//#import "c:\program files\common files\system\ado\msado15.dll" no_namespace rename("EOF", "adoEOF") 

#define SQLOK(rc) ((SQL_SUCCESS==rc) || (SQL_SUCCESS_WITH_INFO==rc))

const int ROW_ARRAY_SIZE = 100000;

BOOL MakeSQLServerODBCDSN(char* strDBServer, char* strDBName, char* strDSN, char* strUID)
{
	long lRtn = 0;
	//ע������Դ��
	HKEY hKey;
	lRtn = RegOpenKeyExA(HKEY_CURRENT_USER, "SOFTWARE\\ODBC\\ODBC.INI", 0, 0, &hKey);
	//if (lRtn != ERROR_SUCCESS) return FALSE;

	char strSubKey[1024] = { 0 };
	sprintf(strSubKey, ("SOFTWARE\\ODBC\\ODBC.INI\\%s"), strDSN);
	lRtn = RegCreateKeyA(HKEY_CURRENT_USER, strSubKey, &hKey);
	if (lRtn != ERROR_SUCCESS) return FALSE;

	char buffer[128] = { 0 };
	int Length = GetSystemDirectoryA(buffer, 128);
	//ע��ODBC��������
	char driver[MAX_PATH] = { 0 };
	sprintf(driver, ("%s\\sqlsrv32.dll"), buffer);
	lRtn = RegSetValueExA(hKey, "Database", 0, REG_SZ, (const unsigned char *)((LPCTSTR)strDBName), strlen(strDBName) + 1);
	if (lRtn != ERROR_SUCCESS) return FALSE;
	lRtn = RegSetValueExA(hKey, "Driver", 0, REG_SZ, (const unsigned char *)((LPCTSTR)driver), strlen(driver) + 1);
	if (lRtn != ERROR_SUCCESS) return FALSE;
	lRtn = RegSetValueExA(hKey, "LastUser", 0, REG_SZ, (const unsigned char *)(strUID), strlen(strUID) + 1);
	if (lRtn != ERROR_SUCCESS) return FALSE;
	lRtn = RegSetValueExA(hKey, "Server", 0, REG_SZ, (const unsigned char *)(strDBServer), strlen(strDBServer) + 1);
	if (lRtn != ERROR_SUCCESS) return FALSE;
	lRtn = RegSetValueExA(hKey, "Trusted_Connection", 0, REG_SZ, (const unsigned char *)((LPCTSTR)"Yes"), 3);
	if (lRtn != ERROR_SUCCESS) return FALSE;
	//����Server ODBC����Դ
	lRtn = RegOpenKeyExA(HKEY_CURRENT_USER, "Software\\ODBC\\ODBC.INI\\ODBC Data Sources", 0, 0, &hKey);
	lRtn = RegCreateKeyA(HKEY_CURRENT_USER, "Software\\ODBC\\ODBC.INI\\ODBC Data Sources", &hKey);
	lRtn = RegSetValueExA(hKey, strDSN, 0, REG_SZ, (const unsigned char *)((LPCTSTR)"SQL Server"), 10);
	return TRUE;
}

SQLHENV g_henv;
SQLHDBC g_hdbc;
SQLHSTMT g_hstmt;
LPRECORD pRecords = NULL;

short iRowStatusPtr[ROW_ARRAY_SIZE] = { 0 };
UINT uRowsFetched = 0;
UINT nTotalCount = 0;

BOOL FreeSQL()
{
	SQLEndTran(SQL_HANDLE_DBC, g_hdbc, SQL_COMMIT);
	if (NULL != pRecords) LocalFree(pRecords);
	if (NULL != g_hstmt) SQLFreeHandle(SQL_HANDLE_STMT, g_hstmt);
	return TRUE;
}

BOOL InitSQL()
{
	SQLRETURN retcode;
	retcode = SQLAllocHandle(SQL_HANDLE_ENV, SQL_NULL_HANDLE, &g_henv);
	retcode = SQLSetEnvAttr(g_henv, SQL_ATTR_ODBC_VERSION, (SQLPOINTER*)SQL_OV_ODBC3, 0);
	retcode = SQLAllocHandle(SQL_HANDLE_DBC, g_henv, &g_hdbc);
	retcode = SQLSetConnectAttr(g_hdbc, SQL_LOGIN_TIMEOUT, (SQLPOINTER)10, 0);
	retcode = SQLConnectA(g_hdbc, (SQLCHAR*)DSNNAME, SQL_NTS, (SQLCHAR*)USRID, SQL_NTS, (SQLCHAR*)NULL, SQL_NTS);
	retcode = SQLAllocHandle(SQL_HANDLE_STMT, g_hdbc, &g_hstmt);
	SQLRETURN rc = SQLPrepare(g_hstmt, (SQLTCHAR *)szSQL, SQL_NTS);
	if (!SQLOK(retcode) || !SQLOK(retcode))
	{
		FreeSQL();
		return FALSE;
	}

	return TRUE;
}

BOOL AllocMem()
{
	SQLRETURN rc;
	rc = SQLSetStmtAttr(g_hstmt, SQL_ATTR_PARAM_BIND_TYPE, (SQLPOINTER)sizeof(RECORD), SQL_IS_INTEGER);
	rc = SQLSetStmtAttr(g_hstmt, SQL_ATTR_PARAMSET_SIZE, (SQLPOINTER)ROW_ARRAY_SIZE, SQL_IS_INTEGER);
	rc = SQLSetStmtAttr(g_hstmt, SQL_ATTR_PARAM_STATUS_PTR, (SQLPOINTER)iRowStatusPtr, SQL_IS_INTEGER);
	rc = SQLSetStmtAttr(g_hstmt, SQL_ATTR_PARAMS_PROCESSED_PTR, (SQLPOINTER)&uRowsFetched, SQL_IS_INTEGER);

	pRecords = (LPRECORD)LocalAlloc(LPTR, sizeof(RECORD)*ROW_ARRAY_SIZE);
	if (NULL == pRecords) return FALSE;

	rc = SQLSetConnectAttr(g_hdbc, SQL_ATTR_AUTOCOMMIT, (SQLPOINTER)SQL_AUTOCOMMIT_OFF, 0);
	return SQLOK(rc) ? TRUE : FALSE;
}
//��������Ҫ�����ݿ�����һ�£���Ӧ��ϵ����
//#define SQL_C_CHAR    SQL_CHAR             /* CHAR, VARCHAR, DECIMAL, NUMERIC */
//#define SQL_C_LONG    SQL_INTEGER          /* INTEGER                      */
//#define SQL_C_SHORT   SQL_SMALLINT         /* SMALLINT                     */
//#define SQL_C_FLOAT   SQL_REAL             /* REAL                         */
//#define SQL_C_DOUBLE  SQL_DOUBLE           /* FLOAT, DOUBLE                */
//��Ҫ��עSQLBindParameter�ĵ�4��5������
BOOL InitParams()
{
	SQLRETURN rc;
	for (int i = 0; i < 8; i++)
	{
		rc = SQLBindParameter(g_hstmt, i + 1, SQL_PARAM_INPUT, SQL_C_LONG, SQL_INTEGER, 10, 0, &pRecords[0].F[i], 0, &pRecords[0].cb[i]);
	}

	rc = SQLBindParameter(g_hstmt, 9, SQL_PARAM_INPUT, SQL_C_CHAR, SQL_CHAR, 10, 0, &pRecords[0].F[8], 0, &pRecords[0].cb[8]);

	return TRUE;
}

BOOL AddDtaAndExecu()
{
	SQLRETURN rc;
	int j = 1;
	{
		for (int i = 0; i < ROW_ARRAY_SIZE; i++)
		{
			RECORD &data = pRecords[i];
			//��������
			data.F[0] = j*ROW_ARRAY_SIZE + i * 1;
			data.F[1] = j*ROW_ARRAY_SIZE + i * 2;
			data.F[2] = j*ROW_ARRAY_SIZE + i * 3;
			data.F[3] = j*ROW_ARRAY_SIZE + i * 4;
			data.F[4] = j*ROW_ARRAY_SIZE + i * 5;
			data.F[5] = j*ROW_ARRAY_SIZE + i * 6;
			data.F[6] = j*ROW_ARRAY_SIZE + i * 7;
			data.F[7] = j*ROW_ARRAY_SIZE + i * 8;
			strcpy(data.F9, "hello");

			//����������ռ�ֽ�
			data.cb[0] = 4;
			data.cb[1] = 4;
			data.cb[2] = 4;
			data.cb[3] = 4;
			data.cb[4] = 4;
			data.cb[5] = 4;
			data.cb[6] = 4;
			data.cb[7] = 4;
			data.cb[8] = 6;
		}
		rc = SQLExecute(g_hstmt);
		if (!SQLOK(rc))
		{
			FreeSQL();
			return FALSE;
		}
		nTotalCount += uRowsFetched;
		uRowsFetched = 0;
	}
	return TRUE;
}

#define ERRORRETURN( INFO ) if (!bRe){\
		MessageBox(NULL, INFO, NULL, MB_OK);\
		return 0;}

int _tmain(int argc, _TCHAR* argv[])
{
	//��������Դ
	BOOL bRe = MakeSQLServerODBCDSN(DBSERVER, DBNAME, DSNNAME, USRID); ERRORRETURN(TEXT("����Դ����ʧ��!"));
	bRe = InitSQL(); ERRORRETURN(TEXT("���ݿ��ʼ��ʧ��!"));

	bRe = AllocMem(); ERRORRETURN(TEXT("�ڴ�����ʧ��!"));

	DWORD dwTime = GetTickCount();
	bRe = InitParams(); ERRORRETURN(TEXT("��������ʧ��!"));
	bRe = AddDtaAndExecu(); ERRORRETURN(TEXT("ִ�й���ʧ��!"));

	dwTime = GetTickCount() - dwTime;
	TCHAR szText[MAX_PATH] = TEXT("");
	wsprintf(szText, TEXT("%ld rows inserted.\r\nTime Ellipsed=%ldms."), nTotalCount, dwTime);
	MessageBox(NULL, szText, 0, 0);
	FreeSQL();

	return 0;
}

