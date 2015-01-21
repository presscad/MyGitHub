#ifndef READTABLETOEXCEL_H__
#define READTABLETOEXCEL_H__

#include "command.h"

struct tDBTextInfor
{
	CString strDBText;
	double dX;
	double dY;
	tDBTextInfor()
	{
		strDBText = _T("");
		dX = 0.0;
		dY = 0.0;
	}
};//处理表格
bool sortDVTextByDx(tDBTextInfor text1, tDBTextInfor text2);


class CReadTableToExcel:public Command
{
	
public:
	CReadTableToExcel();

	void Run();

	bool readTable(ACHAR * strPrompt, vector<vector<CString> > &vec2Table);
	void printToFile(CString strPath, const vector<vector<CString> > vec2Table);

	void sortByTxtCoordinate(vector<tDBTextInfor> vecTmpTable, double dHight);

private:
	bool bReverse;
	std::set<double> m_setRow;
	std::set<double> m_setCol;


};

#endif