// Command.h: interface for the Command class.
//
//////////////////////////////////////////////////////////////////////

#if !defined(AFX_COMMAND_H__461E48FA_2391_4E18_A1D5_2C68571F3B81__INCLUDED_)
#define AFX_COMMAND_H__461E48FA_2391_4E18_A1D5_2C68571F3B81__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

// ��������Ļ���
//����svn
class Command  
{
public:
	Command();
	virtual ~Command();

	virtual void Run() = 0;			// ÿ�������඼Ҫʵ�ֵ�������
};

#endif // !defined(AFX_COMMAND_H__461E48FA_2391_4E18_A1D5_2C68571F3B81__INCLUDED_)
