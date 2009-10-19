#ifndef __HS_EXEARC_READ_H
#define __HS_EXEARC_READ_H

//
// exearc_read.h
// (c)2002 Jonathan Bennett (jon@hiddensoft.com)
//
// Version: v3
//
// STANDALONE CLASS
//

#define HS_EXEARC_FILEVERSION		3			// File wrapper version

// Error codes
#define HS_EXEARC_E_OK				0
#define HS_EXEARC_E_OPENEXE			1
#define HS_EXEARC_E_OPENINPUT		2		
#define HS_EXEARC_E_NOTARC			3
#define HS_EXEARC_E_BADVERSION		4
#define HS_EXEARC_E_BADPWD			5
#define HS_EXEARC_E_FILENOTFOUND	6
#define HS_EXEARC_E_OPENOUTPUT		7		
#define HS_EXEARC_E_MEMALLOC		8

#define HS_EXEARC_MAXPWDLEN			256


class HS_EXEArc_Read
{
public:
	// Functions
	int		Open(const char *szEXEArcFile, const char *szPwd);
	void	Close(void);
	int		FileExtract(const char *szFileID, const char *szFileName);
	int		FileExtractToMem(const char *szFileID, UCHAR **lpData, ULONG *nDataSize);
	int		FileGetFullPath(const char *szFileID, char *szFilename);


private:
	// Variables
	FILE	*m_fEXE;
	ULONG	m_nArchivePtr;
	ULONG	m_nFileSectionPtr;
	char	m_szPwd[HS_EXEARC_MAXPWDLEN+1];
	UINT	m_nPwdHash;

	// Functions
	void	Decrypt(UCHAR *bData, UINT nLen, UINT nSeed);
	int		FileFind(const char *szFileID, char *szFilename);
	void	FileSetTime(const char *szFilename, FILETIME &ftCreated, FILETIME &ftModified);

};


#endif
