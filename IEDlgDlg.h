// IEDlgDlg.h : header file
//

#pragma once


// CIEDlgDlg dialog
class CIEDlgDlg : public CDialog
{
// Construction
public:
	CIEDlgDlg(CWnd* pParent = NULL);	// standard constructor
	~CIEDlgDlg();

// Dialog Data
	enum { IDD = IDD_IEDLG_DIALOG };

protected:
	virtual void DoDataExchange(CDataExchange* pDX);	// DDX/DDV support

public:
	DWORD m_dwCookie;
	IWebBrowser2* m_pBrowser2;
	CWnd *WebWindow;
	void DocumentComplete(IDispatch *pDisp, VARIANT *URL);

protected:
	HICON m_hIcon;

	// Generated message map functions
	bool IsBackCompatDocument();
	virtual BOOL OnInitDialog();
	afx_msg void OnPaint();
	afx_msg void OnSize(UINT nType, int cx, int cy);
	afx_msg HCURSOR OnQueryDragIcon();
	DECLARE_MESSAGE_MAP()
	DECLARE_DISPATCH_MAP()	
public:
	afx_msg void OnBnClickedButton1();
public:
	afx_msg void OnBnClickedButton2();
};