// IEDlgDlg.cpp : implementation file
//

#include "stdafx.h"
#include "IEDlg.h"
#include "IEDlgDlg.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#endif


// CIEDlgDlg dialog
CIEDlgDlg::CIEDlgDlg(CWnd* pParent /*=NULL*/)
	: CDialog(CIEDlgDlg::IDD, pParent)
{
	m_hIcon = AfxGetApp()->LoadIcon(IDR_MAINFRAME);
	EnableAutomation();
	m_pBrowser2 = NULL;
}

CIEDlgDlg::~CIEDlgDlg()
{
	if(m_pBrowser2)
	{
		LPUNKNOWN pUnkSink = GetIDispatch(FALSE);
		AfxConnectionUnadvise((LPUNKNOWN)m_pBrowser2, DIID_DWebBrowserEvents2, pUnkSink, FALSE, m_dwCookie);
		m_pBrowser2->Quit();
		m_pBrowser2->Release();
	}
	OleUninitialize();
	if(WebWindow) delete [] WebWindow;
}

void CIEDlgDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
}

BEGIN_MESSAGE_MAP(CIEDlgDlg, CDialog)
	ON_WM_PAINT()
	ON_WM_QUERYDRAGICON()
	ON_WM_SIZE()
	//}}AFX_MSG_MAP
	ON_BN_CLICKED(IDC_BUTTON1, &CIEDlgDlg::OnBnClickedButton1)
	ON_BN_CLICKED(IDC_BUTTON2, &CIEDlgDlg::OnBnClickedButton2)
END_MESSAGE_MAP()
BEGIN_DISPATCH_MAP(CIEDlgDlg, CCmdTarget)
   DISP_FUNCTION_ID(CIEDlgDlg, "DocumentComplete", DISPID_DOCUMENTCOMPLETE, DocumentComplete,
					 VT_EMPTY, VTS_DISPATCH VTS_PVARIANT)
END_DISPATCH_MAP()


// CIEDlgDlg message handlers

BOOL CIEDlgDlg::OnInitDialog()
{
	CDialog::OnInitDialog();

	// Set the icon for this dialog.  The framework does this automatically
	//  when the application's main window is not a dialog
	SetIcon(m_hIcon, TRUE);			// Set big icon
	SetIcon(m_hIcon, FALSE);		// Set small icon

	// TODO: Add extra initialization here
	//new
	CWnd *DlgWindow = this->GetWindow(0);
	DlgWindow->SetWindowPos(NULL,200,200,1800,800,NULL);
	WebWindow = new CWnd;
	CRect DlgRect;
	GetClientRect(&DlgRect);
	//DlgRect.bottom -= 40;
	DlgRect.top += 40;
	WebWindow->CreateControl(CLSID_WebBrowser, NULL, WS_VISIBLE, DlgRect, this, 1005);
	WebWindow->ShowWindow(1);
	OleInitialize(NULL);
	IUnknown *pUnk = WebWindow->GetControlUnknown(); 
	HRESULT res = pUnk->QueryInterface(IID_IWebBrowser2, (void **)&m_pBrowser2);
	if(!SUCCEEDED(res)) m_pBrowser2 = NULL;

	return TRUE;  // return TRUE  unless you set the focus to a control
}

// If you add a minimize button to your dialog, you will need the code below
//  to draw the icon.  For MFC applications using the document/view model,
//  this is automatically done for you by the framework.

void CIEDlgDlg::OnPaint()
{
	if (IsIconic())
	{
		CPaintDC dc(this); // device context for painting

		SendMessage(WM_ICONERASEBKGND, reinterpret_cast<WPARAM>(dc.GetSafeHdc()), 0);

		// Center icon in client rectangle
		int cxIcon = GetSystemMetrics(SM_CXICON);
		int cyIcon = GetSystemMetrics(SM_CYICON);
		CRect rect;
		GetClientRect(&rect);
		int x = (rect.Width() - cxIcon + 1) / 2;
		int y = (rect.Height() - cyIcon + 1) / 2;

		// Draw the icon
		dc.DrawIcon(x, y, m_hIcon);
	}
	else
	{
		CDialog::OnPaint();
	}
}

void CIEDlgDlg::OnSize(UINT nType, int cx, int cy)
{
	SetWindowPos(NULL, 50, 50, 1200, 900, NULL);
}

// The system calls this function to obtain the cursor to display while the user drags
//  the minimized window.
HCURSOR CIEDlgDlg::OnQueryDragIcon()
{
	return static_cast<HCURSOR>(m_hIcon);
}

void CIEDlgDlg::DocumentComplete(IDispatch *pDisp, VARIANT *URL)
{
	HRESULT hr;
    IUnknown* pUnkBrowser = NULL;
    IUnknown* pUnkDisp = NULL;

    // Is this the DocumentComplete event for the top frame window?
    // Check COM identity: compare IUnknown interface pointers.
    hr = m_pBrowser2->QueryInterface(IID_IUnknown, (void**)&pUnkBrowser);
    if (SUCCEEDED(hr))
    {
        hr = pDisp->QueryInterface(IID_IUnknown, (void**)&pUnkDisp);
        if (SUCCEEDED(hr))
        {
            if (pUnkBrowser == pUnkDisp)
			{
				IHTMLDocument2* pDoc = NULL;	
				IViewObject	*pViewObject = NULL;
				hr = m_pBrowser2->get_Document(&pDisp);
				if ((hr == S_OK) && (pDisp != NULL))
				{				 
					pDisp->QueryInterface(IID_IHTMLDocument2, (LPVOID *)&pDoc);	
					pDisp->Release();
					pDoc->QueryInterface(IID_IViewObject, (LPVOID *)&pViewObject);
					pDoc->Release();
					if(pViewObject)
					{
						long scrollWidth(0);
						long scrollHeight(0);
						//BackCompat
						if(IsBackCompatDocument())
						{
							CComPtr<IHTMLElement> body;
							HRESULT hr = pDoc->get_body(&body);
							if(FAILED(hr) || body == NULL) return;
							CComQIPtr<IHTMLElement2> body2 = body;
							if(body2 == NULL) return;
							body2->get_scrollWidth(&scrollWidth);
							body2->get_scrollHeight(&scrollHeight);
						}
						//CSS1Compat
						else
						{
							CComQIPtr<IHTMLDocument3> pDoc3 = pDoc;
							if (pDoc3 == NULL) return;
							CComPtr<IHTMLElement> rootElement = NULL;
							HRESULT hr = pDoc3->get_documentElement(&rootElement);
							if (FAILED(hr) || rootElement == NULL) return;
							CComQIPtr<IHTMLElement2> rootElement2 = rootElement;
							if (rootElement2 == NULL) return;
							rootElement2->get_scrollWidth(&scrollWidth);
							rootElement2->get_scrollHeight(&scrollHeight);
						}
						if(!scrollWidth || !scrollHeight)
						{
							MessageBox("Nulls in width or height!");
							return;
						}

						//prepare bitmap
						using namespace Gdiplus;
						GdiplusStartupInput gdiplusStartupInput;
						ULONG_PTR gdiplusToken;
						GdiplusStartup(&gdiplusToken, &gdiplusStartupInput, NULL);
						Bitmap bmp(scrollWidth, scrollHeight);
						Graphics g(&bmp);
						HDC bmpHdc = g.GetHDC();

						//draw
						HDC hDC = ::GetDC(WebWindow->m_hWnd);
						RECTL rcSource = {0, 0, scrollWidth, scrollHeight};
						hr = pViewObject->Draw(DVASPECT_CONTENT, -1, NULL, NULL, hDC, bmpHdc, &rcSource, NULL, NULL, 0);
						::ReleaseDC(WebWindow->m_hWnd, hDC);
						pViewObject->Release();

						//save bitmap
						g.ReleaseHDC(bmpHdc);
						CLSID jpgClsid;
						GetEncoderClsid(L"image/jpeg", &jpgClsid);
						bmp.Save(L"test.jpeg", &jpgClsid, NULL);
					}
				}
			}
		}
		pUnkDisp->Release();
	}
	pUnkBrowser->Release();
}
void CIEDlgDlg::OnBnClickedButton1()
{
	if(!m_pBrowser2) return;
	
	LPUNKNOWN pUnkSink = GetIDispatch(FALSE);
	AfxConnectionAdvise((LPUNKNOWN)m_pBrowser2, DIID_DWebBrowserEvents2, pUnkSink, FALSE, &m_dwCookie); 

	VARIANT vEmpty;
	VariantInit(&vEmpty);
	BSTR bstrURL = SysAllocString(L"http://www.ru");
	HRESULT hr = m_pBrowser2->Navigate(bstrURL, &vEmpty, &vEmpty, &vEmpty, &vEmpty);
	if(!SUCCEEDED(hr)) return;
	m_pBrowser2->put_Visible(VARIANT_TRUE);
}

void CIEDlgDlg::OnBnClickedButton2()
{
	if(!m_pBrowser2) return;
	
	VARIANT vEmpty;
	VariantInit(&vEmpty);
	BSTR bstrURL = SysAllocString(L"http://www.ru");
	HRESULT hr = m_pBrowser2->Navigate(bstrURL, &vEmpty, &vEmpty, &vEmpty, &vEmpty);
	if(!SUCCEEDED(hr)) return;
}

bool CIEDlgDlg::IsBackCompatDocument()
{
	IDispatch *pDisp;
	HRESULT hr = m_pBrowser2->get_Document(&pDisp);
	if(hr != S_OK || !pDisp) return false;
	IHTMLDocument5 *pDoc5;
	pDisp->QueryInterface(IID_IHTMLDocument5, (LPVOID *)&pDoc5);	 
	pDisp->Release();
	if(!pDoc5) return false;
	BSTR bstr;
	char buff[255]={0};
	pDoc5->get_compatMode(&bstr);
    pDoc5->Release();
	if(SysStringLen(bstr)) WideCharToMultiByte(CP_ACP, 0, bstr, -1, buff, 255, 0, 0);
	if(!strcmp("BackCompat", buff)) return true;
	return false;
}