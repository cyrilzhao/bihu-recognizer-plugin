#include <node.h>
#include <atlbase.h>
#include <mshtml.h>
#include "exdisp.h"
#include "comdef.h"

namespace addon {

  using namespace v8;

  std::string ConvertWCSToMBS(const wchar_t* pstr, long wslen)
  {
	  int len = ::WideCharToMultiByte(CP_ACP, 0, pstr, wslen, NULL, 0, NULL, NULL);

	  std::string dblstr(len, '\0');
	  len = ::WideCharToMultiByte(CP_ACP, 0,
		  pstr, wslen,
		  &dblstr[0], len,
		  NULL, NULL);

	  return dblstr;
  }

  std::string ConvertBSTRToMBS(BSTR bstr)
  {
	  int wslen = ::SysStringLen(bstr);
	  return ConvertWCSToMBS((wchar_t*)bstr, wslen);
  }

  // Converts a IHTMLWindow2 object to a IWebBrowser2. Returns NULL in case of failure.
  CComQIPtr<IWebBrowser2> HtmlWindowToHtmlWebBrowser(CComQIPtr<IHTMLWindow2> spWindow)
  {
	  ATLASSERT(spWindow != NULL);

	  CComQIPtr<IServiceProvider>  spServiceProvider = (CComQIPtr<IServiceProvider>)spWindow;
	  if (spServiceProvider == NULL)
	  {
		  return CComQIPtr<IWebBrowser2>();
	  }

	  CComQIPtr<IWebBrowser2> spWebBrws;
	  HRESULT hRes = spServiceProvider->QueryService(IID_IWebBrowserApp, IID_IWebBrowser2, (void**)& spWebBrws);
	  if (hRes != S_OK)
	  {
		  return CComQIPtr<IWebBrowser2>();
	  }

	  return spWebBrws;
  }

  // Converts a IHTMLWindow2 object to a IHTMLDocument2. Returns NULL in case of failure.
  // It takes into account accessing the DOM across frames loaded from different domains.
  CComQIPtr<IHTMLDocument2> HtmlWindowToHtmlDocument(CComQIPtr<IHTMLWindow2> spWindow)
  {
	  ATLASSERT(spWindow != NULL);

	  CComQIPtr<IHTMLDocument2> spDocument;
	  HRESULT hRes = spWindow->get_document(&spDocument);

	  if ((S_OK == hRes) && (spDocument != NULL))
	  {
		  // The html document was properly retrieved.
		  return spDocument;
	  }

	  // hRes could be E_ACCESSDENIED that means a security restriction that
	  // prevents scripting across frames that loads documents from different internet domains.
	  CComQIPtr<IWebBrowser2>  spBrws = HtmlWindowToHtmlWebBrowser(spWindow);
	  if (spBrws == NULL)
	  {
		  return CComQIPtr<IHTMLDocument2>();
	  }

	  // Get the document object from the IWebBrowser2 object.
	  CComQIPtr<IDispatch> spDisp;
	  hRes = spBrws->get_Document(&spDisp);
	  spDocument = spDisp;

	  return spDocument;
  }

  std::vector<std::string> GetFrames(CComQIPtr<IHTMLDocument2> pDoc) {
	  std::vector<std::string> frames;

	  CComQIPtr<IHTMLElement> body;
	  HRESULT result = pDoc->get_body(&body);
	  if (SUCCEEDED(result)) {
		  CComBSTR strBody;
		  result = body->get_outerHTML(&strBody);
		  if (SUCCEEDED(result)) {
			  frames.push_back(ConvertBSTRToMBS(strBody));
		  }
	  }

	  CComQIPtr<IHTMLFramesCollection2> pFrames;
	  pDoc->get_frames(&pFrames);
	  if (NULL == pFrames) {
		  return frames;
	  }
	  long len = 0;
	  VARIANT varIndex, varResult;
	  pFrames->get_length(&len);
	  for (int i = 0; i < len; i++) {
		  varIndex.vt = VT_I4;
		  varIndex.llVal = i;
		  pFrames->item(&varIndex, &varResult);

		  CComQIPtr<IHTMLWindow2> frameWindow;
		  HRESULT hr;
		  hr = varResult.pdispVal->QueryInterface(IID_IHTMLWindow2, (void**)& frameWindow);
		  if (FAILED(hr)) continue;

		  CComQIPtr<IHTMLDocument2> pDoc = HtmlWindowToHtmlDocument(frameWindow);

		  CComQIPtr<IHTMLElement> pBody;
		  hr = pDoc->get_body(&pBody);
		  if (FAILED(hr)) continue;
		  CComBSTR bstrBody;
		  hr = pBody->get_outerHTML(&bstrBody);
		  if (FAILED(hr)) continue;
		  std::string body = ConvertBSTRToMBS(bstrBody);
		  if (body.size() > 100) {
			  frames.push_back(body);
		  }

		  std::vector<std::string> subFrames = GetFrames(pDoc);
		  frames.insert(frames.end(), subFrames.begin(), subFrames.end());
	  }
	  return frames;
  }

  bool StartsWith(std::string left, std::string right) {
	  if (left.size() < right.size()) {
		  return false;
	  }
	  for (int i = 0; i < right.size(); i++) {
		  if (left[i] != right[i]) {
			  return false;
		  }
	  }
	  return true;
  }

  bool InWhitelist(std::vector<std::string> whitelist, std::string url) {
	  if (whitelist.size() == 0) {
		  return true;
	  }
	  for (int i = 0; i < whitelist.size(); i++) {
		  if (StartsWith(url, whitelist[i])) {
			  return true;
		  }
	  }
	  return false;
  }

  struct UrlFrame {
	  std::string url;
	  std::vector<std::string> frames;
  };

  BOOL CALLBACK EnumChildProc(HWND hwnd, LPARAM lParam) {
	  CHAR szClassName[100];
	  ::GetClassName(hwnd, szClassName, sizeof(szClassName));
	  if (_tcscmp(szClassName, _T("Internet Explorer_Server")) == 0) {
		  //*(HWND*)lParam = hwnd;
		  std::vector<HWND>* list = (std::vector<HWND>*)lParam;
		  list->push_back(hwnd);
		  //return FALSE;
	  }
	  return TRUE;
  }

  std::vector<CComQIPtr<IHTMLDocument2>> FindFromWindow(CHAR* clsName) {
	  std::vector<CComQIPtr<IHTMLDocument2>> pDocs;

	  HWND hWnd = ::FindWindowEx(0, 0, clsName, 0);
	  std::vector<HWND> children;
	  ::EnumChildWindows(hWnd, EnumChildProc, (LPARAM)& children);

	  for (int i = 0; i < children.size(); i++) {
		  HWND hWndChild = children[i];
		  UINT nMsg = ::RegisterWindowMessage(_T("WM_HTML_GETOBJECT"));
		  LRESULT lRes;
		  ::SendMessageTimeout(hWndChild, nMsg, 0L, 0L, SMTO_ABORTIFHUNG, 1000, (PDWORD_PTR)& lRes);
		  DWORD e = GetLastError();
		  CComQIPtr <IHTMLDocument2> spDoc;
		  ::ObjectFromLresult(lRes, IID_IHTMLDocument2, 0, (LPVOID*)& spDoc);
		  if (spDoc != NULL) {
			  pDocs.push_back(spDoc);
		  }
	  }

	  return pDocs;
  }

  std::vector<CComQIPtr<IHTMLDocument2>> FindFromShell()
  {
	  std::vector<CComQIPtr<IHTMLDocument2>> pDocs;

	  CComQIPtr<IShellWindows> spShellWin;
	  HRESULT hr = spShellWin.CoCreateInstance(CLSID_ShellWindows);
	  if (FAILED(hr))    return pDocs;

	  long nCount = 0;
	  spShellWin->get_Count(&nCount);   // 取得浏览器实例个数

	  for (long i = 0; i < nCount; i++)
	  {
		  CComQIPtr< IDispatch > spDisp;
		  hr = spShellWin->Item(CComVariant(i), &spDisp);
		  if (FAILED(hr)) continue;

		  CComQIPtr< IWebBrowser2 > spBrowser(spDisp);
		  if (!spBrowser) continue;

		  CComQIPtr< IDispatch > dispDoc;
		  hr = spBrowser->get_Document(&dispDoc);
		  if (FAILED(hr)) continue;

		  CComQIPtr< IHTMLDocument2 > spDoc(dispDoc);
		  if (!spDoc) continue;

		  pDocs.push_back(spDoc);
	  }
	  
	  return pDocs;
  }

  std::vector<UrlFrame> FindAllFrams(std::vector<std::string> whitelist, bool fromShell)
  {
	  std::vector<UrlFrame> frames;

	  std::vector<CComQIPtr<IHTMLDocument2>> pDocs;
	  if (fromShell) {
		  pDocs = FindFromShell();
	  }
	  else {
		  pDocs = FindFromWindow("IEFrame");
		  std::vector<CComQIPtr<IHTMLDocument2>> seDocs = FindFromWindow("360se6_Frame");
		  pDocs.insert(pDocs.end(), seDocs.begin(), seDocs.end());
	  }
	  for (long i = 0; i < pDocs.size(); i++)
	  {
		  CComQIPtr< IHTMLDocument2 > spDoc = pDocs[i];
		  if (!spDoc) continue;

		  CComBSTR url;
		  HRESULT hr = spDoc->get_URL(&url);
		  if (FAILED(hr)) continue;

		  std::string urlStr = ConvertBSTRToMBS(url);
		  if (InWhitelist(whitelist, urlStr)) {
			  std::vector<std::string> subFrames = GetFrames(spDoc);
			  UrlFrame urlFrame = {
				  urlStr,
				  subFrames
			  };
			  frames.push_back(urlFrame);
		  }
	  }
	  return frames;
  }

  std::string GBKToUTF8(std::string& strGBK)
  {
	  std::string strOutUTF8 = "";
	  WCHAR* str1;
	  int n = MultiByteToWideChar(CP_ACP, 0, strGBK.c_str(), -1, NULL, 0);
	  str1 = new WCHAR[n];
	  MultiByteToWideChar(CP_ACP, 0, strGBK.c_str(), -1, str1, n);
	  n = WideCharToMultiByte(CP_UTF8, 0, str1, -1, NULL, 0, NULL, NULL);
	  char* str2 = new char[n];
	  WideCharToMultiByte(CP_UTF8, 0, str1, -1, str2, n, NULL, NULL);
	  strOutUTF8 = str2;
	  delete[]str1;
	  str1 = NULL;
	  delete[]str2;
	  str2 = NULL;
	  return strOutUTF8;
  }

  static void GetAllFrames(const FunctionCallbackInfo<Value>& args) {
	  // isolate当前的V8执行环境，每个isolate执行环境相互独立
	  Isolate* isolate = args.GetIsolate();

	  // 不再检查传入的参数的个数
	  //if (args.Length() < 1) {
		 // // 抛出一个错误并传回到 JavaScript
		 // isolate->ThrowException(Exception::TypeError(
			//  String::NewFromUtf8(isolate, "参数的数量错误")));
		 // return;
	  //}

	  std::vector<std::string> urls;
	  if (args.Length() > 0 && args[0]->IsArray()) {
		  Local<Array> input = Local<Array>::Cast(args[0]);
		  for (int i = 0; i < input->Length(); i++) {
			  String::Utf8Value value(input->Get(i)->ToString());
			  urls.push_back(std::string(*value));
		  }
	  }
	  bool fromShell = true;
	  if (args.Length() > 1) {
		  fromShell = args[1]->BooleanValue();
	  }

	  std::vector<UrlFrame> urlFrames = FindAllFrams(urls, fromShell);

	  Local<Array> result = Array::New(isolate);
	  for (int i = 0; i < urlFrames.size(); i++) {
		  Local<Object> ob = Object::New(isolate);
		  ob->Set(String::NewFromUtf8(isolate, "url"), String::NewFromUtf8(isolate, urlFrames[i].url.c_str()));

		  Local<Array> frames = Array::New(isolate);
		  for (int j = 0; j < urlFrames[i].frames.size(); j++) {
			  std::string f = GBKToUTF8(urlFrames[i].frames[j]);
			  frames->Set(j, String::NewFromUtf8(isolate, f.c_str()));
		  }
		  ob->Set(String::NewFromUtf8(isolate, "frames"), frames);

		  result->Set(i, ob);
		 // MessageBox(NULL, frames[i].c_str(), "html", MB_OK);
	  }

    // 设定返回值
    args.GetReturnValue().Set(result);
  }
  
  static void init(Local<Object> exports, Local<Object> module) {
	  HRESULT hr;

	  hr = CoInitialize(NULL);
    // 设定module.exports为HelloWorld函数
    NODE_SET_METHOD(module, "exports", GetAllFrames);
  }
  // 所有的 Node.js 插件必须以以下形式模式的初始化函数
  NODE_MODULE(NODE_GYP_MODULE_NAME, init)

}
