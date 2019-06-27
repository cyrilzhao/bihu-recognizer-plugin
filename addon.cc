#include <node.h>
#include <atlbase.h>
#include <mshtml.h>
#include "exdisp.h"

namespace helloWorld {

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
		  if (body.size() > 1) {
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
	  long hresult;
  };

  std::vector<UrlFrame> FindFromShell(std::vector<std::string> whitelist)
  {
	  std::vector<UrlFrame> frames;

	  CComQIPtr<IShellWindows> spShellWin;
	  HRESULT hr = spShellWin.CoCreateInstance(CLSID_ShellWindows);
	  if (FAILED(hr))    return frames;

	  long nCount = 0;
	  spShellWin->get_Count(&nCount);   // ȡ�������ʵ������

	  for (long i = 0; i < nCount; i++)
	  {
		  UrlFrame urlFrame;
		  CComQIPtr< IDispatch > spDisp;
		  hr = spShellWin->Item(CComVariant(i), &spDisp);
		  if (FAILED(hr)) continue;

		  CComQIPtr< IWebBrowser2 > spBrowser = (CComQIPtr< IWebBrowser2 >)spDisp;
		  if (!spBrowser) continue;

		  spDisp.Release();
		  hr = spBrowser->get_Document(&spDisp);
		  if (FAILED(hr)) continue;

		  CComQIPtr< IHTMLDocument2 > spDoc = (CComQIPtr< IHTMLDocument2 >)spDisp;
		  if (!spDoc) continue;

		  CComBSTR url;
		  hr = spDoc->get_URL(&url);
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
	  MultiByteToWideChar(CP_ACP, 0, strGBK.c_str(), -1, str1, n); n = WideCharToMultiByte(CP_UTF8, 0, str1, -1, NULL, 0, NULL, NULL);
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
	  // isolate��ǰ��V8ִ�л�����ÿ��isolateִ�л����໥����
	  Isolate* isolate = args.GetIsolate();

	  // ���ټ�鴫��Ĳ����ĸ���
	  //if (args.Length() < 1) {
		 // // �׳�һ�����󲢴��ص� JavaScript
		 // isolate->ThrowException(Exception::TypeError(
			//  String::NewFromUtf8(isolate, "��������������")));
		 // return;
	  //}

	  std::vector<std::string> urls;
	  for (int i = 0; i < args.Length(); i++) {
		  String::Utf8Value value(args[i]->ToString());
		  urls.push_back(std::string(*value));
	  }

	  std::vector<UrlFrame> urlFrames = FindFromShell(urls);

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

    // �趨����ֵ
    args.GetReturnValue().Set(result);
  }
  
  static void init(Local<Object> exports, Local<Object> module) {
	  HRESULT hr;

	  hr = CoInitialize(NULL);
    // �趨module.exportsΪHelloWorld����
    NODE_SET_METHOD(module, "exports", GetAllFrames);
  }
  // ���е� Node.js ���������������ʽģʽ�ĳ�ʼ������
  NODE_MODULE(NODE_GYP_MODULE_NAME, init)

}
