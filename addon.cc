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
	  len = ::WideCharToMultiByte(CP_ACP, 0 /* no flags */,
		  pstr, wslen /* not necessary NULL-terminated */,
		  &dblstr[0], len,
		  NULL, NULL /* no default char */);

	  return dblstr;
  }

  std::string ConvertBSTRToMBS(BSTR bstr)
  {
	  int wslen = ::SysStringLen(bstr);
	  return ConvertWCSToMBS((wchar_t*)bstr, wslen);
  }

  std::vector<std::string> getFrames(CComPtr<IHTMLDocument2> pDoc) {
	  std::vector<std::string> frames;
	  CComPtr<IHTMLFramesCollection2> pFrames;
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

		  CComPtr<IHTMLWindow2> frameWindow;
		  varResult.pdispVal->QueryInterface(IID_IHTMLWindow2, (void**)& frameWindow);
		  CComPtr<IHTMLDocument2> pDoc;
		  frameWindow->get_document(&pDoc);
		  CComBSTR bstrTitle;
		  pDoc->get_title(&bstrTitle);

		  CComPtr<IHTMLElement> pBody;
		  pDoc->get_body(&pBody);
		  CComBSTR bstrBody;
		  pBody->get_outerHTML(&bstrBody);
		  std::string body = ConvertBSTRToMBS(bstrBody);
		  if (body.size() > 100) {
			  frames.push_back(body);
		  }

		  std::vector<std::string> subFrames = getFrames(pDoc);
		  frames.insert(frames.end(), subFrames.begin(), subFrames.end());
	  }
	  return frames;
  }

  bool startWith(std::string left, std::string right) {
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

  bool inWhitelist(std::vector<std::string> whitelist, std::string url) {
	  for (int i = 0; i < whitelist.size(); i++) {
		  if (startWith(url, whitelist[i])) {
			  return true;
		  }
	  }
	  return false;
  }

  struct UrlFrame {
	  std::string url;
	  std::vector<std::string> frames;
  };

  std::vector<UrlFrame> FindFromShell(std::vector<std::string> whitelist)
  {
	  std::vector<UrlFrame> frames;

	  CComPtr<IShellWindows> spShellWin;
	  HRESULT hr = spShellWin.CoCreateInstance(CLSID_ShellWindows);
	  if (FAILED(hr))    return frames;

	  long nCount = 0;
	  spShellWin->get_Count(&nCount);   // 取得浏览器实例个数

	  for (long i = 0; i < nCount; i++)
	  {

		  CComPtr< IDispatch > spDisp;
		  hr = spShellWin->Item(CComVariant(i), &spDisp);
		  if (FAILED(hr))   continue;

		  CComQIPtr< IWebBrowser2 > spBrowser = (CComQIPtr< IWebBrowser2 >)spDisp;
		  if (!spBrowser)     continue;

		  spDisp.Release();
		  hr = spBrowser->get_Document(&spDisp);
		  if (FAILED(hr))  continue;

		  CComQIPtr< IHTMLDocument2 > spDoc = (CComQIPtr< IHTMLDocument2 >)spDisp;
		  if (!spDoc)         continue;

		  // 程序运行到此，已经找到了 IHTMLDocument2 的接口指针
		  CComBSTR bstrTitle;
		  spDoc->get_URL(&bstrTitle);

		  CComBSTR url;
		  spDoc->get_URL(&url);
		  std::string urlStr = ConvertBSTRToMBS(url);
		  if (inWhitelist(whitelist, urlStr)) {
			  std::vector<std::string> subFrames = getFrames(spDoc);
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
	  // isolate当前的V8执行环境，每个isolate执行环境相互独立
	  Isolate* isolate = args.GetIsolate();

	  // 检查传入的参数的个数
	  if (args.Length() < 1) {
		  // 抛出一个错误并传回到 JavaScript
		  isolate->ThrowException(Exception::TypeError(
			  String::NewFromUtf8(isolate, "参数的数量错误")));
		  return;
	  }

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
