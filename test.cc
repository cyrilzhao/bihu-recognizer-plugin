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

  std::vector<std::string> FindFromShell(std::vector<std::string> whitelist)
  {
	  std::vector<std::string> frames;

	  CComPtr<IShellWindows> spShellWin;
	  HRESULT hr = spShellWin.CoCreateInstance(CLSID_ShellWindows);
	  if (FAILED(hr))    return frames;

	  long nCount = 0;
	  spShellWin->get_Count(&nCount);   // ȡ�������ʵ������

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

		  // �������е��ˣ��Ѿ��ҵ��� IHTMLDocument2 �Ľӿ�ָ��
		  CComBSTR bstrTitle;
		  spDoc->get_URL(&bstrTitle);

		  CComBSTR url;
		  spDoc->get_URL(&url);
		  if (inWhitelist(whitelist, ConvertBSTRToMBS(url))) {
			  std::vector<std::string> subFrames = getFrames(spDoc);
			  frames.insert(frames.end(), subFrames.begin(), subFrames.end());
		  }
	  }
	  return frames;
  }


  static void HelloWorld(const FunctionCallbackInfo<Value>& args) {
	  // isolate��ǰ��V8ִ�л�����ÿ��isolateִ�л����໥����
	  Isolate* isolate = args.GetIsolate();

	  // ��鴫��Ĳ����ĸ���
	  if (args.Length() < 1) {
		  // �׳�һ�����󲢴��ص� JavaScript
		  isolate->ThrowException(Exception::TypeError(
			  String::NewFromUtf8(isolate, "��������������")));
		  return;
	  }

	  std::vector<std::string> urls;
	  for (int i = 0; i < args.Length(); i++) {
		  String::Utf8Value value(args[i]->ToString());
		  urls.push_back(std::string(*value));
	  }

	  std::vector<std::string> frames = FindFromShell(urls);

	  Local<Array> result = Array::New(isolate);
	  for (int i = 0; i < frames.size(); i++) {
		  result->Set(i, String::NewFromUtf8(isolate, frames[i].c_str()));
		 // MessageBox(NULL, frames[i].c_str(), "html", MB_OK);
	  }

    // �趨����ֵ
    args.GetReturnValue().Set(result);
  }

  void FindFromShell()
  {
	  CComPtr< IShellWindows > spShellWin;
	  HRESULT hr = spShellWin.CoCreateInstance(CLSID_ShellWindows);
	  if (FAILED(hr))    return;

	  long nCount = 0;
	  spShellWin->get_Count(&nCount);   // ȡ�������ʵ������

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

		  // �������е��ˣ��Ѿ��ҵ��� IHTMLDocument2 �Ľӿ�ָ��
		  BSTR bstrTitle;
		  spDoc->get_title(&bstrTitle);
		  ::SysFreeString(bstrTitle);


	  }
  }

  char* getIFrame(CComPtr<IHTMLDocument2> pDoc) {
	  CComPtr<IHTMLFramesCollection2> pFrames;
	  pDoc->get_frames(&pFrames);
	  if (NULL == pFrames) {
		  return NULL;
	  }
	  long len = 0;
	  VARIANT varIndex, varResult;
	  pFrames->get_length(&len);
	  for (int i = 0; i < len; i++) {
		  varIndex.vt = VT_I4;
		  varIndex.llVal = i;
		  pFrames->item(&varIndex, &varResult);

		  IHTMLWindow2* frameWindow;
		  varResult.pdispVal->QueryInterface(IID_IHTMLWindow2, (void**)& frameWindow);
		  IHTMLDocument2* pDoc;
		  frameWindow->get_document(&pDoc);
		  BSTR bstrTitle;
		  pDoc->get_title(&bstrTitle);
		  ::SysFreeString(bstrTitle);
	  }
  }
  
  static void init(Local<Object> exports, Local<Object> module) {
	  HRESULT hr;

	  hr = CoInitialize(NULL);
    // �趨module.exportsΪHelloWorld����
    NODE_SET_METHOD(module, "exports", HelloWorld);
  }
  // ���е� Node.js ���������������ʽģʽ�ĳ�ʼ������
  NODE_MODULE(NODE_GYP_MODULE_NAME, init)

}
