/*
  client.cc
*/

#include "client.h"
#include "v8variant.h"
#include <node.h>
#include <nan.h>

using namespace v8;
using namespace ole32core;

namespace node_win32ole {

Nan::Persistent<FunctionTemplate> Client::clazz;

void Client::Init(Handle<Object> target)
{
  Local<FunctionTemplate> t = Nan::New<FunctionTemplate>(New);
  t->InstanceTemplate()->SetInternalFieldCount(2);
  t->SetClassName(Nan::New("Client").ToLocalChecked());
//  Nan::SetPrototypeMethod(clazz, "New", New);
  Nan::SetPrototypeMethod(t, "Dispatch", Dispatch);
  Nan::SetPrototypeMethod(t, "Finalize", Finalize);
  Nan::Set(target, Nan::New("Client").ToLocalChecked(), Nan::GetFunction(t).ToLocalChecked());
  clazz.Reset(t);
}

NAN_METHOD(Client::New)
{
  DISPFUNCIN();
  if(!info.IsConstructCall())
    return Nan::ThrowError(Exception::TypeError(
      Nan::New("Use the new operator to create new Client objects").ToLocalChecked()));
  std::string cstr_locale(".ACP"); // default
  if(info.Length() >= 1){
    if(!info[0]->IsString())
      return Nan::ThrowError(Exception::TypeError(
        Nan::New("Argument 1 is not a String").ToLocalChecked()));
    String::Utf8Value u8s_locale(info[0]);
    cstr_locale = std::string(*u8s_locale);
  }
  OLE32core *oc = new OLE32core();
  if(!oc)
    return Nan::ThrowError(Exception::TypeError(
      Nan::New("Can't create new Client object (null OLE32core)").ToLocalChecked()));
  bool cnresult = false;
  try{
    cnresult = oc->connect(cstr_locale);
  }catch(OLE32coreException e){
    std::cerr << e.errorMessage((char *)cstr_locale.c_str());
  }catch(char *e){
    std::cerr << e << cstr_locale.c_str() << std::endl;
  }
  if(!cnresult)
    return Nan::ThrowError(Exception::TypeError(
      Nan::New("May be CoInitialize() is failed.").ToLocalChecked()));
  Local<Object> thisObject = info.This();
  Client *cl = new Client(); // must catch exception
  cl->Wrap(thisObject); // InternalField[0]
  thisObject->SetInternalField(1, Nan::New<External>(oc));
  // dunno what to do with this
  // thisObject.SetWeak(oc, Dispose);
  DISPFUNCOUT();
 return info.GetReturnValue().Set(info.This());
}

NAN_METHOD(Client::Dispatch)
{
  DISPFUNCIN();
  BEVERIFY(done, info.Length() >= 1);
  BEVERIFY(done, info[0]->IsString());
  wchar_t *wcs;
  {
    String::Utf8Value u8s(info[0]); // must create here
    wcs = u8s2wcs(*u8s);
  }
  BEVERIFY(done, wcs);
#ifdef DEBUG
  char *mbs = wcs2mbs(wcs);
  if(!mbs) free(wcs);
  BEVERIFY(done, mbs);
  fprintf(stderr, "ProgID: %s\n", mbs);
  free(mbs);
#endif
  CLSID clsid;
  HRESULT hr = CLSIDFromProgID(wcs, &clsid);
  free(wcs);
  BEVERIFY(done, !FAILED(hr));
#ifdef DEBUG
  fprintf(stderr, "clsid:"); // 00024500-0000-0000-c000-000000000046 (Excel) ok
  for(int i = 0; i < sizeof(CLSID); ++i)
    fprintf(stderr, " %02x", ((unsigned char *)&clsid)[i]);
  fprintf(stderr, "\n");
#endif
  Handle<Object> vApp = V8Variant::CreateUndefined();
  BEVERIFY(done, !vApp.IsEmpty());
  BEVERIFY(done, !vApp->IsUndefined());
  BEVERIFY(done, vApp->IsObject());
  OCVariant *app = castedInternalField<OCVariant>(vApp);
  CHECK_OCV(app);
  app->v.vt = VT_DISPATCH;
  // When 'CoInitialize(NULL)' is not called first (and on the same instance),
  // next functions will return many errors.
  // (old style) GetActiveObject() returns 0x000036b7
  //   The requested lookup key was not found in any active activation context.
  // (OLE2) CoCreateInstance() returns 0x000003f0
  //   An attempt was made to reference a token that does not exist.
  REFIID riid = IID_IDispatch; // can't connect to Excel etc with IID_IUnknown
  // C -> C++ changes types (&clsid -> clsid, &IID_IDispatch -> IID_IDispatch)
  // options (CLSCTX_INPROC_SERVER CLSCTX_INPROC_HANDLER CLSCTX_LOCAL_SERVER)
  DWORD ctx = CLSCTX_INPROC_SERVER|CLSCTX_LOCAL_SERVER;
  hr = CoCreateInstance(clsid, NULL, ctx, riid, (void **)&app->v.pdispVal);
  if(FAILED(hr)){
    // Retry with WOW6432 bridge option.
    // This may not be a right way, but better.
    BDISPFUNCDAT("FAILED CoCreateInstance: %d: 0x%08x\n", 0, hr);
#if defined(_WIN64)
    ctx |= CLSCTX_ACTIVATE_32_BIT_SERVER; // 32bit COM server on 64bit OS
#else
    ctx |= CLSCTX_ACTIVATE_64_BIT_SERVER; // 64bit COM server on 32bit OS
#endif
    hr = CoCreateInstance(clsid, NULL, ctx, riid, (void **)&app->v.pdispVal);
  }
  if(FAILED(hr)) BDISPFUNCDAT("FAILED CoCreateInstance: %d: 0x%08x\n", 1, hr);
  BEVERIFY(done, !FAILED(hr));
  DISPFUNCOUT();
  return info.GetReturnValue().Set(vApp);
done:
  DISPFUNCOUT();
  Nan::ThrowError(Exception::TypeError(Nan::New("Dispatch failed").ToLocalChecked()));
}

NAN_METHOD(Client::Finalize)
{
  DISPFUNCIN();
#if(0)
  std::cerr << __FUNCTION__ << " Finalizer is called\a" << std::endl;
  std::cerr.flush();
#endif
  Local<Object> thisObject = info.This();
#if(0)
  Client *cl = node::ObjectWrap::Unwrap<Client>(thisObject);
  if(cl) delete cl; // it has been already deleted ?
  thisObject->SetInternalField(0, External::New(NULL));
#endif
#if(1) // now GC will call Disposer automatically
  OLE32core *oc = castedInternalField<OLE32core>(thisObject);
  if(oc){
    try{
      delete oc; // will call oc->disconnect();
    }catch(OLE32coreException e){ std::cerr << e.errorMessage(__FUNCTION__);
    }catch(char *e){ std::cerr << e << __FUNCTION__ << std::endl;
    }
  }
#endif
  thisObject->SetInternalField(1, ExternalNew(NULL));
  DISPFUNCOUT();
 return info.GetReturnValue().Set(info.This());
}

template<typename T>
void Client::Dispose(const Nan::WeakCallbackInfo<T> &data)
{
  Handle<Object> handle = data.GetValue();
  DISPFUNCIN();
#if(0)
//  std::cerr << __FUNCTION__ << " Disposer is called\a" << std::endl;
  std::cerr << __FUNCTION__ << " Disposer is called" << std::endl;
  std::cerr.flush();
#endif
  Local<Object> thisObject = handle->ToObject();
#if(0) // it has been already deleted ?
  Client *cl = node::ObjectWrap::Unwrap<Client>(thisObject);
  if(!cl){
    std::cerr << __FUNCTION__;
    std::cerr << " InternalField[0] has been already deleted" << std::endl;
    std::cerr.flush();
  }else delete cl; // it has been already deleted ?
  BEVERIFY(done, thisObject->InternalFieldCount() > 0);
  thisObject->SetInternalField(0, External::New(NULL));
#endif
  OLE32core *p = castedInternalField<OLE32core>(thisObject);
  if(!p){
    std::cerr << __FUNCTION__;
    std::cerr << " InternalField[1] has been already deleted" << std::endl;
    std::cerr.flush();
  }
//  else{
    OLE32core *oc = data.GetParameter(); // oc may be same as p
    if(oc){
      try{
        delete oc; // will call oc->disconnect();
      }catch(OLE32coreException e){ std::cerr << e.errorMessage(__FUNCTION__);
      }catch(char *e){ std::cerr << e << __FUNCTION__ << std::endl;
      }
    }
//  }
  BEVERIFY(done, thisObject->InternalFieldCount() > 1);
  thisObject->SetInternalField(1, ExternalNew(NULL));
done:
  DISPFUNCOUT();
}

void Client::Finalize()
{
  assert(!finalized);
  finalized = true;
}

} // namespace node_win32ole
