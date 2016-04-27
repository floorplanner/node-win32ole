/*
  v8variant.cc
*/

#include "v8variant.h"
#include <node.h>
#include <nan.h>

using namespace v8;
using namespace ole32core;

namespace node_win32ole {

#define CHECK_OLE_ARGS(info, n, av0, av1) do{ \
    if(info.Length() < n) \
      return Nan::ThrowError(Exception::TypeError( \
        Nan::New(__FUNCTION__" takes exactly " #n " argument(s)").ToLocalChecked())); \
    if(!info[0]->IsString()) \
      return Nan::ThrowError(Exception::TypeError( \
        Nan::New(__FUNCTION__" the first argument is not a Symbol").ToLocalChecked())); \
    if(n == 1) \
      if(info.Length() >= 2) \
        if(!info[1]->IsArray()) \
          return Nan::ThrowError(Exception::TypeError(Nan::New( \
            __FUNCTION__" the second argument is not an Array").ToLocalChecked())); \
        else av1 = info[1]; /* Array */ \
      else av1 = Nan::New<Array>(0); /* change none to Array[] */ \
    else av1 = info[1]; /* may not be Array */ \
    av0 = info[0]; \
  }while(0)

Nan::Persistent<FunctionTemplate> V8Variant::clazz;

void V8Variant::Init(Handle<Object> target)
{
  Local<FunctionTemplate> t = Nan::New<FunctionTemplate>(New);
  t->InstanceTemplate()->SetInternalFieldCount(2);
  t->SetClassName(Nan::New("V8Variant").ToLocalChecked());
  Nan::SetPrototypeMethod(t, "isA", OLEIsA);
  Nan::SetPrototypeMethod(t, "vtName", OLEVTName);
  Nan::SetPrototypeMethod(t, "toBoolean", OLEBoolean);
  Nan::SetPrototypeMethod(t, "toInt32", OLEInt32);
  Nan::SetPrototypeMethod(t, "toInt64", OLEInt64);
  Nan::SetPrototypeMethod(t, "toNumber", OLENumber);
  Nan::SetPrototypeMethod(t, "toDate", OLEDate);
  Nan::SetPrototypeMethod(t, "toUtf8", OLEUtf8);
  Nan::SetPrototypeMethod(t, "toValue", OLEValue);
//  Nan::SetPrototypeMethod(t, "New", New);
  Nan::SetPrototypeMethod(t, "call", OLECall);
  Nan::SetPrototypeMethod(t, "get", OLEGet);
  Nan::SetPrototypeMethod(t, "set", OLESet);
/*
 In ParseUnaryExpression() < v8/src/parser.cc >
 v8::Object::ToBoolean() is called directly for unary operator '!'
 instead of v8::Object::valueOf()
 so NamedPropertyHandler will not be called
 Local<Boolean> ToBoolean(); // How to fake ? override v8::Value::ToBoolean
*/
  Local<ObjectTemplate> instancetpl = t->InstanceTemplate();
  Nan::SetCallAsFunctionHandler(instancetpl, OLECallComplete);
  Nan::SetNamedPropertyHandler(instancetpl, OLEGetAttr, OLESetAttr);
//  Nan::SetIndexedPropertyHandler(instancetpl, OLEGetIdxAttr, OLESetIdxAttr);
  Nan::SetPrototypeMethod(t, "Finalize", Finalize);
  Nan::Set(target, Nan::New("V8Variant").ToLocalChecked(), Nan::GetFunction(t).ToLocalChecked());
  clazz.Reset(t);
}

std::string V8Variant::CreateStdStringMBCSfromUTF8(Handle<Value> v)
{
  String::Utf8Value u8s(v);
  wchar_t * wcs = u8s2wcs(*u8s);
  if(!wcs){
    std::cerr << "[Can't allocate string (wcs)]" << std::endl;
    std::cerr.flush();
    return std::string("'!ERROR'");
  }
  char *mbs = wcs2mbs(wcs);
  if(!mbs){
    free(wcs);
    std::cerr << "[Can't allocate string (mbs)]" << std::endl;
    std::cerr.flush();
    return std::string("'!ERROR'");
  }
  std::string s(mbs);
  free(mbs);
  free(wcs);
  return s;
}

OCVariant *V8Variant::CreateOCVariant(Handle<Value> v)
{
  if (v->IsNull() || v->IsUndefined()) {
    // todo: make separate undefined type
		return new OCVariant();
  }

  BEVERIFY(done, !v.IsEmpty());
  BEVERIFY(done, !v->IsExternal());
  BEVERIFY(done, !v->IsNativeError());
// NOTE: commenting the following fixes for node 6 / v8
//       needs investigaton v->IsFunction deprecated?
//  BEVERIFY(done, !v->IsFunction());
// VT_USERDEFINED VT_VARIANT VT_BYREF VT_ARRAY more...
  if(v->IsBoolean()){
    return new OCVariant(Nan::To<bool>(v).FromJust());
  }else if(v->IsArray()){
// VT_BYREF VT_ARRAY VT_SAFEARRAY
    std::cerr << "[Array (not implemented now)]" << std::endl; return NULL;
    std::cerr.flush();
  }else if(v->IsInt32()){
    return new OCVariant((long)Nan::To<int32_t>(v).FromJust());
  }else if(v->IsNumber()){
    std::cerr << "[Number (VT_R8 or VT_I8 bug?)]" << std::endl;
    std::cerr.flush();
    return new OCVariant((double)Nan::To<int64_t>(v).FromJust()); // double
  }else if(v->IsNumberObject()){
    std::cerr << "[NumberObject (VT_R8 or VT_I8 bug?)]" << std::endl;
    std::cerr.flush();
    return new OCVariant((double)Nan::To<int64_t>(v).FromJust()); // double
  }else if(v->IsDate()){
    double d = Nan::To<int64_t>(v).FromJust();
    time_t sec = (time_t)(d / 1000.0);
    int msec = (int)(d - sec * 1000.0);
    struct tm *t = localtime(&sec); // *** must check locale ***
    if(!t){
      std::cerr << "[Date may not be valid]" << std::endl;
      std::cerr.flush();
      return NULL;
    }
    SYSTEMTIME syst;
    syst.wYear = t->tm_year + 1900;
    syst.wMonth = t->tm_mon + 1;
    syst.wDay = t->tm_mday;
    syst.wHour = t->tm_hour;
    syst.wMinute = t->tm_min;
    syst.wSecond = t->tm_sec;
    syst.wMilliseconds = msec;
    SystemTimeToVariantTime(&syst, &d);
    return new OCVariant(d, true); // date
  }else if(v->IsRegExp()){
    std::cerr << "[RegExp (bug?)]" << std::endl;
    std::cerr.flush();
    return new OCVariant(CreateStdStringMBCSfromUTF8(v));
  }else if(v->IsString()){
    return new OCVariant(CreateStdStringMBCSfromUTF8(v));
  }else if(v->IsStringObject()){
    std::cerr << "[StringObject (bug?)]" << std::endl;
    std::cerr.flush();
    return new OCVariant(CreateStdStringMBCSfromUTF8(v));
  }else if(v->IsObject()){
#if(0)
    std::cerr << "[Object (test)]" << std::endl;
    std::cerr.flush();
#endif
    OCVariant *ocv = castedInternalField<OCVariant>(v->ToObject());
    if(!ocv){
      std::cerr << "[Object may not be valid (null OCVariant)]" << std::endl;
      std::cerr.flush();
      return NULL;
    }
    // std::cerr << ocv->v.vt;
    return new OCVariant(*ocv);
  }else{
    std::cerr << "[unknown type (not implemented now)]" << std::endl;
    std::cerr.flush();
  }
done:
  return NULL;
}

NAN_METHOD(V8Variant::OLEIsA)
{
  DISPFUNCIN();
  OCVariant *ocv = castedInternalField<OCVariant>(info.This());
  CHECK_OCV(ocv);
  DISPFUNCOUT();
 return info.GetReturnValue().Set(Nan::New(ocv->v.vt));
}

NAN_METHOD(V8Variant::OLEVTName)
{
  DISPFUNCIN();
  OCVariant *ocv = castedInternalField<OCVariant>(info.This());
  CHECK_OCV(ocv);
  Local<Object> target = Nan::New(module_target);
  Array *a = Array::Cast(*(GET_PROP(target, "vt_names").ToLocalChecked()));
  DISPFUNCOUT();
 return info.GetReturnValue().Set(ARRAY_AT(a, ocv->v.vt));
}

NAN_METHOD(V8Variant::OLEBoolean)
{
  DISPFUNCIN();
  OCVariant *ocv = castedInternalField<OCVariant>(info.This());
  CHECK_OCV(ocv);
  if(ocv->v.vt != VT_BOOL)
    return Nan::ThrowError(Exception::TypeError(
      Nan::New("OLEBoolean source type OCVariant is not VT_BOOL").ToLocalChecked()));
  bool c_boolVal = ocv->v.boolVal == VARIANT_FALSE ? 0 : !0;
  DISPFUNCOUT();
 return info.GetReturnValue().Set(c_boolVal);
}

NAN_METHOD(V8Variant::OLEInt32)
{
  DISPFUNCIN();
  OCVariant *ocv = castedInternalField<OCVariant>(info.This());
  CHECK_OCV(ocv);
  if(ocv->v.vt != VT_I4 && ocv->v.vt != VT_INT
  && ocv->v.vt != VT_UI4 && ocv->v.vt != VT_UINT)
    return Nan::ThrowError(Exception::TypeError(
      Nan::New("OLEInt32 source type OCVariant is not VT_I4 nor VT_INT nor VT_UI4 nor VT_UINT").ToLocalChecked()));
  DISPFUNCOUT();
 return info.GetReturnValue().Set(Nan::New(ocv->v.lVal));
}

NAN_METHOD(V8Variant::OLEInt64)
{
  DISPFUNCIN();
  OCVariant *ocv = castedInternalField<OCVariant>(info.This());
  CHECK_OCV(ocv);
  if(ocv->v.vt != VT_I8 && ocv->v.vt != VT_UI8)
    return Nan::ThrowError(Exception::TypeError(
      Nan::New("OLEInt64 source type OCVariant is not VT_I8 nor VT_UI8").ToLocalChecked()));
  DISPFUNCOUT();
 return info.GetReturnValue().Set(Nan::New<Number>(ocv->v.llVal));
}

NAN_METHOD(V8Variant::OLENumber)
{
  DISPFUNCIN();
  OCVariant *ocv = castedInternalField<OCVariant>(info.This());
  CHECK_OCV(ocv);
  if(ocv->v.vt != VT_R8)
    return Nan::ThrowError(Exception::TypeError(
      Nan::New("OLENumber source type OCVariant is not VT_R8").ToLocalChecked()));
  DISPFUNCOUT();
 return info.GetReturnValue().Set(Nan::New(ocv->v.dblVal));
}

NAN_METHOD(V8Variant::OLEDate)
{
  DISPFUNCIN();
  OCVariant *ocv = castedInternalField<OCVariant>(info.This());
  CHECK_OCV(ocv);
  if(ocv->v.vt != VT_DATE)
    return Nan::ThrowError(Exception::TypeError(
      Nan::New("OLEDate source type OCVariant is not VT_DATE").ToLocalChecked()));
  SYSTEMTIME syst;
  VariantTimeToSystemTime(ocv->v.date, &syst);
  struct tm t = {0}; // set t.tm_isdst = 0
  t.tm_year = syst.wYear - 1900;
  t.tm_mon = syst.wMonth - 1;
  t.tm_mday = syst.wDay;
  t.tm_hour = syst.wHour;
  t.tm_min = syst.wMinute;
  t.tm_sec = syst.wSecond;
  DISPFUNCOUT();
  return info.GetReturnValue().Set(Nan::New<Date>(mktime(&t) * 1000LL + syst.wMilliseconds).ToLocalChecked());
}

NAN_METHOD(V8Variant::OLEUtf8)
{
  DISPFUNCIN();
  OCVariant *ocv = castedInternalField<OCVariant>(info.This());
  CHECK_OCV(ocv);
  if(ocv->v.vt != VT_BSTR)
    return Nan::ThrowError(Exception::TypeError(
      Nan::New("OLEUtf8 source type OCVariant is not VT_BSTR").ToLocalChecked()));
  Handle<Value> result;
  if(!ocv->v.bstrVal) result = Nan::Undefined(); // or Null();
  else {
    std::wstring wstr(ocv->v.bstrVal);
    char *cs = wcs2u8s(wstr.c_str());
    result = Nan::New(cs).ToLocalChecked();
    free(cs);
  }
  DISPFUNCOUT();
 return info.GetReturnValue().Set(result);
}

NAN_METHOD(V8Variant::OLEValue)
{
  OLETRACEIN();
  OLETRACEVT(info.This());
  OLETRACEFLUSH();
  Local<Object> thisObject = info.This();
  OLE_PROCESS_CARRY_OVER(thisObject);
  OLETRACEVT(thisObject);
  OLETRACEFLUSH();
  OCVariant *ocv = castedInternalField<OCVariant>(thisObject);
  if(!ocv){ std::cerr << "ocv is null"; std::cerr.flush(); }
  CHECK_OCV(ocv);
  if(ocv->v.vt == VT_EMPTY) ; // do nothing
  else if (ocv->v.vt == VT_NULL) return info.GetReturnValue().SetNull();
  else if (ocv->v.vt == VT_DISPATCH) {
    if (ocv->v.pdispVal == NULL) {
      return info.GetReturnValue().SetNull();
    }
    return info.GetReturnValue().Set(thisObject); // through it
  }
  else if(ocv->v.vt == VT_BOOL) OLEBoolean(info);
  else if(ocv->v.vt == VT_I4 || ocv->v.vt == VT_INT
  || ocv->v.vt == VT_UI4 || ocv->v.vt == VT_UINT) OLEInt32(info);
  else if(ocv->v.vt == VT_I8 || ocv->v.vt == VT_UI8) OLEInt64(info);
  else if(ocv->v.vt == VT_R8) OLENumber(info);
  else if(ocv->v.vt == VT_DATE) OLEDate(info);
  else if(ocv->v.vt == VT_BSTR) OLEUtf8(info);
  else if(ocv->v.vt == VT_ARRAY || ocv->v.vt == VT_SAFEARRAY){
    std::cerr << "[Array (not implemented now)]" << std::endl;
    std::cerr.flush();
  }else{
    Handle<Value> s = INSTANCE_CALL(thisObject, "vtName", 0, NULL);
    std::cerr << "[unknown type " << ocv->v.vt << ":" << *String::Utf8Value(s);
    std::cerr << " (not implemented now)]" << std::endl;
    std::cerr.flush();
  }
done:
  OLETRACEOUT();
}

static std::string GetName(ITypeInfo *typeinfo, MEMBERID id) {
  BSTR name;
  UINT numNames = 0;
  typeinfo->GetNames(id, &name, 1, &numNames);
  if (numNames > 0) {
    return BSTR2MBCS(name);
  }
  return "";
}

/**
 * Like OLEValue, but goes one step further and reduces IDispatch objects
 * to actual JS objects. This enables things like console.log().
 **/
NAN_METHOD(V8Variant::OLEPrimitiveValue) {
  Local<Object> thisObject = info.This();
  OLE_PROCESS_CARRY_OVER(thisObject);
  OLETRACEVT(thisObject);
  OLETRACEFLUSH();
  OCVariant *ocv = castedInternalField<OCVariant>(thisObject);
  CHECK_OCV(ocv);
  if (ocv->v.vt == VT_DISPATCH) {
    IDispatch *dispatch = ocv->v.pdispVal;
    if (dispatch == NULL) {
      return info.GetReturnValue().SetNull();
    }
    Local<Object> object = Nan::New<Object>();
    ITypeInfo *typeinfo = NULL;
    HRESULT hr = dispatch->GetTypeInfo(0, LOCALE_USER_DEFAULT, &typeinfo);
    if (typeinfo) {
      TYPEATTR* typeattr;
      BASSERT(SUCCEEDED(typeinfo->GetTypeAttr(&typeattr)));
      for (int i = 0; i < typeattr->cFuncs; i++) {
        FUNCDESC *funcdesc;
        typeinfo->GetFuncDesc(i, &funcdesc);
        if (funcdesc->invkind != INVOKE_FUNC) {
          Nan::Set(object, Nan::New(GetName(typeinfo, funcdesc->memid)).ToLocalChecked(), Nan::New("Function").ToLocalChecked());
        }
        typeinfo->ReleaseFuncDesc(funcdesc);
      }
      for (int i = 0; i < typeattr->cVars; i++) {
        VARDESC *vardesc;
        typeinfo->GetVarDesc(i, &vardesc);
        Nan::Set(object, Nan::New(GetName(typeinfo, vardesc->memid)).ToLocalChecked(), Nan::New("Variable").ToLocalChecked());
        typeinfo->ReleaseVarDesc(vardesc);
      }
      typeinfo->ReleaseTypeAttr(typeattr);
    }
    return info.GetReturnValue().Set(object);
  } else {
    V8Variant::OLEValue(info);
  }
}

Handle<Object> V8Variant::CreateUndefined(void)
{
  DISPFUNCIN();
  Local<FunctionTemplate> localClazz = Nan::New(clazz);
  Local<Object> instance = Nan::NewInstance(Nan::GetFunction(localClazz).ToLocalChecked(), 0, NULL).ToLocalChecked();
  DISPFUNCOUT();
  return instance;
}

NAN_METHOD(V8Variant::New)
{
  DISPFUNCIN();
  if(!info.IsConstructCall())
    return Nan::ThrowError(Exception::TypeError(
      Nan::New("Use the new operator to create new V8Variant objects").ToLocalChecked()));
  OCVariant *ocv = new OCVariant();
  CHECK_OCV(ocv);
  Local<Object> thisObject = info.This();
  V8Variant *v = new V8Variant(); // must catch exception
  v->Wrap(thisObject); // InternalField[0]
  thisObject->SetInternalField(1, ExternalNew(ocv));
  // IDK
  // thisObject.SetWeak(ocv, Dispose);
  DISPFUNCOUT();
 return info.GetReturnValue().Set(thisObject);
}

Handle<Value> V8Variant::OLEFlushCarryOver(Handle<Value> v)
{
  OLETRACEIN();
  OLETRACEVT(v->ToObject());
  Handle<Value> result = Nan::Undefined();
  V8Variant *v8v = ObjectWrap::Unwrap<V8Variant>(v->ToObject());
  if(v8v->property_carryover.empty()){
    std::cerr << " *** carryover empty *** " << __FUNCTION__ << std::endl;
    std::cerr.flush();
    // *** or throw exception
  }else{
    const char *name = v8v->property_carryover.c_str();
    {
      OLETRACEPREARGV(Nan::New(name).ToLocalChecked());
      OLETRACEARGV();
    }
    OLETRACEFLUSH();
    Handle<Value> argv[] = {Nan::New(name).ToLocalChecked(), Nan::New<Array>(0)};
    int argc = sizeof(argv) / sizeof(argv[0]);
    v8v->property_carryover.erase();
    result = INSTANCE_CALL(v->ToObject(), "call", argc, argv);
    OCVariant *rv = V8Variant::CreateOCVariant(result);
    CHECK_OCV(rv);
    OCVariant *o = castedInternalField<OCVariant>(v->ToObject());
    CHECK_OCV(o);
    *o = *rv; // copy and don't delete rv
  }
  OLETRACEOUT();
  return result;
}

template<bool isCall>
NAN_METHOD(OLEInvoke)
{
  OLETRACEIN();
  OLETRACEVT(info.This());
  OLETRACEARGS();
  OLETRACEFLUSH();
  OCVariant *ocv = castedInternalField<OCVariant>(info.This());
  CHECK_OCV(ocv);
  Handle<Value> av0, av1;
  CHECK_OLE_ARGS(info, 1, av0, av1);
  OCVariant *argchain = NULL;
  Array *a = Array::Cast(*av1);
  for(size_t i = 0; i < a->Length(); ++i){
    OCVariant *o = V8Variant::CreateOCVariant(
      ARRAY_AT(a, (i ? i : a->Length()) - 1));
    CHECK_OCV(o);
    if(!i) argchain = o;
    else argchain->push(o);
  }
  Handle<Object> vResult = V8Variant::CreateUndefined();
  String::Utf8Value u8s(av0);
  wchar_t *wcs = u8s2wcs(*u8s);
  if(!wcs && argchain) delete argchain;
  BEVERIFY(done, wcs);
  try{
    OCVariant *rv = isCall ? // argchain will be deleted automatically
      ocv->invoke(wcs, argchain, true) : ocv->getProp(wcs, argchain);
    if(rv){
      OCVariant *o = castedInternalField<OCVariant>(vResult);
      CHECK_OCV(o);
      *o = *rv; // copy and don't delete rv
    }
  }catch(OLE32coreException e){ std::cerr << e.errorMessage(*u8s); goto done;
  }catch(char *e){ std::cerr << e << *u8s << std::endl; goto done;
  }
  free(wcs); // *** it may leak when error ***
  Handle<Value> result = INSTANCE_CALL(vResult, "toValue", 0, NULL);
  OLETRACEOUT();
  return info.GetReturnValue().Set(result);
done:
  OLETRACEOUT();
  Nan::ThrowError(Exception::TypeError(
    Nan::New(__FUNCTION__" failed").ToLocalChecked()));
}

NAN_METHOD(V8Variant::OLECall)
{
  OLETRACEIN();
  OLETRACEVT(info.This());
  OLETRACEARGS();
  OLETRACEFLUSH();
  OLEInvoke<true>(info);
  OLETRACEOUT();
}

NAN_METHOD(V8Variant::OLEGet)
{
  OLETRACEIN();
  OLETRACEVT(info.This());
  OLETRACEARGS();
  OLETRACEFLUSH();
  OLEInvoke<false>(info);
  OLETRACEOUT();
}

NAN_METHOD(V8Variant::OLESet)
{
  OLETRACEIN();
  OLETRACEVT(info.This());
  OLETRACEARGS();
  OLETRACEFLUSH();
  Local<Object> thisObject = info.This();
  OLE_PROCESS_CARRY_OVER(thisObject);
  OCVariant *ocv = castedInternalField<OCVariant>(thisObject);
  CHECK_OCV(ocv);
  Handle<Value> av0, av1;
  CHECK_OLE_ARGS(info, 2, av0, av1);
  OCVariant *argchain = V8Variant::CreateOCVariant(av1);
  if(!argchain)
    return Nan::ThrowError(Exception::TypeError(Nan::New(
      __FUNCTION__" the second argument is not valid (null OCVariant)").ToLocalChecked()));
  bool result = false;
  String::Utf8Value u8s(av0);
  wchar_t *wcs = u8s2wcs(*u8s);
  BEVERIFY(done, wcs);
  try{
    ocv->putProp(wcs, argchain); // argchain will be deleted automatically
  }catch(OLE32coreException e){ std::cerr << e.errorMessage(*u8s); goto done;
  }catch(char *e){ std::cerr << e << *u8s << std::endl; goto done;
  }
  free(wcs); // *** it may leak when error ***
  result = true;
  OLETRACEOUT();
  return info.GetReturnValue().Set(result);
done:
  OLETRACEOUT();
  Nan::ThrowError(Exception::TypeError(
    Nan::New(__FUNCTION__" failed").ToLocalChecked()));
}

NAN_METHOD(V8Variant::OLECallComplete)
{
  OLETRACEIN();
  OLETRACEVT(info.This());
  Handle<Value> result = Nan::Undefined();
  V8Variant *v8v = ObjectWrap::Unwrap<V8Variant>(info.This());
  if(v8v->property_carryover.empty()){
    std::cerr << " *** carryover empty *** " << __FUNCTION__ << std::endl;
    std::cerr.flush();
    // *** or throw exception
  }else{
    const char *name = v8v->property_carryover.c_str();
    {
      OLETRACEPREARGV(Nan::New(name).ToLocalChecked());
      OLETRACEARGV();
    }
    OLETRACEARGS();
    OLETRACEFLUSH();
    Handle<Array> a = Nan::New<Array>(info.Length());
    for(int i = 0; i < info.Length(); ++i) ARRAY_SET(a, i, info[i]);
    Handle<Value> argv[] = {Nan::New(name).ToLocalChecked(), a};
    int argc = sizeof(argv) / sizeof(argv[0]);
    v8v->property_carryover.erase();
    result = INSTANCE_CALL(info.This(), "call", argc, argv);
  }
//_
//Handle<Value> r = INSTANCE_CALL(Handle<Object>::Cast(v), "toValue", 0, NULL);
  OLETRACEOUT();
 return info.GetReturnValue().Set(result);
}

NAN_PROPERTY_GETTER(V8Variant::OLEGetAttr)
{
  OLETRACEIN();
  OLETRACEVT(info.This());
  {
    OLETRACEPREARGV(property);
    OLETRACEARGV();
  }
  OLETRACEFLUSH();
  String::Utf8Value u8name(property);
  Local<Object> thisObject = info.This();

  // Why GetAttr comes twice for () in the third loop instead of CallComplete ?
  // Because of the Crankshaft v8's run-time optimizer ?
  {
    V8Variant *v8v = ObjectWrap::Unwrap<V8Variant>(thisObject);
    if(!v8v->property_carryover.empty()){
      if(v8v->property_carryover == *u8name){
        OLETRACEOUT();
        return info.GetReturnValue().Set(thisObject); // through it
      }
    }
  }

#if(0)
  if(std::string("call") == *u8name || std::string("get") == *u8name
  || std::string("_") == *u8name || std::string("toValue") == *u8name
//|| std::string("valueOf") == *u8name || std::string("toString") == *u8name
  ){
    OLE_PROCESS_CARRY_OVER(thisObject);
  }
#else
  if(std::string("set") != *u8name
  && std::string("toBoolean") != *u8name
  && std::string("toInt32") != *u8name && std::string("toInt64") != *u8name
  && std::string("toNumber") != *u8name && std::string("toDate") != *u8name
  && std::string("toUtf8") != *u8name
  && std::string("inspect") != *u8name && std::string("constructor") != *u8name
  && std::string("valueOf") != *u8name && std::string("toString") != *u8name
  && std::string("toLocaleString") != *u8name
  && std::string("toJSON") != *u8name
  && std::string("hasOwnProperty") != *u8name
  && std::string("isPrototypeOf") != *u8name
  && std::string("propertyIsEnumerable") != *u8name
//&& std::string("_") != *u8name
  ){
    OLE_PROCESS_CARRY_OVER(thisObject);
  }
#endif
  OLETRACEVT(thisObject);
  // Can't use INSTANCE_CALL here. (recursion itself)
  // So it returns Object's fundamental function and custom function:
  //   inspect ?, constructor valueOf toString toLocaleString
  //   hasOwnProperty isPrototypeOf propertyIsEnumerable
  static fundamental_attr fundamentals[] = {
    {0, "call", OLECall}, {0, "get", OLEGet}, {0, "set", OLESet},
    {0, "isA", OLEIsA}, {0, "vtName", OLEVTName}, // {"vt_names", ???},
    {!0, "toBoolean", OLEBoolean},
    {!0, "toInt32", OLEInt32}, {!0, "toInt64", OLEInt64},
    {!0, "toNumber", OLENumber}, {!0, "toDate", OLEDate},
    {!0, "toUtf8", OLEUtf8},
    {0, "toValue", OLEValue},
    {0, "inspect", OLEPrimitiveValue}, {0, "constructor", NULL}, {0, "valueOf", OLEPrimitiveValue},
    {0, "toString", OLEPrimitiveValue}, {0, "toLocaleString", OLEPrimitiveValue},
    {0, "toJSON", OLEPrimitiveValue},
    {0, "hasOwnProperty", NULL}, {0, "isPrototypeOf", NULL},
    {0, "propertyIsEnumerable", NULL}
  };
  for(int i = 0; i < sizeof(fundamentals) / sizeof(fundamentals[0]); ++i){
    if(std::string(fundamentals[i].name) != *u8name) continue;
    if(fundamentals[i].obsoleted){
      std::cerr << " *** ## [." << fundamentals[i].name;
      std::cerr << "()] is obsoleted. ## ***" << std::endl;
      std::cerr.flush();
    }
    OLETRACEFLUSH();
    OLETRACEOUT();
    return info.GetReturnValue().Set(Nan::GetFunction(Nan::New<FunctionTemplate>(
      fundamentals[i].func, thisObject)).ToLocalChecked());
  }
  if(std::string("_") == *u8name){ // through it when "_"
#if(0)
    std::cerr << " *** ## [._] is obsoleted. ## ***" << std::endl;
    std::cerr.flush();
#endif
  }else{
    Handle<Object> vResult = V8Variant::CreateUndefined(); // uses much memory
    OCVariant *rv = V8Variant::CreateOCVariant(thisObject);
    CHECK_OCV(rv);
    OCVariant *o = castedInternalField<OCVariant>(vResult);
    CHECK_OCV(o);
    *o = *rv; // copy and don't delete rv
    V8Variant *v8v = ObjectWrap::Unwrap<V8Variant>(vResult);
    v8v->property_carryover.assign(*u8name);
    OLETRACEPREARGV(property);
    OLETRACEARGV();
    OLETRACEFLUSH();
    OLETRACEOUT();
    return info.GetReturnValue().Set(vResult); // convert and hold it (uses much memory)
  }
  OLETRACEFLUSH();
  OLETRACEOUT();
 return info.GetReturnValue().Set(thisObject); // through it
}

NAN_PROPERTY_SETTER(V8Variant::OLESetAttr)
{
  OLETRACEIN();
  OLETRACEVT(info.This());
  Handle<Value> argv[] = {property, value};
  int argc = sizeof(argv) / sizeof(argv[0]);
  OLETRACEARGV();
  OLETRACEFLUSH();
  Handle<Value> r = INSTANCE_CALL(info.This(), "set", argc, argv);
  OLETRACEOUT();
 return info.GetReturnValue().Set(r);
}

NAN_METHOD(V8Variant::Finalize)
{
  DISPFUNCIN();
#if(0)
  std::cerr << __FUNCTION__ << " Finalizer is called\a" << std::endl;
  std::cerr.flush();
#endif
  Local<Object> thisObject = info.This();
#if(0)
  V8Variant *v = ObjectWrap::Unwrap<V8Variant>(thisObject);
  if(v) delete v; // it has been already deleted ?
  thisObject->SetInternalField(0, External::New(NULL));
#endif
#if(1) // now GC will call Disposer automatically
  OCVariant *ocv = castedInternalField<OCVariant>(thisObject);
  if(ocv) delete ocv;
#endif
  thisObject->SetInternalField(1, ExternalNew(NULL));
  DISPFUNCOUT();
 return info.GetReturnValue().Set(thisObject);
}

template<typename T>
void V8Variant::Dispose(const Nan::WeakCallbackInfo<T> &data)
{
  DISPFUNCIN();
#if(0)
//  std::cerr << __FUNCTION__ << " Disposer is called\a" << std::endl;
  std::cerr << __FUNCTION__ << " Disposer is called" << std::endl;
  std::cerr.flush();
#endif
  Local<Object> thisObject = data.GetValue();
#if(0) // it has been already deleted ?
  V8Variant *v = ObjectWrap::Unwrap<V8Variant>(thisObject);
  if(!v){
    std::cerr << __FUNCTION__;
    std::cerr << "InternalField[0] has been already deleted" << std::endl;
    std::cerr.flush();
  }else delete v; // it has been already deleted ?
  BEVERIFY(done, thisObject->InternalFieldCount() > 0);
  thisObject->SetInternalField(0, ExternalNew(NULL));
#endif
  OCVariant *p = castedInternalField<OCVariant>(thisObject);
  if(!p){
    std::cerr << __FUNCTION__;
    std::cerr << "InternalField[1] has been already deleted" << std::endl;
    std::cerr.flush();
  }
//  else{
    OCVariant *ocv = data.GetParameter(); // ocv may be same as p
    if(ocv) delete ocv;
//  }
  BEVERIFY(done, thisObject->InternalFieldCount() > 1);
  thisObject->SetInternalField(1, ExternalNew(NULL));
done:
  DISPFUNCOUT();
}

void V8Variant::Finalize()
{
  assert(!finalized);
  finalized = true;
}

} // namespace node_win32ole
