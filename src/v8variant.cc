/*
  v8variant.cc
*/

#include "v8variant.h"
#include <node.h>
#include <nan.h>

using namespace v8;
using namespace ole32core;

namespace node_win32ole {

#define CHECK_OLE_ARGS(args, n, av0, av1) do{ \
    if(args.Length() < n) \
      NanThrowError(Exception::TypeError( \
        NanNew<String>(__FUNCTION__" takes exactly " #n " argument(s)"))); \
    if(!args[0]->IsString()) \
      NanThrowError(Exception::TypeError( \
        NanNew<String>(__FUNCTION__" the first argument is not a Symbol"))); \
    if(n == 1) \
      if(args.Length() >= 2) \
        if(!args[1]->IsArray()) \
          NanThrowError(Exception::TypeError(NanNew<String>( \
            __FUNCTION__" the second argument is not an Array"))); \
        else av1 = args[1]; /* Array */ \
      else av1 = NanNew<Array>(0); /* change none to Array[] */ \
    else av1 = args[1]; /* may not be Array */ \
    av0 = args[0]; \
  }while(0)

Persistent<FunctionTemplate> V8Variant::clazz;

void V8Variant::Init(Handle<Object> target)
{
  NanScope();
  Local<FunctionTemplate> t = NanNew<FunctionTemplate>(New);
  t->InstanceTemplate()->SetInternalFieldCount(2);
  t->SetClassName(NanNew<String>("V8Variant"));
  NODE_SET_PROTOTYPE_METHOD(t, "isA", OLEIsA);
  NODE_SET_PROTOTYPE_METHOD(t, "vtName", OLEVTName);
  NODE_SET_PROTOTYPE_METHOD(t, "toBoolean", OLEBoolean);
  NODE_SET_PROTOTYPE_METHOD(t, "toInt32", OLEInt32);
  NODE_SET_PROTOTYPE_METHOD(t, "toInt64", OLEInt64);
  NODE_SET_PROTOTYPE_METHOD(t, "toNumber", OLENumber);
  NODE_SET_PROTOTYPE_METHOD(t, "toDate", OLEDate);
  NODE_SET_PROTOTYPE_METHOD(t, "toUtf8", OLEUtf8);
  NODE_SET_PROTOTYPE_METHOD(t, "toValue", OLEValue);
//  NODE_SET_PROTOTYPE_METHOD(t, "New", New);
  NODE_SET_PROTOTYPE_METHOD(t, "call", OLECall);
  NODE_SET_PROTOTYPE_METHOD(t, "get", OLEGet);
  NODE_SET_PROTOTYPE_METHOD(t, "set", OLESet);
/*
 In ParseUnaryExpression() < v8/src/parser.cc >
 v8::Object::ToBoolean() is called directly for unary operator '!'
 instead of v8::Object::valueOf()
 so NamedPropertyHandler will not be called
 Local<Boolean> ToBoolean(); // How to fake ? override v8::Value::ToBoolean
*/
  Local<ObjectTemplate> instancetpl = t->InstanceTemplate();
  instancetpl->SetCallAsFunctionHandler(OLECallComplete);
  instancetpl->SetNamedPropertyHandler(OLEGetAttr, OLESetAttr);
//  instancetpl->SetIndexedPropertyHandler(OLEGetIdxAttr, OLESetIdxAttr);
  NODE_SET_PROTOTYPE_METHOD(t, "Finalize", Finalize);
  target->Set(NanNew<String>("V8Variant"), t->GetFunction());
  NanAssignPersistent(clazz, t);
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
  BEVERIFY(done, !v->IsFunction());
// VT_USERDEFINED VT_VARIANT VT_BYREF VT_ARRAY more...
  if(v->IsBoolean()){
    return new OCVariant((bool)(v->BooleanValue() ? !0 : 0));
  }else if(v->IsArray()){
// VT_BYREF VT_ARRAY VT_SAFEARRAY
    std::cerr << "[Array (not implemented now)]" << std::endl; return NULL;
    std::cerr.flush();
  }else if(v->IsInt32()){
    return new OCVariant((long)v->Int32Value());
#if(0) // may not be supported node.js / v8
  }else if(v->IsInt64()){
    return new OCVariant((long long)v->Int64Value());
#endif
  }else if(v->IsNumber()){
    std::cerr << "[Number (VT_R8 or VT_I8 bug?)]" << std::endl;
    std::cerr.flush();
// if(v->ToInteger()) =64 is failed ? double : OCVariant((longlong)VT_I8)
    return new OCVariant((double)v->NumberValue()); // double
  }else if(v->IsNumberObject()){
    std::cerr << "[NumberObject (VT_R8 or VT_I8 bug?)]" << std::endl;
    std::cerr.flush();
// if(v->ToInteger()) =64 is failed ? double : OCVariant((longlong)VT_I8)
    return new OCVariant((double)v->NumberValue()); // double
  }else if(v->IsDate()){
    double d = v->NumberValue();
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
    return new OCVariant(CreateStdStringMBCSfromUTF8(v->ToDetailString()));
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
  NanScope();
  DISPFUNCIN();
  OCVariant *ocv = castedInternalField<OCVariant>(args.This());
  CHECK_OCV(ocv);
  DISPFUNCOUT();
  NanReturnValue(NanNew(ocv->v.vt));
}

NAN_METHOD(V8Variant::OLEVTName)
{
  NanScope();
  DISPFUNCIN();
  OCVariant *ocv = castedInternalField<OCVariant>(args.This());
  CHECK_OCV(ocv);
  Local<Object> target = NanNew(module_target);
  Array *a = Array::Cast(*GET_PROP(target, "vt_names"));
  DISPFUNCOUT();
  NanReturnValue(ARRAY_AT(a, ocv->v.vt));
}

NAN_METHOD(V8Variant::OLEBoolean)
{
  NanScope();
  DISPFUNCIN();
  OCVariant *ocv = castedInternalField<OCVariant>(args.This());
  CHECK_OCV(ocv);
  if(ocv->v.vt != VT_BOOL)
    NanThrowError(Exception::TypeError(
      NanNew<String>("OLEBoolean source type OCVariant is not VT_BOOL")));
  bool c_boolVal = ocv->v.boolVal == VARIANT_FALSE ? 0 : !0;
  DISPFUNCOUT();
  NanReturnValue(c_boolVal);
}

NAN_METHOD(V8Variant::OLEInt32)
{
  NanScope();
  DISPFUNCIN();
  OCVariant *ocv = castedInternalField<OCVariant>(args.This());
  CHECK_OCV(ocv);
  if(ocv->v.vt != VT_I4 && ocv->v.vt != VT_INT
  && ocv->v.vt != VT_UI4 && ocv->v.vt != VT_UINT)
    NanThrowError(Exception::TypeError(
      NanNew<String>("OLEInt32 source type OCVariant is not VT_I4 nor VT_INT nor VT_UI4 nor VT_UINT")));
  DISPFUNCOUT();
  NanReturnValue(NanNew(ocv->v.lVal));
}

NAN_METHOD(V8Variant::OLEInt64)
{
  NanScope();
  DISPFUNCIN();
  OCVariant *ocv = castedInternalField<OCVariant>(args.This());
  CHECK_OCV(ocv);
  if(ocv->v.vt != VT_I8 && ocv->v.vt != VT_UI8)
    NanThrowError(Exception::TypeError(
      NanNew<String>("OLEInt64 source type OCVariant is not VT_I8 nor VT_UI8")));
  DISPFUNCOUT();
  NanReturnValue(NanNew<Number>(ocv->v.llVal));
}

NAN_METHOD(V8Variant::OLENumber)
{
  NanScope();
  DISPFUNCIN();
  OCVariant *ocv = castedInternalField<OCVariant>(args.This());
  CHECK_OCV(ocv);
  if(ocv->v.vt != VT_R8)
    NanThrowError(Exception::TypeError(
      NanNew<String>("OLENumber source type OCVariant is not VT_R8")));
  DISPFUNCOUT();
  NanReturnValue(NanNew(ocv->v.dblVal));
}

NAN_METHOD(V8Variant::OLEDate)
{
  NanScope();
  DISPFUNCIN();
  OCVariant *ocv = castedInternalField<OCVariant>(args.This());
  CHECK_OCV(ocv);
  if(ocv->v.vt != VT_DATE)
    NanThrowError(Exception::TypeError(
      NanNew<String>("OLEDate source type OCVariant is not VT_DATE")));
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
  NanReturnValue(NanNew<Date>(mktime(&t) * 1000LL + syst.wMilliseconds));
}

NAN_METHOD(V8Variant::OLEUtf8)
{
  NanScope();
  DISPFUNCIN();
  OCVariant *ocv = castedInternalField<OCVariant>(args.This());
  CHECK_OCV(ocv);
  if(ocv->v.vt != VT_BSTR)
    NanThrowError(Exception::TypeError(
      NanNew<String>("OLEUtf8 source type OCVariant is not VT_BSTR")));
  Handle<Value> result;
  if(!ocv->v.bstrVal) result = NanUndefined(); // or Null();
  else {
    std::wstring wstr(ocv->v.bstrVal);
    char *cs = wcs2u8s(wstr.c_str());
    result = NanNew<String>(cs);
    free(cs);
  }
  DISPFUNCOUT();
  NanReturnValue(result);
}

NAN_METHOD(V8Variant::OLEValue)
{
  NanEscapableScope();
  OLETRACEIN();
  OLETRACEVT(args.This());
  OLETRACEFLUSH();
  Local<Object> thisObject = args.This();
  OLE_PROCESS_CARRY_OVER(thisObject);
  OLETRACEVT(thisObject);
  OLETRACEFLUSH();
  OCVariant *ocv = castedInternalField<OCVariant>(thisObject);
  if(!ocv){ std::cerr << "ocv is null"; std::cerr.flush(); }
  CHECK_OCV(ocv);
  if(ocv->v.vt == VT_EMPTY) ; // do nothing
  else if (ocv->v.vt == VT_NULL) NanReturnNull();
  else if (ocv->v.vt == VT_DISPATCH) NanReturnValue(thisObject); // through it
  else if(ocv->v.vt == VT_BOOL) OLEBoolean(args);
  else if(ocv->v.vt == VT_I4 || ocv->v.vt == VT_INT
  || ocv->v.vt == VT_UI4 || ocv->v.vt == VT_UINT) OLEInt32(args);
  else if(ocv->v.vt == VT_I8 || ocv->v.vt == VT_UI8) OLEInt64(args);
  else if(ocv->v.vt == VT_R8) OLENumber(args);
  else if(ocv->v.vt == VT_DATE) OLEDate(args);
  else if(ocv->v.vt == VT_BSTR) OLEUtf8(args);
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
}

/**
 * Dump all available variables/methods of an OLE object.
 **/
NAN_METHOD(V8Variant::OLEInspect) {
  Local<Object> thisObject = args.This();
  OLE_PROCESS_CARRY_OVER(thisObject);
  OLETRACEVT(thisObject);
  OLETRACEFLUSH();
  OCVariant *ocv = castedInternalField<OCVariant>(thisObject);
  CHECK_OCV(ocv);
  if (ocv->v.vt == VT_DISPATCH) {
    IDispatch *dispatch = ocv->v.pdispVal;
    if (dispatch == NULL) {
      NanReturnNull();
    }
    Local<Object> object = NanNew<Object>();
    ITypeInfo *typeinfo = NULL;
    HRESULT hr = dispatch->GetTypeInfo(0, LOCALE_USER_DEFAULT, &typeinfo);
    if (typeinfo) {
      TYPEATTR* typeattr;
      BASSERT(SUCCEEDED(typeinfo->GetTypeAttr(&typeattr)));
      for (int i = 0; i < typeattr->cFuncs; i++) {
        FUNCDESC *funcdesc;
        typeinfo->GetFuncDesc(i, &funcdesc);
        if (funcdesc->invkind != INVOKE_FUNC) {
          object->Set(NanNew(GetName(typeinfo, funcdesc->memid)), NanNew("Function"));
        }
        typeinfo->ReleaseFuncDesc(funcdesc);
      }
      for (int i = 0; i < typeattr->cVars; i++) {
        VARDESC *vardesc;
        typeinfo->GetVarDesc(i, &vardesc);
        object->Set(NanNew(GetName(typeinfo, vardesc->memid)), NanNew("Variable"));
        typeinfo->ReleaseVarDesc(vardesc);
      }
      typeinfo->ReleaseTypeAttr(typeattr);
    }
    NanReturnValue(object);
  } else {
    V8Variant::OLEValue(args);
  }
}

Handle<Object> V8Variant::CreateUndefined(void)
{
  DISPFUNCIN();
  Local<FunctionTemplate> localClazz = NanNew(clazz);
  Local<Object> instance = localClazz->GetFunction()->NewInstance(0, NULL);
  DISPFUNCOUT();
  return instance;
}

NAN_METHOD(V8Variant::New)
{
  NanScope();
  DISPFUNCIN();
  if(!args.IsConstructCall())
    NanThrowError(Exception::TypeError(
      NanNew<String>("Use the new operator to create new V8Variant objects")));
  OCVariant *ocv = new OCVariant();
  CHECK_OCV(ocv);
  Local<Object> thisObject = args.This();
  V8Variant *v = new V8Variant(); // must catch exception
  v->Wrap(thisObject); // InternalField[0]
  thisObject->SetInternalField(1, ExternalNew(ocv));
  NanMakeWeakPersistent(thisObject, ocv, Dispose);
  DISPFUNCOUT();
  NanReturnThis();
}

Handle<Value> V8Variant::OLEFlushCarryOver(Handle<Value> v)
{
  OLETRACEIN();
  OLETRACEVT(v->ToObject());
  Handle<Value> result = NanUndefined();
  V8Variant *v8v = ObjectWrap::Unwrap<V8Variant>(v->ToObject());
  if(v8v->property_carryover.empty()){
    std::cerr << " *** carryover empty *** " << __FUNCTION__ << std::endl;
    std::cerr.flush();
    // *** or throw exception
  }else{
    const char *name = v8v->property_carryover.c_str();
    {
      OLETRACEPREARGV(NanNew<String>(name));
      OLETRACEARGV();
    }
    OLETRACEFLUSH();
    Handle<Value> argv[] = {NanNew<String>(name), Array::New(v8::Isolate::GetCurrent(), 0)};
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
  NanScope();
  OLETRACEIN();
  OLETRACEVT(args.This());
  OLETRACEARGS();
  OLETRACEFLUSH();
  OCVariant *ocv = castedInternalField<OCVariant>(args.This());
  CHECK_OCV(ocv);
  Handle<Value> av0, av1;
  CHECK_OLE_ARGS(args, 1, av0, av1);
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
  NanReturnValue(result);
done:
  OLETRACEOUT();
  NanThrowError(Exception::TypeError(
    NanNew<String>(__FUNCTION__" failed")));
}

NAN_METHOD(V8Variant::OLECall)
{
  NanScope();
  OLETRACEIN();
  OLETRACEVT(args.This());
  OLETRACEARGS();
  OLETRACEFLUSH();
  OLEInvoke<true>(args);
  OLETRACEOUT();
}

NAN_METHOD(V8Variant::OLEGet)
{
  NanScope();
  OLETRACEIN();
  OLETRACEVT(args.This());
  OLETRACEARGS();
  OLETRACEFLUSH();
  OLEInvoke<false>(args);
  OLETRACEOUT();
}

NAN_METHOD(V8Variant::OLESet)
{
  NanScope();
  OLETRACEIN();
  OLETRACEVT(args.This());
  OLETRACEARGS();
  OLETRACEFLUSH();
  Local<Object> thisObject = args.This();
  OLE_PROCESS_CARRY_OVER(thisObject);
  OCVariant *ocv = castedInternalField<OCVariant>(thisObject);
  CHECK_OCV(ocv);
  Handle<Value> av0, av1;
  CHECK_OLE_ARGS(args, 2, av0, av1);
  OCVariant *argchain = V8Variant::CreateOCVariant(av1);
  if(!argchain)
    NanThrowError(Exception::TypeError(NanNew<String>(
      __FUNCTION__" the second argument is not valid (null OCVariant)")));
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
  NanReturnValue(result);
done:
  OLETRACEOUT();
  NanThrowError(Exception::TypeError(
    NanNew<String>(__FUNCTION__" failed")));
}

NAN_METHOD(V8Variant::OLECallComplete)
{
  NanScope();
  OLETRACEIN();
  OLETRACEVT(args.This());
  Handle<Value> result = NanUndefined();
  V8Variant *v8v = ObjectWrap::Unwrap<V8Variant>(args.This());
  if(v8v->property_carryover.empty()){
    std::cerr << " *** carryover empty *** " << __FUNCTION__ << std::endl;
    std::cerr.flush();
    // *** or throw exception
  }else{
    const char *name = v8v->property_carryover.c_str();
    {
      OLETRACEPREARGV(NanNew<String>(name));
      OLETRACEARGV();
    }
    OLETRACEARGS();
    OLETRACEFLUSH();
    Handle<Array> a = NanNew<Array>(args.Length());
    for(int i = 0; i < args.Length(); ++i) ARRAY_SET(a, i, args[i]);
    Handle<Value> argv[] = {NanNew<String>(name), a};
    int argc = sizeof(argv) / sizeof(argv[0]);
    v8v->property_carryover.erase();
    result = INSTANCE_CALL(args.This(), "call", argc, argv);
  }
//_
//Handle<Value> r = INSTANCE_CALL(Handle<Object>::Cast(v), "toValue", 0, NULL);
  OLETRACEOUT();
  NanReturnValue(result);
}

NAN_PROPERTY_GETTER(V8Variant::OLEGetAttr)
{
  NanScope();
  OLETRACEIN();
  OLETRACEVT(args.This());
  {
    OLETRACEPREARGV(property);
    OLETRACEARGV();
  }
  OLETRACEFLUSH();
  String::Utf8Value u8name(property);
  Local<Object> thisObject = args.This();

  // Why GetAttr comes twice for () in the third loop instead of CallComplete ?
  // Because of the Crankshaft v8's run-time optimizer ?
  {
    V8Variant *v8v = ObjectWrap::Unwrap<V8Variant>(thisObject);
    if(!v8v->property_carryover.empty()){
      if(v8v->property_carryover == *u8name){
        OLETRACEOUT();
        NanReturnValue(thisObject); // through it
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
    {!0, "toBoolean", OLEValue},
    {!0, "toInt32", OLEValue}, {!0, "toInt64", OLEValue},
    {!0, "toNumber", OLEValue}, {!0, "toDate", OLEValue},
    {!0, "toUtf8", OLEValue},
    {0, "toValue", OLEValue},
    {0, "inspect", OLEInspect}, {0, "constructor", NULL}, {0, "valueOf", OLEValue},
    {0, "toString", OLEValue}, {0, "toLocaleString", OLEValue},
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
    NanReturnValue(NanNew<FunctionTemplate>(
      fundamentals[i].func, thisObject)->GetFunction());
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
    NanReturnValue(vResult); // convert and hold it (uses much memory)
  }
  OLETRACEFLUSH();
  OLETRACEOUT();
  NanReturnValue(thisObject); // through it
}

NAN_PROPERTY_SETTER(V8Variant::OLESetAttr)
{
  NanScope();
  OLETRACEIN();
  OLETRACEVT(args.This());
  Handle<Value> argv[] = {property, value};
  int argc = sizeof(argv) / sizeof(argv[0]);
  OLETRACEARGV();
  OLETRACEFLUSH();
  Handle<Value> r = INSTANCE_CALL(args.This(), "set", argc, argv);
  OLETRACEOUT();
  NanReturnValue(r);
}

NAN_METHOD(V8Variant::Finalize)
{
  NanScope();
  DISPFUNCIN();
#if(0)
  std::cerr << __FUNCTION__ << " Finalizer is called\a" << std::endl;
  std::cerr.flush();
#endif
  Local<Object> thisObject = args.This();
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
  NanReturnThis();
}

NAN_WEAK_CALLBACK(V8Variant::Dispose)
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
