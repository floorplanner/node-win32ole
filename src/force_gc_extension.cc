/*
  force_gc_extension.cc
*/

#include "node_win32ole.h"
#include <node.h>
#include <nan.h>
#include "ole32core.h"

using namespace v8;
using namespace ole32core;

namespace node_win32ole {

NAN_METHOD(Method_force_gc_extension) // v8/gc : gc()
{
  BDISPFUNCDAT("context %s "__FUNCTION__" %s\n", "preset", "start");
  // create context with extension(s)
  const char *extensionNames[] = {"v8/gc",};
  ExtensionConfiguration extensions(
    sizeof(extensionNames) / sizeof(extensionNames[0]), extensionNames);
  Handle<ObjectTemplate> global = ObjectTemplate::New();
  Local<Context> context = NanNew<Context>(&extensions, global);
  Context::Scope context_scope(context);
  BDISPFUNCDAT("context %s "__FUNCTION__" %s\n", "preset", "end");
  NanScope();
  BDISPFUNCIN();
  Local<String> sourceObj = NanNew("gc()");
  TryCatch try_catch;
  Local<Script> scriptObj = Script::Compile(sourceObj);
  if(scriptObj.IsEmpty()){
    std::string msg("Can't compile v8/gc : gc();\n");
    String::Utf8Value u8s(try_catch.Exception());
    msg += *u8s;
    NanThrowError(Exception::TypeError(NanNew(msg.c_str())));
  }
  Local<Value> local_result = scriptObj->Run();
  if(local_result.IsEmpty()){
    std::string msg("Can't get execution result of v8/gc : gc();\n");
    String::Utf8Value u8s(try_catch.Exception());
    msg += *u8s;
    NanThrowError(Exception::TypeError(NanNew(msg.c_str())));
  }
  String::Utf8Value result(local_result);
  BDISPFUNCOUT();
  NanReturnValue(NanNew(*result));
}

} // namespace node_win32ole
