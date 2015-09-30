/*
  node_win32ole.cc
  $ cp ~/.node-gyp/0.8.18/ia32/node.lib to ~/.node-gyp/0.8.18/node.lib
  $ node-gyp configure
  $ node-gyp build
  $ node test/init_win32ole.test.js
*/

#include "node_win32ole.h"
#include "client.h"
#include "v8variant.h"

using namespace v8;

namespace node_win32ole {

Persistent<Object> module_target;

NAN_METHOD(Method_version)
{
  NanScope();
  Handle<Object> local = NanNew<Object>(module_target);
  NanReturnValue(local->Get(NanNew("VERSION")));
}

NAN_METHOD(Method_printACP) // UTF-8 to MBCS (.ACP)
{
  NanScope();
  if(args.Length() >= 1){
    String::Utf8Value s(args[0]);
    char *cs = *s;
    printf(UTF82MBCS(std::string(cs)).c_str());
  }
  NanReturnValue(true);
}

NAN_METHOD(Method_print) // through (as ASCII)
{
  NanScope();
  if(args.Length() >= 1){
    String::Utf8Value s(args[0]);
    char *cs = *s;
    printf(cs); // printf("%s\n", cs);
  }
  NanReturnValue(true);
}

} // namespace node_win32ole

using namespace node_win32ole;

namespace {

void init(Handle<Object> target)
{
  NanAssignPersistent(module_target, target);
  V8Variant::Init(target);
  Client::Init(target);
  target->ForceSet(NanNew("VERSION"),
    NanNew("0.0.0 (will be set later)"),
    static_cast<PropertyAttribute>(DontDelete));
  target->ForceSet(NanNew("MODULEDIRNAME"),
    NanNew("/tmp"),
    static_cast<PropertyAttribute>(DontDelete));
  target->ForceSet(NanNew("SOURCE_TIMESTAMP"),
    NanNew(__FILE__ " " __DATE__ " " __TIME__),
    static_cast<PropertyAttribute>(ReadOnly | DontDelete));
  target->Set(NanNew("version"),
    NanNew<FunctionTemplate>(Method_version)->GetFunction());
  target->Set(NanNew("printACP"),
    NanNew<FunctionTemplate>(Method_printACP)->GetFunction());
  target->Set(NanNew("print"),
    NanNew<FunctionTemplate>(Method_print)->GetFunction());
  target->Set(NanNew("gettimeofday"),
    NanNew<FunctionTemplate>(Method_gettimeofday)->GetFunction());
  target->Set(NanNew("sleep"),
    NanNew<FunctionTemplate>(Method_sleep)->GetFunction());
  target->Set(NanNew("force_gc_extension"),
    NanNew<FunctionTemplate>(Method_force_gc_extension)->GetFunction());
  target->Set(NanNew("force_gc_internal"),
    NanNew<FunctionTemplate>(Method_force_gc_internal)->GetFunction());
}

} // namespace

NODE_MODULE(node_win32ole, init)
