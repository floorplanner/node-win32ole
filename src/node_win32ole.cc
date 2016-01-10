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

Nan::Persistent<Object> module_target;

NAN_METHOD(Method_version)
{
  Handle<Object> local = Nan::New<Object>(module_target);
  auto ver = Nan::Get(local, Nan::New("VERSION").ToLocalChecked());
  if (!ver.IsEmpty()) {
   return info.GetReturnValue().Set(ver.ToLocalChecked());
  }
}

NAN_METHOD(Method_printACP) // UTF-8 to MBCS (.ACP)
{
  if(info.Length() >= 1){
    String::Utf8Value s(info[0]);
    char *cs = *s;
    printf(UTF82MBCS(std::string(cs)).c_str());
  }
 return info.GetReturnValue().Set(true);
}

NAN_METHOD(Method_print) // through (as ASCII)
{
  if(info.Length() >= 1){
    String::Utf8Value s(info[0]);
    char *cs = *s;
    printf(cs); // printf("%s\n", cs);
  }
 return info.GetReturnValue().Set(true);
}

} // namespace node_win32ole

using namespace node_win32ole;

namespace {

void init(Handle<Object> target)
{
  module_target.Reset(target);
  V8Variant::Init(target);
  Client::Init(target);
  Nan::ForceSet(target, Nan::New("VERSION").ToLocalChecked(),
    Nan::New("0.0.0 (will be set later)").ToLocalChecked(),
    static_cast<PropertyAttribute>(DontDelete));
  Nan::ForceSet(target, Nan::New("MODULEDIRNAME").ToLocalChecked(),
    Nan::New("/tmp").ToLocalChecked(),
    static_cast<PropertyAttribute>(DontDelete));
  Nan::ForceSet(target, Nan::New("SOURCE_TIMESTAMP").ToLocalChecked(),
    Nan::New(__FILE__ " " __DATE__ " " __TIME__).ToLocalChecked(),
    static_cast<PropertyAttribute>(ReadOnly | DontDelete));
  Nan::Set(target, Nan::New("version").ToLocalChecked(),
    Nan::GetFunction(Nan::New<FunctionTemplate>(Method_version)).ToLocalChecked());
  Nan::Set(target, Nan::New("printACP").ToLocalChecked(),
    Nan::GetFunction(Nan::New<FunctionTemplate>(Method_printACP)).ToLocalChecked());
  Nan::Set(target, Nan::New("print").ToLocalChecked(),
    Nan::GetFunction(Nan::New<FunctionTemplate>(Method_print)).ToLocalChecked());
  Nan::Set(target, Nan::New("gettimeofday").ToLocalChecked(),
    Nan::GetFunction(Nan::New<FunctionTemplate>(Method_gettimeofday)).ToLocalChecked());
  Nan::Set(target, Nan::New("sleep").ToLocalChecked(),
    Nan::GetFunction(Nan::New<FunctionTemplate>(Method_sleep)).ToLocalChecked());
  Nan::Set(target, Nan::New("force_gc_extension").ToLocalChecked(),
    Nan::GetFunction(Nan::New<FunctionTemplate>(Method_force_gc_extension)).ToLocalChecked());
  Nan::Set(target, Nan::New("force_gc_internal").ToLocalChecked(),
    Nan::GetFunction(Nan::New<FunctionTemplate>(Method_force_gc_internal)).ToLocalChecked());
}

} // namespace

NODE_MODULE(node_win32ole, init)
