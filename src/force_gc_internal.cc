/*
  force_gc_internal.cc
*/

#include "node_win32ole.h"
#include <iostream>
#include <node.h>
#include <nan.h>

using namespace v8;

namespace node_win32ole {

NAN_METHOD(Method_force_gc_internal)
{
  NanScope();
  std::cerr << "-in: " __FUNCTION__ << std::endl;
  if(args.Length() < 1)
    NanThrowError(Exception::TypeError(
      NanNew("this function takes at least 1 argument(s)")));
  if(!args[0]->IsInt32())
    NanThrowError(Exception::TypeError(
      NanNew("type of argument 1 must be Int32")));
  int flags = (int)args[0]->Int32Value();
  while (!NanIdleNotification(100)){}
  std::cerr << "-out: " __FUNCTION__ << std::endl;
  NanReturnValue(true);
}

} // namespace node_win32ole
