#ifndef __CLIENT_H__
#define __CLIENT_H__

#include <node.h>
#include <nan.h>
#include "node_win32ole.h"

using namespace v8;

namespace node_win32ole {

class Client : public node::ObjectWrap {
public:
  static Nan::Persistent<FunctionTemplate> clazz;
  static void Init(Handle<Object> target);
  static NAN_METHOD(New);
  static NAN_METHOD(Dispatch);
  static NAN_METHOD(Finalize);
public:
  Client() : finalized(false) {}
  ~Client() { if(!finalized) Finalize(); }
protected:
  template<typename T>
  void Dispose(const Nan::WeakCallbackInfo<T> &data);
  void Finalize();
protected:
  bool finalized;
};

} // namespace node_win32ole

#endif // __CLIENT_H__
