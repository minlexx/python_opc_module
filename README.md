# python_opc_module
Python interface for OPC protocol

I needed a way to access OPC Servers from python program, especially Python version 3.
OpenOPC for python works only with python2, pythoncom from pywin32 didn't work for me for some reason :(
So I decided to write a simple small python module myself.
Inside there is a lot of Win32/COM C/C++ code, python interface is much simpler (but not as simple as in OpenOPC).
Currently only read is supported. I didn't need a write sunctionality for now, but it can be easily done.

#How to compile and install:
CMake power! Use appropriate compiler for your python. vc2010 for python3.4, vc2015 for python3.5.
Python must be installed with headers and libs. They are the only dependency...
CMake can even install the module for you python, if you set the correct path (PYTHON_DIR) in cmake configure.

#How to use:
```
import opc_helper

opc_helper.initialize_com()
computer_name = None  # can be Netbios computer name or IP address
opcsrvs = opc_helper.opc_enum_query(computer_name)
for opcsrv in opcsrvs:
  print(opcsrv)

server = opc_helper.opc_connect(opcsrvs[0]['guid'], computer_name)
print(server)
print(server.guid)
print(server.progid)
print('status:', server.get_status())
print('supports v3:', server.supports_v3())
print(server.browse())
print(server.get_item('Numeric._I4'))
print(server.get_item_info('Numeric._I4'))
```

That's all API for now.
