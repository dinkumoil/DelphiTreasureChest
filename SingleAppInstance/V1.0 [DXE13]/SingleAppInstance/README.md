# SingleAppInstance

Written in Delphi 10.4 Sydney.

**Target OS:** MS Windows (Vista and later)

**Author:** Andreas Heim, 2020-06


With the help of these Delphi units it is possible to implement a single-instance-mode for an arbitrary application.

The control logic for this purpose is implemented in unit `_src\lib\class_SingleAppInstance.pas`. The necessary communication between the main application instance and subsequently started instances in order to exchange command line arguments is done via a client-server architecture using named pipes. The implementing units are `_src\lib\class_SingleAppInstance.PipeServer.pas` and `_src\lib\class_SingleAppInstance.PipeClient.pas`.

This server part of this project was inspired by the _Microsoft Learn_ article [_Named Pipe Server Using Completion Routines_](https://learn.microsoft.com/en-us/windows/win32/ipc/named-pipe-server-using-completion-routines). For the client part it was the article [_Named Pipe Client_](https://learn.microsoft.com/en-us/windows/win32/ipc/named-pipe-client).


# History

v1.0 - December 2023
- Initial version
