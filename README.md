# CMDBackDoor
This program for executing `DOS commands` remotely on Windows.

# About Program

This is written in `Vb6 without any OCX`.
tested on `Win Server 2019 and Win10` and It works properly.
It has two parts, client and server.


  
# How to use
## 1. Server:
server must be running on the system from which we want to control the victim system .
All you have to do on the server is enter the free port number(not used by Other Program) and click on listen button .

## 2. Client:
This file must be executed on the victim system where we want execute the DOS commands .

After the first run, it asks you for parameters :
Name : <Name show in Server File>
ServerIp: <The system Ip in which we run the server >
ServerPort: <The port number we entered on the server >
After this, the program closes and you have to run the client again. But this time it does not ask for the parameters and Running Hiden. 

## Result
```bash
	After executing both programs and entered the correct parameters,
	the client-server connection is established and you can execute
	the commands on the client and see the result on the server by 
	writing and send DOS commands on the server. 
```
<h3 align="left">Connect with me:</h3>
	
TelegramID : [@mrbugbounty](https://t.me/mrbugbounty)
