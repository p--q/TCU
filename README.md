# TCU - Tree Command for UNO

Output the API tee from the UNO object or IDL name.

## System requirements

- The confirmed environment is as follows.

  - LibreOffice 5.2, 5.3, 5.4 in Ubuntu 14.04 32bit

  - LibreOffice 5.4 in Windows 10 Home 64bit

In description.xml, LibreOffice-minimal-version is 5.2.

## Installation

- Add <a href="https://github.com/p--q/TCU/tree/master/TCU/oxt">TCU.oxt</a> with Extension Manager.

- Restart LibreOffice.

## Usage

	# Instantiate with IDL name "pq.Tcu".
	tcu = smgr.createInstanceWithContext("pq.Tcu", ctx)  # smgr: service manager, ctx: component context
	
	# wtree() method outputs trees to the default browser.
	# When outputting to the web browser, anchors are attached to the IDL reference.
	tcu.wtree(arg)  # arg is an UNO object or a string of IDL full name.
	
	# tree() method returns a list of lines.
	s = tcu.tree(arg)    # arg is a UNO object or a string of IDL full name.
	print("\n".join(ｓ))　　# By joining the elements of the list with a line feed code we get the following tree.

	# When component context is used as an argument.
	# "com.sun.star" is omitted.
	object
	├─.container.XNameContainer
	│   │   void  insertByName( [in] string aName,
	│   │                       [in]    any aElement
	│   │            ) raises ( .lang.WrappedTargetException,
	│   │                       .container.ElementExistException,
	│   │                       .lang.IllegalArgumentException)
	│   │   void  removeByName( [in] string Name
	│   │            ) raises ( .lang.WrappedTargetException,
	│   │                       .container.NoSuchElementException)
	│   └─.container.XNameReplace
	│   	  │   void  replaceByName( [in] string aName,
	│   	  │                        [in]    any aElement
	│   	  │             ) raises ( .lang.WrappedTargetException,
	│   	  │                        .container.NoSuchElementException,
	│   	  │                        .lang.IllegalArgumentException)
	│   	  └─.container.XNameAccess
	│   	  	  │        any  getByName( [in] string aName
	│   	  	  │             ) raises ( .lang.WrappedTargetException,
	│   	  	  │                        .container.NoSuchElementException)
	│   	  	  │   [string]  getElementNames()
	│   	  	  │    boolean  hasByName( [in] string aName)
	│   	  	  └─.container.XElementAccess
	│   	  	  	  │      type  getElementType()
	│   	  	  	  │   boolean  hasElements()
	│   	  	  	  └─.uno.XInterface
	│   	  	  	  	  	  void  acquire()
	│   	  	  	  	  	   any  queryInterface( [in] type aType)
	│   	  	  	  	  	  void  release()
	├─.lang.XComponent
	│   	  void  addEventListener( [in] .lang.XEventListener xListener)
	│   	  void  dispose()
	│   	  void  removeEventListener( [in] .lang.XEventListener aListener)
	├─.lang.XTypeProvider
	│   	  [byte]  getImplementationId()
	│   	  [type]  getTypes()
	├─.uno.XComponentContext
	│   	  .lang.XMultiComponentFactory  getServiceManager()
	│   	                           any  getValueByName( [in] string Name)
	└─.uno.XWeak
			  .uno.XAdapter  queryAdapter()

## Script Examples

- For automation

  - <a href="https://github.com/p--q/TCU/blob/master/TCU/automationexample.py">TCU/automationexample.py</a>

    - Output the API tree of the component context.

- For macro (It also contains the code for automation.)

  - <a href="https://github.com/p--q/TCU/blob/master/TCU/macroexample.py">TCU/macroexample.py</a>

    - Output the API tree of the document that launched the macro.

## Options

![2017-09-24_002602](https://user-images.githubusercontent.com/6964955/30774573-a9179286-a0bf-11e7-907f-2131c148ceae.png)

- **API Reference URL** : Specify the address to <a href="https://api.libreoffice.org/docs/idl/ref/">LibreOffice: Main Page</a>.

- **Use Local Reference** : If checked, use anchors to the API reference to locally installed SDK.

  - You can not check it unless you install SDK.

- **Ignored Interfaces** : Enter the interface name that you do not want to output.

  - By default, it suppresses the output of the <a href="https://wiki.openoffice.org/wiki/Documentation/DevGuide/WritingUNO/Core_Interfaces_to_Implement">core interfaces</a> that comes up frequently.

- **Restore Defaults** : Get the path to the API reference of locally installed SDK.

  - It is useful when you want to switch the version of LibreOffice of API to use.

## Release notes

2017-9-23 version 0.9.0 First release.


## Tools

<a href="https://github.com/p--q/TCU/tree/master/TCU/tools">TCU/TCU/tools</a> contains helper scripts for creating this extension.

  - createIDLs.py

  - createRDB.py
  
  - createXcs.py
  
  - createOptionsDialogXcu.py
  
  - createXMLs.py
  
  - createOXT.py
  
  - deployOXT.py
  
  - execAtOnce.py
  
  - helper.py
  
  - config.py
  
  - createOXT.py
  
  
