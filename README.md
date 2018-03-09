# TCU - Tree Command for UNO

Output the API tee from the UNO object or IDL name.

	├─pq.Tcu
	│   └─pq.XTcu
	│   	  	  [string]  treelines( [in] any Object)
	│   	  	      void  wcompare( [in] any Object1,
	│   	  	                      [in] any Object2)
	│   	  	  [string]  wcomparelines( [in] any Object1,
	│   	  	                           [in] any Object2)
	│   	  	      void  wtree( [in] any Object)
	│   	  	  [string]  wtreelines( [in] any Object)
	└─.awt.XContainerWindowEventHandler
			   boolean  callHandlerMethod( [in] .awt.XWindow xWindow,
						       [in]          any EventObject,
						       [in]       string MethodName
					    ) raises ( .lang.WrappedTargetException)
			  [string]  getSupportedMethodNames()



## System requirements

- The confirmed environment is as follows.

  - LibreOffice 5.4 in Ubuntu 14.04 32bit

  - LibreOffice 5.4 in Windows 10 Home 64bit

In description.xml, LibreOffice-minimal-version is 5.2.

## Installation

- Click the Download button to download <a href="https://github.com/p--q/TCU/blob/master/TCU/oxt/TCU.oxt">TCU.oxt</a>. 

- Add TCU.oxt with Extension Manager.

- Restart LibreOffice.

## Usage

	# Instantiate with IDL name "pq.Tcu".
	# wtree() method outputs trees to the default browser.
	# When outputting to the web browser, anchors are attached to the IDL reference.
	# wtree() method can be executed only once.
	# arg is an UNO object or a string of IDL full name.
	# tree() method returns a tuple of lines.
	# arg is a UNO object or a string of IDL full name.
	# By joining the elements of the tuple with a line feed code we get the following tree.
	# When component context is used as an argument.
	# "com.sun.star" is omitted.
	# The interface outputted once does not displayed.
	
	tcu = smgr.createInstanceWithContext("pq.Tcu", ctx)  
	print("\n".join(tcu.treelines(ctx)))
	
	Connected to a running office ...
	Using LibreOffice 5.4
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
	│   	  	  	  	     type  getElementType()
	│   	  	  	  	  boolean  hasElements()
	├─.lang.XComponent
	│   	  void  addEventListener( [in] .lang.XEventListener xListener)
	│   	  void  dispose()
	│   	  void  removeEventListener( [in] .lang.XEventListener aListener)
	└─.uno.XComponentContext
			  .lang.XMultiComponentFactory  getServiceManager()
						   any  getValueByName( [in] string Name)

	# wcompare() method compares the services and interfaces of the two objects and outputs the results to the web browser.
	# It compares interfaces acquired by getTypes() method.


## Example of Output of wtree() method

	def macro():
		ctx = XSCRIPTCONTEXT.getComponentContext() 
		smgr = ctx.getServiceManager()
		tcu = smgr.createInstanceWithContext("pq.Tcu", ctx)
		tcu.wtree(ctx)

![2018-03-02_195200](https://user-images.githubusercontent.com/6964955/36895712-3c803c38-1e53-11e8-9e85-a7f2f3dc865d.png)

## Script Examples

- For automation

  - <a href="https://github.com/p--q/TCU/blob/master/TCU/usageexample/automationexample.py">TCU/automationexample.py at master · p--q/TCU · GitHub</a>

    - Output the API tree of the component context.

- For macro (It also contains the code for automation.)

  - <a href="https://github.com/p--q/TCU/blob/master/TCU/usageexample/macroexample.py">TCU/macroexample.py at master · p--q/TCU · GitHub</a>

    - Output the API tree of the document that launched the macro.

- wtreelines() method

  - <a href="https://github.com/p--q/TCU/blob/master/TCU/usageexample/tcuclaclines.py">TCU/tcuclaclines.py at master · p--q/TCU · GitHub</a>

## Links to pages with output of TCU

<a href="https://p--q.blogspot.jp/2017/11/libreoffice594.html">LibreOffice5(94)サービスとインターフェース一覧が載っているページの一覧</a>

## Options

![2018-03-02_194047](https://user-images.githubusercontent.com/6964955/36895375-0827341a-1e52-11e8-87ba-f72d2f98ba11.png)

- **API Reference URL** : Specify the address to <a href="https://api.libreoffice.org/docs/idl/ref/">LibreOffice: Main Page</a>.

- **Use Local Reference** : If checked, use anchors to the API reference to locally installed SDK.

  - You can not check it unless you install SDK.

- **Ignored Interfaces** : Enter the interface name that you do not want to output.

  - By default, it suppresses the output of the <a href="https://wiki.openoffice.org/wiki/Documentation/DevGuide/WritingUNO/Core_Interfaces_to_Implement">core interfaces</a> that come up frequently.

- **Restore Defaults** : Get the path to the API reference of locally installed SDK.

  - It is useful when you want to switch the version of LibreOffice of API to use.

## Known Issues

- wtree() and wcompare() methods can be executed only once per process.

## Release notes

2017-9-23 version 0.9.0 First release.

2017-9-24 version 0.9.2 Fixed a serious bug. The output of the services were missing.

2017-9-24 version 0.9.3 Fixed a bug. Default Ignored Interfaces were not used.

2017-10-1 version 0.9.4 Suppress duplicate service output.

2017-10-2 version 0.9.5 Support for services that can not get TypeDescription object.

2017-10-3 version 0.9.6 Refactoring.

2017-10-13 version 1.0.1 Added wcompare() method.

2017-11-10 version 1.0.2 Fixed a bug.　Output the tree without error when the service that can not acquire a TypeDescription object has no getPropertySetInfo() method.

2018-1-26 version 1.0.3 Fixed a bug. Resolved an issue that wcompare() does not suppress output of services not in IDL.

2018-2-2 version 1.0.5 Fixed a serious bug. Output property not via service.

2018-2-3 version 1.0.6 Added function to correct incorrect IDL(com.sun.star.AccessibleSpreadsheetDocumentView). 

2018-2-8 version 1.0.7 Changed branch of property not via service.

2018-3-2 version 2.0.0 Changed algorithm. Fixed an incorrect output of wcompare() method. Changed default suppression interfaces. Python regular expressions enabled.

2018-3-2 version 2.0.1 Stop comparing by property names in wcompare() method.

2018-3-3 version 2.0.2 The property name and the attribute name are prevented from overlapping. Fixed a serious bug. 

2018-3-5 version 2.1.0 Comparing by property names in wcompare() method.

2018-3-6 version 2.1.1 Fixed a bug.

2018-3-7 version 2.1.2 Deal with when the return value of the object's getPropertySetInfo() method is None.

2018-3-9 version 3.0.1 Add treelines(), wtreelines(), wcomparelines() methods. Eliminate 404 error to API reference.

## Tools

This repository is Eclipse's PyDev project.

For PyDev project interpreter, specify LibreOffice bundle Python.

<a href="https://github.com/p--q/TCU/tree/master/TCU/tools">TCU/TCU/tools</a> contains helper scripts for creating this extension.

  - config.py
  
    - This script gets the necessary information for configuring the extension from <a href="https://github.com/p--q/TCU/blob/master/TCU/src/config.ini">config.ini</a> and <a href="https://github.com/p--q/TCU/blob/master/TCU/src/pyunocomponent.py">pyunocomponent.py</a>.
    
  - createIDLs.py

    - Create idl files.
    
    - The content of the idl file is defined in the function defineIDLs() of this script.

  - createRDB.py
  
    - Compile the idl files and create a rdb file.
  
  - createXcs.py
  
    - Create config.xcs file that defines the component schema node that stores the configurations for the extension.
    
    - The definition of the xcs file is defined in the function createXcs() of this script.
  
  - createOptionsDialogXcu.py
  
    - Create OptionsDialog.xcs file to display the options page.
    
    - The definition of the xcs file is defined in the function createOptionsDialogXcu() of this script.
  
  - createXMLs.py
  
    - Create manifest.xml, description.xml, .components files.
    
    - Since these files are created from the existing information, the definition in this script is unnecessary.
  
    - It must be executed when newly creating components, rdb, xcu, xcs files.
    
    - The xml files output by this script are not formatted.
    
    - By default, createComponentsFile() and createManifestFile() are commented out. Although there are few opportunities to update .component file and manifest.xml created by these, description.xml needs to be created each time a extension version is updated.
  
  - createOXT.py
  
    - Collect the files you have created and create an oxt file in the oxt folder.
  
  - deployOXT.py
  
    -  Install this extension to LibreOffice specified as the interpreter.
    
    -  You must execute it with LibreOffice terminated.
    
    -  If it failed, you have to delete the extension in the GUI.
  
  - execAtOnce.py
  
    - Execute multiple scripts in order.
  
  - helper.py
  
    - Class for creating xml nodes.
  
  - createProtocolHandlerXcu.py
  
    -  This script is not used with this extension.

After creating the necessary file, execute createOXT.py and deployOXT.py in execAtOnce.py every time I edit the py file and check the operation.

## How to start the debugger from inside the extension

<a href="https://github.com/p--q/TCU/blob/master/TCU/src/pythonpath/inoxt/common.py">pythonpath/inoxt/common.py</a> contains enableRemoteDebugging decorator to invoke debugger.

1. Write the path to the folder containing pydevd.py on sites.pth.

2. Place sites.pth on the valid path of bundle Python.

    - When the bundle Python version is 3.5, the LibreOffice version is 5.4, Ubuntu 14.04, it becomes as follows.
    
    
**~/.local/lib/python3.5/site-packages**

or

**/opt/libreoffice5.4/program/python-core-3.5.0/lib/python3.5/site-packages**

When you start Extension with PyDev Debug Server running in Eclipse, you can debug the method or function with the enableRemoteDebugging decorator.

However, __init __ () can not be decorated by enableRemoteDebugging.
