# TCU - Tree Command for UNO

Output the API tee from the UNO object or IDL name.

## Installing TCU extension

Add <a href="https://github.com/p--q/TCU/tree/master/TCU/oxt">TCU.oxt</a> with Extension Manager.

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
