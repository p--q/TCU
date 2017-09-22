#!/opt/libreoffice5.2/program/python
# -*- coding: utf-8 -*-
import unohelper
from com.sun.star.beans import PropertyValue
from com.sun.star.lang import XServiceInfo
from com.sun.star.awt import XContainerWindowEventHandler
from pq import XTcu  # 拡張機能で定義したインターフェイスをインポート。
from .optiondialog import dilaogHandler
from .common import enableRemoteDebugging  # デバッグ用デコレーター
IMPLE_NAME = None
SERVICE_NAME = None
def create(ctx, *args, imple_name, service_name):
	global IMPLE_NAME
	global SERVICE_NAME
	if IMPLE_NAME is None:
		IMPLE_NAME = imple_name 
	if SERVICE_NAME is None:
		SERVICE_NAME = service_name
	return TreeCommand(ctx, *args)
class TreeCommand(unohelper.Base, XServiceInfo, XTcu, XContainerWindowEventHandler):  
	def __init__(self, ctx, *args):  # argsはcreateInstanceWithArgumentsAndContext()でインスタンス化したときの引数。
		self.args = args  # 引数がないときもNoneではなくタプルが入る。
		smgr = ctx.getServiceManager()  # サービスマネジャーの取得。
		configurationprovider = smgr.createInstanceWithContext("com.sun.star.configuration.ConfigurationProvider", ctx)
		css = "com.sun.star"  # IDL名の先頭から省略する部分。
		properties = "OffLine", "RefURL", "RefDir", "IgnoredIDLs"  # config.xcuのpropノード名。
		nodepath = "/pq.Tcu.ExtensionData/Leaves/XUnoTreeCommandSettings/"  # config.xcuのノードへのパス。
		simplefileaccess = smgr.createInstanceWithContext("com.sun.star.ucb.SimpleFileAccess", ctx)  	
		self.consts = ctx, smgr, configurationprovider, css, properties, nodepath, simplefileaccess
	# XServiceInfo
	def getImplementationName(self):
		return IMPLE_NAME
	def supportsService(self, name):
		return name == SERVICE_NAME
	def getSupportedServiceNames(self):
		return (SERVICE_NAME,)		
	# XContainerWindowEventHandler
# 	@enableRemoteDebugging
	def callHandlerMethod(self, dialog, eventname, methodname):  # ブーリアンを返す必要あり。dialogはUnoControlDialog。 eventnameは文字列initialize, ok, backのいずれか。methodnameは文字列external_event。
		if methodname=="external_event":  # Falseのときがありうる?
			try:
				dilaogHandler(self.consts, dialog, eventname)
			except:
				import traceback; traceback.print_exc()
				return False
		return True		
	def getSupportedMethodNames(self):
		return "tree", "itree", "wtree"
	# XUnoTreeCommand
# 	@enableRemoteDebugging
	def tree(self, obj):
		
		
		return tuple()
	def itree(self, obj):
		pass
	
	def wtree(self, obj):
		pass	
