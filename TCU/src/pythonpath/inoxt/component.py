#!/opt/libreoffice5.2/program/python
# -*- coding: utf-8 -*-
import unohelper, os
from com.sun.star.beans import PropertyValue
from com.sun.star.lang import XServiceInfo
from com.sun.star.awt import XContainerWindowEventHandler
from pq import XTcu  # 拡張機能で定義したインターフェイスをインポート。
from .optiondialog import dilaogHandler
def enableRemoteDebugging(func):  # デバッグサーバーに接続したい関数やメソッドにつけるデコレーター。主にリスナーのメソッドのデバッグ目的。__init()__には不可。
	def wrapper(*args, **kwargs):  # /opt/libreoffice5.2/program/python-core-3.5.0/lib/python3.5/site-packages/sites.pthにpydevd.pyがあるフォルダへのパス設定が必要。
		import pydevd; pydevd.settrace(stdoutToServer=True, stderrToServer=True)  # デバッグサーバーを起動していた場合はここでブレークされる。import pydevdは時間がかかる。
		try:
			func(*args, **kwargs)  # Step Intoして中に入る。
		except:
			import traceback; traceback.print_exc()  # これがないとPyDevのコンソールにトレースバックが表示されない。stderrToServer=Trueが必須。
	return wrapper
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
		self.css = "com.sun.star"  # IDL名の先頭から省略する部分。
		smgr = ctx.getServiceManager()  # サービスマネジャーの取得。
		self.args = args  # 引数がないときもNoneではなくタプルが入る。
		configurationprovider = smgr.createInstanceWithContext("com.sun.star.configuration.ConfigurationProvider", ctx)
		readConfigs, self.writeConfig = createConfigAccessor(configurationprovider, "/pq.Tcu.ExtensionData/Leaves/XUnoTreeCommandSettings/OptionDialog")
		self._offline, self._refurl, self._refdir, ignoredidls = readConfigs("OffLine", "RefURL", "RefDir", "IgnoredIDLs")  # LibreOfficeの設定ファイルからデータを読み込む。
		self._ignoredidls = set(ignoredidls.split(","))
		self.ctx = ctx
		self.smgr = smgr
		self.configurationprovider = configurationprovider
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
				dilaogHandler(self, dialog, eventname)
			except:
				return False
		return True			
	# XUnoTreeCommand
# 	@enableRemoteDebugging
	def tree(self, obj):
		
		
		return tuple()
	def itree(self, obj):
		pass
	
	def wtree(self, obj):
		pass	

# 	 XUnoTreeCommandSettings
	def getRefURL(self):	
		return self._refurl
	def setRefURL(self, refurl):
		self._refurl = refurl
		self.writeConfig("RefURL", refurl)  # LibreOfficeの設定ファイルに書き込む。
	RefURL = property(getRefURL, setRefURL)  # メソッドをPythonオブジェクトのアトリビュートに変換(それをPythonではプロパティという)。
	def getRefDir(self):	
		return self._refdir
	def setRefDir(self, refdir):
		self._refdir = refdir
		self.writeConfig("RefURL", refdir)  # LibreOfficeの設定ファイルに書き込む。
	RefDir = property(getRefDir, setRefDir)  # メソッドをPythonオブジェクトのアトリビュートに変換(それをPythonではプロパティという)。
	def getOffLine(self):
		return self._offline
	def setOffLine(self, flag):
		if flag:
			if not os.path.exists(self.RefDir):  # SDKのリファレンスフォルダが存在しない時。
				raise RuntimeError("SDK reference folder does not exists.".format(self.RefDir))  # 例外を出す。
		self._offline = flag
		self.writeConfig("OffLine", flag)  # LibreOfficeの設定ファイルに書き込む。
	OffLine = property(getOffLine, setOffLine)  # メソッドをPythonオブジェクトのアトリビュートに変換(それをPythonではプロパティという)。
	def getIgnoredIDLs(self):	
		return tuple("{}{}".format(self.css, i) if i.startswith(".") else i for i in self._ignoredidls)
	def setIgnoredIDLs(self, idls):
		self._ignoredidls.update([i.replace(self.css, "") if i.startswith(self.css) else i for i in idls])
		ignoredidls = ",".join(self._ignoredidls)
		self.writeConfig("IgnoredIDLs", ignoredidls)  # LibreOfficeの設定ファイルに書き込む。
	IgnoredIDLs = property(getIgnoredIDLs, setIgnoredIDLs)  # メソッドをPythonオブジェクトのアトリビュートに変換(それをPythonではプロパティという)。
def createConfigAccessor(configurationprovider, rootpath):  # コンポーネントデータノードへのアクセス。
	node = PropertyValue(Name="nodepath", Value=rootpath)
	root = configurationprovider.createInstanceWithArguments("com.sun.star.configuration.ConfigurationUpdateAccess", (node,))
# 	@enableRemoteDebugging
	def readConfigs(*args):  # 値の取得。整数か文字列かブーリアンのいずれか。コンポーネントスキーマノードの設定に依存。
		return root.getPropertyValues(args)
	def writeConfig(name, value):  # 値の書き込み。整数か文字列かブーリアンのいずれか。コンポーネントスキーマノードの設定に依存。
		try:
			root.setPropertyValue(name, value)
			root.commitChanges()  # 変更値の書き込み。
		except:
			import traceback; traceback.print_exc()
	return readConfigs, writeConfig