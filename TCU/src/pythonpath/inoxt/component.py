#!/opt/libreoffice5.4/program/python
# -*- coding: utf-8 -*-
import unohelper
import re, os
from com.sun.star.beans import PropertyValue  # Struct
from com.sun.star.lang import XServiceInfo
from com.sun.star.awt import XContainerWindowEventHandler
from com.sun.star.uno.TypeClass import ENUM, TYPEDEF, STRUCT, EXCEPTION, INTERFACE, CONSTANTS  # enum
from pq import XTcu  # 拡張機能で定義したインターフェイスをインポート。
from .optiondialog import dilaogHandler
from .wsgi import Wsgi, createHTMLfile
from .wcompare import wCompare
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
	def callHandlerMethod(self, dialog, eventname, methodname):  # ブーリアンを返す必要あり。dialogはUnoControlDialog。 eventnameは文字列initialize, ok, backのいずれか。methodnameは文字列external_event。
		if methodname=="external_event":  # Falseのときがありうる?
			try:
				dilaogHandler(self.consts, dialog, eventname)
			except:
				import traceback; traceback.print_exc()
				return False
		return True		
	def getSupportedMethodNames(self):
		return "tree", "wtree"
	# XUnoTreeCommand
	def treelines(self, obj):  # 一行ずつの文字列のシークエンスを返す。
		ctx, configurationprovider, css, fns_keys, dummy_offline, dummy_prefix, idlsset = getConfigs(self.consts)
		outputs = []
		fns = {key: outputs.append for key in fns_keys}
		args = ctx, configurationprovider, css, fns, idlsset, outputs
		wCompare(args, obj, None)
		return outputs
	def wtreelines(self, obj):  # 一行ずつの文字列のシークエンスを返す。連続スペースはnbspに置換の必要あり。
		ctx, configurationprovider, css, fns_keys, dummy, prefix, idlsset = getConfigs(self.consts)
		outputs = []
		fns = createFns(ctx, css, prefix, fns_keys, outputs)
		args = ctx, configurationprovider, css, fns, idlsset, outputs
		wCompare(args, obj, None)		
		return outputs
	def wcomparelines(self, obj1, obj2):  # 一行ずつの文字列のシークエンスを返す。連続スペースはnbspに置換の必要あり。
		ctx, configurationprovider, css, fns_keys, dummy, prefix, idlsset = getConfigs(self.consts)
		outputs = []  # 出力行を収納するリストを初期化。等幅フォントのタグを指定。
		fns = createFns(ctx, css, prefix, fns_keys, outputs)
		args = ctx, configurationprovider, css, fns, idlsset, outputs
		wCompare(args, obj1, obj2)
		return outputs
	def wtree(self, obj):  # obj1とobj2を比較して結果をウェブブラウザに出力する。
		ctx, configurationprovider, css, fns_keys, offline, prefix, idlsset = getConfigs(self.consts)
		outputs = ['<code style="white-space: nowrap;">']  # 出力行を収納するリストを初期化。等幅フォントのタグを指定。
		fns = createFns(ctx, css, prefix, fns_keys, outputs)
		args = ctx, configurationprovider, css, fns, idlsset, outputs
		wCompare(args, obj, None)
		createHtml(ctx, offline, outputs)  # ウェブブラウザに出力。
	def wcompare(self, obj1, obj2):  # obj1とobj2を比較して結果をウェブブラウザに出力する。
		ctx, configurationprovider, css, fns_keys, offline, prefix, idlsset = getConfigs(self.consts)
		outputs = ['<code style="white-space: nowrap;">']  # 出力行を収納するリストを初期化。等幅フォントのタグを指定。
		fns = createFns(ctx, css, prefix, fns_keys, outputs)
		args = ctx, configurationprovider, css, fns, idlsset, outputs
		wCompare(args, obj1, obj2)
		createHtml(ctx, offline, outputs)  # ウェブブラウザに出力。
def createHtml(ctx, offline, outputs):  # ウェブブラウザに出力。
	outputs.append("</code>")	
	html = "<br/>".join(outputs).replace(" ", chr(0x00A0))  # 半角スペースをノーブレークスペースに置換する。
	html = re.sub(r'(?<!\u00A0)\u00A0(?!\u00A0)', " ", html)  # タグ内にノーブレークスペースはエラーになるので連続しないノーブレークスペースを半角スペースに戻す。
	title = "TCU - Tree Command for UNO"
	if offline:  # ローカルリファレンスを使うときはブラウザのセキュリティの制限のためにhtmlファイルを開くようにしないとローカルファイルが開けない。
		pathsettingssingleton = ctx.getByName('/singletons/com.sun.star.util.thePathSettings')
		fileurl = pathsettingssingleton.getPropertyValue("Temp")
		systempath = unohelper.fileUrlToSystemPath(fileurl)
		filepath = os.path.join(systempath, "tcu_output.html")
		createHTMLfile(filepath, title, html)
	else:
		server = Wsgi(title, html)
		server.wsgiServer()		
def getConfigs(consts):
	ctx, smgr, configurationprovider, css, properties, nodepath, simplefileaccess = consts
	fns_keys = "SERVICE", "INTERFACE", "PROPERTY", "INTERFACE_METHOD", "INTERFACE_ATTRIBUTE", "NOLINK"  # fnsのキーのタプル。
	node = PropertyValue(Name="nodepath", Value="{}OptionDialog".format(nodepath))
	root = configurationprovider.createInstanceWithArguments("com.sun.star.configuration.ConfigurationAccess", (node,))
	offline, refurl, refdir, idlstext = root.getPropertyValues(properties)  # コンポーネントデータノードから値を取得する。		
	prefix = "https://{}".format(refurl)
	if offline:  # ローカルリファレンスを使うときはprefixを置換する。
		pathsubstservice = smgr.createInstanceWithContext("com.sun.star.comp.framework.PathSubstitution", ctx)
		fileurl = pathsubstservice.substituteVariables(refdir, True)  # $(inst)を変換する。fileurlが返ってくる。
		if simplefileaccess.exists(fileurl):
			systempath = os.path.normpath(unohelper.fileUrlToSystemPath(fileurl))  # fileurlをシステムパスに変換して正規化する。
			prefix = "file://{}".format(systempath)
		else:
			raise RuntimeError("Local API Reference does not exists.")
	if not prefix.endswith("/"):
		prefix = "{}/".format(prefix)
	idls = "".join(idlstext.split()).split(",")  # xmlがフォーマットされていると空白やタブが入ってくるのでそれを除去してリストにする。
	idlsset = set("{}{}".format(css, i) if i.startswith(".") else i for i in idls)  # "com.sun.star"が略されていれば付ける。
	return ctx, configurationprovider, css, fns_keys, offline, prefix, idlsset	
def createFns(ctx, css, prefix, fns_keys, outputs):
	reg_idl = re.compile(r'(?<!\w)\.[\w\.]+')  # IDL名を抽出する正規表現オブジェクト。
	reg_i = re.compile(r'(?<!\w)\.[\w\.]+\.X[\w]+')  # インターフェイス名を抽出する正規表現オブジェクト。
	tdm = ctx.getByName('/singletons/com.sun.star.reflection.theTypeDescriptionManager')  # TypeDescriptionManagerをシングルトンでインスタンス化。
	def _make_anchor(typ, i, item_with_branch, fragment):
		m = ".".join(i.split(".")[:-1]) if fragment else i  # フラグメントがあるとき(ENUMやTYPEDEFのとき)モジュールへのパスの取得。
		lnk = "<a href='{}{}com_1_1sun_1_1star{}.html{}' target='_blank' style='text-decoration:none;'>{}</a>".format(prefix, typ, m.replace(".", "_1_1"), fragment, i)  # 下線はつけない。
		return item_with_branch.replace(i, lnk)
	def _make_link(typ, regex, item_with_branch):
		idl = regex.findall(item_with_branch)  # 正規表現でIDL名を抽出する。
		if idl:
			lnk = "<a href='{}{}com_1_1sun_1_1star{}.html' target='_blank'>{}</a>".format(prefix, typ, idl[0].replace(".", "_1_1"), idl[0])  # サービス名のアンカータグを作成。
			outputs.append(item_with_branch.replace(idl[0], lnk)) 
		else:
			outputs.append(item_with_branch)
	def _fn(item_with_branch):  # サービス名とインターフェイス名以外を出力するときの関数。
		idl = reg_idl.findall(item_with_branch) # 正規表現でIDL名を抽出する。
		for i in idl:  # STRUCT, EXCEPTION, INTERFACE, CONSTANTSのIDLのみアンカーを付ける。ENUM, TYPEDEFはリンクを取得できない。
			j = tdm.getByHierarchicalName("{}{}".format(css, i) if i.startswith(".") else i)  # TypeDescriptionオブジェクトを取得。
			typeclass = j.getTypeClass()  # enum TypeClassを取得。辞書のキーにはなれない、uno.RuntimeException: <class 'TypeError'>: unhashable type: 'Enum'となる。
			fragment = ""
			if typeclass==INTERFACE:
				t = "interface"
			elif typeclass==EXCEPTION:
				t = "exception"
			elif typeclass==STRUCT:
				t = "struct"		
			elif typeclass== CONSTANTS:
				t = "namespace"				
			elif typeclass==ENUM:				
				t = "namespace"	
				fragment = "#enum-members"
			elif  typeclass==TYPEDEF:  
				t = "namespace"	
				fragment = "#typedef-members"
			item_with_branch = _make_anchor(t, i, item_with_branch, fragment)
		outputs.append(item_with_branch)
	def _fn_s(item_with_branch):  # サービス名にアンカータグをつける。
		_make_link("service", reg_idl, item_with_branch)
	def _fn_i(item_with_branch):  # インターフェイス名にアンカータグをつける。
		_make_link("interface", reg_i, item_with_branch)	
	def _fn_nolink(item_with_branch):
		outputs.append(item_with_branch)
	fns = {key: _fn for key in fns_keys[2:5]}  # キー PROPERTY, INTERFACE_METHOD, INTERFACE_ATTRIBUTE
	fns[fns_keys[0]] = _fn_s  # SERVICE		
	fns[fns_keys[1]] = _fn_i  # INTERFACE
	fns[fns_keys[5]] = _fn_nolink  # NOLINK
	return fns
