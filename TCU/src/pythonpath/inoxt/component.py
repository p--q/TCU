#!/opt/libreoffice5.4/program/python
# -*- coding: utf-8 -*-
import unohelper
import re, os
from com.sun.star.beans import PropertyValue
from com.sun.star.lang import XServiceInfo
from com.sun.star.awt import XContainerWindowEventHandler
from pq import XTcu  # 拡張機能で定義したインターフェイスをインポート。
from .optiondialog import dilaogHandler
from .trees import createTree
from .wsgi import Wsgi, createHTMLfile
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
		return "tree", "wtree"
	# XUnoTreeCommand
	def tree(self, obj):
		ctx, configurationprovider, css, fns_keys, dummy_offline, dummy_prefix, idlsset = getConfigs(self.consts)
		outputs = []
		fns = {key: outputs.append for key in fns_keys}
		args = ctx, configurationprovider, css, fns, idlsset, outputs, " "
		createTree(args, obj)
		return outputs
	def wtree(self, obj):
		ctx, configurationprovider, css, fns_keys, offline, prefix, idlsset = getConfigs(self.consts)
		outputs = ['<tt>']  # 出力行を収納するリストを初期化。等幅フォントのタグを指定。
		fns = createFns(prefix, fns_keys, outputs)
		args = ctx, configurationprovider, css, fns, idlsset, outputs, "&nbsp;"
		createTree(args, obj)
		outputs.append("</tt>")		
		html = "<br>".join(outputs)
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
def createFns(prefix, fns_keys, outputs):
	reg_idl = re.compile(r'(?<!\w)\.[\w\.]+')  # IDL名を抽出する正規表現オブジェクト。
	reg_i = re.compile(r'(?<!\w)\.[\w\.]+\.X[\w]+')  # インターフェイス名を抽出する正規表現オブジェクト。
	reg_e = re.compile(r'(?<!\w)\.[\w\.]+\.[\w]+Exception')  # 例外名を抽出する正規表現オブジェクト。
	def _make_anchor(typ, i, item_with_branch):
		lnk = "<a href='{}{}com_1_1sun_1_1star{}.html' target='_blank' style='text-decoration:none;'>{}</a>".format(prefix, typ, i.replace(".", "_1_1"), i)  # 下線はつけない。
		return item_with_branch.replace(i, lnk)
	def _make_link(typ, regex, item_with_branch):
		idl = regex.findall(item_with_branch)  # 正規表現でIDL名を抽出する。
		if idl:
			lnk = "<a href='{}{}com_1_1sun_1_1star{}.html' target='_blank'>{}</a>".format(prefix, typ, idl[0].replace(".", "_1_1"), idl[0])  # サービス名のアンカータグを作成。
			outputs.append(item_with_branch.replace(" ", "&nbsp;").replace(idl[0], lnk))  # 半角スペースを置換後にサービス名をアンカータグに置換。
		else:
			outputs.append(item_with_branch.replace(" ", "&nbsp;"))  # 半角スペースを置換。	
	def _fn(item_with_branch):  # サービス名とインターフェイス名以外を出力するときの関数。
		idl = set(reg_idl.findall(item_with_branch)) # 正規表現でIDL名を抽出する。
		inf = reg_i.findall(item_with_branch) # 正規表現でインターフェイス名を抽出する。
		exc = reg_e.findall(item_with_branch) # 正規表現で例外名を抽出する。
		idl.difference_update(inf, exc)  # IDL名のうちインターフェイス名と例外名を除く。
		idl = list(idl)  # 残ったIDL名はすべてStructと考えて処理する。
		item_with_branch = item_with_branch.replace(" ", "&nbsp;")  # まず半角スペースをHTMLに置換する。
		for i in inf:  # インターフェイス名があるとき。
			item_with_branch = _make_anchor("interface", i, item_with_branch)
		for i in exc:  # 例外名があるとき。
			item_with_branch = _make_anchor("exception", i, item_with_branch)
		for i in idl:  # インターフェイス名と例外名以外について。
			item_with_branch = _make_anchor("struct", i, item_with_branch)
		outputs.append(item_with_branch)
	def _fn_s(item_with_branch):  # サービス名にアンカータグをつける。
		_make_link("service", reg_idl, item_with_branch)
	def _fn_i(item_with_branch):  # インターフェイス名にアンカータグをつける。
		_make_link("interface", reg_i, item_with_branch)	
	def _fn_nolink(item_with_branch):
		outputs.append(item_with_branch.replace(" ", "&nbsp;"))
	fns = {key: _fn for key in fns_keys[2:5]}
	fns[fns_keys[0]] = _fn_s		
	fns[fns_keys[1]] = _fn_i
	fns[fns_keys[5]] = _fn_nolink
	return fns
