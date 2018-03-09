#!/opt/libreoffice5.4/program/python
# -*- coding: utf-8 -*-
import unohelper  # オートメーションには必須(必須なのはuno)。
def macro():
	ctx = XSCRIPTCONTEXT.getComponentContext()  # コンポーネントコンテクストの取得。
	smgr = ctx.getServiceManager()  # サービスマネージャーの取得。 
	tcu = smgr.createInstanceWithContext("pq.Tcu", ctx)  # サービス名か実装名でインスタンス化。
	doc = XSCRIPTCONTEXT.getDocument()
	controller = doc.getCurrentController()  # コントローラの取得。
	frame = controller.getFrame()  # フレームを取得。
	containerwindow = frame.getContainerWindow()
	lines = tcu.wtreelines(containerwindow.getToolkit())
	toBrowser(lines)
g_exportedScripts = macro, #マクロセレクターに限定表示させる関数をタプルで指定。
from wsgiref.simple_server import make_server
import webbrowser
_resp = '''\
<html>
  <head>
	 <title>{0}</title>
	 <meta charset="UTF-8">
   </head>
   <body>
	 {1}
   </body>
</html>'''  # ローカルファイルとして開くときは<meta charset="UTF-8">がないと文字化けする。
class Wsgi:
	def __init__(self, title, html):
		self.resp = _resp.format(title, html)
	def app(self, environ, start_response):  # WSGIアプリ。引数はWSGIサーバから渡されるデフォルト引数。
		start_response('200 OK', [ ('Content-type','text/html; charset=utf-8')])  # charset=utf-8'がないと文字化けする時がある
		yield self.resp.encode()  # デフォルトエンコードはutf-8
	def wsgiServer(self): 
		host, port = "localhost", 8080  # サーバが受け付けるポート番号を設定。
		httpd = make_server(host, port, self.app)  # appへの接続を受け付けるWSGIサーバを生成。
		url = "http://localhost:{}".format(port)  # 出力先のurlを取得。
		webbrowser.open_new_tab(url)   # デフォルトブラウザでurlを開く。
		httpd.handle_request()  # リクエストを1回だけ受け付けたらサーバを終了させる。ローカルファイルはセキュリティの制限で開けない。
def toBrowser(lines):
	outputs = ['<tt style="white-space: nowrap;">']  # 出力行を収納するリストを初期化。等幅フォントのタグを指定。
	outputs.extend(lines)
	outputs.append("</tt>")		
	html = "<br>".join(outputs).replace(" "*2, "&nbsp;"*2).replace(" "*3, "&nbsp;"*3).replace(" "*4, "&nbsp;"*4)  # 連続したスペースはブラウザで一つにされるのでnbspに置換する。一つのスペースを置換するとタグ内まで置換されるのでダメ。
	title = "TCU - Tree Command for UNO"	
	server = Wsgi(title, html)
	server.wsgiServer()	
if __name__ == "__main__":  # オートメーションで実行するとき
	def automation():  # オートメーションのためにglobalに出すのはこの関数のみにする。
		import officehelper
		from functools import wraps
		import sys
		from com.sun.star.beans import PropertyValue
		from com.sun.star.script.provider import XScriptContext  
		def connectOffice(func):  # funcの前後でOffice接続の処理
			@wraps(func)
			def wrapper():  # LibreOfficeをバックグラウンドで起動してコンポーネントテクストとサービスマネジャーを取得する。
				try:
					ctx = officehelper.bootstrap()  # コンポーネントコンテクストの取得。
				except:
					print("Could not establish a connection with a running office.", file=sys.stderr)
					sys.exit()
				print("Connected to a running office ...")
				smgr = ctx.getServiceManager()  # サービスマネジャーの取得。
				print("Using {} {}".format(*_getLOVersion(ctx, smgr)))  # LibreOfficeのバージョンを出力。
				return func(ctx, smgr)  # 引数の関数の実行。
			def _getLOVersion(ctx, smgr):  # LibreOfficeの名前とバージョンを返す。
				cp = smgr.createInstanceWithContext('com.sun.star.configuration.ConfigurationProvider', ctx)
				node = PropertyValue(Name = 'nodepath', Value = 'org.openoffice.Setup/Product' )  # share/registry/main.xcd内のノードパス。
				ca = cp.createInstanceWithArguments('com.sun.star.configuration.ConfigurationAccess', (node,))
				return ca.getPropertyValues(('ooName', 'ooSetupVersion'))  # LibreOfficeの名前とバージョンをタプルで返す。
			return wrapper
		@connectOffice  # createXSCRIPTCONTEXTの引数にctxとsmgrを渡すデコレータ。
		def createXSCRIPTCONTEXT(ctx, smgr):  # XSCRIPTCONTEXTを生成。
			class ScriptContext(unohelper.Base, XScriptContext):
				def __init__(self, ctx):
					self.ctx = ctx
				def getComponentContext(self):
					return self.ctx
				def getDesktop(self):
					return ctx.getByName('/singletons/com.sun.star.frame.theDesktop')  # com.sun.star.frame.Desktopはdeprecatedになっている。
				def getDocument(self):
					return self.getDesktop().getCurrentComponent()
			return ScriptContext(ctx)  
		XSCRIPTCONTEXT = createXSCRIPTCONTEXT()  # XSCRIPTCONTEXTの取得。
		doc = XSCRIPTCONTEXT.getDocument()  # 現在開いているドキュメントを取得。
		doctype = "scalc", "com.sun.star.sheet.SpreadsheetDocument"  # Calcドキュメントを開くとき。
# 		doctype = "swriter", "com.sun.star.text.TextDocument"  # Writerドキュメントを開くとき。
		if (doc is None) or (not doc.supportsService(doctype[1])):  # ドキュメントが取得できなかった時またはCalcドキュメントではない時
			XSCRIPTCONTEXT.getDesktop().loadComponentFromURL("private:factory/{}".format(doctype[0]), "_blank", 0, ())  # ドキュメントを開く。ここでdocに代入してもドキュメントが開く前にmacro()が呼ばれてしまう。
		flg = True
		while flg:
			doc = XSCRIPTCONTEXT.getDocument()  # 現在開いているドキュメントを取得。
			if doc is not None:
				flg = (not doc.supportsService(doctype[1]))  # ドキュメントタイプが確認できたらwhileを抜ける。
		return XSCRIPTCONTEXT
	XSCRIPTCONTEXT = automation()  # XSCRIPTCONTEXTを取得。 
	macro()  # マクロの実行。
