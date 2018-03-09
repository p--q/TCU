#!/opt/libreoffice5.4/program/python
# -*- coding: utf-8 -*-
import unohelper  # オートメーションには必須(必須なのはuno)。
from wsgiref.simple_server import make_server
import webbrowser
import xml.etree.ElementTree as ET
def macro():	
	ctx = XSCRIPTCONTEXT.getComponentContext()  # コンポーネントコンテクストの取得。
	smgr = ctx.getServiceManager()  # サービスマネージャーの取得。 
	tcu = smgr.createInstanceWithContext("pq.Tcu", ctx)  # サービス名か実装名でインスタンス化。
	desktop = XSCRIPTCONTEXT.getDesktop() 
	doc = XSCRIPTCONTEXT.getDocument()
	controller = doc.getCurrentController()  # コントローラの取得。
	frame = controller.getFrame()  # フレームを取得。
	containerwindow = frame.getContainerWindow()
	componentwindow = frame.getComponentWindow()
	toolkit = containerwindow.getToolkit()
	
	dic_obj = {"Toolkit": toolkit, "Frame": frame}
	dic_objs = {"Container Window vs. Component Window": (containerwindow, componentwindow),\
			"Desktop vs. Frame": (desktop, frame)}
	createTrees(tcu, dic_obj, dic_objs)
def createTrees(tcu, dic_obj, dic_objs):	
	colors = "#FFCC99", "#FFCCCC", "#FF99CC", "#FFCCFF", "#CC99FF", "#CCCCFF", "#99CCFF", "#CCFFFF", "#99FFCC", "#CCFFCC" , "#CCFF99"  # Orange10, Red10, Pink10, Magenta10, Violet10, Blue10, SkyBlue10, Cyan10, Turquoise10, Green10, YellowGreen10	
	nodepairs = []
	[nodepairs.append(createNodes(key, tcu.wtreelines(val))) for key, val in dic_obj.items()]
	[nodepairs.append(createNodes(key, tcu.wcomparelines(*val))) for key, val in dic_objs.items()]
	n = len(colors)
	style = "\n".join(["#tabcontrol a:nth-child({0}), #tabbody div:nth-child({0}) {{background-color: {1};}}".format(i, colors[i%n]) for i in range(1, len(nodepairs)+1)])
	stylenode = Elem("style", text=style)				
	tabnodes = Elem("p", {"id": "tabcontrol"})
	tabbodynodes = Elem("div", {"id": "tabbody"})
	for tabnode, tabbodynode in nodepairs:
		tabnodes.append(tabnode)
		tabbodynodes.append(tabbodynode)
	root, scriptnode = createRoot()
	bodynode = root[1]
	
	# LibreOfficeのバージョン、かかった時間を入れる。
	
	bodynode.append(stylenode)
	bodynode.append(tabnodes)
	bodynode.append(tabbodynodes)
	bodynode.append(scriptnode)
	toBrowser(root)
def createNodes(name, lines):
	id = name.replace(" ", "").replace(".", "")
	tabnode = Elem("a", {"href": "#{}".format(id)}, text=name)
	tabbodynode = Elem("div", {"id": id})
	tabbodynode.append(Elem("p", text=name))
	nbsp = chr(0x00A0)  # Unicodeのノーブレークスペースの取得。
	txt = "<br/>".join(lines).replace(" "*2, nbsp*2).replace(" "*3, nbsp*3).replace(" "*4, nbsp*4)  # 連続したスペースはブラウザで一つにされるのでnbspに置換する。一つのスペースを置換するとタグ内まで置換されるのでダメ。
	xml = "<tt style='white-space: nowrap;'>{}</tt>".format(txt)
	tabbodynode.append(ET.XML(xml))
	return tabnode, tabbodynode
def createRoot():
	rt = Elem("html")
	rt.append(Elem("head"))
	rt[0].append(Elem("title", text="TCU - Tree Command for UNO"))
	rt[0].append(Elem("meta", {"meta": "UTF-8"}))
	rt.append(Elem("body"))
	style = """\
#tabcontrol {
	margin: 0;  /* タブ領域全体 */
}
/* タブ */
#tabcontrol a {
	display: inline-block;			/* インラインブロック化 */
	border-width: 1px 1px 0px 1px;	/* 下以外の枠線を引く */
	border-style: solid;			  /* 枠線の種類：実線 */
	border-color: black;			  /* 枠線の色：黒色 */
	border-radius: 0.75em 0.75em 0 0; /* 枠線の左上角と右上角だけを丸く */
	padding: 0.75em 1em;			  /* 内側の余白 */
	text-decoration: none;			/* リンクの下線を消す */
	color: black;					 /* 文字色：黒色 */
	background-color: white;		  /* 背景色：白色 */
	font-weight: bold;				/* 太字 */
	position: relative;			   /* JavaScriptでz-indexを調整するために必要 */
}
/* タブにマウスポインタが載った際 */
#tabcontrol a:hover {
	text-decoration: underline;   /* リンクの下線を引く */
}
/* タブの中身 */
#tabbody div {
	border: 1px solid black; /* 枠線：黒色の実線を1pxの太さで引く */
	margin-top: -1px;		/* 上側にあるタブと1pxだけ重ねるために「-1px」を指定 */
	padding: 1em;			/* 内側の余白量 */
	background-color: white; /* 背景色：白色 */
	position: relative;	  /* z-indexを調整するために必要 */
	z-index: 0;			  /* 重なり順序を「最も背面」にするため */
}"""
	rt[1].append(Elem("style", text=style))
	script = """\
var tabs = document.getElementById('tabcontrol').getElementsByTagName('a');
var pages = document.getElementById('tabbody').getElementsByTagName('div');	
function changeTab() {
	var targetid = this.href.substring(this.href.indexOf('#')+1, this.href.length);  // href属性値から対象のid名を抜き出す
	for (var i=0;i<pages.length; i++) {
		if (pages[i].id!=targetid) {
			pages[i].style.display = "none";
		} else { 
			pages[i].style.display = "inline-block";  // 指定のタブページだけを表示する
		}
	}
	for (var i=0;i<tabs.length; i++) {
		tabs[i].style.zIndex = "0";
	}
	this.style.zIndex = "10";  // クリックされたタブを前面に表示する
	return false;  // ページ遷移しないようにfalseを返す
}
for (var i=0;i<tabs.length; i++) {
	tabs[i].onclick = changeTab;  // すべてのタブに対して、クリック時にchangeTab関数が実行されるよう指定する
}
tabs[0].onclick();  // 最初は先頭のタブを選択	"""
	scriptnode = Elem("script", text=script)
	return rt, scriptnode
class Wsgi:
	def __init__(self, html):
		self.resp = html
	def app(self, environ, start_response):  # WSGIアプリ。引数はWSGIサーバから渡されるデフォルト引数。
		start_response('200 OK', [ ('Content-type','text/html; charset=utf-8')])  # charset=utf-8'がないと文字化けする時がある
		yield self.resp  # デフォルトエンコードはutf-8。
	def wsgiServer(self): 
		host, port = "localhost", 8080  # サーバが受け付けるポート番号を設定。
		httpd = make_server(host, port, self.app)  # appへの接続を受け付けるWSGIサーバを生成。
		url = "http://localhost:{}".format(port)  # 出力先のurlを取得。
		webbrowser.open_new_tab(url)   # デフォルトブラウザでurlを開く。
		httpd.handle_request()  # リクエストを1回だけ受け付けたらサーバを終了させる。ローカルファイルはセキュリティの制限で開けない。
def toBrowser(root):
	html = ET.tostring(root, encoding="utf-8",  method="html")  # utf-8にエンコードする。utf-8ではなくunicodeにすると文字列になる。method="html"にしないと<script>内がhtmlエンティティになってしまう。
	server = Wsgi(html)
	server.wsgiServer()	
class Elem(ET.Element):  
    '''
    キーワード引数textでテキストノードを付加するxml.etree.ElementTree.Element派生クラス。
    '''
    def __init__(self, tag, attrib={},  **kwargs):  # textキーワードは文字列のみしか受け取らない。  
        if "text" in kwargs:  # textキーワードがあるとき
            txt = kwargs["text"]
            del kwargs["text"]  
            super().__init__(tag, attrib, **kwargs)
            self._text(txt)
        else:
            super().__init__(tag, attrib, **kwargs)
    def _text(self, txt):
        self.text = txt
g_exportedScripts = macro, #マクロセレクターに限定表示させる関数をタプルで指定。
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
