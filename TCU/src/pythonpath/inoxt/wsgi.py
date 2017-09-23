#!/opt/libreoffice5.4/program/python
# -*- coding: utf-8 -*-
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

def createHTMLfile(filepath, title, html):
	with open(filepath, 'w', encoding='UTF-8') as f:  # htmlファイルをUTF-8で作成。すでにあるときは上書き。ホームフォルダに出力される。
		f.writelines(_resp.format(title, html))  # シークエンスデータをファイルに書き出し。
		webbrowser.open_new_tab(f.name)  # デフォルトのブラウザの新しいタブでhtmlファイルを開く。
