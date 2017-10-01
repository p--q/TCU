#!/opt/libreoffice5.4/program/python
# -*- coding: utf-8 -*-
import unohelper  # オートメーションには必須(必須なのはuno)。
def macro():
	ctx = XSCRIPTCONTEXT.getComponentContext()  # コンポーネントコンテクストの取得。
	smgr = ctx.getServiceManager()  # サービスマネージャーの取得。
	doc = XSCRIPTCONTEXT.getDocument()
	
	conv = doc.createInstance("com.sun.star.table.CellAddressConversion")

# 	mri = ctx.getServiceManager().createInstanceWithContext("mytools.Mri", ctx)
# 	mri.inspect(conv)
	
	tcu = smgr.createInstanceWithContext("pq.Tcu", ctx)  # サービス名か実装名でインスタンス化。
	
	tcu.wtree(conv)  

g_exportedScripts = macro, #マクロセレクターに限定表示させる関数をタプルで指定。
