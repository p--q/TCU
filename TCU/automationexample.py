#!/opt/libreoffice5.4/program/python
# -*- coding: utf-8 -*-
import officehelper
import traceback
def connect():  # LibreOfficeをバックグラウンドで起動してコンポネントテクストを取得する。
	ctx = None
	try:
		ctx = officehelper.bootstrap()
		if ctx:
			print("Connected to a running office ...")
		return ctx
	except:
		pass
	return None
ctx = connect()
if ctx:
	smgr = ctx.getServiceManager()
	print("Using remote servicemanager\n")
	if not smgr:
		print( "ERROR: no service manager" )
		 
# PyUnoComponentの実行
try:	   
	tcu = smgr.createInstanceWithContext("pq.Tcu", ctx)  # サービス名か実装名でインスタンス化。

	tcu.wtree(ctx)
# 	print("\n".join(tcu.tree("com.sun.star.frame.Desktop")))
# 	print("\n".join(tcu.tree("com.sun.star.table.CellAddressConversion")))  # Can not get TypeDescription object of com.sun.star.table.CellAddressConversion.

except:
	traceback.print_exc()

 
# soffice.binの終了処理。これをしないとLibreOfficeを起動できなくなる。
desktop = smgr.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)
from com.sun.star.beans import PropertyValue
prop = PropertyValue(Name="Hidden",Value=True)
desktop.loadComponentFromURL("private:factory/swriter", "_blank", 0, (prop,))  # バックグラウンドでWriterのドキュメントを開く。
terminated = desktop.terminate()  # LibreOfficeをデスクトップに展開していない時はエラーになる。
if terminated:
	print("\nThe Office has been terminated.")  # 未保存のドキュメントがないとき。
else:
	print("\nThe Office is still running. Someone else prevents termination.")  # 未保存のドキュメントがあってキャンセルボタンが押された時。
	  