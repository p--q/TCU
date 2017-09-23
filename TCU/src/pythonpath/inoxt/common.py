#!/opt/libreoffice5.4/program/python
# -*- coding: utf-8 -*-
import os, gettext
from com.sun.star.beans import PropertyValue
def localization(configurationprovider):  # 地域化。moファイルの切替。Linuxでは不要だがWindowsでは設定が必要。
	node = PropertyValue(Name="nodepath", Value="/org.openoffice.Setup/L10N")
	root = configurationprovider.createInstanceWithArguments("com.sun.star.configuration.ConfigurationAccess", (node,))
	lang = root.getPropertyValue("ooLocale"),  # 現在のロケールを取得。タプルかリストで渡す。
	lodir = os.path.join(os.path.abspath(os.path.dirname(__file__)), "locale")  # このスクリプトと同じファルダにあるlocaleフォルダの絶対パスを取得。
	mo = "default"  # moファイル名。拡張子なし。
	t = gettext.translation(mo, lodir, lang, fallback=True)  # Translations インスタンスを取得。moファイルがなくてもエラーはでない。
	return t.gettext  # 関数t.gettextを返す。
def enableRemoteDebugging(func):  # デバッグサーバーに接続したい関数やメソッドにつけるデコレーター。主にリスナーのメソッドのデバッグ目的。__init()__には不可。
	def wrapper(*args, **kwargs):  # /opt/libreoffice5.2/program/python-core-3.5.0/lib/python3.5/site-packages/sites.pthにpydevd.pyがあるフォルダへのパス設定が必要。
		import pydevd; pydevd.settrace(stdoutToServer=True, stderrToServer=True)  # デバッグサーバーを起動していた場合はここでブレークされる。import pydevdは時間がかかる。
		try:
			func(*args, **kwargs)  # Step Intoして中に入る。
		except:
			import traceback; traceback.print_exc()  # これがないとPyDevのコンソールにトレースバックが表示されない。stderrToServer=Trueが必須。
	return wrapper
