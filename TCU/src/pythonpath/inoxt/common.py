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
