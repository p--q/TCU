#!/opt/libreoffice5.2/program/python
# -*- coding: utf-8 -*-
import os, gettext
from com.sun.star.beans import PropertyValue

def localization(configurationprovider):  # 地域化。moファイルの切替。Linuxでは不要だがWindowsでは設定が必要。
	global _
	node = PropertyValue(Name="nodepath", Value="/org.openoffice.Setup/L10N")
	root = configurationprovider.createInstanceWithArguments("com.sun.star.configuration.ConfigurationAccess", (node,))
	lang = root.getHierarchicalPropertyValue("ooLocale"),  # 現在のロケールを取得。タプルかリストで渡す。
	lodir = os.path.join(os.path.abspath(os.path.dirname(__file__)), "locale")  # このスクリプトと同じファルダにあるlocaleフォルダの絶対パスを取得。
	mo = os.path.splitext(os.path.basename(__file__))[0]  # moファイル名。拡張子なし。このスクリプトファイルと同名と想定。
	t = gettext.translation(mo, lodir, lang, fallback=True)  # Translations インスタンスを取得。moファイルがなくてもエラーはでない。
	_ = t.gettext  # _にt.gettext関数を代入。
def dilaogHandler(treecommand, dialog, eventname):
	ctx = treecommand.ctx
	smgr = treecommand.smgr
	configurationprovider = treecommand.configurationprovider
# 	localization(configurationprovider)  # 地域化。
	if eventname=="initialize":  # オプションダイアログがアクティブになった時
		addControl = controlCreator(ctx, smgr, dialog)  # オプションダイアログdialogにコントロールを追加する関数を取得。
		addControl("FixedLine", {"PositionX": 5, "PositionY": 13, "Width": 250, "Height": 10, "Label": "TCU - Tree Command for UNO"})  # 文字付き水平線。



# IgnoredIDLs
# self.RefURL = "api.libreoffice.org/docs/idl/ref/" 
# self.RefDir = os.path.join(inst_path, "sdk", "docs", "idl", "ref") 
# OffLine
		
	
	elif eventname=="ok":  # OKボタンが押された時
		pass
	elif eventname=="back":  # 元に戻すボタンが押された時
		pass



# 	def __init__(self, ctx, *args):
# 		self.ctx = ctx
# 		self.smgr = ctx.getServiceManager()
# 		self.readConfig, self.writeConfig = createConfigAccessor(ctx, self.smgr, "/com.pq.blogspot.comp.ExtensionExample.ExtensionData/Leaves/MaximumPaperSize")  # config.xcsに定義していあるコンポーネントデータノードへのパス。
# 		self.cfgnames = "Width", "Height"
# 		self.defaults = self.readConfig("Defaults/Width", "Defaults/Height")
# 	# XContainerWindowEventHandler
# 	def callHandlerMethod(self, dialog, eventname, methodname):  # ブーリアンを返す必要あり。dialogはUnoControlDialog。 eventnameは文字列initialize, ok, backのいずれか。methodnameは文字列external_event。
# 		if methodname==self.METHODNAME:  # Falseのときがありうる?
# 			try:
# 				if eventname=="initialize":  # オプションダイアログがアクティブになった時
# 					maxwidth, maxheight = self.readConfig(*self.cfgnames)  # コンポーネントデータノードの値を取得。取得した値は文字列。
# 					maxwidth = maxwidth or self.defaults[0]
# 					maxheight = maxheight or self.defaults[1]
# 					buttonlistener = ButtonListener(dialog, self.defaults)  # ボタンリスナーをインスタンス化。
# 					addControl = controlCreator(self.ctx, self.smgr, dialog)  # オプションダイアログdialogにコントロールを追加する関数を取得。
# 					addControl("FixedLine", {"PositionX": 5, "PositionY": 13, "Width": 250, "Height": 10, "Label": _("Maximum page size")})  # 文字付き水平線。
# 					addControl("FixedText", {"PositionX": 11, "PositionY": 39, "Width": 49, "Height": 15, "Label": _("Width"), "NoLabel": True})  # 文字列。
# 					addControl("NumericField", {"PositionX": 65, "PositionY": 39, "Width": 60, "Height": 15, "Spin": True, "ValueMin": 0, "Value": float(maxwidth), "DecimalAccuracy": 2, "HelpText": _("Width")})  # 上下ボタン付き数字枠。小数点2桁、floatに変換して値を代入。
# 					addControl("NumericField", {"PositionX": 65, "PositionY": 64, "Width": 60, "Height": 15, "Spin": True, "ValueMin": 0, "Value": float(maxheight), "DecimalAccuracy": 2, "HelpText": _("Height")})  # 同上。
# 					addControl("FixedText", {"PositionX": 11, "PositionY": 66, "Width": 49, "Height": 15, "Label": _("Height"), "NoLabel": True})  # 文字列。
# 					addControl("FixedText", {"PositionX": 127, "PositionY": 42, "Width": 25, "Height": 15, "Label": "cm", "NoLabel": True})  # 文字列。
# 					addControl("FixedText", {"PositionX": 127, "PositionY": 68, "Width": 25, "Height": 15, "Label": "cm", "NoLabel": True})  # 文字列。
# 					addControl("Button", {"PositionX": 155, "PositionY": 39, "Width": 50, "Height": 15, "Label": _("~Default")}, {"setActionCommand": "width", "addActionListener": buttonlistener})  # ボタン。
# 					addControl("Button", {"PositionX": 155, "PositionY": 64, "Width": 50, "Height": 15, "Label": _("~Default")}, {"setActionCommand": "height", "addActionListener": buttonlistener})  # ボタン。
# 				elif eventname=="ok":  # OKボタンが押された時
# 					maxwidth = dialog.getControl("NumericField1").getModel().Value  # NumericFieldコントロールから値を取得。
# 					maxheight = dialog.getControl("NumericField2").getModel().Value  # NumericFieldコントロールから値を取得。
# 					self.writeConfig(self.cfgnames, (str(maxwidth), str(maxheight)))  # 取得した値を文字列にしてコンポーネントデータノードに保存。
# 				elif eventname=="back":  # 元に戻すボタンが押された時
# 					maxwidth, maxheight = self.readConfig(*self.cfgnames)
# 					dialog.getControl("NumericField1").getModel().Value= float(maxwidth)  # コンポーネントデータノードの値を取得。
# 					dialog.getControl("NumericField2").getModel().Value= float(maxheight)  # コンポーネントデータノードの値を取得。
# 			except:
# 				traceback.print_exc()  # トレースバックはimport pydevd; pydevd.settrace(stdoutToServer=True, stderrToServer=True)でブレークして取得できるようになる。
# 				return False
# 		return True


	
def controlCreator(ctx, smgr, dialog):  # コントロールを追加する関数を返す。
	dialogmodel = dialog.getModel()  # ダイアログモデルを取得。
	def addControl(controltype, props, attrs=None):  # props: コントロールモデルのプロパティ、attr: コントロールの属性。
		controlmodel = _createControlModel(controltype, props)  # コントロールモデルの生成。
		dialogmodel.insertByName(props["Name"], controlmodel)  # ダイアログモデルにモデルを追加するだけでコントロールも作成される。
		control = dialog.getControl(props["Name"])  # コントロールコンテナに追加された後のコントロールを取得。
		if attrs is not None:  # Dialogに追加したあとでないと各コントロールへの属性は追加できない。
			for key, val in attrs.items():  # メソッドの引数がないときはvalをNoneにしている。
				if val is None:
					getattr(control, key)()
				else:
					getattr(control, key)(val)
		return control
	def _createControlModel(controltype, props):  # コントロールモデルの生成。
		if not "Name" in props:
			props["Name"] = _generateSequentialName(controltype)  # Nameがpropsになければ通し番号名を生成。
		controlmodel = dialogmodel.createInstance("com.sun.star.awt.UnoControl{}Model".format(controltype))  # コントロールモデルを生成。UnoControlDialogElementサービスのためにUnoControlDialogModelからの作成が必要。
		if props:
			values = props.values()  # プロパティの値がタプルの時にsetProperties()でエラーが出るのでその対応が必要。
			if any(map(isinstance, values, [tuple]*len(values))):
				[setattr(controlmodel, key, val) for key, val in props.items()]  # valはリストでもタプルでも対応可能。XMultiPropertySetのsetPropertyValues()では[]anyと判断されてタプルも使えない。
			else:
				controlmodel.setPropertyValues(tuple(props.keys()), tuple(values))
		return controlmodel
	def _generateSequentialName(controltype):  # 連番名の作成。
		i = 1
		flg = True
		while flg:
			name = "{}{}".format(controltype, i)
			flg = dialog.getControl(name)  # 同名のコントロールの有無を判断。
			i += 1
		return name
	return addControl  # コントロールコンテナとそのコントロールコンテナにコントロールを追加する関数を返す。
