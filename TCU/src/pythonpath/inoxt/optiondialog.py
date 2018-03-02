#!/opt/libreoffice5.4/program/python
# -*- coding: utf-8 -*-
import os, unohelper
from com.sun.star.beans import PropertyValue
from com.sun.star.style.VerticalAlignment import MIDDLE, BOTTOM
from com.sun.star.awt import XActionListener, XMouseListener
from com.sun.star.ui.dialogs.ExecutableDialogResults import OK
from .common import localization
# from .common import enableRemoteDebugging  # デバッグ用デコレーター
def dilaogHandler(consts, dialog, eventname):
	ctx, smgr, configurationprovider, css, properties, nodepath, simplefileaccess = consts
	global _  # グローバルな_を地域化関数に置換する。
	_ = localization(configurationprovider)  # グローバルな_を地域化関数に置換。
	node = PropertyValue(Name="nodepath", Value="{}OptionDialog".format(nodepath))
	root = configurationprovider.createInstanceWithArguments("com.sun.star.configuration.ConfigurationAccess", (node,))
	configs = root.getPropertyValues(properties)  # コンポーネントデータノードから値を取得する。	
	state, refurl, path, idlsedit = toDialog(ctx, smgr, css, simplefileaccess, configs)  # ダイアログ用データの取得。pathはシステムパス。
	if eventname=="initialize":  # オプションダイアログがアクティブになった時
		actionlistener = ActionListener(dialog, consts)
		addControl = controlCreator(ctx, smgr, dialog)  # オプションダイアログdialogにコントロールを追加する関数を取得。
		addControl("FixedLine", {"PositionX": 5, "PositionY": 6, "Width": 268, "Height": 10, "Label": _("TCU - Tree Command for UNO")})  # 文字付き水平線。
		addControl("FixedText", {"PositionX": 5, "PositionY": 22, "Width": 83, "Height": 15, "Label": _("API Reference URL: https://"), "NoLabel": True, "VerticalAlign": MIDDLE, "Align": 2})
		addControl("Edit", {"Name": "RefUrl", "PositionX": 89, "PositionY": 22, "Width": 184, "Height": 15, "Text": refurl})
		addControl("FixedHyperlink", {"PositionX": 89, "PositionY": 38, "Width": 184, "Height": 10, "Label": _("Jump to this URL"), "TextColor": 0x3D578C, "Align": 2}, {"addMouseListener": MouseListener(dialog)})  # ActionListenerをつけるとリンクが開かない。
		addControl("GroupBox", {"PositionX": 5, "PositionY": 47, "Width": 268, "Height": 58, "Label": "Local Reference"})   
		addControl("CheckBox", {"Name": "OffLine", "PositionX": 11, "PositionY": 56, "Width": 258, "Height": 8, "Label": _("~Use Local Reference"), "State": state, "Enabled": False}) 
		addControl("FixedText", {"PositionX": 11, "PositionY": 67, "Width": 258, "Height": 15, "Label": _("Local Reference Path:"), "NoLabel": True, "VerticalAlign": BOTTOM}) 
		addControl("FixedText", {"Name": "RefDir" , "PositionX": 11, "PositionY": 82, "Width": 230, "Height": 15, "Label": path, "NoLabel": True, "VerticalAlign": BOTTOM})  
		addControl("Button", {"PositionX": 242, "PositionY": 82, "Width": 27, "Height": 15, "Label": _("~Browse")}, {"setActionCommand": "folderpicker", "addActionListener": actionlistener})
		addControl("FixedText", {"PositionX": 11, "PositionY": 107, "Width": 258, "Height": 15, "Label": _("Ignored Interfaces (Python regex patterns can be used, excluding dots):"), "NoLabel": True, "VerticalAlign": MIDDLE}) 
		addControl("Edit", {"Name": "IgnoredIdls", "PositionX": 5, "PositionY": 124, "Width": 268, "Height": 96, "MultiLine": True, "Text": idlsedit})  
		addControl("Button", {"PositionX": 218, "PositionY": 224, "Width": 55, "Height": 15, "Label": _("~Restore Defaults")}, {"setActionCommand": "restore","addActionListener": actionlistener})
		if os.path.exists(path):
			dialog.getControl("OffLine").setEnable(True)
	elif eventname=="ok":  # OKボタンが押された時
		state = dialog.getControl("OffLine").getState()
		refurl = dialog.getControl("RefUrl").getText()
		path = dialog.getControl("RefDir").getText()  # システムパスが返ってくる。実存は入力時に確認済。
		idlsedit = dialog.getControl("IgnoredIdls").getText()
		offline = True if state==1 else False
		refdir = unohelper.systemPathToFileUrl(path)  # システムパスをfileurlに変換する。
		idlstext = "".join(idlsedit.split()).replace(css, "")  # 空白、タブ、改行とcom.sun.starを除去。
		configs = offline, refurl, refdir, idlstext  # コンポーネントデータノード用の値を取得。
		node = PropertyValue(Name="nodepath", Value="{}OptionDialog".format(nodepath))
		root = configurationprovider.createInstanceWithArguments("com.sun.star.configuration.ConfigurationUpdateAccess", (node,))
		root.setPropertyValues(properties, configs)  # コンポーネントデータノードに値を代入。
		root.commitChanges()  # registrymodifications.xcuに変更後の値を書き込む。
	elif eventname=="back":  # 元に戻すボタンが押された時
		toControls(dialog, (state, refurl, path, idlsedit))  # 各コントロールに値を入力する。		
def toDialog(ctx, smgr, css, simplefileaccess, configs):  # ダイアログ向けにデータを変換する。
	offline, refurl, refdir, idlstext = configs  # コンポーネントデータノード用の値を取得。	
	state = 1 if offline else 0
	pathsubstservice = smgr.createInstanceWithContext("com.sun.star.comp.framework.PathSubstitution", ctx)
	fileurl = pathsubstservice.substituteVariables(refdir, True)  # $(inst)を変換する。fileurlが返ってくる。
	path = os.path.normpath(unohelper.fileUrlToSystemPath(fileurl)) if simplefileaccess.exists(fileurl) else "Local API Reference does not exists."  # fileurlをシステムパスに変換する。パスの実存を確認する。
	idls = "".join(idlstext.split()).split(",")  # xmlがフォーマットされていると空白やタブが入ってくるのでそれを除去してリストにする。
	idlsedit = ", ".join("{}{}".format(css, i) if i.startswith(".") else i for i in idls)	
	return state, refurl, path, idlsedit
def toControls(dialog, configs):  # 各コントロールに値を入力する。		
	state, refurl, path, idlsedit = configs  # ダイアログ用データの取得。
	dialog.getControl("RefUrl").setText(refurl)
	checkbox = dialog.getControl("OffLine")
	checkbox.setState(state)
	if os.path.exists(path):
		checkbox.setEnable(True)	
	dialog.getControl("RefDir").setText(path)
	dialog.getControl("IgnoredIdls").setText(idlsedit)			
class ActionListener(unohelper.Base, XActionListener):
	def __init__(self, dialog, consts):
		self.dialog = dialog
		self.consts = consts
	def actionPerformed(self, actionevent):	
		cmd = actionevent.ActionCommand
		ctx, smgr, configurationprovider, css, properties, nodepath, simplefileaccess = self.consts
		dialog = self.dialog
		if cmd=="folderpicker":
			fixedtext = dialog.getControl("RefDir")
			path = fixedtext.getText()  # システムパスが返ってくる。
			folderpicker = smgr.createInstanceWithContext("com.sun.star.ui.dialogs.FolderPicker", ctx)
			if os.path.exists(path):  # pathが存在するとき
				fileurl = unohelper.systemPathToFileUrl(path)  # システムパスをfileurlに変換する。
				folderpicker.setDisplayDirectory(fileurl)  # フォルダ選択ダイアログに設定する。
			folderpicker.setTitle(_("Select ref folder"))
			if folderpicker.execute()==OK:
				fileurl = folderpicker.getDirectory()
				checkbox = dialog.getControl("OffLine")
				if simplefileaccess.exists(fileurl):
					path = unohelper.fileUrlToSystemPath(fileurl)  # fileurlをシステムパスに変換する。
					checkbox.setEnable(True)
				else:
					path = "Local API Reference does not exists."
					checkbox.setEnable(False)
				fixedtext.setText(path)
		elif cmd=="restore":
			node = PropertyValue(Name="nodepath", Value="{}Defaults".format(nodepath))
			root = configurationprovider.createInstanceWithArguments("com.sun.star.configuration.ConfigurationAccess", (node,))
			configs = root.getPropertyValues(properties)  # コンポーネントデータノードからデフォルト値を取得する。	
			toControls(dialog, toDialog(ctx, smgr, css, simplefileaccess, configs))  # 各コントロールに値を入力する。		
	def disposing(self, eventobject):
		pass	
class MouseListener(unohelper.Base, XMouseListener):
	def __init__(self, dialog):
		self.dialog = dialog
	def mousePressed(self, mouseevent):
		pass			
	def mouseReleased(self, mouseevent):
		pass
	def mouseEntered(self, mouseevent):
		control, dummy_controlmodel, name = eventSource(mouseevent)
		if name == "FixedHyperlink1":
			refurl = self.dialog.getControl("RefUrl").getText()
			control.setURL("https://{}".format(refurl))
	def mouseExited(self, mouseevent):
		pass
	def disposing(self, eventobject):
		pass
def eventSource(event):  # イベントからコントロール、コントロールモデル、コントロール名を取得。
	control = event.Source  # イベントを駆動したコントロールを取得。
	controlmodel = control.getModel()  # コントロールモデルを取得。
	name = controlmodel.getPropertyValue("Name")  # コントロール名を取得。	
	return control, controlmodel, name	
def controlCreator(ctx, smgr, dialog):  # コントロールを追加する関数を返す。
	dialogmodel = dialog.getModel()  # ダイアログモデルを取得。
	def addControl(controltype, props, attrs=None):  # props: コントロールモデルのプロパティ、attr: コントロールの属性。
		controlmodel = _createControlModel(controltype, props)  # コントロールモデルの生成。
		dialogmodel.insertByName(props["Name"], controlmodel)  # ダイアログモデルにモデルを追加するだけでコントロールも作成される。
		if attrs is not None:  # Dialogに追加したあとでないと各コントロールへの属性は追加できない。
			control = dialog.getControl(props["Name"])  # コントロールコンテナに追加された後のコントロールを取得。
			for key, val in attrs.items():  # メソッドの引数がないときはvalをNoneにしている。
				if val is None:
					getattr(control, key)()
				else:
					getattr(control, key)(val)
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
