#!/opt/libreoffice5.4/program/python
# -*- coding: utf-8 -*-
from com.sun.star.container import NoSuchElementException
from com.sun.star.uno.TypeClass import SERVICE, INTERFACE, PROPERTY, INTERFACE_METHOD, INTERFACE_ATTRIBUTE
# from .common import enableRemoteDebugging  # デバッグ用デコレーター
# @enableRemoteDebugging
def getAttrbs(args, obj):
	ctx, css, outputs= args  # outputs:結果を入れるリスト。
	stack = []  # スタックを初期化。
	st_oms = set()  # サービス名を入れる集合。
	st_omi = set()  # インターフェイス名を入れる集合。
	st_si = set()  #  最初のサポートインターフェイス名の集合。
	st_nontyps = set()  # TypeDescriptionオブジェクトを取得できないサービスの集合。
	tdm = ctx.getByName('/single tons/com.sun.star.reflection.theTypeDescriptionManager')  # TypeDescriptionManagerをシングルトンでインスタンス化。
	def _idl_check(idl): # IDL名からTypeDescriptionオブジェクトを取得。
		try:
			return tdm.getByHierarchicalName(idl)  # IDL名からTypeDescriptionオブジェクトを取得。
		except NoSuchElementException:  # TypeDescriptionを取得できないIDLのとき
			st_nontyps.add(idl)  # TypeDescriptionオブジェクトを取得できないサービスの集合に追加。
			return None
	if isinstance(obj, str): # objが文字列(IDL名)のとき
		idl = "{}{}".format(css, obj) if obj.startswith(".") else obj  # 先頭が.で始まっているときcom.sun.starが省略されていると考えて、com.sun.starを追加する。
		j = _idl_check(idl)  # IDL名からTypeDescriptionオブジェクトを取得。
		if j is None:  # TypeDescriptionオブジェクトを取得できなかったとき。
			return
		typcls = j.getTypeClass()  # jのタイプクラスを取得。
		stack = [j]  # TypeDescriptionオブジェクトをスタックに取得			
		if typcls == INTERFACE:  # インターフェイスの時
			st_omi.add(idl)  # インターフェイス名の集合に追加する。
		elif typcls == SERVICE:  # サービスの時
			st_oms.add(idl)  # サービス名の集合に追加する。			
		else:  # サービスかインターフェイス以外のときは未対応。
			return
	else:  # objが文字列以外の時
		if hasattr(obj, "getSupportedServiceNames"):  # オブジェクトがサービスを持っているとき。
			supportedservicenames = obj.getSupportedServiceNames()
			incorrectidl = "com.sun.star.AccessibleSpreadsheetDocumentView"
			if incorrectidl in supportedservicenames:  # 間違ったIDL名が存在する時。
				supportedservicenames = set(supportedservicenames)  # タプルを集合に変換。
				supportedservicenames.remove(incorrectidl)  # 間違ったIDL名を除く。
				supportedservicenames.add("com.sun.star.sheet.AccessibleSpreadsheetDocumentView")  # 正しいIDL名を追加する。
			st_ss = set(i for i in supportedservicenames if _idl_check(i))  # サポートサービス名一覧からTypeDescriptionオブジェクトを取得できないサービス名を除いた集合を得る。	
			st_sups = set()  # 親サービスを入れる集合。
			if len(st_ss) > 1:  # サポートしているサービス名が複数ある場合。
				stack = [tdm.getByHierarchicalName(i) for i in st_ss]  # サポートサービスのTypeDescriptionオブジェクトをスタックに取得。
				while stack:  # スタックがある間実行。
					j = stack.pop()  # サービスのTypeDescriptionオブジェクトを取得。
					t_std = j.getMandatoryServices() + j.getOptionalServices()  # 親サービスのタプルを取得。
					lst_std = [i for i in t_std if not i.Name in st_sups]  # 親サービスのTypeDescriptionオブジェクトのうち既に取得した親サービスにないものだけを取得。
					stack.extend(lst_std)  # スタックに新たなサービスのTypeDescriptionオブジェクトのみ追加。
					st_sups.update(i.Name for i in lst_std)  # 既に取得した親サービス名の集合型に新たに取得したサービス名を追加。
			st_ss.difference_update(st_sups)  # オブジェクトのサポートサービスのうち親サービスにないものだけにする=これがサービスの末裔。
			st_ss.difference_update(st_oms)  # 出力を抑制するサービス名を除く。
			if st_ss:  # サービス名があるとき
				lst_ss = sorted(st_ss, reverse=True)  # サービス名の集合を降順のリストにする。
				stack = [tdm.getByHierarchicalName(i) for i in lst_ss]  # TypeDescriptionオブジェクトに変換してスタックに取得。
				st_oms.update(st_ss)  # スタックに追加しないサービス名に追加する。	
		if hasattr(obj, "getTypes"):  # サービスを介さないインターフェイスがある場合。elifにしてはいけない。
			types = obj.getTypes()  # インターフェイス名ではなくtype型が返ってくる。
			if types:  # type型が取得できたとき。
				st_si = set(i.typeName for i in types).difference(st_omi)  # サポートインターフェイス名を集合型で取得して、除外するインターフェイス名を除く。
				if not stack:  # サポートするサービスがないとき
					if st_si:  # サポートするインターフェイスがあるとき
						stack = [tdm.getByHierarchicalName(i) for i in sorted(st_si, reverse=True)]  # 降順にしてTypeDescriptionオブジェクトに変換してスタックに取得。
						st_omi.update(st_si)  # すでにでてきたインターフェイス名をst_omiに追加して次は使わないようにする。
	if stack:  # 起点となるサービスかインターフェイスがあるとき。
		args = st_omi, st_oms, stack, st_si, tdm, st_nontyps, obj, outputs
		generateOutputs(args)
def generateOutputs(args):  # 末裔から祖先を得てサービス名、インターフェイス名、プロパティ名を取得する。
	st_omi, st_oms, stack, st_si, tdm, st_nontyps, obj, outputs = args
	st_omp = set()  # プロパティ名を入れる集合。
	t_itd = tuple()  # インターフェイスのTypeDescriptionオブジェクトの入れ物を初期化。
	t_spd = tuple()  # サービス属性のTypeDescriptionオブジェクトの入れ物を初期化。
	def _consumeStack(stack):  	
		def _stack_interface(t_itd):  # インターフェイスをスタックに追加する。
			lst_itd = [i for i in t_itd if not i.Name in st_omi]  # st_omiを除く。
			stack.extend(sorted(lst_itd, key=lambda x: x.Name, reverse=True))  # 降順にしてスタックに追加。
			st_omi.update(i.Name for i in lst_itd)  # すでにでてきたインターフェイス名をst_omiに追加して次は使わないようにする。
			t_itd = tuple()  # インターフェイスのTypeDescriptionオブジェクトの入れ物を初期化。	
		while stack:  # スタックがある間実行。
			j = stack.pop()  # スタックからTypeDescriptionオブジェクトをpop。
			typcls = j.getTypeClass()  # jのタイプクラスを取得。
			if typcls==INTERFACE:  # インターフェイスのとき。XInterfaceTypeDescription2インターフェイスをもつTypeDescriptionオブジェクト。
				t_itd = j.getBaseTypes() + j.getOptionalBaseTypes()  # 親インターフェイスを取得。	
				if t_itd:  # 親インターフェイスがあるとき。(TypeDescriptionオブジェクト)
					_stack_interface(t_itd)
			elif typcls==SERVICE:  # jがサービスのときtdはXServiceTypeDescriptionインターフェイスをもつ。
				t_std = j.getMandatoryServices() + j.getOptionalServices()  # 親サービスを取得。
				lst_std = [i for i in t_std if not i.Name in st_oms]  # st_ssを除く。
				stack.extend(sorted(lst_std, key=lambda x: x.Name, reverse=True))  # 親サービス名で降順に並べてサービスのTypeDescriptionオブジェクトをスタックに追加。
				st_oms.update(i.Name for i in lst_std)  # すでにでてきたサービス名をst_ssに追加して次は使わないようにする。
				itd = j.getInterface()  # new-styleサービスのインターフェイスを取得。TypeDescriptionオブジェクト。
				if itd:  # new-styleサービスのインターフェイスがあるとき。
					t_itd = itd,  # XInterfaceTypeDescription2インターフェイスをもつTypeDescriptionオブジェクト。
				else:  # new-styleサービスのインターフェイスがないときはold-styleサービスのインターフェイスを取得。
					t_itd = j.getMandatoryInterfaces() + j.getOptionalInterfaces()  # XInterfaceTypeDescriptionインターフェイスをもつTypeDescriptionオブジェクト。
				if t_itd:  # 親インターフェイスがあるとき。(TypeDescriptionオブジェクト)
					_stack_interface(t_itd)
				t_spd = j.Properties  # サービスからXPropertyTypeDescriptionインターフェイスをもつオブジェクトのタプルを取得。
				if t_spd:  # プロパティがあるとき。
					st_omp.add([k.Name for k in t_spd ])  # プロパティ名の集合に追加。					
	_consumeStack(stack)
	# サービスを介さないインターフェイスの出力。
	if isinstance(st_si, set):  # st_siが集合のとき、それはオブジェクトから得たインターフェイス名。
		st_rem = st_si.difference(st_omi)  # まだでてきていないインターフェイスがサービスを介さないインターフェイス。
		if st_rem:  # まだ出力していないインターフェイスが残っているとき
			stack =  [tdm.getByHierarchicalName(i) for i in sorted(st_rem, reverse=True)]  # 降順にしてTypeDescriptionオブジェクトに変換してスタックに取得。CSSが必要。
			st_omi.update(st_rem)  # すでにでてきたインターフェイス名をst_omiに追加して次は使わないようにする。
			_consumeStack(stack)	
	# サービスを介さないプロパティの出力。IDLにでてこないプロパティ?IDLにないサービス名の出力時の枝の決定のために存在だけまず確認する。
	if hasattr(obj, "getPropertySetInfo"):	# objにgetPropertySetInfoがあるとき。サービスに属さないプロパティを出力。
		properties = obj.getPropertySetInfo().getProperties()  # オブジェクトのプロパティを取得。すべてのプロパティのProperty Structのタプルが返ってくるので集合にする。
		st_omp.extend([p.Name for p in properties])
	st_oms.extend(st_nontyps)
	outputs.extend((st_oms, st_omi, st_omp))
