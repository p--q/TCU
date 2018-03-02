#!/opt/libreoffice5.4/program/python
# -*- coding: utf-8 -*-
import re
from .common import localization
from com.sun.star.container import NoSuchElementException
from com.sun.star.uno.TypeClass import SERVICE, INTERFACE, PROPERTY, INTERFACE_METHOD, INTERFACE_ATTRIBUTE
# from .common import enableRemoteDebugging  # デバッグ用デコレーター import pydevd; pydevd.settrace(stdoutToServer=True, stderrToServer=True)
# @enableRemoteDebugging
def wCompare(args, obj1, obj2):
	ctx, configurationprovider, css, fns, st_omi, outputs = args  # st_omi: スタックに追加しないインターフェイス名の集合。
	tdm = ctx.getByName('/singletons/com.sun.star.reflection.theTypeDescriptionManager')  # TypeDescriptionManagerをシングルトンでインスタンス化。
	global _  # グローバルな_を置換する。
	_ = localization(configurationprovider)  # 地域化関数に置換。
	reg_p = re.compile(r'[^_a-zA-Z0-9\.]')  # アンダースコア、ドット、数字、文字以外のものを含んでいるものは正規表現パターンと判断する。
	patterns = set(i for i in st_omi if reg_p.search(i))  # 正規表現のパターンになっている要素のみをst_omiから集合に取得。
	st_omi.difference_update(patterns)  # st_omiから正規表現のパターンを除く。
	args = outputs, tdm, css
	ss_obj1, nontdm_obj1, is_obj1, ps_obj1 = getAttrbs(args, obj1)  # obj1のサービス名、TypeDescriptionオブジェクトがないサービス名、インターフェイス名、プロパティのみProperty Struct、の集合。
	is_obj1.difference_update(st_omi)  # 出力しないインターフェイス名を除いておく。
	[st_omi.update(filter(lambda x: re.search(p, x), is_obj1)) for p in patterns]  # オブジェックトのインターフェイス名の集合から正規表現のパターンに一致する名前をst_omiに追加する。
	is_obj1.difference_update(st_omi)  # 正規表現で取得した出力しないインターフェイス名を除いておく。
	if obj2 is None:  # obj2がないときは比較しない。
		treeCreator(tdm, css, fns, outputs, st_omi)(ss_obj1, nontdm_obj1, is_obj1, ps_obj1)  # obj1のサービスとインターフェイスのみ出力する。			
	else:  # obj2があるときはobj1と比較する。	
		ss_obj2, nontdm_obj2, is_obj2, ps_obj2 = getAttrbs(args, obj2)  # obj2のサービス名、TypeDescriptionオブジェクトがないサービス名、インターフェイス名、プロパティのみProperty Struct、の集合。
		is_obj2.difference_update(st_omi)  # 出力しないインターフェイス名を除いておく。
		[st_omi.update(filter(lambda x: re.search(p, x), is_obj2)) for p in patterns]  # オブジェックトのインターフェイス名の集合から正規表現のパターンに一致する名前をst_omiに追加する。
		createTree = treeCreator(tdm, css, fns, outputs, st_omi)  # createTreeを取得。
		is_obj2.difference_update(st_omi)  # 正規表現で取得した出力しないインターフェイス名を除いておく。
		ps_obj1name = set(i.Name for i in ps_obj1)  # プロパティだけProperty Structなので名前の集合を求めておく。
		ps_obj2name = set(i.Name for i in ps_obj2)
		outputs.append(_("Services and interface common to object1 and object2."))  # object1とobject2に共通するサービスとイターフェイス一覧。
		st_s = ss_obj1 & ss_obj2  # 共通するサービス名。
		st_non = nontdm_obj1 & nontdm_obj2  # 共通するnontdmサービス名。
		st_i = is_obj1 & is_obj2  # 共通するインターフェイス名。
		st_pname = ps_obj1name & ps_obj2name  # 共通するプロパティ名。	
		st_p = [i for i in ps_obj1 if i.Name in st_pname]
		omis = createTree(st_s, st_non, st_i, st_p)  # 共通するサービスとインターフェイスを出力する。出力したサービス名、インターフェイス名、プロパティ名が返る。
		outputs.append("")	
		outputs.append(_("Services and interfaces that only object1 has."))  # object1だけがもつサービスとインターフェイス一覧。
		st_s = ss_obj1 - ss_obj2
		st_non = nontdm_obj1 - nontdm_obj2
		st_i = is_obj1 - is_obj2
		st_pname = ps_obj1name - ps_obj2name
		st_p = [i for i in ps_obj1 if i.Name in st_pname]
		omis = createTree(st_s, st_non, st_i, st_p, omis=omis)  # obj1のみのサービスとインターフェイスを出力する。すでに出力したサービス名、インターフェイス名、プロパティ名を渡して抑制する。
		outputs.append("")	
		outputs.append(_("Services and interfaces that only object2 has."))  # object2だけがもつサービスとインターフェイス一覧。
		st_s = ss_obj2 - ss_obj1
		st_non = nontdm_obj2 - nontdm_obj1
		st_i = is_obj2 - is_obj1
		st_pname = ps_obj2name - ps_obj1name
		st_p = [i for i in ps_obj2 if i.Name in st_pname]
		createTree(st_s, st_non, st_i, st_p, omis=omis)  # obj2のみのサービスとインターフェイスを出力する。	すでに出力したサービス名、インターフェイス名、プロパティ名を渡して抑制する。	
def getAttrbs(args, obj):
	outputs, tdm, css = args
	st_ss, st_nontdm, st_is, st_ps = [set() for i in range(4)]  # サービス名、TypeDescriptionオブジェクトを取得できないサービス名、インターフェイス名を入れる集合、プロパティのみProperty Structを返す。
	if isinstance(obj, str): # objが文字列(IDL名)のとき
		idl = "{}{}".format(css, obj) if obj.startswith(".") else obj  # 先頭が.で始まっているときcom.sun.starが省略されていると考えて、com.sun.starを追加する。
		if tdm.hasByHierarchicalName(idl):  #  TypeDescriptionが取得できるとき。
			j = tdm.getByHierarchicalName(idl)  # IDL名からTypeDescriptionオブジェクトを取得。
			typcls = j.getTypeClass()  # jのタイプクラスを取得。
			if typcls == SERVICE:  # サービスの時
				st_ss.add(j.Name)
				args = st_ss, st_is, [j]
				getSuperService(args)
			elif typcls == INTERFACE:  # インターフェイスの時
				st_is.add(j.Name)
				getSuperInterface(st_is, [j])	
			else:  # サービスかインターフェイス以外のときは未対応。
				outputs.append(_("{} is not a service name or an interface name, so it is not supported yet.".format(idl)))  # はサービス名またはインターフェイス名ではないので未対応です。
		else:  # TypeDescriptionを取得できないIDL名のとき
			st_nontdm.add(idl)  # TypeDescriptionオブジェクトを取得できないサービスの集合に追加。		
	else:  # objが文字列以外の時
		if hasattr(obj, "getSupportedServiceNames"):  # オブジェクトがサービスを持っているとき。
			st_ss.update(obj.getSupportedServiceNames())  # サポートサービス名を集合にして取得。
			incorrectidls = {"com.sun.star.AccessibleSpreadsheetDocumentView": "com.sun.star.sheet.AccessibleSpreadsheetDocumentView"}  # 間違って返ってくるIDL名の辞書。
			st_incorrects = st_ss.intersection(incorrectidls.keys())  # 間違っているIDL名の集合を取得。
			for incorrect in st_incorrects:
				st_ss.remove(incorrect)  # 間違ったIDL名を除く。
				st_ss.add(incorrectidls[incorrect])  # 正しいIDL名を追加する。
			if st_ss:
				st_ss = set(i if tdm.hasByHierarchicalName(i) else st_nontdm.add(i) for i in st_ss)  # TypeDescriptionオブジェクトが取得できないサービス名を削除する。
				st_ss.discard(None)  # tdm.hasByHierarchicalName(i)がFalseのときに入ってくるNoneを削除する。remove()では要素がないときにエラーになる。
				args = st_ss, st_is, [tdm.getByHierarchicalName(i) for i in st_ss]
				getSuperService(args)  # TypeDescriptionオブジェクトに変換して渡す。	
		if hasattr(obj, "getTypes"):  # サービスを介さないインターフェイスがある場合。elifにしてはいけない。
			types = obj.getTypes()  # インターフェイス名ではなくtype型が返ってくる。
			if types:  # type型が取得できたとき。
				typenames = [i.typeName for i in types]  #  # types型をインターフェイス名のリストに変換。
				st_is.update(typenames)
				getSuperInterface(st_is, [tdm.getByHierarchicalName(i) for i in typenames]) 
		if hasattr(obj, "getPropertySetInfo"):	# objにgetPropertySetInfoがあるとき。サービスに属さないプロパティを出力。
			properties = obj.getPropertySetInfo().getProperties()  # オブジェクトのプロパティを取得。すべてのプロパティのProperty Structのタプルが返ってくるので集合にする。
			st_ps.update(properties)
		if not any([st_ss, st_nontdm, st_is, st_ps]):
			outputs.append(_("There is no service or interface to support."))  # サポートするサービスやインターフェイスがありません。
	return st_ss, st_nontdm, st_is, st_ps  # プロパティのみProperty Structを返す。
def getSuperService(args):  # 再帰的にサービスのスーパークラスとインターフェイス名を取得する。XServiceTypeDescription2インターフェイスをもつTypeDescriptionオブジェクト
	st_ss, st_is, tdms = args
	for j in tdms:  # 各サービスのTypeDescriptionオブジェクトについて。
		if j.Name in st_ss:  # すでに追加したサービス名のときは次のループにいく。
			continue
		else:
			lst_itd = []  # サービスがもっているインターフェイスを入れるリストを初期化。
			if j.isSingleInterfaceBased():  # new-styleサービスのとき。
				lst_itd.append(j.getInterface())  # new-styleサービスのインターフェイスを取得。TypeDescriptionオブジェクト。
			else:  # old-styleサービスのときはスーパークラスのサービスがありうる。
				lst_std = list(j.getMandatoryServices())
				lst_std.extend(j.getOptionalServices())
				if lst_std:
					st_ss.update(i.Name for i in lst_std)  # スーパークラスのサービス名を取得。
					args = st_ss, st_is, lst_std
					getSuperService(args)  # サービスのスーパークラスについて。 
				lst_itd.extend(j.getMandatoryInterfaces())  # old-styleサービスのインターフェイスを取得。
				lst_itd.extend(j.getOptionalInterfaces())  # old-styleサービスのインターフェイスを取得。			
			if lst_itd:  # スーパークラスのインターフェイスがあるとき。
				st_is.update(i.Name for i in lst_itd)  # スーパークラスのインターフェイス名を取得。
				getSuperInterface(st_is, lst_itd)  # インターフェイスのスーパークラスについて。	
def getSuperInterface(st_is, tdms):  # 再帰的にインターフェイスのスーパークラスを取得する。XInterfaceTypeDescription2インターフェイスをもつTypeDescriptionオブジェクト。
	for j in tdms:
		if j.Name in st_is:  # すでに追加したインターフェイス名のときは次のループにいく。
			continue
		else:
			lst_itd = list(j.getBaseTypes())
			lst_itd.extend(j.getOptionalBaseTypes())
			if lst_itd:  # スーパークラスのインターフェイスがあるとき。
				st_is.update(i.Name for i in lst_itd)  # スーパークラスのインターフェイス名を取得。
				getSuperInterface(st_is, lst_itd)
def treeCreator(tdm, css, fns, outputs, omi):				
	def createTree(st_s, st_non, st_i, st_p, *, omis=None):  # st_pの要素はプロパティ名ではなくProperty Struct。omisは出力を抑制する名前のタプルのタプル。
		if omis is None:
			st_oms, st_omi, st_omp = set(), omi.copy(), set()  # すでに取得したサービス名、プロパティ名を入れる集合。
		else:
			st_oms, st_omi, st_omp = omis  
			st_omi.update(omi)  # デフォルトの出力抑制インターフェイスを追加する。
		indent = "	  "  # インデントを設定。
		consumeStack = createStackConsumer(indent, css, fns, st_oms, st_omi, st_omp)  # クロージャーに値を渡す。
		non_ss = selectNonSupers(tdm, st_s.copy(), getSuperServices)  # ツリーでスーパークラスが先にでてこないようにスーパークラスにあたる名前を削除する。
		if non_ss:  
			stack = [tdm.getByHierarchicalName(i) for i in sorted(non_ss, reverse=True)]  # サービス名を降順にしてTypeDescriptionオブジェクトをスタックに取得。		
			st_oms.update(non_ss)  # すでに取得したサービス名の集合に追加。
		else:  # サービス名がないとき。
			non_si = selectNonSupers(tdm, st_i.copy(), getSuperInterfaces)  # ツリーでスーパークラスが先にでてこないようにスーパークラスにあたる名前を削除する。 
			if non_si:  # インターフェイスがあるとき。
				stack = [tdm.getByHierarchicalName(i) for i in sorted(non_si, reverse=True)]  # 降順にしてTypeDescriptionオブジェクトに変換してスタックに取得。
				st_omi.update(non_si)  # すでに取得したインターフェイス名の集合に追加。
		consumeStack(stack)  # ツリーを出力。
		st_s.difference_update(st_oms)  # すでに出力されたサービス名を除く。比較のときは出力抑制されたスーパークラスのサービス名がでてくる。
		if st_s:  # まだ出力されていないサービスが残っているとき。
			non_ss = selectNonSupers(tdm, st_s.copy(), getSuperServices)  # ツリーでスーパークラスが先にでてこないようにスーパークラスにあたる名前を削除する。
			if non_ss:  
				stack = [tdm.getByHierarchicalName(i) for i in sorted(non_ss, reverse=True)]  # サービス名を降順にしてTypeDescriptionオブジェクトをスタックに取得。		
				st_oms.update(non_ss) # すでに取得したサービス名の集合に追加。
				consumeStack(stack)  # ツリーを出力。
		st_i.difference_update(st_omi)  # すでに出力されたインターフェイス名を除く。サービスがあるのに、サービスを介さないインターフェイス名があるときにそれが残る。
		if st_i:  # まだ出力されていないインターフェイスが残っているとき。
			non_si = selectNonSupers(tdm, st_i.copy(), getSuperInterfaces)  # ツリーでスーパークラスが先にでてこないようにスーパークラスにあたる名前を削除する。
			if non_si:  # インターフェイスがあるとき。
				stack = [tdm.getByHierarchicalName(i) for i in sorted(non_si, reverse=True)]  # 降順にしてTypeDescriptionオブジェクトに変換してスタックに取得。
				st_omi.update(non_si)  # すでに取得したインターフェイス名の集合に追加。
				consumeStack(stack)  # ツリーを出力。
		if st_non:  # TypeDescriptionオブジェクトを取得できないサービスを出力する。コントロール関係は実装サービス名がここにでてくる。
			for i in sorted(st_non):  # 昇順に並べて取得。
				branch = ["├─"] 
				branch.append(i.replace(css, ""))  # 一番最後のサービス名をbranchの要素に追加。
				fns["NOLINK"]("".join(branch))  # リンクをつけずに出力。	
		properties = [i for i in st_p if not i.Name in st_omp]  # すでに出力されたプロパティを除く。Property Structのリストが返る。		
		if properties:  # まだ出力していないプロパティが存在する時。
			props = sorted(properties, key=lambda x: x.Name)  #Name属性で昇順に並べる。
			m = max(len(i.Type.typeName.replace(css, "")) for i in props)  # プロパティの型のうち最大文字数を取得。
			fns["NOLINK"]("└──")  # 枝の最後なので下に枝を出さない。
			for i in props:  # 各プロパティについて。
				branch = [indent*2]  # 枝をリセット。
				branch.append("{}  {}".format(i.Type.typeName.replace(css, "").rjust(m), i.Name))  # 型は最大文字数で右寄せにする。
				fns["PROPERTY"]("".join(branch))  # プロパティの行を出力。	
		else:			
			n = len(outputs)  # 出力行数を取得。
			for j in reversed(range(n)):  # 出力行を下からみていく。
				line = outputs[j]  # 行の内容を取得。
				if "│   " in line:  # 下に続く縦棒があるとき
					outputs[j] = line.replace("│   ", indent, 1)  # j行目のline内の左端の不要な縦棒を空白に置換。
				elif "├─" in line:  # 下に続く分岐があるとき。
					outputs[j] = line.replace("├─", "└─", 1)  # 分岐を終了枝に置換してループを出る。
					break	
		return st_oms, st_omi-omi, st_omp  # 出力したサービス名、インターフェイス名、プロパティ名を返す。
	return createTree			
def selectNonSupers(tdm, st, getSupers):  # ツリーでスーパークラスが先にでてこないようにスーパークラスにあたる名前を削除する。
	s = st.copy()  # 元の集合。
	st_sup = set()  # スーパークラス名を入れる集合。
	while st:  # サービス名がなくなるまで実行。
		j = tdm.getByHierarchicalName(st.pop())  # スタックのサービス名のTypeDescriptionオブジェクトを取得。
		getSupers(st, st_sup, j)  # スーパークラスをスタックから除きながら再帰的にスーパーサービスをst_supに取得する。
	return s - st_sup	# スーパークラスのサービス名を除いた集合を返す。	
def getSuperServices(st, st_sup, j):  # スーパーサービス名を取得してstから削除する。	
	if not j.isSingleInterfaceBased():  # old-styleサービスのときはスーパークラスのサービスがありうる。
		lst_super = list(j.getMandatoryServices())  # スーパークラスのTypeDescriptionオブジェクトを取得。
		lst_super.extend(j.getOptionalServices())  # スーパークラスのTypeDescriptionオブジェクトを取得。	
		names = [i.Name for i in lst_super]
		st.difference_update(names)  # スーパーサービス名をスタックから削除する。
		st_sup.update(names)  # スーパークラス名を取得する。
		for j in lst_super:  # 各スーパークラスのTypeDescriptionオブジェクトについて。
			getSuperServices(st, st_sup, j)
def getSuperInterfaces(st, st_sup, j):  # スーパーインターフェイス名を取得してstから削除する。	
	lst_super = list(j.getBaseTypes())  # スーパークラスのTypeDescriptionオブジェクトを取得。
	lst_super.extend(j.getOptionalBaseTypes())  # スーパークラスのTypeDescriptionオブジェクトを取得。	
	names = [i.Name for i in lst_super]
	st.difference_update(names)  # スーパーサービス名をスタックから削除する。
	st_sup.update(names)  # スーパークラス名を取得する。
	for j in lst_super:  # 各スーパークラスのTypeDescriptionオブジェクトについて。
		getSuperInterfaces(st, st_sup, j)	
def createStackConsumer(indent, css, fns, st_oms, st_omi, st_omp):	
		reg_sqb = re.compile(r'\[\]')  # 型から角括弧ペアを取得する正規表現オブジェクト。
		inout_dic = {(True, False): "[in]", (False, True): "[out]", (True, True): "[inout]"}  # メソッドの引数のinout変換辞書。	
		def consumeStack(stack):  # サービスかインターフェイスのTypeDescriptionのリストを引数とする。スタックの最後から出力される。
			m = 0  # 最大文字数を初期化。
			lst_level = [1]*len(stack)  # stackの要素すべてについて枝分かれ番号を取得。1から始まる。
			def _format_type(typ):  # 属性がシークエンスのとき[]の表記を修正。
				n = len(reg_sqb.findall(typ))  # 角括弧のペアのリストの数を取得。
				for dummy in range(n):  # 角括弧の数だけ繰り返し。
					typ = typ.replace("]", "", 1) + "]" 
				return typ  # 訂正した表記を返す。			
			def _stack_interface(lst_itd):  # インターフェイスをスタックに追加する。引数はインターフェイスのTypeDescriptionオブジェクトのリスト。
				lst_itd = [i for i in lst_itd if not i.Name in st_omi]  # st_omiを除く。
				stack.extend(sorted(lst_itd, key=lambda x: x.Name, reverse=True))  # 降順にしてスタックに追加。
				lst_level.extend(level+1 for i in lst_itd)  # 枝分かれ番号を1増やして設定。
				st_omi.update(i.Name for i in lst_itd)  # インターフェイス名をst_omiに追加して次は使わないようにする。
			while stack:  # スタックがある間実行。
				j = stack.pop()  # スタックからTypeDescriptionオブジェクトをpop。
				level = lst_level.pop()  # jの枝分かれ番号を取得。
				typcls = j.getTypeClass()  # jのタイプクラスを取得。
				branch = ["", ""]  # 枝をリセット。縦枝はインデックス0に文字列として追加する。枝分かれの枝はインデックス1の要素に入れる。
				if level>1:  # 枝分かれ番号が2以上のとき。つまり左端の縦線のみのとき以外の行のときなので一番最初の枝分かれ以外のときすべてに該当。
					branch[0] = "│   "  # 左端に縦枝をつける。不要な縦枝は最後に削ることにする。
					p = 2  # 枝分かれ番号2から処理する。
					for i in range(p, level):  # 枝分かれpからみていく。
						branch[0] += "│   " if i in lst_level else indent  # スタックに同じ枝分かれ回数の要素があるときは縦枝を表示する。
				if typcls==INTERFACE or typcls==SERVICE:  # jがサービスかインターフェイスのとき。
					if level==1: 
						branch[1] = "├─"  # 枝分かれ番号1のときは常に下につづく分岐をつけて、すべての枝を出力してから修正する。
					else:
						branch[1] = "├─" if level in lst_level else "└─"  # スタックに同じ枝分かれ番号の要素があるときは├─そうでなければ└─で枝を止める 。
					if typcls==INTERFACE:  # インターフェイスのとき。XInterfaceTypeDescription2インターフェイスをもつTypeDescriptionオブジェクト。
						branch.append(j.Name.replace(css, ""))  # インターフェイス名をbranchのインデックス1の要素に追加。
						fns["INTERFACE"]("".join(branch))  # インターフェイス名の行を出力。
						lst_itd = list(j.getBaseTypes())  # スーパークラスのインターフェイスを取得。
						lst_itd.extend(j.getOptionalBaseTypes())  # スーパークラスのインターフェイスを取得。
						if lst_itd:  # インターフェイスのスーパークラスがあるとき。(要素はTypeDescriptionオブジェクト)
							_stack_interface(lst_itd)  # スタックに追加。
						t_md = j.getMembers()  # インターフェイスアトリビュートとメソッドのTypeDescriptionオブジェクトを取得。
						if t_md:  # インターフェイスアトリビュートとメソッドがあるとき。
							stack.extend(sorted(t_md, key=lambda x: x.Name, reverse=True))  # 降順にしてスタックに追加。
							lst_level.extend([level+1 for i in t_md])  # 枝分かれ番号1増やして設定。
							stringlengths = [len(i.ReturnType.Name.replace(css, "")) for i in t_md if i.getTypeClass()==INTERFACE_METHOD]  # メソッド名の長さのリスト。
							stringlengths.extend([len(i.Type.Name.replace(css, "")) for i in t_md if i.getTypeClass()==INTERFACE_ATTRIBUTE])  # アトリビュート名の長さのリストを追加。
							m = max(stringlengths)  # インターフェイスアトリビュートとメソッドの型のうち最大文字数を取得。
					elif typcls==SERVICE:  # jがサービスのときXServiceTypeDescription2インターフェイスをもつTypeDescriptionオブジェクト。
						branch.append(j.Name.replace(css, ""))  # サービス名をbranchのインデックス2の要素に追加。
						fns["SERVICE"]("".join(branch))  # サービス名の行を出力。
						lst_itd = []  # インターフェイスのTypeDescriptionオブジェクトの入れ物を初期化。	
						if j.isSingleInterfaceBased():  # new-styleサービスのとき。
							lst_itd.append(j.getInterface())  # new-styleサービスのインターフェイスを取得。TypeDescriptionオブジェクト。
						else:  # old-styleサービスのとき。
							lst_std = list(j.getMandatoryServices())  # スーパークラスのサービスを取得。
							lst_std.extend(j.getOptionalServices())  # スーパークラスのサービスを取得。
							lst_std = [i for i in lst_std if not i.Name in st_oms]  # st_omsを除く。
							if lst_std:  # サービスのスーパークラスがあるとき。(要素はTypeDescriptionオブジェクト)  
								stack.extend(sorted(lst_std, key=lambda x: x.Name, reverse=True))  # サービス名で降順に並べてサービスのTypeDescriptionオブジェクトをスタックに追加。
								lst_level.extend(level+1 for i in lst_std)  # 枝分かれ番号1増やして設定。
								st_oms.update(i.Name for i in lst_std)  # サービス名をst_omsに追加して次は使わないようにする。
							lst_itd.extend(j.getMandatoryInterfaces())  # old-styleサービスのインターフェイスを取得。
							lst_itd.extend(j.getOptionalInterfaces())  # old-styleサービスのインターフェイスを取得。					
						if lst_itd:  # インターフェイスがあるとき。(TypeDescriptionオブジェクト)
							_stack_interface(lst_itd)  # インターフェイスをスタックに追加する。
						t_spd = j.Properties  # サービスからXPropertyTypeDescriptionインターフェイスをもつオブジェクトのタプルを取得。
						if t_spd:  # プロパティがあるとき。
							stack.extend(sorted(list(t_spd), key=lambda x: x.Name, reverse=True))  # 降順にしてスタックに追加。
							lst_level.extend(level+1 for i in t_spd)  # 枝分かれ番号を取得。
							m = max(len(i.getPropertyTypeDescription().Name.replace(css, "")) for i in t_spd)  # プロパティの型のうち最大文字数を取得。
							st_omp.update(i.Name for i in t_spd) # すでに出力したプロパティ名の集合に追加。
				else:  # jがインターフェイスかサービス以外のとき。
					branch[1] = indent  # 枝分かれはしない。
					if level in lst_level:  # スタックに同じ枝分かれ番号があるとき。
						typcls2 = stack[lst_level.index(level)].getTypeClass()  # スタックにある同じ枝分かれ番号のものの先頭の要素のTypeClassを取得。
						if typcls2==INTERFACE or typcls2==SERVICE: branch[1] = "│   "  # サービスかインターフェイスのとき。横枝だったのを縦枝に書き換える。
					if typcls==INTERFACE_METHOD:  # jがメソッドのとき。
						typ = _format_type(j.ReturnType.Name.replace(css, ""))  # 戻り値の型を取得。
						stack2 = list(j.Parameters)[::-1]  # メソッドの引数について逆順(降順ではない)にスタック2に取得。
						if not stack2:  # 引数がないとき。
							branch.append("{}  {}()".format(typ.rjust(m), j.MemberName.replace(css, "")))  # 「戻り値の型(固定幅mで右寄せ) メソッド名()」をbranchの3番の要素に取得。
							fns["INTERFACE_METHOD"]("".join(branch))  # メソッドの行を出力。
						else:  # 引数があるとき。
							m3 = max(len(i.Type.Name.replace(css, "")) for i in stack2)  # 引数の型の最大文字数を取得。
							k = stack2.pop()  # 先頭の引数を取得。
							inout = inout_dic[(k.isIn(), k.isOut())]  # 引数の[in]の判定、[out]の判定
							typ2 = _format_type(k.Type.Name.replace(css, ""))  # 戻り値の型を取得。
							branch.append("{}  {}( {} {} {}".format(typ.rjust(m), j.MemberName.replace(css, ""), inout, typ2.rjust(m3), k.Name.replace(css, "")))  # 「戻り値の型(固定幅で右寄せ)  メソッド名(inout判定　引数の型(固定幅m3で左寄せ) 引数名」をbranchの3番の要素に取得。
							m2 = len("{}  {}( ".format(typ.rjust(m), j.MemberName.replace(css, "")))  # メソッドの引数の部分をインデントする文字数を取得。
							if stack2:  # 引数が複数あるとき。
								branch.append(",")  # branchの4番の要素に「,」を取得。
								fns["INTERFACE_METHOD"]("".join(branch))  # 枝をつけてメソッド名とその0番の引数の行を出力。
								del branch[2:]  # branchの2番以上の要素は破棄する。
								while stack2:  # 1番以降の引数があるとき。
									k = stack2.pop()
									inout = inout_dic[(k.isIn(), k.isOut())]  # 引数の[in]の判定、[out]の判定
									typ2 = _format_type(k.Type.Name.replace(css, ""))  # 戻り値の型を取得。
									branch.append("{}{} {} {}".format(" ".rjust(m2), inout, typ2.rjust(m3), k.Name.replace(css, "")))  # 「戻り値の型とメソッド名の固定幅m2 引数の型(固定幅m3で左寄せ) 引数名」をbranchの2番の要素に取得。
									if stack2:  # 最後の引数でないとき。
										branch.append(",")  # branchの3番の要素に「,」を取得。
										fns["INTERFACE_METHOD"]("".join(branch))  # 枝をつけて引数を出力。
										del branch[2:]  # branchの2番以上の要素は破棄する。
							t_ex = j.Exceptions  # 例外を取得。
							if t_ex:  # 例外があるとき。
								fns["INTERFACE_METHOD"]("".join(branch))  # 最後の引数を出力。
								del branch[2:]  # branchの2番以降の要素を削除。
								n = ") raises ( "  # 例外があるときに表示する文字列。
								m4 = len(n)  # nの文字数。
								stack2 = list(t_ex)  # 例外のタプルをリストに変換。
								branch.append(" ".rjust(m2-m4))  # 戻り値の型とメソッド名の固定幅m2からnの文字数を引いて2番の要素に取得。
								branch.append(n)  # nを3番要素に取得。他の要素と別にしておかないとn in branchがTrueにならない。
								while stack2:  # stack2があるとき
									k = stack2.pop()  # 例外を取得。
									branch.append(k.Name.replace(css, ""))  # branchの3番(初回以外のループでは4番)の要素に例外名を取得。
									if stack2:  # まだ次の例外があるとき。
										fns["INTERFACE_METHOD"]("{},".format("".join(branch)))  # 「,」をつけて出力。
									else:  # 最後の要素のとき。
										fns["INTERFACE_METHOD"]("{})".format("".join(branch)))  # 閉じ括弧をつけて出力。
									if n in branch:  # nが枝にあるときはnのある3番以上のbranchの要素を削除。
										del branch[3:]
										branch.append(" ".rjust(m4))  # nを削った分の固定幅をbranchの3番の要素に追加。
									else:
										del branch[4:]  # branchの4番以上の要素を削除。
							else:  # 例外がないとき。
								fns["INTERFACE_METHOD"]("{})".format("".join(branch)))  # 閉じ括弧をつけて最後の引数を出力。
					elif typcls==PROPERTY:  # プロパティのとき。
						typ = _format_type(j.getPropertyTypeDescription().Name.replace(css, ""))  # プロパティの型を取得。
						branch.append("{}  {}".format(typ.rjust(m), j.Name))  # 型は最大文字数で右寄せにする。
						fns["PROPERTY"]("".join(branch))  # プロパティの行を出力。
					elif typcls==INTERFACE_ATTRIBUTE:  # アトリビュートのとき。
						typ = _format_type(j.Type.Name.replace(css, ""))  # 戻り値の型を取得。
						branch.append("{}  {}".format(typ.rjust(m), j.MemberName.replace(css, "")))  # 型は最大文字数で右寄せにする。
						fns["INTERFACE_METHOD"]("".join(branch))  # アトリビュートの行を出力。				
		return consumeStack
