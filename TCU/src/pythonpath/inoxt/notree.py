#!/opt/libreoffice5.4/program/python
# -*- coding: utf-8 -*-
import re
from .common import localization
from com.sun.star.container import NoSuchElementException
from com.sun.star.uno.TypeClass import SERVICE, INTERFACE, PROPERTY, INTERFACE_METHOD, INTERFACE_ATTRIBUTE
# from .common import enableRemoteDebugging  # デバッグ用デコレーター
# @enableRemoteDebugging
def getAttrbs(args):
	configurationprovider, outputs, tdm, css, obj = args
	st_ss, st_nontdm, st_is, st_ps = [set() for i in range(4)]  # サービス名、TypeDescriptionオブジェクトを取得できないサービス名、インターフェイス名を入れる集合、プロパティのみProperty Structを返す。
	_ = localization(configurationprovider)  # 地域化関数に置換。
	def _idl_check(idl): # IDL名からTypeDescriptionオブジェクトを取得。
		try:
			return tdm.getByHierarchicalName(idl)  # IDL名からTypeDescriptionオブジェクトを取得して返す。
		except NoSuchElementException:  # TypeDescriptionを取得できないIDL名のとき
			st_nontdm.add(idl)  # TypeDescriptionオブジェクトを取得できないサービスの集合に追加。
	if isinstance(obj, str): # objが文字列(IDL名)のとき
		idl = "{}{}".format(css, obj) if obj.startswith(".") else obj  # 先頭が.で始まっているときcom.sun.starが省略されていると考えて、com.sun.starを追加する。
		j = _idl_check(idl)  # TypeDescriptionオブジェクトを取得。
		if j:
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
		else:  # TypeDescriptionオブジェクトを取得できなかったとき。
			outputs.append(_("Can not get TypeDescription object of {}.".format(idl)))  # {}のTypeDescriptionが取得できません。
	else:  # objが文字列以外の時
		if hasattr(obj, "getSupportedServiceNames"):  # オブジェクトがサービスを持っているとき。
			st_ss.update(obj.getSupportedServiceNames())  # サポートサービス名を集合にして取得。
			incorrectidls = {"com.sun.star.AccessibleSpreadsheetDocumentView": "com.sun.star.sheet.AccessibleSpreadsheetDocumentView"}  # 間違って返ってくるIDL名の辞書。
			st_incorrects = st_ss.intersection(incorrectidls.keys())  # 間違っているIDL名の集合を取得。
			for incorrect in st_incorrects:
				st_ss.remove(incorrect)  # 間違ったIDL名を除く。
				st_ss.add(incorrectidls[incorrect])  # 正しいIDL名を追加する。
			if st_ss:
				args = st_ss, st_is, [_idl_check(i) for i in st_ss]
				getSuperService(args)  # TypeDescriptionオブジェクトに変換して渡す。	
		if hasattr(obj, "getTypes"):  # サービスを介さないインターフェイスがある場合。elifにしてはいけない。
			types = obj.getTypes()  # インターフェイス名ではなくtype型が返ってくる。
			if types:  # type型が取得できたとき。
				st_is.update(i.Name for i in types)
				getSuperInterface(st_is, types)
		if hasattr(obj, "getPropertySetInfo"):	# objにgetPropertySetInfoがあるとき。サービスに属さないプロパティを出力。
			properties = obj.getPropertySetInfo().getProperties()  # オブジェクトのプロパティを取得。すべてのプロパティのProperty Structのタプルが返ってくるので集合にする。
			st_ps.update(properties)
		if not any([st_ss, st_nontdm, st_is, st_ps]):
			outputs.append(_("There is no service or interface to support."))  # サポートするサービスやインターフェイスがありません。
	return st_ss, st_nontdm, st_is, st_ps  # プロパティのみProperty Structを返す。
def getSuperService(args):
	st_ss, st_is, tdms = args
	for j in tdms:
		lst_std = list(j.getMandatoryServices())
		lst_std.extend(j.getOptionalServices())
		if lst_std:
			st_ss.update(i.Name for i in lst_std)
			args = st_ss, st_is, lst_std
			getSuperService(args) 
		lst_itd = [j.getInterface()]  # new-styleサービスのインターフェイスを取得。TypeDescriptionオブジェクト。
		if not lst_itd:  # new-styleサービスのインターフェイスがないときはold-styleサービスのインターフェイスを取得。
			lst_itd.extend(j.getMandatoryInterfaces())
			lst_itd.extend(j.getOptionalInterfaces())
		if lst_itd:
			st_is.update(i.Name for i in lst_itd)
			getSuperInterface(st_is, lst_itd)	
# 		st_ps.update(i.Name for i in j.Properties)  # サービスからXPropertyTypeDescriptionインターフェイスのタプルを取得してプロパティ名を取得。
def getSuperInterface(st_is, tdms):
	for j in tdms:
		lst_itd = list(j.getBaseTypes())
		lst_itd.extend(j.getOptionalBaseTypes())
		if lst_itd:
			st_is.update(i.Name for i in lst_itd)
			getSuperInterface(st_is, lst_itd)
def createTree2(args):
	tdm, css, fns, st_s, st_non, st_i , st_p, st_omi = args  # st_pの要素はプロパティ名ではなくProperty Struct。
	indent = "	  "  # インデントを設定。
	st_oms, st_omp = set(), set()  # すでに取得したサービス名、プロパティ名を入れる集合。
	consumeStack = createStackConsumer(indent, css, fns, st_oms, st_omi, st_omp)
	non_ss = getNonSuperServices(tdm, st_s)  # st_sからスーパークラスになるサービス名を除いたサービス名の集合を取得。
	if non_ss:  
		stack = [tdm.getByHierarchicalName(i) for i in sorted(non_ss, reverse=True)]  # サービス名を降順にしてTypeDescriptionオブジェクトをスタックに取得。		
		st_oms.update(non_ss)
	else:  # サービス名がないとき。
		non_si = getNonSuperInterfaces(tdm, st_i)  # st_iからスーパークラスになるインターフェイス名を除いたインターフェイス名の集合を取得。
		if non_si:  # インターフェイスがあるとき。
			stack = [tdm.getByHierarchicalName(i) for i in sorted(non_si, reverse=True)]  # 降順にしてTypeDescriptionオブジェクトに変換してスタックに取得。
			st_omi.update(non_si)  # すでに取得したインターフェイス名の集合に追加。
	consumeStack(stack)  # スーパークラスのないサービス名を元にツリーを出力。
	st_s.difference_update(st_oms)  # すでに出力されたサービス名を除く。比較のときは出力抑制されたスーパークラスのサービス名がでてくる。
	if st_s:  # まだ出力されていないサービスが残っているとき。
		non_ss = getNonSuperServices(tdm, st_s)  # st_sからスーパークラスになるサービス名を除いたサービス名の集合を取得。
		if non_ss:  
			stack = [tdm.getByHierarchicalName(i) for i in sorted(non_ss, reverse=True)]  # サービス名を降順にしてTypeDescriptionオブジェクトをスタックに取得。		
			st_oms.update(non_ss) # すでに取得したサービス名の集合に追加。
			consumeStack(stack)  # スーパークラスのないサービス名を元にツリーを出力。
	st_i.difference_update(st_omi)  # すでに出力されたインターフェイス名を除く。サービスがあるのに、サービスを介さないインターフェイス名があるときにそれが残る。
	if st_i:  # まだ出力されていないインターフェイスが残っているとき。
		non_si = getNonSuperInterfaces(tdm, st_i)  # st_iからスーパークラスになるインターフェイス名を除いたインターフェイス名の集合を取得。
		if non_si:  # インターフェイスがあるとき。
			stack = [tdm.getByHierarchicalName(i) for i in sorted(non_si, reverse=True)]  # 降順にしてTypeDescriptionオブジェクトに変換してスタックに取得。
			st_omi.update(non_si)  # すでに取得したインターフェイス名の集合に追加。
			consumeStack(stack)  # スーパークラスのないインターフェイス名を元にツリーを出力。
	properties = [i for i in st_p if not i.Name in st_omp]  # すでに出力されたプロパティを除く。Property Structのリストが返る。			
	branchfirst = "├─" if properties else "└─"  # まだ出力していないプロパティがあるときは枝を変える。
	if st_non:  # TypeDescriptionオブジェクトを取得できないサービスを出力する。コントロール関係は実装サービス名がここにでてくる。
		if len(st_non)==1:  # サービスがひとつの時
			branch = [branchfirst] 
			branch.append(st_non[0].replace(css, ""))  # サービス名をbranchの要素に追加。
			fns["NOLINK"]("".join(branch))  # リンクをつけずに出力。			
		else:  # サービスが複数のとき(UnoControlDialogModelなど)
			lst_nontyps = sorted(st_non)  # 昇順に並べる
			for i in lst_nontyps[:-1]:  # 一番最後の要素以外について
				branch = ["├─"] 
				branch.append(i.replace(css, ""))  # サービス名をbranchの要素に追加。
				fns["NOLINK"]("".join(branch))  # リンクをつけずに出力。	
			branch = [branchfirst] 
			branch.append(lst_nontyps[-1].replace(css, ""))  # 一番最後のサービス名をbranchの要素に追加。
			fns["NOLINK"]("".join(branch))  # リンクをつけずに出力。	
def getNonSuperInterfaces(tdm, st_i):
	stack = [tdm.getByHierarchicalName(i) for i in st_i]  # インターフェイス名一覧のTypeDescriptionオブジェクトをスタックに取得。
	st_supi = set()  # スーパークラスになるインターフェイスを入れる集合。
	while stack:  # スタックがある間実行。
		j = stack.pop()  # インターフェイスのTypeDescriptionオブジェクトを取得。
		lst_super = list(j.getBaseTypes())  # スーパークラスの取得。
		lst_super.extend(j.getOptionalBaseTypes())  # スーパークラスの取得。
		new_itd = [i for i in lst_super if not i.Name in st_supi]  # st_supiにまだないTypeDescriptionオブジェクトのみ取得。
		stack.extend(new_itd)  # スタックに新たなサービスのTypeDescriptionオブジェクトのみ追加。
		st_supi.update(i.Name for i in new_itd)  # 既に取得した親サービス名の集合型に新たに取得したインターフェイス名を追加。	
	return st_i.difference(st_supi)  # オブジェクトのサポートインターフェイスのうちスーパークラスになるインターフェイスを除いた集合を返す。
def getNonSuperServices(tdm, st_s):
	stack = [tdm.getByHierarchicalName(i) for i in st_s]  # サービス名一覧のTypeDescriptionオブジェクトをスタックに取得。
	st_sups = set()  # スーパークラスになるサービスを入れる集合。
	while stack:  # スタックがある間実行。
		j = stack.pop()  # サービスのTypeDescriptionオブジェクトを取得。
		lst_super = list(j.getMandatoryServices())  # スーパークラスの取得。
		lst_super.extend(j.getOptionalServices())  # スーパークラスの取得。
		new_std = [i for i in lst_super if not i.Name in st_sups]  # st_supsにまだないTypeDescriptionオブジェクトのみ取得。
		stack.extend(new_std)  # スタックに新たなサービスのTypeDescriptionオブジェクトのみ追加。
		st_sups.update(i.Name for i in new_std)  # 既に取得した親サービス名の集合型に新たに取得したサービス名を追加。	
	return st_s.difference(st_sups)  # オブジェクトのサポートサービスのうちスーパークラスになるサービスを除いた集合を返す。	
def createStackConsumer(indent, css, fns, st_oms, st_omi, st_omp):	
		reg_sqb = re.compile(r'\[\]')  # 型から角括弧ペアを取得する正規表現オブジェクト。
		inout_dic = {(True, False): "[in]", (False, True): "[out]", (True, True): "[inout]"}  # メソッドの引数のinout変換辞書。	
	# 		lst_itd = tuple()  # インターフェイスのTypeDescriptionオブジェクトの入れ物を初期化。	
		def consumeStack(stack):
			m = 0  # 最大文字数を初期化。
			lst_level = [1]*len(stack)  # stackの要素すべてについて階層を取得。
			def _format_type(typ):  # 属性がシークエンスのとき[]の表記を修正。
				n = len(reg_sqb.findall(typ))  # 角括弧のペアのリストの数を取得。
				for dummy in range(n):  # 角括弧の数だけ繰り返し。
					typ = typ.replace("]", "", 1) + "]" 
				return typ			
			def _stack_interface(lst_itd):  # インターフェイスをスタックに追加する。
				lst_itd = [i for i in lst_itd if not i.Name in st_omi]  # st_omiを除く。
				stack.extend(sorted(lst_itd, key=lambda x: x.Name, reverse=True))  # 降順にしてスタックに追加。
				lst_level.extend(level+1 for i in lst_itd)  # 階層を取得。
				st_omi.update(i.Name for i in lst_itd)  # すでにでてきたインターフェイス名をst_omiに追加して次は使わないようにする。
			while stack:  # スタックがある間実行。
				j = stack.pop()  # スタックからTypeDescriptionオブジェクトをpop。
				level = lst_level.pop()  # jの階層を取得。
				typcls = j.getTypeClass()  # jのタイプクラスを取得。
				branch = ["", ""]  # 枝をリセット。jがサービスまたはインターフェイスのときjに直接つながる枝は1番の要素に入れる。それより左の枝は0番の要素に加える。
				if level>1:  # 階層が2以上のとき。
# 					p = 1  # 処理開始する階層を設定。
# 					if st_i:  # サービスを介さないインターフェイスがあるとき
					branch[0] = "│   "  # 階層1に立枝をつける。
					p = 2  # 階層2から処理する。
					for i in range(p, level):  # 階層iから出た枝が次行の階層i-1の枝になる。
						branch[0] += "│   " if i in lst_level else indent  # 階層iから出た枝が次行の階層i-1の枝になる。iは枝の階層ではなく、枝のより上の行にあるその枝がでた階層になる。
				if typcls==INTERFACE or typcls==SERVICE:  # jがサービスかインターフェイスのとき。
# 					if level==1 and st_i:  # 階層1かつサービスを介さないインターフェイスがあるとき
					if level==1:  # 階層1のとき
						branch[1] = "├─"  # 階層1のときは下につづく分岐をつける。
					else:
						branch[1] = "├─" if level in lst_level else "└─"  # スタックに同じ階層があるときは"├─" 。
					if typcls==INTERFACE:  # インターフェイスのとき。XInterfaceTypeDescription2インターフェイスをもつTypeDescriptionオブジェクト。
						branch.append(j.Name.replace(css, ""))  # インターフェイス名をbranchの2番要素に追加。
						fns["INTERFACE"]("".join(branch))  # 枝をつけて出力。
		# 				st_omi.add(j.Name)  # すでにでてきたインターフェイス名をst_omiに追加して次は使わないようにする。
						lst_itd = list(j.getBaseTypes())
						lst_itd.extend(j.getOptionalBaseTypes())
						if lst_itd:  # インターフェイスのスーパークラスがあるとき。(TypeDescriptionオブジェクト)
							_stack_interface(lst_itd)
						t_md = j.getMembers()  # インターフェイス属性とメソッドのTypeDescriptionオブジェクトを取得。
						if t_md:  # インターフェイス属性とメソッドがあるとき。
							stack.extend(sorted(t_md, key=lambda x: x.Name, reverse=True))  # 降順にしてスタックに追加。
							lst_level.extend([level+1 for i in t_md])  # 階層を取得。
							m = max([len(i.ReturnType.Name.replace(css, "")) for i in t_md if i.getTypeClass()==INTERFACE_METHOD] + [len(i.Type.Name.replace(css, "")) for i in t_md if i.getTypeClass()==INTERFACE_ATTRIBUTE])  # インターフェイス属性とメソッドの型のうち最大文字数を取得。
		# 					t_md = tuple()  # メソッドのTypeDescriptionオブジェクトの入れ物を初期化。
					elif typcls==SERVICE:  # jがサービスのときtdはXServiceTypeDescriptionインターフェイスをもつ。
						branch.append(j.Name.replace(css, ""))  # サービス名をbranchの2番要素に追加。
						fns["SERVICE"]("".join(branch))  # 枝をつけて出力。
		# 				st_oms.add(j.Name)  # すでにでてきたサービス名をst_omsに追加して次は使わないようにする。
						lst_std = list(j.getMandatoryServices())
						lst_std.extend(j.getOptionalServices())
		# 				t_std = j.getMandatoryServices() + j.getOptionalServices()  # 親サービスを取得。
						lst_std = [i for i in lst_std if not i.Name in st_oms]  # st_omsを除く。
						stack.extend(sorted(lst_std, key=lambda x: x.Name, reverse=True))  # サービス名で降順に並べてサービスのTypeDescriptionオブジェクトをスタックに追加。
						lst_level.extend(level+1 for i in lst_std)  # 階層を取得。
						st_oms.update(i.Name for i in lst_std)  # すでに取得したサービス名をst_omsに追加して次は使わないようにする。
						lst_itd = [j.getInterface()]  # new-styleサービスのインターフェイスを取得。TypeDescriptionオブジェクト。
		# 				if itd:  # new-styleサービスのインターフェイスがあるとき。
		# 					lst_itd = [itd]  # XInterfaceTypeDescription2インターフェイスをもつTypeDescriptionオブジェクト。
		# 				else:  # new-styleサービスのインターフェイスがないときはold-styleサービスのインターフェイスを取得。
						if not lst_itd:  # new-styleサービスのインターフェイスがないときはold-styleサービスのインターフェイスを取得。
							lst_itd.extend(j.getMandatoryInterfaces())
							lst_itd.extend(j.getOptionalInterfaces())					
		# 					t_itd = j.getMandatoryInterfaces() + j.getOptionalInterfaces()  # XInterfaceTypeDescriptionインターフェイスをもつTypeDescriptionオブジェクト。
						if lst_itd:  # 親インターフェイスがあるとき。(TypeDescriptionオブジェクト)
							_stack_interface(lst_itd)  # インターフェイスをスタックに追加する。
						t_spd = j.Properties  # サービスからXPropertyTypeDescriptionインターフェイスをもつオブジェクトのタプルを取得。
						if t_spd:  # プロパティがあるとき。
							stack.extend(sorted(list(t_spd), key=lambda x: x.Name, reverse=True))  # 降順にしてスタックに追加。
							lst_level.extend(level+1 for i in t_spd)  # 階層を取得。
							m = max(len(i.getPropertyTypeDescription().Name.replace(css, "")) for i in t_spd)  # サービス属性の型のうち最大文字数を取得。
							st_omp.update(i.Name for i in t_spd) # すでに出力したプロパティ名の集合に追加。
		# 					t_spd = tuple()  # サービス属性のTypeDescriptionオブジェクトの入れ物を初期化。
				else:  # jがインターフェイスかサービス以外のとき。
					branch[1] = indent  # 横枝は出さない。
					if level in lst_level:  # スタックに同じ階層があるとき。
						typcls2 = stack[lst_level.index(level)].getTypeClass()  # スタックにある同じ階層のものの先頭の要素のTypeClassを取得。
						if typcls2==INTERFACE or typcls2==SERVICE: branch[1] = "│   "  # サービスかインターフェイスのとき。横枝だったのを縦枝に書き換える。
					if typcls==INTERFACE_METHOD:  # jがメソッドのとき。
						typ = _format_type(j.ReturnType.Name.replace(css, ""))  # 戻り値の型を取得。
						stack2 = list(j.Parameters)[::-1]  # メソッドの引数について逆順(降順ではない)にスタック2に取得。
						if not stack2:  # 引数がないとき。
							branch.append("{}  {}()".format(typ.rjust(m), j.MemberName.replace(css, "")))  # 「戻り値の型(固定幅mで右寄せ) メソッド名()」をbranchの3番の要素に取得。
							fns["INTERFACE_METHOD"]("".join(branch))  # 枝をつけてメソッドを出力。
						else:  # 引数があるとき。
							m3 = max(len(i.Type.Name.replace(css, "")) for i in stack2)  # 引数の型の最大文字数を取得。
							k = stack2.pop()  # 先頭の引数を取得。
							inout = inout_dic[(k.isIn(), k.isOut())]  # 引数の[in]の判定、[out]の判定
							typ2 = _format_type(k.Type.Name.replace(css, ""))  # 戻り値の型を取得。
							branch.append("{}  {}( {} {} {}".format(typ.rjust(m), j.MemberName.replace(css, ""), inout, typ2.rjust(m3), k.Name.replace(css, "")))  # 「戻り値の型(固定幅で右寄せ)  メソッド名(inout判定　引数の型(固定幅m3で左寄せ) 引数名」をbranchの3番の要素に取得。
							m2 = len("{}  {}( ".format(typ.rjust(m), j.MemberName.replace(css, "")))  # メソッドの引数の部分をインデントする文字数を取得。
							if stack2:  # 引数が複数あるとき。
								branch.append(",")  # branchの4番の要素に「,」を取得。
								fns["INTERFACE_METHOD"]("".join(branch))  # 枝をつけてメソッド名とその0番の引数を出力。
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
					elif typcls==PROPERTY:  # サービス属性(つまりプロパティ)のとき。
						typ = _format_type(j.getPropertyTypeDescription().Name.replace(css, ""))  # 属性の型を取得。
						name = j.Name  # プロパティ名を取得。
						branch.append("{}  {}".format(typ.rjust(m), name))  # 型は最大文字数で右寄せにする。
						fns["PROPERTY"]("".join(branch))  # 枝をつけて出力。
					elif typcls==INTERFACE_ATTRIBUTE:  # インターフェイス属性(つまりアトリビュート)のとき。
						typ = _format_type(j.Type.Name.replace(css, ""))  # 戻り値の型を取得。
						branch.append("{}  {}".format(typ.rjust(m), j.MemberName.replace(css, "")))  # 型は最大文字数で右寄せにする。
						fns["INTERFACE_METHOD"]("".join(branch))  # 枝をつけて出力。				
		return consumeStack

