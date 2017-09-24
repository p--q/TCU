#!/opt/libreoffice5.4/program/python
# -*- coding: utf-8 -*-
import re
from com.sun.star.container import NoSuchElementException
from com.sun.star.uno.TypeClass import SERVICE, INTERFACE, PROPERTY, INTERFACE_METHOD, INTERFACE_ATTRIBUTE
from .common import localization
from .common import enableRemoteDebugging  # デバッグ用デコレーター
def createTree(args, obj):
	ctx, configurationprovider, css, fns, st_omi, outputs = args
	global _  # グローバルな_を地域化関数に置換する。
	_ = localization(configurationprovider)  # グローバルな_を地域化関数に置換。
	tdm = ctx.getByName('/singletons/com.sun.star.reflection.theTypeDescriptionManager')  # TypeDescriptionManagerをシングルトンでインスタンス化。
	def _idl_check(idl): # IDL名からTypeDescriptionオブジェクトを取得。
		try:
			return tdm.getByHierarchicalName(idl)  # IDL名からTypeDescriptionオブジェクトを取得。
		except NoSuchElementException:
			return None
	flag = False
	if isinstance(obj, str): # objが文字列(IDL名)のとき
		idl = "{}{}".format(css, obj) if obj.startswith(".") else obj  # 先頭が.で始まっているときcom.sun.starが省略されていると考えて、com.sun.starを追加する。
		j = _idl_check(idl)  # IDL名からTypeDescriptionオブジェクトを取得。
		if j is None:  # TypeDescriptionオブジェクトを取得できなかったとき。
			outputs.append(_("{} is not an IDL name.".format(idl)))  # はIDL名ではありません。
			return
		typcls = j.getTypeClass()  # jのタイプクラスを取得。
		if typcls == INTERFACE or typcls == SERVICE:  # jがサービスかインターフェイスのとき。
			outputs.append(idl)  # treeの根にIDL名を表示
			stack = [j]  # TypeDescriptionオブジェクトをスタックに取得
		else:  # サービスかインターフェイス以外のときは未対応。
			outputs.append(_("{} is not a service name or an interface name, so it is not supported yet.".format(idl)))  # はサービス名またはインターフェイス名ではないので未対応です。
			return
	else:  # objが文字列以外の時
		outputs.append("object")  # treeの根に表示させるもの。
		if hasattr(obj, "getSupportedServiceNames"):  # オブジェクトがサービスを持っているとき。
			if hasattr(obj, "getTypes"):
				flag = True  # サービスを介さないインターフェイスがあるときフラグを立てる。
			st_ss = set(i for i in obj.getSupportedServiceNames() if _idl_check(i))  # サポートサービス名一覧からTypeDescriptionオブジェクトを取得できないサービス名を除いた集合を得る。
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
			stack = [tdm.getByHierarchicalName(i) for i in st_ss]  # TypeDescriptionオブジェクトに変換。
			if stack: 
				stack.sort(key=lambda x: x.Name, reverse=True)  # Name属性で降順に並べる。
		elif hasattr(obj, "getTypes"):  # サポートしているインターフェイスがある場合。
			st_si = set(i.typeName.replace(css, "") for i in obj.getTypes())  # サポートインターフェイス名を集合型で取得。
			lst_si = sorted(st_si.difference(st_omi), reverse=True)  # 除外するインターフェイス名を除いて降順のリストにする。
			stack = [tdm.getByHierarchicalName("{}{}".format(css, i) if i.startswith(".") else i) for i in lst_si]  # TypeDescriptionオブジェクトに変換。CSSが必要。
		else:  # サポートするサービスやインターフェイスがないとき。
			outputs.append(_("There is no service or interface to support."))  # サポートするサービスやインターフェイスがありません。
			return
	args = flag, css, fns, st_omi, stack
	generateOutputs(args)
# @enableRemoteDebugging
def generateOutputs(args):  # 末裔から祖先を得て木を出力する。flagはオブジェクトが直接インターフェイスをもっているときにTrueになるフラグ。
	reg_sqb = re.compile(r'\[\]')  # 型から角括弧ペアを取得する正規表現オブジェクト。
	def _format_type(typ):  # 属性がシークエンスのとき[]の表記を修正。
		n = len(reg_sqb.findall(typ))  # 角括弧のペアのリストの数を取得。
		for dummy in range(n):  # 角括弧の数だけ繰り返し。
			typ = typ.replace("]", "", 1) + "]" 
		return typ
	flag, css, fns, st_omi, stack = args
	if stack:  # 起点となるサービスかインターフェイスがあるとき。
		lst_level = [1]*len(stack)  # stackの要素すべてについて階層を取得。
		indent = "	  "  # インデントを設定。
		m = 0  # 最大文字数を初期化。
		inout_dic = {(True, False): "[in]", (False, True): "[out]", (True, True): "[inout]"}  # メソッドの引数のinout変換辞書。
		t_itd = tuple()  # インターフェイスのTypeDescriptionオブジェクトの入れ物を初期化。
		t_md = tuple()  # メソッドのTypeDescriptionオブジェクトの入れ物を初期化。
		t_spd = tuple()  # サービス属性のTypeDescriptionオブジェクトの入れ物を初期化。
		while stack:  # スタックがある間実行。
			j = stack.pop()  # スタックからTypeDescriptionオブジェクトをpop。
			level = lst_level.pop()  # jの階層を取得。
			typcls = j.getTypeClass()  # jのタイプクラスを取得。
			branch = ["", ""]  # 枝をリセット。jがサービスまたはインターフェイスのときjに直接つながる枝は1番の要素に入れる。それより左の枝は0番の要素に加える。
			if level > 1:  # 階層が2以上のとき。
				p = 1  # 処理開始する階層を設定。
				if flag:  # サービスを介さないインターフェイスがあるとき
					branch[0] = "│   "  # 階層1に立枝をつける。
					p = 2  # 階層2から処理する。
				for i in range(p, level):  # 階層iから出た枝が次行の階層i-1の枝になる。
					branch[0] += "│   " if i in lst_level else indent  # iは枝の階層ではなく、枝のより上の行にあるその枝がでた階層になる。
			if typcls == INTERFACE or typcls == SERVICE:  # jがサービスかインターフェイスのとき。
				if level == 1 and flag:  # 階層1かつサービスを介さないインターフェイスがあるとき
					branch[1] = "├─"  # 階層1のときは下につづく分岐をつける。
				else:
					branch[1] = "├─" if level in lst_level else "└─"  # スタックに同じ階層があるときは"├─" 。
			else:  # jがインターフェイスかサービス以外のとき。
				branch[1] = indent  # 横枝は出さない。
				if level in lst_level:  # スタックに同じ階層があるとき。
					typcls2 = stack[lst_level.index(level)].getTypeClass()  # スタックにある同じ階層のものの先頭の要素のTypeClassを取得。
					if typcls2 == INTERFACE or typcls2 == SERVICE: branch[1] = "│   "  # サービスかインターフェイスのとき。横枝だったのを縦枝に書き換える。
			if typcls == INTERFACE_METHOD:  # jがメソッドのとき。
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
			else:  # jがメソッド以外のとき。
				if typcls == INTERFACE:  # インターフェイスのとき。XInterfaceTypeDescription2インターフェイスをもつTypeDescriptionオブジェクト。
					branch.append(j.Name.replace(css, ""))  # インターフェイス名をbranchの2番要素に追加。
					t_itd = j.getBaseTypes() + j.getOptionalBaseTypes()  # 親インターフェイスを取得。
					t_md = j.getMembers()  # インターフェイス属性とメソッドのTypeDescriptionオブジェクトを取得。
					fns["INTERFACE"]("".join(branch))  # 枝をつけて出力。
				elif typcls == PROPERTY:  # サービス属性のとき。
					typ = _format_type(j.getPropertyTypeDescription().Name.replace(css, ""))  # 属性の型を取得。
					branch.append("{}  {}".format(typ.rjust(m), j.Name.replace(css, "")))  # 型は最大文字数で右寄せにする。
					fns["PROPERTY"]("".join(branch))  # 枝をつけて出力。
				elif typcls == INTERFACE_ATTRIBUTE:  # インターフェイス属性のとき。
					typ = _format_type(j.Type.Name.replace(css, ""))  # 戻り値の型を取得。
					branch.append("{}  {}".format(typ.rjust(m), j.MemberName.replace(css, "")))  # 型は最大文字数で右寄せにする。
					fns["INTERFACE_METHOD"]("".join(branch))  # 枝をつけて出力。
				elif typcls == SERVICE:  # jがサービスのときtdはXServiceTypeDescriptionインターフェイスをもつ。
					branch.append(j.Name.replace(css, ""))  # サービス名をbranchの2番要素に追加。
					t_std = j.getMandatoryServices() + j.getOptionalServices()  # 親サービスを取得。
					stack.extend(sorted(list(t_std), key=lambda x: x.Name, reverse=True))  # 親サービス名で降順に並べてサービスのTypeDescriptionオブジェクトをスタックに追加。
					lst_level.extend(level + 1 for i in t_std)  # 階層を取得。
					itd = j.getInterface()  # new-styleサービスのインターフェイスを取得。TypeDescriptionオブジェクト。
					if itd:  # new-styleサービスのインターフェイスがあるとき。
						t_itd = itd,  # XInterfaceTypeDescription2インターフェイスをもつTypeDescriptionオブジェクト。
					else:  # new-styleサービスのインターフェイスがないときはold-styleサービスのインターフェイスを取得。
						t_itd = j.getMandatoryInterfaces() + j.getOptionalInterfaces()  # XInterfaceTypeDescriptionインターフェイスをもつTypeDescriptionオブジェクト。
					t_spd = j.Properties  # サービスからXPropertyTypeDescriptionインターフェイスをもつオブジェクトのタプルを取得。
					fns["SERVICE"]("".join(branch))  # 枝をつけて出力。
			if t_itd:  # 親インターフェイスがあるとき。(TypeDescriptionオブジェクト)
				lst_itd = [i for i in t_itd if not i.Name.replace(css, "") in st_omi]  # st_omiを除く。
				stack.extend(sorted(lst_itd, key=lambda x: x.Name, reverse=True))  # 降順にしてスタックに追加。
				lst_level.extend(level + 1 for i in lst_itd)  # 階層を取得。
				st_omi.update(i.Name.replace(css, "") for i in lst_itd)  # すでにでてきたインターフェイス名をst_omiに追加して次は使わないようにする。
				t_itd = tuple()  # インターフェイスのTypeDescriptionオブジェクトの入れ物を初期化。
			if t_md:  # インターフェイス属性とメソッドがあるとき。
				stack.extend(sorted(t_md, key=lambda x: x.Name, reverse=True))  # 降順にしてスタックに追加。
				lst_level.extend([level + 1 for i in t_md])  # 階層を取得。
				m = max([len(i.ReturnType.Name.replace(css, "")) for i in t_md if i.getTypeClass()==INTERFACE_METHOD] + [len(i.Type.Name.replace(css, "")) for i in t_md if i.getTypeClass()==INTERFACE_ATTRIBUTE])  # インターフェイス属性とメソッドの型のうち最大文字数を取得。
				t_md = tuple()  # メソッドのTypeDescriptionオブジェクトの入れ物を初期化。
			if t_spd:  # サービス属性があるとき。
				stack.extend(sorted(list(t_spd), key=lambda x: x.Name, reverse=True))  # 降順にしてスタックに追加。
				lst_level.extend(level + 1 for i in t_spd)  # 階層を取得。
				m = max(len(i.getPropertyTypeDescription().Name.replace(css, "")) for i in t_spd)  # サービス属性の型のうち最大文字数を取得。
				t_spd = tuple()  # サービス属性のTypeDescriptionオブジェクトの入れ物を初期化。
