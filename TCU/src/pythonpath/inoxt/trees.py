#!/opt/libreoffice5.4/program/python
# -*- coding: utf-8 -*-
import re
from com.sun.star.container import NoSuchElementException
from com.sun.star.uno.TypeClass import SERVICE, INTERFACE, PROPERTY, INTERFACE_METHOD, INTERFACE_ATTRIBUTE
from .common import localization
from .common import enableRemoteDebugging  # デバッグ用デコレーター
# @enableRemoteDebugging
def createTree(args, obj):
	ctx, configurationprovider, css, fns, st_omi, outputs, s, st_oms = args  # st_omi: スタックに追加しないインターフェイス名の集合。ss_omi: : スタックに追加しないサービス名の集合。
# 	st_ss = set()  # スタックに追加しいないサービス名の集合。
	stack = []  # スタックを初期化。
	st_si = False  #  サポートインターフェイス名の集合。
	st_nontyps = set()  # TypeDescriptionオブジェクトを取得できないサービス。
	global _  # グローバルな_を地域化関数に置換する。
	_ = localization(configurationprovider)  # グローバルな_を地域化関数に置換。
	tdm = ctx.getByName('/singletons/com.sun.star.reflection.theTypeDescriptionManager')  # TypeDescriptionManagerをシングルトンでインスタンス化。
	def _idl_check(idl): # IDL名からTypeDescriptionオブジェクトを取得。
		try:
			return tdm.getByHierarchicalName(idl)  # IDL名からTypeDescriptionオブジェクトを取得。
		except NoSuchElementException:  # TypeDescriptionを取得できないIDLのとき
			st_nontyps.add(idl)  # 最後に枝を出す。
			st_si = True  # 縦枝が最後まで伸びるようにしておく。
			return None
	if isinstance(obj, str): # objが文字列(IDL名)のとき
		idl = "{}{}".format(css, obj) if obj.startswith(".") else obj  # 先頭が.で始まっているときcom.sun.starが省略されていると考えて、com.sun.starを追加する。
		j = _idl_check(idl)  # IDL名からTypeDescriptionオブジェクトを取得。
		if j is None:  # TypeDescriptionオブジェクトを取得できなかったとき。
			outputs.append(_("Can not get TypeDescription object of {}.".format(idl)))  # {}のTypeDescriptionが取得できません。
			return
		typcls = j.getTypeClass()  # jのタイプクラスを取得。
		outputs.append(idl)  # treeの根にIDL名を表示
		stack = [j]  # TypeDescriptionオブジェクトをスタックに取得			
		if typcls == INTERFACE:  # インターフェイスの時
			st_omi.add(idl)  # スタックに追加しないインターフェイス名に追加する。
		elif typcls == SERVICE:  # サービスの時
			st_oms.add(idl)  # スタックに追加しないサービス名に追加する。			
		else:  # サービスかインターフェイス以外のときは未対応。
			outputs.append(_("{} is not a service name or an interface name, so it is not supported yet.".format(idl)))  # はサービス名またはインターフェイス名ではないので未対応です。
			return
	else:  # objが文字列以外の時
		outputs.append("object")  # treeの根に表示させるもの。
		if hasattr(obj, "getSupportedServiceNames"):  # オブジェクトがサービスを持っているとき。elifにしてはいけない。
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
		args = css, fns, st_omi, st_oms, stack, st_si, tdm, st_nontyps, obj
		generateOutputs(args)
		removeBranch(s, outputs)  # 不要な枝を削除。置換する空白を渡す。
	else:
		outputs.append(_("There is no service or interface to support."))  # サポートするサービスやインターフェイスがありません。
def removeBranch(s, outputs):  # 不要な枝を削除。引数は半角スペース。" "か"&nbsp;"
	def _replaceBar(j, line, ss):  #  不要な縦棒を空白に置換。
		line = line.replace("│",ss,1)
		outputs[j] = line	
	i = None;  # 縦棒の位置を初期化。	
	n = len(outputs)  # 出力行数を取得。
	for j in reversed(range(n)):  # 出力行を下からみていく。
		line = outputs[j]  # 行の内容を取得。
		if j == n - 1 :  # 最終行のとき
			if "│" in line:  # 本来あってはならない縦棒があるとき。
				i = line.find("│")  # 縦棒の位置を取得。
				_replaceBar(j, line, s*2) #  不要な縦棒を空白に置換。
			else:
				break  # 最終行に縦棒がなければループを出る。
		else:
			if "│" in line[i]:  # 消去するべき縦棒があるとき
				_replaceBar(j, line, s*2) #  不要な縦棒を空白に置換。
			else:  # 縦棒が途切れたとき
				line = line.replace("├─","└─",1)  # 下向きの分岐を消す。
				outputs[j] = line
				break
# @enableRemoteDebugging
def generateOutputs(args):  # 末裔から祖先を得て木を出力する。flagはオブジェクトが直接インターフェイスをもっているときにTrueになるフラグ。
	reg_sqb = re.compile(r'\[\]')  # 型から角括弧ペアを取得する正規表現オブジェクト。
	css, fns, st_omi, st_oms, stack, st_si, tdm, st_nontyps, obj = args
	lst_level = [1]*len(stack)  # stackの要素すべてについて階層を取得。
	indent = "	  "  # インデントを設定。
	m = 0  # 最大文字数を初期化。
	inout_dic = {(True, False): "[in]", (False, True): "[out]", (True, True): "[inout]"}  # メソッドの引数のinout変換辞書。
	t_itd = tuple()  # インターフェイスのTypeDescriptionオブジェクトの入れ物を初期化。
	t_md = tuple()  # メソッドのTypeDescriptionオブジェクトの入れ物を初期化。
	t_spd = tuple()  # サービス属性のTypeDescriptionオブジェクトの入れ物を初期化。
	def _consumeStack(stack):  # fnsの関数による出力順を変更してはいけない。	
		def _format_type(typ):  # 属性がシークエンスのとき[]の表記を修正。
			n = len(reg_sqb.findall(typ))  # 角括弧のペアのリストの数を取得。
			for dummy in range(n):  # 角括弧の数だけ繰り返し。
				typ = typ.replace("]", "", 1) + "]" 
			return typ			
		def _stack_interface(t_itd):  # インターフェイスをスタックに追加する。
			lst_itd = [i for i in t_itd if not i.Name in st_omi]  # st_omiを除く。
			stack.extend(sorted(lst_itd, key=lambda x: x.Name, reverse=True))  # 降順にしてスタックに追加。
			lst_level.extend(level+1 for i in lst_itd)  # 階層を取得。
			st_omi.update(i.Name for i in lst_itd)  # すでにでてきたインターフェイス名をst_omiに追加して次は使わないようにする。
			t_itd = tuple()  # インターフェイスのTypeDescriptionオブジェクトの入れ物を初期化。	
		while stack:  # スタックがある間実行。
			j = stack.pop()  # スタックからTypeDescriptionオブジェクトをpop。
			level = lst_level.pop()  # jの階層を取得。
			typcls = j.getTypeClass()  # jのタイプクラスを取得。
			branch = ["", ""]  # 枝をリセット。jがサービスまたはインターフェイスのときjに直接つながる枝は1番の要素に入れる。それより左の枝は0番の要素に加える。
			if level>1:  # 階層が2以上のとき。
				p = 1  # 処理開始する階層を設定。
				if st_si:  # サービスを介さないインターフェイスがあるとき
					branch[0] = "│   "  # 階層1に立枝をつける。
					p = 2  # 階層2から処理する。
				for i in range(p, level):  # 階層iから出た枝が次行の階層i-1の枝になる。
					branch[0] += "│   " if i in lst_level else indent  # 階層iから出た枝が次行の階層i-1の枝になる。iは枝の階層ではなく、枝のより上の行にあるその枝がでた階層になる。
			if typcls==INTERFACE or typcls==SERVICE:  # jがサービスかインターフェイスのとき。
				if level==1 and st_si:  # 階層1かつサービスを介さないインターフェイスがあるとき
					branch[1] = "├─"  # 階層1のときは下につづく分岐をつける。
				else:
					branch[1] = "├─" if level in lst_level else "└─"  # スタックに同じ階層があるときは"├─" 。
				if typcls==INTERFACE:  # インターフェイスのとき。XInterfaceTypeDescription2インターフェイスをもつTypeDescriptionオブジェクト。
					branch.append(j.Name.replace(css, ""))  # インターフェイス名をbranchの2番要素に追加。
					fns["INTERFACE"]("".join(branch))  # 枝をつけて出力。
					t_itd = j.getBaseTypes() + j.getOptionalBaseTypes()  # 親インターフェイスを取得。	
					if t_itd:  # 親インターフェイスがあるとき。(TypeDescriptionオブジェクト)
						_stack_interface(t_itd)
					t_md = j.getMembers()  # インターフェイス属性とメソッドのTypeDescriptionオブジェクトを取得。
					if t_md:  # インターフェイス属性とメソッドがあるとき。
						stack.extend(sorted(t_md, key=lambda x: x.Name, reverse=True))  # 降順にしてスタックに追加。
						lst_level.extend([level+1 for i in t_md])  # 階層を取得。
						m = max([len(i.ReturnType.Name.replace(css, "")) for i in t_md if i.getTypeClass()==INTERFACE_METHOD] + [len(i.Type.Name.replace(css, "")) for i in t_md if i.getTypeClass()==INTERFACE_ATTRIBUTE])  # インターフェイス属性とメソッドの型のうち最大文字数を取得。
						t_md = tuple()  # メソッドのTypeDescriptionオブジェクトの入れ物を初期化。
				elif typcls==SERVICE:  # jがサービスのときtdはXServiceTypeDescriptionインターフェイスをもつ。
					branch.append(j.Name.replace(css, ""))  # サービス名をbranchの2番要素に追加。
					fns["SERVICE"]("".join(branch))  # 枝をつけて出力。
					t_std = j.getMandatoryServices() + j.getOptionalServices()  # 親サービスを取得。
					lst_std = [i for i in t_std if not i.Name in st_oms]  # st_ssを除く。
					stack.extend(sorted(lst_std, key=lambda x: x.Name, reverse=True))  # 親サービス名で降順に並べてサービスのTypeDescriptionオブジェクトをスタックに追加。
					lst_level.extend(level+1 for i in lst_std)  # 階層を取得。
					st_oms.update(i.Name for i in lst_std)  # すでにでてきたサービス名をst_ssに追加して次は使わないようにする。
					itd = j.getInterface()  # new-styleサービスのインターフェイスを取得。TypeDescriptionオブジェクト。
					if itd:  # new-styleサービスのインターフェイスがあるとき。
						t_itd = itd,  # XInterfaceTypeDescription2インターフェイスをもつTypeDescriptionオブジェクト。
					else:  # new-styleサービスのインターフェイスがないときはold-styleサービスのインターフェイスを取得。
						t_itd = j.getMandatoryInterfaces() + j.getOptionalInterfaces()  # XInterfaceTypeDescriptionインターフェイスをもつTypeDescriptionオブジェクト。
					if t_itd:  # 親インターフェイスがあるとき。(TypeDescriptionオブジェクト)
						_stack_interface(t_itd)
					t_spd = j.Properties  # サービスからXPropertyTypeDescriptionインターフェイスをもつオブジェクトのタプルを取得。
					if t_spd:  # サービス属性があるとき。
						stack.extend(sorted(list(t_spd), key=lambda x: x.Name, reverse=True))  # 降順にしてスタックに追加。
						lst_level.extend(level+1 for i in t_spd)  # 階層を取得。
						m = max(len(i.getPropertyTypeDescription().Name.replace(css, "")) for i in t_spd)  # サービス属性の型のうち最大文字数を取得。
						t_spd = tuple()  # サービス属性のTypeDescriptionオブジェクトの入れ物を初期化。
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
				elif typcls==PROPERTY:  # サービス属性のとき。
					typ = _format_type(j.getPropertyTypeDescription().Name.replace(css, ""))  # 属性の型を取得。
					branch.append("{}  {}".format(typ.rjust(m), j.Name.replace(css, "")))  # 型は最大文字数で右寄せにする。
					fns["PROPERTY"]("".join(branch))  # 枝をつけて出力。
				elif typcls==INTERFACE_ATTRIBUTE:  # インターフェイス属性のとき。
					typ = _format_type(j.Type.Name.replace(css, ""))  # 戻り値の型を取得。
					branch.append("{}  {}".format(typ.rjust(m), j.MemberName.replace(css, "")))  # 型は最大文字数で右寄せにする。
					fns["INTERFACE_METHOD"]("".join(branch))  # 枝をつけて出力。						
	_consumeStack(stack)
	if isinstance(st_si, set):  # st_siが集合のとき、それはオブジェクトから得たインターフェイス名。
		st_rem = st_si.difference(st_omi)  # まだでてきていないインターフェイスがサービスを介さないインターフェイス。
		if st_rem:  # まだ出力していないインターフェイスが残っているとき
			stack =  [tdm.getByHierarchicalName(i) for i in sorted(st_rem, reverse=True)]  # 降順にしてTypeDescriptionオブジェクトに変換してスタックに取得。CSSが必要。
			lst_level = [1]*len(stack)  # stackの要素すべてについて階層を取得。
			st_omi.update(st_rem)  # すでにでてきたインターフェイス名をst_omiに追加して次は使わないようにする。
			_consumeStack(stack)		
	st_nontyps.difference_update(st_oms)  # TypeDescriptionオブジェクトを取得できないサービスから出力を抑制するサービスを除く。
	if st_nontyps:  # TypeDescriptionオブジェクトを取得できないサービスを最後に出力する。
		if len(st_nontyps)==1:  # サービスがひとつの時
			branch = ["└─"] 
			branch.append(st_nontyps.pop().replace(css, ""))  # サービス名をbranchの要素に追加。
			fns["NOLINK"]("".join(branch))  # リンクをつけずに出力。			
		else:  # サービスが複数のとき(があるだろうか)?
			lst_nontyps = sorted(st_nontyps)  # 昇順に並べる
			for i in lst_nontyps[:-1]:  # 一番最後の要素以外について
				branch = ["├─"] 
				branch.append(i.replace(css, ""))  # サービス名をbranchの要素に追加。
				fns["NOLINK"]("".join(branch))  # リンクをつけずに出力。	
			branch = ["└─"] 
			branch.append(lst_nontyps[-1].replace(css, ""))  # 一番最後のサービス名をbranchの要素に追加。
			fns["NOLINK"]("".join(branch))  # リンクをつけずに出力。		
		if hasattr(obj, "getPropertySetInfo"):	# objにgetPropertySetInfoがあるとき
			properties = obj.getPropertySetInfo().getProperties()  # オブジェクトのプロパティーを取得。
			props = sorted(properties, key=lambda x: x.Name)  #Name属性で昇順に並べる。
			m = max(len(i.Type.typeName.replace(css, "")) for i in props)  # プロパティの型のうち最大文字数を取得。
			for i in props:  # 各プロパティについて。
				branch = [indent*2]  # 枝をリセット。
				branch.append("{}  {}".format(i.Type.typeName.replace(css, "").rjust(m), i.Name))  # 型は最大文字数で右寄せにする。
				fns["PROPERTY"]("".join(branch))  # 枝をつけて出力。
