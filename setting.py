# ここでは帳簿においてどの列に何を書いているなどの.
# デザイン変更した時に更新が面倒なものについて記述.
class o:
    # 原本の配置.
    date=0
    id=1
    receipt=2
    payment=3
    key=4
    origname=5
    balance=6
    # その他使う情報
    max_col=7
    # 1は見出し.
    min_row=2
    max_row=1500

class c:
    # 帳簿の配置.
    # tupleで指定する時用.なのでindexで表記.
    id=7
    date=0
    sub=1
    category=10
    account=4
    name=5
    receipt=8
    payment=9
    balance=12
    note=11
    link=14
    # alphabetで指定する時.
    n_chr='F'
    r_chr='I'
    p_chr='J'
    b_chr='M'
    # merge_cells(min_row=...)で使う時.Aと1が対応.
    # visualを整えるための列指定.
    v_date_col=3
    v_sub_col=4
    v_balance_col=14
    v_name_col=7
    # その他の情報.
    max_col=15
    # 1は見出し、2は前年度引き継ぎ.
    min_row=3
    max_row=2000

minou_min_row=2
match_max_row=500
keisan_min_row=2
keisan_max_row=2000