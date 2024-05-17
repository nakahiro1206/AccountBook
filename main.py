import Make
import Check
"""
下準備
python3 -m pip install openpyxl

ファイル構造
-- chobo2023
 L keiri.py

ディレクトリ移動
cd /Users/nakanohiroki/taido/keiri/prototype/chobo2023/
フォルダ作成してそこに移動
mkdir -p 2023Jun; cd 2023Jun
ファイル複製
(一括で複製したい時)for f in ../2023May/*.xlsx;do cp $f ".${f:10:-14}230617更新).xlsx";done;
"""
def main():
    filepath="sample.xlsx"

    #口座の種類で異なるのでここで設定。現在は使っていない.
    account=None
    # 出力先のパス
    outputpath=filepath
    #実行部分
    workbook, dict=Check.main(fp=filepath)
    # MakeChoboは保存は行わない.
    complete, dict=Make.main(workbook,dict,account)
    # MakeChoboの後にCheckを行う.save=Trueで保存できる.
    Check.main(fp=outputpath,wb=workbook,save=True,complete=complete,newdict=dict)

if __name__ == '__main__':
    main()