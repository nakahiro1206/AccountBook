import Make
import Check
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
