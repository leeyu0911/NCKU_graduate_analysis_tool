import pandas as pd


def to_excel(data, excel_name, advisor_name):
    # data : [paper, exam_method, recommend_method, other_method, year_data, advisor_name]
    df_list = [pd.DataFrame(columns=['畢業年度', '入學方式', '就學期間', '論文名稱']) for _ in data]
    for i, d in enumerate(data):
        try:
            df_list[i]['畢業年度'] = [i.graduate_year for i in d[0]]
            df_list[i]['入學方式'] = [i.addmission_method for i in d[0]]
            df_list[i]['就學期間'] = [i.duringSchool for i in d[0]]
            df_list[i]['論文名稱'] = [i.paper_title for i in d[0]]

            if len(df_list[i]) < 15:  # 為了讓入學方式和畢業時間都可以顯示出來
                for ii in range(len(df_list[i]), 15):
                    df_list[i].loc[ii] = ['' for _ in range(4)]

            df_list[i][''] = pd.DataFrame([None])
            df_list[i]['入學方式統計'] = pd.DataFrame(['考試入學', '推甄入學', '其他'])
            df_list[i]['人數'] = pd.DataFrame([d[1], d[2], d[3]])
            df_list[i][' '] = pd.DataFrame([None])
            df_list[i]['畢業時間統計'] = pd.DataFrame([str(i) + ' 年畢業' for i, j in enumerate(d[4]) if j != 0])
            df_list[i]['人數 '] = pd.DataFrame([i for i in d[4] if i != 0])

        except ValueError:
            print('查無此資料：' + str(advisor_name[i]))

    # 寫入excel
    def adj_column_width(sheet_name):
        writer.sheets[sheet_name].column_dimensions['A'].width = 10
        writer.sheets[sheet_name].column_dimensions['B'].width = 10
        writer.sheets[sheet_name].column_dimensions['C'].width = 10
        writer.sheets[sheet_name].column_dimensions['D'].width = 70
        writer.sheets[sheet_name].column_dimensions['E'].width = 8
        writer.sheets[sheet_name].column_dimensions['F'].width = 15
        writer.sheets[sheet_name].column_dimensions['G'].width = 6
        writer.sheets[sheet_name].column_dimensions['H'].width = 2
        writer.sheets[sheet_name].column_dimensions['I'].width = 15
        writer.sheets[sheet_name].column_dimensions['J'].width = 6

    with pd.ExcelWriter(excel_name) as writer:
        df_list[0].to_excel(writer, sheet_name=advisor_name[0], index=False)
        adj_column_width(advisor_name[0])
    for i in range(1, len(advisor_name)):
        with pd.ExcelWriter(excel_name, mode='a') as writer:
            df_list[i].to_excel(writer, sheet_name=advisor_name[i], index=False)
            adj_column_width(advisor_name[i])
