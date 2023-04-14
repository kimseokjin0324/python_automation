from os import listdir
import pandas as pd
pd.set_option('display.max_columns',None)


def make_email_list(data_path,partners_filename):
    filenames = listdir(data_path)

    df_partners = pd.read_excel(partners_filename)


    #업체명 추출
    for filename in filenames:
        filename = filename.replace('.xlsx','')
        partner_name = filename.replace('[패스트몰]','')
        print(partner_name)
        found_row = df_partners[df_partners['업체명'].str.contains(partner_name)]
        email1 = str(found_row['이메일1'].values[0])
        partner_manager_name = str(found_row['컨택담당자'].values[0])
        email_cc = str(found_row['참조이메일'].values[0])
        # print(found_row)
        print(email1,partner_manager_name,email_cc)               # 이메일1  컨텍담당자 참조이메일

if __name__ == '__main__':
  make_email_list('data/','파트너목록.xlsx')