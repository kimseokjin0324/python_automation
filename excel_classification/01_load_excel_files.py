import pandas as pd
import openpyxl


class ClassificationExcel:

    def __init__(self, order_xlsx_filename, partner_info_xlsx_file_name):
        #주문 목록
        df = pd.read_excel(order_xlsx_filename)
        df = df.rename(columns=df.iloc[1])  # 1번을 열제목으로 쓰겠다 라는 뜻
        '''
        0                                      NaN        NaN  ...         NaN         NaN
        1                                     주문번호       상품번호  ...          주소     주문시요구사항
        '''
        df = df.drop([df.index[0],df.index[1]])
        df = df.reset_index(drop = True)
        self.order_list = df

        # 파트너 목록
        df_partners = pd.read_excel(partner_info_xlsx_file_name)

        self.brands = df_partners['브랜드'].to_list()
        self.partners = df_partners['업체명'].to_list()
        print(self.brands)
        print(self.partners)


if __name__ == '__main__':
    order_excel_filename = '주문목록20221112.xlsx'
    partner_info_excel_filename = '파트너목록.xlsx'
    ce = ClassificationExcel(order_excel_filename, partner_info_excel_filename)
