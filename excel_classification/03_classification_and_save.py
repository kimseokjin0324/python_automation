import pandas as pd
# pd.set_option('display.max_columns',None)
import openpyxl


class ClassificationExcel:

    path = ''

    def __init__(self, order_xlsx_filename, partner_info_xlsx_file_name,path = 'result'):
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
        self.path =path
        # 파트너 목록
        df_partners = pd.read_excel(partner_info_xlsx_file_name)

        self.brands = df_partners['브랜드'].to_list()
        self.partners = df_partners['업체명'].to_list()


    def classify(self):

        for i , row in self.order_list.iterrows():  #iterrows()는 행들을 반복할것이다
            brand_name =''
            partner_name =''
            for j in range(len(self.brands)):
                if self.brands[j] in row['상품명']:
                    brand_name = self.brands[j]
                    partner_name =self.partners[j]
                    break
            if partner_name != '':
                # 자기 자신에 대괄호를 이용해서 filter를 가능하다
                df_filtered = self.order_list[self.order_list['상품명'].str.contains(brand_name)]
                df_filtered.to_excel(f"{self.path}/[테스트몰] {partner_name}.xlsx")
            else :
                print('없는 brand name : ',brand_name, row['상품명'])


if __name__ == '__main__':
    order_excel_filename = '주문목록20221112.xlsx'
    partner_info_excel_filename = '파트너목록.xlsx'
    ce = ClassificationExcel(order_excel_filename, partner_info_excel_filename)
    ce.classify()
