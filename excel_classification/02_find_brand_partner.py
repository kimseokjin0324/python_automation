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


    def classify(self):
        for i , row in self.order_list.head(5).iterrows():  #iterrows()는 행들을 반복할것이다
            brand_name =''
            idx_partners = 0
            for j in range(len(self.brands)):
                if self.brands[j] in row['상품명']:
                    brand_name = self.brands[j]
                    idx_partners = j
                    break
            print(f'{row["상품명"]}은 {brand_name} 브랜드 입니다. {idx_partners}번째')
            print(f'업체명 : {self.partners[idx_partners]}')

            #print(row['상품명'])

        print(len(self.brands), self.brands)
        print(len(self.partners),self.partners)

if __name__ == '__main__':
    order_excel_filename = '주문목록20221112.xlsx'
    partner_info_excel_filename = '파트너목록.xlsx'
    ce = ClassificationExcel(order_excel_filename, partner_info_excel_filename)
    ce.classify()
