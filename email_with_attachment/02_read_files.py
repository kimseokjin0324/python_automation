from os import listdir
def make_email_list(data_path):
    filenames = listdir(data_path)

    for filename in filenames:
        filename = filename.replace('.xlsx','')
        partner_name = filename.replace('[패스트몰]','')
        print(partner_name)


if __name__ == '__main__':
  make_email_list('data/')