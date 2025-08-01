import pandas as pd
import re
import unicodedata
import argparse

DEFAULT_FILE = '2025_YKS_KONTENJAN_KILAVUZU_sayfa138_540_rotated.xlsx'
EXTRA_FILE = 'tablo4_01082025d.xls'

provinces = {
    'MARMARA': ['BALIKESİR','BİLECİK','BURSA','ÇANAKKALE','EDİRNE','İSTANBUL','KIRKLARELİ','KOCAELİ','SAKARYA','TEKİRDAĞ','YALOVA'],
    'EGE': ['AFYONKARAHİSAR','AYDIN','DENİZLİ','İZMİR','MANİSA','MUĞLA','KÜTAHYA','UŞAK'],
    'AKDENIZ': ['ADANA','ANTALYA','BURDUR','HATAY','ISPARTA','MERSİN','OSMANIYE','KAHRAMANMARAŞ'],
    'IC ANADOLU': ['ANKARA','AKSARAY','ÇANKIRI','ESKİŞEHİR','KAYSERİ','KIRIKKALE','KIRŞEHİR','KONYA','NEVŞEHİR','NİĞDE','SİVAS','YOZGAT','KARAMAN'],
    'KARADENIZ': ['AMASYA','ARTVİN','BARTIN','BAYBURT','BOLU','ÇORUM','DÜZCE','GİRESUN','GÜMÜŞHANE','KASTAMONU','ORDU','RİZE','SAMSUN','SİNOP','TOKAT','TRABZON','ZONGULDAK','KARABÜK'],
    'DOGU ANADOLU': ['AĞRI','BİNGÖL','BİTLİS','ELAZIĞ','ERZİNCAN','ERZURUM','HAKKARİ','KARS','MALATYA','MUŞ','TUNCELİ','VAN','ARDAHAN','IĞDIR'],
    'GUNEYDOGU': ['ADIYAMAN','BATMAN','DİYARBAKIR','GAZİANTEP','KİLİS','MARDİN','SİİRT','ŞANLIURFA','ŞIRNAK']
}
REGION_MAP = {prov:region for region,cities in provinces.items() for prov in cities}

ALLOWED_REGIONS = {'KARADENIZ','AKDENIZ','MARMARA','EGE','IC ANADOLU'}

HEADER_ROW = 3


def parse_args():
    parser = argparse.ArgumentParser(description="Generate preference list")
    parser.add_argument('--file', default=DEFAULT_FILE,
                        help='path to Excel file, e.g. %s or %s' % (DEFAULT_FILE, EXTRA_FILE))
    parser.add_argument('--output', default='preference_list.csv',
                        help='output CSV file')
    return parser.parse_args()


def extract_city(uni:str) -> str:
    if not uni:
        return None
    m = re.search(r"\(([^()]+)\)", uni)
    if m and 'Üniversitesi' not in m.group(1) and 'ÜNİVERSİTESİ' not in m.group(1):
        return m.group(1).strip().upper()
    return uni.split('ÜNİVERSİTESİ')[0].split()[0].strip().upper()

def categorize(name:str):
    if not isinstance(name, str):
        return None
    n = unicodedata.normalize('NFD', name)
    n = ''.join(c for c in n if unicodedata.category(c) != 'Mn')
    n = n.lower()
    if 'ingilizce ogretmenligi' in n:
        return 1
    if ('mütercim' in n or 'mutercim' in n) and 'ingilizce' in n:
        return 2
    if 'dilbilim' in n or 'dil bilim' in n:
        return 3
    return None

def main():
    args = parse_args()
    df = pd.read_excel(args.file, header=HEADER_ROW)
    df.columns = df.columns.str.replace('\n', ' ').str.strip()

    unis = []
    current = None
    for val in df['PROGRAM ADI (2)']:
        if isinstance(val, str) and 'ÜNİVERSİTESİ' in val:
            current = val.strip()
        unis.append(current)
    df['uni'] = unis
    df = df[df['PROGRAM KODU (1)'].notna()]  # keep program rows
    df['city'] = df['uni'].apply(extract_city)
    df['region'] = df['city'].map(REGION_MAP)
    df['rank'] = pd.to_numeric(df['2024-YKS BAŞARI SIRASI (12)'], errors='coerce')
    df['priority'] = df['PROGRAM ADI (2)'].apply(categorize)
    df = df[df['priority'].notna()]
    df['quota'] = pd.to_numeric(df['GENEL KONT. (5)'], errors='coerce')
    staff_cols = ['P.DR. SAYI (14)', 'D.DR. SAYI (15)', 'DR.ÖĞR. ÜYE SAYI (16)']
    for c in staff_cols:
        df[c] = pd.to_numeric(df[c], errors='coerce')
    df = df[(df['rank'] >= 13000) & (df['rank'] <= 30000)]
    df = df[df['region'].isin(ALLOWED_REGIONS)]
    result = df[['uni','PROGRAM ADI (2)','rank','city','region','priority','quota'] + staff_cols].sort_values(['priority','rank'])
    result.to_csv(args.output, index=False)

if __name__ == '__main__':
    main()
