import pandas as pd
import re
import unicodedata
import argparse

DEFAULT_FILE = 'tablo4_01082025d.xlsx'

provinces = {
    'MARMARA': ['BALIKESIR','BILECIK','BURSA','CANAKKALE','EDIRNE','ISTANBUL','KIRKLARELI','KOCAELI','SAKARYA','TEKIRDAG','YALOVA'],
    'EGE': ['AFYONKARAHISAR','AYDIN','DENIZLI','IZMIR','MANISA','MUGLA','KUTAHYA','USAK'],
    'AKDENIZ': ['ADANA','ANTALYA','BURDUR','HATAY','ISPARTA','MERSIN','OSMANIYE','KAHRAMANMARAS'],
    'IC ANADOLU': ['ANKARA','AKSARAY','CANKIRI','ESKISEHIR','KAYSERI','KIRIKKALE','KIRSEHIR','KONYA','NEVSEHIR','NIGDE','SIVAS','YOZGAT','KARAMAN'],
    'KARADENIZ': ['AMASYA','ARTVIN','BARTIN','BAYBURT','BOLU','CORUM','DUZCE','GIRESUN','GUMUSHANE','KASTAMONU','ORDU','RIZE','SAMSUN','SINOP','TOKAT','TRABZON','ZONGULDAK','KARABUK'],
    'DOGU ANADOLU': ['AGRI','BINGOL','BITLIS','ELAZIG','ERZINCAN','ERZURUM','HAKKARI','KARS','MALATYA','MUS','TUNCELI','VAN','ARDAHAN','IGDIR'],
    'GUNEYDOGU': ['ADIYAMAN','BATMAN','DIYARBAKIR','GAZIANTEP','KILIS','MARDIN','SIIRT','SANLIURFA','SIRNAK']
}
REGION_MAP = {prov:region for region,cities in provinces.items() for prov in cities}

ALLOWED_REGIONS = {'KARADENIZ','AKDENIZ','MARMARA','EGE','IC ANADOLU'}

HEADER_ROWS = [0,1,2]


def parse_args():
    parser = argparse.ArgumentParser(description="Generate preference list")
    parser.add_argument('--file', default=DEFAULT_FILE,
                        help='path to Excel file, default %(default)s')
    parser.add_argument('--output', default='preference_list.csv',
                        help='output CSV file')
    return parser.parse_args()


def extract_city(uni: str) -> str:
    if not uni:
        return None
    m = re.search(r"\(([^()]+)\)", uni)
    if m:
        inner = ''.join(c for c in unicodedata.normalize('NFD', m.group(1)) if unicodedata.category(c) != 'Mn').upper()
        if 'UNIVERSITESI' not in inner:
            return m.group(1).strip().upper()
    base = ''.join(c for c in unicodedata.normalize('NFD', uni) if unicodedata.category(c) != 'Mn').upper()
    return base.split('UNIVERSITESI')[0].split()[0].strip()


def categorize(name: str):
    if not isinstance(name, str):
        return None
    n = unicodedata.normalize('NFD', name)
    n = ''.join(c for c in n if unicodedata.category(c) != 'Mn')
    n = n.lower()
    if 'ingilizce ogretmenligi' in n:
        return 1
    if ('mutercim' in n or 'm%C3%BCtercim' in n) and 'ingilizce' in n:
        return 2
    if 'dilbilim' in n or 'dil bilim' in n:
        return 3
    return None


def read_excel(path: str) -> pd.DataFrame:
    df = pd.read_excel(path, header=HEADER_ROWS)
    new_cols = []
    for col in df.columns:
        parts = [str(x).strip() for x in col if str(x) != 'nan' and not str(x).startswith('Unnamed:')]
        new_cols.append(' '.join(parts).strip())
    df.columns = pd.Index(new_cols)
    df.columns = df.columns.str.replace('\n', ' ').str.strip()
    return df


def main():
    args = parse_args()
    df = read_excel(args.file)

    unis = []
    current = None
    for val in df['PROGRAM ADI (2)']:
        if isinstance(val, str):
            norm = ''.join(c for c in unicodedata.normalize('NFD', val) if unicodedata.category(c) != 'Mn').upper()
            if 'UNIVERSITESI' in norm:
                current = val.strip()
        unis.append(current)
    df['uni'] = unis
    df = df[df['PROGRAM KODU (1)'].notna()]
    df['city'] = df['uni'].apply(extract_city)
    df['region'] = df['city'].map(REGION_MAP)
    df['rank'] = pd.to_numeric(df['2024-YKS BAŞARI SIRASI (12)'], errors='coerce')
    df['priority'] = df['PROGRAM ADI (2)'].apply(categorize)
    df = df[df['priority'].notna()]
    df['quota'] = pd.to_numeric(df['GENEL KONT. (5)'], errors='coerce')
    staff_cols = ['P.DR. SAYI (14)', 'D.DR. SAYI (15)', 'DR.ÖĞR. ÜYE SAYI (16)']
    for c in staff_cols:
        if c in df.columns:
            df[c] = pd.to_numeric(df[c], errors='coerce')
        else:
            df[c] = None
    df = df[(df['rank'] >= 13000) & (df['rank'] <= 30000)]
    df = df[df['region'].isin(ALLOWED_REGIONS)]
    result = df[['uni','PROGRAM ADI (2)','rank','city','region','priority','quota'] + staff_cols].sort_values(['priority','rank'])
    result.to_csv(args.output, index=False)

if __name__ == '__main__':
    main()
