import pandas as pd
from datetime import datetime
import re

# Read the Excel file
df = pd.read_excel('19-12_08.01.xlsx')

# Column mapping from old names to new names
column_position_mapping = {
    0: 'ECOMM_ORD_NO',              # Amazon Sipariş Numarası
    4: 'ECOMM_CATALOG_NO',          # Amazon Siparişinin Ürün Numarası
    6: 'ORDER_DATE',                # Satın Alma Tarihi
    7: 'PAYMENT_DATE',              # Ödeme Tarihi
    9: 'REPORT_DATE',               # Rapor Tarihi
    10: 'CUSTOMER_EMAIL',           # Alıcı E-postası
    11: 'CUSTOMER_NAME',            # Alıcı Adı
    13: 'SKU',                      # Satıcı SKU'su
    14: 'CATALOG_DESCRIPTION',      # Başlık
    15: 'QUANTITY',                 # Gönderilen Adet
    16: 'CURRENCY',                 # Para Birimi
    17: 'PRICE',                    # Ürün Fiyatı
    18: 'PRICE_TAX',                # Ürün Vergisi
    19: 'CARGO_COST',               # Kargo Ücreti
    20: 'CARGO_COST_TAX',           # Kargo Vergisi
    23: 'CARGO_SRV_LEVEL',          # Kargo Hizmet Düzeyi
    33: 'INVOICE_ADDRESS_1',        # Fatura Adresi 1
    34: 'INVOICE_ADDRESS_2',        # Fatura Adresi 2
    36: 'INVOICE_CITY',             # Fatura Şehri
    37: 'INVOICE_COUNTY',           # Fatura Eyaleti
    38: 'INVOICE_ZIP_CODE',         # bill-postal-code
    39: 'INVOICE_COUNTRY',          # bill-country
    40: 'PROMOSYON_PRICE',          # Ürün Promosyon İndirimi
    41: 'KARGO_PROMOSYON_INDIRIMI', # Kargo Promosyon İndirimi
    43: 'CAROG_TRACK_NO',           # Takip Numarası
    47: 'SALES_CHANNEL'             # Satış Kanalı
}

sku_to_catalog = {
    "8694407670151_fba": "10002",
    "8694407670168_fba": "10003",
    "8694407109200_fba": "10027",
    "8694407100108_fba": "10014",
    "8694407670144_fba": "10001",
    "8694407670717_fba": "10013",
    "8694407128775_fba": "10033",
    "8694407128737_fba": "10028",
    "8694407128751_fba": "10031",
    "8694407220332_fba": "10001234",
    "8694407160317_fba": "42001701",
    "8694407197931_fba": "56001701",
    "8694407734464_fba": "86111701",
    "8694407737281_fba": "86280701",
    "8694407149213_fba": "42001702",
    "8694407197948_fba": "56001702",
    "8694407204622_fba": "56050701",
    "8694407737137_fba": "86280702",
    "8694407734433_fba": "86111702",
    "8694407201973_fba": "56057701",
    "8694407139849_fba": "32001028",
    "8694407139863_fba": "32001001",
    "8694407160829_fba": "42001057",
    "8694407110282_fba": "45182",
    "8694407158338_fba": "42001011",
    "8694407192417_fba": "56001082",
    "8694407146113_fba": "36064099",
    "8694407160874_fba": "43001028",
    "8694407197863_fba": "57001028",
    "8694407162526_fba": "43001001",
    "8694407197832_fba": "57001001",
    "8694407675712_fba": "45108",
    "8694407204660_fba": "57050028",
    "8694407707598_fba": "57054028",
    "8694407162571_fba": "43001003",
    "8694407197849_fba": "57001003",
    "8694407616852_fba": "BC08.0W-0227M",
    "8694407616944_fba": "BC10.0W-0127M",
    "8694407596512_fba": "BC05.0W-0127M",
    "8694407596635_fba": "BC09.0W-0127M",
    "8694407596390_fba": "BM03.0W-0114A",
    "8694407616821_fba": "BC06.0W-0127M",
    "8694407596369_fba": "BM03.0W-0214A",
    "8694407601469_fba": "DC06.0W-02005",
    "8694407596604_fba": "BC09.0W-0227M",
    "8694407589309_fba": "AC00.0W-07001",
    "8694407596482_fba": "BC05.0W-0227M",
    "8694407589330_fba": "AC00.0W-07002",
    "8694407642271_fba": "N3201063",
    "8694407632227_fba": "N6311053",
    "8694407632197_fba": "N6311043",
    "8694407642240_fba": "N3201043",
    "8694407642363_fba": "N3301043",
    "8694407147080_fba": "10001205",
    "8694407593054_fba": "N4210063",
    "8694407631954_fba": "N4311043",
    "8694407642219_fba": "N3201033",
    "8694407642905_fba": "N6201043",
    "8694407592903_fba": "N6310043",
    "8694407598103_fba": "N4310063",
    "8694407642394_fba": "N3301063",
    "8694407700667_fba": "10015",
    "8694407642844_fba": "N6201000",
    "8694407591913_fba": "N4310033",
    "8694407721433_fba": "25064A99",
    "8694407642691_fba": "N5301000",
    "8694407642875_fba": "N6201033",
    "8694407642189_fba": "N3201000",
    "8694407672469_fba": "MGP133",
    "8694407683915_fba": "MGP211",
    "8694407672452_fba": "MGP132",
    "8694407128898_fba": "63140",
    "8694407128584_fba": "63012",
    "8694407129796_fba": "63016",
    "8694407128607_fba": "63024",
    "8694407128591_fba": "63112",
    "8694407128560_fba": "63008",
    "8694407128577_fba": "63108",
    "8694407128904_fba": "63141",
    "8694407129802_fba": "63116",
    "8694407128553_fba": "63106",
    "8694407109910_fba": "10080",
    "8694407109958_fba": "10081",
    "8694407683281_fba": "18300",
    "8694407683298_fba": "18301",
    "8694407683328_fba": "18304",
    "8694407707123_fba": "18308",
    "8694407128355_fba": "18354",
    "8694407675668_fba": "45101",
    "8694407675682_fba": "45103",
    "8694407108371_fba": "45104",
    "8694407675705_fba": "45107",
    "8694407139825_fba": "32001003",
    "8694407144843_fba": "32001004",
    "8694407140364_fba": "32001005",
    "8694407139948_fba": "32001014",
    "8694407142085_fba": "32001027",
    "8694407140302_fba": "32001029",
    "8694407143167_fba": "32001057",
    "8694407149329_fba": "42001001",
    "8694407149336_fba": "42001003",
    "8694407149350_fba": "42001005",
    "8694407149367_fba": "42001007",
    "8694407149374_fba": "42001014",
    "8694407149381_fba": "42001028",
    "8694407149398_fba": "42001029",
    "8694407158475_fba": "42001050",
    "8694407192073_fba": "56001001",
    "8694407192158_fba": "56001003",
    "8694407192257_fba": "56001004",
    "8694407192189_fba": "56001005",
    "8694407192431_fba": "56001007",
    "8694407192592_fba": "56001028",
    "8694407192622_fba": "56001029",
    "8694407192394_fba": "56001036",
    "8694407192288_fba": "56001050",
    "8694407200426_fba": "56010001",
    "8694407200488_fba": "56010003",
    "8694407200433_fba": "56010028",
    "8694407212337_fba": "56010029"
}

# Rename columns based on their positions
for position, new_name in column_position_mapping.items():
    df.columns.values[position] = new_name

# Filter out unwanted columns
df = df[column_position_mapping.values()]

# Add new columns
df['STATE'] = ''  # Empty column
df['DATE_ADDED'] = datetime.now().strftime('%d/%m/%Y')  # Add today's date in DD/MM/YYYY format

# Replace '--' with 'MERKEZ' in INVOICE_COUNTY
df['INVOICE_COUNTY'] = df['INVOICE_COUNTY'].replace('--', 'MERKEZ')

# Ensure DATE_ADDED is in the correct format
df['DATE_ADDED'] = pd.to_datetime(df['DATE_ADDED'], format='%d/%m/%Y').dt.strftime('%d/%m/%Y')

turkish_cities = ["Adana", "Ankara", "İstanbul","Istanbul","istanbul","ıstanbul","İzmir", "Antalya", "Bursa", "Adıyaman", "Afyonkarahisar", "Ağrı", "Aksaray", "Amasya", "Ardahan", "Artvin", "Aydın", "Balıkesir", "Bartın", "Batman", "Bayburt", "Bilecik", "Bingöl", "Bitlis", "Bolu", "Burdur", "Çanakkale", "Çankırı", "Çorum", "Denizli", "Diyarbakır", "Düzce", "Edirne", "Elazığ", "Erzincan", "Erzurum", "Eskişehir", "Gaziantep", "Giresun", "Gümüşhane", "Hakkari", "Hatay", "Iğdır", "Isparta", "Kahramanmaraş", "Karabük", "Karaman", "Kars", "Kastamonu", "Kayseri", "Kırıkkale", "Kırklareli", "Kırşehir", "Kilis", "Kocaeli", "Konya", "Kütahya", "Malatya", "Manisa", "Mardin", "Mersin", "Muğla", "Muş", "Nevşehir", "Niğde", "Ordu", "Osmaniye", "Rize", "Sakarya", "Samsun", "Şanlıurfa", "Siirt", "Sinop", "Sivas", "Şırnak", "Tekirdağ", "Tokat", "Trabzon", "Tunceli", "Uşak", "Van", "Yalova", "Yozgat", "Zonguldak"]

# Function to abbreviate an address
def abbreviate_address(address, abbreviations):
    # Convert the address to string in case it's not
    address_str = str(address) if not pd.isna(address) else ""

    # Use case-insensitive regular expressions for abbreviation
    for full, abbrev in abbreviations.items():
        address_str = re.sub(r'\b' + re.escape(full) + r'\b', abbrev, address_str, flags=re.IGNORECASE)

    # Remove space after separation symbols ('/', '\', '-')
    address_str = re.sub(r'([/\\-])\s', r'\1', address_str)

    # Remove extra spaces
    address_str = re.sub(r'\s+', ' ', address_str).strip()

    return address_str

def remove_city_names(address, city_list):
    # Remove city names from the address
    for city in city_list:
        address = re.sub(r'\b' + re.escape(city) + r'\b', '', address, flags=re.IGNORECASE)
     
    # Remove extra spaces that may have been created
    address = re.sub(r'\s+', ' ', address).strip()

    return address

# Function to split long addresses
def split_address(address1, address2, max_length=35):
    # Convert both addresses to strings in case they are not
    address1_str = str(address1) if not pd.isna(address1) else ""
    address2_str = str(address2) if not pd.isna(address2) else ""

    # Find a natural breakpoint near the max_length limit
    if len(address1_str) > max_length:
        breakpoint = address1_str.rfind(' ', 0, max_length)
        if breakpoint == -1:
            # No natural breakpoint found, default to the hard limit
            breakpoint = max_length

        remaining = address1_str[breakpoint:].strip()
        address1_str = address1_str[:breakpoint].strip()
        address2_str = (remaining + " " + address2_str).strip()

    return address1_str, address2_str

def swap_addresses_if_needed(address1, address2):
    if len(address2) > len(address1):
        return address2, address1
    return address1, address2

def remove_vowels_to_fit(address, max_length=35):
    vowels = "aeıioöuüAEIİOÖUÜ"  # Include both lowercase and uppercase Turkish vowels
    for char in address:
        if len(address) <= max_length:
            break
        if char in vowels:
            address = address.replace(char, '', 1)  # Remove the first occurrence of the vowel
    return address

def right_trim_to_fit(address, max_length=35):
    return address[:max_length] if len(address) > max_length else address


def transform_customer_name(name):
    # Turkish to English character mapping
    char_mapping = {
        'ç': 'c', 'ğ': 'g', 'ı': 'i', 'ö': 'o', 'ş': 's', 'ü': 'u',
        'Ç': 'C', 'Ğ': 'G', 'İ': 'I', 'Ö': 'O', 'Ş': 'S', 'Ü': 'U'
    }
    
    # Check if the name is a string instance
    if isinstance(name, str):
        # Replace Turkish characters with English equivalents and convert to uppercase
        for tr_char, en_char in char_mapping.items():
            name = name.replace(tr_char, en_char)
        return name.upper()
    else:
        # If name is not a string (possibly NaN), just return it as it is
        return name


def adjust_zip_code(zip_code):
    zip_code_str = str(zip_code)  # Ensure it's a string
    while len(zip_code_str) < 5:
        zip_code_str = '0' + zip_code_str
    return zip_code_str

def convert_date_format(date_str):
    try:
        # First, try parsing the date with ISO 8601 format
        date_obj = pd.to_datetime(date_str, errors='coerce')

        # If parsing fails (resulting in NaT), try 'DD/MM/YYYY' format
        if pd.isna(date_obj):
            date_obj = pd.to_datetime(date_str, format='%d/%m/%Y', errors='coerce')

        # Check if date_obj is still NaT after both attempts
        if pd.isna(date_obj):
            print(f"Warning: Unable to parse date '{date_str}'. Defaulting to original value.")
            return date_str

        return date_obj.strftime('%d/%m/%Y')

    except Exception as e:
        print(f"Error processing date '{date_str}': {e}")
        return date_str



#LAST TWO COLUMNS ARE GONE???

def consolidate_orders(df):
    # Convert both 'ECOMM_ORD_NO' and 'SKU' to string type
    df['ECOMM_ORD_NO'] = df['ECOMM_ORD_NO'].astype(str)
    df['SKU'] = df['SKU'].astype(str)

    # Create a new column that concatenates 'ECOMM_ORD_NO' and 'SKU'
    df['ORDER_SKU'] = df['ECOMM_ORD_NO'] + '_' + df['SKU']
    
    # Group by this new column and sum the relevant columns
    grouped = df.groupby('ORDER_SKU').agg({
        'QUANTITY': 'sum',
        'PRICE': 'sum',
        'PRICE_TAX': 'sum',
        'CARGO_COST': 'sum',
        'CARGO_COST_TAX': 'sum',
        'PROMOSYON_PRICE': 'sum',
        'KARGO_PROMOSYON_INDIRIMI': 'sum'
    }).reset_index()
    
    # Map back the summed values to the original DataFrame
    df = df.drop_duplicates(subset='ORDER_SKU').drop(columns=['QUANTITY', 'PRICE', 'PRICE_TAX', 'CARGO_COST', 'CARGO_COST_TAX', 'PROMOSYON_PRICE', 'KARGO_PROMOSYON_INDIRIMI'])
    df = df.merge(grouped, on='ORDER_SKU', how='left')
    
    # Remove the 'ORDER_SKU' as it has served its purpose
    df.drop(columns=['ORDER_SKU'], inplace=True)
    
    return df

def make_positive(value):
    # If the value is less than zero, make it positive
    if value < 0:
        return -value
    # Else, return the value as it is (covers zero and any unexpected positive values)
    return value

    # Function to get CATALOG_NO from SKU
def get_catalog_no(sku):
    return sku_to_catalog.get(sku, "Unknown")  # Returns 'Unknown' if SKU not found in dictionary
 
# Abbreviations dictionary
abbreviations = {
    "Caddesi": "Cd",
    "Cadde": "Cd",
    "Caddes": "Cd",
    "Caddessi": "Cd",
    "Caddess": "Cd",
    "Sokak": "Sk",
    "Sokağı": "Sk",
    "Sokagi": "Sk",
    "Sok": "Sk",
    "Mahallesi": "Mh",
    "Mahalle": "Mh",
    "Mahalles": "Mh",
    "Mah": "Mh",
    "Bulvarı": "Blv",
    "Bulvari": "Blv",
    "Bulvar": "Blv",
    "Bulv": "Blv",
    "Köyü": "Köy",
    "Köy": "Köy",
    "İlçesi": "İlçe",
    "İlçe": "İlçe",
    "İli": "İl",
    "Apartmanı": "Apt",
    "Apartmani": "Apt",
    "Apartman": "Apt",
    "Apt": "Apt",
    "Dairesi": "D",
    "Daire": "D",
    "Kat": "Kt",
    "Numarası": "No",
    "Numara": "No",
    "No": "No",
    "Blok": "Blk",
    "Sitesi": "Sit",
    "Site": "Sit",
    "İş Merkezi": "İş Mrk",
    "İş Merk": "İş Mrk",
    "Plaza": "Plz",
    "Lojman": "Loj",
    "Bölgesi": "Blg",
    "Bölge": "Blg",
    "Karşısı": "Krs",
    "Karşı": "Krs",
    "Yanı": "Yn",
    "Üstü": "Üst",
    "Altı": "Alt",
    "Sağ": "Sğ",
    "Sol": "Sl",
    "Merkez": "Mrkz",
    "Merkezi": "Mrkz",
    "Merkez": "Mrkz",
    "Çarşı": "Çrş",
    "Pazar": "Pzr",
    "Meydanı": "Myd",
    "Meydan": "Myd",
    "Alay": "Aly",
    "Apart": "Aprt",
    "Bağlar": "Bğl",
    "Bağları": "Bğl",
    "Bahçe": "Bçe",
    "Bahçesi": "Bçe",
    "Cami": "Cmi",
    "Camiisi": "Cmi",
    "Çıkmazı": "Çkmz",
    "Çıkmaz": "Çkmz",
    "Derbent": "Drb",
    "Derbenti": "Drb",
    "Dere": "Dr",
    "Deresi": "Dr",
    "Evleri": "Evl",
    "Ferahlık": "Ferh",
    "Geçidi": "Gçd",
    "Geçit": "Gçd",
    "Hastane": "Hst",
    "Hastanesi": "Hst",
    "Hatun": "Htn",
    "Havuzu": "Hvz",
    "Havuz": "Hvz",
    "Hayat": "Hyt",
    "Hürriyet": "Hrryt",
    "Irmak": "Irmk",
    "Irmakları": "Irmk",
    "Kale": "Kl",
    "Kaleleri": "Kl",
    "Kampüs": "Kmp",
    "Kampüsü": "Kmp",
    "Kavşak": "Kvş",
    "Kavşağı": "Kvş",
    "Kışla": "Kşl",
    "Kışlası": "Kşl",
    "Konutları": "Knt",
    "Konut": "Knt",
    "Koru": "Kru",
    "Koruluğu": "Kru",
    "Köprü": "Kpr",
    "Köprüsü": "Kpr",
    "Küme": "Küm",
    "Kümeevler": "Küm",
    "Mescit": "Mst",
    "Meydan": "Mydn",
    "Meydanı": "Mydn",
    "Muhteşem": "Mhtşm",
    "Müstakil": "Mstk",
    "Orman": "Orm",
    "Park": "Prk",
    "Parkı": "Prk",
    "Polis": "Pls",
    "Polisi": "Pls",
    "Residence": "Rsnc",
    "Residencesi": "Rsnc",
    "Saray": "Sry",
    "Sarayı": "Sry",
    "Sayfiye": "Syf",
    "Sebil": "Sbl",
    "Sokaklar": "Sk",
    "Sultan": "Sltn",
    "Sultanı": "Sltn",
    "Tepe": "Tp",
    "Tepesi": "Tp",
    "Turistik": "Trst",
    "Tünel": "Tnl",
    "Tüneli": "Tnl",
    "Vadisi": "Vds",
    "Vadi": "Vds",
    "Villa": "Vll",
    "Villası": "Vll",
    "Yalı": "Ylı",
    "Yalısı": "Ylı",
    "Yokuş": "Ykş",
    "Yokuşu": "Ykş",
    "Zafer": "Zfr",
    "Zaferi": "Zfr",
    "Çarşı": "Çrş",
    "Pazar": "Pzr",
    "Meydanı": "Myd",
    "Meydan": "Myd",
    "Belediye": "Bld",
    "Belediyesi": "Bld",
    "Eczane": "Ecz",
    "Başkanlık": "Bşknlk",
    "Başkanlığı": "Bşknlk",
    "Müdürlük": "Mdr",
    "Müdürlüğü": "Mdr",
    "Şube": "Şb",
    "Şubesi": "Şb",
    "Sağlık Ocağı": "S.O",
    "Aile Sağlık Merkezi": "ASM",
    "Poliklinik": "Plk",
    "Klinik": "Kln",
    "Anadolu": "And",
    "Lisesi": "Lis",
    "Üniversitesi": "Ünv",
    "Üniversite": "Ünv",
    "Fakültesi": "Fkl",
    "Fakülte": "Fkl",
    "Enstitüsü": "Enst",
    "Enstitü": "Enst",
    "İlkokulu": "İlk",
    "Ortaokulu": "Ort",
    "Kreş": "Krş",
    "Anaokulu": "Ank",
    "Yurt": "Yrt",
    "Lojmanları": "Loj",
    "Vakfı": "Vkf",
    "Derneği": "Drn",
    "Kütüphanesi": "Ktp",
    "Kütüphane": "Ktp",
    "Müzesi": "Mz",
    "Müze": "Mz",
    "Sanat Galerisi": "Snt Gls",
    "Sanat": "Snt",
    "Kültür Merkezi": "Klt Mrk",
    "Sergi Salonu": "Srg Sl",
    "Opera Binası": "Opr Bn",
    "Opera": "Opr",
    "Konservatuvarı": "Kns",
    "Konservatuvar": "Kns",
    "Tiyatrosu": "Tyt",
    "Tiyatro": "Tyt",
    "Sineması": "Sin",
    "Sinema": "Sin",
    "Stadyumu": "Std",
    "Stadyum": "Std",
    "Spor Salonu": "Spr Sl",
    "Alışveriş Merkezi": "Avş Mrk",
    "Alışveriş": "Avş",
    "Market": "Mkt",
    "Supermarket": "Spmkt",
    "Mağaza": "Mğz",
    "Mağazası": "Mğz",
    "Ofisi": "Ofs",
    "Ofis": "Ofs",
    "Şirketi": "Şrk",
    "Şirket": "Şrk",
    "Firması": "Frm",
    "Firma": "Frm",
    "Fabrika": "Fbr",
    "Fabrikası": "Fbr",
    "Atölyesi": "Atly",
    "Atölye": "Atly",
    "Depo": "Dp",
    "Deposu": "Dp",
    "Antrepo": "Antrp",
    "Lojistik Merkezi": "Ljst Mrk",
    "Terminali": "Trm",
    "Terminal": "Trm",
    "Garajı": "Grj",
    "Garaj": "Grj",
    "İstasyonu": "İst",
    "İstasyon": "İst",
    "Limanı": "Lmn",
    "Liman": "Lmn",
    "Havaalanı": "Hvl",
    "Havaalan": "Hvl",
    "Havalimanı": "Hvlmn",
    "Havaliman": "Hvlmn",
    "Otobüs Durağı": "Otbs Dr",
    "Durağı": "Dr",
    "Durağ": "Dr",
    "Helikopter Pisti": "Hlkptr Pst",
    "Pist": "Pst",
    "Atış Poligonu": "Atş Plg",
    "Poligon": "Plg",
    "Kamp Alanı": "Kmp Aln",
    "Kamp": "Kmp",
    "Mesire Alanı": "Msr Aln",
    "Mesire": "Msr",
    "/ ": "/",
    "Başkanlık" : "Bşk" 
}

# Specify the data types for certain columns now that they are renamed
df = df.astype({'ECOMM_CATALOG_NO': str, 'CAROG_TRACK_NO': str})

df['CUSTOMER_NAME'] = df['CUSTOMER_NAME'].apply(transform_customer_name)
df['INVOICE_ZIP_CODE'] = df['INVOICE_ZIP_CODE'].apply(adjust_zip_code)

# Convert and reformat the date columns
df['ORDER_DATE'] = df['ORDER_DATE'].apply(convert_date_format)
df['PAYMENT_DATE'] = df['PAYMENT_DATE'].apply(convert_date_format)
df['REPORT_DATE'] = df['REPORT_DATE'].apply(convert_date_format)
#df['DATE_ADDED'] = df['DATE_ADDED'].apply(convert_date_format)
df['INVOICE_CITY'] = df['INVOICE_CITY'].str.upper()

# Call the function to consolidate orders
df = consolidate_orders(df)

df['PROMOSYON_PRICE'] = df['PROMOSYON_PRICE'].apply(make_positive)

# Apply the function to create a new column
df['CATALOG_NO'] = df['SKU'].apply(get_catalog_no)

#NET_PRICE CALCULATION 
df['NET_PRICE'] = ((df['PRICE'] / df['QUANTITY']) - (df['PROMOSYON_PRICE'] / df['QUANTITY'])) / 1.2

# Define the desired column order
desired_order = [
    'ECOMM_ORD_NO', 'ECOMM_CATALOG_NO', 'ORDER_DATE', 'PAYMENT_DATE', 
    'REPORT_DATE', 'CUSTOMER_EMAIL', 'CUSTOMER_NAME', 'SKU', 
    'CATALOG_DESCRIPTION', 'QUANTITY', 'CURRENCY', 'PRICE', 
    'PRICE_TAX', 'CARGO_COST', 'CARGO_COST_TAX', 'CARGO_SRV_LEVEL', 
    'INVOICE_ADDRESS_1', 'INVOICE_ADDRESS_2', 'INVOICE_CITY', 
    'INVOICE_COUNTY', 'INVOICE_ZIP_CODE', 'INVOICE_COUNTRY', 
    'KARGO_PROMOSYON_INDIRIMI', 'CAROG_TRACK_NO', 'SALES_CHANNEL', 
    'STATE', 'DATE_ADDED', 'PROMOSYON_PRICE', 'CATALOG_NO', 'NET_PRICE'
]

# Reorder the columns in the DataFrame
df = df.reindex(columns=desired_order)

# Updated Process each row loop
for index, row in df.iterrows():
    # Apply abbreviation and remove city names for both addresses
    address_1 = abbreviate_address(row['INVOICE_ADDRESS_1'], abbreviations)
    address_1 = remove_city_names(address_1, turkish_cities)
    address_2 = abbreviate_address(row['INVOICE_ADDRESS_2'], abbreviations)
    address_2 = remove_city_names(address_2, turkish_cities)

    # Swap addresses if needed
    address_1, address_2 = swap_addresses_if_needed(address_1, address_2)

    # Split address if too long
    modified_address_1, modified_address_2 = split_address(address_1, address_2)

    # Check and modify INVOICE_ADDRESS_2 length if needed
    modified_address_2 = remove_vowels_to_fit(modified_address_2, 35)
    modified_address_2 = right_trim_to_fit(modified_address_2, 35)

    # Update the DataFrame with the modified values
    df.at[index, 'INVOICE_ADDRESS_1'] = modified_address_1
    df.at[index, 'INVOICE_ADDRESS_2'] = modified_address_2


# Save the processed data
df.to_excel('processed_file.xlsx', index=False, engine='openpyxl')

