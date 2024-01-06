import pandas as pd
from datetime import datetime
import re

# Read the Excel file
df = pd.read_excel('AMZ_G_SIP.xlsx')

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
    41: 'KARGO_PROMOSYON_INDIRIMI', # Kargo Promosyon İndirimi
    43: 'CAROG_TRACK_NO',           # Takip Numarası
    46: 'SALES_CHANNEL'             # Satış Kanalı
}

# Rename columns based on their positions
for position, new_name in column_position_mapping.items():
    df.columns.values[position] = new_name

# Filter out unwanted columns
df = df[column_position_mapping.values()]

# Add new columns
df['STATE'] = ''  # Empty column
df['DATE_ADDED'] = datetime.now().strftime('%d/%m/%Y')  # Add today's date in DD/MM/YYYY format

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

    # Replace Turkish characters with English equivalents and convert to uppercase
    for tr_char, en_char in char_mapping.items():
        name = name.replace(tr_char, en_char)

    return name.upper()

def adjust_zip_code(zip_code):
    zip_code_str = str(zip_code)  # Ensure it's a string
    while len(zip_code_str) < 5:
        zip_code_str = '0' + zip_code_str
    return zip_code_str

def convert_date_format(date_str):
    # Parse the date string to a datetime object
    date_obj = pd.to_datetime(date_str)

    # Format the datetime object to the desired string format
    return date_obj.strftime('%d/%m/%Y')

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
df['DATE_ADDED'] = df['DATE_ADDED'].apply(convert_date_format)
df['INVOICE_CITY'] = df['INVOICE_CITY'].str.upper()

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

