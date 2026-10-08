import urllib.parse
import json
import sys

def generate_research_queries(commodity, location, raw_materials=None, equipment=None):
    """
    Generates structured search queries and Google URLs for market verification.
    """
    if raw_materials is None:
        raw_materials = ["bahan baku utama", "gula", "minyak", "tepung", "susu"]
    if equipment is None:
        equipment = ["mesin produksi", "etalase", "chiller", "meja kursi"]

    encoded_loc = urllib.parse.quote_plus(location)
    
    queries = {
        "umr_and_salaries": [
            f"UMR UMK {location} terbaru keputusan gubernur",
            f"standar gaji karyawan {commodity} {location} jobstreet indeed",
            f"gaji kasir barista koki {location} glints"
        ],
        "property_and_rent": [
            f"sewa ruko tempat usaha {location} site:lamudi.co.id OR site:olx.co.id",
            f"harga sewa kios dekat kampus {location} rumah123",
            f"biaya renovasi interior cafe ruko per m2 {location}"
        ],
        "utilities": [
            "tarif listrik komersial B-1 B-2 PLN per kWh",
            "tarif air bersih pdam komersial per m3",
            "harga isi ulang tabung gas 12kg agen pangkalan resmi"
        ],
        "raw_materials": [
            f"harga grosir {mat} sak karton site:tokopedia.com" for mat in raw_materials
        ],
        "equipment": [
            f"harga komersial {eq} distributor f&b" for eq in equipment
        ],
        "competitors": [
            f"daftar kedai usaha {commodity} di {location} google maps review",
            f"kompetitor bisnis {commodity} terbaik {location}"
        ]
    }
    return queries

def print_research_sheet(commodity, location):
    queries = generate_research_queries(commodity, location)
    print(f"============================================================")
    print(f" SKB MARKET VERIFICATION QUERIES: {commodity.upper()} @ {location.upper()}")
    print(f"============================================================\n")
    for cat, q_list in queries.items():
        print(f"--- KATEGORI: {cat.upper().replace('_', ' ')} ---")
        for q in q_list:
            search_url = f"https://www.google.com/search?q={urllib.parse.quote_plus(q)}"
            print(f"  * Query : {q}")
            print(f"    Link  : {search_url}")
        print()

if __name__ == "__main__":
    com = sys.argv[1] if len(sys.argv) > 1 else "Kopi & Minuman Kekinian"
    loc = sys.argv[2] if len(sys.argv) > 2 else "Tanjung Duren Jakarta Barat"
    print_research_sheet(com, loc)
