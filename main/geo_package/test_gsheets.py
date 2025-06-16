import asyncio
# test for Gsheet column generated: setting of hyperlink column
import os, sys
main_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.append(main_dir)
from config import SHEETS_KEY
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import pandas as pd
from geo_scraper_regex import GeoScrapper

sheets_key = SHEETS_KEY
SCOPES = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive', 'https://www.googleapis.com/auth/spreadsheets']
credentials = ServiceAccountCredentials.from_json_keyfile_dict(sheets_key, SCOPES)
client = gspread.authorize(credentials)
gsheet_link ='https://docs.google.com/spreadsheets/d/1sgqVH-iOmJpAWwl3fGeLsyZCsz8RMh-TaTT1BbtvOq8/edit?usp=sharing'
url_provided = 'https://www.georgiapublicnotice.com/(S(41gfwvxnbs4n4zyaueam3d2y))/Details.aspx?SID=41gfwvxnbs4n4zyaueam3d2y&ID=3641052'
worksheet = client.open_by_url(gsheet_link).sheet1
print("Worksheet opened")

COLUMN_ALIAS = 'F'  # Column for aliases
COLUMN_URL = 'E'    # Column for URLs

# Get the values in the alias column
alias_values = worksheet.col_values(ord(COLUMN_ALIAS) - ord('A') + 1)[1:]

def set_to_hyperlink():
    # Generate hyperlinks and update the URL column
    for i in range(len(alias_values)):
        alias = alias_values[i]
        if alias:  # Skip empty cells
            hyperlink_formula = f'=HYPERLINK("{url_provided}", "{alias}")'
            worksheet.update_cell(i + 2, ord(COLUMN_URL) - ord('A') + 1, hyperlink_formula)


set_to_hyperlink()


def test_sheet_creation(scraper: GeoScrapper):
    """Test function to create a Google Sheet with sample data without scraping."""
    # Create sample data
    sample_data = {
        'First name': ['John', 'Jane'],
        'Last name': ['Doe', 'Smith'],
        'Mortgage Balance': ['$200,000', '$300,000'],
        'Auction Date': ['June 1, 2025', 'June 15, 2025'],
        'Link to Foreclosure': ['http://example.com/1', 'http://example.com/2'],
        'Property Address': ['123 Main St', '456 Oak Ave'],
        'City': ['Atlanta', 'Marietta'],
        'State': ['GA', 'GA'],
        'Zipcode': ['30301', '30060'],
        'Loopnet Search': ['test', 'test'],
        'Crexi Search': ['test', 'test'],
        'Zillow Search': ['test', 'test'],
    }

    # Create sample search URLs
    # for name, address in zip(sample_data['First name'], sample_data['Property Address']):
    #     full_address = f"{address}, {sample_data['City'][sample_data['First name'].index(name)]}, {sample_data['State'][sample_data['First name'].index(name)]} {sample_data['Zipcode'][sample_data['First name'].index(name)]}"
    #     for site in scraper.search_sites:
    #         site_key = f"{site.split('.')[0].title()} Search"
    #         if site_key not in sample_data:
    #             sample_data[site_key] = []
    #         sample_data[site_key].append(scraper.create_google_search_url(site, full_address))

    text = """
    5d43c2d4-bc6f-4cac-b6a7-a8299a83f0b3Commercial use of this website is strictly prohibited.Contact the administrator with any questions.1afc856d-7ae7-42d7-b101-8dd0a235f61c
gpn11 MDJ-4294 MDJ-4294 GPN-11 State of Georgia County of Cobb Notice of Sale Under Power By virtue of the power of sale contained in that Deed to Secure Debt, Assignment of Rents and Leases, Security Agreement, and Fixture Filing from WEB IV, LLC ("WEB IV" or "Grantor"),  held by SouthState Bank, N.A., f/k/a CenterState Bank, N.A., f/k/a (Private Bank of Buckhead, a division of National Bank of Commerce) ("SSB"), dated January 23, 2017, recorded January 30, 2017, in Deed Book 15414, Page 3726, records of Cobb County, Georgia, as modified or amended, ("Security Deed"), securing a Promissory Note and Construction and Term Loan Agreement, dated January 23, 2017, as modified and amended (hereinafter the "Note") there will be sold by the undersigned at public outcry to the highest bidder for cash before the Courthouse at Cobb County, Georgia, within the legal hours of sale on Tuesday, July 1, 2025, the property described on Exhibit "A" attached hereto and incorporated herein by reference,  including all of the following: Together with all buildings, structures and improvements of every nature whatsoever now or hereafter situated on, under or above the Land (the "Improvements"), and all goods, inventory, machinery, equipment, fixtures (including, without limitation, all heating, air conditioning, plumbing, lighting, communications and elevator fixtures), furnishing, building supplies and materials, and all other personal property of every kind and nature whatsoever owned by Grantor (or in which Grantor has or hereafter acquires an interest) and now or hereafter located upon, or appurtenant to, the Land or the Improvements or used or useable in the present or future operation and occupancy of the Land or the Improvements, along with all accessions, replacements or substitutions of all or any portion thereof; all of which are hereby declared and shall be deemed to be fixtures and accessions to the Land and a part of the Premises as between the parties hereto and all persons claiming by, through or under them, and which shall be deemed to be a portion of the security for the indebtedness herein described and to be secured by this Instrument. Together with all machinery, equipment, vehicles, and other personal property of Grantor either located on or used in connection with the Land (the "Personal Property"); Together with all right, title and interest of Grantor in and to all policies of insurance and all condemnation proceeds, which in any way now or hereafter belong, relate, or appertain to the Land, the Improvements, or the Personal Property, or any part thereof; Together with all income, rents, issues, profits, revenues, deposits, accounts and other benefits belonging or payable to Grantor from the operation of the Land or the Improvements, including, without limitation, all revenues, all receivables, customer obligations, installment payment obligations and other obligations now existing or hereafter arising or created out of sale, lease, sublease, license, concession, or other grant of the right of the possession, use or occupancy of all or any portion of the Land, the Improvements or personalty located thereon, and all proceeds, if any, from business interruption or other loss of income insurance relating to the use, enjoyment or occupancy of the Land and/or the Improvements whether paid or accruing before or after the filing by or against Grantor of any petition for relief under any present or future state or federal law regarding bankruptcy (each a "Bankruptcy Code"), reorganization or other relief to debtors and all proceeds from the sale or other disposition of the Tenant Leases (hereinafter defined). All leases, subleases, licenses and other agreements granting others the right to use or occupy all or any part of the Land or the Improvements together with all restatements, renewals, extensions, amendments and supplements thereto, (collectively, the "Tenant Leases"), now existing or hereafter entered into, and whether entered before or after the filing by or against Grantor of any petition for relief under a Bankruptcy Code, and all of Grantor's right, title and interest in the Tenant Leases, including, without limitation (i) all guarantees, letters of credit and any other credit support given by any tenant or guarantor in collection therewith (collectively, the "Tenant Lease Guaranties"), (ii) all cash, notes, or security deposited thereunder to secure the performance by the tenants of their obligations thereunder (collectively, the "Tenant Security Deposits"), (iii) all claims and rights to the payment of damages and other claims arising from any rejection by a tenant of its Tenant Lease under a Bankruptcy Code ("Bankruptcy Claims"), (iv) all of the landlord's rights in casualty or condemnation proceeds of a tenant in respect of the leased premises (collectively, the "Tenant Claims"), (v) all rents, ground rents, additional rents, revenues, termination and similar payments, issues and profits (including all oil and gas or other mineral royalties and bonuses) from the Land or the Improvements (collectively with the Tenant Lease Guaranties, Tenant Security Deposits, Bankruptcy Claims and Tenant Claims, the "Rents"), whether paid or accruing before or after the filing by or against Borrower of any petition for relief under a Bankruptcy Code, (vi) all proceeds or streams of payment from the sale or other disposition of the Tenant Leases or disposition of any Rents, and (vii) the right to receive and apply the Rents to the payment of Borrower's obligations to Lender and to do all other things which Borrower or a lessor is or may become entitled to do under the Tenant Leases or with respect to the Rents; Together with all agreements, equipment leases, contracts (including, without limitation, any construction, architectural, service, supply and maintenance contracts), equipment leases, registrations, permits, licenses, franchise agreements, plans, specifications and other documents, now or hereafter entered into, and all rights therein and thereto, respecting or pertaining to the use, occupation, construction, management or operation of the Land or the Improvements, or respecting any business or activity conducted from the Land or the Improvements, and all right, title and interest of Grantor therein and thereunder, including, without limitation, the right, while an Event of Default remains uncured, to receive and collect any sums payable to Grantor thereunder; Together all deposits (including but not limited to Grantor's rights in Tenant Security Deposits, deposits with respect to utility services to the Land and Improvements, and any deposits or reserves hereunder or under any other Loan Documents for taxes, insurance or otherwise), rebates or refunds of impact fees, taxes, assessments or charges, and all other contracts, purchase agreements, instruments and documents as such may arise from or be related to the Land and Improvements; Together with all accounts, escrows, chattel paper, claims, deposits, trade names, trademarks, service marks, logos. copyrights, goodwill, licenses, permits, plans and specifications, environmental audits, engineering reports, warranties, guaranties, books and records and all other general intangibles and payment intangibles owned by Grantor relating to or used in collection with the operation of the Land or the Improvements; Together with all reserves, escrows and deposit accounts maintained by Grantor or with respect to the Land or the Improvements, together with all cash, checks, drafts, certificates, accounts receivable, documents, letter of credit rights, commercial tort claims, securities, investment property, financial assets, instruments and other property from time to time held therein, and all proceeds, products, distributions, dividends or substitutions thereon or thereof; Together with all permits, licenses, franchises, certificates, development rights, commitments and rights for utilities, and other rights and privileges obtained in connection with the Land and Improvements; Together with all oil, gas and other hydrocarbons and other minerals produced from or allocated to the Land and all products processed or obtained therefrom, and the proceeds thereof; Together with all proceeds, products, substitutions, and accessions of the foregoing of every type. The Security Deed secures a Judgment on the Note, plus post Judgment interest, entered in Civil Action No. 25GC00661 on March 26, 2025 by the Superior Court of Cobb County in favor of SSB (hereinafter the "Debt").  The Debt, secured by the Security Deed, has been and is hereby declared due and payable, and all required notices have been given. The Debt remaining in default, this sale will be made by non-judicial foreclosure as provided for in the Security Deed, for the purpose of paying the Debt and all expenses of this sale, including attorney's fees. Said property will be sold subject to any outstanding ad valorem taxes (including taxes which are a lien, but not yet due and payable), any matters which might be disclosed by an accurate survey and inspection of the property, any assessments, liens, encumbrances, zoning ordinances, restrictions, easements, covenants, and matters of record superior to the Security Deed first set out above, including, but not necessarily limited to, senior encumbrances that will not be extinguished by the foreclosure sale contemplated by this notice. 	Please note that SSB is the entity that has full authority to negotiate, amend, and modify all terms of the Debt and Security Deed, and that SSB can be contacted through the following representative: Neal J. Quirk, Esq. Quirk & Parnell, LLC 6000 Lake Forrest Drive NW, Ste. 300 Atlanta, Georgia, 30328 Email: njq@quirklaw.com Telephone: (404) 252-1425  To the best knowledge and belief of the undersigned, the party or parties in possession of the property is WEB IV, LLC, and said property is more commonly known as 1005 Cobb Place Blvd., NW, Kennesaw, GA 30144. However, please rely only on the legal description contained in this notice for the location of the property. The undersigned reserves the right to sell the property, or any part thereof, together in its entirety or in one or more parcels as determined by the undersigned in its sole discretion. SouthState Bank, N.A. As Attorney-in-Fact for WEB IV, LLC Under Power of Sale in the Security Deed  Neal J. Quirk, Esq. Quirk & Parnell, LLC 6000 Lake Forrest Drive NW, Ste. 300 Atlanta, Georgia, 30328 Email: njq@quirklaw.com Telephone: (404) 252-1425 Run Dates: 06/06, 06/13, 06/20, 06/27, 2025 EXHIBIT "A" LEGAL DESCRIPTION All that tract or parcel of land lying in and being a portion of Land Lot 173 of the 20th District, 2nd Section, Cobb County, Georgia, being more fully and particularly described as follows: To find the TRUE POINT OF BEGINNING commence at the intersection of the Southeasterly right of way line of Cobb Place Boulevard (120-foot right of way) with the Northwesterly end of a miter at the Northerly right of way line of Roberts Boulevard; thence Northeasterly along the said right of way line of Cobb Place Boulevard a distance of 1192.65 fee to a corner marked by a 1/2-inch reinforcing rod and the TRUE POINT OF BEGINNING; thence continuing Northeasterly along the said right of way line along a curve to the left an arc distance of 167.17 feet (said arc subtended by a chord of North 43 degrees, 24 minutes, 22 seconds East with a length of 166.87 feet) to a point; thence North 37 degrees, 29 minutes, 43 seconds East continuing along the said right of way line a distance of 263.69 feet to a corner marked by a 1/2-inch reinforcing rod; thence leaving the said right of way line South 44 degrees, 38 minutes, 38 seconds East a distance of 827.95 to a corner marked by a 1/2-in reinforcing rod; thence South 71 degrees, 14 minutes, 00 seconds west a distance of 590.52 feet to a corner marked by a 1/2-inch reinforcing rod; thence North 43 degrees, 40 minutes, 44 seconds West a distance of 211.23 fee to a corner marked by a 1/2-inch reinforcing rod; thence North 46 degrees, 51 minutes, 35 seconds East a distance of 87.99 feet to a corner marked by a 1/2-inch reinforcing rod; thence North 43 degrees, 25 minutes, 20 seconds West a distance of 23.04 feet to a corner marked by a PK nail; thence North 43 degrees, 26 minutes, 52 seconds West a distance of 179.98 feet to a corner marked by a PK nail; thence North 78 degrees, 54 minutes, 12 seconds West a distance of 26.02 feet to a corner marked by a PK nail; thence North 27 degrees, 33 minutes, 08 seconds West a distance of 99.33 feet to a corner at the Southeasterly right of way line of Cobb Place Boulevard and the TRUE POINT OF BEGINNING. Said tract contains 7.512 acres and is delineated on an ALTA/NSPS Land Title Survey Plat prepared for WEB IV, LLC, First American Title Insurance Company, Private /Bank of Buckhead and ISAOA, dated 1/12/2017 prepared by Crusselle Company bearing the stamp and seal of Benjamin W. Crusselle, GRLS No. 2841, said survey being incorporated herein by reference.  Parcel ID: 20017300590 6:6,13,20,27-2025"""
    url = "https://www.georgiapublicnotice.com/(S(41gfwvxnbs4n4zyaueam3d2y))/Details.aspx?SID=41gfwvxnbs4n4zyaueam3d2y&ID=3643477"

    # Create DataFrame from sample data
    scraper.data_frame = pd.DataFrame(sample_data)
    asyncio.run(scraper.data_cleaner(text, url, 1))

    # Create and format worksheet
    result_url = asyncio.run(scraper.create_worksheet())
    return result_url


# Function to run the test
def gsheet_test():
    print("Running gsheet_test...")
    scraper = GeoScrapper(1, "06/13/2025", "06/14/2025")
    result_url = test_sheet_creation(scraper)
    print(f"Test spreadsheet created: {result_url}")


if __name__ == "__main__":
    gsheet_test()
