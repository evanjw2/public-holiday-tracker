import requests
import pandas as pd
from bs4 import BeautifulSoup
from openpyxl import Workbook


workbook = Workbook()
url = f"https://www.officeholidays.com/countries/"
page = requests.get(url)
if page.status_code == 200:
    soup = BeautifulSoup(page.content, "html.parser")
    list = soup.find("div", class_="twelve columns")
    countries_link = list.find_all("a")
    for country_link in countries_link:
        country_linkurl = country_link["href"]

        url = f"{country_linkurl}"
        page = requests.get(url)
        if page.status_code == 200:
            soup = BeautifulSoup(page.content, "html.parser")
            table = soup.find("table", class_="country-table")

                # Create a DataFrame for the country's holiday data
            holidays_data = []
            holidays = table.find_all("tr")
            for holiday in holidays[1:]:  # Skipping the header row
                cells = holiday.find_all("td")
                date = cells[0].text.strip()
                name = cells[1].text.strip()
                holiday_type = cells[2].text.strip()
                comments = cells[3].text.strip()
                holidays_data.append([date, name + " 2024", holiday_type, comments])

            df = pd.DataFrame(holidays_data, columns=["Date", "Name", "Type", "Comments"])

                # Truncate sheet name if it exceeds 31 characters
            sheet_name = country_linkurl[41:].capitalize()[:31]

                # Create a worksheet for the country
            worksheet = workbook.create_sheet(title=sheet_name)

                # Write headers to worksheet
            worksheet.append(["Date", "Name", "Type", "Comments"])

                # Write data to worksheet
            for index, row in df.iterrows():
                worksheet.append(row.tolist())

                # Adjust column widths
            for column_cells in worksheet.columns:
                max_length = 0
                for cell in column_cells:
                    try:
                        if len(str(cell.value)) > max_length:
                            max_length = len(cell.value)
                    except:
                        pass
                adjusted_width = (max_length + 2) * 1.2  # Adding some extra space
                worksheet.column_dimensions[column_cells[0].column_letter].width = adjusted_width

# Remove the default sheet created by Workbook
workbook.remove(workbook["Sheet"])

# Save the Excel file
workbook.save(filename="public_holidays.xlsx")

#%%