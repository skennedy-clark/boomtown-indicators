from bs4 import BeautifulSoup

base_url = "https://www.qgso.qld.gov.au"

with open("input.html", "r", encoding="utf-8") as f:
    soup = BeautifulSoup(f, "html.parser")

rows = soup.find_all("tr")

output = []

for row in rows:
    name_tag = row.find("div", class_="field--name-field-short-title")
    link_tag = row.find("a", href=True)

    if name_tag and link_tag:
        name = name_tag.get_text(strip=True)
        link = base_url + link_tag["href"]
        output.append(f"{name},{link}")

# write to file
with open("output.csv", "w", encoding="utf-8") as f:
    f.write("\n".join(output))

print(f"Extracted {len(output)} rows")