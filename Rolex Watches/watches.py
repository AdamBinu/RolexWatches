import pandas as pd
import os
from jinja2 import Template

# Load the Excel file
df = pd.read_excel("watch_inventory.xlsx")  # Replace with your actual file name

# HTML template (Jinja2 format)
html_template = """ 
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>{{ name }}</title>
</head>
<body>
    <h1>{{ name }}</h1>
    <p><strong>Dial:</strong> {{ dial }}</p>
    <p><strong>Case Size:</strong> {{ case_size }}mm</p>
    <p><strong>Reference:</strong> {{ reference }}</p>
    <p><strong>Serial:</strong> {{ serial }}</p>
    <p><strong>Year of Production:</strong> {{ year_production }}</p>
    <p><strong>Papers/Card:</strong> {{ papers_card }}</p>
    <p><strong>Year on Card:</strong> {{ year_on_card }}</p>
    <p><strong>Cost:</strong> ${{ cost }}</p>
    <p><strong>Price:</strong> ${{ price }}</p>
    <p><strong>Location:</strong> {{ location }}</p>
    <p><strong>Misc:</strong> {{ misc }}</p>
</body>
</html>
"""

# Ensure output folder exists
output_folder = "generated_watch_pages"
os.makedirs(output_folder, exist_ok=True)

# Loop through each row and generate an HTML page
for index, row in df.iterrows():
    template = Template(html_template)
    
    # Render template with actual data
    html_content = template.render(
        name=row["Name"], 
        dial=row["Dial"], 
        case_size=row["Case Size"], 
        reference=row["Reference"], 
        serial=row["Serial"], 
        year_production=row["Year of Production"], 
        papers_card=row["Papers/Card?"], 
        year_on_card=row["Year on Card"] if pd.notna(row["Year on Card"]) else "N/A",
        cost=f"{row['Cost']:,}",  # Format with commas
        price=f"{row['Price']:,}",
        location=row["Location"], 
        misc=row["MISC"] if pd.notna(row["MISC"]) else "N/A"
    )
    
    # Define file name based on ID
    filename = f"{output_folder}/watch_{row['ID']}.html"
    
    # Save HTML file
    with open(filename, "w", encoding="utf-8") as file:
        file.write(html_content)

    print(f"Generated: {filename}")

print("All watch pages have been created successfully!")
