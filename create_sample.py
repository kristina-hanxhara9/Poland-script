#!/usr/bin/env python3
"""Helper script to create the sample_input.xlsx template."""

import pandas as pd

data = {
    "name": [
        "Castorama Polska Sp. z o.o.",
        "Leroy Merlin Polska Sp. z o.o.",
        "OBI Centrala Systemowa Sp. z o.o.",
        "Śnieżka SA",
        "Farby Kabe Polska Sp. z o.o.",
        "Dekoral Professional",
        "Grupa PSB Handel S.A.",
        "Bricoman Polska Sp. z o.o.",
        "Mrówka Market Budowlany",
        "Hurtownia Budowlana BAT",
        "Nomi S.A.",
        "Tikkurila Polska S.A.",
        "Skład Budowlany Materiały",
        "Bricomarché Polska",
        "Centrum Budowlane Bełchatów",
    ],
    "zip": [
        "02-255",
        "03-734",
        "02-672",
        "39-102",
        "32-060",
        "00-108",
        "25-323",
        "62-080",
        "35-001",
        "60-001",
        "43-300",
        "05-077",
        "30-001",
        "61-001",
        "97-400",
    ],
    "channel": [
        "DIY shops",
        "DIY shops",
        "DIY shops",
        "Paint specialists",
        "Paint specialists",
        "Paint specialists",
        "Builders merchants",
        "DIY shops",
        "DIY shops",
        "Builders merchants",
        "DIY shops",
        "Paint specialists",
        "Builders merchants",
        "DIY shops",
        "Builders merchants",
    ],
}

df = pd.DataFrame(data)
df.to_excel("sample_input.xlsx", index=False)
print(f"Created sample_input.xlsx with {len(df)} rows")
print(df.to_string(index=False))
