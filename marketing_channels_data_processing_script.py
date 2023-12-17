# -*- coding: utf-8 -*-
"""
Created on Thu Jul 27 10:40:12 2023

@author: dlysh
"""

import pandas as pd

# 1) Ask the user to paste a list of websites
websites = input("Please paste the list of websites, separated by a new line: ").split("\n")

# 2) Create Excel files and tables for all marketing channels
df_direct = pd.DataFrame(columns=["Date"] + websites)
df_organic = pd.DataFrame(columns=["Date"] + websites)
df_paid_search = pd.DataFrame(columns=["Date"] + websites)
df_display_ads = pd.DataFrame(columns=["Date"] + websites)
df_email = pd.DataFrame(columns=["Date"] + websites)
df_referrals = pd.DataFrame(columns=["Date"] + websites)
df_social = pd.DataFrame(columns=["Date"] + websites)

# 3) Open the CSV file from the same folder as the script
csv_data = pd.read_csv("similarweb-results_marketing-channels.csv")

# Function to populate DataFrames for each marketing channel
def populate_dataframe(df, channel, device):
    channel_data = csv_data[(csv_data["Channel"] == channel) & (csv_data["Device"] == device)]
    unique_dates = channel_data["Date"].unique()
    for date in unique_dates:
        df_row = {"Date": date}
        date_data = channel_data[channel_data["Date"] == date]
        for website in websites:
            website_data = date_data[date_data["Domain"] == website]
            value = website_data["Visits"].sum()
            df_row[website] = value
        df = pd.concat([df, pd.DataFrame([df_row])], ignore_index=True)
    return df

# 4) Populate DataFrames for all marketing channels and combine Desktop and Mobile Web data
df_direct = populate_dataframe(df_direct, "Direct", "Desktop")
df_direct = populate_dataframe(df_direct, "Direct", "Mobile Web")

df_organic = populate_dataframe(df_organic, "Organic Search", "Desktop")
df_organic = populate_dataframe(df_organic, "Organic Search", "Mobile Web")

df_paid_search = populate_dataframe(df_paid_search, "Paid Search", "Desktop")
df_paid_search = populate_dataframe(df_paid_search, "Paid Search", "Mobile Web")

df_display_ads = populate_dataframe(df_display_ads, "Display Ads", "Desktop")
df_display_ads = populate_dataframe(df_display_ads, "Display Ads", "Mobile Web")

df_email = populate_dataframe(df_email, "Email", "Desktop")
df_email = populate_dataframe(df_email, "Email", "Mobile Web")

df_referrals = populate_dataframe(df_referrals, "Referrals", "Desktop")
df_referrals = populate_dataframe(df_referrals, "Referrals", "Mobile Web")

df_social = populate_dataframe(df_social, "Social", "Desktop")
df_social = populate_dataframe(df_social, "Social", "Mobile Web")

# 5) Merge DataFrames to combine "Desktop" and "Mobile Web" data for each marketing channel
df_direct = df_direct.groupby("Date", as_index=False).sum()
df_organic = df_organic.groupby("Date", as_index=False).sum()
df_paid_search = df_paid_search.groupby("Date", as_index=False).sum()
df_display_ads = df_display_ads.groupby("Date", as_index=False).sum()
df_email = df_email.groupby("Date", as_index=False).sum()
df_referrals = df_referrals.groupby("Date", as_index=False).sum()
df_social = df_social.groupby("Date", as_index=False).sum()

# 6) Save resulting DataFrames to Excel file
with pd.ExcelWriter("output.xlsx") as writer:
    df_direct.to_excel(writer, sheet_name="Direct", index=False)
    df_organic.to_excel(writer, sheet_name="Organic", index=False)
    df_paid_search.to_excel(writer, sheet_name="Paid Search", index=False)
    df_display_ads.to_excel(writer, sheet_name="Display Ads", index=False)
    df_email.to_excel(writer, sheet_name="Email", index=False)
    df_referrals.to_excel(writer, sheet_name="Referrals", index=False)
    df_social.to_excel(writer, sheet_name="Social", index=False)

print("Data has been saved to output.xlsx file.")