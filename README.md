# Marketing Channels Data Processing Script

This Python script is designed to aggregate and organize marketing channel data from a CSV file, specifically formatted for the SimilarWeb platform. The script prompts the user to input a list of websites, then processes the CSV data to create separate Excel files and tables for various marketing channels, such as Direct, Organic Search, Paid Search, Display Ads, Email, Referrals, and Social.

## Instructions for Use

### 1. List of Websites:
- The script begins by asking the user to paste a list of websites, with each website on a new line.

    ```python
    websites = input("Please paste the list of websites, separated by a new line: ").split("\n")
    ```

### 2. CSV Data:
- The script reads marketing channel data from a CSV file named "similarweb-results_marketing-channels.csv" in the same folder. Ensure that the CSV file is correctly formatted with columns for "Date," "Channel," "Device," and "Visits."

    ```python
    csv_data = pd.read_csv("similarweb-results_marketing-channels.csv")
    ```

### 3. Data Aggregation:
- The script creates separate DataFrames for each marketing channel (Direct, Organic, Paid Search, Display Ads, Email, Referrals, Social) for both "Desktop" and "Mobile Web."

    ```python
    df_direct = populate_dataframe(df_direct, "Direct", "Desktop")
    df_direct = populate_dataframe(df_direct, "Direct", "Mobile Web")
    # ... Repeat for other marketing channels ...
    ```

### 4. Combining Data:
- The script combines data for "Desktop" and "Mobile Web" for each marketing channel.

    ```python
    df_direct = df_direct.groupby("Date", as_index=False).sum()
    # ... Repeat for other marketing channels ...
    ```

### 5. Save to Excel:
- The aggregated data is saved to an Excel file named "output.xlsx" with separate sheets for each marketing channel.

    ```python
    with pd.ExcelWriter("output.xlsx") as writer:
        df_direct.to_excel(writer, sheet_name="Direct", index=False)
        # ... Repeat for other marketing channels ...
    ```

### 6. Completion Message:
- Finally, a message is printed to indicate that the data has been saved.

    ```python
    print("Data has been saved to output.xlsx file.")
    ```

## Important Note
- Ensure that the required Python packages (`pandas`) are installed before running the script.

Feel free to use and modify this script according to your specific needs. If you encounter any issues or have suggestions for improvement, please feel free to open an issue or contribute to the code. Happy data aggregating!

**Author**: @dlysh
