import pandas as pd

# Load the dataframe if it's not already loaded. Assuming the file is in the content directory.
try:
  df = pd.read_excel(r'C:\Users\mahmoud.aboelhadid\Documents\UiPath\Excel Manipulation\Excel Manipulation\Output\Employees_Cleaned.xlsx')
except FileNotFoundError:
  print("Error: 'Employees_Cleaned.xlsx' not found. Please upload the file to the /content/ directory.")


# Count the number of unique companies
unique_companies_count = df['Company'].nunique()
print(f"Number of unique companies: {unique_companies_count}")

# Find the top 5 most common email domains
# Extract email domains
df['EmailDomain'] = df['Email'].apply(lambda x: x.split('@')[1])

# Count the frequency of each domain
email_domain_counts = df['EmailDomain'].value_counts().reset_index()
email_domain_counts.columns = ['EmailDomain', 'NumberOfEmployees']

# Get the top 5 domains
top_5_email_domains = email_domain_counts.head(5)
print("\nTop 5 most common email domains:")
print(top_5_email_domains.to_string(index=False))

# Count the number of employees per company
company_counts = df['Company'].value_counts().reset_index()
company_counts.columns = ['Company', 'NumberOfEmployees']
print("\nNumber of employees per company:")
print(company_counts.to_string(index=False))

# Save the results into Summary_Report.csv
# Combine the relevant results into a dictionary or list of dictionaries
summary_data = {
    'Metric': ['Unique Companies', 'Top 5 Most Common Email Domains', 'Employee Count per Company'],
    'Value': [
        unique_companies_count,
        " | ".join([f"{row['EmailDomain']} ({row['NumberOfEmployees']})" for index, row in top_5_email_domains.iterrows()]),
        " | ".join([f"{row['Company']} ({row['NumberOfEmployees']})" for index, row in company_counts.iterrows()])
    ]
}

# Create a DataFrame from the summary data
summary_df = pd.DataFrame(summary_data)

# Save the DataFrame to a CSV file
summary_df.to_csv('Summary_Report.csv', index=False)
print("\nSummary report saved to Summary_Report_Python2.csv")