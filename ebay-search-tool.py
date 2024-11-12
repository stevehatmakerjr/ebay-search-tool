import requests
import pandas as pd
from datetime import datetime, timedelta
from dateutil import parser
import pytz 

def get_api_response(app_id, api_endpoint, params):
    """Makes the API request and returns the JSON response."""
    response = requests.get(api_endpoint, params=params)
    print("Status Code:", response.status_code)  # Debugging statement

    if response.status_code != 200:
        print(f"HTTP Error: {response.status_code}")
        exit()

    try:
        data = response.json()
    except ValueError as e:
        print("Error parsing JSON response:", e)
        exit()

    return data

def parse_items(data):
    """Parses the JSON response and returns a list of items."""
    if 'findItemsAdvancedResponse' in data:
        response_data = data['findItemsAdvancedResponse'][0]
        if 'errorMessage' in response_data:
            errors = response_data['errorMessage'][0]['error']
            for error in errors:
                print(f"Error Code: {error['errorId'][0]}, Message: {error['message'][0]}")
            exit()
        else:
            items = response_data['searchResult'][0].get('item', [])
            items = items[:50]  # Limit to 50 items - limit is not hard set 
            if not items:
                print("No items found matching your criteria.")
                exit()
            return items
    else:
        print("Unexpected response format.")
        exit()

def process_item(item):
    """Processes an individual item and returns a dictionary of extracted data."""
    title = item.get('title', ['N/A'])[0]
    url = item.get('viewItemURL', ['N/A'])[0]
    # Extract price
    price_info = item.get('sellingStatus', [{}])[0].get('currentPrice', [{}])[0]
    price = price_info.get('__value__', 'N/A')
    currency = price_info.get('@currencyId', 'N/A')
    # Extract shipping price
    shipping_info = item.get('shippingInfo', [{}])[0]
    shipping_cost_info = shipping_info.get('shippingServiceCost', [{}])[0]
    shipping_price = shipping_cost_info.get('__value__', 'N/A')
    if shipping_price in ['0.0', '0.00', '0']:
        shipping_price = 'FREE'
    elif shipping_price != 'N/A':
        shipping_price = f"{float(shipping_price):.2f}"
    # Extract listing type
    listing_type = item.get('listingInfo', [{}])[0].get('listingType', ['N/A'])[0]
    # Extract item condition
    condition = item.get('condition', [{}])[0].get('conditionDisplayName', ['N/A'])[0]
    # Extract end time
    end_time_str = item.get('listingInfo', [{}])[0].get('endTime', [''])[0]
    if end_time_str:
        # Parse the end time string to a datetime object in UTC
        end_time_utc = parser.parse(end_time_str)
        # Convert to CST timezone
        # Change to any timezone
        cst = pytz.timezone('US/Central')
        end_time_cst = end_time_utc.astimezone(cst)
        # Format the datetime as a string in 12-hour format with AM/PM
        end_time_formatted = end_time_cst.strftime('%Y-%m-%d %I:%M:%S %p')
    else:
        end_time_formatted = 'N/A'
    # Return the processed item data
    return {
        'Title': title,
        'Price': price,
        'Currency': currency,
        'Shipping Price': shipping_price,
        'Listing Type': listing_type,
        'Item Condition': condition,
        'End Time': end_time_formatted,
        'URL': url
    }

def save_to_excel(results, filename='ebay_listings.xlsx'):
    """Saves the results to an Excel file with conditional formatting."""
    import pandas as pd

    df = pd.DataFrame(results)
    if df.empty:
        print("No data to save to Excel.")
        exit()
    else:
        # Reorder columns as per your request
        df = df[['Title', 'Price', 'Currency', 'Shipping Price', 'Listing Type',
                 'Item Condition', 'End Time', 'URL']]

        # Use ExcelWriter with xlsxwriter engine
        with pd.ExcelWriter(filename, engine='xlsxwriter') as writer:
            df.to_excel(writer, index=False, sheet_name='Sheet1')
            workbook = writer.book
            worksheet = writer.sheets['Sheet1']

            # Define formats
            green_format = workbook.add_format({'bg_color': '#C6EFCE',
                                                'font_color': '#006100'})
            yellow_format = workbook.add_format({'bg_color': '#FFEB9C',
                                                 'font_color': '#9C6500'})

            # Get the dimensions of the dataframe
            max_row, max_col = df.shape

            # Convert column index to Excel column letter
            def colnum_string(n):
                string = ''
                while n >= 0:
                    n, remainder = divmod(n, 26)
                    string = chr(65 + remainder) + string
                    n -= 1
                return string

            # Get the column letter for "Listing Type"
            listing_type_col_idx = df.columns.get_loc('Listing Type')
            listing_type_col_letter = colnum_string(listing_type_col_idx)

            # Define the cell range for conditional formatting
            start_row = 1  # Row 0 is the header
            end_row = max_row  # Number of data rows
            cell_range = f'{listing_type_col_letter}{start_row + 1}:{listing_type_col_letter}{end_row + 1}'

            # Apply green format to 'FixedPrice'
            worksheet.conditional_format(cell_range, {'type': 'text',
                                                      'criteria': 'containing',
                                                      'value': 'FixedPrice',
                                                      'format': green_format})

            # Apply yellow format to 'Auction'
            worksheet.conditional_format(cell_range, {'type': 'text',
                                                      'criteria': 'containing',
                                                      'value': 'Auction',
                                                      'format': yellow_format})

            # Adjust column widths
            for idx, col in enumerate(df.columns):
                # Get the maximum length of the data in the column
                max_len = df[col].astype(str).map(len).max()
                # Set the column width
                worksheet.set_column(idx, idx, max_len + 5)

            print(f"Data saved to {filename} with conditional formatting.")

def main():
    # Your eBay App ID (Client ID)
    APP_ID = 'EBAY_APP_ID'  # Replace with your actual App ID

    # eBay Finding API endpoint
    API_ENDPOINT = 'https://svcs.ebay.com/services/search/FindingService/v1'

    # Keyword to search for
    KEYWORD = 'Nintendo 64 Game Console'  # Replace with your desired keyword

    # Calculate the end time (current time + 24 hours) in ISO 8601 format
    end_time_to = (datetime.utcnow() + timedelta(hours=24)).strftime('%Y-%m-%dT%H:%M:%SZ')

    # Set up parameters for the API request
    params = {
        'OPERATION-NAME': 'findItemsAdvanced',
        'SERVICE-VERSION': '1.0.0',
        'SECURITY-APPNAME': APP_ID,
        'RESPONSE-DATA-FORMAT': 'JSON',
        'keywords': KEYWORD,
        'paginationInput.entriesPerPage': '50',  # Limit to 50 items per page
        'itemFilter(0).name': 'EndTimeTo',
        'itemFilter(0).value': end_time_to,
        # Include both Auction and Buy It Now listings
        'itemFilter(1).name': 'ListingType',
        'itemFilter(1).value(0)': 'Auction',
        'itemFilter(1).value(1)': 'FixedPrice',
        'sortOrder': 'EndTimeSoonest',
    }

    # Get API response
    data = get_api_response(APP_ID, API_ENDPOINT, params)

    # Parse items from response
    items = parse_items(data)

    # Process each item
    results = []
    for item in items:
        processed_item = process_item(item)
        results.append(processed_item)

    # Save results to Excel
    save_to_excel(results)

if __name__ == '__main__':
    main()

