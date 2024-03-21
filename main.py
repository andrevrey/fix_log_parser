"""
Processes Financial Information eXchange (FIX) messages from text files.
Extracts message components and saves the data into an Excel (xlsx) file.
Dynamically handles known and unknown FIX tags, ensuring comprehensive data capture.
"""

from datetime import datetime
import os
import pandas as pd
from openpyxl.utils import get_column_letter

# Define mapping of FIX tag numbers to human-readable column names
columns_mapping = {
    'TID': 'TID',
    'Order Type': 'Order Type',
    '8': '8 BeginString',
    '9': '9 BodyLength',
    '35': '35 MsgType',
    '34': '34 MsgSeqNum',
    '49': '49 SenderCompID',
    '56': '56 TargetCompID',
    '57': '57 TargetSubID',
    '52': '52 SendingTime',
    '11': '11 ClOrdID',
    '17': '17 ExecID',
    '37': '37 OrderID',
    '198': '198 SecondaryOrderId',
    '150': '150 ExecType',
    '453': '453 NoPartyIDs',
    '448': '448 PartyID',
    '447': '447 PartyIDSource',
    '452': '452 PartyRole',
    '55': '55 Symbol',
    '48': '48 SecurityID',
    '22': '22 SecurityIDSource',
    '762': '762 SecuritySubType',
    '1': '1 Account',
    '14': '14 CumQty',
    '31': '31 LastPx',
    '32': '32 LastQty',
    '38': '38 OrderQty',
    '110': '110 MinQty',
    '39': '39 OrdStatus',
    '40': '40 OrdType',
    '44': '44 Price',
    '847': '847 TargetStrategy',
    '54': '54 Side',
    '59': '59 TimeInForce',
    '60': '60 TransactTime',
    '75': '75 TradeDate',
    '64': '64 SettlDate',
    '151': '151 LeavesQty',
    '880': '880 TrdMatchID',
    '1891': '1891 TrdMatchSubID',
    '1057': '1057 AggressorIndicator',
    '381': '381 GrossTradeAmt',
    '797': '797 CopyMsgIndicator',
    '10': '10 CheckSum',
}

def extract_preamble_details(line):
    """
    Extracts the transaction ID (TID) and Order Type from the line preamble.
    
    Parameters:
    - line (str): The full line from the log file.
    
    Returns:
    - tuple: (TID, Order Type)
    """
    preamble, _ = line.split(']: ', 1)
    tid_marker = preamble.find("(TID=")
    if tid_marker != -1:
        tid_end = preamble.find(")", tid_marker) + 1
        tid = preamble[:tid_end].strip()
    else:
        tid = "Unknown TID"
        
    order_type_marker = preamble.find("| ")
    if order_type_marker != -1:
        order_type = preamble[order_type_marker+2:].strip()
    else:
        order_type = "Unknown Order Type"
    
    return tid, order_type

def parse_message(line):
    """
    Parses a single FIX message line, extracting both known and unknown tag values
    and handling repeated '448', '447', and '452' tags specifically.
    
    Parameters:
    - line (str): The FIX message line.
    
    Returns:
    - dict: Parsed tag values, including TID and Order Type, and handling for repeated tags.
    """
    data = {'Unknown Tags': {}}
    tid, order_type = extract_preamble_details(line)
    data['TID'] = tid
    data['Order Type'] = order_type
    
    # Initialize containers for repeated tags
    partyIDs, partyIDSources, partyRoles = [], [], []
    
    _, message_str = line.split(']: ', 1)
    for part in message_str.split('|'):
        if '=' not in part:
            continue  # Skip malformed parts
        key, value = part.split('=', 1)
        if key in columns_mapping:
            if key in ['448', '447', '452']:  # Check if the tag is a repeated tag
                if key == '448':
                    partyIDs.append(value)
                elif key == '447':
                    partyIDSources.append(value)
                elif key == '452':
                    partyRoles.append(value)
            else:
                # Use the human-readable column name if the tag is known
                data[columns_mapping[key]] = value
        else:
            # Create a separate column for each unknown tag
            data[f'Unknown Tag {key}'] = value
    
    # Add the repeated tag lists to the data dictionary
    if partyIDs:
        data['448 PartyID'] = partyIDs
        data['447 PartyIDSource'] = partyIDSources
        data['452 PartyRole'] = partyRoles

    return data


def process_file(filepath):
    """
    Processes each line of a FIX log file.
    
    Parameters:
    - filepath (str): Path to the FIX log file.
    
    Returns:
    - list: A list of dictionaries with parsed message data.
    """
    with open(filepath, 'r') as file:
        messages = [parse_message(line) for line in file if line.strip()]
    return messages

def adjust_column_widths(worksheet):
    """
    Adjusts the column widths of a worksheet based on the length of its contents.
    
    Parameters:
    - worksheet: The Excel worksheet object.
    """
    for column_cells in worksheet.columns:
        length = max(len(str(cell.value)) for cell in column_cells)
        worksheet.column_dimensions[get_column_letter(column_cells[0].column)].width = length + 2

def save_to_excel(messages, filename):
    """
    Saves the parsed FIX messages to an Excel file, adjusting column widths and ensuring column order,
    including dynamically handling unknown tags.
    
    Parameters:
    - messages (list): List of dictionaries containing the parsed FIX message data.
    - filename (str): Filename for the output Excel file.
    """
    df = pd.DataFrame(messages)
    # Generate the ordered list of columns based on the `columns_mapping`, plus any additional columns for unknown tags
    ordered_columns = ['TID', 'Order Type'] + [columns_mapping[tag] for tag in columns_mapping if tag not in ['TID', 'Order Type']]
    # Determine if there are any extra columns in the DataFrame not in the ordered_columns list
    extra_columns = [col for col in df.columns if col not in ordered_columns]
    final_columns_order = ordered_columns + extra_columns
    df = df.reindex(columns=final_columns_order)  # Reorder the DataFrame columns

    with pd.ExcelWriter(filename, engine='openpyxl') as writer:
        df.to_excel(writer, index=False)
        adjust_column_widths(writer.sheets['Sheet1'])

if __name__ == "__main__":
    input_filename = input("Enter the full file path to parse: ")
    messages = process_file(input_filename)

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    base_filename = os.path.splitext(os.path.basename(input_filename))[0]
    output_filename = f'Parsed_{base_filename}_{timestamp}.xlsx'

    save_to_excel(messages, output_filename)
    print(f"Parsed file: '{output_filename}'.")
