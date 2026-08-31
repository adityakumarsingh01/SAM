"""
Excel Helper Copilot Plugin
Uses COM to interact with the active Excel spreadsheet to read, write, and format cells safely.
"""
import json
try:
    import win32com.client
    import pythoncom
except ImportError:
    win32com = None

PLUGIN = {
    "name": "excel_helper",
    "description": (
        "Use this tool to interact directly with the active Microsoft Excel application. "
        "You can read whole tables to see formulas and values, write data to specific cells, "
        "format cells (colors, bold), clear ranges, and get sheet info. "
        "ALWAYS use this tool instead of computer_control when interacting with Excel. "
        "CRITICAL: When fixing a table, ONLY modify the cells that are broken or causing errors (like text instead of numbers). "
        "DO NOT change the values of valid cells or guess new data. "
        "CRITICAL 2: If fixing multiple cells, DO NOT call write_range multiple times! "
        "Pass a 2D JSON array to 'write_range' (e.g. for A1:B2) to fix everything in a single, lightning-fast batch update."
    ),
    "parameters": {
        "type": "OBJECT",
        "properties": {
            "action": {
                "type": "STRING", 
                "description": "The action to perform: 'read_range', 'write_range', 'format_cell', 'clear_range', or 'get_active_sheet_info'."
            },
            "cell": {
                "type": "STRING", 
                "description": "The cell reference or range, e.g., 'B3', 'D4', 'A1:C5'. Not required for 'get_active_sheet_info'."
            },
            "value": {
                "type": "STRING", 
                "description": (
                    "For 'write_range': The value to write. To write multiple rows/columns, provide a JSON string of a 2D array (e.g. '[[\"A\",\"B\"], [1, 2]]'). "
                    "For 'format_cell': Provide a JSON object with optional keys: 'bold' (bool), 'italic' (bool), 'color' (hex string like '#FF0000')."
                )
            },
            "read_formulas": {
                "type": "BOOLEAN",
                "description": "If True and action is 'read_range', reads the formulas instead of the computed values."
            }
        },
        "required": ["action"],
    },
}

def _parse_value(value_str):
    if value_str is None:
        return ""
    
    if not isinstance(value_str, str):
        value_str = str(value_str)
        
    if not value_str:
        return value_str
    
    # Check if it's a JSON array
    if value_str.startswith('[') and value_str.endswith(']'):
        try:
            return json.loads(value_str)
        except json.JSONDecodeError:
            pass

    # Basic numeric conversion
    if value_str.isdigit() or (value_str.startswith('-') and value_str[1:].isdigit()):
        return int(value_str)
    try:
        if '.' in value_str:
            return float(value_str)
    except ValueError:
        pass
    
    return value_str

def run(parameters: dict, player=None, session_memory=None) -> str:
    if not win32com:
        return "Error: pywin32 is not installed. Run 'pip install pywin32'."

    action = parameters.get("action", "").lower()
    cell = parameters.get("cell", "").upper()
    value_str = parameters.get("value", "")
    read_formulas = parameters.get("read_formulas", False)
    
    if action != "get_active_sheet_info" and not cell:
        return "Error: 'cell' is required for this action."

    try:
        pythoncom.CoInitialize()
        excel = win32com.client.GetActiveObject("Excel.Application")
        wb = excel.ActiveWorkbook
        if wb is None:
            return "Error: No active Excel workbook found."
            
        sheet = wb.ActiveSheet
        
        if action == "get_active_sheet_info":
            used_range = sheet.UsedRange
            address = used_range.Address
            return f"Active Sheet: '{sheet.Name}'. Used range: {address.replace('$', '')}"

        target_range = sheet.Range(cell)
        
        if action == "write_range":
            if player:
                player.write_log(f"SAM: Writing to Excel {cell}...")
            
            parsed_val = _parse_value(value_str)
            
            if isinstance(parsed_val, list):
                # 2D array write
                num_rows = len(parsed_val)
                num_cols = len(parsed_val[0]) if num_rows > 0 and isinstance(parsed_val[0], list) else 1
                
                # Resize the target range to fit the array precisely
                target_range = target_range.Resize(num_rows, num_cols)
                
                # Convert 1D list to 2D list for single column arrays
                if num_cols == 1 and not isinstance(parsed_val[0], list):
                    parsed_val = [[x] for x in parsed_val]
                    
                target_range.Value = parsed_val
                return f"Successfully wrote 2D array to {target_range.Address.replace('$', '')}"
            else:
                # Single value write
                if isinstance(parsed_val, str) and parsed_val.startswith("="):
                    target_range.Formula = parsed_val
                else:
                    target_range.Value = parsed_val
                return f"Successfully wrote '{parsed_val}' to cell {cell}."
                
        elif action == "read_range":
            if player:
                player.write_log(f"SAM: Reading Excel {cell}...")
                
            if read_formulas:
                val = target_range.Formula
            else:
                val = target_range.Value
                
            if val is None:
                val = ""
                
            # Formatting nicely for the LLM
            if isinstance(val, tuple):
                # Format 2D array as JSON string for compactness
                result = json.dumps([list(row) for row in val])
                return f"Data in {cell}:\n{result}"
            else:
                return f"Value in {cell}: {val}"
                
        elif action == "clear_range":
            if player:
                player.write_log(f"SAM: Clearing Excel {cell}...")
            target_range.ClearContents()
            return f"Successfully cleared {cell}."
            
        elif action == "format_cell":
            if player:
                player.write_log(f"SAM: Formatting Excel {cell}...")
                
            try:
                fmt = json.loads(value_str)
            except:
                return "Error: 'value' must be a valid JSON object for format_cell (e.g., {\"bold\": true})."
                
            if "bold" in fmt:
                target_range.Font.Bold = fmt["bold"]
            if "italic" in fmt:
                target_range.Font.Italic = fmt["italic"]
            if "color" in fmt:
                # Convert hex to BGR format for Excel
                hex_color = fmt["color"].lstrip('#')
                if len(hex_color) == 6:
                    r, g, b = tuple(int(hex_color[i:i+2], 16) for i in (0, 2, 4))
                    target_range.Interior.Color = b * 65536 + g * 256 + r
                    
            return f"Successfully formatted {cell}."
            
        else:
            return f"Error: Unknown action '{action}'."
            
    except Exception as e:
        if "Operation unavailable" in str(e):
            return "Error: Excel is currently in 'Edit Mode' (a cell is double-clicked). Tell the user to press Enter or Escape to exit Edit Mode first."
        return f"Failed to interact with Excel: {str(e)}"
    finally:
        try:
            pythoncom.CoUninitialize()
        except:
            pass
