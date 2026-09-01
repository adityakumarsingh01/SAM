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
        "You can read whole tables, write data, format cells, clear ranges, and get sheet info. "
        "CRITICAL: When fixing a table, ONLY modify the exact cells that are broken. Do NOT change valid cells. "
        "CRITICAL 2: If fixing multiple scattered cells, use the 'batch_update' action! "
        "For 'batch_update', pass a JSON array of objects in 'value' (e.g. '[{\"cell\": \"C2\", \"value\": 25}, {\"cell\": \"B5\", \"value\": 2}]'). "
        "This is the ONLY correct way to fix multiple non-contiguous errors in one step. "
        "For 'analyze_data', pass a JSON object in 'value' (e.g. '{\"group_by\": \"Product\", \"sum\": \"Revenue\"}') to aggregate totals. "
        "For 'highlight_rows', pass a JSON object in 'value' (e.g. '{\"column\": \"Product\", \"match\": \"Smartphone\", \"format\": {\"bold\": true, \"color\": \"#FFFF00\"}}')."
    ),
    "parameters": {
        "type": "OBJECT",
        "properties": {
            "action": {
                "type": "STRING", 
                "description": "The action to perform: 'read_range', 'write_range', 'batch_update', 'format_cell', 'clear_range', 'analyze_data', 'highlight_rows', or 'get_active_sheet_info'."
            },
            "cell": {
                "type": "STRING", 
                "description": "The cell reference or range, e.g., 'B3'. Not required for 'get_active_sheet_info' or 'batch_update'."
            },
            "value": {
                "type": "STRING", 
                "description": (
                    "For 'write_range': The value to write (or 2D JSON array for contiguous block). "
                    "For 'batch_update': A JSON array of cell/value objects. "
                    "For 'format_cell': A JSON object like {\"bold\": true, \"color\": \"#FF0000\"}."
                )
            },
            "read_formulas": {
                "type": "BOOLEAN",
                "description": "If True and action is 'read_range', reads formulas instead of values."
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
    
    if action not in ("get_active_sheet_info", "batch_update", "analyze_data", "highlight_rows") and not cell:
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

        target_range = None
        if cell:
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
                
        elif action == "batch_update":
            if player:
                player.write_log("SAM: Performing batch update in Excel...")
            updates = _parse_value(value_str)
            if not isinstance(updates, list):
                return "Error: 'value' must be a JSON array of objects for 'batch_update'."
            for update in updates:
                c = update.get("cell")
                v = update.get("value")
                r = sheet.Range(c)
                if isinstance(v, str) and v.startswith("="):
                    r.Formula = v
                else:
                    r.Value = v
            return f"Successfully processed {len(updates)} scattered cell updates."
                
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
                
        elif action == "analyze_data":
            if player:
                player.write_log(f"SAM: Analyzing Excel data...")
            try:
                params = json.loads(value_str)
                group_col_name = params.get("group_by")
                sum_col_name = params.get("sum")
                
                used_range = sheet.UsedRange.Value
                if not used_range or len(used_range) < 2:
                    return "Error: Not enough data."
                
                headers = [str(h).strip() for h in used_range[0]]
                if group_col_name not in headers or sum_col_name not in headers:
                    return f"Error: Columns '{group_col_name}' or '{sum_col_name}' not found. Available: {headers}"
                
                g_idx = headers.index(group_col_name)
                s_idx = headers.index(sum_col_name)
                
                results = {}
                for row in used_range[1:]:
                    if len(row) > max(g_idx, s_idx):
                        key = str(row[g_idx]).strip()
                        val = row[s_idx]
                        if not key or key == 'None': continue
                        try:
                            val = float(val)
                        except (ValueError, TypeError):
                            continue
                        results[key] = results.get(key, 0.0) + val
                
                sorted_results = sorted(results.items(), key=lambda x: x[1], reverse=True)
                
                md = f"| {group_col_name} | Total {sum_col_name} |\n| --- | ---: |\n"
                for k, v in sorted_results:
                    md += f"| {k} | {v:,.0f} |\n"
                
                if sorted_results:
                    best = sorted_results[0]
                    md += f"\n🏆 Highest: {best[0]} ({best[1]:,.0f})"
                
                return f"Analysis Complete. Show this to the user:\n{md}"
            except Exception as e:
                return f"Error in analyze_data: {str(e)}"
                
        elif action == "highlight_rows":
            if player:
                player.write_log("SAM: Highlighting rows...")
            try:
                params = json.loads(value_str)
                column_name = params.get("column")
                match_text = str(params.get("match")).lower()
                fmt = params.get("format", {})
                
                used_range = sheet.UsedRange.Value
                if not used_range or len(used_range) < 2:
                    return "Error: Not enough data."
                
                headers = [str(h).strip() for h in used_range[0]]
                if column_name not in headers:
                    return f"Error: Column '{column_name}' not found. Available: {headers}"
                
                col_idx = headers.index(column_name)
                total_cols = len(headers)
                
                rows_to_highlight = []
                for i, row in enumerate(used_range[1:]):
                    if len(row) > col_idx:
                        val = str(row[col_idx]).lower()
                        if match_text in val:
                            rows_to_highlight.append(i + 2) # +1 for 0-index, +1 for header
                
                if not rows_to_highlight:
                    return f"No rows found matching '{match_text}'."
                
                for row_num in rows_to_highlight:
                    r = sheet.Range(sheet.Cells(row_num, 1), sheet.Cells(row_num, total_cols))
                    if "bold" in fmt:
                        r.Font.Bold = fmt["bold"]
                    if "color" in fmt:
                        hex_color = fmt["color"].lstrip('#')
                        if len(hex_color) == 6:
                            red, g, b = tuple(int(hex_color[i:i+2], 16) for i in (0, 2, 4))
                            r.Interior.Color = b * 65536 + g * 256 + red
                            
                return f"Successfully highlighted {len(rows_to_highlight)} rows matching '{match_text}'."
            except Exception as e:
                return f"Error in highlight_rows: {str(e)}"

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
