import xml.etree.ElementTree as ET
import os
from datetime import datetime, timedelta
import copy

def get_date(date_str):
    return datetime.strptime(date_str.split(';')[0], "%Y%m%d")

def merge_ibkr_xmls(folder_path, output_file):
    error_file = 'error.xml'
    xml_files = [f for f in os.listdir(folder_path) if f.endswith('.xml')]
    
    if not xml_files:
        print("❌ ERROR: No XML files found in the ./xml/ folder.")
        return

    print(f"🔍 Found {len(xml_files)} files. Analyzing dates...")

    # 1. Parse and Inspect
    statements = []
    for f in xml_files:
        try:
            tree = ET.parse(os.path.join(folder_path, f))
            root = tree.getroot()
            # IBKR usually nests: FlexQueryResponse -> FlexStatements -> FlexStatement
            stmt = root.find(".//FlexStatement")
            
            if stmt is not None:
                start_date = stmt.get('fromDate')
                end_date = stmt.get('toDate')
                statements.append({
                    'file': f, 'root': root, 'stmt': stmt,
                    'from': get_date(start_date), 'to': get_date(end_date)
                })
                print(f" ✅ Loaded {f}: {start_date} to {end_date}")
            else:
                print(f" ⚠️ WARNING: {f} has no <FlexStatement> tag. Skipping.")
        except Exception as e:
            print(f" ❌ ERROR: Could not read {f}: {e}")
            return

    # Sort: Earliest (2024) first
    statements.sort(key=lambda x: x['from'])

    # 2. Check for Gaps
    print("\n📅 Checking for timeline gaps...")
    for i in range(len(statements) - 1):
        curr = statements[i]
        nxt = statements[i+1]
        if nxt['from'] > curr['to'] + timedelta(days=1):
            print(f" 🛑 STOPPED: Gap detected between {curr['file']} and {nxt['file']}!")
            print(f"    {curr['file']} ends {curr['to'].date()}")
            print(f"    {nxt['file']} starts {nxt['from'].date()}")
            if os.path.exists(output_file): os.rename(output_file, error_file)
            else:
                with open(error_file, 'w') as f: f.write("Gap detected")
            return
    print(" ✅ Timeline is continuous. No gaps found.")

    # 3. Merging
    print("\n🏗️  Merging data into template...")
    final_root = statements[0]['root']
    final_stmt = final_root.find(".//FlexStatement")
    
    # Collect trades/cash from all statements BEFORE modifying template (copy elements)
    per_file_trades = {}
    per_file_cash = {}
    for s in statements:
        per_file_trades[s['file']] = [copy.deepcopy(t) for t in s['stmt'].findall(".//Trade")]
        per_file_cash[s['file']] = [copy.deepcopy(c) for c in s['stmt'].findall(".//CashTransaction")]
    
    # Get or Create Containers
    trades_container = final_stmt.find(".//Trades")
    cash_container = final_stmt.find(".//CashTransactions")
    
    if trades_container is None: 
        trades_container = ET.SubElement(final_stmt, "Trades")
    if cash_container is None: 
        cash_container = ET.SubElement(final_stmt, "CashTransactions")

    # clear existing children to start fresh
    trades_container.clear()
    cash_container.clear()
    
    seen_trades = set()
    seen_cash = set()

    for s in statements:
        t_count = 0
        c_count = 0
        
        # Pull trades (from copied lists)
        for t in per_file_trades[s['file']]:
            tid = t.get("tradeID")
            if tid and tid not in seen_trades:
                trades_container.append(t)
                seen_trades.add(tid)
                t_count += 1
        
        # Pull cash (from copied lists)
        for c in per_file_cash[s['file']]:
            fp = (c.get("dateTime"), c.get("symbol"), c.get("type"), c.get("amount"))
            if fp not in seen_cash:
                cash_container.append(c)
                seen_cash.add(fp)
                c_count += 1
        
        print(f" ➕ Merged {s['file']}: {t_count} trades, {c_count} cash items.")

    # 4. Finalize
    final_stmt.set("fromDate", statements[0]['stmt'].get("fromDate"))
    final_stmt.set("toDate", statements[-1]['stmt'].get("toDate"))
    
    tree = ET.ElementTree(final_root)
    tree.write(output_file, encoding="utf-8", xml_declaration=True)
    
    if os.path.exists(error_file): os.remove(error_file)

    print(f"\n✨ SUCCESS! Created {output_file}")
    print(f"📊 FINAL TOTALS: {len(seen_trades)} Trades | {len(seen_cash)} Cash Items")

merge_ibkr_xmls('./xml/', 'merged.xml')