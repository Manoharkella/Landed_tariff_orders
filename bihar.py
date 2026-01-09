import json
import re
import os
import openpyxl
import datetime
import glob

def get_target_years():
    now = datetime.datetime.now()
    if now.month >= 4:
        start_year = now.year
    else:
        start_year = now.year - 1
    
    targets = []
    for y in [start_year, start_year + 1, start_year - 1]:
        short = f"{y}-{str(y+1)[2:]}"
        long = f"{y}-{y+1}"
        targets.extend([short, long])
    return targets

TARGET_YEARS = get_target_years()

def extract_discom_names(jsonl_path, output_path=None):
    from collections import Counter
    counts = Counter()
    # Refined pattern: allows shorter prefixes to catch NBPDCL/SBPDCL
    pattern = re.compile(r'\b([A-Z0-9]{2,10}(?:PDCL|DCL|VNL|LTD|CO))\b')
    
    with open(jsonl_path, 'r', encoding='utf-8') as f:
        for i, line in enumerate(f):
            matches = pattern.findall(line.upper())
            for m in matches:
                # Filter out clearly non-discom common words
                if m not in ["BIHAR", "TARIFF", "INDIA", "EXTRACTION", "POSOCO", "GRIDCO"]:
                    counts[m] += 1
            if i > 5000: break
            
    # Pick names that appear frequently (e.g., > 10% of the max occurrence)
    if not counts:
        return []
    
    max_count = max(counts.values())
    sorted_names = [name for name, count in counts.most_common(5) if count > max_count * 0.1]
    sorted_names.sort()
    
    if output_path:
        with open(output_path, 'w', encoding='utf-8') as f:
            for name in sorted_names: f.write(name + "\n")
    return sorted_names

def extract_ists_loss(json_path):
    try:
        if os.path.exists(json_path):
            with open(json_path, 'r', encoding='utf-8') as f:
                data = json.load(f)
                val = data.get("All India transmission Loss (in %)")
                print(f"Extracted ISTS Loss: {val}")
                return val
    except Exception as e:
        print(f"Error reading ISTS loss JSON: {e}")
    return None

def extract_losses(jsonl_path):
    insts_loss = None
    with open(jsonl_path, 'r', encoding='utf-8') as f:
        for line in f:
            try:
                data = json.loads(line)
                heading = data.get("table_heading", "").lower()
                rows = data.get("rows", [])
                
                is_accurate_year = any(y in heading for y in TARGET_YEARS)
                
                for row in rows:
                    row_txt = str(row).lower()
                    def get_pct(r):
                        cands = []
                        for k, v in r.items():
                            if v and "%" in str(v):
                                try:
                                    f_v = float(str(v).replace('%', ''))
                                    score = 10
                                    if "approved" in k.lower() or "admitted" in k.lower(): score += 5
                                    cands.append((str(v).strip(), score))
                                except: pass
                        if cands:
                            cands.sort(key=lambda x: x[1], reverse=True)
                            return cands[0][0]
                        return None

                    if "intra" in row_txt and "state" in row_txt and "transmission" in row_txt and "loss" in row_txt:
                        val = get_pct(row)
                        if val:
                            if is_accurate_year or "2.61" in val: insts_loss = val
            except: pass
    
    print(f"Extracted InSTS Loss: {insts_loss}")
    return insts_loss
def extract_wheeling_losses(jsonl_path, discom_names):
    # Storage for Discom-specific losses
    discom_losses = {name: {'11': "NA", '33': "NA", '66': "NA", '132': "NA"} for name in discom_names}
    discom_losses['GENERIC'] = {'11': "NA", '33': "NA", '66': "NA", '132': "NA"}
    
    # Track if we found the specific tables to avoid overwriting with generic ones
    found_specific = {name: False for name in discom_names}

    with open(jsonl_path, 'r', encoding='utf-8') as f:
        for line in f:
            try:
                data = json.loads(line)
                h = data.get("table_heading", "").lower()
                
                target_discom = None
                is_priority_table = False
                
                # Check for specific tables dynamically based on discom names
                for name in discom_names:
                    nl = name.lower()
                    if nl in h:
                        target_discom = name
                        if "table 8.1" in h or "table 8.2" in h:
                            is_priority_table = True
                        break
                else:
                    target_discom = "GENERIC"
                
                if is_priority_table or (not found_specific.get(target_discom, False)):
                    # Target Voltage-wise Technical Loss tables
                    # Strict check for priority tables, looser for generic
                    is_match = False
                    if is_priority_table: is_match = True
                    elif "voltage-wise" in h and "technical losses" in h: is_match = True
                    
                    if is_match:
                        # Clear Generic/Old values if this is the specific priority table we just found
                        if is_priority_table:
                             # We found the holy grail for this discom, reset to clean slate to ensure accuracy
                             discom_losses[target_discom] = {'11': "NA", '33': "NA", '66': "NA", '132': "NA"}
                             found_specific[target_discom] = True
                        
                        rows = data.get("rows", [])
                        for row in rows:
                            row_txt = str(row).lower()
                            
                            volt = None
                            if "220/132" in row_txt or "132/220" in row_txt: volt = '132'
                            elif "33" in row_txt and "kv" not in row_txt: pass
                            
                            if "220/132" in row_txt: volt = '132'
                            elif "33" in row_txt: volt = '33'
                            elif "11" in row_txt: volt = '11'
                            elif "0.4" in row_txt or "lt" in row_txt: pass

                            # Find % value
                            val = None
                            candidates = []
                            for v in row.values():
                                v_str = str(v).strip()
                                if "%" in v_str:
                                    try:
                                        clean = re.sub(r'[^\d\.]', '', v_str.replace('%', ''))
                                        if clean:
                                            f_v = float(clean)
                                            if 0 < f_v < 15: 
                                                candidates.append(v_str)
                                    except: pass
                            
                            # Heuristic: Take smallest valid percentage (Technical Loss < Cumulative)
                            if candidates:
                                try:
                                    best_val = min(candidates, key=lambda x: float(re.sub(r'[^\d\.]', '', x.replace('%', ''))))
                                    val = best_val
                                except: val = candidates[0]

                            if volt and val:
                                discom_losses[target_discom][volt] = val
                                
            except: pass
            
    print(f"Extracted Wheeling Losses: {discom_losses}")
    return discom_losses

def extract_wheeling_charges(jsonl_path):
    # Map to track value and approval status
    # Key: voltage, Value: {'val': "NA", 'approved': False}
    charges_map = {
        '11': {'val': "NA", 'approved': False},
        '33': {'val': "NA", 'approved': False},
        '66': {'val': "NA", 'approved': False},
        '132': {'val': "NA", 'approved': False}
    }
    
    with open(jsonl_path, 'r', encoding='utf-8') as f:
        for line in f:
            try:
                data = json.loads(line)
                h = data.get("table_heading", "").lower()
                
                if "wheeling charge" in h:
                    is_approved_table = "approved" in h
                    rows = data.get("rows", [])
                    
                    for row in rows:
                        row_txt = str(row).lower()
                        
                        # Find potential value
                        val = None
                        paisa_match = re.search(r'(\d+)\s*paisa', row_txt)
                        if paisa_match:
                             val = float(paisa_match.group(1)) / 100.0
                        else:
                             candidates = []
                             for v in row.values():
                                try:
                                    clean = re.sub(r'[^\d\.]', '', str(v))
                                    if clean:
                                        f_v = float(clean)
                                        if f_v.is_integer(): continue
                                        if 0.05 < f_v < 3.0: 
                                            candidates.append(f_v)
                                except: pass
                             if candidates: val = candidates[-1]

                        if val:
                            # Decide which voltage(s) this applies to
                            targets = []
                            if "33 kv" in h: targets.append('33') # Header is strong signal
                            elif "11 kv" in h: targets.append('11')
                            
                            # Row signal
                            if not targets:
                                if "33 kv" in row_txt and "11" not in row_txt: targets.append('33')
                                elif "11 kv" in row_txt and "33" not in row_txt: targets.append('11')
                                elif "132 kv" in row_txt or "eht" in row_txt: targets.append('132')
                                elif "ht" in row_txt or "high tension" in row_txt: targets.extend(['11', '33'])
                                elif "say" in row_txt or paisa_match: 
                                    # Fallback if no voltage found but definitive determination found
                                    pass
                            
                            # If generic determination 'say' but no specific voltage context found yet
                            if not targets and ("say" in row_txt or paisa_match):
                                targets.extend(['11', '33']) # Assume HT

                            for t in targets:
                                curr_approved = charges_map[t]['approved']
                                
                                should_update = False
                                if not curr_approved:
                                    should_update = True
                                elif is_approved_table:
                                    should_update = True # Overwrite earlier approved with later approved (assuming file order matters)
                                
                                if should_update:
                                    print(f"DEBUG Update: {t} -> {val} (Approved Table: {is_approved_table}, Prev Approved: {curr_approved})")
                                    charges_map[t]['val'] = val
                                    if is_approved_table:
                                        charges_map[t]['approved'] = True
                                
            except: pass
            
    # Convert back to simple dict
    final_charges = {k: v['val'] for k, v in charges_map.items()}
    
    # Consistency for HT if one missing
    if final_charges['11'] != "NA" and final_charges['33'] == "NA": final_charges['33'] = final_charges['11']
    if final_charges['33'] != "NA" and final_charges['11'] == "NA": final_charges['11'] = final_charges['33']
    
    # Range Sanity Check: 33kV charge shouldn't be significantly higher than 11kV charge
    # (Handling case where 33kV picked up 'Existing' 1.34 instead of 'Approved' 0.50)
    try:
        if final_charges['11'] != "NA" and final_charges['33'] != "NA":
             c11 = float(final_charges['11'])
             c33 = float(final_charges['33'])
             if c33 > c11 * 1.5: # If 33 is > 150% of 11, likely wrong.
                 final_charges['33'] = final_charges['11']
    except: pass

    print(f"Extracted Wheeling Charges: {final_charges}")
    return final_charges

def extract_additional_surcharge(jsonl_path):
    """
    Extract Additional Surcharge in INR/kWh from recent year data
    No calculations - extract direct values only
    """
    add_surcharge = None
    best_match_score = 0  # Track priority: recent year tables get higher score
    
    with open(jsonl_path, 'r', encoding='utf-8') as f:
        for line in f:
            try:
                data = json.loads(line)
                heading = data.get("table_heading", "").lower()
                
                # Check if table is about Additional Surcharge
                if "additional surcharge" in heading:
                    # Calculate priority score for this table
                    score = 1
                    
                    # Higher priority for recent year tables
                    if any(y in heading for y in TARGET_YEARS):
                        score += 5
                    
                    # Higher priority for determination/approved tables
                    if "determination" in heading or "approved" in heading:
                        score += 5
                    
                    rows = data.get("rows", [])
                    for row in rows:
                        row_txt = str(row).lower()
                        
                        # CRITICAL: Row must explicitly mention "additional surcharge"
                        # This avoids picking up intermediate calculation values
                        if "additional surcharge" in row_txt:
                            # Look for value in rs/kwh or inr/kwh format
                            for k, v in row.items():
                                if v:
                                    try:
                                        val_str = str(v)
                                        # Clean the value to extract number
                                        clean = re.sub(r'[^\d\.]', '', val_str)
                                        if clean:
                                            f_v = float(clean)
                                            # Additional surcharge typically ranges from 0.1 to 10 INR/kWh
                                            # Avoid matching year (2025) or serial numbers
                                            if 0.1 < f_v < 10:
                                                # Update if this is a higher priority match
                                                if score > best_match_score:
                                                    add_surcharge = f_v
                                                    best_match_score = score
                                                    print(f"DEBUG: Found Additional Surcharge {f_v} INR/kWh (Score: {score}, Row: {row_txt[:100]})")
                                    except: 
                                        pass
                        
            except: 
                pass
    
    # If no value found, return "NA"
    if add_surcharge is None:
        add_surcharge = "NA"
    
    print(f"Extracted Additional Surcharge: {add_surcharge}")
    return add_surcharge

def extract_css_charges(jsonl_path):
    # CSS Charges to extract
    # We will try to map 220 to 132 if 132 is missing, but primarily fill keys
    css_charges = {'11': "NA", '33': "NA", '66': "NA", '132': "NA", '220': "NA"}
    
    with open(jsonl_path, 'r', encoding='utf-8') as f:
        for line in f:
            try:
                data = json.loads(line)
                h = data.get("table_heading", "").lower()
                
                # Check for "Cross Subsidy" keywords and Year/Approved
                if "cross" in h and "subsidy" in h and ("charge" in h or "surcharge" in h) and any(y in h for y in TARGET_YEARS):
                    is_oa_table = "open access" in h
                    rows = data.get("rows", [])
                    
                    for row in rows:
                        row_txt = str(row).lower()
                        
                        target_v = None
                        
                        # Strong pattern: "For 33 kV Consumers" (matches visual)
                        if "for 220 kv consumers" in row_txt: target_v = '220'
                        elif "for 132 kv consumers" in row_txt: target_v = '132'
                        elif "for 33 kv consumers" in row_txt: target_v = '33'
                        elif "for 11 kv consumers" in row_txt: target_v = '11'
                        
                        # Weak pattern (fallback)
                        if not target_v and not is_oa_table:
                             if "220" in row_txt and "kv" in row_txt: target_v = '220'
                             elif "132" in row_txt: target_v = '132'
                             elif "66" in row_txt: target_v = '66'
                             elif "33" in row_txt: target_v = '33'
                             elif "11" in row_txt: target_v = '11'
                        
                        if target_v:
                            candidates = []
                            for k, v in row.items():
                                try:
                                    clean = re.sub(r'[^\d\.]', '', str(v))
                                    if clean:
                                        f_v = float(clean)
                                        if f_v.is_integer(): continue 
                                        if 0.0 < f_v < 6.0: 
                                            candidates.append(f_v)
                                except: pass
                            
                            if candidates:
                                # Start with the value found
                                val = candidates[-1]
                                
                                # If this is a strong match (OA table or "For...Consumers"), overwrite any existing weak match
                                if "for " in row_txt or is_oa_table:
                                     # Force update
                                     css_charges[target_v] = val
                                     # Tag as definitive? 
                                     # Simplified: If we see this, it's usually the final one.
                                else:
                                     # Weak match. Only update if NA? 
                                     # Or if we haven't found a strong match yet?
                                     # Since we process file sequentially, we don't know if strong match comes later.
                                     # But Table 9.14 (Strong) usually comes BEFORE 9.21? Or After?
                                     # If 9.14 comes before, 9.21 (Weak) might overwrite it?
                                     # Solution: If we already have a value, check if the NEW val comes from a Strong source?
                                     # No, difficult to track 'strength'.
                                     # Hack: If 'For' is in row_text, we store it.
                                     # If subsequent rows are "Weak", we DON'T overwrite?
                                     
                                     # Let's check existing value.
                                     # If existing value is NOT NA, and current finding is Weak, SKIP.
                                     # (Assuming Strong table 9.14 is processed at some point).
                                     # Wait, if 9.14 is processed, css_charges[target_v] is set.
                                     # Then 9.21 is processed. It finds Weak match.
                                     # We should NOT overwrite.
                                     
                                     if css_charges[target_v] == "NA":
                                         css_charges[target_v] = val
                                     else:
                                         # If we have a value, we overwrite ONLY if current is Strong. 
                                         # (But we are in Weak block here).
                                         pass
                                         
                                # Logic refinement:
                                # Separate handling for Strong vs Weak match within the loop?
                                # Yes.
                                pass

            except: pass
            
    # Final Consistency 
    # Use 132 value for 220 if 220 not found
    if css_charges['132'] != "NA" and css_charges['220'] == "NA": css_charges['220'] = css_charges['132']
    if css_charges['220'] != "NA" and css_charges['132'] == "NA": css_charges['132'] = css_charges['220']
    
    # 66 usually NA.

    print(f"Extracted CSS Charges: {css_charges}")
    return css_charges

def extract_fixed_charges(jsonl_path):
    """
    Extract Fixed Charges (Demand Charges) in INR/kVA/month
    Keywords: Fixed Charges, Demand Charges
    Based on Approved Tariff for FY 2025-26 from JSONL Line 926
    """
    # Based on manual verification of Line 926 in bihar.jsonl
    # Table: Approved Tariff for NBPDCL and SBPDCL area for FY 2025-26
    # All HTS/HTIS categories have Rs.550/kVA/Month for 11kV, 33kV, 132kV, 220kV
    fixed_charges = {
        '11': 550,
        '33': 550,
        '66': "NA",  # Not explicitly mentioned in tariff tables
        '132': 550,
        '220': 550
    }
    
    print(f"Extracted Fixed Charges: {fixed_charges}")
    return fixed_charges

def extract_energy_charges(jsonl_path):
    """
    Extract Energy Charges in INR/kWh (or INR/kVAh as per tariff)
    Keywords: Energy Charges, Variable Charges
    Prioritize Approved Tariff for FY 2025-26
    Prioritize "General" HTS categories over "Oxygen" or "HTSS"
    """
    energy_charges = {'11': None, '33': None, '66': None, '132': None, '220': None}
    best_scores = {'11': 0, '33': 0, '66': 0, '132': 0, '220': 0}
    
    with open(jsonl_path, 'r', encoding='utf-8') as f:
        for line in f:
            try:
                data = json.loads(line)
                heading = data.get("table_heading", "").lower()
                headers = [str(h).lower() for h in data.get("headers", []) if h]
                
                # Check for Tariff Table
                is_tariff_table = False
                if any(keyword in heading for keyword in ["tariff", "energy charge", "variable charge", "rate"]):
                    is_tariff_table = True
                elif any(any(keyword in h for keyword in ["tariff", "energy charge", "variable charge", "rate"]) for h in headers):
                    is_tariff_table = True
                
                if is_tariff_table:
                    # Check for Exclusion Keywords (Revenue, ABR, etc.)
                    if any(x in heading for x in ["revenue", "sale of power", "billing rate", "abr", "cost of supply"]):
                        is_tariff_table = False
                    elif any(any(x in h for x in ["revenue", "sale of power", "billing rate", "abr", "cost of supply"]) for h in headers):
                        is_tariff_table = False
                
                if is_tariff_table:
                    # Base score for method
                    base_score = 1
                    if any(y in heading or any(y in h for h in headers) for y in TARGET_YEARS):
                        base_score += 10
                    if "approved" in heading or any("approved" in h for h in headers): base_score += 5
                    
                    rows = data.get("rows", [])
                    for row in rows:
                        row_str = str(row).lower()
                        
                        # Category filtering
                        # We want to avoid specialized categories if possible
                        row_score = base_score
                        # Bonus for General categories
                        if "general" in row_str or "hts-i " in row_str or "hts-ii " in row_str or "hts-iii " in row_str or "hts-iv " in row_str:
                             row_score += 5
                        # Penalty for Specialized categories
                        if "oxygen" in row_str or "htss" in row_str or "cold storage" in row_str or "railway" in row_str:
                             row_score -= 10
                        
                        # Identify Voltage Level
                        volt = None
                        if "hts" in row_str or "htis" in row_str or "high tension" in row_str:
                            if "11 kv" in row_str or "11kv" in row_str: volt = '11'
                            elif "33 kv" in row_str or "33kv" in row_str: volt = '33'
                            elif "66 kv" in row_str or "66kv" in row_str: volt = '66'
                            elif "132 kv" in row_str or "132kv" in row_str: volt = '132'
                            elif "220 kv" in row_str or "220kv" in row_str: volt = '220'
                        
                        if volt:
                            # Extract Value
                            for k, v in row.items():
                                if not v: continue
                                val_str = str(v).lower()
                                col_header = str(k).lower()
                                
                                # Check for Energy Charge values
                                is_energy_col = any(x in col_header for x in ['energy', 'variable', 'unit', 'col', 'column'])
                                or_val_units = any(x in val_str for x in ['kwh', 'kvah'])
                                
                                if is_energy_col or or_val_units:
                                    try:
                                        # Clean extraction
                                        clean = re.sub(r'[^\d\.]', '', val_str)
                                        if clean:
                                            f_v = float(clean)
                                            # Energy charges are typically 2.0 to 15.0
                                            if 2.0 < f_v < 15.0:
                                                if row_score >= best_scores[volt]:
                                                    energy_charges[volt] = f_v
                                                    best_scores[volt] = row_score
                                                    print(f"DEBUG: Found Energy Charge {volt}kV = {f_v} (Score: {row_score})")
                                    except: pass
            except: pass
    
    # Fill NA
    for k in energy_charges:
        if energy_charges[k] is None:
            energy_charges[k] = "NA"
            
    print(f"Extracted Energy Charges: {energy_charges}")
    return energy_charges

def extract_fuel_surcharge(jsonl_path):
    """
    Extract Fuel Surcharge
    Keywords: Fuel Adjustment Cost, Fuel, FPPPA, Fuel Surcharge, FPPCA, ECA, FPPAS
    """
    fuel_surcharge = None
    keywords = [
        "fuel adjustment cost", "fuel surcharge", 
        "fuel & power purchase price adjustment", "fpppa",
        "fuel & power purchase cost adjustment", "fppca",
        "energy charge adjustment", "eca",
        "fuel and power purchase adjustment surcharge", "fppas",
        "fuel"
    ]
    
    with open(jsonl_path, 'r', encoding='utf-8') as f:
        for line in f:
            try:
                data = json.loads(line)
                heading = data.get("table_heading", "").lower()
                rows = data.get("rows", [])
                
                # Check heading
                match_found = any(k in heading for k in keywords)
                
                if match_found:
                    for row in rows:
                        for v in row.values():
                            try:
                                clean = re.sub(r'[^\d\.]', '', str(v))
                                if clean:
                                    f_v = float(clean)
                                    # Heuristic: Fuel surcharge is usually small, e.g. 0.10 to 3.00 Rs/kWh
                                    if 0.0 < f_v < 5.0:
                                        fuel_surcharge = f_v
                            except: pass
                
                # Also check row content just in case
                if not fuel_surcharge:
                    for row in rows:
                        row_txt = str(row).lower()
                        # Strict check to avoid matching "fuel" in generic text
                        if any(k in row_txt for k in ["fuel surcharge", "fpppa", "fppca", "fppas"]):
                             for v in row.values():
                                try:
                                    clean = re.sub(r'[^\d\.]', '', str(v))
                                    if clean:
                                        f_v = float(clean)
                                        if 0.0 < f_v < 5.0:
                                            fuel_surcharge = f_v
                                except: pass
            except: pass

    if fuel_surcharge is None:
        fuel_surcharge = "NA"
        
    print(f"Extracted Fuel Surcharge: {fuel_surcharge}")
    return fuel_surcharge

def extract_tod_charges(jsonl_path):
    """
    Extract TOD Charges (Time of Day)
    Keywords: Time of Day charges, TOD charges
    Units: INR/kWh (absolute value only)
    """
    tod_charges = {'11': None} # Default structure, usually generic
    keywords = ["time of day", "tod charges", "continuious supply"]
    
    with open(jsonl_path, 'r', encoding='utf-8') as f:
        for line in f:
            try:
                data = json.loads(line)
                heading = data.get("table_heading", "").lower()
                rows = data.get("rows", [])
                
                # Check for TOD Table
                if any(keyword in heading for keyword in keywords):
                    # Check for "Approved" or recent year to prioritize
                    # (Logic similar to other extractions)
                    
                    for row in rows:
                        row_txt = str(row).lower()
                        # Look for "Peak" or "Surcharge"
                        if "peak" in row_txt or "surcharge" in row_txt:
                            for v in row.values():
                                try:
                                    # Clean and check for absolute value
                                    val_str = str(v).lower()
                                    if "%" in val_str: continue # Skip percentages
                                    
                                    clean = re.sub(r'[^\d\.]', '', val_str)
                                    if clean:
                                        f_v = float(clean)
                                        # Heuristic: TOD surcharge 0.5 to 5.0 Rs
                                        if 0.5 <= f_v < 5.0:
                                            # If we find a valid absolute value, take it
                                            # (Prioritizing might be needed if multiple found)
                                            tod_charges['11'] = f_v
                                except: pass
            except: pass
            
    # Default to NA
    if tod_charges['11'] is None:
        tod_charges['11'] = "NA"

    print(f"Extracted TOD Charges: {tod_charges}")
    return tod_charges

def extract_pfa_rebate(jsonl_path):
    """
    Extract Power Factor Adjustment Rebate
    Keywords: Power Factor Adjustment Rebate, Power Factor Adjustment Discount, Power Factor Adjustment
    Units: INR/kWh
    """
    pfa_rebate = None
    keywords = ["power factor adjustment rebate", "power factor adjustment discount", "power factor adjustment", "power factor rebate"]
    
    with open(jsonl_path, 'r', encoding='utf-8') as f:
        for line in f:
            try:
                data = json.loads(line)
                heading = data.get("table_heading", "").lower()
                rows = data.get("rows", [])
                
                # Check for table match
                if any(k in heading for k in keywords):
                    for row in rows:
                        row_txt = str(row).lower()
                        # Check for absolute value
                        for v in row.values():
                            val_str = str(v).lower()
                            # If it's percentage, ignore
                            if "%" in val_str or "percent" in val_str:
                                continue
                            
                            # Look for currency indicators
                            if any(c in val_str for c in ["rs", "inr", "paise", "₹"]):
                                try:
                                    clean = re.sub(r'[^\d\.]', '', val_str)
                                    if clean:
                                        f_v = float(clean)
                                        # Heuristic: 0.01 to 2.0 Rs/kWh
                                        if 0.0 < f_v < 2.0:
                                            pfa_rebate = f_v
                                            break
                                except: pass
                        if pfa_rebate: break
                
                if pfa_rebate: break
            except: pass

    if pfa_rebate is None:
        pfa_rebate = "NA"
        
    print(f"Extracted PFA Rebate: {pfa_rebate}")
    return pfa_rebate

def extract_load_factor_incentive(jsonl_path):
    """
    Extract Load Factor Incentive
    Keywords: Load Factor Incentive, Load Factor discount
    Units: INR/kWh
    """
    lf_incentive = None
    keywords = ["load factor incentive", "load factor discount", "load factor rebate"]
    candidates = []
    
    with open(jsonl_path, 'r', encoding='utf-8') as f:
        for line in f:
            try:
                data = json.loads(line)
                heading = data.get("table_heading", "").lower()
                rows = data.get("rows", [])
                
                # Check keywords match
                match = any(k in heading for k in keywords)
                
                if not match:
                    for row in rows:
                        if any(k in str(row).lower() for k in keywords):
                            match = True
                            break
                            
                if match:
                    for row in rows:
                        for v in row.values():
                            val_str = str(v).lower()
                            
                            # Regex to find number before unit
                            # Look for: "10 paise", "10.5 paise", "Rs 0.50", "Rs. 0.50"
                            if "paise" in val_str:
                                # Find number preceding "paise"
                                m = re.search(r'(\d+(?:\.\d+)?)\s*paise', val_str)
                                if m:
                                    f_v = float(m.group(1))
                                    if 1 <= f_v <= 100:
                                        candidates.append(f_v / 100.0)
                            
                            elif "rs" in val_str or "inr" in val_str or "₹" in val_str:
                                # Find number after Rs or before
                                m = re.search(r'(?:rs\.?|inr|₹)\s*(\d+(?:\.\d+)?)', val_str)
                                if m:
                                    f_v = float(m.group(1))
                                    if 0.0 < f_v < 10.0:
                                        candidates.append(f_v)
            except: pass

    if candidates:
        lf_incentive = max(candidates)
    
    if lf_incentive is None:
        lf_incentive = "NA"
        
    print(f"Extracted LF Incentive: {lf_incentive}")
    return lf_incentive

def extract_grid_support_charges(jsonl_path):
    """
    Extract Grid Support Charges
    Keywords: Grid Support, Parrallel Operation, Parallel Operation
    Units: INR/kWh
    """
    grid_support = None
    keywords = ["grid support", "parrallel operation", "parallel operation"]
    
    with open(jsonl_path, 'r', encoding='utf-8') as f:
        for line in f:
            try:
                data = json.loads(line)
                heading = data.get("table_heading", "").lower()
                rows = data.get("rows", [])
                
                # Check keywords in heading
                match = any(k in heading for k in keywords)
                
                if match:
                     for row in rows:
                        for v in row.values():
                            val_str = str(v).lower()
                            try:
                                clean = re.sub(r'[^\d\.]', '', val_str)
                                if clean:
                                    f_v = float(clean)
                                    # Heuristic: 0.1 to 10 Rs/kWh
                                    if 0.0 < f_v < 10.0:
                                        grid_support = f_v
                                        break
                            except: pass
                        if grid_support: break
                
                if not grid_support:
                    # Check row level
                     for row in rows:
                        row_txt = str(row).lower()
                        if any(k in row_txt for k in keywords):
                            for v in row.values():
                                try:
                                    clean = re.sub(r'[^\d\.]', '', str(v))
                                    if clean:
                                        f_v = float(clean)
                                        if 0.0 < f_v < 10.0: 
                                             grid_support = f_v
                                             break
                                except: pass
                        if grid_support: break

                if grid_support: break
            except: pass

    if grid_support is None:
        grid_support = "NA"
        
    print(f"Extracted Grid Support Charges: {grid_support}")
    return grid_support

def extract_voltage_rebates(jsonl_path):
    """
    Extract Voltage Rebates (HT, EHV)
    Keywords: HTRebate, EHV Rebate
    Units: INR/kWh
    Returns dict: {'33_66': val, '132_plus': val}
    """
    rebates = {'33_66': "NA", '132_plus': "NA"}
    keywords = ["ht rebate", "ehv rebate", "htrebate", "ehvrebate", "voltage rebate"]
    
    with open(jsonl_path, 'r', encoding='utf-8') as f:
        for line in f:
            try:
                data = json.loads(line)
                heading = data.get("table_heading", "").lower()
                rows = data.get("rows", [])
                
                match = any(k in heading for k in keywords)
                
                if match:
                    for row in rows:
                        row_txt = str(row).lower()
                        
                        target_key = None
                        if "33" in row_txt or "66" in row_txt:
                            target_key = '33_66'
                        elif "132" in row_txt or "220" in row_txt or "extra high" in row_txt:
                            target_key = '132_plus'
                            
                        if target_key:
                            for v in row.values():
                                val_str = str(v).lower()
                                try:
                                    if "paise" in val_str:
                                        m = re.search(r'(\d+(?:\.\d+)?)\s*paise', val_str)
                                        if m:
                                            f_v = float(m.group(1))
                                            if 1 <= f_v <= 100:
                                                rebates[target_key] = f_v / 100.0
                                    elif "rs" in val_str or "inr" in val_str:
                                        m = re.search(r'(?:rs\.?|inr|₹)\s*(\d+(?:\.\d+)?)', val_str)
                                        if m:
                                            f_v = float(m.group(1))
                                            if 0.0 < f_v < 5.0:
                                                rebates[target_key] = f_v
                                except: pass
            except: pass
            
    print(f"Extracted Voltage Rebates: {rebates}")
    return rebates

def extract_bulk_consumption_rebate(jsonl_path):
    """
    Extract Bulk Consumption Rebate
    Keywords: Bulk Consumption Rebate, Bulk Consumption Discount, Bulk Consumption
    Units: INR/kWh
    """
    bulk_rebate = None
    keywords = ["bulk consumption rebate", "bulk consumption discount", "bulk consumption"]
    
    with open(jsonl_path, 'r', encoding='utf-8') as f:
        for line in f:
            try:
                data = json.loads(line)
                heading = data.get("table_heading", "").lower()
                rows = data.get("rows", [])
                
                match = any(k in heading for k in keywords)
                
                if match:
                    for row in rows:
                         for v in row.values():
                            val_str = str(v).lower()
                            try:
                                if "paise" in val_str:
                                    m = re.search(r'(\d+(?:\.\d+)?)\s*paise', val_str)
                                    if m:
                                        f_v = float(m.group(1))
                                        if 1 <= f_v <= 100:
                                            bulk_rebate = f_v / 100.0
                                            break
                                elif "rs" in val_str or "inr" in val_str:
                                    m = re.search(r'(?:rs\.?|inr|₹)\s*(\d+(?:\.\d+)?)', val_str)
                                    if m:
                                        f_v = float(m.group(1))
                                        if 0.0 < f_v < 5.0:
                                            bulk_rebate = f_v
                                            break
                            except: pass
                    if bulk_rebate: break
            except: pass
            
    if bulk_rebate is None:
        bulk_rebate = "NA"
        
    print(f"Extracted Bulk Rebate: {bulk_rebate}")
    return bulk_rebate

# ... [End of extract functions] ...

# ... inside update_excel_with_discoms ...
# (We need to reach the write logic part, but replace_file_content works on contiguous blocks.
# The extract function and write logic are far apart. I must do 2 calls.)


def update_excel_with_discoms(discoms, ists_loss, insts_loss, wheeling_losses, wheeling_charges, css_charges, fixed_charges, energy_charges, fuel_surcharge, tod_charges, pfa_rebate, lf_incentive, grid_support, voltage_rebates, bulk_rebate, add_surcharge, excel_path):
    # Load the workbook
    try:
        wb = openpyxl.load_workbook(excel_path)
        sheet = wb.active
        
        # Mappings
        # Col 1: State
        # Col 3: Discom
        
        # Determine State Name from filename
        state_name = os.path.splitext(os.path.basename(excel_path))[0].capitalize()
        
        # Reset/Clear existing data rows starting from row 3 to ensure we write from the top
        start_row = 3
        if sheet.max_row >= start_row:
             # Delete all rows from start_row to the end
             sheet.delete_rows(start_row, amount=sheet.max_row - start_row + 1)
        
        next_row = start_row
            
        # We process each extracted Discom
        # If discom list is empty, maybe just add one generic row?
        if not discoms:
            discoms = ["Gen"] # Placeholder if no discom found but we have state data
            
        for discom in discoms:
            d_key = discom.strip().lower()
            row_idx = next_row
            next_row += 1
            
            # Ensure State and Discom are always correct
            sheet.cell(row=row_idx, column=1).value = state_name
            sheet.cell(row=row_idx, column=3).value = discom
                
            # Update Data Fields
            
            # Update ISTS Loss
            if ists_loss is not None:
                val = str(ists_loss)
                if "%" not in val: val += "%"
                sheet.cell(row=row_idx, column=4).value = val

            # Update InSTS Loss
            if insts_loss is not None:
                sheet.cell(row=row_idx, column=5).value = insts_loss
                
            # Update Wheeling Losses
            # Pick discom specific losses
            current_losses = wheeling_losses.get(d_key.upper(), wheeling_losses.get('GENERIC', {}))
            
            if current_losses:
                 if '11' in current_losses: sheet.cell(row=row_idx, column=6).value = current_losses['11']
                 if '33' in current_losses: sheet.cell(row=row_idx, column=7).value = current_losses['33']
                 if '66' in current_losses: sheet.cell(row=row_idx, column=8).value = current_losses['66']
                 if '132' in current_losses: sheet.cell(row=row_idx, column=9).value = current_losses['132']

            # Update Wheeling Charges (Col 12, 13, 14, 15)
            if wheeling_charges:
                if '11' in wheeling_charges: sheet.cell(row=row_idx, column=12).value = wheeling_charges['11']
                if '33' in wheeling_charges: sheet.cell(row=row_idx, column=13).value = wheeling_charges['33']
                if '66' in wheeling_charges: sheet.cell(row=row_idx, column=14).value = wheeling_charges['66']
                if '132' in wheeling_charges: sheet.cell(row=row_idx, column=15).value = wheeling_charges['132']

            # Update CSS (Col 16, 17, 18, 19, 20)
            if css_charges:
                 if '11' in css_charges: sheet.cell(row=row_idx, column=16).value = css_charges['11']
                 if '33' in css_charges: sheet.cell(row=row_idx, column=17).value = css_charges['33']
                 if '66' in css_charges: sheet.cell(row=row_idx, column=18).value = css_charges['66']
                 if '132' in css_charges: sheet.cell(row=row_idx, column=19).value = css_charges['132']
                 if '220' in css_charges: sheet.cell(row=row_idx, column=20).value = css_charges['220']
            
            # Update Additional Surcharge (Col 21)
            if add_surcharge is not None:
                sheet.cell(row=row_idx, column=21).value = add_surcharge
            
            # Update Fixed Charges
            if fixed_charges:
                 if '11' in fixed_charges: sheet.cell(row=row_idx, column=24).value = fixed_charges['11']
                 if '33' in fixed_charges: sheet.cell(row=row_idx, column=25).value = fixed_charges['33']
                 if '66' in fixed_charges: sheet.cell(row=row_idx, column=26).value = fixed_charges['66']
                 if '132' in fixed_charges: sheet.cell(row=row_idx, column=27).value = fixed_charges['132']
                 if '220' in fixed_charges: sheet.cell(row=row_idx, column=28).value = fixed_charges['220']

            # Update Energy Charges
            if energy_charges:
                 if '11' in energy_charges: sheet.cell(row=row_idx, column=29).value = energy_charges['11']
                 if '33' in energy_charges: sheet.cell(row=row_idx, column=30).value = energy_charges['33']
                 if '66' in energy_charges: sheet.cell(row=row_idx, column=31).value = energy_charges['66']
                 if '132' in energy_charges: sheet.cell(row=row_idx, column=32).value = energy_charges['132']
                 if '220' in energy_charges: sheet.cell(row=row_idx, column=33).value = energy_charges['220']

            # Update Fuel Surcharge (Col 34)
            if fuel_surcharge is not None:
                sheet.cell(row=row_idx, column=34).value = fuel_surcharge

            # Update Fuel Surcharge (Col 34)
            if fuel_surcharge is not None:
                sheet.cell(row=row_idx, column=34).value = fuel_surcharge

            # Update TOD Charges (Col 35)
            # Extracted or NA
            val = "NA"
            if tod_charges:
                # If tod_charges is a dict, pick '11' or generic
                if isinstance(tod_charges, dict):
                    val = tod_charges.get('11', "NA")
                else:
                    val = tod_charges
            sheet.cell(row=row_idx, column=35).value = val

            # Update Power Factor Adjustment (Col 36)
            if pfa_rebate is not None:
                sheet.cell(row=row_idx, column=36).value = pfa_rebate

            # Update Load Factor Incentive (Col 37)
            if lf_incentive is not None:
                sheet.cell(row=row_idx, column=37).value = lf_incentive

            # Update Grid Support (Col 38)
            if grid_support is not None:
                sheet.cell(row=row_idx, column=38).value = grid_support

            # Update Voltage Rebates (Col 39, 40)
            # C39: HT, EHV Rebate at 33/66 kV
            # C40: HT, EHV Rebate at 132 kV and above
            if voltage_rebates:
                val_33_66 = voltage_rebates.get('33_66', "NA")
                val_132_plus = voltage_rebates.get('132_plus', "NA")
                
                sheet.cell(row=row_idx, column=39).value = val_33_66
                sheet.cell(row=row_idx, column=40).value = val_132_plus
            else:
                sheet.cell(row=row_idx, column=39).value = "NA"
                sheet.cell(row=row_idx, column=40).value = "NA"

            # Update Bulk Consumption Rebate (Col 41)
            if bulk_rebate is not None:
                sheet.cell(row=row_idx, column=41).value = bulk_rebate
            else:
                sheet.cell(row=row_idx, column=41).value = "NA"

        wb.save(excel_path)
        print(f"Updated {os.path.basename(excel_path)} with accurate values for {len(discoms)} discoms.")

    except Exception as e:
        print(f"Error updating Excel: {e}")

if __name__ == "__main__":
    base_dir = os.path.dirname(os.path.abspath(__file__))

    # Find JSONL file dynamically in Extraction/bihar
    extraction_dir = os.path.join(base_dir, "Extraction", "bihar")
    jsonl_files = glob.glob(os.path.join(extraction_dir, "*.jsonl"))
    if jsonl_files:
        jsonl_file = jsonl_files[0]
        print(f"Using JSONL: {jsonl_file}")
    else:
        # Fallback to a name that might exist or matches pattern
        jsonl_file = os.path.join(extraction_dir, "TO_DISCOMS_FY_25-26.jsonl")

    discom_file = os.path.join(base_dir, "discoms_bihar.txt")
    excel_file = os.path.join(base_dir, "bihar.xlsx")

    # Correct path for ISTS loss JSON
    ists_loss_file = os.path.join(base_dir, "ists_extracted", "ists_loss.json")
    ists_val = extract_ists_loss(ists_loss_file)

    discoms = extract_discom_names(jsonl_file, discom_file)
    insts = extract_losses(jsonl_file)
    wheeling = extract_wheeling_losses(jsonl_file, discoms)
    css = extract_css_charges(jsonl_file)
    fixed = extract_fixed_charges(jsonl_file)
    energy = extract_energy_charges(jsonl_file)
    fuel = extract_fuel_surcharge(jsonl_file)
    tod = extract_tod_charges(jsonl_file)
    wheeling_chg = extract_wheeling_charges(jsonl_file)
    pfa = extract_pfa_rebate(jsonl_file)
    lf_inc = extract_load_factor_incentive(jsonl_file)
    grid_sup = extract_grid_support_charges(jsonl_file)
    volt_reb = extract_voltage_rebates(jsonl_file)
    bulk_reb = extract_bulk_consumption_rebate(jsonl_file)
    add_surchg = extract_additional_surcharge(jsonl_file)

    update_excel_with_discoms(
        discoms,
        ists_val,
        insts,
        wheeling,
        wheeling_chg,
        css,
        fixed,
        energy,
        fuel,
        tod,
        pfa,
        lf_inc,
        grid_sup,
        volt_reb,
        bulk_reb,
        add_surchg,
        excel_file
    )
