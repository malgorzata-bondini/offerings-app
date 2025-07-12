import datetime as dt
import re, warnings
from pathlib import Path
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Alignment

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

need_cols = [
    "Name (Child Service Offering lvl 1)", "Parent Offering",
    "Service Offerings | Depend On (Application Service)", "Service Commitments",
    "Delivery Manager", "Subscribed by Location", "Phase", "Status",
    "Life Cycle Stage", "Life Cycle Status", "Support group", "Managed by Group",
    "Subscribed by Company", "Visibility group",
]

discard_lc  = {"retired", "retiring", "end of life", "end of support"}

def extract_parent_info(parent_offering):
    """Extract the content between [Parent ...] from parent offering"""
    match = re.search(r'\[Parent\s+(.*?)\]', str(parent_offering), re.I)
    if match:
        return match.group(1).strip()
    return ""

def extract_catalog_name(parent_offering):
    """Extract the catalog name after the brackets"""
    parts = str(parent_offering).split(']', 1)
    if len(parts) > 1:
        return parts[1].strip()
    return ""

def build_standard_name(parent_offering, sr_or_im, app, schedule_suffix, special_dept=None):
    """Build standard name when not CORP"""
    parent_content = extract_parent_info(parent_offering)
    catalog_name = extract_catalog_name(parent_offering)
    
    # Check if catalog name or parent offering contains keywords that exclude "Prod"
    no_prod_keywords = ["hardware", "mailbox", "network"]
    parent_lower = parent_offering.lower()
    catalog_lower = catalog_name.lower()
    exclude_prod = any(keyword in parent_lower or keyword in catalog_lower for keyword in no_prod_keywords)
    
    if special_dept == "Medical":
        # Extract division and country from parent content
        parts = parent_content.split()
        division = ""
        country = ""
        topic = ""
        
        for i, part in enumerate(parts):
            if part in ["HS", "DS"]:
                division = part
            elif len(part) == 2 and part.isupper() and part not in ["IT", "HR"]:
                country = part
            elif part not in ["HS", "DS"] and not (len(part) == 2 and part.isupper()):
                topic = part
                break
        
        # Build Medical name - NO PROD
        prefix_parts = [sr_or_im]
        if division:
            prefix_parts.append(division)
        if country:
            prefix_parts.append(country)
        prefix_parts.append("Medical")
        
        # Use topic from parent and lowercase catalog name
        return f"[{' '.join(prefix_parts)}] {topic} {catalog_name.lower()} {schedule_suffix}"
    
    elif special_dept == "DAK":
        # Replace DAK with Business Services - NO PROD
        parts = parent_content.split()
        division = ""
        country = ""
        
        for part in parts:
            if part in ["HS", "DS"]:
                division = part
            elif len(part) == 2 and part.isupper() and part not in ["IT", "HR", "DAK"]:
                country = part
        
        prefix_parts = [sr_or_im]
        if division:
            prefix_parts.append(division)
        if country:
            prefix_parts.append(country)
        prefix_parts.append("Business Services")
        
        # Add app only if provided - NO PROD
        if app:
            return f"[{' '.join(prefix_parts)}] {catalog_name} {app} {schedule_suffix}"
        else:
            return f"[{' '.join(prefix_parts)}] {catalog_name} {schedule_suffix}"
    
    elif special_dept == "HR":
        # HR - NO PROD
        parts = parent_content.split()
        division = ""
        country = ""
        for part in parts:
            if part in ["HS", "DS"]:
                division = part
            elif len(part) == 2 and part.isupper() and part not in ["IT", "HR"]:
                country = part
        
        prefix_parts = [sr_or_im]
        if division:
            prefix_parts.append(division)
        if country:
            prefix_parts.append(country)
        prefix_parts.append("HR")
        
        # Extract the topic from parent content
        topic_parts = []
        for part in parts:
            if part not in ["HS", "DS"] and not (len(part) == 2 and part.isupper()):
                topic_parts.append(part)
        
        topic = " ".join(topic_parts) if topic_parts else "Software"
        
        # Add app if provided - NO PROD
        if app:
            return f"[{' '.join(prefix_parts)}] {topic} {catalog_name.lower()} {app} {schedule_suffix}"
        else:
            return f"[{' '.join(prefix_parts)}] {topic} {catalog_name.lower()} {schedule_suffix}"
    
    elif special_dept == "IT":
        # IT - special handling
        parts = parent_content.split()
        division = ""
        country = ""
        topic = ""
        
        # Find division, country and topic
        for i, part in enumerate(parts):
            if part in ["HS", "DS"]:
                division = part
            elif len(part) == 2 and part.isupper() and part not in ["IT", "HR"]:
                country = part
            elif part not in ["HS", "DS"] and not (len(part) == 2 and part.isupper()):
                # This is the topic (e.g., "Hardware", "Software", etc.)
                topic = part
                break
        
        prefix_parts = [sr_or_im]
        if division:
            prefix_parts.append(division)
        if country:
            prefix_parts.append(country)
        prefix_parts.append("IT")
        
        # For IT, use the topic from parent (e.g., "Hardware") and lowercase catalog name
        name_parts = [f"[{' '.join(prefix_parts)}]", topic, catalog_name.lower()]
        
        # Add app if provided
        if app:
            name_parts.append(app)
        
        # Add "solving" for IM
        if sr_or_im == "IM":
            name_parts.append("solving")
        
        # Only add Prod if no hardware/mailbox/network keywords
        if not exclude_prod:
            name_parts.append("Prod")
        
        name_parts.append(schedule_suffix)
        
        return " ".join(name_parts)
    
    else:
        # Standard case - just replace Parent with SR/IM
        if app:
            return f"[{sr_or_im} {parent_content}] {catalog_name} {app} Prod {schedule_suffix}"
        else:
            return f"[{sr_or_im} {parent_content}] {catalog_name} Prod {schedule_suffix}"

def build_corp_name(parent_offering, sr_or_im, app, schedule_suffix, receiver, delivering_tag):
    """Build name for CORP offerings"""
    parent_content = extract_parent_info(parent_offering)
    catalog_name = extract_catalog_name(parent_offering)
    
    # Extract parts from parent content
    parts = parent_content.split()
    country = ""
    topic = ""
    
    for part in parts:
        if len(part) == 2 and part.isupper() and part not in ["HS", "DS"]:
            country = part
        elif part not in ["HS", "DS", "Parent", "RecP"]:
            topic = part
    
    # Build CORP name
    prefix_parts = [sr_or_im]
    delivering_parts = delivering_tag.split() if delivering_tag else ["HS", country]
    prefix_parts.extend(delivering_parts[:2])  # Take division and country from delivering tag
    prefix_parts.extend(["CORP", receiver])
    # Only add topic if it exists, don't default to IT for CORP
    if topic:
        prefix_parts.append(topic)
    
    if sr_or_im == "SR":
        if app:
            return f"[{' '.join(prefix_parts)}] {catalog_name} {app} Prod {schedule_suffix}"
        else:
            return f"[{' '.join(prefix_parts)}] {catalog_name} Prod {schedule_suffix}"
    else:
        if app:
            return f"[{' '.join(prefix_parts)}] {catalog_name} solving {app} Prod {schedule_suffix}"
        else:
            return f"[{' '.join(prefix_parts)}] {catalog_name} solving Prod {schedule_suffix}"

def update_commitments(orig, sched, rsp, rsl):
    """Update existing commitments and ensure OLA is present"""
    out = []
    has_ola = False
    country_code = None
    
    for line in str(orig).splitlines():
        line = line.strip()
        if not line:
            continue
            
        if "RSP" in line:
            # Extract country code from line like [PL] SLA SR RSP...
            match = re.search(r'\[(\w+)\]', line)
            if match:
                country_code = match.group(1)
            # Extract P values (P1-P4, P1-P3, etc)
            p_match = re.search(r'(P\d+-P\d+)', line)
            p_values = p_match.group(1) if p_match else "P1-P4"
            # Update schedule and duration
            line = re.sub(r"RSP\s+[^P]+", f"RSP {sched} ", line)
            line = re.sub(r"(P\d+-P\d+)\s+.*$", f"{p_values} {rsp}", line)
        elif "RSL" in line:
            # Extract P values
            p_match = re.search(r'(P\d+-P\d+)', line)
            p_values = p_match.group(1) if p_match else "P1-P4"
            # Update schedule and duration
            line = re.sub(r"RSL\s+[^P]+", f"RSL {sched} ", line)
            line = re.sub(r"(P\d+-P\d+)\s+.*$", f"{p_values} {rsl}", line)
        elif "OLA" in line:
            has_ola = True
            # Extract P values
            p_match = re.search(r'(P\d+-P\d+)', line)
            p_values = p_match.group(1) if p_match else "P1-P4"
            # Update schedule and duration
            line = re.sub(r"RSL\s+[^P]+", f"RSL {sched} ", line)
            line = re.sub(r"(P\d+-P\d+)\s+.*$", f"{p_values} {rsl}", line)
        out.append(line)
    
    # If no OLA found, create it from the last RSL line
    if not has_ola and country_code:
        # Find the last RSL line to copy its format
        rsl_line = None
        for line in out:
            if "RSL" in line and "SLA" in line:
                rsl_line = line
        
        if rsl_line:
            # Create OLA by replacing SLA with OLA in the RSL line
            ola_line = rsl_line.replace("SLA", "OLA")
            out.append(ola_line)
    
    return "\n".join(out)

def commit_block(cc, schedule_suffix, rsp_duration, rsl_duration):
    """Create commitment block with OLA for all countries"""
    lines=[
        f"[{cc}] SLA SR RSP {schedule_suffix} P1-P4 {rsp_duration}",
        f"[{cc}] SLA SR RSL {schedule_suffix} P1-P4 {rsl_duration}",
        f"[{cc}] OLA SR RSL {schedule_suffix} P1-P4 {rsl_duration}"
    ]
    return "\n".join(lines)

def run_generator(*,
    keywords_parent, keywords_child, new_apps, schedule_suffixes,
    delivery_manager, global_prod,
    rsp_duration, rsl_duration,
    sr_or_im, require_corp, delivering_tag,
    support_group, managed_by_group, aliases_on, aliases_value,
    src_dir: Path, out_dir: Path,
    special_it=False, special_hr=False, special_medical=False, special_dak=False):

    sheets, seen = {}, set()
    existing_offerings = set()  # Track existing offerings to detect duplicates

    def parse_keywords(keyword_string):
        """Parse keywords - returns (keywords_list, use_and_logic)"""
        if not keyword_string.strip():
            return [], False
        
        # Check if it contains commas (AND logic)
        if ',' in keyword_string:
            keywords = [k.strip() for k in keyword_string.split(',') if k.strip()]
            return keywords, True
        else:
            # Line separated (OR logic)
            keywords = [k.strip() for k in keyword_string.split('\n') if k.strip()]
            return keywords, False

    def row_keywords_ok(row):
        # Parse parent keywords
        parent_keywords, parent_use_and = parse_keywords(keywords_parent)
        
        # Check parent offering
        if parent_keywords:
            p = str(row["Parent Offering"]).lower()
            if parent_use_and:
                # AND logic - all keywords must match
                if not all(re.search(rf"\b{re.escape(k.lower())}\b", p) for k in parent_keywords):
                    return False
            else:
                # OR logic - any keyword must match
                if not any(re.search(rf"\b{re.escape(k.lower())}\b", p) for k in parent_keywords):
                    return False
        
        # Parse child keywords
        child_keywords, child_use_and = parse_keywords(keywords_child)
        
        # Check child name
        if child_keywords:
            n = str(row["Name (Child Service Offering lvl 1)"]).lower()
            if child_use_and:
                # AND logic - all keywords must match
                if not all(re.search(rf"\b{re.escape(k.lower())}\b", n) for k in child_keywords):
                    return False
            else:
                # OR logic - any keyword must match
                if not any(re.search(rf"\b{re.escape(k.lower())}\b", n) for k in child_keywords):
                    return False
        
        return True

    def lc_ok(row):
        return all(str(row[c]).strip().lower() not in discard_lc
                   for c in ("Phase","Status","Life Cycle Stage","Life Cycle Status"))

    def name_prefix_ok(name):
        return name.lower().startswith(f"[{sr_or_im.lower()} ")

    # Process apps - support both newline and comma separation
    all_apps = []
    for app_line in new_apps:
        # Split by comma and add non-empty values
        for app in app_line.split(','):
            app = app.strip()
            if app:
                all_apps.append(app)
    
    # If no apps provided, use empty string as placeholder
    if not all_apps:
        all_apps = [""]

    # Determine special department
    special_dept = None
    if special_it and not require_corp:
        special_dept = "IT"
    elif special_hr and not require_corp:
        special_dept = "HR"
    elif special_medical and not require_corp:
        special_dept = "Medical"
    elif special_dak and not require_corp:
        special_dept = "DAK"

    # First, collect all existing offerings from the source files
    for wb in src_dir.glob("ALL_Service_Offering_*.xlsx"):
        try:
            df = pd.read_excel(wb, sheet_name="Child SO lvl1")
            if "Name (Child Service Offering lvl 1)" in df.columns:
                # Clean the names before adding to set - remove extra whitespace and convert to string
                existing_names = df["Name (Child Service Offering lvl 1)"].dropna().astype(str).str.strip()
                existing_offerings.update(existing_names)
        except Exception:
            # Skip if there's an error reading the file
            continue

    # Now process the files
    for wb in src_dir.glob("ALL_Service_Offering_*.xlsx"):
        df = pd.read_excel(wb, sheet_name="Child SO lvl1")
        
        # Add missing columns as empty
        for col in need_cols:
            if col not in df.columns:
                df[col] = ""

        mask=(df.apply(row_keywords_ok,axis=1)
              & df["Name (Child Service Offering lvl 1)"].astype(str).apply(name_prefix_ok)
              & df.apply(lc_ok,axis=1)
              & (df["Service Commitments"].astype(str).str.strip().replace({"nan":""})!="-"))
        
        if require_corp:
            mask &= df["Name (Child Service Offering lvl 1)"].str.contains(r"\bCORP\b",case=False)
            mask &= df["Name (Child Service Offering lvl 1)"].str.contains(rf"\b{re.escape(delivering_tag)}\b",case=False)
        else:
            mask &= ~df["Name (Child Service Offering lvl 1)"].str.contains(r"\bCORP\b",case=False)

        base_pool=df.loc[mask]
        if base_pool.empty:
            continue
        
        # Process ALL matching rows, not just the first one
        for idx, base_row in base_pool.iterrows():
            base_row_df = base_row.to_frame().T.copy()
            country=wb.stem.split("_")[-1].upper()
            tag_hs, tag_ds = f"HS {country}", f"DS {country}"

            # Determine receivers based on country and CORP setting
            if require_corp:
                if country=="DE":         
                    receivers=["DS DE","HS DE"]
                elif country in {"UA","MD"}: 
                    receivers=[f"DS {country}"]
                elif country=="PL":         
                    # Check if DS PL exists in the data
                    if "DS PL" in base_pool["Name (Child Service Offering lvl 1)"].str.cat(sep=" "):
                        receivers=["DS PL"]
                    else:
                        receivers=["HS PL"]
                elif country=="CY":         
                    receivers=["DS CY","HS CY"]
                else:                       
                    receivers=[f"DS {country}"]
            else:
                receivers=[""]

            parent_full=str(base_row["Parent Offering"])

            for app in all_apps:
                for recv in receivers:
                    for schedule_suffix in schedule_suffixes:
                        if require_corp:
                            new_name = build_corp_name(
                                parent_full, sr_or_im, app, schedule_suffix, recv, delivering_tag
                            )
                        else:
                            new_name = build_standard_name(
                                parent_full, sr_or_im, app, schedule_suffix, special_dept
                            )
                        
                        # Normalize the name for comparison (remove extra spaces)
                        new_name_normalized = ' '.join(new_name.split())
                        
                        # Check for duplicates within this generation run
                        if new_name_normalized in seen:
                            # Skip this one instead of raising error - it's already being created
                            continue
                        
                        # Check against existing offerings in source files
                        found_in_existing = False
                        for existing in existing_offerings:
                            existing_normalized = ' '.join(str(existing).split())
                            if existing_normalized == new_name_normalized:
                                found_in_existing = True
                                break
                        
                        if found_in_existing:
                            raise ValueError(f"Sorry, it would be a duplicate - we already have this offering in the system: {new_name}")
                        
                        seen.add(new_name_normalized)

                        row=base_row_df.copy()
                        row["Name (Child Service Offering lvl 1)"]=new_name
                        row["Delivery Manager"]=delivery_manager
                        row["Support group"]=support_group
                        # If Managed by Group is empty but Support Group is filled, copy Support Group
                        row["Managed by Group"]=managed_by_group if managed_by_group else support_group
                        
                        # Handle aliases
                        for c in [c for c in row.columns if "Aliases" in c]:
                            row[c]=aliases_value if aliases_on else "-"
                        
                        if country=="DE":
                            row["Subscribed by Company"]="DE Internal Patients\nDE External Patients" if recv=="HS DE" else "DE IFLB Laboratories\nDE IMD Laboratories"
                        elif country=="UA":
                            row["Subscribed by Company"]="Сiнево Україна"
                        elif country=="CY":
                            row["Subscribed by Company"]="CY Healthcare Services\nCY Medical Centers" if recv=="HS CY" else "CY Diagnostic Laboratories"
                        else:
                            row["Subscribed by Company"]=recv or tag_hs
                            
                        orig_comm=str(row.iloc[0]["Service Commitments"]).strip()
                        row["Service Commitments"]=commit_block(country, schedule_suffix, rsp_duration, rsl_duration) if not orig_comm or orig_comm=="-" else update_commitments(orig_comm,schedule_suffix,rsp_duration,rsl_duration)
                        
                        if global_prod:
                            if app:
                                row["Service Offerings | Depend On (Application Service)"]=f"[Global Prod] {app}"
                            else:
                                row["Service Offerings | Depend On (Application Service)"]="[Global Prod]"
                        else:
                            depend_tag = f"{delivering_tag} Prod" if require_corp else f"{recv or tag_hs} Prod"
                            if app:
                                row["Service Offerings | Depend On (Application Service)"]=f"[{depend_tag}] {app}"
                            else:
                                row["Service Offerings | Depend On (Application Service)"]=f"[{depend_tag}]"
                        
                        sheets.setdefault(country,pd.DataFrame())
                        sheets[country]=pd.concat([sheets[country],row],ignore_index=True)

    if not sheets:
        raise ValueError("No matching rows found with the specified keywords.")

    out_dir.mkdir(parents=True,exist_ok=True)
    outfile=out_dir / f"Offerings_NEW_{dt.datetime.now():%Y%m%d_%H%M%S}.xlsx"
    
    # Write to Excel with special handling for empty values
    with pd.ExcelWriter(outfile,engine="openpyxl") as w:
        for cc,dfc in sheets.items():
            # Ensure unique names per country
            df_final = dfc.drop_duplicates(subset=["Name (Child Service Offering lvl 1)"])
            
            if "Number" in df_final.columns:
                df_final = df_final.drop(columns=["Number"])
            
            # Replace all forms of empty/null values with empty string
            df_final = df_final.fillna('')
            
            for col in df_final.columns:
                if df_final[col].dtype == 'bool':
                    df_final[col] = df_final[col].map({True: 'true', False: 'false'})
                elif df_final[col].dtype == 'object':
                    # Replace all variants of empty values
                    df_final[col] = df_final[col].astype(str).replace({
                        'nan': '', 
                        'NaN': '', 
                        'None': '', 
                        'none': '',
                        'NULL': '',
                        'null': '',
                        '<NA>': '',
                        'True': 'true', 
                        'False': 'false'
                    })
                    # Also handle when the string is literally "nan"
                    df_final[col] = df_final[col].apply(lambda x: '' if str(x).lower() == 'nan' else x)
            
            df_final.to_excel(w,sheet_name=cc,index=False)
    
    # Apply formatting
    wb=load_workbook(outfile)
    for ws in wb.worksheets:
        ws.auto_filter.ref=ws.dimensions
        for col in ws.columns:
            ws.column_dimensions[col[0].column_letter].width=max(len(str(c.value)) if c.value else 0 for c in col)+2
            for c in col:
                c.alignment=Alignment(wrap_text=True)
                # Ensure empty cells stay empty in Excel
                if c.value in ['nan', 'NaN', 'None', None, 'none', 'NULL', 'null', '<NA>']:
                    c.value = None
    wb.save(outfile)
    return outfile