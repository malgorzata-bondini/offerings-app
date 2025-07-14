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

def build_corp_it_name(parent_offering, sr_or_im, app, schedule_suffix, receiver, delivering_tag):
    """Build name for CORP IT offerings"""
    parent_content = extract_parent_info(parent_offering)
    catalog_name = extract_catalog_name(parent_offering)
    
    # Extract parts from parent content
    parts = parent_content.split()
    country = ""
    topic = ""
    
    for part in parts:
        if len(part) == 2 and part.isupper() and part not in ["HS", "DS"]:
            country = part
        elif part not in ["HS", "DS", "Parent", "RecP"] and not (len(part) == 2 and part.isupper()):
            topic = part
            break
    
    # Build CORP IT name - always ends with IT
    prefix_parts = [sr_or_im]
    
    # Add delivering tag parts (who delivers the service - from user input)
    delivering_parts = delivering_tag.split() if delivering_tag else ["HS", country]
    prefix_parts.extend(delivering_parts)
    
    prefix_parts.append("CORP")
    
    # Add receiver parts (who receives the service - DS DE or HS DE for Germany)
    if receiver:
        prefix_parts.append(receiver)
    else:
        prefix_parts.extend(delivering_parts)
    
    prefix_parts.append("IT")
    
    # Build the name
    name_prefix = f"[{' '.join(prefix_parts)}]"
    
    # Use topic from parent or "Software" as default
    if not topic:
        topic = "Software"
    
    name_parts = [name_prefix, topic, catalog_name.lower()]
    
    # Add app if provided
    if app:
        name_parts.append(app)
    
    # Add "solving" for IM
    if sr_or_im == "IM":
        name_parts.append("solving")
    
    #name_parts.append("Prod")
    name_parts.append(schedule_suffix)
    
    return " ".join(name_parts)

def build_corp_dedicated_name(parent_offering, sr_or_im, app, schedule_suffix, receiver, delivering_tag):
    """Build name for CORP Dedicated Services offerings"""
    parent_content = extract_parent_info(parent_offering)
    catalog_name = extract_catalog_name(parent_offering)
    
    # Extract parts from parent content
    parts = parent_content.split()
    country = ""
    topic = ""
    
    for part in parts:
        if len(part) == 2 and part.isupper() and part not in ["HS", "DS"]:
            country = part
        elif part not in ["HS", "DS", "Parent", "RecP"] and not (len(part) == 2 and part.isupper()):
            topic = part
            break
    
    # Build CORP Dedicated Services name
    prefix_parts = [sr_or_im]
    
    # Add delivering tag parts (who delivers the service - from user input)
    delivering_parts = delivering_tag.split() if delivering_tag else ["HS", country]
    prefix_parts.extend(delivering_parts)
    
    prefix_parts.append("CORP")
    
    # Add receiver (e.g., HS DE)
    prefix_parts.append(receiver)
    prefix_parts.append("Dedicated Services")
    
    # Build the name
    name_prefix = f"[{' '.join(prefix_parts)}]"
    
    name_parts = [name_prefix, catalog_name]
    
    # Add app if provided
    if app:
        name_parts.append(app)
    
    # Add "solving" for IM
    if sr_or_im == "IM":
        name_parts.append("solving")
    
    name_parts.append("Prod")
    name_parts.append(schedule_suffix)
    
    return " ".join(name_parts)

def build_recp_name(parent_offering, sr_or_im, app, schedule_suffix, receiver, delivering_tag):
    """Build name for RecP offerings"""
    parent_content = extract_parent_info(parent_offering)
    catalog_name = extract_catalog_name(parent_offering)
    
    # Extract parts from parent content
    parts = parent_content.split()
    country = ""
    topic = ""
    
    for part in parts:
        if len(part) == 2 and part.isupper() and part not in ["HS", "DS"]:
            country = part
        elif part not in ["HS", "DS", "Parent", "RecP"] and not (len(part) == 2 and part.isupper()):
            topic = part
            break
    
    # Build RecP name - always ends with IT
    prefix_parts = [sr_or_im]
    
    # Extract division and country from parent
    parent_division = ""
    for part in parts:
        if part in ["HS", "DS"]:
            parent_division = part
            break
    
    if parent_division:
        prefix_parts.append(parent_division)
    if country:
        prefix_parts.append(country)
    
    prefix_parts.append("CORP")
    
    # Add delivering tag parts
    delivering_parts = delivering_tag.split() if delivering_tag else ["HS", country]
    prefix_parts.extend(delivering_parts)
    prefix_parts.append("IT")
    
    # Build the name with topic from parent
    name_prefix = f"[{' '.join(prefix_parts)}]"
    
    # For RecP, use topic (e.g., "Software") + catalog name in lowercase + "solving" for IM
    name_parts = [name_prefix]
    
    if topic:
        name_parts.append(topic)
    
    name_parts.append(catalog_name.lower())
    
    if sr_or_im == "IM":
        name_parts.append("solving")
    
    if app:
        name_parts.append(app)
    
    name_parts.append("Prod")
    name_parts.append(schedule_suffix)
    
    return " ".join(name_parts)

def build_standard_name(parent_offering, sr_or_im, app, schedule_suffix, special_dept=None):
    """Build standard name when not CORP"""
    parent_content = extract_parent_info(parent_offering)
    catalog_name = extract_catalog_name(parent_offering)
    
    # Check if catalog name, parent offering, or parent content contains keywords that exclude "Prod"
    no_prod_keywords = ["hardware", "mailbox", "network", "mobile"]
    parent_lower = parent_offering.lower()
    catalog_lower = catalog_name.lower()
    parent_content_lower = parent_content.lower()
    exclude_prod = any(keyword in parent_lower or keyword in catalog_lower or keyword in parent_content_lower for keyword in no_prod_keywords)
    
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
            else:
                # Any other word is the topic (e.g., "Hardware", "Software", "Permissions", etc.)
                if part not in ["HS", "DS", "RecP"] and not (len(part) == 2 and part.isupper()):
                    topic = part
                    break
        
        # If no topic found, use the first significant word from catalog name
        if not topic:
            catalog_words = catalog_name.split()
            for word in catalog_words:
                if word.lower() not in ["the", "a", "an", "and", "or", "for", "of", "in", "on", "to"]:
                    topic = word
                    break
        
        prefix_parts = [sr_or_im]
        if division:
            prefix_parts.append(division)
        if country:
            prefix_parts.append(country)
        prefix_parts.append("IT")
        
        # For IT, build the name with topic BEFORE the brackets
        # Format: [SR DS MD IT] Hardware configuration laptop Mon-Fri 8-16
        name_parts = []
        
        if topic:
            # Topic goes BEFORE the brackets
            name_parts.append(f"[{' '.join(prefix_parts)}] {topic}")
        else:
            name_parts.append(f"[{' '.join(prefix_parts)}]")
        
        name_parts.append(catalog_name.lower())
        
        # Add app if provided
        if app:
            name_parts.append(app)
        
        # Add "solving" for IM - BEFORE checking for Prod
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
    
    # Add delivering tag parts (who delivers the service - from user input)
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

def update_commitments(orig, sched, rsp, rsl, sr_or_im, country):
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
            # Update schedule and duration - OLA uses same pattern as RSL
            line = re.sub(r"RSL\s+[^P]+", f"RSL {sched} ", line)
            line = re.sub(r"(P\d+-P\d+)\s+.*$", f"{p_values} {rsl}", line)
        out.append(line)
    
    # For IM, never add OLA
    # For SR, only add OLA if not PL (HS PL and DS PL already have OLA)
    if sr_or_im == "SR" and not has_ola and country_code and country != "PL":
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

def commit_block(cc, schedule_suffix, rsp_duration, rsl_duration, sr_or_im):
    """Create commitment block with OLA for all countries"""
    if sr_or_im == "IM":
        # For IM, no OLA
        lines=[
            f"[{cc}] SLA IM RSP {schedule_suffix} P1-P4 {rsp_duration}",
            f"[{cc}] SLA IM RSL {schedule_suffix} P1-P4 {rsl_duration}"
        ]
    elif cc == "PL":
        # For SR and PL, include OLA (PL already has OLA in source data)
        lines=[
            f"[{cc}] SLA SR RSP {schedule_suffix} P1-P4 {rsp_duration}",
            f"[{cc}] SLA SR RSL {schedule_suffix} P1-P4 {rsl_duration}",
            f"[{cc}] OLA SR RSL {schedule_suffix} P1-P4 {rsl_duration}"
        ]
    else:
        # For SR and other countries, include OLA
        lines=[
            f"[{cc}] SLA SR RSP {schedule_suffix} P1-P4 {rsp_duration}",
            f"[{cc}] SLA SR RSL {schedule_suffix} P1-P4 {rsl_duration}",
            f"[{cc}] OLA SR RSL {schedule_suffix} P1-P4 {rsl_duration}"
        ]
    return "\n".join(lines)

def custom_commit_block(cc, sr_or_im, rsp_enabled, rsl_enabled, rsp_schedule, rsl_schedule, 
                       rsp_priority, rsl_priority, rsp_time, rsl_time):
    """Create custom commitment block based on user selections"""
    lines = []
    
    if rsp_enabled and rsp_schedule and rsp_priority and rsp_time:
        lines.append(f"[{cc}] SLA {sr_or_im} RSP {rsp_schedule} {rsp_priority} {rsp_time}")
    
    if rsl_enabled and rsl_schedule and rsl_priority and rsl_time:
        lines.append(f"[{cc}] SLA {sr_or_im} RSL {rsl_schedule} {rsl_priority} {rsl_time}")
        # Add OLA for SR only
        if sr_or_im == "SR":
            lines.append(f"[{cc}] OLA {sr_or_im} RSL {rsl_schedule} {rsl_priority} {rsl_time}")
    
    return "\n".join(lines) if lines else ""

def run_generator(*,
    keywords_parent, keywords_child, new_apps, schedule_suffixes,
    delivery_manager, global_prod,
    rsp_duration, rsl_duration,
    sr_or_im, require_corp, require_recp, delivering_tag,
    support_group, managed_by_group, aliases_on, aliases_value,
    src_dir: Path, out_dir: Path,
    special_it=False, special_hr=False, special_medical=False, special_dak=False,
    use_custom_commitments=False, custom_commitments_str="", commitment_country=None,
    rsp_enabled=False, rsl_enabled=False,
    rsp_schedule="", rsl_schedule="",
    rsp_priority="", rsl_priority="",
    rsp_time="", rsl_time="",
    require_corp_it=False, require_corp_dedicated=False):

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
    if special_it and not require_corp and not require_recp and not require_corp_it and not require_corp_dedicated:
        special_dept = "IT"
    elif special_hr and not require_corp and not require_recp and not require_corp_it and not require_corp_dedicated:
        special_dept = "HR"
    elif special_medical and not require_corp and not require_recp and not require_corp_it and not require_corp_dedicated:
        special_dept = "Medical"
    elif special_dak and not require_corp and not require_recp and not require_corp_it and not require_corp_dedicated:
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
        
        # Ensure Visibility group column exists
        if "Visibility group" not in df.columns:
            df["Visibility group"] = ""

        mask=(df.apply(row_keywords_ok,axis=1)
              & df["Name (Child Service Offering lvl 1)"].astype(str).apply(name_prefix_ok)
              & df.apply(lc_ok,axis=1)
              & (df["Service Commitments"].astype(str).str.strip().replace({"nan":""})!="-"))
        
        if require_corp:
            mask &= df["Name (Child Service Offering lvl 1)"].str.contains(r"\bCORP\b",case=False)
            mask &= df["Name (Child Service Offering lvl 1)"].str.contains(rf"\b{re.escape(delivering_tag)}\b",case=False)
            # Exclude IT entries when looking for regular CORP
            mask &= ~df["Parent Offering"].str.contains(r"\bIT\b",case=False)
            # Exclude Dedicated Services entries
            mask &= ~df["Parent Offering"].str.contains(r"\bDedicated\b",case=False)
        elif require_recp:
            # For RecP, filter parent offering that contains RecP
            mask &= df["Parent Offering"].str.contains(r"\bRecP\b",case=False)
        elif require_corp_it:
            # For CORP IT, we don't require CORP to already be in the name
            # We're creating new CORP IT entries from regular entries
            # Just apply the standard filtering without additional CORP requirements
            pass  # No additional filtering needed
        elif require_corp_dedicated:
            # For CORP Dedicated Services, look for Dedicated in parent offering
            mask &= df["Parent Offering"].str.contains(r"\bDedicated\b",case=False)
            # Also ensure it's CORP related
            mask &= (df["Name (Child Service Offering lvl 1)"].str.contains(r"\bCORP\b",case=False) | 
                     df["Parent Offering"].str.contains(r"\bCORP\b",case=False))
        else:
            # For standard processing, exclude CORP entries
            # But don't apply this exclusion for special departments
            if not (special_dept in ["IT", "HR", "Medical", "DAK"]):
                mask &= ~df["Name (Child Service Offering lvl 1)"].str.contains(r"\bCORP\b",case=False)
            
            # Also exclude RecP from standard processing
            mask &= ~df["Parent Offering"].str.contains(r"\bRecP\b",case=False)
            
            # For IT department, filter by IT in either parent offering OR child name
            if special_dept == "IT":
                mask &= (df["Parent Offering"].str.contains(r"\bIT\b",case=False) | 
                         df["Name (Child Service Offering lvl 1)"].str.contains(r"\bIT\b",case=False))
            # For HR department, filter by HR in parent offering
            elif special_dept == "HR":
                mask &= df["Parent Offering"].str.contains(r"\bHR\b",case=False)
            # For Medical department, filter by Medical in parent offering
            elif special_dept == "Medical":
                mask &= df["Parent Offering"].str.contains(r"\bMedical\b",case=False)
            # For DAK department, filter by DAK in parent offering
            elif special_dept == "DAK":
                mask &= df["Parent Offering"].str.contains(r"\bDAK\b",case=False)

        base_pool=df.loc[mask]
        if base_pool.empty:
            continue
        
        # Process ALL matching rows, not just the first one
        for idx, base_row in base_pool.iterrows():
            base_row_df = base_row.to_frame().T.copy()
            country=wb.stem.split("_")[-1].upper()
            tag_hs, tag_ds = f"HS {country}", f"DS {country}"

            # Determine receivers based on country and CORP/RecP/CORP IT/CORP Dedicated setting
            if require_corp or require_recp or require_corp_it or require_corp_dedicated:
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
            
            # Store original depend on value
            original_depend_on = str(base_row.get("Service Offerings | Depend On (Application Service)", "")).strip()

            for app in all_apps:
                for schedule_suffix in schedule_suffixes:
                    for recv in receivers:
                    # For DE, find the matching row (DS DE or HS DE) in the original data
                        if country == "DE" and (require_corp or require_recp or require_corp_it or require_corp_dedicated):
                        # Search for the specific receiver in the base pool
                            recv_mask = base_pool["Name (Child Service Offering lvl 1)"].str.contains(rf"\b{re.escape(recv)}\b", case=False)
                            matching_rows = base_pool[recv_mask]
                        
                            if not matching_rows.empty:
                            # Use the first matching row as base
                                base_row = matching_rows.iloc[0]
                                base_row_df = base_row.to_frame().T.copy()
                                original_depend_on = str(base_row.get("Service Offerings | Depend On (Application Service)", "")).strip()
                
                        if require_corp:
                            new_name = build_corp_name(
                                parent_full, sr_or_im, app, schedule_suffix, recv, delivering_tag
                            )
                        elif require_recp:
                            new_name = build_recp_name(
                                parent_full, sr_or_im, app, schedule_suffix, recv, delivering_tag
                            )
                        elif require_corp_it:
                            new_name = build_corp_it_name(
                                parent_full, sr_or_im, app, schedule_suffix, recv, delivering_tag
                            )
                        elif require_corp_dedicated:
                            new_name = build_corp_dedicated_name(
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
                        
                        # Handle aliases - copy from original row if aliases_on is False
                        if not aliases_on:
                            # Keep original aliases values
                            pass
                        else:
                            # Apply new alias value
                            for c in [c for c in row.columns if "Aliases" in c]:
                                row[c]=aliases_value if aliases_value else "-"
                        
                        # Handle Visibility group - ensure it exists for PL
                        if country == "PL" and "Visibility group" not in row.columns:
                            row["Visibility group"] = ""
                        
                        if country=="DE":
                            row["Subscribed by Company"]="DE Internal Patients\nDE External Patients" if recv=="HS DE" else "DE IFLB Laboratories\nDE IMD Laboratories"
                        elif country=="UA":
                            row["Subscribed by Company"]="Сiнево Україна"
                        elif country=="CY":
                            row["Subscribed by Company"]="CY Healthcare Services\nCY Medical Centers" if recv=="HS CY" else "CY Diagnostic Laboratories"
                        else:
                            row["Subscribed by Company"]=recv or tag_hs
                            
                        orig_comm=str(row.iloc[0]["Service Commitments"]).strip()
                        
                        # Handle custom commitments - use both approaches
                        if use_custom_commitments and custom_commitments_str:
                           # Use the direct string if provided
                           row["Service Commitments"] = custom_commitments_str
                        elif use_custom_commitments and commitment_country:
                            # Use custom_commit_block function if country provided
                            row["Service Commitments"] = custom_commit_block(
                                commitment_country, sr_or_im, rsp_enabled, rsl_enabled,
                                rsp_schedule, rsl_schedule, rsp_priority, rsl_priority,
                                rsp_time, rsl_time
                            )
                        else:
                           # Use existing logic
                           row["Service Commitments"]=commit_block(country, schedule_suffix, rsp_duration, rsl_duration, sr_or_im) if not orig_comm or orig_comm=="-" else update_commitments(orig_comm,schedule_suffix,rsp_duration,rsl_duration,sr_or_im,country)
                        
                        # Special handling for IT with UA/MD - always use DS
                        if (special_dept == "IT" or require_corp_it) and country in ["UA", "MD"]:
                            depend_tag = f"DS {country} Prod"
                        elif global_prod:
                            depend_tag = "Global Prod"
                        else:
                            depend_tag = f"{delivering_tag} Prod" if (require_corp or require_recp or require_corp_it or require_corp_dedicated) else f"{recv or tag_hs} Prod"
                        
                        # Handle Service Offerings | Depend On - preserve original values
                        if original_depend_on in ["-", "", "nan", "NaN", None]:
                            # Preserve the original empty/dash value
                            if original_depend_on == "-":
                                row["Service Offerings | Depend On (Application Service)"] = "-"
                            else:
                                row["Service Offerings | Depend On (Application Service)"] = ""
                        else:
                            # Only update if original had a real value
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
        # Sort country codes alphabetically
        sorted_countries = sorted(sheets.keys())
        
        for cc in sorted_countries:
            dfc = sheets[cc]
            # Ensure unique names per country
            df_final = dfc.drop_duplicates(subset=["Name (Child Service Offering lvl 1)"])
            
            if "Number" in df_final.columns:
                df_final = df_final.drop(columns=["Number"])
            # Remove Visibility group column for PL
            if cc == "PL" and "Visibility group" in df_final.columns:
               df_final = df_final.drop(columns=["Visibility group"])
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
                    # normalize any True/False (any case) to lowercase
                    df_final[col] = df_final[col].replace({
                       'True': 'true',  'TRUE': 'true',
                    'False': 'false','FALSE': 'false'
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