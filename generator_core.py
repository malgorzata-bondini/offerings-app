import datetime as dt
import re
import time
import warnings
from pathlib import Path
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Alignment, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.comments import Comment
from dataclasses import dataclass
from typing import List, Dict, Optional, Any

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

need_cols = [
    "Name (Child Service Offering lvl 1)", "Parent Offering",
    "Service Offerings | Depend On (Application Service)", "Service Commitments",
    "Delivery Manager", "Subscribed by Location", "Phase", "Status",
    "Life Cycle Stage", "Life Cycle Status", "Support group", "Managed by Group",
    "Subscribed by Company", "Visibility group", "Business Criticality",
    "Record view", "Approval required", "Approval group"  # Add "Approval group"
]

discard_lc = {"retired", "retiring", "end of life", "end of support"}

def ensure_incident_naming(name):
    """
    Ensure that if any keyword is 'incident', then 'solving' is right after 'incident' and then app name
    This function reorganizes the name to ensure proper order: incident solving [app] [other parts]
    """
    # First, check if we already have "incident solving" in the correct order
    if re.search(r'\bincident\s+solving\b', name, re.IGNORECASE):
        return name
    
    # Split the name into parts
    parts = name.split()
    new_parts = []
    i = 0
    
    while i < len(parts):
        if parts[i].lower() == "incident":
            # Add "incident solving"
            new_parts.append(parts[i])
            new_parts.append("solving")
            
            # Skip if the next word was already "solving"
            if i + 1 < len(parts) and parts[i + 1].lower() == "solving":
                i += 1
        elif parts[i].lower() == "solving":
            # Skip standalone "solving" as we'll add it after "incident"
            pass
        else:
            new_parts.append(parts[i])
        i += 1
    
    # Now we need to ensure app name comes AFTER "incident solving", not between
    # Find if there's an app name that got placed between incident and solving
    final_parts = []
    i = 0
    
    while i < len(new_parts):
        if i > 0 and new_parts[i-1].lower() == "incident" and new_parts[i].lower() == "solving":
            # We found "incident solving" - good
            final_parts.append(new_parts[i])
        elif new_parts[i].lower() == "incident":
            # Check if there's something between incident and solving
            j = i + 1
            app_parts = []
            
            # Collect everything until we find "solving"
            while j < len(new_parts) and new_parts[j].lower() != "solving":
                app_parts.append(new_parts[j])
                j += 1
            
            # Add incident solving first
            final_parts.append(new_parts[i])
            if j < len(new_parts) and new_parts[j].lower() == "solving":
                final_parts.append("solving")
                # Then add the app parts that were in between
                final_parts.extend(app_parts)
                i = j
            else:
                # No solving found, just add solving after incident
                final_parts.append("solving")
                final_parts.extend(app_parts)
                i = j - 1
        else:
            final_parts.append(new_parts[i])
        i += 1
    
    return " ".join(final_parts)

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

def get_division_and_country(parent_content, country, delivering_tag):
    """Get division and country with special handling for MD, UA, RO, and TR"""
    # Special handling for UA, MD, RO, and TR - always use DS
    if country in ["UA", "MD", "RO", "TR"]:
        return "DS", country
    
    # Extract parts from parent content
    parts = parent_content.split()
    division = ""
    
    for part in parts:
        if part in ["HS", "DS"]:
            division = part
            break
    
    # If no division found in parent and we have delivering_tag, use it
    if not division and delivering_tag:
        delivering_parts = delivering_tag.split()
        if delivering_parts and delivering_parts[0] in ["HS", "DS"]:
            division = delivering_parts[0]
    
    # Default to HS if still no division
    if not division:
        division = "HS"
    
    return division, country

def build_lvl2_name(parent_offering, sr_or_im, app, schedule_suffix, service_type_lvl2):
    """Build name for Lvl2 entries - SR/IM is added to both Parent Offering parsing and final name"""
    parent_content = extract_parent_info(parent_offering)
    catalog_name = extract_catalog_name(parent_offering)
    
    # Check if SR/IM already exists in parent content
    parts = parent_content.split()
    has_sr_im = "SR" in parts or "IM" in parts
    
    # If SR/IM not already there, insert it after Parent
    if not has_sr_im:
        # Insert SR/IM as the first element after "Parent" was removed
        parts.insert(0, sr_or_im)
        parent_content = " ".join(parts)
    
    # Now extract parts for building
    parts = parent_content.split()
    country = ""
    division = ""
    dept = ""  # IT, HR, etc.
    sr_im_pos = ""
    
    for part in parts:
        if part in ["SR", "IM"]:
            sr_im_pos = part
        elif len(part) == 2 and part.isupper() and part not in ["HS", "DS", "IT", "HR"]:
            country = part
        elif part in ["HS", "DS"]:
            division = part
        elif part in ["IT", "HR", "Medical", "Business Services"]:
            dept = part
    
    # Build the prefix - use the SR/IM from parent or the one provided
    prefix_parts = [sr_im_pos if sr_im_pos else sr_or_im]
    
    # Special handling for UA, MD, RO, and TR - always use DS
    if country in ["UA", "MD", "RO", "TR"]:
        prefix_parts.append("DS")
    elif division:
        prefix_parts.append(division)
    else:
        prefix_parts.append("HS")  # Default
    
    if country:
        prefix_parts.append(country)
    
    if dept:
        prefix_parts.append(dept)
    
    # Build the name
    name_prefix = f"[{' '.join(prefix_parts)}]"
    
    # Extract the core part of catalog name (e.g., "Software incident solving")
    name_parts = [name_prefix, catalog_name]
    
    # Add app if provided
    if app:
        name_parts.append(app)
    
    # Check if name contains Microsoft - if so, don't add Prod
    name_so_far = " ".join(name_parts).lower()
    if "microsoft" not in name_so_far:
        name_parts.append("Prod")
    
    # Add service type if provided (e.g., "Application issue")
    if service_type_lvl2:
        name_parts.append(service_type_lvl2)
    
    # Add schedule
    name_parts.append(schedule_suffix)
    
    # Join and ensure incident naming
    final_name = " ".join(name_parts)
    return ensure_incident_naming(final_name)

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
    
    # Get division and country with special handling for MD/UA/RO/TR
    division, country_code = get_division_and_country(parent_content, country, delivering_tag)
    
    # Add delivering tag parts (who delivers the service - from user input)
    if delivering_tag:
        delivering_parts = delivering_tag.split()
        # For MD/UA/RO/TR, override with DS
        if country in ["UA", "MD", "RO", "TR"]:
            prefix_parts.extend(["DS", country])
        else:
            prefix_parts.extend(delivering_parts)
    else:
        prefix_parts.extend([division, country])
    
    prefix_parts.append("CORP")
    
    # Add receiver parts (who receives the service - DS DE or HS DE for Germany)
    if receiver:
        prefix_parts.append(receiver)
    else:
        prefix_parts.extend([division, country])
    
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
    
    # name_parts.append("Prod")
    name_parts.append(schedule_suffix)
    
    # Join and ensure incident naming
    final_name = " ".join(name_parts)
    return ensure_incident_naming(final_name)

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
    
    # Get division and country with special handling for MD/UA/RO/TR
    division, country_code = get_division_and_country(parent_content, country, delivering_tag)
    
    # Add delivering tag parts (who delivers the service - from user input)
    if delivering_tag:
        delivering_parts = delivering_tag.split()
        # For MD/UA/RO/TR, override with DS
        if country in ["UA", "MD", "RO", "TR"]:
            prefix_parts.extend(["DS", country])
        else:
            prefix_parts.extend(delivering_parts)
    else:
        prefix_parts.extend([division, country])
    
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
    
    # Join and ensure incident naming
    final_name = " ".join(name_parts)
    return ensure_incident_naming(final_name)

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
    
    # Get division and country with special handling for MD/UA/RO/TR
    division, country_code = get_division_and_country(parent_content, country, delivering_tag)
    
    # Extract division and country from parent
    parent_division = ""
    for part in parts:
        if part in ["HS", "DS"]:
            parent_division = part
            break
    
    # For MD/UA/RO/TR, always use DS
    if country in ["UA", "MD", "RO", "TR"]:
        prefix_parts.extend(["DS", country])
    else:
        if parent_division:
            prefix_parts.append(parent_division)
        if country:
            prefix_parts.append(country)
    
    prefix_parts.append("CORP")
    
    # Add delivering tag parts
    if delivering_tag:
        delivering_parts = delivering_tag.split()
        # For MD/UA/RO/TR, override with DS
        if country in ["UA", "MD", "RO", "TR"]:
            prefix_parts.extend(["DS", country])
        else:
            prefix_parts.extend(delivering_parts)
    else:
        # For MD/UA/RO/TR, use DS
        if country in ["UA", "MD", "RO", "TR"]:
            prefix_parts.extend(["DS", country])
        else:
            prefix_parts.extend([division, country])
    
    prefix_parts.append("IT")
    
    # Build the name with topic from parent
    name_prefix = f"[{' '.join(prefix_parts)}]"
    
    # For RecP, use topic (e.g., "Software") + catalog name in lowercase
    name_parts = [name_prefix]
    
    if topic:
        name_parts.append(topic)
    
    name_parts.append(catalog_name.lower())
    
    # Add app if provided
    if app:
        name_parts.append(app)
    
    # Add "solving" for IM
    if sr_or_im == "IM":
        name_parts.append("solving")
    
    name_parts.append("Prod")
    name_parts.append(schedule_suffix)
    
    # Join and ensure incident naming
    final_name = " ".join(name_parts)
    return ensure_incident_naming(final_name)

def build_standard_name(parent_offering, sr_or_im, app, schedule_suffix, special_dept=None, receiver=None):
    """Build standard name when not CORP"""
    parent_content = extract_parent_info(parent_offering)
    catalog_name = extract_catalog_name(parent_offering)
    
    # Extract country from parent content
    parts = parent_content.split()
    country = ""
    for part in parts:
        if len(part) == 2 and part.isupper() and part not in ["HS", "DS", "IT", "HR"]:
            country = part
            break
    
    # Check if catalog name, parent offering, or parent content contains keywords that exclude "Prod"
    no_prod_keywords = ["hardware", "mailbox", "network", "mobile", "security"]
    parent_lower = parent_offering.lower()
    catalog_lower = catalog_name.lower()
    parent_content_lower = parent_content.lower()
    exclude_prod = any(keyword in parent_lower or keyword in catalog_lower or keyword in parent_content_lower 
                      for keyword in no_prod_keywords)
    
    if special_dept == "Medical":
        # Extract division and country from parent content
        parts = parent_content.split()
        division = ""
        country = ""
        topic_parts = []
        
        for i, part in enumerate(parts):
            if part in ["HS", "DS"]:
                division = part
            elif len(part) == 2 and part.isupper() and part not in ["IT", "HR"]:
                country = part
            elif part not in ["HS", "DS"] and not (len(part) == 2 and part.isupper()):
                topic_parts.append(part)
        
        topic = " ".join(topic_parts) if topic_parts else "Software"
        
        # Build Medical name - NO PROD
        prefix_parts = [sr_or_im]
        
        # Special handling for UA, MD, RO, and TR - always use DS
        if country in ["UA", "MD", "RO", "TR"]:
            prefix_parts.extend(["DS", country])
        else:
            if division:
                prefix_parts.append(division)
            if country:
                prefix_parts.append(country)
        
        prefix_parts.append("Medical")
        
        # Use topic from parent and lowercase catalog name
        final_name = f"[{' '.join(prefix_parts)}] {topic} {catalog_name.lower()} {schedule_suffix}"
        return ensure_incident_naming(final_name)
    
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
        
        # Special handling for UA, MD, RO, and TR - always use DS
        if country in ["UA", "MD", "RO", "TR"]:
            prefix_parts.extend(["DS", country])
        else:
            if division:
                prefix_parts.append(division)
            if country:
                prefix_parts.append(country)
        
        prefix_parts.append("Business Services")
        
        # Add app only if provided - NO PROD
        if app:
            final_name = f"[{' '.join(prefix_parts)}] {catalog_name} {app} {schedule_suffix}"
        else:
            final_name = f"[{' '.join(prefix_parts)}] {catalog_name} {schedule_suffix}"
        return ensure_incident_naming(final_name)
    
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
        
        # Special handling for UA, MD, RO, and TR - always use DS
        if country in ["UA", "MD", "RO", "TR"]:
            prefix_parts.extend(["DS", country])
        else:
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
            final_name = f"[{' '.join(prefix_parts)}] {topic} {catalog_name.lower()} {app} {schedule_suffix}"
        else:
            final_name = f"[{' '.join(prefix_parts)}] {topic} {catalog_name.lower()} {schedule_suffix}"
        return ensure_incident_naming(final_name)
    
    elif special_dept == "IT":
        # IT - special handling
        parts = parent_content.split()
        division = ""
        country = ""
        topic = ""
        
        # Find division, country and topic
        topic_parts = []
        for i, part in enumerate(parts):
            if part in ["HS", "DS"]:
                division = part
            elif len(part) == 2 and part.isupper() and part not in ["IT", "HR"]:
                country = part
            else:
                # Collect all remaining words as topic (e.g., "Security & Privacy", "Hardware", etc.)
                if part not in ["HS", "DS", "RecP"] and not (len(part) == 2 and part.isupper()):
                    topic_parts.append(part)
        
        # Join all topic parts to get the full topic phrase
        topic = " ".join(topic_parts) if topic_parts else ""
        
        # If no topic found, use the first significant word from catalog name
        if not topic:
            catalog_words = catalog_name.split()
            for word in catalog_words:
                if word.lower() not in ["the", "a", "an", "and", "or", "for", "of", "in", "on", "to"]:
                    topic = word
                    break
        
        prefix_parts = [sr_or_im]
        
        # For DE with receiver, extract division from receiver (e.g., "HS DE" -> "HS")
        if country == "DE" and receiver:
            recv_division = receiver.split()[0]  # Extract HS or DS
            prefix_parts.append(recv_division)
        # Special handling for UA, MD, RO, and TR - always use DS
        elif country in ["UA", "MD", "RO", "TR"]:
            prefix_parts.append("DS")
        elif division:
            prefix_parts.append(division)
        else:
            # Default to HS if no division found
            prefix_parts.append("HS")
        
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
        
        # Add "solving" for IM
        if sr_or_im == "IM":
            name_parts.append("solving")
        
        # Check if topic contains any no-prod keywords
        topic_lower = topic.lower() if topic else ""
        topic_exclude_prod = any(keyword in topic_lower for keyword in no_prod_keywords)
        
        # Only add Prod if no hardware/mailbox/network/mobile/security keywords in any source
        if not exclude_prod and not topic_exclude_prod:
            name_parts.append("Prod")
        
        name_parts.append(schedule_suffix)
        
        # Join and ensure incident naming
        final_name = " ".join(name_parts)
        return ensure_incident_naming(final_name)
    
    else:
        # Standard case - replace Parent with SR/IM and add IT for RecP entries
        parts = parent_content.split()
        division = ""
        country_code = ""
        dept = ""
        other_parts = []
        
        # Parse parent content to extract components
        for part in parts:
            if part in ["HS", "DS"]:
                division = part
            elif len(part) == 2 and part.isupper() and part not in ["HS", "DS", "IT", "HR"]:
                country_code = part
            elif part in ["IT", "HR", "Medical", "Business Services"]:
                dept = part
            else:
                other_parts.append(part)
        
        # Build the new name components
        name_parts = [sr_or_im]
        
        # Special handling for UA, MD, RO, and TR - always use DS
        if country in ["UA", "MD", "RO", "TR"]:
            name_parts.append("DS")
        elif division:
            name_parts.append(division)
        else:
            # Default to HS if no division found
            name_parts.append("HS")
        
        if country_code:
            name_parts.append(country_code)
        
        # Add other parts (like RecP, Software, etc.)
        name_parts.extend(other_parts)
        
        # For RecP entries, add IT department if not already present
        if "RecP" in other_parts and not dept:
            name_parts.append("IT")
        elif dept:
            name_parts.append(dept)
        
        # Build the final name
        prefix = f"[{' '.join(name_parts)}]"
        
        # Add catalog name
        final_parts = [prefix, catalog_name]
        
        # Check if we should add Prod
        catalog_lower = catalog_name.lower()
        exclude_prod = any(keyword in catalog_lower for keyword in ["hardware", "mailbox", "network", "mobile", "security"])
        
        # Add app if provided
        if app:
            final_parts.append(app)
        
        # Add "solving" for IM
        if sr_or_im == "IM":
            final_parts.append("solving")
        
        if not exclude_prod:
            final_parts.append("Prod")
        
        final_parts.append(schedule_suffix)
        
        final_name = " ".join(final_parts)
        return ensure_incident_naming(final_name)

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
    
    # Get division and country with special handling for MD/UA/RO/TR
    division, country_code = get_division_and_country(parent_content, country, delivering_tag)
    
    # Add delivering tag parts (who delivers the service - from user input)
    if delivering_tag:
        delivering_parts = delivering_tag.split()
        # For MD/UA/RO/TR, override with DS
        if country in ["UA", "MD", "RO", "TR"]:
            prefix_parts.extend(["DS", country])
        else:
            prefix_parts.extend(delivering_parts[:2])  # Take division and country from delivering tag
    else:
        prefix_parts.extend([division, country])
    
    prefix_parts.extend(["CORP", receiver])
    
    # Only add topic if it exists, don't default to IT for CORP
    if topic:
        prefix_parts.append(topic)
    
    if sr_or_im == "SR":
        if app:
            final_name = f"[{' '.join(prefix_parts)}] {catalog_name} {app} Prod {schedule_suffix}"
        else:
            final_name = f"[{' '.join(prefix_parts)}] {catalog_name} Prod {schedule_suffix}"
    else:
        # For IM: always add solving after the app (or after catalog if no app)
        if app:
            final_name = f"[{' '.join(prefix_parts)}] {catalog_name} {app} solving Prod {schedule_suffix}"
        else:
            final_name = f"[{' '.join(prefix_parts)}] {catalog_name} solving Prod {schedule_suffix}"
    
    return ensure_incident_naming(final_name)

def commit_block(cc, schedule_suffix, rsp_duration, rsl_duration, sr_or_im):
    """Create commitment block with OLA for all countries"""
    if sr_or_im == "IM":
        # For IM, no OLA
        lines = [
            f"[{cc}] SLA IM RSP {schedule_suffix} P1-P4 {rsp_duration}",
            f"[{cc}] SLA IM RSL {schedule_suffix} P1-P4 {rsl_duration}"
        ]
    else:
        # For SR, include OLA (but only once)
        lines = [
            f"[{cc}] SLA SR RSP {schedule_suffix} P1-P4 {rsp_duration}",
            f"[{cc}] SLA SR RSL {schedule_suffix} P1-P4 {rsl_duration}",
            f"[{cc}] OLA SR RSL {schedule_suffix} P1-P4 {rsl_duration}"
        ]
    return "\n".join(lines)

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
    
    return "\n".join(out)

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

def create_new_parent_row(new_parent_offering, new_parent, country, business_criticality="", approval_required=False, approval_required_value="empty", change_subscribed_location=False, custom_subscribed_location="Global"):
    """Create a new row with the specified parent offering and parent values"""
    # Create a basic row structure with required columns
    new_row = {
        "Name (Child Service Offering lvl 1)": "",  # Will be filled later
        "Parent Offering": new_parent_offering,  # Use the user-provided value
        "Parent": new_parent,  # Use the user-provided value (not hardcoded)
        "Service Offerings | Depend On (Application Service)": "",
        "Service Commitments": "",
        "Delivery Manager": "",
        "Subscribed by Location": "",  # Will be set below
        "Phase": "Catalog",
        "Status": "Operational",
        "Life Cycle Stage": "Operational",
        "Life Cycle Status": "In Use",
        "Support group": "",
        "Managed by Group": "",
        "Subscribed by Company": "",  # Will be set during processing based on receiver and CORP type
        "Visibility group": "",
        "Business Criticality": business_criticality,
        "Record view": "",  # Will be set based on SR/IM
        "Approval required": "false" if not approval_required else approval_required_value,
        "Approval group": "empty" if not approval_required else approval_required_value
    }
    
    # Set Subscribed by Location based on user choice
    if change_subscribed_location:
        new_row["Subscribed by Location"] = custom_subscribed_location
    else:
        new_row["Subscribed by Location"] = "Global"
    
    return pd.Series(new_row)

def get_support_group_for_country(country, support_group, support_groups_per_country, division=None):
    """Get the appropriate support group for a given country and division"""
    # For PL, use division-specific support groups
    if country == "PL" and division:
        division_key = f"{division} PL"
        if support_groups_per_country and division_key in support_groups_per_country:
            return support_groups_per_country[division_key]
    
    # For other countries, use the country key directly
    if support_groups_per_country and country in support_groups_per_country:
        return support_groups_per_country[country]
    return support_group

def get_managed_by_group_for_country(country, managed_by_group, managed_by_groups_per_country, 
                                     support_group_for_country, division=None):
    """Get the appropriate managed by group for a given country and division"""
    # For PL, use division-specific managed by groups
    if country == "PL" and division:
        division_key = f"{division} PL"
        if managed_by_groups_per_country and division_key in managed_by_groups_per_country:
            managed_by_value = managed_by_groups_per_country[division_key]
            # If managed_by is empty but support_group is filled, use support_group
            return managed_by_value if managed_by_value else support_group_for_country
    
    # For other countries, use the country key directly
    if managed_by_groups_per_country and country in managed_by_groups_per_country:
        managed_by_value = managed_by_groups_per_country[country]
        # If managed_by is empty but support_group is filled, use support_group
        return managed_by_value if managed_by_value else support_group_for_country
    # Fall back to global managed_by_group, or support_group_for_country if empty
    return managed_by_group if managed_by_group else support_group_for_country

def get_support_groups_list_for_country(country, support_group, support_groups_per_country, 
                                       managed_by_groups_per_country, division=None):
    """Get list of support groups for a country (handles multiple groups for DE)"""
    # Get support groups
    if country == "PL" and division:
        division_key = f"{division} PL"
        country_support_groups = support_groups_per_country.get(division_key, support_group) if support_groups_per_country else support_group
        country_managed_groups = managed_by_groups_per_country.get(division_key, "") if managed_by_groups_per_country else ""
    else:
        country_support_groups = support_groups_per_country.get(country, support_group) if support_groups_per_country else support_group
        country_managed_groups = managed_by_groups_per_country.get(country, "") if managed_by_groups_per_country else ""
    
    # Handle multiple groups (separated by newlines)
    if country_support_groups and '\n' in str(country_support_groups):
        support_list = [sg.strip() for sg in str(country_support_groups).split('\n') if sg.strip()]
        managed_list = []
        if country_managed_groups and '\n' in str(country_managed_groups):
            managed_list = [mg.strip() for mg in str(country_managed_groups).split('\n') if mg.strip()]
        else:
            managed_list = [country_managed_groups.strip()] * len(support_list) if country_managed_groups else support_list
        
        # Ensure managed_list has same length as support_list
        while len(managed_list) < len(support_list):
            managed_list.append(support_list[len(managed_list)])
            
        return list(zip(support_list, managed_list))
    else:
        # Single support group
        single_support = country_support_groups or support_group or ""
        single_managed = country_managed_groups or single_support or ""
        return [(single_support, single_managed)] if single_support else [("", "")]

def get_schedule_suffixes_for_country(country, receiver, schedule_settings_per_country, default_schedule_suffixes):
    """Get the appropriate schedule suffixes for a given country and receiver"""
    # For countries that split into DS/HS (PL), use receiver key
    # For DE, it doesn't split in the same way, so use country directly
    # For CY (now DS-only), RO, TR, use country directly
    if country in ["PL"] and receiver:
        key = receiver  # e.g., "HS PL", "DS PL"
    else:
        key = country  # e.g., "DE", "MD", "UA", "CY", "RO", "TR"
    
    # Check if there are custom schedules for this country/receiver
    if key in schedule_settings_per_country:
        custom_schedules = schedule_settings_per_country[key]
        if isinstance(custom_schedules, str):
            # Split on newlines if it's a string
            return [s.strip() for s in custom_schedules.split('\n') if s.strip()]
        elif isinstance(custom_schedules, list):
            return custom_schedules
    
    # Fallback to default schedule suffixes
    return default_schedule_suffixes

def get_de_company_and_ldap(support_group, receiver, original_row=None):
    """Get the Subscribed by Company and LDAP values for DE based on support group"""
    # Normalize the support group name for comparison (remove extra spaces, normalize case)
    normalized_sg = ' '.join(support_group.strip().split()) if support_group else ""
    
    # Special mappings for DE support groups - using normalized comparison
    if "HS DE IT Service Desk HC" in normalized_sg and receiver == "HS DE":
        return "DE Internal Patients", "CALDOM1.DE [Hospital Calbe]"
    elif "HS DE IT Service Desk - MCC" in normalized_sg and receiver == "HS DE":
        return "DE External Patients", "mednet-de.world [Medicover Clinics]"
    elif ("DS DE IT Service Desk -Labs" in normalized_sg or "DS DE IT Service Desk - Labs" in normalized_sg) and receiver == "DS DE":
        return "DE IFLB Laboratories\nDE IMD Laboratories", "imd-labore.intern [General]"
    else:
        # For other support groups, return the original "Subscribed by Company" value if available
        if original_row is not None and "Subscribed by Company" in original_row.index:
            original_company = str(original_row["Subscribed by Company"]).strip()
            if original_company and original_company not in ["nan", "NaN", "", "None"]:
                return original_company, ""
        # Fallback to support group name if no original value available
        return support_group, ""

@dataclass
class GeneratorConfig:
    # Basic parameters
    keywords_parent: str
    keywords_child: str
    new_apps: List[str]
    schedule_suffixes: List[str]
    delivery_manager: str
    global_prod: bool
    
    # Service parameters
    rsp_duration: str
    rsl_duration: str
    sr_or_im: str
    
    # Type flags
    require_corp: bool = False
    require_recp: bool = False
    require_corp_it: bool = False
    require_corp_dedicated: bool = False
    
    # Special department flags
    special_it: bool = False
    special_hr: bool = False
    special_medical: bool = False
    special_dak: bool = False
    
    # Other parameters
    delivering_tag: str = ""
    support_group: str = ""
    managed_by_group: str = ""
    aliases_on: bool = False
    aliases_value: str = ""
    
    # Directories
    src_dir: Optional[Path] = None
    out_dir: Optional[Path] = None
    
    # Custom commitments
    use_custom_commitments: bool = False
    custom_commitments_str: str = ""
    commitment_country: Optional[str] = None
    
    # Per-country settings
    support_groups_per_country: Optional[Dict[str, str]] = None
    managed_by_groups_per_country: Optional[Dict[str, str]] = None
    schedule_settings_per_country: Optional[Dict[str, str]] = None
    aliases_per_country: Optional[Dict[str, str]] = None
    
    # Additional flags
    use_new_parent: bool = False
    new_parent_offering: str = ""
    new_parent: str = ""
    keywords_excluded: str = ""
    use_lvl2: bool = False
    service_type_lvl2: str = ""
    use_custom_depend_on: bool = False
    custom_depend_on_value: str = ""
    business_criticality: str = ""

def run_generator(
    keywords_parent, keywords_child, new_apps, schedule_suffixes,
    delivery_manager, global_prod,
    rsp_duration, rsl_duration,
    sr_or_im, require_corp, require_recp, delivering_tag,
    support_group, managed_by_group, aliases_on, aliases_value,
    src_dir, out_dir,
    special_it=False, special_hr=False, special_medical=False, special_dak=False,
    use_custom_commitments=False, custom_commitments_str="", commitment_country=None,
    rsp_enabled=False, rsl_enabled=False,
    rsp_schedule="", rsl_schedule="",
    rsp_priority="", rsl_priority="",
    rsp_time="", rsl_time="",
    require_corp_it=False, require_corp_dedicated=False,
    use_new_parent=False, new_parent_offering="", new_parent="",
    keywords_excluded="",
    use_lvl2=False, service_type_lvl2="",
    support_groups_per_country=None, managed_by_groups_per_country=None,
    schedule_settings_per_country=None,
    use_custom_depend_on=False, custom_depend_on_value="",
    aliases_per_country=None,
    selected_languages=None,
    business_criticality="",
    approval_required=False,
    approval_required_value="empty",
    change_subscribed_location=False,
    custom_subscribed_location="Global"):
    """
    Main generator function refactored for reduced complexity.
    """

    def initialize_settings():
        """Initialize all per-country settings dictionaries"""
        nonlocal support_groups_per_country, managed_by_groups_per_country, schedule_settings_per_country, aliases_per_country, selected_languages
        if support_groups_per_country is None:
            support_groups_per_country = {}
        if managed_by_groups_per_country is None:
            managed_by_groups_per_country = {}
        if schedule_settings_per_country is None:
            schedule_settings_per_country = {}
        if aliases_per_country is None:
            aliases_per_country = {}
        if selected_languages is None:
            selected_languages = []

    def parse_all_apps(new_apps):
        """Parse and split application names from input"""
        all_apps = []
        for raw in new_apps:
            for app in re.split(r'[,\n;]+', str(raw)):
                app = app.strip()
                if app:
                    all_apps.append(app)
        if not all_apps:
            all_apps = [None]
        return all_apps

    def determine_special_dept(special_it, special_hr, special_medical, special_dak):
        """Determine special department based on flags"""
        if special_it:
            return "IT"
        elif special_hr:
            return "HR"
        elif special_medical:
            return "Medical"
        elif special_dak:
            return "DAK"
        return None

    def collect_existing_offerings(src_dir):
        """Collect all existing offerings from source files"""
        existing_offerings = set()
        original_ldap_data = {}
        excel_cache = {}
        column_order_cache = {}
        
        for wb in src_dir.glob("ALL_Service_Offering_*.xlsx"):
            try:
                if wb not in excel_cache:
                    excel_cache[wb] = pd.ExcelFile(wb)
                excel_file = excel_cache[wb]
                
                for sheet_name in ["Child SO lvl1", "Child SO lvl2"]:
                    try:
                        if sheet_name in excel_file.sheet_names:
                            df = pd.read_excel(excel_file, sheet_name=sheet_name)
                            
                            # Store column order
                            country = wb.stem.split("_")[-1].upper()
                            column_key = f"{country}_{sheet_name}"
                            if column_key not in column_order_cache:
                                column_order_cache[column_key] = list(df.columns)
                            
                            if "Name (Child Service Offering lvl 1)" in df.columns:
                                existing_names = df["Name (Child Service Offering lvl 1)"].dropna().astype(str)
                                normalized_names = existing_names.apply(lambda x: ' '.join(str(x).split()))
                                existing_offerings.update(normalized_names)
                            
                            # Collect LDAP data for DE
                            if wb.stem.endswith("_DE") and "Support group" in df.columns:
                                ldap_cols = [col for col in df.columns if "LDAP" in col.upper() or "Ldap" in col or "ldap" in col]
                                if ldap_cols and "Support group" in df.columns:
                                    for idx, row in df.iterrows():
                                        sg = str(row.get("Support group", "")).strip()
                                        if sg and sg not in ["nan", "NaN", ""]:
                                            ldap_values = {}
                                            for ldap_col in ldap_cols:
                                                ldap_val = str(row.get(ldap_col, "")).strip()
                                                if ldap_val and ldap_val not in ["nan", "NaN", "", "None", "none"]:
                                                    ldap_values[ldap_col] = ldap_val
                                            if ldap_values:
                                                original_ldap_data[sg] = ldap_values
                    except Exception:
                        continue
            except Exception:
                continue
        
        return existing_offerings, original_ldap_data, excel_cache, column_order_cache

    def process_files(src_dir, excel_cache, column_order_cache, existing_offerings, original_ldap_data):
        """Process all source files and generate offerings"""
        sheets_data = {}
        seen = set()
        missing_schedule_info = {}
        
        total_files = len(list(src_dir.glob("ALL_Service_Offering_*.xlsx")))
        processed_files = 0
        
        for wb in src_dir.glob("ALL_Service_Offering_*.xlsx"):
            processed_files += 1
            country = wb.stem.split("_")[-1].upper()
            print(f"Processing file {processed_files}/{total_files}: {wb.name}")
            
            levels_to_process = [1, 2] if use_lvl2 else [1]
            
            for current_level in levels_to_process:
                sheet_name = f"Child SO lvl{current_level}"
                is_lvl2 = (current_level == 2)
                
                try:
                    base_pool = get_base_pool(wb, sheet_name, country, is_lvl2, excel_cache, column_order_cache)
                    if base_pool.empty:
                        continue
                    
                    # Get schedule checking data
                    all_country_names, corp_names, non_corp_names = get_schedule_checking_data(base_pool)
                    
                    # Process each base row
                    process_base_rows(
                        base_pool, country, is_lvl2, current_level, sheet_name,
                        sheets_data, seen, missing_schedule_info,
                        existing_offerings, original_ldap_data,
                        all_country_names, corp_names, non_corp_names
                    )
                except Exception as e:
                    if "Worksheet" not in str(e):
                        print(f"Error processing {sheet_name} in {wb}: {e}")
                    continue
        
        return sheets_data, missing_schedule_info

    def get_base_pool(wb, sheet_name, country, is_lvl2, excel_cache, column_order_cache):
        """Get the base pool of rows to process"""
        if use_new_parent:
            new_row = create_new_parent_row(
                new_parent_offering, new_parent, country, 
                business_criticality, approval_required, 
                approval_required_value, change_subscribed_location, 
                custom_subscribed_location
            )
            return pd.DataFrame([new_row])
        else:
            return load_and_filter_data(wb, sheet_name, country, is_lvl2, excel_cache, column_order_cache)

    def load_and_filter_data(wb, sheet_name, country, is_lvl2, excel_cache, column_order_cache):
        """Load and filter data from Excel file"""
        if wb not in excel_cache:
            excel_cache[wb] = pd.ExcelFile(wb)
        excel_file = excel_cache[wb]
        
        if sheet_name not in excel_file.sheet_names:
            return pd.DataFrame()
            
        df = pd.read_excel(excel_file, sheet_name=sheet_name)
        
        # Store column order and add missing columns
        column_key = f"{country}_{sheet_name}"
        if column_key not in column_order_cache:
            column_order_cache[column_key] = list(df.columns)
        
        for col in need_cols:
            if col not in df.columns:
                df[col] = ""
        
        if "Visibility group" not in df.columns:
            df["Visibility group"] = ""
        
        # Add LDAP columns for DE
        ldap_cols = [col for col in df.columns if "LDAP" in col.upper() or "Ldap" in col or "ldap" in col]
        if not ldap_cols and country == "DE":
            df["LDAP"] = ""
        
        # Apply filters
        return apply_filters(df, is_lvl2)

    def parse_keywords(keywords_str):
        """
        Parse keywords string and determine if AND logic should be used.
        Returns (list of keywords, use_and: bool)
        """
        keywords_str = str(keywords_str).strip()
        if not keywords_str:
            return [], False
        # Use AND logic if ' AND ' is present, otherwise OR logic
        if ' AND ' in keywords_str.upper():
            keywords = [k.strip() for k in re.split(r'AND', keywords_str, flags=re.IGNORECASE) if k.strip()]
            return keywords, True
        else:
            keywords = [k.strip() for k in re.split(r'[;,]', keywords_str) if k.strip()]
            return keywords, False

    def row_keywords_ok(row):
        """
        Check if all parent keywords are present in the Parent Offering column for AND logic.
        """
        parent_keywords, parent_use_and = parse_keywords(keywords_parent)
        parent_value = str(row["Parent Offering"]).lower()
        if parent_use_and and parent_keywords:
            return all(k.lower() in parent_value for k in parent_keywords)
        elif parent_keywords:
            return any(k.lower() in parent_value for k in parent_keywords)
        return True

    def row_excluded_keywords_ok(row):
        """
        Check if excluded keywords are NOT present in the Parent Offering column.
        """
        excluded_keywords = [k.strip() for k in re.split(r'[;,]', str(keywords_excluded)) if k.strip()]
        parent_value = str(row["Parent Offering"]).lower()
        if excluded_keywords:
            return not any(k.lower() in parent_value for k in excluded_keywords)
        return True

    def lc_ok(row):
        """
        Returns True if the Life Cycle Status and Life Cycle Stage are not in discard_lc.
        """
        lc_status = str(row.get("Life Cycle Status", "")).strip().lower()
        lc_stage = str(row.get("Life Cycle Stage", "")).strip().lower()
        return lc_status not in discard_lc and lc_stage not in discard_lc

    def name_prefix_ok(name):
        """
        Returns True if the name does not start with a forbidden prefix (if any such logic is needed).
        For now, always returns True.
        """
        return True

    def apply_filters(df, is_lvl2):
        """Apply all filters to the dataframe"""
        # Pre-filter with keywords
        if keywords_parent.strip():
            parent_col = df["Parent Offering"].astype(str).str.lower()
            parent_keywords, parent_use_and = parse_keywords(keywords_parent)
            if parent_keywords:
                if parent_use_and:
                    # AND logic - all keywords must be present
                    parent_mask = parent_col.str.contains('|'.join([re.escape(k.lower()) for k in parent_keywords]), na=False)
                    # Further filter with apply for exact AND logic
                    parent_mask = parent_mask & df.apply(row_keywords_ok, axis=1)
                else:
                    # OR logic - any keyword must be present
                    parent_mask = parent_col.str.contains('|'.join([re.escape(k.lower()) for k in parent_keywords]), na=False)
            else:
                parent_mask = pd.Series([True] * len(df), index=df.index)
            df = df[parent_mask]
        
        if df.empty:
            return df
        
        # Apply all filters
        if is_lvl2:
            mask = (df.apply(row_keywords_ok, axis=1)
                    & df.apply(row_excluded_keywords_ok, axis=1)
                    & df.apply(lc_ok, axis=1))
        else:
            mask = (df.apply(row_keywords_ok, axis=1)
                    & df.apply(row_excluded_keywords_ok, axis=1)
                    & df["Name (Child Service Offering lvl 1)"].astype(str).apply(name_prefix_ok)
                    & df.apply(lc_ok, axis=1)
                    & (df["Service Commitments"].astype(str).str.strip().replace({"nan": ""}) != "-"))
        
        return df.loc[mask]

    def get_schedule_checking_data(base_pool):
        """Get schedule checking data from base pool"""
        if "Name (Child Service Offering lvl 1)" in base_pool.columns:
            all_names = base_pool["Name (Child Service Offering lvl 1)"].astype(str).str.upper()
            corp_names = all_names[all_names.str.contains('CORP|DEDICATED|RECP', na=False)]
            non_corp_names = all_names[~all_names.str.contains('CORP|DEDICATED|RECP', na=False)]
        else:
            all_names = corp_names = non_corp_names = pd.Series([], dtype=str)
        
        return all_names, corp_names, non_corp_names

    def process_base_rows(base_pool, country, is_lvl2, current_level, sheet_name,
                         sheets_data, seen, missing_schedule_info,
                         existing_offerings, original_ldap_data,
                         all_country_names, corp_names, non_corp_names):
        """Process all base rows"""
        total_base_rows = len(base_pool)
        print(f"Processing {total_base_rows} base rows...")
        start_time = time.time()
        
        for row_idx, (idx, base_row) in enumerate(base_pool.iterrows()):
            if row_idx % 10 == 0 and row_idx > 0:
                elapsed = time.time() - start_time
                avg_time = elapsed / row_idx
                remaining = (total_base_rows - row_idx) * avg_time
                print(f"  Processed {row_idx}/{total_base_rows} base rows... ETA: {remaining:.1f}s")
            
            process_single_base_row(
                base_row, country, is_lvl2, current_level, sheet_name,
                sheets_data, seen, missing_schedule_info,
                existing_offerings, original_ldap_data,
                all_country_names, corp_names, non_corp_names,
                base_pool
            )
    def process_single_base_row(base_row, country, is_lvl2, current_level, sheet_name,
                               sheets_data, seen, missing_schedule_info,
                               existing_offerings, original_ldap_data,
                               all_country_names, corp_names, non_corp_names,
                               base_pool):
        """Process a single base row through all combinations"""
        base_row_df = base_row.to_frame().T.copy()
        receivers = get_receivers_for_country(country)
        parent_full = str(base_row["Parent Offering"])
        original_depend_on = str(base_row.get("Service Offerings | Depend On (Application Service)", "")).strip()
        
        for app in all_apps:
            for recv in receivers:
                country_schedule_suffixes = get_schedule_suffixes_for_country(
                    country, recv, schedule_settings_per_country, schedule_suffixes
                )
                
                for schedule_suffix in country_schedule_suffixes:
                    missing_schedule = check_missing_schedule(
                        schedule_suffix, all_country_names, corp_names, non_corp_names
                    )
                    
                    # Handle DE-specific row matching
                    if country == "DE" and not use_new_parent:
                        base_row, base_row_df, original_depend_on = handle_de_row_matching(
                            base_pool, recv, base_row, base_row_df
                        )
                    
                    # Build the name
                    new_name = build_offering_name(
                        parent_full, app, schedule_suffix, recv, is_lvl2, country
                    )
                    
                    new_name_normalized = ' '.join(new_name.split())
                    
                    # Check for duplicates
                    if new_name_normalized in existing_offerings:
                        raise ValueError(f"Sorry, it would be a duplicate - we already have this offering in the system: {new_name}")
                    
                    # Process support groups
                    process_support_groups(
                        country, recv, new_name, new_name_normalized, app, schedule_suffix,
                        base_row_df, is_lvl2, current_level, sheet_name,
                        sheets_data, seen, missing_schedule_info, missing_schedule,
                        original_ldap_data, original_depend_on
                    )
                    

    def get_receivers_for_country(country):
        """Get receivers for a specific country"""
        if country == "PL":
            return ["HS PL", "DS PL"]
        elif country == "CY":
            return ["DS CY"]
        elif country == "DE":
            return ["HS DE", "DS DE"]
        elif country == "UA":
            return ["DS UA"]
        elif country == "MD":
            return ["DS MD"]
        elif country == "RO":
            return ["DS RO"]
        elif country == "TR":
            return ["DS TR"]
        else:
            return [f"HS {country}", f"DS {country}"]

    def check_missing_schedule(schedule_suffix, all_country_names, corp_names, non_corp_names):
        """Check if schedule is missing from source data"""
        if len(all_country_names) == 0:
            return False
        
        is_corp_type = (require_corp or require_recp or require_corp_it or require_corp_dedicated)
        schedule_pattern = schedule_suffix.strip()
        
        if is_corp_type:
            if len(corp_names) > 0:
                return not any(schedule_pattern in name for name in corp_names)
            else:
                return True
        else:
            if len(non_corp_names) > 0:
                return not any(schedule_pattern in name for name in non_corp_names)
            else:
                return True

    def handle_de_row_matching(base_pool, recv, base_row, base_row_df):
        """Handle DE-specific row matching"""
        recv_mask = base_pool["Name (Child Service Offering lvl 1)"].str.contains(
            rf"\b{re.escape(recv)}\b", case=False
        )
        if recv_mask.any():
            base_row = base_pool[recv_mask].iloc[0]
            base_row_df = base_row.to_frame().T.copy()
            original_depend_on = str(base_row.get("Service Offerings | Depend On (Application Service)", "")).strip()
        else:
            original_depend_on = str(base_row.get("Service Offerings | Depend On (Application Service)", "")).strip()
        
        return base_row, base_row_df, original_depend_on

    def build_offering_name(parent_full, app, schedule_suffix, recv, is_lvl2, country):
        """Build the offering name based on type and parameters"""
        if is_lvl2:
            return build_lvl2_name(parent_full, sr_or_im, app, schedule_suffix, service_type_lvl2)
        elif require_corp:
            return build_corp_name(parent_full, sr_or_im, app, schedule_suffix, recv, delivering_tag)
        elif require_recp:
            return build_recp_name(parent_full, sr_or_im, app, schedule_suffix, recv, delivering_tag)
        elif require_corp_it:
            return build_corp_it_name(parent_full, sr_or_im, app, schedule_suffix, recv, delivering_tag)
        elif require_corp_dedicated:
            return build_corp_dedicated_name(parent_full, sr_or_im, app, schedule_suffix, recv, delivering_tag)
        else:
            return build_standard_offering_name(parent_full, app, schedule_suffix, recv, country)

    def build_standard_offering_name(parent_full, app, schedule_suffix, recv, country):
        """Build standard offering name with DE special handling"""
        if country == "DE" and recv:
            parent_content = extract_parent_info(parent_full)
            catalog_name = extract_catalog_name(parent_full)
            
            parts = parent_content.split()
            new_parts = [sr_or_im]
            
            recv_division = recv.split()[0]
            new_parts.append(recv_division)
            
            for part in parts:
                if part in ["HS", "DS"]:
                    continue
                elif len(part) == 2 and part.isupper() and part not in ["IT", "HR"]:
                    new_parts.append(part)
                elif part in ["IT", "HR", "Medical", "Business Services"] or (part not in ["HS", "DS"] and not (len(part) == 2 and part.isupper())):
                    new_parts.append(part)
            
            new_parent_offering = f"[Parent {' '.join(new_parts)}] {catalog_name}"
            return build_standard_name(new_parent_offering, sr_or_im, app, schedule_suffix, special_dept, recv)
        else:
            return build_standard_name(parent_full, sr_or_im, app, schedule_suffix, special_dept, recv)

    def process_support_groups(country, recv, new_name, new_name_normalized, app, schedule_suffix,
                              base_row_df, is_lvl2, current_level, sheet_name,
                              sheets_data, seen, missing_schedule_info, missing_schedule,
                              original_ldap_data, original_depend_on):
        """Process support groups for the offering"""
        division = get_division_for_pl(country, new_name, base_row_df)
        support_groups_list = get_support_groups_for_offering(country, recv, division)
        
        # Filter for DE
        if country == "DE" and recv:
            prefix = recv
            matching = [(sg, mg) for sg, mg in support_groups_list if sg.strip().startswith(prefix)]
            if matching:
                support_groups_list = matching
        
        for support_group_for_country, managed_by_group_for_country in support_groups_list:
            key = (new_name_normalized, recv, app, schedule_suffix, support_group_for_country, managed_by_group_for_country)
            
            if key in seen:
                continue
            seen.add(key)
            
            create_offering_row(
                base_row_df, new_name, app, support_group_for_country, managed_by_group_for_country,
                country, recv, schedule_suffix, is_lvl2, current_level, sheet_name,
                sheets_data, missing_schedule, original_ldap_data, original_depend_on
            )

    def get_division_for_pl(country, new_name, base_row_df):
        """Get division for PL offerings"""
        if country != "PL":
            return None
        
        if "HS PL" in new_name or "HS PL" in str(base_row_df.iloc[0].get("Name (Child Service Offering lvl 1)", "")):
            return "HS"
        elif "DS PL" in new_name or "DS PL" in str(base_row_df.iloc[0].get("Name (Child Service Offering lvl 1)", "")):
            return "DS"
        else:
            return "HS"  # Default

    def get_support_groups_for_offering(country, recv, division):
        """Get support groups list for the offering"""
        if country == "PL":
            key = recv
            sg_dict = support_groups_per_country if support_groups_per_country is not None else {}
            mg_dict = managed_by_groups_per_country if managed_by_groups_per_country is not None else {}
            country_supports = sg_dict.get(key, "")
            country_managed = mg_dict.get(key, "")
            
            if country_supports:
                sg = str(country_supports).strip()
                mg = str(country_managed or sg).strip()
                return [(sg, mg)]
            else:
                return [("", "")]
        else:
            return get_support_groups_list_for_country(
                country, support_group, support_groups_per_country if support_groups_per_country is not None else {}, 
                managed_by_groups_per_country if managed_by_groups_per_country is not None else {}, division
            )

    def create_offering_row(base_row_df, new_name, app, support_group_for_country, managed_by_group_for_country,
                           country, recv, schedule_suffix, is_lvl2, current_level, sheet_name,
                           sheets_data, missing_schedule, original_ldap_data, original_depend_on):
        """Create a single offering row with all settings"""
        row = base_row_df.copy()
        
        # Basic settings
        row.loc[:, "Name (Child Service Offering lvl 1)"] = new_name
        row.loc[:, "Parent"] = new_parent if use_new_parent else ""
        row.loc[:, "Delivery Manager"] = delivery_manager
        
        # Business criticality
        if business_criticality:
            row.loc[:, "Business Criticality"] = business_criticality
        
        # Record view
        if sr_or_im == "SR":
            row.loc[:, "Record view"] = "Request Item"
        elif sr_or_im == "IM":
            row.loc[:, "Record view"] = "Incident, Major Incident"
        
        # Approval settings
        if approval_required:
            row.loc[:, "Approval required"] = approval_required_value
            row.loc[:, "Approval group"] = approval_required_value
        else:
            row.loc[:, "Approval required"] = "false"
            row.loc[:, "Approval group"] = "empty"
        
        # Location settings
        if change_subscribed_location:
            row.loc[:, "Subscribed by Location"] = custom_subscribed_location
        else:
            row.loc[:, "Subscribed by Location"] = "Global"
        
        # Support groups
        row.loc[:, "Support group"] = support_group_for_country if support_group_for_country else ""
        row.loc[:, "Managed by Group"] = managed_by_group_for_country if managed_by_group_for_country else ""
        
        # Handle aliases
        handle_aliases(row, country, recv, app)
        
        # Handle company and LDAP
        handle_company_and_ldap(row, country, recv, support_group_for_country, new_name, original_ldap_data, base_row_df)
        
        # Handle commitments
        handle_commitments(row, country, schedule_suffix, is_lvl2, missing_schedule)
        
        # Handle depend on
        handle_depend_on(row, country, recv, app, new_name, original_depend_on)
        
        # Add to sheets data
        add_to_sheets_data(row, current_level, country, sheet_name, sheets_data, missing_schedule)

    def handle_aliases(row, country, recv, app):
        """Handle aliases based on settings"""
        if not aliases_on:
            return
        
        alias_value_to_use = ""
        
        if aliases_per_country:
            if country == "PL" and recv:
                alias_value_to_use = aliases_per_country.get(recv, "")
            else:
                alias_value_to_use = aliases_per_country.get(country, "")
        
        if not alias_value_to_use:
            alias_value_to_use = aliases_value
        
        if alias_value_to_use == "USE_APP_NAMES":
            alias_value_to_use = app if app else ""
        
        if not (alias_value_to_use and selected_languages):
            return
        
        alias_columns = [c for c in row.columns if "Alias" in c]
        matching_columns = []
        
        for col in alias_columns:
            for lang in selected_languages:
                if (f"- {lang}" in col or f"({lang})" in col or 
                    f"_{lang}" in col or col.endswith(f" {lang}")):
                    matching_columns.append(col)
                    break
                elif lang == "ENG" and any(x in col for x in ["- EN", "- ENGLISH", "(EN)", "(ENGLISH)"]):
                    matching_columns.append(col)
                    break
                elif lang == "DE" and any(x in col for x in ["- GER", "- GERMAN", "(GER)", "(GERMAN)"]):
                    matching_columns.append(col)
                    break
        
        matching_columns = list(dict.fromkeys(matching_columns))
        
        for col in matching_columns:
            row.loc[:, col] = alias_value_to_use

    def handle_company_and_ldap(row, country, recv, support_group_for_country, new_name, original_ldap_data, base_row_df):
        """Handle company and LDAP settings"""
        if use_new_parent:
            if require_corp or require_recp or require_corp_it or require_corp_dedicated:
                match = re.search(r'\[.*?CORP\s+([A-Z]{2}\s+[A-Z]{2})', new_name)
                if match:
                    row.loc[:, "Subscribed by Company"] = match.group(1)
                else:
                    row.loc[:, "Subscribed by Company"] = recv
            else:
                row.loc[:, "Subscribed by Company"] = recv
        elif country == "DE":
            company, ldap = get_de_company_and_ldap(support_group_for_country, recv, base_row_df.iloc[0])
            row.loc[:, "Subscribed by Company"] = company
            
            ldap_cols = [col for col in row.columns if "LDAP" in col.upper() or "Ldap" in col or "ldap" in col]
            if ldap_cols:
                for ldap_col in ldap_cols:
                    row.loc[:, ldap_col] = ""
                
                if ldap and ldap not in ["", "nan", "NaN", "None"]:
                    row.loc[:, ldap_cols[0]] = ldap
                else:
                    if support_group_for_country in original_ldap_data:
                        ldap_data = original_ldap_data[support_group_for_country]
                        for ldap_col, ldap_val in ldap_data.items():
                            if ldap_col in row.columns and ldap_val not in ["", "nan", "NaN", "None"]:
                                row.loc[:, ldap_col] = ldap_val
        elif require_corp or require_recp or require_corp_it or require_corp_dedicated:
            match = re.search(r'\[.*?CORP\s+([A-Z]{2}\s+[A-Z]{2})', new_name)
            if match:
                row.loc[:, "Subscribed by Company"] = match.group(1)
            else:
                row.loc[:, "Subscribed by Company"] = recv
        else:
            if "Subscribed by Company" in base_row_df.iloc[0].index:
                original_company = str(base_row_df.iloc[0]["Subscribed by Company"]).strip()
                if original_company and original_company not in ["nan", "NaN", "", "None", "none"]:
                    row.loc[:, "Subscribed by Company"] = original_company
                else:
                    row.loc[:, "Subscribed by Company"] = ""
            else:
                row.loc[:, "Subscribed by Company"] = ""

    def handle_commitments(row, country, schedule_suffix, is_lvl2, missing_schedule):
        """Handle service commitments"""
        orig_comm = str(row.iloc[0]["Service Commitments"]).strip()
        
        if missing_schedule:
            if not orig_comm or orig_comm in ["-", "nan", "NaN", "", None]:
                row.loc[:, "Service Commitments"] = commit_block(country, schedule_suffix, rsp_duration, rsl_duration, sr_or_im)
            else:
                row.loc[:, "Service Commitments"] = update_commitments(orig_comm, schedule_suffix, rsp_duration, rsl_duration, sr_or_im, country)
        elif is_lvl2 and (not orig_comm or orig_comm in ["-", "nan", "NaN", "", None]):
            row.loc[:, "Service Commitments"] = ""
        else:
            if use_custom_commitments and custom_commitments_str:
                row.loc[:, "Service Commitments"] = custom_commitments_str
            elif use_custom_commitments and commitment_country:
                row.loc[:, "Service Commitments"] = custom_commit_block(
                    commitment_country, sr_or_im, rsp_enabled, rsl_enabled,
                    rsp_schedule, rsl_schedule, rsp_priority, rsl_priority,
                    rsp_time, rsl_time
                )
            else:
                if not orig_comm or orig_comm == "-":
                    row.loc[:, "Service Commitments"] = commit_block(country, schedule_suffix, rsp_duration, rsl_duration, sr_or_im)
                else:
                    row.loc[:, "Service Commitments"] = update_commitments(orig_comm, schedule_suffix, rsp_duration, rsl_duration, sr_or_im, country)

    def handle_depend_on(row, country, recv, app, new_name, original_depend_on):
        """Handle depend on application service"""
        if (special_dept == "IT" or require_corp_it) and country in ["UA", "MD", "RO", "TR"]:
            depend_tag = f"DS {country} Prod"
        elif global_prod:
            depend_tag = "Global Prod"
        else:
            if country == "PL":
                if re.search(r'\bHS\s+PL\b', new_name, re.IGNORECASE):
                    depend_tag = "HS PL Prod"
                elif re.search(r'\bDS\s+PL\b', new_name, re.IGNORECASE):
                    depend_tag = "DS PL Prod"
                else:
                    depend_tag = "DS PL Prod"
            elif recv:
                depend_tag = f"{recv} Prod"
            else:
                tag_hs = f"HS {country}"
                depend_tag = f"{delivering_tag} Prod" if (require_corp or require_recp or require_corp_it or require_corp_dedicated) else f"{tag_hs} Prod"
        
        if use_custom_depend_on and custom_depend_on_value:
            if app:
                row.loc[:, "Service Offerings | Depend On (Application Service)"] = f"{custom_depend_on_value} {app}"
            else:
                row.loc[:, "Service Offerings | Depend On (Application Service)"] = custom_depend_on_value
        elif app:
            row.loc[:, "Service Offerings | Depend On (Application Service)"] = f"[{depend_tag}] {app}"
        else:
            row.loc[:, "Service Offerings | Depend On (Application Service)"] = ""

    def add_to_sheets_data(row, current_level, country, sheet_name, sheets_data, missing_schedule):
        """Add row to sheets data"""
        sheet_key = f"{country} lvl{current_level}"
        column_key = f"{country}_{sheet_name}"
        
        sheets_data.setdefault(sheet_key, [])
        
        if isinstance(row, pd.DataFrame):
            row_dict = row.iloc[0].to_dict()
        else:
            row_dict = row.to_dict()
        
        row_dict["_missing_schedule"] = missing_schedule
        row_dict["_column_order_key"] = column_key
        
        sheets_data[sheet_key].append(row_dict)

    def write_output(sheets_data, missing_schedule_info, column_order_cache):
        """Write output to Excel file"""
        if not sheets_data or all(not rows_list for rows_list in sheets_data.values()):
            print("❌ No matching offerings found with the specified keywords.")
            raise ValueError("No matching offerings found. Please adjust your search criteria.")
        
        outfile = out_dir / f"Generated_Service_Offerings_{dt.datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        out_dir.mkdir(parents=True, exist_ok=True)
        
        with pd.ExcelWriter(outfile, engine="openpyxl") as w:
            sheets = {}
            
            for sheet_key, rows_list in sheets_data.items():
                if rows_list:
                    print(f"  Processing {sheet_key}: {len(rows_list)} rows")
                    df = pd.DataFrame(rows_list)
                    
                    # Get column order and clean dataframe
                    df_cleaned = clean_dataframe(df, sheet_key, column_order_cache, missing_schedule_info)
                    
                    sheets[sheet_key] = df_cleaned
                    df_cleaned.to_excel(w, sheet_name=sheet_key, index=False)
        
        return outfile, sheets

    def clean_dataframe(df, sheet_key, column_order_cache, missing_schedule_info):
        """Clean and reorder dataframe columns"""
        column_order_key = None
        if "_column_order_key" in df.columns and len(df) > 0:
            column_order_key = df.iloc[0]["_column_order_key"]
        
        sheet_name = sheet_key
        missing_schedule_rows = []
        if "_missing_schedule" in df.columns:
            missing_schedule_rows = df[df["_missing_schedule"] == True].index.tolist()
            missing_schedule_info[sheet_name] = missing_schedule_rows
        
        for col in ["_missing_schedule", "_column_order_key"]:
            if col in df.columns:
                df = df.drop(columns=[col])
        
        # Reorder columns
        if column_order_key and column_order_key in column_order_cache:
            df = reorder_columns(df, column_order_cache[column_order_key])
        
        # Clean data
        return clean_dataframe_values(df, sheet_key)

    def reorder_columns(df, original_order):
        """Reorder columns to match original order"""
        ordered_cols = [col for col in original_order if col in df.columns and col != "Number"]
        new_cols = [col for col in df.columns if col not in original_order]
        return df[ordered_cols + new_cols]

    def clean_dataframe_values(df, sheet_key):
        """Clean dataframe values for Excel output"""
        df_final = df.copy()
        
        for col in df_final.columns:
            try:
                if df_final[col].dtype == 'object':
                    df_final[col] = df_final[col].fillna('')
                    df_final[col] = df_final[col].apply(lambda x: str(x) if pd.notna(x) and x != '' else '')
                elif df_final[col].dtype in ['int64', 'float64']:
                    df_final[col] = df_final[col].fillna('')
                else:
                    df_final[col] = df_final[col].fillna('')
                
                if df_final[col].dtype == 'object':
                    df_final[col] = df_final[col].replace({
                        'nan': '', 'NaN': '', 'None': '', 'none': '',
                        'NULL': '', 'null': '', '<NA>': ''
                    })
            except Exception as e:
                print(f"Warning: Error processing column {col}: {e}")
                df_final[col] = df_final[col].fillna('').astype(str)
        
        # Handle special columns
        if "Approval required" in df_final.columns:
            df_final["Approval required"] = df_final["Approval required"].apply(clean_approval_value)
        
        if "Approval group" in df_final.columns:
            df_final["Approval group"] = df_final["Approval group"].apply(clean_approval_group)
        
        return df_final

    def clean_approval_value(val):
        """Clean approval required values"""
        if pd.isna(val) or str(val).strip() in ['', 'nan', 'NaN', 'None', 'none']:
            return 'false'
        val_str = str(val).strip().lower()
        if val_str in ['true', 'false']:
            return val_str
        elif val_str in ['yes', 'y', '1']:
            return 'true'
        elif val_str in ['no', 'n', '0']:
            return 'false'
        else:
            return str(val).strip()

    def clean_approval_group(val):
        """Clean approval group values"""
        if pd.isna(val) or str(val).strip() in ['', 'nan', 'NaN', 'None', 'none', 'NULL', 'null', '<NA>']:
            return 'empty'
        return str(val).strip()

    def apply_formatting(outfile, missing_schedule_info):
        """Apply Excel formatting"""
        try:
            wb = load_workbook(outfile)
            red_fill = PatternFill(start_color="FFCCCC", end_color="FFCCCC", fill_type="solid")
            
            for ws in wb.worksheets:
                sheet_name = ws.title
                
                # Apply red formatting for missing schedules
                name_col_idx = None
                for idx, cell in enumerate(ws[1], start=1):
                    if cell.value == "Name (Child Service Offering lvl 1)":
                        name_col_idx = idx
                        break
                
                if name_col_idx is not None:
                    name_col_letter = get_column_letter(name_col_idx)
                    for row_idx in missing_schedule_info.get(sheet_name, []):
                        excel_row_idx = row_idx + 2
                        if excel_row_idx <= ws.max_row:
                            cell = ws[f"{name_col_letter}{excel_row_idx}"]
                            cell.fill = red_fill
                
                # Apply column formatting
                for col_idx, col in enumerate(ws.columns, start=1):
                    col_letter = get_column_letter(col_idx)
                    
                    max_length = 10
                    for cell in col:
                        try:
                            cell_length = len(str(cell.value)) if cell.value else 0
                            if cell_length > max_length:
                                max_length = min(cell_length, 100)
                        except:
                            pass
                    
                    ws.column_dimensions[col_letter].width = max_length + 2
                    
                    for cell in col:
                        try:
                            if hasattr(cell, 'alignment'):
                                cell.alignment = Alignment(wrap_text=True)
                        
                            from openpyxl.cell.cell import MergedCell
                            if not isinstance(cell, MergedCell) and hasattr(cell, 'value') and cell.value is not None:
                                cell_val = str(cell.value).strip()
                                if cell_val.lower() in ['nan', 'none', 'null', '<na>', 'n/a'] or cell_val == '':
                                    cell.value = None
                        except Exception:
                            continue
            
            wb.save(outfile)
        except Exception as e:
            print(f"Warning: Error applying formatting: {e}")

    # Main execution flow
    initialize_settings()
    all_apps = parse_all_apps(new_apps)
    special_dept = determine_special_dept(special_it, special_hr, special_medical, special_dak)
    
    existing_offerings, original_ldap_data, excel_cache, column_order_cache = collect_existing_offerings(src_dir)
    
    sheets_data, missing_schedule_info = process_files(
        src_dir, excel_cache, column_order_cache, existing_offerings, original_ldap_data
    )
    
    outfile, sheets = write_output(sheets_data, missing_schedule_info, column_order_cache)
    
    apply_formatting(outfile, missing_schedule_info)
    
    # Cleanup
    for excel_file in excel_cache.values():
        try:
            if isinstance(excel_file, pd.ExcelFile):
                excel_file.close()
        except:
            pass
    
    print("Processing complete. Output saved to:", outfile)
    return outfile