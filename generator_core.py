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
    "Subscribed by Company",
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
    
    if special_dept:  # IT or HR case
        # Extract division and country from parent content
        parts = parent_content.split()
        division = ""
        country = ""
        for part in parts:
            if part in ["HS", "DS"]:
                division = part
            elif len(part) == 2 and part.isupper() and part not in ["IT", "HR"]:
                country = part
        
        # Build the name with special department
        prefix_parts = [sr_or_im]
        if division:
            prefix_parts.append(division)
        if country:
            prefix_parts.append(country)
        prefix_parts.append(special_dept)
        
        # Extract the topic from parent content (what's after division and country)
        topic_parts = []
        skip_next = False
        for i, part in enumerate(parts):
            if part in ["HS", "DS"] or (len(part) == 2 and part.isupper() and part not in ["IT", "HR"]):
                skip_next = True
                continue
            if skip_next:
                skip_next = False
                continue
            topic_parts.append(part)
        
        topic = " ".join(topic_parts) if topic_parts else "Software"
        
        return f"[{' '.join(prefix_parts)}] {topic} {catalog_name.lower()} {app} Prod {schedule_suffix}"
    else:
        # Standard case - just replace Parent with SR/IM
        return f"[{sr_or_im} {parent_content}] {catalog_name} {app} Prod {schedule_suffix}"

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
        return f"[{' '.join(prefix_parts)}] {catalog_name} {app} Prod {schedule_suffix}"
    else:
        return f"[{' '.join(prefix_parts)}] {catalog_name} solving {app} Prod {schedule_suffix}"

def run_generator(*,
    keywords, new_apps, schedule_suffix,
    delivery_manager, global_prod,
    rsp_duration, rsl_duration,
    sr_or_im, require_corp, delivering_tag,
    support_group, managed_by_group, aliases_on, aliases_value,
    src_dir: Path, out_dir: Path,
    special_it=False, special_hr=False):  # IT and HR as separate flags

    sheets, seen = {}, set()

    def row_keywords_ok(row):
        if not keywords:
            return True
        
        # First keyword filters Parent Offering
        if len(keywords) >= 1:
            p = str(row["Parent Offering"]).lower()
            if not re.search(rf"\b{re.escape(keywords[0].lower())}\b", p):
                return False
        
        # Subsequent keywords filter Name (Child Service Offering lvl 1) - ANY match (OR logic)
        if len(keywords) > 1:
            n = str(row["Name (Child Service Offering lvl 1)"]).lower()
            # Check if ANY of the subsequent keywords match
            matches = False
            for k in keywords[1:]:
                if re.search(rf"\b{re.escape(k.lower())}\b", n):
                    matches = True
                    break
            if not matches:
                return False
        
        return True

    def lc_ok(row):
        return all(str(row[c]).strip().lower() not in discard_lc
                   for c in ("Phase","Status","Life Cycle Stage","Life Cycle Status"))

    def name_prefix_ok(name):
        return name.lower().startswith(f"[{sr_or_im.lower()} ")

    def update_commitments(orig, sched, rsp, rsl):
        out=[]
        for line in str(orig).splitlines():
            if "RSP" in line:
                line=re.sub(r"RSP .*? P1-P4",f"RSP {sched} P1-P4",line)
                line=re.sub(r"P1-P4 .*?$",f"P1-P4 {rsp}",line)
            elif "RSL" in line:
                line=re.sub(r"RSL .*? P1-P4",f"RSL {sched} P1-P4",line)
                line=re.sub(r"P1-P4 .*?$",f"P1-P4 {rsl}",line)
            out.append(line)
        return "\n".join(out)

    def commit_block(cc):
        lines=[
            f"[{cc}] SLA SR RSP {schedule_suffix} P1-P4 {rsp_duration}",
            f"[{cc}] SLA SR RSL {schedule_suffix} P1-P4 {rsl_duration}",
            f"[{cc}] OLA SR RSL {schedule_suffix} P1-P4 {rsl_duration}"
        ]
        return "\n".join(lines)

    # Process apps - support both newline and comma separation
    all_apps = []
    for app_line in new_apps:
        # Split by comma and add non-empty values
        for app in app_line.split(','):
            app = app.strip()
            if app:
                all_apps.append(app)

    # Determine special department
    special_dept = None
    if special_it and not require_corp:
        special_dept = "IT"
    elif special_hr and not require_corp:
        special_dept = "HR"

    for wb in src_dir.glob("ALL_Service_Offering_*.xlsx"):
        df=pd.read_excel(wb,sheet_name="Child SO lvl1")
        if any(c not in df.columns for c in need_cols):
            continue

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
        base_row=base_pool.iloc[0].to_frame().T.copy()

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

        parent_full=str(base_row.iloc[0]["Parent Offering"])

        for app in all_apps:
            for recv in receivers:
                if require_corp:
                    new_name = build_corp_name(
                        parent_full, sr_or_im, app, schedule_suffix, recv, delivering_tag
                    )
                else:
                    new_name = build_standard_name(
                        parent_full, sr_or_im, app, schedule_suffix, special_dept
                    )
                
                if new_name in seen:
                    continue
                seen.add(new_name)

                row=base_row.copy()
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
                row["Service Commitments"]=commit_block(country) if not orig_comm or orig_comm=="-" else update_commitments(orig_comm,schedule_suffix,rsp_duration,rsl_duration)
                
                if global_prod:
                    row["Service Offerings | Depend On (Application Service)"]=f"[Global Prod] {app}"
                else:
                    depend_tag = f"{delivering_tag} Prod" if require_corp else f"{recv or tag_hs} Prod"
                    row["Service Offerings | Depend On (Application Service)"]=f"[{depend_tag}] {app}"
                
                sheets.setdefault(country,pd.DataFrame())
                sheets[country]=pd.concat([sheets[country],row],ignore_index=True)

    if not sheets:
        raise ValueError("No matching rows found with the specified keywords.")

    out_dir.mkdir(parents=True,exist_ok=True)
    outfile=out_dir / f"Offerings_NEW_{dt.datetime.now():%Y%m%d_%H%M%S}.xlsx"
    with pd.ExcelWriter(outfile,engine="openpyxl") as w:
        for cc,dfc in sheets.items():
            # Ensure unique names per country
            df_final = dfc.drop_duplicates(subset=["Name (Child Service Offering lvl 1)"])
            
            if "Number" in df_final.columns:
                df_final = df_final.drop(columns=["Number"])
            
            for col in df_final.columns:
                if df_final[col].dtype == 'bool':
                    df_final[col] = df_final[col].map({True: 'true', False: 'false'})
                elif df_final[col].dtype == 'object':
                    df_final[col] = df_final[col].astype(str).replace({'True': 'true', 'False': 'false'})
            
            df_final.to_excel(w,sheet_name=cc,index=False)
    
    wb=load_workbook(outfile)
    for ws in wb.worksheets:
        ws.auto_filter.ref=ws.dimensions
        for col in ws.columns:
            ws.column_dimensions[col[0].column_letter].width=max(len(str(c.value)) if c.value else 0 for c in col)+2
            for c in col:
                c.alignment=Alignment(wrap_text=True)
    wb.save(outfile)
    return outfile