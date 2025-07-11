import datetime as dt
import re, warnings
from pathlib import Path
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Alignment

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

NAMING_CONVENTIONS = {
    "parent_non_software": "Parent Service Offering (does not apply to Service Offerings related to software model)",
    "parent_sam": "Parent Service Offering (Service Offerings related to software model - SAM)",
    "child_lvl1_parent": "Child Offering LVL1 PARENT (applicable only for Software & Permissions, if there is a Child Offering Lvl2)",
    "child_lvl1_software": "Child Offering LVL1 SOFTWARE & PERMISSIONS (applicable only if there is no Child Offering Lvl2)",
    "child_lvl1_hardware": "Child Offering LVL1 HARDWARE or NETWORK",
    "child_lvl1_other": "Child Offering LVL1 OTHER (i.e. MAILBOX, DEDICATED SERVICES etc.)",
    "child_lvl2_non_software": "Child Offering LVL2 (does not apply to Service Offerings related to software model)",
    "child_lvl2_sam": "Child Offering LVL2 (Service Offerings related to software model - SAM)",
    "exception": "EXCEPTION Child Offering LVL1 / LVL2",
    "child_lvl3": "Child Offering LVL 3 (application functionality)",
}

need_cols = [
    "Name (Child Service Offering lvl 1)", "Parent Offering",
    "Service Offerings | Depend On (Application Service)", "Service Commitments",
    "Delivery Manager", "Subscribed by Location", "Phase", "Status",
    "Life Cycle Stage", "Life Cycle Status", "Support group", "Managed by Group",
    "Subscribed by Company",
]

skip_hs_for = {"MD", "UA", "PL"}
discard_lc  = {"retired", "retiring", "end of life", "end of support"}

def extract_info_from_parent(parent_offering):
    info = {
        "country": "",
        "division": "",
        "topic": "",
        "catalog_name": "",
        "parent_inner": ""
    }
    
    match = re.search(r'\[Parent\s+(.*?)\]', str(parent_offering), re.I)
    if match:
        inner = match.group(1).strip()
        info["parent_inner"] = inner
        
        parts = inner.split()
        for part in parts:
            if part in ["HS", "DS"]:
                info["division"] = part
            elif len(part) == 2 and part.isupper():
                info["country"] = part
            elif part not in ["Parent", "RecP"]:
                info["topic"] = part
    
    parts = str(parent_offering).split(']', 1)
    if len(parts) > 1:
        info["catalog_name"] = parts[1].strip()
    
    return info

def build_name_parent_non_software(info, sr_or_im):
    parts = ["Parent"]
    if info["division"]:
        parts.append(info["division"])
    parts.extend([info["country"], info["topic"]])
    
    return f"[{' '.join(parts)}] {info['catalog_name']}"

def build_name_parent_sam(info, sr_or_im, app, schedule_suffix):
    parts = [sr_or_im]
    if info["division"]:
        parts.append(info["division"])
    parts.extend([info["country"], info["topic"]])
    
    return f"[{' '.join(parts)}] Parent {info['catalog_name']} {app} {schedule_suffix}"

def build_name_child_lvl1_parent(info, sr_or_im, app):
    parts = [sr_or_im]
    if info["division"]:
        parts.append(info["division"])
    parts.extend([info["country"], info["topic"]])
    
    if sr_or_im == "SR":
        return f"[{' '.join(parts)}] Parent {info['catalog_name']} {app} Prod"
    else:
        return f"[{' '.join(parts)}] Parent {info['catalog_name']} solving {app} Prod"

def build_name_child_lvl1_software(info, sr_or_im, app, schedule_suffix, require_corp, receiver, delivering_tag):
    parts = [sr_or_im]
    
    if require_corp:
        delivering_division = info["division"] or "HS"
        parts.extend([delivering_division, info["country"], "CORP", receiver, info["topic"] or "IT"])
    else:
        if info["division"]:
            parts.append(info["division"])
        parts.extend([info["country"], info["topic"] or "IT"])
    
    base = f"[{' '.join(parts)}]"
    
    if sr_or_im == "SR":
        return f"{base} {info['catalog_name']} {app} Prod {schedule_suffix}"
    else:
        return f"{base} {info['catalog_name']} solving {app} Prod {schedule_suffix}"

def build_name_child_lvl1_hardware(info, sr_or_im, schedule_suffix):
    parts = [sr_or_im]
    if info["division"]:
        parts.append(info["division"])
    parts.extend([info["country"], info["topic"] or "IT"])
    
    base = f"[{' '.join(parts)}]"
    
    if sr_or_im == "SR":
        return f"{base} {info['catalog_name']} {schedule_suffix}"
    else:
        return f"{base} {info['catalog_name']} solving {schedule_suffix}"

def build_name_child_lvl1_other(info, sr_or_im, schedule_suffix, service_type=None):
    parts = [sr_or_im]
    if info["division"]:
        parts.append(info["division"])
    parts.extend([info["country"], info["topic"] or "IT"])
    
    base = f"[{' '.join(parts)}]"
    
    catalog = info['catalog_name']
    if service_type:
        catalog = f"{catalog} {service_type}"
    
    if sr_or_im == "SR":
        return f"{base} {catalog} {schedule_suffix}"
    else:
        return f"{base} {catalog} solving {schedule_suffix}"

def build_name_by_convention(naming_convention, **kwargs):
    info = extract_info_from_parent(kwargs.get("parent_offering", ""))
    
    if naming_convention == "parent_non_software":
        return build_name_parent_non_software(info, kwargs.get("sr_or_im"))
    
    elif naming_convention == "parent_sam":
        return build_name_parent_sam(
            info, 
            kwargs.get("sr_or_im"),
            kwargs.get("app"),
            kwargs.get("schedule_suffix")
        )
    
    elif naming_convention == "child_lvl1_parent":
        return build_name_child_lvl1_parent(
            info,
            kwargs.get("sr_or_im"),
            kwargs.get("app")
        )
    
    elif naming_convention == "child_lvl1_software":
        return build_name_child_lvl1_software(
            info,
            kwargs.get("sr_or_im"),
            kwargs.get("app"),
            kwargs.get("schedule_suffix"),
            kwargs.get("require_corp"),
            kwargs.get("receiver"),
            kwargs.get("delivering_tag")
        )
    
    elif naming_convention == "child_lvl1_hardware":
        return build_name_child_lvl1_hardware(
            info,
            kwargs.get("sr_or_im"),
            kwargs.get("schedule_suffix")
        )
    
    elif naming_convention == "child_lvl1_other":
        return build_name_child_lvl1_other(
            info,
            kwargs.get("sr_or_im"),
            kwargs.get("schedule_suffix"),
            kwargs.get("service_type")
        )
    
    else:
        return build_name_child_lvl1_software(
            info,
            kwargs.get("sr_or_im"),
            kwargs.get("app"),
            kwargs.get("schedule_suffix"),
            kwargs.get("require_corp"),
            kwargs.get("receiver"),
            kwargs.get("delivering_tag")
        )

def run_generator(*,
    keywords, new_apps, schedule_suffix,
    delivery_manager, global_prod,
    rsp_duration, rsl_duration,
    sr_or_im, require_corp, delivering_tag,
    support_group, managed_by_group, aliases_on,
    src_dir: Path, out_dir: Path,
    naming_convention: str = "child_lvl1_software"):

    aliases_value = "" if aliases_on else "-"
    sheets, seen = {}, set()

    def row_keywords_ok(row):
        p = str(row["Parent Offering"]).lower()
        n = str(row["Name (Child Service Offering lvl 1)"]).lower()
        hit_p = {k for k in keywords if re.search(rf"\b{re.escape(k)}\b", p)}
        hit_n = {k for k in keywords if re.search(rf"\b{re.escape(k)}\b", n)}
        return bool(hit_p) and bool(hit_n) and hit_p | hit_n == set(keywords)

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

        if require_corp:
            if country=="DE":         receivers=["DS DE","HS DE"]
            elif country in {"UA","MD"}: receivers=[f"DS {country}"]
            elif country=="PL":         receivers=["DS PL"] if "DS PL" in base_pool["Name (Child Service Offering lvl 1)"].str.cat(sep=" ") else ["HS PL"]
            elif country=="CY":         receivers=["DS CY","HS CY"]
            else:                       receivers=[f"DS {country}"]
        else:
            receivers=[""]

        parent_full=str(base_row.iloc[0]["Parent Offering"])

        for app in new_apps:
            for recv in receivers:
                new_name = build_name_by_convention(
                    naming_convention=naming_convention,
                    parent_offering=parent_full,
                    sr_or_im=sr_or_im,
                    app=app,
                    schedule_suffix=schedule_suffix,
                    require_corp=require_corp,
                    receiver=recv,
                    delivering_tag=delivering_tag,
                    country=country
                )
                
                if new_name in seen:
                    continue
                seen.add(new_name)

                row=base_row.copy()
                row["Name (Child Service Offering lvl 1)"]=new_name
                row["Delivery Manager"]=delivery_manager
                row["Support group"]=support_group
                row["Managed by Group"]=managed_by_group
                
                for c in [c for c in row.columns if "Aliases" in c]:
                    row[c]="-"
                    
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
        raise ValueError("No matching rows.")

    out_dir.mkdir(parents=True,exist_ok=True)
    outfile=out_dir / f"Offerings_NEW_{dt.datetime.now():%Y%m%d_%H%M%S}.xlsx"
    with pd.ExcelWriter(outfile,engine="openpyxl") as w:
        for cc,dfc in sheets.items():
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