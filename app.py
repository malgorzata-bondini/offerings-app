import streamlit as st
import pandas as pd
from pathlib import Path
import shutil
import tempfile
from generator_core import run_generator

st.set_page_config(page_title="Service Offerings Generator", layout="wide")

st.title("🔧 Service Offerings Generator")
st.markdown("---")

col1, col2 = st.columns([1, 1])

with col1:
    st.header("📁 Input Files")
    uploaded_files = st.file_uploader(
        "Upload ALL_Service_Offering Excel files",
        type=["xlsx"],
        accept_multiple_files=True,
        help="Upload one or more Excel files with naming pattern: ALL_Service_Offering_*.xlsx"
    )
    
    if uploaded_files:
        st.success(f"✅ Uploaded {len(uploaded_files)} file(s)")
        for file in uploaded_files:
            st.text(f"• {file.name}")

with col2:
    st.header("⚙️ Configuration")
    
    tab1, tab2, tab3, tab4, tab5, tab6, tab7 = st.tabs(["Basic", "New Parent Offering", "Schedule", "Service Commitments", "Groups", "Naming", "Advanced"])
    
    with tab1:
        st.subheader("Basic Settings")
        
        # Existing keyword filtering section
        st.info("""
        **Keywords filtering:**
        - Line-separated keywords: OR logic (any match)
        - Comma-separated keywords: AND logic (all must match)
        - When "Include Level 2" is checked: Searches in BOTH Child SO lvl1 AND lvl2 sheets
        - When unchecked: Searches only in Child SO lvl1 sheet
        """)
        
        keywords_parent = st.text_area(
            "Keywords in Parent Offering",
            value="",
            placeholder="Enter keywords (one per line for OR, comma-separated for AND)",
            help="Filter by Parent Offering column"
        )
        
        keywords_child = st.text_area(
            "Keywords in Name (Child Service Offering lvl 1)",
            value="",
            placeholder="Enter keywords (one per line for OR, comma-separated for AND)",
            help="Filter by Child Service Offering Name column"
        )
        
        keywords_excluded = st.text_area(
            "Keywords to Exclude",
            value="",
            placeholder="Enter keywords to exclude (one per line for OR, comma-separated for AND)",
            help="Exclude rows containing these keywords in either Parent Offering or Child Name"
        )
        
        new_apps = st.text_area(
            "Applications (one per line or comma-separated)",
            value="",
            help="Optional - Enter application names. If empty, offerings will be created without application names"
        ).strip().split('\n')
        new_apps = [a.strip() for a in new_apps if a.strip()]
        
        sr_or_im = st.radio("Service Type", ["SR", "IM"], horizontal=True, help="This will be used in naming for both lvl1 and lvl2 entries")
        
        # Add Lvl2 checkbox
        st.markdown("---")
        use_lvl2 = st.checkbox(
            "Include Level 2 (Child SO lvl2)",
            help="When checked, search in BOTH Child SO lvl1 AND Child SO lvl2 sheets. When unchecked, only search lvl1."
        )
        
        if use_lvl2:
            service_type = st.text_input(
                "Service Type (for Lvl2 entries)",
                value="",
                placeholder="e.g., Application issue, Hardware problem",
                help="Optional - This will be added to Lvl2 entries after Prod and before the schedule"
            )
            st.info("📌 Will search in BOTH levels: Child SO lvl1 AND Child SO lvl2. For Lvl2 entries, the selected SR/IM will be added to Parent Offering if not already present.")
        else:
            service_type = ""
            st.info("📌 Will search only in Child SO lvl1")
        
        delivery_manager = st.text_input("Delivery Manager", value="")
    
    with tab2:
        st.subheader("Direct Parent Offering Selection")
        
        use_new_parent = st.checkbox(
            "Use specific parent offering (instead of parent keyword search)",
            help="When checked, you can enter a completely new Parent Offering and Parent name"
        )
        
        if use_new_parent:
            st.info("📝 Enter the exact Parent Offering and Parent values to use")
            
            col_parent1, col_parent2 = st.columns(2)
            with col_parent1:
                new_parent_offering = st.text_input(
                    "New Parent Offering",
                    value="",
                    placeholder="e.g., [Parent HS PL IT] Software assistance",
                    help="Enter the complete Parent Offering text"
                )
            
            with col_parent2:
                new_parent = st.text_input(
                    "New Parent",
                    value="",
                    placeholder="e.g., PL IT Support",
                    help="Enter the Parent value"
                )
            
            # Show preview
            if new_parent_offering and new_parent:
                st.success(f"📋 Will use: Parent Offering = '{new_parent_offering}', Parent = '{new_parent}'")
                # Note about Level 2 behavior
                st.info("📌 If 'Include Level 2' is checked in Basic tab, SR/IM will be added to Level 2 entries")
        else:
            new_parent_offering = ""
            new_parent = ""
            
        # Update keyword filtering visibility based on use_new_parent
        if use_new_parent:
            st.info("🔒 Keyword filtering in Basic tab is disabled when using specific parent offering")
            keywords_parent = ""
            keywords_child = ""
            keywords_excluded = ""
        elif use_lvl2:
            st.info("""
            📝 Note: Keywords will search in BOTH "Child SO lvl1" AND "Child SO lvl2" sheets.
            For Lvl2 entries, SR/IM will be automatically added to Parent Offering when building names.
            """)
    
    with tab3:
        st.subheader("Schedule Settings")
        
        schedule_type = st.checkbox("Custom schedule per period")
        create_multiple_schedules = st.checkbox("Create multiple schedules", help="Generate the same offerings with different schedules")
        
        if not schedule_type and create_multiple_schedules:
            st.info("📅 Enter multiple schedules (one per line)")
            schedule_simple = st.text_area(
                "Schedules", 
                value="", 
                placeholder="Mon-Fri 9-17\nMon-Fri 8-16\nMon-Sun 24/7",
                height=150
            )
            schedule_suffixes = [s.strip() for s in schedule_simple.split('\n') if s.strip()]
            
            # Show preview of all schedules
            if schedule_suffixes:
                st.success(f"📋 Will create {len(schedule_suffixes)} offering(s) with different schedules:")
                for i, sched in enumerate(schedule_suffixes, 1):
                    st.text(f"  {i}. {sched}")
        elif schedule_type and not create_multiple_schedules:
            st.info("📅 Enter schedule periods (e.g., Mon-Thu 9-17, Fri 9-16, Sat 8-12)")
            
            # Allow up to 5 schedule periods
            schedule_parts = []
            num_periods = st.number_input("Number of periods", min_value=1, max_value=5, value=2)
            
            for i in range(num_periods):
                st.markdown(f"**Period {i+1}**")
                col1, col2 = st.columns([2, 1])
                with col1:
                    period = st.text_input(
                        f"Days", 
                        value="", 
                        placeholder="e.g. Mon-Thu or Fri or Sat-Sun",
                        key=f"schedule_period_{i}",
                        label_visibility="collapsed"
                    )
                with col2:
                    hours = st.text_input(
                        f"Hours", 
                        value="", 
                        placeholder="e.g. 9-17",
                        key=f"schedule_hours_{i}",
                        label_visibility="collapsed"
                    )
                
                if period and hours:
                    schedule_parts.append(f"{period} {hours}")
            
            # Join all non-empty schedule parts
            schedule_suffix = " ".join(schedule_parts) if schedule_parts else ""
            schedule_suffixes = [schedule_suffix] if schedule_suffix else []
            
            # Show preview
            if schedule_suffix:
                st.success(f"📋 Schedule: **{schedule_suffix}**")
        elif schedule_type and create_multiple_schedules:
            st.info("📅 Enter multiple custom schedule periods")
            num_schedules = st.number_input("Number of schedules", min_value=1, max_value=5, value=2)
            schedule_suffixes = []
            
            for sched_idx in range(num_schedules):
                st.markdown(f"### Schedule {sched_idx + 1}")
                schedule_parts = []
                num_periods = st.number_input(f"Number of periods for schedule {sched_idx + 1}", min_value=1, max_value=5, value=2, key=f"num_periods_{sched_idx}")
                
                for i in range(num_periods):
                    st.markdown(f"**Period {i+1}**")
                    col1, col2 = st.columns([2, 1])
                    with col1:
                        period = st.text_input(
                            f"Days", 
                            value="", 
                            placeholder="e.g. Mon-Thu or Fri or Sat-Sun",
                            key=f"schedule_period_{sched_idx}_{i}",
                            label_visibility="collapsed"
                        )
                    with col2:
                        hours = st.text_input(
                            f"Hours", 
                            value="", 
                            placeholder="e.g. 9-17",
                            key=f"schedule_hours_{sched_idx}_{i}",
                            label_visibility="collapsed"
                        )
                    
                    if period and hours:
                        schedule_parts.append(f"{period} {hours}")
                
                if schedule_parts:
                    schedule_suffix = " ".join(schedule_parts)
                    schedule_suffixes.append(schedule_suffix)
                    st.success(f"📋 Schedule {sched_idx + 1}: **{schedule_suffix}**")
        else:
            schedule_simple = st.text_input("Schedule", value="", placeholder="e.g. Mon-Fri 9-17")
            schedule_suffixes = [schedule_simple] if schedule_simple else []
        
        st.markdown("---")
        col_rsp, col_rsl = st.columns(2)
        with col_rsp:
            rsp_duration = st.text_input("RSP Duration", value="")
        with col_rsl:
            rsl_duration = st.text_input("RSL Duration", value="")
    
    with tab4:
        st.subheader("Service Commitments")
        
        use_custom_commitments = st.checkbox("Define custom Service Commitments", help="If unchecked, commitments will be copied from source files")
        
        if use_custom_commitments:
            commitment_country = st.selectbox("Country", ["CY", "DE", "MD", "PL", "UA"])
            
            st.markdown("### Service Commitments Configuration")
            
            # Initialize commitment lines list
            commitment_lines = []
            
            # Create multiple commitment entries
            num_commitments = st.number_input("Number of commitment lines", min_value=1, max_value=10, value=2)
            
            for i in range(num_commitments):
                st.markdown(f"#### Commitment Line {i+1}")
                col1, col2, col3, col4 = st.columns([1, 1, 2, 1])
                
                with col1:
                    line_type = st.selectbox(f"Type", ["RSP", "RSL"], key=f"commit_type_{i}")
                
                with col2:
                    priority = st.selectbox(f"Priority", ["P1", "P2", "P3", "P4", "P1-P2", "P3-P4", "P1-P4"], key=f"commit_priority_{i}")
                
                with col3:
                    schedule = st.text_input(f"Schedule", placeholder="e.g. Mon-Fri 6-21", key=f"commit_schedule_{i}")
                
                with col4:
                    time = st.text_input(f"Time", placeholder="e.g. 2h, 1d", key=f"commit_time_{i}")
                
                if schedule and time:
                    line = f"[{commitment_country}] SLA {sr_or_im} {line_type} {schedule} {priority} {time}"
                    commitment_lines.append(line)
                    
                    # Add OLA for SR and RSL
                    if sr_or_im == "SR" and line_type == "RSL":
                        ola_line = f"[{commitment_country}] OLA {sr_or_im} RSL {schedule} {priority} {time}"
                        commitment_lines.append(ola_line)
            
            # Show preview
            if commitment_lines:
                st.markdown("### Preview")
                st.code("\n".join(commitment_lines))
                
            # Store the custom commitments string
            custom_commitments_str = "\n".join(commitment_lines) if commitment_lines else ""
        else:
            custom_commitments_str = ""
            commitment_country = None
    
    with tab5:
        support_group = st.text_input("Support Group", value="")
        managed_by_group = st.text_input(
            "Managed by Group", 
            value="",
            help="Optional - if empty, will use Support Group value"
        )
    
    with tab6:
        st.subheader("Select naming convention to use:")
        
        st.info("⚠️ Important: These checkboxes only control which naming convention is applied to ALL matching entries. They do NOT filter which entries are processed.")

        
        # Create a column to ensure vertical layout - ALPHABETICAL ORDER
        col = st.container()
        with col:
            # List in alphabetical order
            require_corp = st.checkbox("CORP", help="Apply CORP naming to ALL matching entries")
            require_recp = st.checkbox("RecP", help="Apply RecP naming to ALL matching entries")
            require_corp_dedicated = st.checkbox("CORP Dedicated Services", help="Apply CORP Dedicated naming to ALL matching entries")
            require_corp_it = st.checkbox("CORP IT", help="Apply CORP IT naming to ALL matching entries")
            special_dak = st.checkbox("DAK (Business Services)", help="Apply Business Services naming to ALL matching entries")
            special_hr = st.checkbox("HR", help="Apply HR naming to ALL matching entries")
            special_it = st.checkbox("IT", help="Apply IT naming to ALL matching entries")
            special_medical = st.checkbox("Medical", help="Apply Medical naming to ALL matching entries")

        
        # Ensure only one is selected
        all_selected = sum([require_corp, require_recp, special_it, special_hr, special_medical, special_dak, require_corp_it, require_corp_dedicated])
        if all_selected > 1:
            st.error("⚠️ Please select only one naming type")
            # Reset all to handle multiple selection
            require_corp = require_recp = special_it = special_hr = special_medical = special_dak = require_corp_it = require_corp_dedicated = False
        elif all_selected == 0:
            st.info("📌 No special naming selected - will use standard naming for ALL matching entries (including RecP, IT, HR, Medical, etc.)")
        
        if require_corp or require_recp or require_corp_it or require_corp_dedicated:
            delivering_tag = st.text_input(
                "Who delivers the service", 
                value="",
                help="E.g. HS PL, DS DE, IT, Finance, etc."
            )
        else:
            delivering_tag = ""
    
    with tab7:
        # Global settings
        st.markdown("### Global")
        global_prod = st.checkbox("Global Prod", value=False)
        
        # Aliases
        st.markdown("### Aliases")
        aliases_on = st.checkbox("Enable Aliases", value=False)  # Default to False
        
        if aliases_on:
            aliases_value = st.text_input(
                "Alias Value",
                value="",
                help="Enter the value to use for aliases"
            )
        else:
            aliases_value = ""

st.markdown("---")

# Show naming examples based on selections
with st.expander("📋 Naming Convention Examples"):
    if 'use_lvl2' in locals() and use_lvl2:
        st.markdown("**Level 1 Example:**")
        st.code("[SR HS PL IT] Software assistance Outlook Prod Mon-Fri 9-17")
        st.markdown("**Level 2 Example:**")
        st.code("[IM HS PL IT] Software incident solving EMR 2.0 Prod Application issue Mon-Sun 24/7")
        st.markdown("From parent: `[Parent HS PL IT] Software incident solving` → IM added automatically")
        if 'service_type' in locals() and service_type:
            st.info(f"Service Type '{service_type}' will be added to Lvl2 entries after Prod")
        st.markdown("**📂 Output:** Separate sheets like 'PL lvl1' and 'PL lvl2' in the same file")
    elif 'require_corp' in locals() and require_corp:
        st.markdown("**CORP Example:**")
        st.code("[SR HS PL CORP DS DE] Software assistance Outlook Prod Mon-Fri 8-17")
        st.info("Applied to ALL matching entries using CORP naming style")
    elif 'require_corp_it' in locals() and require_corp_it:
        st.markdown("**CORP IT Example:**")
        st.code("[SR HS PL CORP HS PL IT] Software assistance MS Outlook Prod Mon-Fri 8-17")
        st.info("Applied to ALL matching entries using CORP IT naming style")
    elif 'require_corp_dedicated' in locals() and require_corp_dedicated:
        st.markdown("**CORP Dedicated Services Example:**")
        st.code("[SR HS PL CORP HS DE Dedicated Services] HelpMe announcement creation ServiceNow HelpMe Prod Mon-Fri 9-17")
        st.info("Applied to ALL matching entries using CORP Dedicated Services naming style")
    elif 'require_recp' in locals() and require_recp:
        st.markdown("**RecP Example:**")
        st.code("[IM HS PL CORP HS PL IT] Software incident solving Active Directory Prod Mon-Fri 6-21")
        st.info("Applied to ALL matching entries using RecP naming style")
    elif 'special_it' in locals() and special_it:
        st.markdown("**IT Example:**")
        st.code("[SR HS PL IT] Software assistance Outlook Prod Mon-Fri 8-17")
        st.info("Applied to ALL matching entries using IT naming style")
    elif 'special_hr' in locals() and special_hr:
        st.markdown("**HR Example:**")
        st.code("[SR HS PL HR] Software assistance Outlook Prod Mon-Fri 8-17")
        st.info("Applied to ALL matching entries using HR naming style")
    elif 'special_medical' in locals() and special_medical:
        st.markdown("**Medical Example:**")
        st.code("[SR HS PL Medical] TeleCentrum medical procedures and quality of care Mon-Fri 7-20")
        st.info("Applied to ALL matching entries using Medical naming style")
    elif 'special_dak' in locals() and special_dak:
        st.markdown("**DAK (Business Services) Example:**")
        st.code("[SR HS PL Business Services] Product modifications Service removal Mon-Fri 8-17")
        st.info("Applied to ALL matching entries using Business Services naming style")
    else:
        st.markdown("**Standard Example:**")
        st.code("[SR HS PL Permissions] Granting permissions to application Outlook Prod Mon-Fri 9-17")
        st.markdown("From parent: `[Parent HS PL Permissions] Granting permissions to application`")
        st.info("Standard naming is applied to ALL matching entries (no special naming selected)")

# MODIFY THIS VALIDATION SECTION
if st.button("🚀 Generate Service Offerings", type="primary", use_container_width=True):
    if not uploaded_files:
        st.error("⚠️ Please upload at least one Excel file")
    elif use_new_parent and (not new_parent_offering or not new_parent):
        st.error("⚠️ When using specific parent offering, both 'New Parent Offering' and 'New Parent' must be filled")
    elif not use_new_parent and not keywords_parent and not keywords_child:
        st.error("⚠️ Please enter at least one keyword in either Parent Offering or Child Service Offering, or use specific parent offering")
    elif 'schedule_suffixes' not in locals() or not schedule_suffixes or not any(schedule_suffixes):
        st.error("⚠️ Please configure at least one schedule")
    elif all_selected > 1:
        st.error("⚠️ Please select only one naming type")
    else:
        try:
            with tempfile.TemporaryDirectory() as temp_dir:
                src_dir = Path(temp_dir) / "input"
                out_dir = Path(temp_dir) / "output"
                src_dir.mkdir(exist_ok=True)
                
                for uploaded_file in uploaded_files:
                    file_path = src_dir / uploaded_file.name
                    with open(file_path, "wb") as f:
                        f.write(uploaded_file.getbuffer())
                
                with st.spinner("🔄 Generating service offerings..."):
                    result_file = run_generator(
                        keywords_parent=keywords_parent if not use_new_parent else "",
                        keywords_child=keywords_child if not use_new_parent else "",
                        new_apps=new_apps,
                        schedule_suffixes=schedule_suffixes,
                        delivery_manager=delivery_manager,
                        global_prod=global_prod,
                        rsp_duration=rsp_duration,
                        rsl_duration=rsl_duration,
                        sr_or_im=sr_or_im,
                        require_corp=require_corp,
                        require_recp=require_recp if 'require_recp' in locals() else False,
                        delivering_tag=delivering_tag,
                        support_group=support_group,
                        managed_by_group=managed_by_group,
                        aliases_on=aliases_on,
                        aliases_value=aliases_value,
                        src_dir=src_dir,
                        out_dir=out_dir,
                        special_it=special_it if 'special_it' in locals() else False,
                        special_hr=special_hr if 'special_hr' in locals() else False,
                        special_medical=special_medical if 'special_medical' in locals() else False,
                        special_dak=special_dak if 'special_dak' in locals() else False,
                        use_custom_commitments=use_custom_commitments if 'use_custom_commitments' in locals() else False,
                        custom_commitments_str=custom_commitments_str if 'custom_commitments_str' in locals() else "",
                        commitment_country=commitment_country if 'commitment_country' in locals() else None,
                        require_corp_it=require_corp_it if 'require_corp_it' in locals() else False,
                        require_corp_dedicated=require_corp_dedicated if 'require_corp_dedicated' in locals() else False,
                        # ADD THESE NEW PARAMETERS
                        use_new_parent=use_new_parent,
                        new_parent_offering=new_parent_offering,
                        new_parent=new_parent,
                        keywords_excluded=keywords_excluded if not use_new_parent else "",
                        # Add Lvl2 parameters
                        use_lvl2=use_lvl2 if 'use_lvl2' in locals() else False,
                        service_type_lvl2=service_type if 'service_type' in locals() else ""
                    )
                
                st.success("✅ Service offerings generated successfully!")
                
                with open(result_file, "rb") as f:
                    st.download_button(
                        label="📥 Download Generated File",
                        data=f.read(),
                        file_name=result_file.name,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )
                
                st.info(f"Generated file: {result_file.name}")
                if use_lvl2:
                    st.info("📊 The file contains separate sheets for each level (e.g., 'PL lvl1', 'PL lvl2')")
                else:
                    st.info("📊 The file contains sheets named by country with lvl1 suffix (e.g., 'PL lvl1')")
                
        except ValueError as e:
            if "duplicate offering" in str(e).lower():
                st.error(f"❌ {str(e)}")
            else:
                st.error(f"❌ Error: {str(e)}")
        except Exception as e:
            st.error(f"❌ Error: {str(e)}")
            st.exception(e)

st.markdown("---")
st.markdown(
    """
    <div style='text-align: center; color: gray;'>
        Service Offerings Generator v3.3 | Fixed: Naming checkboxes now only control naming style, not filtering
    </div>
    """,
    unsafe_allow_html=True
)