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
    
    tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs(["Basic", "Schedule", "Service Commitments", "Groups", "Naming", "Advanced"])
    
    with tab1:
        st.subheader("Basic Settings")
        
        # ADD THIS SECTION - Direct Parent Offering Selection
        st.markdown("### Direct Parent Offering Selection")
        
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
        else:
            new_parent_offering = ""
            new_parent = ""
        
        st.markdown("---")
        
        # Existing keyword filtering section
        if not use_new_parent:
            st.info("""
            **Keywords filtering:**
            - Line-separated keywords: OR logic (any match)
            - Comma-separated keywords: AND logic (all must match)
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
        else:
            keywords_parent = ""
            keywords_child = ""
            st.info("🔒 Keyword filtering is disabled when using specific parent offering")
        
        new_apps = st.text_area(
            "Applications (one per line or comma-separated)",
            value="",
            help="Optional - Enter application names. If empty, offerings will be created without application names"
        ).strip().split('\n')
        new_apps = [a.strip() for a in new_apps if a.strip()]
        
        sr_or_im = st.radio("Service Type", ["SR", "IM"], horizontal=True)
        
        delivery_manager = st.text_input("Delivery Manager", value="")
    
    with tab2:
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
    
    with tab3:
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
    
    with tab4:
        support_group = st.text_input("Support Group", value="")
        managed_by_group = st.text_input(
            "Managed by Group", 
            value="",
            help="Optional - if empty, will use Support Group value"
        )
    
    with tab5:
        st.subheader("Select one of the following:")

        
        # Create a column to ensure vertical layout - ALPHABETICAL ORDER
        col = st.container()
        with col:
            # List in alphabetical order
            require_corp = st.checkbox("CORP")
            require_recp = st.checkbox("RecP")
            require_corp_dedicated = st.checkbox("CORP Dedicated Services")
            require_corp_it = st.checkbox("CORP IT")
            special_dak = st.checkbox("DAK (Business Services)")
            special_hr = st.checkbox("HR")
            special_it = st.checkbox("IT")
            special_medical = st.checkbox("Medical")

        
        # Ensure only one is selected
        all_selected = sum([require_corp, require_recp, special_it, special_hr, special_medical, special_dak, require_corp_it, require_corp_dedicated])
        if all_selected > 1:
            st.error("⚠️ Please select only one naming type")
            # Reset all to handle multiple selection
            require_corp = require_recp = special_it = special_hr = special_medical = special_dak = require_corp_it = require_corp_dedicated = False
        
        if require_corp or require_recp or require_corp_it or require_corp_dedicated:
            delivering_tag = st.text_input(
                "Who delivers the service", 
                value="",
                help="E.g. HS PL, DS DE, IT, Finance, etc."
            )
        else:
            delivering_tag = ""
    
    with tab6:
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
    if 'require_corp' in locals() and require_corp:
        st.markdown("**CORP Example:**")
        st.code("[SR HS PL CORP DS DE] Software assistance Outlook Prod Mon-Fri 8-17")
    elif 'require_corp_it' in locals() and require_corp_it:
        st.markdown("**CORP IT Example:**")
        st.code("[SR HS PL CORP HS PL IT] Software assistance MS Outlook Prod Mon-Fri 8-17")
    elif 'require_corp_dedicated' in locals() and require_corp_dedicated:
        st.markdown("**CORP Dedicated Services Example:**")
        st.code("[SR HS PL CORP HS DE Dedicated Services] HelpMe announcement creation ServiceNow HelpMe Prod Mon-Fri 9-17")
    elif 'require_recp' in locals() and require_recp:
        st.markdown("**RecP Example:**")
        st.code("[IM HS PL CORP HS PL IT] Software incident solving Active Directory Prod Mon-Fri 6-21")
    elif 'special_it' in locals() and special_it:
        st.markdown("**IT Example:**")
        st.code("[SR HS PL IT] Software assistance Outlook Prod Mon-Fri 8-17")
    elif 'special_hr' in locals() and special_hr:
        st.markdown("**HR Example:**")
        st.code("[SR HS PL HR] Software assistance Outlook Prod Mon-Fri 8-17")
    elif 'special_medical' in locals() and special_medical:
        st.markdown("**Medical Example:**")
        st.code("[SR HS PL Medical] TeleCentrum medical procedures and quality of care Mon-Fri 7-20")
    elif 'special_dak' in locals() and special_dak:
        st.markdown("**DAK (Business Services) Example:**")
        st.code("[SR HS PL Business Services] Product modifications Service removal Mon-Fri 8-17")
    else:
        st.markdown("**Standard Example:**")
        st.code("[SR HS PL Permissions] Granting permissions to application Outlook Prod Mon-Fri 9-17")
        st.markdown("From parent: `[Parent HS PL Permissions] Granting permissions to application`")

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
                        keywords_parent=keywords_parent,
                        keywords_child=keywords_child,
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
                        new_parent=new_parent
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
        Service Offerings Generator v3.0 | With Enhanced Naming Logic
    </div>
    """,
    unsafe_allow_html=True
)