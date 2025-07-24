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
        
        sr_or_im = st.radio("Service Type", ["SR", "IM"], horizontal=True)
        
        # Add Business Criticality control
        st.markdown("---")
        business_criticality = st.selectbox(
            "Business Criticality",
            options=["", "1 - most critical", "2 - somewhat critical", "3 - less critical", "4 - not critical"],
            index=0,
            help="Set Business Criticality for all generated offerings. If empty, original values from source files will be preserved."
        )
        
        # Add Lvl2 checkbox
        st.markdown("---")
        use_lvl2 = st.checkbox(
            "Include Level 2 (Child SO lvl2)",
            help="When checked, search in BOTH Child SO lvl1 AND Child SO lvl2 sheets."
        )
        
        if use_lvl2:
            service_type = st.text_input(
                "Service Type (for Lvl2 entries)",
                value="",
                placeholder="e.g., Application issue, Hardware problem",
                help="Optional - This will be added to Lvl2 entries after Prod and before the schedule"
            )
        else:
            service_type = ""
        
        delivery_manager = st.text_input("Delivery Manager", value="")
    
        # Add Approval Required control
        approval_required = st.checkbox(
            "Approval Required",
            value=False,
            help="Set Approval Required to true for all generated offerings. Default is false."
        )
    
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
        else:
            new_parent_offering = ""
            new_parent = ""
    
    with tab3:
        st.subheader("Schedule Settings")
        
        schedule_type = st.checkbox("Custom schedule per period")
        create_multiple_schedules = st.checkbox("Create multiple schedules", help="Generate the same offerings with different schedules")
        
        if not schedule_type and create_multiple_schedules:
            schedule_simple = st.text_area(
                "Schedules", 
                value="", 
                placeholder="Mon-Fri 9-17\nMon-Fri 8-16\nMon-Sun 24/7",
                height=150
            )
            schedule_suffixes = [s.strip() for s in schedule_simple.split('\n') if s.strip()]
        elif schedule_type and not create_multiple_schedules:
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
            
            schedule_suffix = " ".join(schedule_parts) if schedule_parts else ""
            schedule_suffixes = [schedule_suffix] if schedule_suffix else []
        elif schedule_type and create_multiple_schedules:
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
        else:
            schedule_simple = st.text_input("Schedule", value="", placeholder="e.g. Mon-Fri 9-17")
            schedule_suffixes = [schedule_simple] if schedule_simple else []
        
        st.markdown("---")
        
        # Schedule Settings per Country
        use_per_country_schedules = st.checkbox("Use different schedules per country/division", help="Define specific schedules for different countries or divisions")
        schedule_settings_per_country = {}
        
        if use_per_country_schedules:
            st.markdown("**Schedule Settings per Country/Division**")
            
            # Define available countries/divisions - UPDATED with new countries
            available_countries = ["HS PL", "DS PL", "DE", "MD", "UA", "DS CY", "DS RO", "DS TR"]
            
            # Create tabs or columns for different countries
            country_tabs = st.tabs(available_countries)
            
            for idx, country in enumerate(available_countries):
                with country_tabs[idx]:
                    st.markdown(f"**{country} Schedules**")
                    
                    country_schedules = st.text_area(
                        f"Schedules for {country}",
                        value="",
                        placeholder="Mon-Fri 9-17\nMon-Fri 8-16\nMon-Sun 24/7",
                        height=150,
                        key=f"schedule_{country.replace(' ', '_')}"
                    )
                    
                    if country_schedules.strip():
                        # Map DS CY to CY, DS RO to RO, DS TR to TR for backend compatibility
                        if country == "DS CY":
                            schedule_settings_per_country["CY"] = country_schedules.strip()
                        elif country == "DS RO":
                            schedule_settings_per_country["RO"] = country_schedules.strip()
                        elif country == "DS TR":
                            schedule_settings_per_country["TR"] = country_schedules.strip()
                        else:
                            schedule_settings_per_country[country] = country_schedules.strip()
        
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
            # UPDATED country list with RO and TR
            commitment_country = st.selectbox("Country", ["CY", "DE", "MD", "PL", "RO", "TR", "UA"])
            
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
        st.subheader("Support Groups")
        
        # Option to use same group for all countries or different per country
        use_per_country_groups = st.checkbox("Use different support groups per country", value=False)
        
        if not use_per_country_groups:
            # Single support group for all countries
            st.markdown("#### Global Support Groups")
            support_group = st.text_input("Support Group", value="", help="Same support group for all countries")
            managed_by_group = st.text_input(
                "Managed by Group", 
                value="",
                help="Optional - if empty, will use Support Group value"
            )
            # Store as single values for backward compatibility
            support_groups_per_country = {}
            managed_by_groups_per_country = {}
        else:
            # Per-country support groups
            st.markdown("#### Support Groups by Country")
            st.info("Select countries to configure specific support groups. Unchecked countries will use empty support groups.")
            
            # UPDATED countries list with new DS-only countries
            countries = ["HS PL", "DS PL", "DE", "UA", "MD", "DS CY", "DS RO", "DS TR"]
            support_groups_per_country = {}
            managed_by_groups_per_country = {}
            
            # Create columns for better layout
            cols = st.columns(2)
            
            for i, country in enumerate(countries):
                col_idx = i % 2
                with cols[col_idx]:
                    country_enabled = st.checkbox(f"Configure {country}", key=f"enable_{country}")
                    
                    if country_enabled:
                        # Special handling for DE - multiple support groups
                        if country == "DE":
                            num_de_groups = st.number_input(
                                f"Number of Support Groups for DE", 
                                min_value=1, 
                                max_value=5, 
                                value=1, 
                                key=f"num_groups_DE",
                                help="DE can have multiple support groups for the same offerings"
                            )
                            
                            de_support_groups = []
                            de_managed_groups = []
                            
                            for group_idx in range(num_de_groups):
                                st.markdown(f"**DE Support Group {group_idx + 1}**")
                                de_support_group = st.text_input(
                                    f"Support Group {group_idx + 1}",
                                    value="",
                                    key=f"support_DE_{group_idx}",
                                    placeholder=f"e.g., DE IT Support Team {group_idx + 1}"
                                )
                                de_managed_group = st.text_input(
                                    f"Managed by Group {group_idx + 1}",
                                    value="",
                                    key=f"managed_DE_{group_idx}",
                                    placeholder=f"Optional - defaults to Support Group if empty",
                                    help="If empty, will use the Support Group value"
                                )
                                
                                if de_support_group:
                                    de_support_groups.append(de_support_group)
                                    de_managed_groups.append(de_managed_group if de_managed_group else de_support_group)
                            
                            # Join multiple groups with newlines
                            support_groups_per_country[country] = "\n".join(de_support_groups) if de_support_groups else ""
                            managed_by_groups_per_country[country] = "\n".join(de_managed_groups) if de_managed_groups else ""
                        else:
                            # Standard single support group for other countries
                            # Map display names to backend keys
                            backend_key = country
                            if country == "DS CY":
                                backend_key = "CY"
                            elif country == "DS RO":
                                backend_key = "RO"
                            elif country == "DS TR":
                                backend_key = "TR"
                            
                            support_groups_per_country[backend_key] = st.text_input(
                                f"Support Group",
                                value="",
                                key=f"support_{country}",
                                placeholder=f"e.g., {country} IT Support"
                            )
                            managed_by_groups_per_country[backend_key] = st.text_input(
                                f"Managed by Group",
                                value="",
                                key=f"managed_{country}",
                                placeholder=f"Optional - defaults to Support Group if empty",
                                help="If empty, will use the Support Group value"
                            )
                    else:
                        # Map display names to backend keys for empty values too
                        backend_key = country
                        if country == "DS CY":
                            backend_key = "CY"
                        elif country == "DS RO":
                            backend_key = "RO"
                        elif country == "DS TR":
                            backend_key = "TR"
                        
                        support_groups_per_country[backend_key] = ""
                        managed_by_groups_per_country[backend_key] = ""
            
            # Set global variables to empty for backward compatibility
            support_group = ""
            managed_by_group = ""
    
    with tab6:
        st.subheader("Select naming convention:")
        
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
        elif all_selected == 0:
            st.info("📌 Standard naming will be used")
        
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
        
        # Custom Depend On setting
        st.markdown("### Service Offerings | Depend On")
        use_custom_depend_on = st.checkbox("Use custom 'Service Offerings | Depend On' value", value=False, 
                                          help="Override automatic generation with a custom value for all rows")
        
        if use_custom_depend_on:
            col1, col2 = st.columns([1, 2])
            with col1:
                # UPDATED options with DS-only countries
                depend_on_prefix = st.selectbox(
                    "Select Prefix",
                    options=["HS PL", "DS PL", "HS DE", "DS DE", "DS UA", "DS MD", "DS CY", "DS RO", "DS TR", "Global"],
                    index=0,
                    help="Choose the service prefix"
                )
            
            with col2:
                # Show which apps will be used automatically
                if new_apps:
                    if len(new_apps) == 1:
                        st.text_input(
                            "Application Name",
                            value=new_apps[0],
                            disabled=True,
                            help=f"Will automatically use: {new_apps[0]}"
                        )
                    else:
                        st.text_area(
                            "Application Names",
                            value="\n".join(new_apps),
                            disabled=True,
                            height=80,
                            help=f"Will automatically use each app: {', '.join(new_apps)}"
                        )
                    app_names_display = new_apps
                else:
                    st.text_input(
                        "Application Name",
                        value="(no apps specified)",
                        disabled=True,
                        help="Add applications in Basic tab to see them here"
                    )
                    app_names_display = ["(no app)"]
            
            # Construct the custom depend on value
            if depend_on_prefix == "Global":
                prefix_tag = "Global Prod"
            else:
                prefix_tag = f"{depend_on_prefix} Prod"
            
            # Show preview(s) for the custom depend on values
            if new_apps:
                if len(new_apps) == 1:
                    custom_depend_on_value = f"[{prefix_tag}] {new_apps[0]}"
                    st.info(f"Preview: `{custom_depend_on_value}`")
                else:
                    st.info("Preview for each app:")
                    for app in new_apps:
                        st.text(f"• `[{prefix_tag}] {app}`")
                    # Store prefix only - the backend will handle app names automatically
                    custom_depend_on_value = f"[{prefix_tag}]"
            else:
                custom_depend_on_value = f"[{prefix_tag}]"
                st.info(f"Preview: `{custom_depend_on_value}`")
        else:
            custom_depend_on_value = ""
        
        # Aliases
        st.markdown("### Aliases")
        aliases_on = st.checkbox("Enable Aliases", value=False)
        
        # Initialize variables
        use_per_country_aliases = False
        aliases_per_country = {}
        
        if aliases_on:
            # Option for global or per-country aliases
            use_per_country_aliases = st.checkbox("Use different aliases per country/division", 
                                                 value=False, 
                                                 help="Define specific aliases for different countries or divisions")
            
            if use_per_country_aliases:
                st.markdown("**Aliases per Country/Division**")
                st.info("Configure aliases for specific countries/divisions. Unconfigured countries will use empty aliases.")
                
                # Define available countries/divisions - UPDATED with new countries
                available_countries = ["HS PL", "DS PL", "DE", "MD", "UA", "DS CY", "DS RO", "DS TR"]
                aliases_per_country = {}
                
                # Create tabs for different countries
                alias_country_tabs = st.tabs(available_countries)
                
                for idx, country in enumerate(available_countries):
                    with alias_country_tabs[idx]:
                        st.markdown(f"**{country} Aliases**")
                        
                        country_alias = st.text_input(
                            f"Alias Value for {country}",
                            value="",
                            key=f"alias_{country.replace(' ', '_')}",
                            placeholder=f"e.g., {country}_ALIAS",
                            help=f"Alias value to use specifically for {country}"
                        )
                        
                        if country_alias.strip():
                            # Map display names to backend keys
                            backend_key = country
                            if country == "DS CY":
                                backend_key = "CY"
                            elif country == "DS RO":
                                backend_key = "RO"
                            elif country == "DS TR":
                                backend_key = "TR"
                            
                            aliases_per_country[backend_key] = country_alias.strip()
                
                # Set global alias to empty when using per-country
                aliases_value = ""
            else:
                # Global alias value
                aliases_value = st.text_input(
                    "Global Alias Value",
                    value="",
                    help="Enter the value to use for aliases (same for all countries)"
                )
        else:
            aliases_value = ""

# Add this with your other variable definitions
if 'approval_required' not in locals():
    approval_required = False

st.markdown("---")

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
                        aliases_per_country=aliases_per_country if use_per_country_aliases else {},
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
                        use_new_parent=use_new_parent,
                        new_parent_offering=new_parent_offering,
                        new_parent=new_parent,
                        keywords_excluded=keywords_excluded if not use_new_parent else "",
                        use_lvl2=use_lvl2 if 'use_lvl2' in locals() else False,
                        service_type_lvl2=service_type if 'service_type' in locals() else "",
                        # Add per-country support groups
                        support_groups_per_country=support_groups_per_country if use_per_country_groups else {},
                        managed_by_groups_per_country=managed_by_groups_per_country if use_per_country_groups else {},
                        # Add per-country schedule settings
                        schedule_settings_per_country=schedule_settings_per_country if use_per_country_schedules else {},
                        # Add custom depend on value
                        use_custom_depend_on=use_custom_depend_on if 'use_custom_depend_on' in locals() else False,
                        custom_depend_on_value=custom_depend_on_value if 'custom_depend_on_value' in locals() else "",
                        # Add business criticality
                        business_criticality=business_criticality,
                        # Add approval required
                        approval_required=approval_required
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
        Service Offerings Generator v3.8 | Support
    </div>
    """,
    unsafe_allow_html=True
)