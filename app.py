import streamlit as st
import pandas as pd
from pathlib import Path
import shutil
import tempfile
from generator_core import run_generator

# Add pluralization function AFTER imports
def get_plural_form_preview(word):
    """Get plural form for preview"""
    plural_map = {
        "Laptop": "Laptops", "Desktop": "Desktops", "Docking station": "Docking stations",
        "Printer": "Printers", "Barcode printer": "Barcode printers", "Barcode scanner": "Barcode scanners",
        "Display": "Displays", "Deskphone": "Deskphones", "Smartphone": "Smartphones",
        "Mouse": "Mouses", "Keyboard": "Keyboards", "Headset": "Headsets", "Tablet": "Tablets",
        "Audio equipment": "Audio equipment", "Video surveillance": "Video surveillance",
        "UPS": "UPS", "External webcam": "External webcams", "Projector": "Projectors",
        "External storage device": "External storage devices", "Microphone": "Microphones",
        "Other hardware": "Other hardware", "Server": "Servers", "Router": "Routers",
        "Switch": "Switches", "Firewall": "Firewalls", "Access point": "Access points",
        "Scanner": "Scanners", "Webcam": "Webcams", "Camera": "Cameras", "Monitor": "Monitors",
        "Speaker": "Speakers", "Cable": "Cables", "Adapter": "Adapters"
    }
    return plural_map.get(word, word)

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
        help="Upload one or more Excel files with this naming pattern: ALL_Service_Offering_*.xlsx"
    )
    
    if uploaded_files:
        st.success(f"✅ Uploaded {len(uploaded_files)} file(s)")
        for file in uploaded_files:
            st.text(f"• {file.name}")

with col2:
    st.header("⚙️ Configuration")
    
    tab1, tab2, tab3, tab4, tab5, tab6, tab7 = st.tabs(["Basic", "New Parent Offering", "Schedule", "Service Commitments", "Groups", "Naming", "Other settings"])
    
    with tab1:
        st.subheader("Basic Settings")
        
        keywords_parent = st.text_area(
            "Keywords in Parent Offering",
            value="",
            placeholder="Enter keywords (one per line for OR, comma separated for AND)",
            help="Filter by Parent Offering column in Excel"
        )
        
        keywords_child = st.text_area(
            "Keywords in Name (Child Service Offering lvl 1)",
            value="",
            placeholder="Enter keywords (one per line for OR, comma-separated for AND)",
            help="Filter by Child Service Offering Name column in Excel"
        )
        
        keywords_excluded = st.text_area(
            "Keywords to Exclude",
            value="",
            placeholder="Enter keywords to exclude from the search(one per line for OR, comma-separated for AND)",
            help="Exclude rows containing these keywords in either Parent Offering or Child Name column"
        )
        
        new_apps = st.text_area(
            "Applications/Other (one per line or comma-separated)",
            value="",
            help="It's optional - enter application names. If empty, offerings will be created without the names"
        ).strip().split('\n')
        new_apps = [a.strip() for a in new_apps if a.strip()]
        
        sr_or_im = st.radio("Service Type", ["SR", "IM"], horizontal=True)
        
        # Add Prod checkbox right after Service Type
        add_prod = st.checkbox("Add 'Prod' to naming convention", value=True, help="Decide whether to include 'Prod' in generated service offering names")
        
        # Move Delivery Manager here - right after Service Type
        delivery_manager = st.text_input("Delivery Manager", value="")
    
        # Add Business Criticality control
        st.markdown("---")
        business_criticality = st.selectbox(
            "Business Criticality",
            options=["", "1 - most critical", "2 - somewhat critical", "3 - less critical", "4 - not critical"],
            index=0,
            help="Set Business Criticality for all generated offerings. If empty, original values from source files will be used."
        )
        
        # Add Approval Required control
        approval_required = st.checkbox(
            "Approval Required",
            value=False,
            help="Set Approval Required to true for all generated offerings. Default is always false."
        )
        
        # Add conditional controls for approval required details
        if approval_required:
            # Option to use same group for all apps or different per app
            use_per_app_approval = st.checkbox("Use different approval groups per application", value=False)
            
            if not use_per_app_approval:
                # Single approval group for all apps
                approval_required_value = st.text_input(
                    "Approval Details",
                    value="",
                    placeholder="Enter approval group name",
                    help="Same approval group for all applications"
                )
                approval_groups_per_app = {}
            else:
                # Per-app approval groups
                st.markdown("#### Approval Groups by Application")
                st.info("Configure specific approval groups for each application. Leave empty if no approval group is known.")
                
                approval_groups_per_app = {}
                
                if new_apps:
                    # Create columns for better layout
                    cols = st.columns(2)
                    
                    for i, app in enumerate(new_apps):
                        col_idx = i % 2
                        with cols[col_idx]:
                            approval_groups_per_app[app] = st.text_input(
                                f"Approval Group for {app}",
                                value="",
                                key=f"approval_{app}",
                                placeholder=f"e.g., {app} leave empty if not needed)"
                            )
                else:
                    st.warning("No applications defined. Add applications in Basic tab first.")
                
                # Set global approval value to empty for per-app mode
                approval_required_value = "PER_APP"
        else:
            approval_required_value = "empty"
            approval_groups_per_app = {}
        
        # Add Subscribed by Location control
        st.markdown("---")
        change_subscribed_location = st.checkbox(
            "Change Subscribed by Location",
            value=False,
            help="By default, Subscribed by Location will be set to 'Global'"
        )
        
        if change_subscribed_location:
            custom_subscribed_location = st.text_input(
                "Subscribed by Location",
                value="",
                placeholder="Enter custom location",
                help="Custom value for Subscribed by Location column in Excel"
            )
        else:
            custom_subscribed_location = "Global"
        
        # Add Lvl2 checkbox
        st.markdown("---")
        use_lvl2 = st.checkbox(
            "Include Level 2 (Child SO lvl2)",
            help="When checked, search in BOTH Child SO lvl1 AND Child SO lvl2 sheets in Excel"
        )
        
        if use_lvl2:
            service_type = st.text_input(
                "Service Type (for Lvl2 entries)",
                value="",
                placeholder="e.g., Application issue, Hardware problem",
                help="Optional - this will be added to Lvl2 entries"
            )
        else:
            service_type = ""
    
    with tab2:
        st.subheader("Direct Parent Offering Selection")
        
        use_new_parent = st.checkbox(
            "Use NEW specific parent offering (instead of parent keyword search in Excel)",
            help="When checked, you can enter a completely new Parent Offering and Parent name - not inclided in the file yet"
        )
        
        if use_new_parent:
            st.info("📝 Enter the exact Parent Offering and Parent values to use")
            
            # Initialize session state for dynamic parent offerings
            if 'parent_offerings' not in st.session_state:
                st.session_state.parent_offerings = [{"offering": "", "parent": ""}]
            
            st.markdown("### Parent Offering")
            
            # Display existing pairs
            for i, pair in enumerate(st.session_state.parent_offerings):
                col1, col2, col3 = st.columns([2, 2, 1])
                
                with col1:
                    pair["offering"] = st.text_input(
                        "New Parent Offering",
                        value=pair["offering"],
                        placeholder="e.g., [Parent HS PL IT] Software assistance",
                        key=f"offering_{i}"
                    )
                
                with col2:
                    pair["parent"] = st.text_input(
                        "New Parent",
                        value=pair["parent"],
                        placeholder="e.g., PL Software Support",
                        key=f"parent_{i}"
                    )
                
                with col3:
                    if len(st.session_state.parent_offerings) > 1:
                        if st.button("➖", key=f"remove_{i}", help="Remove this pair"):
                            st.session_state.parent_offerings.pop(i)
                            st.rerun()
            
            # Add/Remove buttons
            col1, col2 = st.columns(2)
            with col1:
                if st.button("➕ Add Parent Offering", use_container_width=True):
                    st.session_state.parent_offerings.append({"offering": "", "parent": ""})
                    st.rerun()
            
            with col2:
                if len(st.session_state.parent_offerings) > 1:
                    if st.button("➖ Remove Last", use_container_width=True):
                        st.session_state.parent_offerings.pop()
                        st.rerun()
            
            # Convert to the format expected by the backend
            new_parent_offerings = "\n".join([pair["offering"] for pair in st.session_state.parent_offerings if pair["offering"]])
            new_parents = "\n".join([pair["parent"] for pair in st.session_state.parent_offerings if pair["parent"]])
            
            # Show preview
            if new_parent_offerings and new_parents:
                st.success("✅ **Preview of Parent Offering pairs:**")
                for i, pair in enumerate(st.session_state.parent_offerings):
                    if pair["offering"] and pair["parent"]:
                        st.write(f"{i+1}. Parent Offering: `{pair['offering']}` → New Parent: `{pair['parent']}`")
        else:
            new_parent_offerings = ""
            new_parents = ""
            # Clear session state when not using new parent
            if 'parent_offerings' in st.session_state:
                del st.session_state.parent_offerings
    
    with tab3:
        st.subheader("Schedule Settings")
        
        schedule_type = st.checkbox("Custom schedule per period")
        create_multiple_schedules = st.checkbox("Create multiple schedules", help="Generate the same offerings with different schedules")
        
        # Initialize schedule_suffixes
        schedule_suffixes = []
        
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
        use_per_country_schedules = st.checkbox("Use different schedules per country", help="Define specific schedules for different countries")
        schedule_settings_per_country = {}
        
        if use_per_country_schedules:
            st.markdown("Schedule Settings per Country")
            
            # Define available countries/divisions - UPDATED with new countries
            available_countries = ["HS PL", "DS PL", "DE", "MD", "UA", "DS CY", "DS RO", "DS TR"]
            
            # Create tabs or columns for different countries
            country_tabs = st.tabs(available_countries)
            
            for idx, country in enumerate(available_countries):
                with country_tabs[idx]:
                    st.markdown(f"{country} Schedules")
                    
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
        
        st.markdown("### SLA")
        col_rsp, col_rsl = st.columns(2)
        with col_rsp:
            rsp_duration = st.text_input("RSP Duration", value="")
        with col_rsl:
            rsl_duration = st.text_input("RSL Duration", value="")
    
    with tab4:
        st.subheader("Service Commitments")
        
        use_custom_commitments = st.checkbox("Define custom Service Commitments", help="If unchecked, commitments will be copied from the source files")
        
        if use_custom_commitments:
            # UPDATED country list with RO and TR
            commitment_country = st.selectbox("Country", ["CY", "DE", "MD", "PL", "RO", "TR", "UA"])
            
            st.markdown("Service Commitments Configuration")
            
            # Initialize commitment lines list
            commitment_lines = []
            
            # Create multiple commitment entries
            num_commitments = st.number_input("Number of commitment", min_value=1, max_value=10, value=2)
            
            for i in range(num_commitments):
                st.markdown(f"Commitment Line {i+1}")
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
                    # Check if we should use custom prefix from Advanced tab
                    if st.session_state.get('use_custom_depend_on', False) and st.session_state.get('depend_on_prefix'):
                        prefix_to_use = st.session_state.get('depend_on_prefix')
                        if prefix_to_use == "Global":
                            if st.session_state.get('special_it', False):
                                prefix_to_use = "Global"
                            else:
                                prefix_to_use = "Global Prod" if st.session_state.get('global_prod', False) else "Global"
                        else:
                            if st.session_state.get('special_it', False):
                                prefix_to_use = st.session_state.get('depend_on_prefix')
                            else:
                                if st.session_state.get('global_prod', False):
                                    prefix_to_use = f"{st.session_state.get('depend_on_prefix')} Prod"
                                else:
                                    prefix_to_use = st.session_state.get('depend_on_prefix')
                    else:
                        prefix_to_use = commitment_country
                    
                    line = f"[{prefix_to_use}] SLA {sr_or_im} {line_type} {schedule} {priority} {time}"
                    commitment_lines.append(line)
                    
                    # Add OLA for SR and RSL
                    if sr_or_im == "SR" and line_type == "RSL":
                        ola_line = f"[{prefix_to_use}] OLA {sr_or_im} RSL {schedule} {priority} {time}"
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
                help="Optional - if empty, Support Group value will be copied"
            )
            # Instead of empty dicts, populate with all countries
            support_groups_per_country = {}
            managed_by_groups_per_country = {}
            
            # Populate dictionaries with global values for all countries
            all_countries = ["HS PL", "DS PL", "DE", "UA", "MD", "CY", "RO", "TR"]
            for country_key in all_countries:
                if support_group:  # Only populate if there's a value
                    support_groups_per_country[country_key] = support_group
                    managed_by_groups_per_country[country_key] = managed_by_group if managed_by_group else support_group
        else:
            # Per-country support groups
            st.markdown("Support Groups by Country")
            st.info("Select countries to configure.")
            
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
                                    placeholder=f"Optional - uses Support Group if empty",
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
                                placeholder=f"Optional - uses Support Group if empty",
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
        st.subheader("Select proper naming convention:")
        
        # Create a column to ensure vertical layout - ALPHABETICAL ORDER
        col = st.container()
        with col:
            # List in alphabetical order
            require_corp = st.checkbox("CORP")
            require_recp = st.checkbox("CORP RecP")
            require_corp_dedicated = st.checkbox("CORP Dedicated Services")
            require_corp_it = st.checkbox("CORP IT")
            require_dedicated = st.checkbox("Dedicated Services")
            special_dak = st.checkbox("DAK (Business Services)")
            special_hr = st.checkbox("HR")
            special_it = st.checkbox("IT")
            special_medical = st.checkbox("Medical")
        
        # Ensure only one is selected
        all_selected = sum([require_corp, require_recp, special_it, special_hr, special_medical, special_dak, require_corp_it, require_corp_dedicated, require_dedicated])
        if all_selected > 1:
            st.error("⚠️ Please select only one naming type")
            # Reset all to handle multiple selection
            require_corp = require_recp = special_it = special_hr = special_medical = special_dak = require_corp_it = require_corp_dedicated = require_dedicated = False
        elif all_selected == 0:
            st.info("Standard naming will be used if nothing is selected")
        
        if require_corp or require_recp or require_corp_it or require_corp_dedicated:
            delivering_tag = st.text_input(
                "Who delivers the service?", 
                value="",
                help="E.g. HS PL, DS DE, IT, Finance, etc."
            )
        else:
            delivering_tag = ""
        
        # Save IT checkbox to session state
        st.session_state['special_it'] = special_it
    
    with tab7:
        # Global settings - MOVED TO TOP
        st.markdown("### Global")
        global_prod = st.checkbox("Global Prod value for Service Offerings column", value=False, key="global_prod_checkbox")
        
        # Store in session state for use in other tabs
        st.session_state['global_prod'] = global_prod
        
        # Remove pluralization checkbox - always use pluralization
        use_pluralization = True  # Always enabled
        
        # Custom Depend On setting - NOW AFTER GLOBAL SETTINGS
        st.markdown("### Service Offerings | Depend On")
        use_custom_depend_on = st.checkbox("Use custom value for column 'Service Offerings | Depend On'", value=False, 
                                          help="Overrides automatic value based on selected prefix and applications")
        
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
                # Store in session state
                st.session_state['depend_on_prefix'] = depend_on_prefix
                st.session_state['use_custom_depend_on'] = True
            
            with col2:
                # Show which apps will be used automatically
                if new_apps:
                    # Calculate dynamic height based on number of apps
                    num_apps = len(new_apps)
                    # Base height of 60px + 25px per additional app, with minimum 80px and maximum 300px
                    dynamic_height = max(80, min(300, 60 + (num_apps * 25)))
                    
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
                            height=dynamic_height,
                            help=f"Will automatically use each app: {', '.join(new_apps)}"
                        )
                    app_names_display = new_apps
                else:
                    st.text_input(
                        "Application Name",
                        value="(no apps specified)",
                        disabled=True,
                        help="Add applications in Basic tab to see them"
                    )
                    app_names_display = ["(no app)"]
            
            # Construct the custom depend on value
            current_special_it = st.session_state.get('special_it', False)
            
            # PREFIX FOR SERVICE OFFERINGS | DEPEND ON - WITHOUT PROD (backend will add if needed)
            depend_on_prefix_tag = depend_on_prefix
            
            # Show preview(s) - WITH PROD FOR DISPLAY ONLY
            if new_apps:
                if len(new_apps) == 1:
                    app_name = get_plural_form_preview(new_apps[0]) if use_pluralization else new_apps[0]
                    # PREVIEW with Prod if Global Prod is checked
                    preview_prefix = f"{depend_on_prefix_tag} Prod" if global_prod else depend_on_prefix_tag
                    preview_value = f"[{preview_prefix}] {app_name}"
                    st.info(f"Preview: `{preview_value}` (IT: {current_special_it}, Global Prod: {global_prod})")
                    # But send to backend without Prod
                    custom_depend_on_value = f"[{depend_on_prefix_tag}] {app_name}"
                else:
                    preview_prefix = f"{depend_on_prefix_tag} Prod" if global_prod else depend_on_prefix_tag
                    st.info(f"Preview for each app: (IT: {current_special_it}, Global Prod: {global_prod})")
                    for app in new_apps:
                        app_name = get_plural_form_preview(app) if use_pluralization else app
                        st.text(f"• `[{preview_prefix}] {app_name}`")
                    custom_depend_on_value = f"[{depend_on_prefix_tag}]"
            else:
                preview_prefix = f"{depend_on_prefix_tag} Prod" if global_prod else depend_on_prefix_tag
                preview_value = f"[{preview_prefix}]"
                st.info(f"Preview: `{preview_value}` (IT: {current_special_it}, Global Prod: {global_prod})")
                custom_depend_on_value = f"[{depend_on_prefix_tag}]"
        else:
            custom_depend_on_value = ""
            st.session_state['use_custom_depend_on'] = False
        
        # Aliases - SIMPLIFIED VERSION
        st.markdown("### Aliases")

        # Option to use same values as app names - MAIN CHECKBOX
        use_same_as_apps = st.checkbox("Use same values as Application Names", 
                                      value=False,
                                      help="When checked, aliases will automatically use the same values as the application names but for English names ONLY")

        if use_same_as_apps:
            # Auto-enable aliases and set to use app names
            aliases_on = True
            aliases_value = "USE_APP_NAMES"
            selected_languages = ["ENG"]  # Automatically set to ENG only
            
            if new_apps:
                st.success(f"✅ **Aliases will use:** {', '.join(new_apps)} **in ENG column**")
            else:
                st.warning("⚠️ No application names defined. Add applications in Basic tab first.")
        else:
            # When aliases are disabled
            aliases_on = False
            aliases_value = ""
            selected_languages = []

        # Remove per-country aliases variables for backend compatibility
        use_per_country_aliases = False
        aliases_per_country = {}

st.markdown("---")

# GENERATE BUTTON AND VALIDATION
if st.button("🚀 Generate Service Offerings", type="primary", use_container_width=True):
    if not uploaded_files:
        st.error("⚠️ Please upload at least one Excel file")
    elif use_new_parent and (not new_parent_offerings or not new_parents):
        st.error("⚠️ When using specific parent offering, please add at least one Parent Offering")
    elif not use_new_parent and not keywords_parent and not keywords_child:
        st.error("⚠️ Please enter at least one keyword in either Parent Offering or Child Service Offering")
    elif not schedule_suffixes or not any(schedule_suffixes):
        st.error("⚠️ Please configure at least one schedule")
    elif all_selected > 1:
        st.error("⚠️ Please select only one naming type")
    else:
        try:
            with tempfile.TemporaryDirectory() as temp_dir:
                src_dir = Path(temp_dir) / "input"
                out_dir = Path(temp_dir) / "output"
                src_dir.mkdir(exist_ok=True)
                out_dir.mkdir(exist_ok=True)
                
                # Save uploaded files
                for uploaded_file in uploaded_files:
                    file_path = src_dir / uploaded_file.name
                    with open(file_path, "wb") as f:
                        f.write(uploaded_file.getbuffer())
                
                with st.spinner("🔄 Generating service offerings..."):
                    try:
                        result_file = run_generator(
                            keywords_parent=keywords_parent if not use_new_parent else "",
                            keywords_child=keywords_child if not use_new_parent else "",
                            new_apps=new_apps,
                            schedule_suffixes=schedule_suffixes,
                            delivery_manager=delivery_manager,
                            global_prod=global_prod,
                            use_pluralization=use_pluralization,
                            rsp_duration=rsp_duration,
                            rsl_duration=rsl_duration,
                            sr_or_im=sr_or_im,
                            require_corp=require_corp,
                            require_recp=require_recp,
                            delivering_tag=delivering_tag,
                            support_group=support_group,
                            managed_by_group=managed_by_group,
                            aliases_on=aliases_on,
                            aliases_value=aliases_value,
                            aliases_per_country=aliases_per_country,
                            src_dir=src_dir,
                            out_dir=out_dir,
                            special_it=special_it,
                            special_hr=special_hr,
                            special_medical=special_medical,
                            special_dak=special_dak,
                            use_custom_commitments=use_custom_commitments,
                            custom_commitments_str=custom_commitments_str,
                            commitment_country=commitment_country,
                            require_corp_it=require_corp_it,
                            require_corp_dedicated=require_corp_dedicated,
                            require_dedicated=require_dedicated,
                            use_new_parent=use_new_parent,
                            new_parent_offering=new_parent_offerings,
                            new_parent=new_parents,
                            keywords_excluded=keywords_excluded if not use_new_parent else "",
                            use_lvl2=use_lvl2,
                            service_type_lvl2=service_type,
                            support_groups_per_country=support_groups_per_country,
                            managed_by_groups_per_country=managed_by_groups_per_country,
                            schedule_settings_per_country=schedule_settings_per_country,
                            use_custom_depend_on=use_custom_depend_on,
                            custom_depend_on_value=custom_depend_on_value,
                            selected_languages=selected_languages,
                            business_criticality=business_criticality,
                            approval_required=approval_required,
                            approval_required_value=approval_required_value,
                            approval_groups_per_app=approval_groups_per_app,
                            change_subscribed_location=change_subscribed_location,
                            custom_subscribed_location=custom_subscribed_location,
                            add_prod=add_prod
                        )
                        
                        # FIXED HANDLING FOR BOTH True AND Path RETURNS
                        if result_file is True:
                            # Legacy behavior - search for the file
                            import time
                            time.sleep(0.5)  # Give file system time to sync
                            
                            # Search for generated files
                            excel_files = list(out_dir.glob("Generated_Service_Offerings_*.xlsx"))
                            if excel_files:
                                # Use the most recent file
                                result_file = max(excel_files, key=lambda p: p.stat().st_mtime)
                                st.info(f"Found generated file: {result_file.name}")
                            else:
                                # Try alternative pattern
                                all_xlsx = list(out_dir.glob("*.xlsx"))
                                if all_xlsx:
                                    result_file = max(all_xlsx, key=lambda p: p.stat().st_mtime)
                                    st.info(f"Found file: {result_file.name}")
                                else:
                                    st.error("❌ No Excel files found in output directory")
                                    result_file = None
                        
                        # Convert string path to Path object
                        elif isinstance(result_file, str):
                            result_file = Path(result_file)
                            
                        # Validate it's a Path object
                        elif result_file and not isinstance(result_file, Path):
                            st.error(f"❌ Invalid result type: {type(result_file)}")
                            result_file = None
                            
                    except Exception as gen_error:
                        st.error(f"❌ Error during generation: {str(gen_error)}")
                        st.exception(gen_error)
                        result_file = None
                
                # Check if file exists and download
                if result_file and isinstance(result_file, Path) and result_file.exists():
                    # Read the file
                    with open(result_file, "rb") as f:
                        file_data = f.read()
                    
                    if len(file_data) > 0:
                        st.success("✅ Service offerings generated successfully!")
                        st.download_button(
                            label="📥 Download generated file",
                            data=file_data,
                            file_name=result_file.name,
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            use_container_width=True
                        )
                        st.info(f"Generated: {result_file.name} ({len(file_data):,} bytes)")
                    else:
                        st.error("❌ Generated file is empty")
                else:
                    st.error("❌ Failed to generate file. Please check your configuration.")
                    if result_file:
                        st.error(f"Debug: result_file = {result_file}, exists = {result_file.exists() if isinstance(result_file, Path) else 'N/A'}")
                    
        except ValueError as e:
            error_msg = str(e)
            if "duplicate" in error_msg.lower() or "no matching offerings" in error_msg.lower():
                st.error(f"❌ {error_msg}")
            else:
                st.error(f"❌ Error: {error_msg}")
        except Exception as e:
            st.error(f"❌ Unexpected error: {str(e)}")
            st.exception(e)

st.markdown("---")
st.markdown(
    """
    <div style='text-align: center; color: gray;'>
        Service Offerings App
    </div>
    """,
    unsafe_allow_html=True
)